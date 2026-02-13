"""
production.py  v2.0.3
=========================
FIXES:
 1. SocketLineReader: raw socket.recv() instead of makefile().readline().
    makefile streams break permanently after socket.timeout in Python.
 2. Reconnect logic moved INTO monitoring thread (was in health_check on
    main thread, causing race condition: health_check closes socket while
    monitoring thread is mid-recv, simulator sees "connection aborted").
 3. Metal parse: extracts first_value from full data lines, threshold >= 100.
 4. sp_InsertLogEntry: 7 params including metal_value.
"""
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import threading, time, socket, re, os
from datetime import datetime, timedelta

from shared_config import (
    APP_TITLE, APP_VERSION, COLORS, DB, make_logger,
    get_local_ip, get_dropdown, add_dropdown,
    validate_ip, validate_port, sanitize,
)

C = COLORS
METAL_THRESHOLD = 100


class SocketLineReader:
    """Read lines from raw TCP socket using recv() + manual buffer.
    Unlike makefile().readline(), this survives socket.timeout cleanly."""

    def __init__(self, sock, encoding='latin-1'):
        self.sock = sock
        self.encoding = encoding
        self.buffer = b''

    def readline(self):
        """Return one decoded line (stripped). Raises socket.timeout if no
        complete line available. Raises ConnectionError if socket closed."""
        while True:
            idx = self.buffer.find(b'\n')
            if idx >= 0:
                line = self.buffer[:idx]
                self.buffer = self.buffer[idx + 1:]
                return line.decode(self.encoding, errors='replace').rstrip('\r')
            chunk = self.sock.recv(4096)
            if not chunk:
                raise ConnectionError("Connection closed by remote")
            self.buffer += chunk

    def clear(self):
        self.buffer = b''


class ProductionApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"{APP_TITLE} - Production v{APP_VERSION}")
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.log = make_logger("production")
        self.db = DB(self.log)
        if not self.db.connect():
            messagebox.showerror("DB Error", "Cannot connect to database.")
        self.pc_ip = get_local_ip()
        self.machine_name = self.pc_ip
        self.weigher_ip = "192.168.0.100"
        self.weigher_port = 50001
        self.metal_ip = None
        self.metal_port = 50001
        self.w_sock = None; self.w_reader = None
        self.m_sock = None; self.m_reader = None
        self.running = False; self.thread = None
        self.stop_evt = threading.Event()
        self.user = self.prod_id = self.session_data = self.screen = None
        self.entries = {}
        self.w_queue = []; self.m_queue = []; self.db_ops = []
        self.next_log = 1; self.live_data = []
        self.stats = dict(total=0, passed=0, under=0, over=0, metal=0)
        self.last_match_debug = datetime.now()
        self.last_cleanup = datetime.now()
        sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
        w, h = min(1400, sw - 60), min(850, sh - 80)
        root.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")
        root.minsize(1000, 600); root.configure(bg=C["bg_dark"])
        self._load_machine_info(); self._styles(); self._show_login()

    # ═══════════════ STYLES ═══════════════
    def _styles(self):
        s = ttk.Style(); s.theme_use("clam")
        s.configure("Dark.TFrame", background=C["bg_dark"])
        s.configure("Card.TFrame", background=C["bg_card"])
        s.configure("Mid.TFrame", background=C["bg_mid"])
        for n, bg, fg, fnt in [
            ("Title.TLabel", C["bg_dark"], C["text"], ("Segoe UI", 18, "bold")),
            ("Sub.TLabel", C["bg_dark"], C["text2"], ("Segoe UI", 10)),
            ("Head.TLabel", C["bg_card"], C["text"], ("Segoe UI", 11, "bold")),
            ("Norm.TLabel", C["bg_card"], C["text2"], ("Segoe UI", 10)),
            ("Card2.TLabel", C["bg_card"], C["accent"], ("Segoe UI", 12, "bold")),
        ]:
            s.configure(n, background=bg, foreground=fg, font=fnt)
        for n, bg, fg in [
            ("Accent.TButton", C["accent"], "#fff"),
            ("Green.TButton", C["green"], "#1a1a2e"),
            ("Red.TButton", C["red"], "#fff"),
            ("Ghost.TButton", C["bg_card"], C["text"]),
        ]:
            s.configure(n, background=bg, foreground=fg,
                        font=("Segoe UI", 10, "bold"),
                        borderwidth=0, padding=(14, 7), focuscolor="none")
            s.map(n, background=[("active", C["accent_hover"])])
        s.configure("Dark.TEntry", fieldbackground=C["bg_input"],
                    foreground=C["text"], borderwidth=1, insertcolor=C["text"])
        s.configure("Dark.TCombobox", fieldbackground=C["bg_input"], foreground=C["text"])
        s.map("Dark.TCombobox", fieldbackground=[("readonly", C["bg_input"])])
        s.configure("TNotebook", background=C["bg_dark"], borderwidth=0)
        s.configure("TNotebook.Tab", background=C["bg_card"], foreground=C["text2"],
                    font=("Segoe UI", 10, "bold"), padding=(14, 7))
        s.map("TNotebook.Tab", background=[("selected", C["accent"])],
              foreground=[("selected", "#fff")])

    # ═══════════════ HELPERS ═══════════════
    def _clear(self):
        for w in self.root.winfo_children(): w.destroy()

    def _header(self, parent, title):
        h = ttk.Frame(parent, style="Mid.TFrame"); h.pack(fill="x")
        inner = ttk.Frame(h, style="Mid.TFrame"); inner.pack(fill="x", padx=20, pady=10)
        ttk.Label(inner, text=title, font=("Segoe UI", 16, "bold"),
                  background=C["bg_mid"], foreground=C["text"]).pack(side="left")
        if self.user:
            r = ttk.Frame(inner, style="Mid.TFrame"); r.pack(side="right")
            if self.running:
                ttk.Label(r, text="LIVE", font=("Segoe UI", 10, "bold"),
                          background=C["bg_mid"], foreground=C["red"]).pack(side="left", padx=8)
            ttk.Label(r, text=f"{self.user} | {self.machine_name}",
                      font=("Segoe UI", 9), background=C["bg_mid"],
                      foreground=C["text2"]).pack(side="left")
        tk.Frame(parent, height=1, bg=C["border"]).pack(fill="x")

    def _card(self, parent, title=None):
        c = ttk.Frame(parent, style="Card.TFrame"); c.configure(padding=14)
        if title:
            ttk.Label(c, text=title, style="Card2.TLabel").pack(anchor="w", pady=(0, 6))
            tk.Frame(c, height=1, bg=C["border"]).pack(fill="x", pady=(0, 10))
        return c

    def _load_machine_info(self):
        try:
            r = self.db.call_sp("sp_GetMachineByIP", [self.pc_ip], fetch=True)
            if r:
                self.machine_name = r[0][0]
                self.weigher_ip = r[0][1]
                self.metal_ip = r[0][2] if len(r[0]) > 2 else None
        except Exception as e:
            self.log.warning(f"Machine info load: {e}")

    # ═══════════════ LOGIN ═══════════════
    def _show_login(self):
        self._clear(); self.screen = "login"
        m = ttk.Frame(self.root, style="Dark.TFrame"); m.pack(fill="both", expand=True)
        ctr = ttk.Frame(m, style="Dark.TFrame"); ctr.place(relx=0.5, rely=0.42, anchor="center")
        ttk.Label(ctr, text=APP_TITLE, style="Title.TLabel").pack(pady=(0, 2))
        ttk.Label(ctr, text="Production Control System", style="Sub.TLabel").pack(pady=(0, 28))
        card = self._card(ctr, "Sign In"); card.pack(ipadx=40, ipady=8)
        f = ttk.Frame(card, style="Card.TFrame"); f.pack(pady=8)
        ttk.Label(f, text="Username", style="Norm.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 3))
        self._login_user = ttk.Entry(f, width=26, font=("Segoe UI", 11), style="Dark.TEntry")
        self._login_user.grid(row=1, column=0, pady=(0, 10), ipady=4); self._login_user.focus()
        ttk.Label(f, text="Password", style="Norm.TLabel").grid(row=2, column=0, sticky="w", pady=(0, 3))
        self._login_pass = ttk.Entry(f, width=26, show="*", font=("Segoe UI", 11), style="Dark.TEntry")
        self._login_pass.grid(row=3, column=0, pady=(0, 14), ipady=4)
        ttk.Button(card, text="Sign In", command=self._do_login, style="Accent.TButton").pack(fill="x")
        ttk.Label(m, text=f"v{APP_VERSION}", style="Sub.TLabel").pack(side="bottom", pady=10)
        self.root.bind("<Return>", lambda e: self._do_login())

    def _do_login(self):
        u, p = sanitize(self._login_user.get(), 50), self._login_pass.get()
        if not u or not p: messagebox.showwarning("Login", "Enter both fields."); return
        try:
            r = self.db.call_sp("sp_VerifyUser", [u, p], fetch=True)
            if r:
                self.user = u; self.root.unbind("<Return>"); self._show_main()
            else:
                messagebox.showerror("Denied", "Invalid credentials.")
                self._login_pass.delete(0, tk.END); self._login_pass.focus()
        except Exception as e:
            self.log.error(f"Login error: {e}"); messagebox.showerror("Error", "Authentication error.")

    # ═══════════════ MAIN ═══════════════
    def _show_main(self):
        self._clear(); self.screen = "main"
        m = ttk.Frame(self.root, style="Dark.TFrame"); m.pack(fill="both", expand=True)
        self._header(m, "Dashboard")
        ct = ttk.Frame(m, style="Dark.TFrame"); ct.pack(fill="both", expand=True, padx=28, pady=18)
        cards = ttk.Frame(ct, style="Dark.TFrame"); cards.pack(fill="x", pady=(0, 16))
        for i, (ttl, desc, cmd) in enumerate([
            ("Machine Config", "Configure weigher & metal detector", self._show_config),
            ("Production", "Start/stop sessions with live data", self._show_production),
            ("Reports", "Launch report app for analytics", self._launch_reports)]):
            c = ttk.Frame(cards, style="Card.TFrame")
            c.pack(side="left", fill="both", expand=True, padx=(0 if i == 0 else 6, 0))
            c.configure(padding=18)
            ttk.Label(c, text=ttl, font=("Segoe UI", 13, "bold"),
                      background=C["bg_card"], foreground=C["text"]).pack(anchor="w", pady=(6, 3))
            ttk.Label(c, text=desc, font=("Segoe UI", 9),
                      background=C["bg_card"], foreground=C["text3"], wraplength=190).pack(anchor="w", pady=(0, 10))
            ttk.Button(c, text="Open", command=cmd, style="Accent.TButton").pack(anchor="w")
        sc = self._card(ct, "System Status"); sc.pack(fill="x")
        g = ttk.Frame(sc, style="Card.TFrame"); g.pack(fill="x")
        db_ok = self.db.alive()
        for i, (l, v, clr) in enumerate([
            ("Database", "Connected" if db_ok else "Disconnected", C["green"] if db_ok else C["red"]),
            ("Machine", self.machine_name, C["text2"]),
            ("Weigher", f"{self.weigher_ip}:{self.weigher_port}", C["text2"]),
            ("Metal", f"{self.metal_ip}:{self.metal_port}" if self.metal_ip else "N/A", C["text3"]),
            ("Production", "LIVE" if self.running else "Idle", C["red"] if self.running else C["green"])]):
            ttk.Label(g, text=f"{l}:", font=("Segoe UI", 9, "bold"),
                      background=C["bg_card"], foreground=C["text3"]).grid(row=i, column=0, sticky="w", padx=(0, 10), pady=2)
            ttk.Label(g, text=v, font=("Segoe UI", 9),
                      background=C["bg_card"], foreground=clr).grid(row=i, column=1, sticky="w", pady=2)
        btm = ttk.Frame(m, style="Dark.TFrame"); btm.pack(fill="x", padx=28, pady=8)
        ttk.Button(btm, text="Logout", command=self._logout, style="Ghost.TButton").pack(side="right")

    def _launch_reports(self):
        import subprocess, sys
        d = os.path.dirname(os.path.abspath(__file__))
        rp = os.path.join(d, "report.py")
        if os.path.exists(rp): subprocess.Popen([sys.executable, rp])
        else: messagebox.showwarning("Not Found", f"report.py not found in:\n{d}")

    # ═══════════════ CONFIG ═══════════════
    def _show_config(self):
        self._clear(); self.screen = "config"
        m = ttk.Frame(self.root, style="Dark.TFrame"); m.pack(fill="both", expand=True)
        self._header(m, "Machine Configuration")
        ct = ttk.Frame(m, style="Dark.TFrame"); ct.pack(fill="both", expand=True, padx=28, pady=18)
        card = self._card(ct, "Network Settings"); card.pack(fill="x")
        f = ttk.Frame(card, style="Card.TFrame"); f.pack(fill="x")
        ttk.Label(f, text="PC IP:", style="Norm.TLabel").grid(row=0, column=0, sticky="w", pady=5)
        ttk.Label(f, text=self.pc_ip, font=("Consolas", 11, "bold"),
                  background=C["bg_card"], foreground=C["accent"]).grid(row=0, column=1, sticky="w", padx=8)
        ttk.Label(f, text="Weigher IP:", style="Head.TLabel").grid(row=1, column=0, sticky="w", pady=5)
        self._cfg_wip = ttk.Entry(f, width=18, font=("Consolas", 11), style="Dark.TEntry")
        self._cfg_wip.grid(row=1, column=1, padx=8, ipady=2); self._cfg_wip.insert(0, self.weigher_ip)
        ttk.Label(f, text="Port:", style="Norm.TLabel").grid(row=1, column=2, padx=(8, 0))
        self._cfg_wpt = ttk.Entry(f, width=7, font=("Consolas", 11), style="Dark.TEntry")
        self._cfg_wpt.grid(row=1, column=3, padx=8, ipady=2); self._cfg_wpt.insert(0, str(self.weigher_port))
        self._cfg_ws = ttk.Label(f, text="", style="Norm.TLabel"); self._cfg_ws.grid(row=1, column=4, padx=8)
        ttk.Button(f, text="Test", command=lambda: self._test_conn(self._cfg_wip, self._cfg_wpt, self._cfg_ws),
                   style="Ghost.TButton").grid(row=1, column=5)
        ttk.Label(f, text="Metal IP:", style="Head.TLabel").grid(row=2, column=0, sticky="w", pady=5)
        self._cfg_mip = ttk.Entry(f, width=18, font=("Consolas", 11), style="Dark.TEntry")
        self._cfg_mip.grid(row=2, column=1, padx=8, ipady=2)
        if self.metal_ip: self._cfg_mip.insert(0, self.metal_ip)
        ttk.Label(f, text="Port:", style="Norm.TLabel").grid(row=2, column=2, padx=(8, 0))
        self._cfg_mpt = ttk.Entry(f, width=7, font=("Consolas", 11), style="Dark.TEntry")
        self._cfg_mpt.grid(row=2, column=3, padx=8, ipady=2); self._cfg_mpt.insert(0, str(self.metal_port))
        self._cfg_ms = ttk.Label(f, text="", style="Norm.TLabel"); self._cfg_ms.grid(row=2, column=4, padx=8)
        ttk.Button(f, text="Test", command=lambda: self._test_conn(self._cfg_mip, self._cfg_mpt, self._cfg_ms),
                   style="Ghost.TButton").grid(row=2, column=5)
        bf = ttk.Frame(ct, style="Dark.TFrame"); bf.pack(fill="x", pady=14)
        ttk.Button(bf, text="Save & Continue", command=self._save_cfg, style="Green.TButton").pack(side="left", padx=(0, 8))
        ttk.Button(bf, text="Back", command=self._show_main, style="Ghost.TButton").pack(side="right")

    def _test_conn(self, ip_e, port_e, lbl):
        ip, pt = ip_e.get().strip(), validate_port(port_e.get())
        if not ip: lbl.config(text="No IP", foreground=C["text3"]); return
        if not validate_ip(ip): lbl.config(text="Bad IP", foreground=C["red"]); return
        if not pt: lbl.config(text="Bad port", foreground=C["red"]); return
        lbl.config(text="...", foreground=C["orange"]); self.root.update()
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            s.settimeout(5); s.connect((ip, pt)); s.close()
            lbl.config(text="OK", foreground=C["green"])
        except Exception: lbl.config(text="Fail", foreground=C["red"])

    def _save_cfg(self):
        ip, pt = self._cfg_wip.get().strip(), validate_port(self._cfg_wpt.get())
        if not validate_ip(ip): messagebox.showwarning("Validation", "Invalid weigher IP"); return
        if not pt: messagebox.showwarning("Validation", "Invalid weigher port"); return
        self.weigher_ip, self.weigher_port = ip, pt
        mip, mpt = self._cfg_mip.get().strip(), validate_port(self._cfg_mpt.get())
        if mip:
            if not validate_ip(mip): messagebox.showwarning("Validation", "Invalid metal IP"); return
            if not mpt: messagebox.showwarning("Validation", "Invalid metal port"); return
            self.metal_ip, self.metal_port = mip, mpt
        else: self.metal_ip = None
        self._show_main()

    # ═══════════════ PRODUCTION SCREEN ═══════════════
    def _show_production(self):
        self._clear(); self.screen = "production"
        m = ttk.Frame(self.root, style="Dark.TFrame"); m.pack(fill="both", expand=True)
        self._header(m, "Production Session")
        nav = ttk.Frame(m, style="Mid.TFrame"); nav.pack(fill="x")
        ni = ttk.Frame(nav, style="Mid.TFrame"); ni.pack(fill="x", padx=20, pady=5)
        ttk.Button(ni, text="Dashboard", command=self._show_main, style="Ghost.TButton").pack(side="left")
        ct = ttk.Frame(m, style="Dark.TFrame"); ct.pack(fill="both", expand=True, padx=12, pady=8)
        pw = ttk.PanedWindow(ct, orient=tk.HORIZONTAL); pw.pack(fill="both", expand=True)
        left = ttk.Frame(pw, style="Dark.TFrame"); pw.add(left, weight=2)
        right = ttk.Frame(pw, style="Dark.TFrame"); pw.add(right, weight=3)
        self._build_form(left); self._build_monitor(right)
        if self.running: self._restore_state()

    def _build_form(self, parent):
        self.entries = {}
        sc = self._card(parent); sc.pack(fill="x", pady=(0, 6))
        self.lbl_id = ttk.Label(sc, text="Ready", font=("Segoe UI", 12, "bold"),
                                 background=C["bg_card"], foreground=C["text3"]); self.lbl_id.pack(side="left")
        self.lbl_st = ttk.Label(sc, text="IDLE", font=("Segoe UI", 10, "bold"),
                                 background=C["bg_card"], foreground=C["green"]); self.lbl_st.pack(side="right")
        nb = ttk.Notebook(parent); nb.pack(fill="both", expand=True, pady=(0, 6))
        t1 = ttk.Frame(nb, style="Card.TFrame", padding=10); nb.add(t1, text="  Basic  ")
        t2 = ttk.Frame(nb, style="Card.TFrame", padding=10); nb.add(t2, text="  Product  ")
        t3 = ttk.Frame(nb, style="Card.TFrame", padding=10); nb.add(t3, text="  Weight  ")
        self._form_basic(t1); self._form_product(t2); self._form_weight(t3)
        bf = ttk.Frame(parent, style="Dark.TFrame"); bf.pack(fill="x")
        self.btn_chk = ttk.Button(bf, text="Check Lot", command=self._check_lot, style="Accent.TButton")
        self.btn_chk.pack(side="left", padx=(0, 4))
        self.btn_go = ttk.Button(bf, text="Start", command=self._start, style="Green.TButton")
        self.btn_go.pack(side="left", padx=4)
        self.btn_stp = ttk.Button(bf, text="Stop", command=self._stop, style="Red.TButton")
        self.btn_stp.pack(side="left", padx=4); self.btn_stp.configure(state="disabled")

    def _dd_row(self, p, row, label, field, table):
        p.columnconfigure(1, weight=1)
        ttk.Label(p, text=label, style="Norm.TLabel").grid(row=row, column=0, sticky="w", pady=3, padx=(0, 6))
        vals = get_dropdown(self.db, table)
        cb = ttk.Combobox(p, values=vals, state="readonly", style="Dark.TCombobox", width=16)
        if vals: cb.set(vals[0])
        cb.grid(row=row, column=1, sticky="ew", pady=3); self.entries[field] = cb
        ttk.Button(p, text="+", command=lambda: self._add_dlg(field, table, label),
                   style="Ghost.TButton", width=2).grid(row=row, column=2, pady=3, padx=(3, 0))

    def _form_basic(self, p):
        for i, (l, f, t) in enumerate([("Lot Number *", "lot_no", "lot_numbers"), ("Shift", "shift", "shifts"),
            ("Product", "product", "products"), ("Buyer", "buyer", "buyers"),
            ("Contract", "contract", "contracts"), ("Tank", "tank", "tanks")]):
            self._dd_row(p, i, l, f, t)

    def _form_product(self, p):
        for i, (l, f, t) in enumerate([("Bag Supplier", "bag_supplier", "bag_suppliers"),
            ("Bag Weight (kg)", "bag_weight", "bag_weights"), ("Bag Batch No", "bag_batch_no", "bag_batch_no"),
            ("Packing Type", "type_of_packing", "packing_types"), ("Qty per Bag", "quantity_per_bag", "quantities")]):
            self._dd_row(p, i, l, f, t)

    def _form_weight(self, p):
        for i, (l, f, t) in enumerate([("Target (g)", "net_weight", "net_weights"),
            ("Under Limit (g)", "under_limit", "under_limits"), ("Over Limit (g)", "over_limit", "over_limits")]):
            self._dd_row(p, i, l, f, t)

    def _build_monitor(self, parent):
        sc = self._card(parent); sc.pack(fill="x", pady=(0, 6))
        sg = ttk.Frame(sc, style="Card.TFrame"); sg.pack(fill="x")
        self.st_lbl = {}
        for i, (l, k, clr) in enumerate([("Total", "total", C["text"]), ("Pass", "passed", C["green"]),
            ("Under", "under", C["orange"]), ("Over", "over", C["red"]),
            ("Metal", "metal", C["red"]), ("Rate", "rate", C["blue"])]):
            f = ttk.Frame(sg, style="Card.TFrame"); f.pack(side="left", expand=True)
            ttk.Label(f, text=l, font=("Segoe UI", 8), background=C["bg_card"], foreground=C["text3"]).pack()
            lb = ttk.Label(f, text="0", font=("Segoe UI", 14, "bold"), background=C["bg_card"], foreground=clr)
            lb.pack(); self.st_lbl[k] = lb
        tc = self._card(parent, "Live Data"); tc.pack(fill="both", expand=True)
        self.term = scrolledtext.ScrolledText(tc, wrap=tk.WORD, font=("Consolas", 9),
            bg=C["terminal_bg"], fg=C["green"], insertbackground=C["green"], relief="flat", borderwidth=0)
        self.term.pack(fill="both", expand=True)
        for tag, color in [("pass", C["green"]), ("under", C["orange"]), ("over", C["red"]),
                           ("info", C["blue"]), ("hdr", C["accent"]), ("metal", "#ff3333")]:
            self.term.tag_configure(tag, foreground=color)
        self._tw("Waiting for production start...\n", "info"); self.term.config(state="disabled")

    def _tw(self, text, tag=None):
        self.term.config(state="normal")
        self.term.insert(tk.END, text, tag) if tag else self.term.insert(tk.END, text)
        lines = int(self.term.index("end-1c").split(".")[0])
        if lines > 500: self.term.delete("1.0", f"{lines - 500}.0")
        self.term.see(tk.END); self.term.config(state="disabled")

    def _add_dlg(self, field, table, label):
        dlg = tk.Toplevel(self.root); dlg.title(f"Add {label}"); dlg.geometry("360x180")
        dlg.configure(bg=C["bg_dark"]); dlg.transient(self.root); dlg.grab_set()
        f = ttk.Frame(dlg, style="Dark.TFrame", padding=18); f.pack(fill="both", expand=True)
        ttk.Label(f, text=f"New {label}:", font=("Segoe UI", 11, "bold"),
                  background=C["bg_dark"], foreground=C["text"]).pack(anchor="w", pady=(0, 5))
        e = ttk.Entry(f, font=("Segoe UI", 11), style="Dark.TEntry")
        e.pack(fill="x", ipady=3, pady=(0, 10)); e.focus()
        def do():
            v = sanitize(e.get(), 100)
            if not v: return
            if add_dropdown(self.db, table, v, self.user):
                nv = get_dropdown(self.db, table)
                if field in self.entries: self.entries[field]["values"] = nv; self.entries[field].set(v)
                dlg.destroy()
        ttk.Button(f, text="Add", command=do, style="Green.TButton").pack(fill="x")
        e.bind("<Return>", lambda ev: do())

    # ═══════════════ PRODUCTION CONTROL ═══════════════
    def _check_lot(self):
        lot = self.entries.get("lot_no")
        if not lot or not lot.get().strip(): messagebox.showwarning("Input", "Select a Lot Number."); return
        try:
            r = self.db.call_sp("sp_GetProductionIdByLot", [lot.get().strip(), self.machine_name], fetch=True)
            if r:
                self.prod_id = r[0][0]
                self.lbl_id.config(text=f"ID: {self.prod_id} (existing)", foreground=C["green"])
            else:
                r2 = self.db.call_sp("sp_GetNextProductionId", fetch=True)
                self.prod_id = r2[0][0] if r2 else 1
                self.lbl_id.config(text=f"ID: {self.prod_id} (new)", foreground=C["blue"])
        except Exception as e: messagebox.showerror("Error", str(e))

    def _validate(self):
        if self.prod_id is None: messagebox.showwarning("", "Check Lot first."); return False
        for f in ["shift", "lot_no", "product", "buyer", "contract", "tank", "bag_supplier", "bag_batch_no", "type_of_packing"]:
            if f in self.entries and not self.entries[f].get().strip():
                messagebox.showwarning("Missing", f.replace("_", " ").title()); return False
        for f in ["net_weight", "under_limit", "over_limit", "quantity_per_bag"]:
            try: int(self.entries[f].get())
            except (ValueError, KeyError): messagebox.showwarning("Invalid", f"{f} must be integer"); return False
        try: float(self.entries["bag_weight"].get())
        except (ValueError, KeyError): messagebox.showwarning("Invalid", "Bag weight must be number"); return False
        return True

    def _start(self):
        if not self._validate(): return
        try:
            dt = datetime.now()
            self.session_data = {k: self.entries[k].get().strip() for k in self.entries}
            self.session_data["under_limit"] = int(self.session_data["under_limit"])
            self.session_data["over_limit"] = int(self.session_data["over_limit"])
            pd_list = [self.prod_id, self.session_data["shift"], self.session_data["lot_no"],
                self.session_data["product"], self.session_data["buyer"], self.session_data["contract"],
                self.session_data["tank"], self.session_data["bag_supplier"],
                float(self.session_data["bag_weight"]), self.session_data["bag_batch_no"],
                int(self.session_data["net_weight"]), self.session_data["type_of_packing"],
                int(self.session_data["quantity_per_bag"]), int(self.session_data["over_limit"]),
                int(self.session_data["under_limit"]), dt, sanitize(self.user, 10), self.machine_name]
            self.db.call_sp("sp_InsertUpdateProductionSession", pd_list)
            if not self._start_mon(): return
            self.running = True; self.stats = dict(total=0, passed=0, under=0, over=0, metal=0); self.live_data = []
            self._toggle_form(False); self.btn_go.configure(state="disabled")
            self.btn_stp.configure(state="normal"); self.btn_chk.configure(state="disabled")
            self.lbl_st.config(text="RUNNING", foreground=C["red"])
            self._tw(f"\n{'='*42}\n", "hdr")
            self._tw(f"STARTED | ID:{self.prod_id} | {dt:%H:%M:%S}\n", "hdr")
            self._tw(f"Lot: {self.session_data['lot_no']} | Machine: {self.machine_name}\n", "info")
            self._tw(f"{'='*42}\n\n", "hdr"); self._update_stats()
        except Exception as e:
            self.log.error(f"Start error: {e}"); messagebox.showerror("Error", str(e))

    def _stop(self):
        if not self.running: return
        if not messagebox.askyesno("Confirm", "Stop production?"): return
        self._stop_mon(); self.running = False; self._toggle_form(True)
        self.btn_go.configure(state="normal"); self.btn_stp.configure(state="disabled"); self.btn_chk.configure(state="normal")
        self.lbl_st.config(text="STOPPED", foreground=C["text3"])
        self._tw(f"\n{'='*42}\n", "hdr")
        self._tw(f"STOPPED | {datetime.now():%H:%M:%S} | Total: {self.stats['total']}\n", "hdr")
        self._tw(f"{'='*42}\n", "hdr")

    def _toggle_form(self, on):
        st = "readonly" if on else "disabled"
        for w in self.entries.values():
            if isinstance(w, ttk.Combobox): w.configure(state=st)

    def _restore_state(self):
        if self.running:
            self._toggle_form(False); self.btn_go.configure(state="disabled")
            self.btn_stp.configure(state="normal"); self.btn_chk.configure(state="disabled")
            self.lbl_st.config(text="RUNNING", foreground=C["red"])
            self.lbl_id.config(text=f"ID: {self.prod_id} (active)", foreground=C["red"])

    # ═══════════════════════════════════════════════════
    # MONITORING ENGINE
    # All socket operations happen ONLY in the monitoring thread.
    # No health_check on main thread touching sockets (was causing race).
    # ═══════════════════════════════════════════════════

    def _conn_w(self):
        self._disc_w()
        try:
            self.w_sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.w_sock.settimeout(5)
            self.w_sock.connect((self.weigher_ip, self.weigher_port))
            self.w_sock.settimeout(0.5)
            self.w_reader = SocketLineReader(self.w_sock)
            self.log.info(f"Weigher connected: {self.weigher_ip}:{self.weigher_port}")
            return True
        except Exception as e:
            self.log.error(f"Weigher fail: {e}"); self._disc_w(); return False

    def _disc_w(self):
        self.w_reader = None
        try:
            if self.w_sock: self.w_sock.close()
        except Exception: pass
        self.w_sock = None

    def _conn_m(self):
        if not self.metal_ip: return False
        self._disc_m()
        try:
            self.m_sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.m_sock.settimeout(10)
            self.m_sock.connect((self.metal_ip, self.metal_port))
            self.m_sock.settimeout(0.5)
            self.m_reader = SocketLineReader(self.m_sock)
            self.log.info(f"Metal connected: {self.metal_ip}:{self.metal_port}")
            return True
        except Exception as e:
            self.log.error(f"Metal fail: {e}"); self._disc_m(); return False

    def _disc_m(self):
        self.m_reader = None
        try:
            if self.m_sock: self.m_sock.close()
        except Exception: pass
        self.m_sock = None

    def _parse_weigher(self, line):
        if not line: return None, None
        if line.startswith('ANR'):
            ts_m = re.search(r'(\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2})', line)
            wt_m = re.search(r'(\d{4,6})\x01\x01', line)
            if ts_m and wt_m:
                try:
                    return int(wt_m.group(1)), datetime.strptime(ts_m.group(1), "%Y-%m-%dT%H:%M:%S")
                except ValueError: pass
        return None, None

    def _parse_metal(self, line):
        if not line: return None, None, None
        if line.strip() in ('0', '1'):
            st = int(line.strip())
            return st, 0 if st == 0 else METAL_THRESHOLD, datetime.now()
        parts = line.split(' - ', 1)
        if len(parts) == 2:
            num_match = re.search(r'(\d+)', parts[1].strip())
            if num_match:
                first_value = int(num_match.group(1))
                metal_status = 1 if first_value >= METAL_THRESHOLD else 0
                return metal_status, first_value, datetime.now()
        return None, None, None

    def _calc_status(self, weight):
        sd = self.session_data or {}
        u, o = sd.get("under_limit", 25025), sd.get("over_limit", 25175)
        if weight < u: return 0
        if weight > o: return 2
        return 1

    def _start_mon(self):
        for attempt in range(3):
            if self._conn_w(): break
            self.log.warning(f"Weigher attempt {attempt + 1} failed")
            if attempt < 2: time.sleep(2)
        else:
            messagebox.showerror("Connect", f"Weigher failed:\n{self.weigher_ip}:{self.weigher_port}"); return False
        if self.metal_ip:
            for attempt in range(2):
                if self._conn_m(): break
                if attempt < 1: time.sleep(1)
            else: messagebox.showwarning("Metal", "Metal detector failed. Continuing without.")
        self.w_queue.clear(); self.m_queue.clear(); self.db_ops.clear()
        self.last_match_debug = self.last_cleanup = datetime.now()
        try: self.next_log = self.db.call_sp("sp_GetNextLogId", fetch=True)[0][0]
        except Exception: self.next_log = 1
        self.stop_evt.clear()
        self.thread = threading.Thread(target=self._mon_loop, daemon=True, name="ProductionMonitor")
        self.thread.start(); self.log.info("Monitoring thread started"); return True

    def _stop_mon(self):
        self.stop_evt.set()
        # Wait for thread to finish
        if self.thread and self.thread.is_alive():
            self.thread.join(timeout=3)
        # Final flush
        self._match(); self._flush()
        self._disc_w(); self._disc_m()
        self.w_queue.clear(); self.m_queue.clear(); self.db_ops.clear()

    def _mon_loop(self):
        """Main monitoring loop. ALL socket operations happen here only.
        No other thread touches the sockets (removed health_check race)."""
        self.log.info("Monitoring loop started...")
        last_w_reconnect = datetime.now()
        last_m_reconnect = datetime.now()

        while not self.stop_evt.is_set():
            try:
                # --- Read weigher ---
                if self.w_sock and self.w_reader:
                    try:
                        line = self.w_reader.readline()
                        weight, machine_ts = self._parse_weigher(line)
                        if weight is not None:
                            now_ts = datetime.now()
                            self.w_queue.append({
                                'weight': weight, 'status': self._calc_status(weight),
                                'timestamp': machine_ts or now_ts,  # DB gets machine time
                                'queue_time': now_ts,               # Matching uses wall clock
                                'log_id': self.next_log})
                            self.next_log += 1
                            self.log.info(f"W: {weight}g at {(machine_ts or now_ts):%H:%M:%S} (Q:{len(self.w_queue)})")
                    except socket.timeout:
                        pass  # Normal - no data right now
                    except (ConnectionError, OSError) as e:
                        self.log.warning(f"W lost: {e}")
                        self._disc_w()
                else:
                    # Reconnect weigher (inside this thread, no race condition)
                    now = datetime.now()
                    if (now - last_w_reconnect).total_seconds() >= 5:
                        self.log.info("W reconnecting...")
                        self._conn_w()
                        last_w_reconnect = now

                # --- Read metal detector ---
                if self.metal_ip:
                    if self.m_sock and self.m_reader:
                        try:
                            line = self.m_reader.readline()
                            result = self._parse_metal(line)
                            if result[0] is not None:
                                metal_status, metal_value, ts = result
                                self.m_queue.append({
                                    'status': metal_status, 'value': metal_value, 'timestamp': ts})
                                self.log.info(f"M: status={metal_status} value={metal_value} (Q:{len(self.m_queue)})")
                        except socket.timeout:
                            pass  # Normal - simulator has travel delay
                        except (ConnectionError, OSError) as e:
                            self.log.warning(f"M lost: {e}")
                            self._disc_m()
                    else:
                        now = datetime.now()
                        if (now - last_m_reconnect).total_seconds() >= 5:
                            self.log.info("M reconnecting...")
                            self._conn_m()
                            last_m_reconnect = now

                # --- Match & process ---
                self._match()

                now = datetime.now()
                if (now - self.last_cleanup).total_seconds() >= 10:
                    self._cleanup_old(); self.last_cleanup = now
                if (now - self.last_match_debug).total_seconds() > 30:
                    self.log.info(f"Queues: W={len(self.w_queue)} M={len(self.m_queue)}")
                    self.last_match_debug = now

                time.sleep(0.01)
            except Exception as e:
                self.log.error(f"Mon error: {e}"); time.sleep(0.1)

    def _match(self):
        matches = 0
        if self.metal_ip:
            while self.w_queue and self.m_queue:
                wr, mr = self.w_queue.pop(0), self.m_queue.pop(0)
                self.db_ops.append((wr, mr))
                self.root.after(0, self._disp, wr, mr); matches += 1
        else:
            while self.w_queue:
                wr = self.w_queue.pop(0)
                mr = {'status': 0, 'value': None, 'timestamp': wr['timestamp']}
                self.db_ops.append((wr, mr))
                self.root.after(0, self._disp, wr, mr); matches += 1
        if matches > 0:
            self.log.info(f"Matched {matches} reading(s)"); self._flush()

    def _cleanup_old(self):
        """Clean up old readings. NEVER force-match during active production.
        The original code explicitly skips cleanup while running - the FIFO
        matching will naturally pair weigher+metal when the metal arrives."""
        now = datetime.now()

        # During active production: only log queue sizes, NEVER force-match.
        # Every bag eventually passes through the metal detector, so the
        # FIFO queue will pair them naturally.
        if self.running:
            if len(self.w_queue) > 10 or len(self.m_queue) > 10:
                self.log.info(f"In-transit bags: W={len(self.w_queue)} M={len(self.m_queue)}")
            return

        # Only cleanup when production is STOPPED
        for name, q in [("W", self.w_queue), ("M", self.m_queue)]:
            before = len(q)
            q[:] = [r for r in q if (now - r.get('queue_time', r['timestamp'])).total_seconds() <= 3600]
            rem = before - len(q)
            if rem > 0: self.log.info(f"Cleaned {rem} old {name} readings")

    def _flush(self):
        for wr, mr in self.db_ops:
            try:
                self.db.call_sp("sp_InsertLogEntry", [
                    wr['log_id'], self.prod_id, wr['timestamp'],
                    wr['weight'], wr['status'],
                    mr['status'], mr.get('value'),
                ])
            except Exception as e: self.log.error(f"DB flush: {e}")
        self.db_ops.clear()

    def _disp(self, wr, mr):
        w, st, ts, lid = wr['weight'], wr['status'], wr['timestamp'], wr['log_id']
        ms, mv = mr['status'], mr.get('value')
        self.stats["total"] += 1

        # Metal detected = always fail, regardless of weight
        if ms == 1:
            self.stats["metal"] += 1
            self._tw(f"[{ts:%H:%M:%S}] #{lid:04d}  {w:>6d}g  FAIL   METAL!({mv})\n", "metal")
        elif st == 0:
            self.stats["under"] += 1
            mv_str = f"({mv})" if mv is not None else ""
            self._tw(f"[{ts:%H:%M:%S}] #{lid:04d}  {w:>6d}g  UNDER  OK{mv_str}\n", "under")
        elif st == 2:
            self.stats["over"] += 1
            mv_str = f"({mv})" if mv is not None else ""
            self._tw(f"[{ts:%H:%M:%S}] #{lid:04d}  {w:>6d}g  OVER   OK{mv_str}\n", "over")
        else:
            self.stats["passed"] += 1
            mv_str = f"({mv})" if mv is not None else ""
            self._tw(f"[{ts:%H:%M:%S}] #{lid:04d}  {w:>6d}g  PASS   OK{mv_str}\n", "pass")

        self.live_data.append(dict(ts=ts, weight=w, status=st, metal=ms, metal_value=mv, lid=lid))
        if len(self.live_data) > 2000: self.live_data = self.live_data[-2000:]
        self._update_stats()

    def _update_stats(self):
        s = self.stats
        for k in ("total", "passed", "under", "over", "metal"):
            if k in self.st_lbl: self.st_lbl[k].config(text=str(s[k]))
        rate = f"{s['passed'] / s['total'] * 100:.1f}%" if s["total"] else "0%"
        if "rate" in self.st_lbl: self.st_lbl["rate"].config(text=rate)

    # ═══════════════ LIFECYCLE ═══════════════
    def _logout(self):
        if self.running:
            if not messagebox.askyesno("Running", "Stop production and logout?"): return
            self._stop_mon(); self.running = False
        if messagebox.askyesno("Logout", "Logout?"):
            self.user = self.prod_id = None; self._show_login()

    def _on_close(self):
        if self.running: self._stop_mon(); self.running = False
        self.db.close(); self.root.destroy()


def main():
    root = tk.Tk(); ProductionApp(root); root.mainloop()

if __name__ == "__main__":
    main()