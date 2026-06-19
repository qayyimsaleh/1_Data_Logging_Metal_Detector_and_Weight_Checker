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
from collections import deque
from datetime import datetime, timedelta

from shared_config import (
    APP_TITLE, APP_VERSION, COLORS, DB, make_logger,
    get_local_ip, get_dropdown, add_dropdown,
    validate_ip, validate_port, sanitize, hash_password,
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
        self._start_time = datetime.now()
        self.log = make_logger("production")
        self.db = DB(self.log)
        if not self.db.connect():
            messagebox.showerror("DB Error", f"Cannot connect to database.\n\n{getattr(self.db, 'last_error', '')}")
        self.pc_ip = get_local_ip()
        self.machine_name = self.pc_ip
        self.app_session_id = None
        try:
            r = self.db.call_sp("sp_AppStarted",
                                [self.pc_ip, socket.gethostname(), APP_VERSION],
                                fetch=True)
            self.app_session_id = r[0][0] if r else None
            self.log.info(f"App uptime tracking started (session_id={self.app_session_id})")
        except Exception as e:
            self.log.warning(f"App uptime tracking unavailable: {e}")
        self.weigher_ip = "192.168.0.100"
        self.weigher_port = 50001
        self.metal_ip = None
        self.metal_port = 50001
        self.w_sock = None; self.w_reader = None
        self.m_sock = None; self.m_reader = None
        self.running = False; self.thread = None
        self.stop_evt = threading.Event()
        self.pause_evt = threading.Event()  # Set = paused, Clear = running
        self.user = self.prod_id = self.session_data = self.screen = None
        self.entries = {}
        self.w_queue = deque(); self.m_queue = deque(); self.db_ops = deque()
        self.next_log = None  # unused; IDs fetched per-bag via sp_GetNextLogId
        self.stats = dict(total=0, passed=0, under=0, over=0, metal=0)
        self.bag_seq = 0
        self.wc_failed = None
        self._wc_shown = set()
        self._md_shown = set()
        self.last_match_debug = datetime.now()
        self.last_cleanup = datetime.now()
        self.info_lot = self.info_target = self.info_range = None
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
        s.configure("Treeview", background=C["bg_input"], foreground=C["text"],
                    fieldbackground=C["bg_input"], borderwidth=0, rowheight=26)
        s.configure("Treeview.Heading", background=C["bg_card"], foreground=C["text2"],
                    font=("Segoe UI", 9, "bold"))
        s.map("Treeview", background=[("selected", C["accent"])],
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
        self._login_pass.grid(row=3, column=0, pady=(0, 10), ipady=4)
        machines = []
        try:
            r = self.db.call_sp("sp_GetAllMachines", fetch=True)
            machines = [row[0] for row in r] if r else []
        except Exception:
            pass
        self._login_machine = ttk.Combobox(f, width=24, font=("Segoe UI", 11), style="Dark.TCombobox",
                                            state="readonly", values=machines)
        if machines:
            ttk.Label(f, text="Line / Machine", style="Norm.TLabel").grid(row=4, column=0, sticky="w", pady=(0, 3))
            if self.machine_name and self.machine_name in machines:
                self._login_machine.set(self.machine_name)
            else:
                self._login_machine.set(machines[0])
            self._login_machine.grid(row=5, column=0, pady=(0, 14), ipady=4)
        ttk.Button(card, text="Sign In", command=self._do_login, style="Accent.TButton").pack(fill="x")
        ttk.Label(m, text=f"v{APP_VERSION}", style="Sub.TLabel").pack(side="bottom", pady=10)
        self.root.bind("<Return>", lambda e: self._do_login())

    def _apply_machine(self, name: str):
        try:
            r = self.db.call_sp("sp_GetMachineByName", [name], fetch=True)
            if r:
                self.machine_name = r[0][0]
                self.weigher_ip = r[0][1]
                self.metal_ip = r[0][2] if len(r[0]) > 2 else None
                self.log.info(f"Machine set: {self.machine_name} (WC:{self.weigher_ip}, MD:{self.metal_ip})")
        except Exception as e:
            self.log.warning(f"Apply machine failed: {e}")
            self.machine_name = name
        if self.app_session_id:
            try:
                self.db.call_sp("sp_AppSetMachine", [self.app_session_id, self.machine_name])
            except Exception as e:
                self.log.warning(f"App uptime machine update failed: {e}")

    def _do_login(self):
        u, p = sanitize(self._login_user.get(), 50), self._login_pass.get()
        selected = self._login_machine.get().strip()
        if not u or not p: messagebox.showwarning("Login", "Enter both fields."); return
        if not selected and self._login_machine["values"]:
            messagebox.showwarning("Login", "Select a line / machine."); return
        try:
            r = self.db.call_sp("sp_VerifyUser", [u, hash_password(p)], fetch=True)
            if r:
                self.user = u
                if selected:
                    self._apply_machine(selected)
                self.root.unbind("<Return>"); self._show_main()
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
        uptime_h = (datetime.now() - self._start_time).total_seconds() / 3600
        if uptime_h >= 12:
            ttk.Label(btm, text=f"App running {uptime_h:.0f}h — restart at next shift change.",
                      font=("Segoe UI", 9), background=C["bg_dark"],
                      foreground=C["orange"]).pack(side="left")
        ttk.Button(btm, text="Logout", command=self._logout, style="Ghost.TButton").pack(side="right")

    def _launch_reports(self):
        import subprocess, sys
        if getattr(sys, 'frozen', False):
            # PyInstaller exe — look for report.exe alongside main exe
            d = os.path.dirname(sys.executable)
            rp = os.path.join(d, "report.exe")
            if os.path.exists(rp):
                subprocess.Popen([rp])
            else:
                messagebox.showwarning("Not Found",
                    "report.exe not found next to production.exe.\n"
                    "Build report.py separately with PyInstaller.")
        else:
            d = os.path.dirname(os.path.abspath(__file__))
            rp = os.path.join(d, "report.py")
            if os.path.exists(rp):
                subprocess.Popen([sys.executable, rp])
            else:
                messagebox.showwarning("Not Found", f"report.py not found in:\n{d}")

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
        ttk.Button(bf, text="Manage Lines", command=self._show_lines, style="Ghost.TButton").pack(side="left", padx=(0, 8))
        ttk.Button(bf, text="Back", command=self._show_main, style="Ghost.TButton").pack(side="right")

    def _test_conn(self, ip_e, port_e, lbl):
        ip, pt = ip_e.get().strip(), validate_port(port_e.get())
        if not ip: lbl.config(text="No IP", foreground=C["text3"]); return
        if not validate_ip(ip): lbl.config(text="Bad IP", foreground=C["red"]); return
        if not pt: lbl.config(text="Bad port", foreground=C["red"]); return
        lbl.config(text="...", foreground=C["orange"])
        def _do():
            try:
                s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                s.settimeout(5); s.connect((ip, pt)); s.close()
                self.root.after(0, lambda: lbl.config(text="OK", foreground=C["green"]))
            except Exception:
                self.root.after(0, lambda: lbl.config(text="Fail", foreground=C["red"]))
        threading.Thread(target=_do, daemon=True).start()

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
        else:
            self.metal_ip = None
        try:
            self.db.call_sp("sp_UpdateMachineByName", [self.machine_name, self.pc_ip, self.weigher_ip, self.metal_ip])
            self.log.info(f"Machine config saved: {self.machine_name} weigher={self.weigher_ip} metal={self.metal_ip}")
        except Exception as e:
            self.log.warning(f"Config DB save failed (in-memory only): {e}")
        self._show_main()

    def _show_lines(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("Manage Lines / Machines")
        dlg.configure(bg=C["bg_dark"])
        dlg.geometry("640x520")
        dlg.resizable(False, False)
        dlg.grab_set()

        ttk.Label(dlg, text="Lines / Machines", font=("Segoe UI", 13, "bold"),
                  background=C["bg_dark"], foreground=C["text"]).pack(pady=(14, 4), padx=20, anchor="w")
        tk.Frame(dlg, height=1, bg=C["border"]).pack(fill="x", padx=20)

        tv_f = ttk.Frame(dlg, style="Dark.TFrame"); tv_f.pack(fill="both", expand=True, padx=20, pady=(10, 0))
        cols = ("machine", "pc_ip", "weigher_ip", "metal_ip")
        tv = ttk.Treeview(tv_f, columns=cols, show="headings", height=8)
        for col, hdr, w in [("machine","Machine",130), ("pc_ip","PC IP",120),
                             ("weigher_ip","Weigher IP",130), ("metal_ip","Metal IP",130)]:
            tv.heading(col, text=hdr); tv.column(col, width=w, anchor="w")
        sb = ttk.Scrollbar(tv_f, orient="vertical", command=tv.yview)
        tv.configure(yscrollcommand=sb.set)
        tv.pack(side="left", fill="both", expand=True); sb.pack(side="right", fill="y")

        def _refresh():
            tv.delete(*tv.get_children())
            try:
                r = self.db.call_sp("sp_GetAllMachinesDetail", fetch=True)
                for row in (r or []):
                    tv.insert("", "end", values=(row[0], row[1], row[2], row[3] or ""))
            except Exception as e:
                self.log.warning(f"Lines load: {e}")
        _refresh()

        form_f = ttk.Frame(dlg, style="Card.TFrame"); form_f.pack(fill="x", padx=20, pady=10)
        form_f.configure(padding=12)
        flds = {}
        for i, (key, label) in enumerate([("machine","Machine"), ("pc_ip","PC IP"),
                                           ("weigher_ip","Weigher IP"), ("metal_ip","Metal IP (opt)")]):
            ttk.Label(form_f, text=label, style="Norm.TLabel").grid(row=0, column=i, sticky="w", padx=(0, 4))
            e = ttk.Entry(form_f, width=14, font=("Consolas", 10), style="Dark.TEntry")
            e.grid(row=1, column=i, padx=(0, 6), ipady=2)
            flds[key] = e

        _selected_name = [None]

        def _on_select(event):
            sel = tv.focus()
            if not sel: return
            vals = tv.item(sel, "values")
            _selected_name[0] = vals[0]
            for k, v in zip(("machine", "pc_ip", "weigher_ip", "metal_ip"), vals):
                flds[k].config(state="normal")
                flds[k].delete(0, tk.END)
                flds[k].insert(0, v)
            flds["machine"].config(state="disabled")
        tv.bind("<<TreeviewSelect>>", _on_select)

        def _clear_form():
            _selected_name[0] = None
            for e in flds.values():
                e.config(state="normal"); e.delete(0, tk.END)

        def _get():
            return {k: sanitize(e.get() if e["state"] != "disabled" else _selected_name[0] or "", 50)
                    for k, e in flds.items()}

        def _add():
            v = _get()
            if not v["machine"] or not v["pc_ip"] or not v["weigher_ip"]:
                messagebox.showwarning("Validation", "Machine, PC IP, and Weigher IP are required.", parent=dlg); return
            if not validate_ip(v["pc_ip"]): messagebox.showwarning("Validation", "Invalid PC IP.", parent=dlg); return
            if not validate_ip(v["weigher_ip"]): messagebox.showwarning("Validation", "Invalid Weigher IP.", parent=dlg); return
            if v["metal_ip"] and not validate_ip(v["metal_ip"]): messagebox.showwarning("Validation", "Invalid Metal IP.", parent=dlg); return
            try:
                self.db.call_sp("sp_AddMachine", [v["machine"], v["pc_ip"], v["weigher_ip"], v["metal_ip"] or None])
                _clear_form(); _refresh()
            except Exception as e:
                messagebox.showerror("Error", str(e), parent=dlg)

        def _update():
            if not _selected_name[0]: messagebox.showwarning("Select", "Select a machine first.", parent=dlg); return
            v = _get()
            if not v["pc_ip"] or not v["weigher_ip"]: messagebox.showwarning("Validation", "PC IP and Weigher IP are required.", parent=dlg); return
            if not validate_ip(v["pc_ip"]): messagebox.showwarning("Validation", "Invalid PC IP.", parent=dlg); return
            if not validate_ip(v["weigher_ip"]): messagebox.showwarning("Validation", "Invalid Weigher IP.", parent=dlg); return
            if v["metal_ip"] and not validate_ip(v["metal_ip"]): messagebox.showwarning("Validation", "Invalid Metal IP.", parent=dlg); return
            try:
                self.db.call_sp("sp_UpdateMachineByName", [_selected_name[0], v["pc_ip"], v["weigher_ip"], v["metal_ip"] or None])
                _clear_form(); _refresh()
            except Exception as e:
                messagebox.showerror("Error", str(e), parent=dlg)

        def _delete():
            if not _selected_name[0]: messagebox.showwarning("Select", "Select a machine first.", parent=dlg); return
            if not messagebox.askyesno("Confirm", f"Delete '{_selected_name[0]}'?", parent=dlg): return
            try:
                self.db.call_sp("sp_DeleteMachine", [_selected_name[0]])
                _clear_form(); _refresh()
            except Exception as e:
                messagebox.showerror("Error", str(e), parent=dlg)

        btn_f = ttk.Frame(dlg, style="Dark.TFrame"); btn_f.pack(fill="x", padx=20, pady=(0, 16))
        ttk.Button(btn_f, text="Add", command=_add, style="Accent.TButton").pack(side="left", padx=(0, 6))
        ttk.Button(btn_f, text="Update", command=_update, style="Green.TButton").pack(side="left", padx=(0, 6))
        ttk.Button(btn_f, text="Delete", command=_delete, style="Red.TButton").pack(side="left", padx=(0, 6))
        ttk.Button(btn_f, text="Clear", command=_clear_form, style="Ghost.TButton").pack(side="left")
        ttk.Button(btn_f, text="Close", command=dlg.destroy, style="Ghost.TButton").pack(side="right")

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
        # ── Last Bag Status ──────────────────────────────────
        lb_f = tk.Frame(parent, bg=C["bg_card"], padx=14, pady=10)
        lb_f.pack(fill="x", pady=(0, 5))
        hdr_f = tk.Frame(lb_f, bg=C["bg_card"]); hdr_f.pack(fill="x")
        tk.Label(hdr_f, text="LAST BAG", font=("Segoe UI", 8, "bold"),
                 bg=C["bg_card"], fg=C["text3"]).pack(side="left")
        self.lbl_bag_time = tk.Label(hdr_f, text="", font=("Segoe UI", 9),
                                      bg=C["bg_card"], fg=C["text3"])
        self.lbl_bag_time.pack(side="right")
        body_f = tk.Frame(lb_f, bg=C["bg_card"]); body_f.pack(fill="x", pady=(6, 0))
        lf = tk.Frame(body_f, bg=C["bg_card"]); lf.pack(side="left")
        self.lbl_bag_status = tk.Label(lf, text="—", font=("Segoe UI", 32, "bold"),
                                        bg=C["bg_card"], fg=C["text3"])
        self.lbl_bag_status.pack(anchor="w")
        self.lbl_bag_weight = tk.Label(lf, text="", font=("Segoe UI", 14),
                                        bg=C["bg_card"], fg=C["text2"])
        self.lbl_bag_weight.pack(anchor="w")
        rf = tk.Frame(body_f, bg=C["bg_card"]); rf.pack(side="right", anchor="e")
        self.lbl_bag_metal = tk.Label(rf, text="", font=("Segoe UI", 11),
                                       bg=C["bg_card"], fg=C["text3"])
        self.lbl_bag_metal.pack(anchor="e")

        # ── Bag Counters ─────────────────────────────────────
        cnt_f = tk.Frame(parent, bg=C["bg_card"], padx=8, pady=8)
        cnt_f.pack(fill="x", pady=(0, 5))
        self.st_lbl = {}
        for l, k, clr in [
            ("TOTAL", "total", C["text"]), ("PASS", "passed", C["green"]),
            ("UNDER WT", "under", C["orange"]), ("OVER WT", "over", C["red"]),
            ("METAL", "metal", "#ff3333"), ("PASS RATE", "rate", C["blue"])]:
            col = tk.Frame(cnt_f, bg=C["bg_card"]); col.pack(side="left", expand=True)
            lv = tk.Label(col, text="0", font=("Segoe UI", 22, "bold"), bg=C["bg_card"], fg=clr)
            lv.pack()
            tk.Label(col, text=l, font=("Segoe UI", 8), bg=C["bg_card"], fg=C["text3"]).pack()
            self.st_lbl[k] = lv

        # ── Session Info ──────────────────────────────────────
        si = tk.Frame(parent, bg=C["bg_card"], padx=14, pady=5)
        si.pack(fill="x", pady=(0, 5))
        self.info_lot = ttk.Label(si, text="Lot: —", font=("Segoe UI", 9, "bold"),
                                   background=C["bg_card"], foreground=C["text3"])
        self.info_lot.pack(side="left", padx=(0, 20))
        self.info_target = ttk.Label(si, text="Target: —", font=("Segoe UI", 9),
                                      background=C["bg_card"], foreground=C["text3"])
        self.info_target.pack(side="left", padx=(0, 20))
        self.info_range = ttk.Label(si, text="Range: —", font=("Segoe UI", 9),
                                     background=C["bg_card"], foreground=C["text3"])
        self.info_range.pack(side="left")

        # ── Dual Log Panels ───────────────────────────────────
        log_row = tk.Frame(parent, bg=C["bg_dark"])
        log_row.pack(fill="both", expand=True)

        def _make_panel(parent, title):
            f = tk.Frame(parent, bg=C["bg_card"], padx=8, pady=8)
            f.pack(side="left", fill="both", expand=True, padx=(0, 2))
            tk.Label(f, text=title, font=("Segoe UI", 8, "bold"),
                     bg=C["bg_card"], fg=C["accent"]).pack(anchor="w", pady=(0, 3))
            tk.Frame(f, height=1, bg=C["border"]).pack(fill="x", pady=(0, 3))
            t = scrolledtext.ScrolledText(f, wrap=tk.WORD, font=("Consolas", 9),
                bg=C["terminal_bg"], fg=C["text2"], insertbackground=C["text2"],
                relief="flat", borderwidth=0)
            t.pack(fill="both", expand=True)
            for tag, color in [("pass", C["green"]), ("under", C["orange"]), ("over", C["red"]),
                               ("info", C["blue"]), ("hdr", C["accent"]), ("metal", "#ff3333")]:
                t.tag_configure(tag, foreground=color)
            return t

        self.term_wc = _make_panel(log_row, "WEIGHT CHECKER")
        self.term_md = _make_panel(log_row, "METAL DETECTOR")
        self._tw("Waiting for production start...\n", "info")
        self.term_wc.config(state="disabled")
        self.term_md.config(state="disabled")

    def _tw_panel(self, term, text, tag=None):
        term.config(state="normal")
        term.insert(tk.END, text, tag) if tag else term.insert(tk.END, text)
        lines = int(term.index("end-1c").split(".")[0])
        if lines > 500: term.delete("1.0", f"{lines - 500}.0")
        term.see(tk.END); term.config(state="disabled")

    def _tw(self, text, tag=None):
        self._tw_panel(self.term_wc, text, tag)
        self._tw_panel(self.term_md, text, tag)

    def _tw_wc(self, text, tag=None):
        self._tw_panel(self.term_wc, text, tag)

    def _tw_md(self, text, tag=None):
        self._tw_panel(self.term_md, text, tag)

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
            self.running = True; self.stats = dict(total=0, passed=0, under=0, over=0, metal=0)
            self._toggle_form(False); self.btn_go.configure(state="disabled")
            self.btn_stp.configure(state="normal"); self.btn_chk.configure(state="disabled")
            self.lbl_st.config(text="RUNNING", foreground=C["red"])
            if self.info_lot:
                self.info_lot.config(text=f"Lot: {self.session_data['lot_no']}", foreground=C["accent"])
                self.info_target.config(text=f"Target: {self.session_data['net_weight']}g", foreground=C["text2"])
                self.info_range.config(text=f"Range: {self.session_data['under_limit']} - {self.session_data['over_limit']}g", foreground=C["text2"])
            self._tw(f"\n{'─'*54}\n", "hdr")
            self._tw(f"  Session {self.prod_id}  started  {dt:%H:%M:%S}  —  {self.machine_name}\n", "hdr")
            self._tw(f"  Lot: {self.session_data['lot_no']}\n", "info")
            self._tw(f"{'─'*54}\n\n", "hdr"); self._update_stats()
        except Exception as e:
            self.log.error(f"Start error: {e}"); messagebox.showerror("Error", str(e))

    def _stop(self, confirm=True):
        if not self.running: return
        if confirm and not messagebox.askyesno("Confirm", "Stop production?"): return
        self._stop_mon(); self.running = False; self._toggle_form(True)
        self.btn_go.configure(state="normal"); self.btn_stp.configure(state="disabled"); self.btn_chk.configure(state="normal")
        self.lbl_st.config(text="STOPPED", foreground=C["text3"])
        end_dt = datetime.now()
        try:
            self.db.call_sp("sp_EndProductionSession", [self.prod_id, end_dt, sanitize(self.user, 10)])
        except Exception as e:
            self.log.warning(f"Session end record failed: {e}")
        self._tw(f"\n{'─'*54}\n", "hdr")
        self._tw(f"  Session ended  {end_dt:%H:%M:%S}  —  {self.stats['total']} bags total\n", "hdr")
        self._tw(f"{'─'*54}\n", "hdr")
        if self.info_lot:
            self.info_lot.config(text="Lot: —", foreground=C["text3"])
            self.info_target.config(text="Target: —", foreground=C["text3"])
            self.info_range.config(text="Range: —", foreground=C["text3"])

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
            if self.info_lot and self.session_data:
                self.info_lot.config(text=f"Lot: {self.session_data['lot_no']}", foreground=C["accent"])
                self.info_target.config(text=f"Target: {self.session_data['net_weight']}g", foreground=C["text2"])
                self.info_range.config(text=f"Range: {self.session_data['under_limit']} - {self.session_data['over_limit']}g", foreground=C["text2"])

    # ═══════════════════════════════════════════════════
    # MONITORING ENGINE
    # All socket operations happen ONLY in the monitoring thread.
    # No health_check on main thread touching sockets (was causing race).
    # ═══════════════════════════════════════════════════

    @staticmethod
    def _set_keepalive(sock):
        try:
            sock.setsockopt(socket.SOL_SOCKET, socket.SO_KEEPALIVE, 1)
            # Windows: probe after 10s idle, every 3s, 5 probes before giving up
            sock.ioctl(socket.SIO_KEEPALIVE_VALS, (1, 10_000, 3_000))
        except (AttributeError, OSError):
            pass  # non-Windows or unavailable

    def _conn_w(self):
        self._disc_w()
        try:
            self.w_sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self._set_keepalive(self.w_sock)
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
            self._set_keepalive(self.m_sock)
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
        self.bag_seq = 0; self.wc_failed = None; self._wc_shown.clear(); self._md_shown.clear()
        self.last_match_debug = self.last_cleanup = datetime.now()
        self.pause_evt.clear()
        self.stop_evt.clear()
        self.thread = threading.Thread(target=self._mon_loop, daemon=True, name="ProductionMonitor")
        self.thread.start(); self.log.info("Monitoring thread started"); return True

    def _stop_mon(self):
        self.stop_evt.set()
        self.pause_evt.clear()  # Unblock thread if it's waiting in pause loop
        if self.thread and self.thread.is_alive():
            self.thread.join(timeout=3)
            if self.thread.is_alive():
                # Thread didn't stop in time — skip final flush to avoid race on db_ops
                self.log.warning("Monitor thread did not stop within 3s — skipping final flush")
                self._disc_w(); self._disc_m()
                return
        # Thread confirmed dead — safe to do final flush on main thread
        self._match(); self._flush()
        self._disc_w(); self._disc_m()
        self.w_queue.clear(); self.m_queue.clear(); self.db_ops.clear()

    def _mon_loop(self):
        """Main monitoring loop. ALL socket operations happen here only.
        No other thread touches the sockets (removed health_check race)."""
        self.log.info("Monitoring loop started...")
        last_w_reconnect = datetime.now()
        last_m_reconnect = datetime.now()
        last_w_data    = datetime.now()   # watchdog: last time WC sent real data
        last_m_data    = datetime.now()   # watchdog: last time MD sent real data
        last_db_refresh = datetime.now()  # proactive DB keepalive
        last_heartbeat  = datetime.now()  # hourly alive log
        SILENT_TIMEOUT = 300              # 5 min with no data → force reconnect
        DB_REFRESH_SECS = 1800            # refresh DB connection every 30 min
        HEARTBEAT_SECS  = 3600           # log heartbeat every hour

        while not self.stop_evt.is_set():
            try:
                if self.pause_evt.is_set():
                    time.sleep(0.1); continue
                # --- Read weigher ---
                if self.w_sock and self.w_reader:
                    try:
                        line = self.w_reader.readline()
                        weight, machine_ts = self._parse_weigher(line)
                        if weight is not None:
                            last_w_data = datetime.now()
                            now_ts = datetime.now()
                            st = self._calc_status(weight)
                            is_wc_retry = (self.wc_failed is not None)
                            if not is_wc_retry:
                                self.bag_seq += 1
                            try:
                                log_id = self.db.call_sp("sp_GetNextLogId", fetch=True)[0][0]
                            except Exception as e:
                                if not is_wc_retry:
                                    self.bag_seq -= 1  # undo increment — bag was never processed
                                self.log.error(f"GetNextLogId failed, bag skipped: {e}")
                                raise  # propagate to outer handler to sleep & retry
                            wr = {
                                'weight': weight, 'status': st,
                                'timestamp': machine_ts or now_ts,
                                'queue_time': now_ts,
                                'log_id': log_id,
                                'bag_seq': self.wc_failed['bag_seq'] if is_wc_retry else self.bag_seq}
                            self.root.after(0, self._disp_wc, dict(wr), is_wc_retry)
                            if st == 1:  # PASS (new or retry cleared)
                                self.wc_failed = None
                                if self.metal_ip:
                                    self.w_queue.append(wr)
                                    self.log.info(f"W: {weight}g PASS → queued for MD (Q:{len(self.w_queue)}){' (WC retry cleared)' if is_wc_retry else ''}")
                                else:
                                    mr = {'status': None, 'value': None, 'timestamp': wr['timestamp']}
                                    self.db_ops.append((wr, mr))
                                    self.root.after(0, self._disp_md, dict(wr), mr)
                                    self.log.info(f"W: {weight}g PASS → direct log{' (WC retry cleared)' if is_wc_retry else ''}")
                            else:  # OVER/UNDER — log immediately, stay in retry mode
                                mr = {'status': None, 'value': None, 'timestamp': wr['timestamp']}
                                self.db_ops.append((wr, mr))
                                self.root.after(0, self._disp_md, dict(wr), mr)
                                self.wc_failed = wr
                                s = 'UNDER' if st == 0 else 'OVER'
                                self.log.info(f"W: {weight}g {s} → direct log{' ↩ RETRY' if is_wc_retry else ''}")
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
                        if self._conn_w():
                            # Only clear MD queue — stale MD readings from before the disconnect
                            # must not pair with new WC bags after reconnect.
                            # w_queue is intentionally preserved: those bags are physically on
                            # the conveyor heading to MD and still need to be matched and logged.
                            self.m_queue.clear()
                            self.wc_failed = None
                            last_w_data = datetime.now()  # reset watchdog on reconnect
                            self.log.info("WC reconnected; MD queue cleared, WC transit queue preserved")
                        last_w_reconnect = now

                # --- Read metal detector ---
                if self.metal_ip:
                    if self.m_sock and self.m_reader:
                        try:
                            line = self.m_reader.readline()
                            result = self._parse_metal(line)
                            if result[0] is not None:
                                last_m_data = datetime.now()
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
                            if self._conn_m():
                                last_m_data = datetime.now()  # reset watchdog on reconnect
                            last_m_reconnect = now

                # --- Match & process ---
                self._match()

                now = datetime.now()
                if (now - self.last_cleanup).total_seconds() >= 10:
                    self._cleanup_old(); self.last_cleanup = now
                if (now - self.last_match_debug).total_seconds() > 30:
                    self.log.info(f"Queues: W={len(self.w_queue)} M={len(self.m_queue)}")
                    self.last_match_debug = now

                # --- Proactive DB keepalive & hourly heartbeat ---
                if (now - last_db_refresh).total_seconds() > DB_REFRESH_SECS:
                    try:
                        self.db.ensure()
                        self.log.info("DB keepalive OK")
                    except Exception as e:
                        self.log.warning(f"DB keepalive failed: {e}")
                    last_db_refresh = now
                if (now - last_heartbeat).total_seconds() > HEARTBEAT_SECS:
                    self.log.info(
                        f"Heartbeat | running={self.running} "
                        f"W={'up' if self.w_sock else 'down'} "
                        f"M={'up' if self.m_sock else 'N/A' if not self.metal_ip else 'down'} "
                        f"wq={len(self.w_queue)} mq={len(self.m_queue)}")
                    last_heartbeat = now

                # --- Silent-disconnect watchdog ---
                # If socket appears connected but no data in SILENT_TIMEOUT seconds,
                # the remote machine likely dropped without sending FIN (cable pulled etc.)
                if self.running and self.w_sock:
                    if (now - last_w_data).total_seconds() > SILENT_TIMEOUT:
                        self.log.warning(f"WC silent for >{SILENT_TIMEOUT}s — forcing reconnect")
                        self._disc_w(); last_w_data = now
                if self.running and self.metal_ip and self.m_sock:
                    if (now - last_m_data).total_seconds() > SILENT_TIMEOUT:
                        self.log.warning(f"MD silent for >{SILENT_TIMEOUT}s — forcing reconnect")
                        self._disc_m(); last_m_data = now

                time.sleep(0.01)
            except Exception as e:
                self.log.error(f"Mon error: {e}"); time.sleep(0.1)

    def _match(self):
        matches = 0
        if self.metal_ip:
            while self.w_queue and self.m_queue:
                wr, mr = self.w_queue.popleft(), self.m_queue.popleft()
                self.db_ops.append((wr, mr))
                self.root.after(0, self._disp_md, wr, mr)
                if mr['status'] == 1:
                    # MD failed → operator retries same bag through MD.
                    # Put weight back at front with a new log_id so each
                    # attempt gets its own audit row (same weight, new ID).
                    # NOTE: if other bags reach MD before the retry (long gap),
                    # those bags will consume this weight entry instead — FIFO
                    # aggregate counts stay correct even if per-bag pairing drifts.
                    retry_wr = dict(wr)
                    retry_wr['log_id'] = self.db.call_sp("sp_GetNextLogId", fetch=True)[0][0]
                    self.w_queue.appendleft(retry_wr)
                    self.log.info(f"MD fail log_id={wr['log_id']}, retry→log_id={retry_wr['log_id']}")
                self._flush()  # save immediately after each MD result
                matches += 1
        else:
            # Safety net: w_queue only has items here if metal_ip was cleared
            # mid-session; normally it's empty because _mon_loop logs directly.
            while self.w_queue:
                wr = self.w_queue.popleft()
                mr = {'status': None, 'value': None, 'timestamp': wr['timestamp']}
                self.db_ops.append((wr, mr))
                self.root.after(0, self._disp_md, wr, mr); matches += 1
        if matches > 0:
            self.log.info(f"Matched {matches} reading(s)")
        if self.db_ops:  # flush any remaining OVER/UNDER bags
            self._flush()

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
            keep = [r for r in q if (now - r.get('queue_time', r['timestamp'])).total_seconds() <= 3600]
            q.clear(); q.extend(keep)
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

    def _disp_wc(self, wr, is_retry=False):
        """Called immediately when WC reads a bag. Updates WC panel and last-bag card."""
        w, st, ts = wr['weight'], wr['status'], wr['timestamp']
        bseq = wr.get('bag_seq', 0)
        retry_sfx = "  ↩ RETRY" if is_retry else ""
        if not is_retry:
            self._wc_shown.add(bseq)
            self.stats["total"] += 1
        if st == 0:
            if not is_retry: self.stats["under"] += 1
            self._tw_wc(f"  {ts:%H:%M:%S}  #{bseq:04d}  {w:>8,} g  UNDER WEIGHT{retry_sfx}\n", "under")
            self.lbl_bag_status.config(text="UNDER WT", fg=C["orange"])
        elif st == 2:
            if not is_retry: self.stats["over"] += 1
            self._tw_wc(f"  {ts:%H:%M:%S}  #{bseq:04d}  {w:>8,} g  OVER WEIGHT{retry_sfx}\n", "over")
            self.lbl_bag_status.config(text="OVER WT", fg=C["red"])
        else:
            cleared_sfx = "  CLEARED" if is_retry else ""
            self._tw_wc(f"  {ts:%H:%M:%S}  #{bseq:04d}  {w:>8,} g  PASS{cleared_sfx}\n", "pass")
            self.lbl_bag_status.config(text="PASS", fg=C["green"])
        self.lbl_bag_weight.config(text=f"{w:,} g")
        self.lbl_bag_time.config(text=f"{ts:%H:%M:%S}")
        self.lbl_bag_metal.config(text="", fg=C["text3"])
        self._update_stats()

    def _disp_md(self, wr, mr):
        """Called when MD pairs with a WC bag. Updates MD panel only."""
        bseq = wr.get('bag_seq', 0)
        ms, mv = mr['status'], mr.get('value')
        ts = mr.get('timestamp', datetime.now())
        if ms == 1:
            self.stats["metal"] += 1
            sfx = "  ↩ RETRY" if bseq in self._md_shown else ""
            self._tw_md(f"  {ts:%H:%M:%S}  #{bseq:04d}  METAL ({mv}){sfx}\n", "metal")
            self.lbl_bag_status.config(text="METAL!", fg="#ff3333")
            self.lbl_bag_metal.config(text=f"METAL ({mv})", fg="#ff3333")
            self._md_shown.add(bseq)
        elif ms == 0:
            self.stats["passed"] += 1
            sfx = "  CLEARED" if bseq in self._md_shown else ""
            self._tw_md(f"  {ts:%H:%M:%S}  #{bseq:04d}  No Metal   PASS{sfx}\n", "pass")
            self.lbl_bag_status.config(text="PASS", fg=C["green"])
            self.lbl_bag_metal.config(text="No Metal", fg=C["green"])
            self._md_shown.discard(bseq)
        else:
            # no MD result (under/over bag, or no metal_ip configured)
            if not self.metal_ip:
                self.stats["passed"] += 1
        self._update_stats()

    def _update_stats(self):
        s = self.stats
        for k in ("total", "passed", "under", "over", "metal"):
            if k in self.st_lbl: self.st_lbl[k].config(text=str(s[k]))
        rate = f"{s['passed'] / s['total'] * 100:.0f}%" if s["total"] else "—"
        if "rate" in self.st_lbl: self.st_lbl["rate"].config(text=rate)

    # ═══════════════ LIFECYCLE ═══════════════
    def _logout(self):
        if self.running:
            if not messagebox.askyesno("Logout", "Stop production and logout?"): return
            self._stop_mon(); self.running = False
            try:
                self.db.call_sp("sp_EndProductionSession",
                                [self.prod_id, datetime.now(), sanitize(self.user, 10)])
            except Exception as e:
                self.log.warning(f"Session end record failed on logout: {e}")
            self.user = self.prod_id = None; self._show_login()
        elif messagebox.askyesno("Logout", "Logout?"):
            self.user = self.prod_id = None; self._show_login()

    def _on_close(self):
        if self.running:
            self._stop_mon(); self.running = False
            try:
                self.db.call_sp("sp_EndProductionSession",
                                [self.prod_id, datetime.now(), sanitize(self.user, 10)])
            except Exception: pass
        if self.app_session_id:
            try:
                self.db.call_sp("sp_AppStopped", [self.app_session_id])
                self.log.info(f"App uptime stamped (session_id={self.app_session_id})")
            except Exception: pass
        self.db.close(); self.root.destroy()


def _show_splash(root):
    s = tk.Toplevel(root)
    s.overrideredirect(True)
    s.configure(bg=C["bg_dark"])
    sw, sh = s.winfo_screenwidth(), s.winfo_screenheight()
    w, h = 440, 190
    s.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")
    s.lift(); s.attributes("-topmost", True)
    tk.Label(s, text=APP_TITLE, bg=C["bg_dark"], fg=C["text"],
             font=("Segoe UI", 14, "bold")).pack(pady=(40, 6))
    tk.Label(s, text=f"v{APP_VERSION}", bg=C["bg_dark"], fg=C["text2"],
             font=("Segoe UI", 9)).pack()
    tk.Label(s, text="Starting, please wait...", bg=C["bg_dark"], fg=C["text3"],
             font=("Segoe UI", 9)).pack(pady=(8, 0))
    pb = ttk.Progressbar(s, mode="indeterminate", length=320)
    pb.pack(pady=16)
    pb.start(10)
    s.update()
    return s


def main():
    root = tk.Tk()
    root.withdraw()
    splash = _show_splash(root)
    ProductionApp(root)
    splash.destroy()
    root.deiconify()
    root.mainloop()

if __name__ == "__main__":
    main()