"""
report.py - Report & Analytics System
==========================================
Standalone reporting: filters -> generate -> view -> export (Excel/PDF)

FIXES FROM ORIGINAL:
 - Excel "Shape mismatch" error: summary rows were raw Row objects not unpacked
 - Excel summary had 16 col headers but SP returns 17 cols (metal_fail_count)
 - PDF was empty: only header built, never the data table
 - PDF now has full paginated tables + statistics summary
 - Added deduplication for detailed reports

Run: python report.py
"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
from datetime import datetime

try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False

try:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                     Paragraph, Spacer, PageBreak)
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    from reportlab.lib.units import mm
    HAS_REPORTLAB = True
except ImportError:
    HAS_REPORTLAB = False

from shared_config import (
    APP_TITLE, APP_VERSION, COLORS, DB, make_logger, sanitize,
)

C = COLORS

# Column definitions matching sp_GetProductionReport exactly
DETAILED_COLS = [
    "Log ID", "Production ID", "Timestamp", "Weight (g)", "Status",
    "Metal Status", "Lot No", "Product", "Shift", "Batch No",
    "Net Weight", "Under Limit", "Over Limit", "Machine"
]

# SP returns 17 columns for summary (includes metal_fail_count)
SUMMARY_COLS = [
    "Production ID", "Machine", "Lot No", "Product", "Shift", "Batch No",
    "First Reading", "Last Reading", "Total Readings", "Pass Count",
    "Under Fail", "Over Fail", "Metal Fail", "Pass Rate %",
    "Min Weight", "Max Weight", "Avg Weight"
]


class ReportApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"{APP_TITLE} - Reports v{APP_VERSION}")
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.log = make_logger("reports")
        self.db = DB(self.log)
        if not self.db.connect():
            messagebox.showerror("DB Error", "Cannot connect to database.")
        self.report_data = None
        self.report_type = None
        self.sort_col = None
        self.sort_rev = False

        sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
        w, h = min(1500, sw - 40), min(900, sh - 60)
        root.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")
        root.minsize(1100, 650)
        root.configure(bg=C["bg_dark"])
        self._styles()
        self._build_ui()

    def _styles(self):
        s = ttk.Style()
        s.theme_use("clam")
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
            ("Orange.TButton", C["orange"], "#1a1a2e"),
        ]:
            s.configure(n, background=bg, foreground=fg,
                        font=("Segoe UI", 10, "bold"),
                        borderwidth=0, padding=(14, 7), focuscolor="none")
            s.map(n, background=[("active", C["accent_hover"])])
        s.configure("Dark.TEntry", fieldbackground=C["bg_input"],
                    foreground=C["text"], borderwidth=1, insertcolor=C["text"])
        s.configure("Dark.TCombobox", fieldbackground=C["bg_input"],
                    foreground=C["text"])
        s.map("Dark.TCombobox",
              fieldbackground=[("readonly", C["bg_input"])])
        s.configure("Report.Treeview",
                    background=C["bg_card"], foreground=C["text"],
                    fieldbackground=C["bg_card"],
                    font=("Segoe UI", 9), rowheight=26, borderwidth=0)
        s.configure("Report.Treeview.Heading",
                    background=C["bg_mid"], foreground=C["text"],
                    font=("Segoe UI", 9, "bold"), borderwidth=0)
        s.map("Report.Treeview",
              background=[("selected", C["accent"])],
              foreground=[("selected", "#fff")])

    def _build_ui(self):
        m = ttk.Frame(self.root, style="Dark.TFrame")
        m.pack(fill="both", expand=True)
        # Header bar
        h = ttk.Frame(m, style="Mid.TFrame")
        h.pack(fill="x")
        hi = ttk.Frame(h, style="Mid.TFrame")
        hi.pack(fill="x", padx=20, pady=10)
        ttk.Label(hi, text="Production Reports & Analytics",
                  font=("Segoe UI", 16, "bold"),
                  background=C["bg_mid"],
                  foreground=C["text"]).pack(side="left")
        ttk.Label(hi, text=f"v{APP_VERSION}",
                  font=("Segoe UI", 9),
                  background=C["bg_mid"],
                  foreground=C["text3"]).pack(side="right")
        tk.Frame(m, height=1, bg=C["border"]).pack(fill="x")

        ct = ttk.Frame(m, style="Dark.TFrame")
        ct.pack(fill="both", expand=True, padx=12, pady=8)
        pw = ttk.PanedWindow(ct, orient=tk.HORIZONTAL)
        pw.pack(fill="both", expand=True)

        left = ttk.Frame(pw, style="Dark.TFrame")
        pw.add(left, weight=1)
        right = ttk.Frame(pw, style="Dark.TFrame")
        pw.add(right, weight=4)
        self._build_filters(left)
        self._build_results(right)

    def _build_filters(self, parent):
        card = ttk.Frame(parent, style="Card.TFrame", padding=14)
        card.pack(fill="both", expand=True, padx=(0, 4))
        ttk.Label(card, text="Filters", style="Card2.TLabel").pack(
            anchor="w", pady=(0, 8))
        tk.Frame(card, height=1, bg=C["border"]).pack(fill="x", pady=(0, 10))

        ttk.Label(card, text="Machine", style="Norm.TLabel").pack(
            anchor="w", pady=(0, 3))
        self.f_machine = ttk.Combobox(card, state="readonly",
                                       style="Dark.TCombobox")
        try:
            machines = [r[0] for r in
                        (self.db.call_sp("sp_GetAllMachines", fetch=True) or [])]
        except Exception:
            machines = []
        self.f_machine["values"] = ["All Machines"] + machines
        self.f_machine.set("All Machines")
        self.f_machine.pack(fill="x", pady=(0, 10))

        ttk.Label(card, text="From", style="Norm.TLabel").pack(
            anchor="w", pady=(0, 3))
        self.f_from = ttk.Entry(card, style="Dark.TEntry",
                                font=("Segoe UI", 10))
        self.f_from.insert(0, datetime.now().strftime("%Y-%m-01"))
        self.f_from.pack(fill="x", ipady=2, pady=(0, 8))

        ttk.Label(card, text="To", style="Norm.TLabel").pack(
            anchor="w", pady=(0, 3))
        self.f_to = ttk.Entry(card, style="Dark.TEntry",
                              font=("Segoe UI", 10))
        self.f_to.insert(0, datetime.now().strftime("%Y-%m-%d"))
        self.f_to.pack(fill="x", ipady=2, pady=(0, 8))

        ttk.Label(card, text="Lot Number", style="Norm.TLabel").pack(
            anchor="w", pady=(0, 3))
        self.f_lot = ttk.Entry(card, style="Dark.TEntry",
                               font=("Segoe UI", 10))
        self.f_lot.pack(fill="x", ipady=2, pady=(0, 8))

        ttk.Label(card, text="Batch No", style="Norm.TLabel").pack(
            anchor="w", pady=(0, 3))
        self.f_batch = ttk.Entry(card, style="Dark.TEntry",
                                 font=("Segoe UI", 10))
        self.f_batch.pack(fill="x", ipady=2, pady=(0, 10))

        ttk.Label(card, text="Report Type", style="Head.TLabel").pack(
            anchor="w", pady=(0, 5))
        self.f_type = tk.StringVar(value="detailed")
        for txt, val in [("Detailed (per reading)", "detailed"),
                         ("Summary (per session)", "monthly")]:
            rb = tk.Radiobutton(card, text=txt, variable=self.f_type,
                                value=val, bg=C["bg_card"], fg=C["text"],
                                selectcolor=C["bg_input"],
                                activebackground=C["bg_card"],
                                activeforeground=C["text"],
                                font=("Segoe UI", 9))
            rb.pack(anchor="w", pady=1)

        ttk.Button(card, text="Generate Report",
                   command=self._generate,
                   style="Green.TButton").pack(fill="x", pady=(14, 4))
        ttk.Button(card, text="Export Excel",
                   command=self._export_excel,
                   style="Accent.TButton").pack(fill="x", pady=4)
        ttk.Button(card, text="Export PDF",
                   command=self._export_pdf,
                   style="Orange.TButton").pack(fill="x", pady=4)
        ttk.Button(card, text="Clear",
                   command=self._clear_filters,
                   style="Ghost.TButton").pack(fill="x", pady=4)

    def _build_results(self, parent):
        card = ttk.Frame(parent, style="Card.TFrame", padding=10)
        card.pack(fill="both", expand=True, padx=(4, 0))
        self.res_header = ttk.Label(card, text="No report generated",
                                     style="Card2.TLabel")
        self.res_header.pack(anchor="w", pady=(0, 6))
        tk.Frame(card, height=1, bg=C["border"]).pack(fill="x", pady=(0, 6))

        tf = ttk.Frame(card, style="Card.TFrame")
        tf.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(tf, columns=("c1",), show="headings",
                                  height=20, style="Report.Treeview")
        vs = ttk.Scrollbar(tf, orient="vertical", command=self.tree.yview)
        hs = ttk.Scrollbar(tf, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        tf.rowconfigure(0, weight=1)
        tf.columnconfigure(0, weight=1)

    # ═══════════════ GENERATE ═══════════════
    def _generate(self):
        try:
            machine = self.f_machine.get()
            machine = "" if machine == "All Machines" else machine
            params = [
                machine, self.f_from.get().strip(),
                self.f_to.get().strip(), self.f_lot.get().strip(),
                self.f_batch.get().strip(), self.f_type.get()
            ]
            results = self.db.call_sp("sp_GetProductionReport", params,
                                       fetch=True)
            if not results:
                messagebox.showinfo("No Data", "No records found.")
                return

            rtype = self.f_type.get()
            # FIX: Convert pyodbc Row objects to plain lists
            rows = [list(r) for r in results]

            # Dedup for detailed by log_id
            if rtype == "detailed":
                seen = set()
                unique = []
                for r in rows:
                    lid = r[0]
                    if lid not in seen:
                        seen.add(lid)
                        unique.append(r)
                rows = unique

            self.report_data = rows
            self.report_type = rtype
            self._populate_tree(rows, rtype)
            label = "Detailed" if rtype == "detailed" else "Summary"
            mach_label = machine or "All"
            self.res_header.config(
                text=f"{len(rows)} records | {label} | Machine: {mach_label}")
        except Exception as e:
            self.log.error(f"Generate error: {e}")
            messagebox.showerror("Error", str(e))

    def _populate_tree(self, rows, rtype):
        for item in self.tree.get_children():
            self.tree.delete(item)
        cols = DETAILED_COLS if rtype == "detailed" else SUMMARY_COLS
        col_ids = [f"c{i}" for i in range(len(cols))]
        self.tree["columns"] = col_ids
        for i, name in enumerate(cols):
            cid = col_ids[i]
            self.tree.heading(cid, text=name,
                              command=lambda c=cid: self._sort(c))
            w = 120 if "Timestamp" in name or "Reading" in name else 85
            self.tree.column(cid, width=w, minwidth=55)

        for row in rows:
            vals = []
            for j, v in enumerate(row):
                if rtype == "detailed":
                    if j == 2:
                        vals.append(v.strftime("%Y-%m-%d %H:%M:%S")
                                    if hasattr(v, "strftime") else str(v))
                    elif j == 4:
                        vals.append(
                            "Pass" if v == 1 else
                            "Under" if v == 0 else "Over")
                    elif j == 5:
                        vals.append(
                            "Fail" if v == 1 else
                            "Pass" if v == 0 else
                            str(v) if v is not None else "N/A")
                    else:
                        vals.append(v if v is not None else "")
                else:
                    if j in (6, 7):
                        vals.append(v.strftime("%Y-%m-%d %H:%M")
                                    if hasattr(v, "strftime") else str(v))
                    elif j == 13:
                        vals.append(f"{float(v):.1f}"
                                    if v is not None else "0")
                    elif j == 16:
                        vals.append(f"{float(v):.1f}"
                                    if v is not None else "")
                    else:
                        vals.append(v if v is not None else "")
            self.tree.insert("", tk.END, values=vals)

    def _sort(self, col):
        if self.sort_col == col:
            self.sort_rev = not self.sort_rev
        else:
            self.sort_col = col
            self.sort_rev = False
        items = [(self.tree.set(k, col), k)
                 for k in self.tree.get_children("")]
        try:
            items.sort(key=lambda t: float(t[0].replace(",", "")),
                       reverse=self.sort_rev)
        except ValueError:
            items.sort(key=lambda t: t[0].lower(),
                       reverse=self.sort_rev)
        for idx, (_, k) in enumerate(items):
            self.tree.move(k, "", idx)

    def _clear_filters(self):
        self.f_machine.set("All Machines")
        self.f_from.delete(0, tk.END)
        self.f_from.insert(0, datetime.now().strftime("%Y-%m-01"))
        self.f_to.delete(0, tk.END)
        self.f_to.insert(0, datetime.now().strftime("%Y-%m-%d"))
        self.f_lot.delete(0, tk.END)
        self.f_batch.delete(0, tk.END)

    # ═══════════════ EXCEL EXPORT (FIXED) ═══════════════
    def _export_excel(self):
        if not HAS_PANDAS:
            messagebox.showerror(
                "Missing",
                "pandas + openpyxl required.\npip install pandas openpyxl")
            return
        if not self.report_data:
            messagebox.showwarning("No Data", "Generate a report first.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            title="Save Excel Report")
        if not path:
            return

        try:
            rtype = self.report_type
            cols = list(DETAILED_COLS if rtype == "detailed"
                        else SUMMARY_COLS)
            rows = self.report_data

            # FIX: Convert each row value properly
            data = []
            for row in rows:
                processed = []
                for v in row:
                    if hasattr(v, "strftime"):
                        processed.append(v.strftime("%Y-%m-%d %H:%M:%S"))
                    elif v is None:
                        processed.append("")
                    else:
                        processed.append(v)
                data.append(processed)

            # FIX: Ensure column count matches data width
            if data:
                ncols = len(data[0])
                if ncols > len(cols):
                    cols += [f"Col_{i}" for i in range(len(cols), ncols)]
                elif ncols < len(cols):
                    cols = cols[:ncols]

            df = pd.DataFrame(data, columns=cols)

            # Detailed: human readable status
            if rtype == "detailed" and "Status" in df.columns:
                status_map = {0: "Under", 1: "Pass", 2: "Over"}
                metal_map = {0: "Pass", 1: "Fail"}
                df["Status"] = df["Status"].map(status_map).fillna(
                    df["Status"])
                df["Metal Status"] = df["Metal Status"].map(
                    metal_map).fillna(df["Metal Status"])

            sheet = ("Detailed_Report" if rtype == "detailed"
                     else "Summary_Report")

            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name=sheet, index=False,
                            startrow=1)
                ws = writer.sheets[sheet]
                ws.cell(row=1, column=1,
                        value=(f"PanCen Production Report - "
                               f"{sheet.replace('_', ' ')} - "
                               f"{datetime.now():%Y-%m-%d %H:%M}"))

                # Auto column widths
                for col_cells in ws.columns:
                    maxw = 0
                    letter = col_cells[0].column_letter
                    for cell in col_cells:
                        try:
                            clen = len(str(cell.value or ""))
                            if clen > maxw:
                                maxw = clen
                        except Exception:
                            pass
                    ws.column_dimensions[letter].width = min(maxw + 3, 45)

                # Add statistics sheet for summary
                if rtype == "monthly" and len(data) > 0:
                    self._excel_stats_sheet(writer, df)

            messagebox.showinfo("Exported",
                                f"Excel saved:\n{path}\n\n{len(rows)} records")
        except Exception as e:
            self.log.error(f"Excel export error: {e}")
            messagebox.showerror("Export Error", f"Error exporting:\n{e}")

    def _excel_stats_sheet(self, writer, df):
        """Add statistics summary sheet to Excel workbook."""
        try:
            rows = []
            rows.append(["PRODUCTION STATISTICS SUMMARY", ""])
            rows.append(["Generated",
                         datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
            rows.append(["", ""])
            rows.append(["Total Sessions", len(df)])

            total_readings = 0
            total_pass = 0
            if "Total Readings" in df.columns:
                total_readings = pd.to_numeric(
                    df["Total Readings"], errors="coerce").sum()
                rows.append(["Total Readings", int(total_readings)])
            if "Pass Count" in df.columns:
                total_pass = pd.to_numeric(
                    df["Pass Count"], errors="coerce").sum()
                rows.append(["Total Pass", int(total_pass)])
                if total_readings > 0:
                    rate = total_pass / total_readings * 100
                    rows.append(["Overall Pass Rate", f"{rate:.2f}%"])
            if "Under Fail" in df.columns:
                val = pd.to_numeric(
                    df["Under Fail"], errors="coerce").sum()
                rows.append(["Total Under Failures", int(val)])
            if "Over Fail" in df.columns:
                val = pd.to_numeric(
                    df["Over Fail"], errors="coerce").sum()
                rows.append(["Total Over Failures", int(val)])
            if "Metal Fail" in df.columns:
                val = pd.to_numeric(
                    df["Metal Fail"], errors="coerce").sum()
                rows.append(["Total Metal Failures", int(val)])
            rows.append(["", ""])
            if "Min Weight" in df.columns:
                rows.append(["Overall Min Weight (g)",
                             pd.to_numeric(
                                 df["Min Weight"], errors="coerce").min()])
                rows.append(["Overall Max Weight (g)",
                             pd.to_numeric(
                                 df["Max Weight"], errors="coerce").max()])
                avg = pd.to_numeric(
                    df["Avg Weight"], errors="coerce").mean()
                rows.append(["Overall Avg Weight (g)", f"{avg:.1f}"])

            sdf = pd.DataFrame(rows, columns=["Metric", "Value"])
            sdf.to_excel(writer, sheet_name="Statistics", index=False)
            ws = writer.sheets["Statistics"]
            ws.column_dimensions["A"].width = 35
            ws.column_dimensions["B"].width = 25
        except Exception as e:
            self.log.warning(f"Stats sheet error: {e}")

    # ═══════════════ PDF EXPORT (FIXED) ═══════════════
    def _export_pdf(self):
        if not HAS_REPORTLAB:
            messagebox.showerror(
                "Missing",
                "reportlab required.\npip install reportlab")
            return
        if not self.report_data:
            messagebox.showwarning("No Data", "Generate a report first.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
            title="Save PDF Report")
        if not path:
            return

        try:
            rtype = self.report_type
            rows = self.report_data

            doc = SimpleDocTemplate(
                path, pagesize=landscape(A4),
                leftMargin=12 * mm, rightMargin=12 * mm,
                topMargin=15 * mm, bottomMargin=15 * mm)
            elements = []
            styles = getSampleStyleSheet()

            # ── Title header ──
            title_style = ParagraphStyle(
                "ReportTitle", parent=styles["Title"],
                fontSize=16, textColor=colors.white,
                spaceAfter=0)
            date_style = ParagraphStyle(
                "ReportDate", parent=styles["Normal"],
                fontSize=9, textColor=colors.HexColor("#cccccc"),
                alignment=2)

            hdr_data = [
                [Paragraph(
                    f"PanCen Production Report - "
                    f"{'Detailed' if rtype == 'detailed' else 'Summary'}",
                    title_style),
                 Paragraph(
                    f"Generated: {datetime.now():%Y-%m-%d %H:%M}",
                    date_style)]
            ]
            hdr_t = Table(hdr_data, colWidths=[440, 200])
            hdr_t.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0),
                 colors.HexColor("#2d2d44")),
                ("TOPPADDING", (0, 0), (-1, 0), 10),
                ("BOTTOMPADDING", (0, 0), (-1, 0), 10),
                ("LEFTPADDING", (0, 0), (-1, 0), 12),
                ("RIGHTPADDING", (0, 0), (-1, 0), 12),
                ("VALIGN", (0, 0), (-1, 0), "MIDDLE"),
            ]))
            elements.append(hdr_t)
            elements.append(Spacer(1, 6 * mm))

            # ── Filter info ──
            filter_info = (
                f"Machine: {self.f_machine.get()} | "
                f"Date: {self.f_from.get()} to {self.f_to.get()} | "
                f"Records: {len(rows)}")
            fi_style = ParagraphStyle(
                "FilterInfo", parent=styles["Normal"],
                fontSize=8, textColor=colors.HexColor("#888888"))
            elements.append(Paragraph(filter_info, fi_style))
            elements.append(Spacer(1, 4 * mm))

            # ── Statistics box for summary ──
            if rtype == "monthly" and rows:
                stat_elements = self._pdf_stats_box(rows, styles)
                elements.extend(stat_elements)
                elements.append(Spacer(1, 5 * mm))

            # ── Data table ──
            elements.extend(self._pdf_data_table(rows, rtype))

            # ── Footer note ──
            elements.append(Spacer(1, 6 * mm))
            foot_style = ParagraphStyle(
                "Footer", parent=styles["Normal"],
                fontSize=7, textColor=colors.HexColor("#666666"),
                alignment=1)
            elements.append(Paragraph(
                f"PanCen Software v{APP_VERSION} | "
                f"Report generated {datetime.now():%Y-%m-%d %H:%M:%S}",
                foot_style))

            doc.build(elements)
            messagebox.showinfo("Exported",
                                f"PDF saved:\n{path}\n\n{len(rows)} records")
        except Exception as e:
            self.log.error(f"PDF export error: {e}")
            messagebox.showerror("Export Error", f"Error exporting:\n{e}")

    def _pdf_stats_box(self, rows, styles):
        """Build statistics summary section for PDF."""
        elements = []
        total_readings = sum(
            (int(r[8]) for r in rows if r[8] is not None), 0)
        total_pass = sum(
            (int(r[9]) for r in rows if r[9] is not None), 0)
        total_under = sum(
            (int(r[10]) for r in rows if r[10] is not None), 0)
        total_over = sum(
            (int(r[11]) for r in rows if r[11] is not None), 0)
        total_metal = sum(
            (int(r[12]) for r in rows if r[12] is not None), 0)
        pass_rate = ((total_pass / total_readings * 100)
                     if total_readings > 0 else 0)

        min_w = min((r[14] for r in rows if r[14] is not None),
                    default=0)
        max_w = max((r[15] for r in rows if r[15] is not None),
                    default=0)
        avg_vals = [float(r[16]) for r in rows
                    if r[16] is not None]
        avg_w = sum(avg_vals) / len(avg_vals) if avg_vals else 0

        stat_data = [
            ["Sessions", "Readings", "Pass", "Under Fail",
             "Over Fail", "Metal Fail", "Pass Rate",
             "Min Wt", "Max Wt", "Avg Wt"],
            [str(len(rows)), str(total_readings), str(total_pass),
             str(total_under), str(total_over), str(total_metal),
             f"{pass_rate:.1f}%",
             str(min_w), str(max_w), f"{avg_w:.1f}"]
        ]

        t = Table(stat_data, colWidths=[58] * 10)
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0),
             colors.HexColor("#7c5cfc")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, 0), 7),
            ("BACKGROUND", (0, 1), (-1, 1),
             colors.HexColor("#2d2d44")),
            ("TEXTCOLOR", (0, 1), (-1, 1), colors.white),
            ("FONTNAME", (0, 1), (-1, 1), "Helvetica-Bold"),
            ("FONTSIZE", (0, 1), (-1, 1), 9),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("GRID", (0, 0), (-1, -1), 0.5,
             colors.HexColor("#3d3d5c")),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ]))
        elements.append(t)
        return elements

    def _pdf_data_table(self, rows, rtype):
        """Build paginated data table for PDF."""
        elements = []

        if rtype == "detailed":
            pdf_cols = ["#", "Prod ID", "Timestamp", "Weight",
                        "Status", "Metal", "Lot No", "Product",
                        "Machine"]
            pdf_idx = [0, 1, 2, 3, 4, 5, 6, 7, 13]
            col_widths = [35, 42, 95, 50, 40, 38, 60, 70, 58]
        else:
            pdf_cols = ["Prod ID", "Machine", "Lot No", "Product",
                        "Total", "Pass", "Under", "Over",
                        "Metal", "Rate%", "Min", "Max", "Avg"]
            pdf_idx = [0, 1, 2, 3, 8, 9, 10, 11, 12, 13, 14, 15, 16]
            col_widths = [40, 52, 52, 58, 36, 36, 36, 36,
                          36, 40, 42, 42, 42]

        # Format all rows
        formatted = []
        for row in rows:
            r = []
            for j in pdf_idx:
                if j >= len(row):
                    r.append("")
                    continue
                v = row[j]
                if hasattr(v, "strftime"):
                    r.append(v.strftime("%m-%d %H:%M"))
                elif j == 4 and rtype == "detailed":
                    r.append("Pass" if v == 1 else
                             "Under" if v == 0 else "Over")
                elif j == 5 and rtype == "detailed":
                    r.append("Fail" if v == 1 else
                             "OK" if v == 0 else
                             str(v or ""))
                elif j == 13 and rtype == "monthly":
                    r.append(f"{float(v):.1f}" if v is not None
                             else "0")
                elif j == 16 and rtype == "monthly":
                    r.append(f"{float(v):.1f}" if v is not None
                             else "")
                elif v is None:
                    r.append("")
                else:
                    r.append(str(v))
            formatted.append(r)

        # Paginate (45 rows per page)
        PAGE_SIZE = 45
        for pg_start in range(0, len(formatted), PAGE_SIZE):
            chunk = formatted[pg_start:pg_start + PAGE_SIZE]
            table_data = [pdf_cols] + chunk

            t = Table(table_data, colWidths=col_widths, repeatRows=1)
            style_cmds = [
                ("BACKGROUND", (0, 0), (-1, 0),
                 colors.HexColor("#7c5cfc")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, 0), 7),
                ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
                ("FONTSIZE", (0, 1), (-1, -1), 6.5),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("GRID", (0, 0), (-1, -1), 0.3,
                 colors.HexColor("#3d3d5c")),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1),
                 [colors.HexColor("#f8f8ff"),
                  colors.HexColor("#eeeef8")]),
            ]

            # Color-code status column for detailed
            if rtype == "detailed":
                for i, r in enumerate(chunk, start=1):
                    if len(r) > 4:
                        if r[4] == "Under":
                            style_cmds.append(
                                ("TEXTCOLOR", (4, i), (4, i),
                                 colors.HexColor("#ff8c00")))
                        elif r[4] == "Over":
                            style_cmds.append(
                                ("TEXTCOLOR", (4, i), (4, i),
                                 colors.HexColor("#ff3333")))
                        elif r[4] == "Pass":
                            style_cmds.append(
                                ("TEXTCOLOR", (4, i), (4, i),
                                 colors.HexColor("#22aa44")))
                    if len(r) > 5 and r[5] == "Fail":
                        style_cmds.append(
                            ("TEXTCOLOR", (5, i), (5, i),
                             colors.HexColor("#ff3333")))

            t.setStyle(TableStyle(style_cmds))
            elements.append(t)

            # Page break between chunks
            if pg_start + PAGE_SIZE < len(formatted):
                elements.append(PageBreak())

        return elements

    # ═══════════════ LIFECYCLE ═══════════════
    def _on_close(self):
        self.db.close()
        self.root.destroy()


def main():
    root = tk.Tk()
    ReportApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()