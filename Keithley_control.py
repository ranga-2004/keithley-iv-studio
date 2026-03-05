"""
Keithley 2450 — I-V Studio
Professional measurement application for I-V characterisation and FET analysis.
Author : Rangaraajan Muralidaran
Version: 3.0
"""

import pyvisa
import time
import csv
import os
import numpy as np
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import matplotlib
matplotlib.use("TkAgg")
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.ticker as mticker

# ═══════════════════════════════════════════════════════════════════════════════
#  DESIGN TOKENS
# ═══════════════════════════════════════════════════════════════════════════════
BG      = "#ECEEF2"
PANEL   = "#FFFFFF"
PANEL2  = "#F4F6F9"
PANEL3  = "#EBF0F7"
BORDER  = "#CBD2DC"
BORDER2 = "#B0BACC"
TEXT    = "#111827"
TEXT2   = "#4B5563"
TEXT3   = "#9CA3AF"
ACCENT  = "#1E3A5F"
ACCENT2 = "#0369A1"
ACCENT3 = "#E8F0FB"
SUCCESS = "#166534"
SUCCESS2= "#DCFCE7"
DANGER  = "#991B1B"
DANGER2 = "#FEE2E2"
WARN    = "#92400E"
WARN2   = "#FEF3C7"
SEP     = "#E5E9EF"
HOVER   = "#F0F4FF"

# Plot colour palettes
Y1_PAL  = ["#1D4ED8","#0E7490","#15803D","#7C3AED","#B45309"]
Y2_PAL  = ["#BE185D","#9333EA","#C2410C","#0F766E","#374151"]
FET_PAL = ["#1D4ED8","#DC2626","#16A34A","#9333EA","#D97706",
           "#0E7490","#65A30D","#BE185D","#0369A1","#374151"]

FONT_H1   = ("Segoe UI", 14, "bold")
FONT_H2   = ("Segoe UI", 11, "bold")
FONT_H3   = ("Segoe UI", 10, "bold")
FONT_BODY = ("Segoe UI",  9)
FONT_SM   = ("Segoe UI",  8)
FONT_MONO = ("Consolas",  9, "bold")
FONT_MONO2= ("Consolas", 10, "bold")

TSP_SAFE_ABORT = (
    "pcall(function() "
    "  local s = trigger.model.state() "
    "  if s == trigger.STATE_RUNNING or "
    "     s == trigger.STATE_WAITING  or "
    "     s == trigger.STATE_PAUSED   then "
    "    trigger.model.abort() "
    "  end "
    "end)"
)

# ═══════════════════════════════════════════════════════════════════════════════
#  REUSABLE UI COMPONENTS
# ═══════════════════════════════════════════════════════════════════════════════

class Card(tk.Frame):
    """Bordered panel with an optional coloured title bar."""
    def __init__(self, parent, title="", accent=ACCENT, **kw):
        kw.setdefault("bg", PANEL)
        kw.setdefault("highlightbackground", BORDER)
        kw.setdefault("highlightthickness", 1)
        super().__init__(parent, **kw)
        if title:
            hdr = tk.Frame(self, bg=accent)
            hdr.pack(fill="x")
            tk.Label(hdr, text=f"  {title}", bg=accent, fg="white",
                     font=FONT_H3, anchor="w").pack(fill="x", padx=2, pady=5)
        self.body = tk.Frame(self, bg=PANEL)
        self.body.pack(fill="x", padx=8, pady=(6, 8))


class Section(tk.Frame):
    """Collapsible section with animated toggle."""
    def __init__(self, parent, title="", open_=True, **kw):
        kw.setdefault("bg", BG)
        super().__init__(parent, **kw)
        self._open = open_

        hdr = tk.Frame(self, bg=PANEL2, cursor="hand2",
                       highlightbackground=BORDER, highlightthickness=1)
        hdr.pack(fill="x")
        tk.Frame(hdr, bg=ACCENT2, width=3).pack(side="left", fill="y")
        self._arr = tk.Label(hdr, text="▾" if open_ else "▸",
                             bg=PANEL2, fg=ACCENT2, font=("Segoe UI", 10))
        self._arr.pack(side="left", padx=(6, 2), pady=3)
        tk.Label(hdr, text=title, bg=PANEL2, fg=ACCENT,
                 font=FONT_H3).pack(side="left", pady=3)
        hdr.bind("<Button-1>", self._toggle)
        for w in hdr.winfo_children():
            w.bind("<Button-1>", self._toggle)

        self.body = tk.Frame(self, bg=PANEL,
                             highlightbackground=BORDER, highlightthickness=1)
        if open_:
            self.body.pack(fill="x", pady=(0, 2))

    def _toggle(self, _=None):
        if self._open:
            self.body.pack_forget()
            self._arr.config(text="▸")
        else:
            self.body.pack(fill="x", pady=(0, 2))
            self._arr.config(text="▾")
        self._open = not self._open


class StatusBar(tk.Frame):
    """Bottom status bar with indicator dot, message, and elapsed timer."""
    STATES = {
        "ready":    ("#22C55E", "●  Ready"),
        "running":  ("#F59E0B", "●  Running…"),
        "stopping": ("#EF4444", "●  Stopping…"),
        "done":     ("#22C55E", "●  Complete"),
        "error":    ("#EF4444", "●  Error"),
        "stopped":  ("#EF4444", "●  Stopped"),
    }

    def __init__(self, parent, **kw):
        kw.setdefault("bg", PANEL)
        kw.setdefault("highlightbackground", BORDER)
        kw.setdefault("highlightthickness", 1)
        super().__init__(parent, height=28, **kw)
        self.pack_propagate(False)
        self._lbl = tk.Label(self, text="●  Ready", bg=PANEL,
                             fg="#22C55E", font=FONT_MONO, anchor="w")
        self._lbl.pack(side="left", padx=12, pady=3)
        self._extra = tk.Label(self, text="", bg=PANEL, fg=TEXT2,
                               font=FONT_BODY, anchor="w")
        self._extra.pack(side="left", padx=4)
        tk.Label(self, text="Developed by Rangaraajan Muralidaran  ·  Keithley 2450  ·  I-V Studio  v3.0",
                 bg=PANEL, fg=TEXT3, font=FONT_SM).pack(side="right", padx=12)

    def set(self, state, extra=""):
        col, txt = self.STATES.get(state, ("#94A3B8", state))
        self._lbl.config(text=txt, fg=col)
        self._extra.config(text=extra)

    def msg(self, text, col=TEXT2):
        self._lbl.config(text=f"●  {text}", fg=col)
        self._extra.config(text="")


class ProgressRow(tk.Frame):
    """Progress bar + percentage label + elapsed time."""
    def __init__(self, parent, **kw):
        kw.setdefault("bg", PANEL)
        super().__init__(parent, **kw)
        self._bar = ttk.Progressbar(self, orient="horizontal",
                                    mode="determinate", maximum=100,
                                    style="App.Horizontal.TProgressbar")
        self._bar.pack(side="left", fill="x", expand=True, ipady=2)
        self._pct = tk.Label(self, text="", bg=PANEL, fg=ACCENT2,
                             font=FONT_MONO, width=6)
        self._pct.pack(side="left", padx=(8, 0))
        self._timer = tk.Label(self, text="", bg=PANEL, fg=TEXT3,
                               font=FONT_SM, width=8)
        self._timer.pack(side="left", padx=(4, 0))
        self._t0 = None
        self._after_id = None

    def start(self):
        self._t0 = time.time()
        self._bar["value"] = 0
        self._pct.config(text="0 %")
        self._tick()

    def _tick(self):
        if self._t0 is None:
            return
        elapsed = int(time.time() - self._t0)
        self._timer.config(text=f"{elapsed//60:02d}:{elapsed%60:02d}")
        self._after_id = self.after(1000, self._tick)

    def update(self, pct):
        self._bar["value"] = pct
        self._pct.config(text=f"{pct} %")

    def stop(self, success=True):
        if self._after_id:
            self.after_cancel(self._after_id)
            self._after_id = None
        if success:
            self._bar["value"] = 100
            self._pct.config(text="100 %")
        else:
            self._pct.config(text="—")
        self._t0 = None

    def reset(self):
        self.stop(False)
        self._bar["value"] = 0
        self._timer.config(text="")


def _chk(parent, text, var, **kw):
    """Styled tk.Checkbutton that works correctly in all themes."""
    kw.setdefault("bg", PANEL)
    return tk.Checkbutton(parent, text=text, variable=var,
                          bg=kw.pop("bg"), activebackground=PANEL,
                          selectcolor=PANEL, relief="flat",
                          highlightthickness=0, font=FONT_BODY,
                          fg=TEXT, **kw)


def _row(parent, label, width=18):
    """Helper: labelled parameter row → returns (frame, label_widget)."""
    f = tk.Frame(parent, bg=PANEL)
    f.pack(fill="x", pady=2)
    lbl = tk.Label(f, text=label, bg=PANEL, fg=TEXT2, font=FONT_BODY,
                   width=width, anchor="w")
    lbl.pack(side="left")
    return f


def _entry(parent, default="", width=12, key=None, store=None):
    e = ttk.Entry(parent, width=width, font=FONT_BODY)
    e.insert(0, default)
    e.pack(side="right", fill="x", expand=True)
    if key and store is not None:
        store[key] = e
    return e


def _combo(parent, values, default, width=12, key=None, store=None, state="readonly"):
    c = ttk.Combobox(parent, values=values, state=state, width=width,
                     font=FONT_BODY)
    c.set(default)
    c.pack(side="right", fill="x", expand=True)
    if key and store is not None:
        store[key] = c
    return c


def _param(parent, label, default=None, key=None, store=None,
           combo_vals=None, combo_key=None, lbl_w=18):
    f = _row(parent, label, lbl_w)
    if combo_vals:
        w = _combo(f, combo_vals, combo_vals[0],
                   key=combo_key, store=store)
    else:
        w = _entry(f, default or "", key=key, store=store)
    return w


# ═══════════════════════════════════════════════════════════════════════════════
#  STYLE SETUP
# ═══════════════════════════════════════════════════════════════════════════════

def setup_styles():
    s = ttk.Style()
    s.theme_use("clam")

    s.configure("TFrame",        background=PANEL)
    s.configure("TLabel",        background=PANEL, foreground=TEXT, font=FONT_BODY)
    s.configure("TCheckbutton",  background=PANEL, foreground=TEXT, font=FONT_BODY)
    s.configure("TRadiobutton",  background=PANEL, foreground=TEXT, font=FONT_BODY)

    s.configure("TEntry",
                fieldbackground=PANEL, foreground=TEXT,
                insertcolor=ACCENT2, bordercolor=BORDER,
                lightcolor=BORDER, darkcolor=BORDER, font=FONT_BODY,
                padding=(4, 3))
    s.map("TEntry",
          bordercolor=[("focus", ACCENT2)],
          fieldbackground=[("disabled", PANEL2)])

    s.configure("TCombobox",
                fieldbackground=PANEL, background=PANEL2,
                foreground=TEXT, arrowcolor=ACCENT2,
                bordercolor=BORDER, font=FONT_BODY, padding=(4, 3))
    s.map("TCombobox",
          fieldbackground=[("readonly", PANEL), ("disabled", PANEL2)],
          bordercolor=[("focus", ACCENT2)])

    s.configure("TButton",
                background=PANEL2, foreground=TEXT,
                bordercolor=BORDER2, font=FONT_BODY, padding=(8, 4),
                relief="flat")
    s.map("TButton",
          background=[("active", ACCENT3), ("pressed", "#DBEAFE")],
          bordercolor=[("active", ACCENT2)])

    for name, bg, abg, fg in [
        ("Run.TButton",  "#16A34A", "#15803D", "white"),
        ("Stop.TButton", "#DC2626", "#B91C1C", "white"),
        ("Accent.TButton", ACCENT2, "#0284C7", "white"),
    ]:
        s.configure(name, background=bg, foreground=fg,
                    font=FONT_H3, padding=(10, 6), relief="flat")
        s.map(name, background=[("active", abg), ("pressed", abg)])

    s.configure("App.Horizontal.TProgressbar",
                troughcolor=PANEL2, background=ACCENT2,
                bordercolor=BORDER, thickness=8, lightcolor=ACCENT2,
                darkcolor=ACCENT2)

    s.configure("Treeview",
                background=PANEL, foreground=TEXT,
                fieldbackground=PANEL, rowheight=24, font=FONT_BODY,
                bordercolor=BORDER)
    s.configure("Treeview.Heading",
                background=PANEL2, foreground=ACCENT, font=FONT_H3,
                bordercolor=BORDER, relief="flat")
    s.map("Treeview",
          background=[("selected", ACCENT3)],
          foreground=[("selected", ACCENT)])

    s.configure("TSeparator", background=SEP)
    s.configure("TScrollbar",
                background=PANEL2, troughcolor=PANEL,
                bordercolor=BORDER, arrowcolor=TEXT2)


# ═══════════════════════════════════════════════════════════════════════════════
#  LIST VALUE EDITOR  (shared popup)
# ═══════════════════════════════════════════════════════════════════════════════

class ListEditor(tk.Toplevel):
    """Reusable modal list-value editor with import support."""

    def __init__(self, parent, title, initial=None, accent=ACCENT, on_save=None):
        super().__init__(parent)
        self.title(title)
        self.geometry("360x460")
        self.configure(bg=PANEL)
        self.resizable(False, True)
        self.grab_set()
        self._on_save = on_save
        self._accent  = accent

        # Header
        hdr = tk.Frame(self, bg=accent)
        hdr.pack(fill="x")
        tk.Label(hdr, text=f"  {title}", bg=accent, fg="white",
                 font=FONT_H3).pack(side="left", padx=8, pady=7)

        # Treeview
        body = tk.Frame(self, bg=PANEL); body.pack(fill="both", expand=True, padx=8, pady=8)

        col_hdr = tk.Frame(body, bg=accent); col_hdr.pack(fill="x")
        for txt, w, anch in [("Index", 60, "center"), ("Value", 240, "center")]:
            tk.Label(col_hdr, text=txt, width=w//8, bg=accent, fg="white",
                     font=FONT_SM, anchor=anch).pack(side="left", ipadx=6, ipady=4,
                                                     fill="x" if txt=="Value" else None,
                                                     expand=(txt=="Value"))

        tv_f = tk.Frame(body, bg=PANEL); tv_f.pack(fill="both", expand=True)
        vsb  = ttk.Scrollbar(tv_f, orient="vertical")
        self.tv = ttk.Treeview(tv_f, columns=("idx","val"), show="",
                               height=12, yscrollcommand=vsb.set,
                               selectmode="browse")
        vsb.config(command=self.tv.yview)
        self.tv.column("idx", width=60,  anchor="center", stretch=False)
        self.tv.column("val", width=240, anchor="center", stretch=True)
        self.tv.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        self.tv.tag_configure("odd",  background=PANEL)
        self.tv.tag_configure("even", background=PANEL2)
        self.tv.bind("<Double-1>", self._inline_edit)

        hint = tk.Label(body, text="Double-click a value to edit",
                        bg=PANEL, fg=TEXT3, font=FONT_SM)
        hint.pack(anchor="w", pady=(2, 0))

        # Buttons
        btn_f = tk.Frame(self, bg=PANEL2,
                         highlightbackground=BORDER, highlightthickness=1)
        btn_f.pack(fill="x")
        btn_left  = tk.Frame(btn_f, bg=PANEL2); btn_left.pack(side="left",  padx=8, pady=6)
        btn_right = tk.Frame(btn_f, bg=PANEL2); btn_right.pack(side="right", padx=8, pady=6)

        ttk.Button(btn_left,  text="+ Add Row",     command=self._add).pack(side="left", padx=(0,4))
        ttk.Button(btn_left,  text="✕ Delete",      command=self._delete).pack(side="left", padx=(0,4))
        ttk.Button(btn_left,  text="⬆ Import CSV",  command=self._import).pack(side="left")
        ttk.Button(btn_right, text="Cancel",         command=self.destroy).pack(side="right", padx=(4,0))
        ttk.Button(btn_right, text="Save & Close",   style="Accent.TButton",
                   command=self._save).pack(side="right")

        # Populate
        data = initial or []
        if data:
            for i, v in enumerate(data, 1):
                self.tv.insert("","end", values=(i, f"{v:g}"),
                               tags=("odd" if i%2 else "even",))
        else:
            for i in range(1, 6):
                self.tv.insert("","end", values=(i,""),
                               tags=("odd" if i%2 else "even",))

    def _inline_edit(self, event):
        if self.tv.identify("region", event.x, event.y) != "cell": return
        if self.tv.identify_column(event.x) != "#2": return
        iid = self.tv.identify_row(event.y)
        if not iid: return
        x, y, w, h = self.tv.bbox(iid, "#2")
        cur = self.tv.item(iid, "values")[1]
        ed = tk.Entry(self.tv, font=FONT_BODY, bg=PANEL, fg=TEXT,
                      insertbackground=ACCENT2, relief="solid",
                      bd=1, highlightthickness=1,
                      highlightcolor=ACCENT2, highlightbackground=BORDER)
        ed.insert(0, cur); ed.select_range(0, tk.END)
        ed.place(x=x, y=y, width=w, height=h); ed.focus_set()
        def commit(_=None):
            idx = self.tv.item(iid, "values")[0]
            tag = "odd" if int(idx)%2 else "even"
            self.tv.item(iid, values=(idx, ed.get().strip()), tags=(tag,))
            ed.destroy()
        ed.bind("<Return>", commit); ed.bind("<Tab>", commit)
        ed.bind("<Escape>", lambda _: ed.destroy())
        ed.bind("<FocusOut>", commit)

    def _add(self):
        n = len(self.tv.get_children()) + 1
        self.tv.insert("","end", values=(n,""),
                       tags=("odd" if n%2 else "even",))

    def _delete(self):
        sel = self.tv.selection()
        if not sel: return
        self.tv.delete(sel[0])
        for i, iid in enumerate(self.tv.get_children(), 1):
            cur = self.tv.item(iid,"values")[1]
            self.tv.item(iid, values=(i, cur),
                         tags=("odd" if i%2 else "even",))

    def _import(self):
        path = filedialog.askopenfilename(
            parent=self, title="Import Values",
            filetypes=[("CSV","*.csv"),("Excel","*.xlsx *.xls"),("All","*.*")])
        if not path: return
        try:
            vals = []; ext = path.rsplit(".", 1)[-1].lower()
            if ext == "csv":
                with open(path, newline="", encoding="utf-8-sig") as f:
                    for row in csv.reader(f):
                        for cell in row:
                            try: vals.append(float(cell.strip())); break
                            except ValueError: continue
            elif ext in ("xlsx","xls"):
                try:
                    import openpyxl
                    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
                    for row in wb.active.iter_rows(values_only=True):
                        for cell in row:
                            try: vals.append(float(cell)); break
                            except: continue
                except ImportError:
                    messagebox.showerror("Missing Package",
                                         "openpyxl not installed.\npip install openpyxl",
                                         parent=self); return
            if not vals:
                messagebox.showwarning("No Data","No numeric values found.", parent=self)
                return
            for iid in self.tv.get_children(): self.tv.delete(iid)
            for i, v in enumerate(vals, 1):
                self.tv.insert("","end", values=(i, f"{v:g}"),
                               tags=("odd" if i%2 else "even",))
        except Exception as e:
            messagebox.showerror("Import Error", str(e), parent=self)

    def _save(self):
        vals = []
        for iid in self.tv.get_children():
            try: vals.append(float(self.tv.item(iid,"values")[1]))
            except: pass
        if self._on_save:
            self._on_save(vals)
        self.destroy()

    def get_values(self):
        vals = []
        for iid in self.tv.get_children():
            try: vals.append(float(self.tv.item(iid,"values")[1]))
            except: pass
        return vals


# ═══════════════════════════════════════════════════════════════════════════════
#  PLOT HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

def _style_plot(fig, ax, ax2=None, title=""):
    fig.patch.set_facecolor(PANEL)
    ax.set_facecolor("#FAFBFD")
    for sp in ax.spines.values():
        sp.set_color(BORDER); sp.set_linewidth(0.8)
    ax.tick_params(colors=TEXT2, labelsize=8, width=0.6)
    ax.grid(True, linestyle="--", color=SEP, linewidth=0.7, alpha=0.9)
    ax.set_title(title, color=TEXT, fontsize=10, fontweight="bold", pad=8)
    if ax2:
        ax2.set_facecolor("none")
        for sp in ax2.spines.values():
            sp.set_color(BORDER); sp.set_linewidth(0.8)
        ax2.tick_params(colors=TEXT2, labelsize=8, width=0.6)


def _eng_formatter(ax, axis="y"):
    """Apply engineering (SI prefix) tick formatter to an axis."""
    fmt = mticker.EngFormatter(sep="")
    if axis == "y":
        ax.yaxis.set_major_formatter(fmt)
    else:
        ax.xaxis.set_major_formatter(fmt)


def _eng_str(val):
    """Return engineering-notation string for a scalar value."""
    try:
        return mticker.EngFormatter(sep="")(val)
    except Exception:
        return str(val)


# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN APPLICATION
# ═══════════════════════════════════════════════════════════════════════════════

class KeithleyApp:
    CHANNELS = ["Voltage (V)", "Current (A)", "Time (s)",
                "Resistance (Ω)", "Power (W)"]

    def __init__(self, root):
        self.root = root
        self.root.title("Keithley 2450  —  I-V Studio")
        self.root.geometry("1700x980")
        self.root.configure(bg=BG)
        self.root.minsize(1280, 720)

        self.save_dir    = r"D:\KEITHLEY 2450\Data"
        self.master_addr = "USB0::0x05E6::0x2450::04465297::INSTR"
        self.rm          = pyvisa.ResourceManager()
        self.inst        = None
        self.is_running  = False
        self._stop_flag  = False
        self.inputs      = {}
        self.sweep_data  = []
        self.ax_x        = None
        self.ax_y1       = {}
        self.ax_y2       = {}

        if not os.path.exists(self.save_dir):
            try: os.makedirs(self.save_dir)
            except OSError: pass

        setup_styles()
        self._build_ui()

    # ── Instrument helpers ────────────────────────────────────────────────────

    def _safe_abort(self):
        try:
            self.inst.write(TSP_SAFE_ABORT)
            time.sleep(0.08)
        except Exception:
            pass

    def _force_stop(self):
        if not self.is_running:
            return
        self._stop_flag = True
        self.status.set("stopping")
        try:
            if self.inst:
                self.inst.write(TSP_SAFE_ABORT)
                self.inst.write("smu.source.output = smu.OFF")
        except Exception:
            pass

    # ── UI construction ───────────────────────────────────────────────────────

    def _build_ui(self):
        # ── Top header bar ────────────────────────────────────────────────────
        hdr = tk.Frame(self.root, bg=ACCENT, height=56)
        hdr.pack(fill="x"); hdr.pack_propagate(False)

        logo_f = tk.Frame(hdr, bg=ACCENT); logo_f.pack(side="left", padx=14, pady=8)
        tk.Label(logo_f, text="KEITHLEY 2450", bg=ACCENT, fg="white",
                 font=FONT_H1).pack(anchor="w")
        tk.Label(logo_f, text="Source Measure Unit  ·  I-V Studio  v3.0",
                 bg=ACCENT, fg="#7EB6E0", font=FONT_SM).pack(anchor="w")

        # Right side of header
        right_f = tk.Frame(hdr, bg=ACCENT); right_f.pack(side="right", padx=14)

        # FET button
        ttk.Button(right_f, text="FET Characterisation  ▶",
                   style="Accent.TButton",
                   command=lambda: FETWindow(self)).pack(side="right", padx=(8,0), pady=12)

        # VISA address
        addr_f = tk.Frame(right_f, bg=ACCENT); addr_f.pack(side="right", padx=(0,8))
        tk.Label(addr_f, text="VISA Address", bg=ACCENT, fg="#7EB6E0",
                 font=FONT_SM).pack(anchor="e")
        self._addr_entry = tk.Entry(addr_f, width=36,
                                    bg="#1A3255", fg="white",
                                    insertbackground="white",
                                    relief="flat", font=FONT_MONO, bd=0,
                                    highlightthickness=1,
                                    highlightbackground="#3A5F8A",
                                    highlightcolor="#5A8FBF")
        self._addr_entry.insert(0, self.master_addr)
        self._addr_entry.pack(pady=(2,0), ipady=3)
        self._addr_entry.bind("<FocusOut>",
            lambda _: setattr(self, "master_addr", self._addr_entry.get().strip()))

        # Sample ID
        sid_f = tk.Frame(right_f, bg=ACCENT); sid_f.pack(side="right", padx=(0,8))
        tk.Label(sid_f, text="Sample ID", bg=ACCENT, fg="#7EB6E0",
                 font=FONT_SM).pack(anchor="e")
        self.sample_id = tk.Entry(sid_f, width=14,
                                  bg="#1A3255", fg="white",
                                  insertbackground="white",
                                  relief="flat", font=FONT_MONO, bd=0,
                                  highlightthickness=1,
                                  highlightbackground="#3A5F8A",
                                  highlightcolor="#5A8FBF")
        self.sample_id.insert(0, "DUT_001")
        self.sample_id.pack(pady=(2,0), ipady=3)

        # ── Status bar (bottom) ───────────────────────────────────────────────
        self.status = StatusBar(self.root)
        self.status.pack(fill="x", side="bottom")

        # ── Main paned layout ─────────────────────────────────────────────────
        paned = tk.PanedWindow(self.root, orient=tk.HORIZONTAL,
                               bg=BG, sashwidth=6,
                               sashrelief="flat", sashpad=2)
        paned.pack(fill="both", expand=True, padx=0, pady=0)

        # Left panel
        lf = tk.Frame(paned, bg=BG, width=560)
        paned.add(lf, minsize=500)
        self._build_left_panel(lf)

        # Right panel
        rf = tk.Frame(paned, bg=BG)
        paned.add(rf, stretch="always")
        self._build_right_panel(rf)

    def _build_left_panel(self, parent):
        """Scrollable controls panel."""
        cvs = tk.Canvas(parent, bg=BG, highlightthickness=0)
        vsb = ttk.Scrollbar(parent, orient="vertical", command=cvs.yview)
        inner = tk.Frame(cvs, bg=BG)
        inner.bind("<Configure>",
                   lambda e: cvs.configure(scrollregion=cvs.bbox("all")))
        self._ctrl_win = cvs.create_window((0,0), window=inner, anchor="nw")
        cvs.bind("<Configure>",
                 lambda e: cvs.itemconfig(self._ctrl_win, width=e.width))
        cvs.configure(yscrollcommand=vsb.set)
        cvs.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        cvs.bind_all("<MouseWheel>",
                     lambda e: cvs.yview_scroll(int(-1*(e.delta/120)), "units"))
        self._build_controls(inner)

    def _build_controls(self, p):
        pad = {"padx": 10, "pady": (6,0)}

        # ── Source settings ───────────────────────────────────────────────────
        src = Card(p, title="Source & Sweep"); src.pack(fill="x", **pad)
        b = src.body

        # Mode row
        mf = _row(b, "Source Mode", 16)
        modes = ["Voltage Sweep","Current Sweep","Voltage List Sweep",
                 "Current List Sweep","Voltage Bias","Current Bias"]
        self.inputs["Mode"] = _combo(mf, modes, "Voltage Sweep",
                                     key="Mode", store=self.inputs)
        self.inputs["Mode"].bind("<<ComboboxSelected>>", self._update_ui_logic)

        # Option checkboxes
        chk_f = tk.Frame(b, bg=PANEL); chk_f.pack(fill="x", pady=(4,2))
        self.inputs["Stepper"] = tk.BooleanVar()
        self.inputs["Dual"]    = tk.BooleanVar()
        self.inputs["PulseEn"] = tk.BooleanVar()
        _chk(chk_f, "Stepper (Node 2)",  self.inputs["Stepper"],
             command=self._update_ui_logic).pack(side="left", padx=(0,12))
        _chk(chk_f, "Dual Sweep",        self.inputs["Dual"]).pack(side="left", padx=(0,12))
        _chk(chk_f, "Pulse Mode",        self.inputs["PulseEn"],
             command=self._update_ui_logic).pack(side="left")

        tk.Frame(b, bg=SEP, height=1).pack(fill="x", pady=(6,4))

        # Sweep params area (shown/hidden by mode)
        self.f_sweep = tk.Frame(b, bg=PANEL)
        self.f_sweep.pack(fill="x")
        _param(self.f_sweep, "Start",  "0",   key="Start",  store=self.inputs, lbl_w=16)
        _param(self.f_sweep, "Stop",   "5",   key="Stop",   store=self.inputs, lbl_w=16)
        _param(self.f_sweep, "Step",   "0.1", key="Step",   store=self.inputs, lbl_w=16)
        _param(self.f_sweep, "Points", "51",  key="Points", store=self.inputs, lbl_w=16)
        self.inputs["Start"].bind("<FocusOut>",  self._calc_points)
        self.inputs["Stop"].bind("<FocusOut>",   self._calc_points)
        self.inputs["Step"].bind("<FocusOut>",   self._calc_points)
        self.inputs["Points"].bind("<FocusOut>", self._calc_step)

        _param(self.f_sweep, "Sweep Type", combo_vals=["Linear","Logarithmic"],
               combo_key="SweepType", store=self.inputs, lbl_w=16)

        # List sweep area
        self.f_list = tk.Frame(b, bg=PANEL)
        self._build_list_table(self.f_list)
        # (initially hidden, shown when list mode selected)

        # Pulse params (hidden by default)
        self.f_pulse = tk.Frame(b, bg=PANEL2,
                                highlightbackground=BORDER, highlightthickness=1)
        pulse_hdr = tk.Frame(self.f_pulse, bg=ACCENT2)
        pulse_hdr.pack(fill="x")
        tk.Label(pulse_hdr, text="  Pulse Parameters",
                 bg=ACCENT2, fg="white", font=FONT_SM).pack(side="left", pady=3, padx=4)
        _param(self.f_pulse, "Bias Level",   "0.00",  key="PulseBias", store=self.inputs, lbl_w=16)
        _param(self.f_pulse, "On Time (s)",  "0.001", key="OnTime",    store=self.inputs, lbl_w=16)
        _param(self.f_pulse, "Off Time (s)", "0.01",  key="OffTime",   store=self.inputs, lbl_w=16)

        tk.Frame(b, bg=SEP, height=1).pack(fill="x", pady=(6,4))
        _param(b, "Source Range",  combo_vals=["Best Fixed","Auto","200mV","2V","20V","200V"],
               combo_key="RangeV", store=self.inputs, lbl_w=16)
        _param(b, "Compliance",    "0.01", key="Limit",  store=self.inputs, lbl_w=16)
        _param(b, "Source Delay (s)", "0.1", key="Delay", store=self.inputs, lbl_w=16)

        # Stepper config (hidden by default)
        self.f_stepper = tk.Frame(b, bg=PANEL3,
                                  highlightbackground=ACCENT2, highlightthickness=1)
        step_hdr = tk.Frame(self.f_stepper, bg=ACCENT)
        step_hdr.pack(fill="x")
        tk.Label(step_hdr, text="  Stepper  —  Node 2 Configuration",
                 bg=ACCENT, fg="white", font=FONT_SM).pack(side="left", pady=3, padx=4)
        _param(self.f_stepper, "Start (V)", "2.0", key="StepStart",  store=self.inputs, lbl_w=16)
        _param(self.f_stepper, "Stop (V)",  "3.0", key="StepStop",   store=self.inputs, lbl_w=16)
        _param(self.f_stepper, "Points",    "3",   key="StepPoints", store=self.inputs, lbl_w=16)

        # ── Measure settings ──────────────────────────────────────────────────
        meas = Card(p, title="Measure"); meas.pack(fill="x", **pad)
        mb = meas.body

        def meas_group(label, en_key, en_val, rk, rv, xlbl, xk, xv, xd):
            grp = tk.Frame(mb, bg=PANEL,
                           highlightbackground=SEP, highlightthickness=1)
            grp.pack(fill="x", pady=3)
            ghdr = tk.Frame(grp, bg=PANEL2); ghdr.pack(fill="x")
            self.inputs[en_key] = tk.BooleanVar(value=en_val)
            _chk(ghdr, f"  {label}", self.inputs[en_key],
                 bg=PANEL2).pack(side="left", padx=4, pady=3)
            gbody = tk.Frame(grp, bg=PANEL); gbody.pack(fill="x", padx=8, pady=(2,4))
            r1 = tk.Frame(gbody, bg=PANEL); r1.pack(fill="x", pady=1)
            tk.Label(r1, text="Range", bg=PANEL, fg=TEXT2, font=FONT_SM,
                     width=14, anchor="w").pack(side="left")
            self.inputs[rk] = ttk.Combobox(r1, values=rv, state="readonly",
                                            width=14, font=FONT_SM)
            self.inputs[rk].set("Auto")
            self.inputs[rk].pack(side="left", padx=2)
            r2 = tk.Frame(gbody, bg=PANEL); r2.pack(fill="x", pady=1)
            tk.Label(r2, text=xlbl, bg=PANEL, fg=TEXT2, font=FONT_SM,
                     width=14, anchor="w").pack(side="left")
            self.inputs[xk] = ttk.Combobox(r2, values=xv, state="readonly",
                                            width=14, font=FONT_SM)
            self.inputs[xk].set(xd)
            self.inputs[xk].pack(side="left", padx=2)

        meas_group("Current", "MeasI", True,
                   "RangeI",
                   ["Auto","10pA","100pA","1nA","10nA","100nA",
                    "1uA","10uA","100uA","1mA","10mA","100mA","1A"],
                   "Min Auto Range","MinRangeI",
                   ["10pA","100pA","1nA","10nA","100nA",
                    "1uA","10uA","100uA","1mA"], "10nA")

        # Voltage group
        vg = tk.Frame(mb, bg=PANEL,
                      highlightbackground=SEP, highlightthickness=1)
        vg.pack(fill="x", pady=3)
        vghdr = tk.Frame(vg, bg=PANEL2); vghdr.pack(fill="x")
        self.inputs["MeasV"] = tk.BooleanVar(value=False)
        _chk(vghdr, "  Voltage", self.inputs["MeasV"],
             bg=PANEL2).pack(side="left", padx=4, pady=3)
        vgb = tk.Frame(vg, bg=PANEL); vgb.pack(fill="x", padx=8, pady=(2,4))
        vt  = tk.Frame(vgb, bg=PANEL); vt.pack(fill="x", pady=1)
        tk.Label(vt, text="Report", bg=PANEL, fg=TEXT2, font=FONT_SM,
                 width=14, anchor="w").pack(side="left")
        self.inputs["MeasVType"] = ttk.Combobox(vt, values=["Programmed","Actual"],
                                                 state="readonly", width=14, font=FONT_SM)
        self.inputs["MeasVType"].set("Programmed")
        self.inputs["MeasVType"].pack(side="left", padx=2)

        meas_group("Resistance", "MeasR", False,
                   "RangeR",
                   ["Auto","20Ω","200Ω","2kΩ","20kΩ","200kΩ","2MΩ","20MΩ","200MΩ"],
                   "Min Auto Range","MinRangeR",
                   ["20Ω","200Ω","2kΩ","20kΩ"], "20Ω")

        extras = tk.Frame(mb, bg=PANEL); extras.pack(fill="x", pady=(4,0))
        self.inputs["Timestamp"] = tk.BooleanVar(value=True)
        self.inputs["MeasP"]     = tk.BooleanVar(value=False)
        _chk(extras, "Include Timestamp", self.inputs["Timestamp"]).pack(side="left")
        _chk(extras, "Power (W)",         self.inputs["MeasP"]).pack(side="left", padx=12)

        # ── Speed & instrument ────────────────────────────────────────────────
        spd = Card(p, title="Instrument Settings"); spd.pack(fill="x", **pad)
        sb = spd.body
        _param(sb, "NPLC",           "1",    key="NPLC",      store=self.inputs, lbl_w=20)
        _param(sb, "Auto Zero",      None,   combo_vals=["Once","Off","On"],
               combo_key="AutoZero", store=self.inputs, lbl_w=20)
        _param(sb, "Input Jacks",    None,   combo_vals=["Rear","Front"],
               combo_key="Terminals", store=self.inputs, lbl_w=20)
        _param(sb, "Sense Mode",     None,   combo_vals=["2-Wire","4-Wire"],
               combo_key="Sense",    store=self.inputs, lbl_w=20)
        _param(sb, "Output OFF State", None, combo_vals=["Normal","High-Z","Zero"],
               combo_key="OffState", store=self.inputs, lbl_w=20)
        _param(sb, "High Capacitance", None, combo_vals=["Off","On"],
               combo_key="HighC",    store=self.inputs, lbl_w=20)

        # ── Run / Stop ────────────────────────────────────────────────────────
        run_card = tk.Frame(p, bg=BG); run_card.pack(fill="x", padx=10, pady=(8,12))
        ttk.Button(run_card, text="▶   RUN SWEEP",
                   style="Run.TButton",
                   command=self.start_thread).pack(side="left", fill="x",
                                                   expand=True, ipady=5, padx=(0,6))
        ttk.Button(run_card, text="■  STOP",
                   style="Stop.TButton",
                   command=self._force_stop).pack(side="left", ipady=5, ipadx=10)

    def _build_list_table(self, parent):
        thdr = tk.Frame(parent, bg=ACCENT); thdr.pack(fill="x")
        tk.Label(thdr, text="Index", width=8, bg=ACCENT, fg="white",
                 font=FONT_SM, anchor="center").pack(side="left", ipadx=4, ipady=3)
        self._list_col_lbl = tk.Label(thdr, text="Value",
                                      bg=ACCENT, fg="white", font=FONT_SM,
                                      anchor="center")
        self._list_col_lbl.pack(side="left", fill="x", expand=True, ipadx=4, ipady=3)
        tf = tk.Frame(parent, bg=PANEL); tf.pack(fill="x")
        vsb = ttk.Scrollbar(tf, orient="vertical")
        self.list_tv = ttk.Treeview(tf, columns=("idx","val"), show="",
                                    height=7, yscrollcommand=vsb.set,
                                    selectmode="browse")
        vsb.config(command=self.list_tv.yview)
        self.list_tv.column("idx", width=60,  anchor="center", stretch=False)
        self.list_tv.column("val", width=150, anchor="center", stretch=True)
        self.list_tv.pack(side="left", fill="x", expand=True)
        vsb.pack(side="right", fill="y")
        self.list_tv.tag_configure("odd",  background=PANEL)
        self.list_tv.tag_configure("even", background=PANEL2)
        for i in range(1, 6):
            self.list_tv.insert("","end", values=(i,""),
                                tags=("odd" if i%2 else "even",))
        self.list_tv.bind("<Double-1>", self._list_inline_edit)
        br = tk.Frame(parent, bg=PANEL); br.pack(fill="x", pady=(3,0))
        ttk.Button(br, text="+ Row",        command=self._list_add).pack(side="left", padx=(0,2))
        ttk.Button(br, text="✕ Delete",     command=self._list_del).pack(side="left", padx=(0,6))
        ttk.Button(br, text="⬆ Import",     command=self._list_import).pack(side="left")
        self._list_import_lbl = tk.Label(br, text="", bg=PANEL, fg=TEXT3, font=FONT_SM)
        self._list_import_lbl.pack(side="left", padx=6)

    def _list_add(self):
        n = len(self.list_tv.get_children()) + 1
        self.list_tv.insert("","end", values=(n,""),
                            tags=("odd" if n%2 else "even",))

    def _list_del(self):
        sel = self.list_tv.selection()
        if not sel: return
        self.list_tv.delete(sel[0])
        for i, iid in enumerate(self.list_tv.get_children(), 1):
            v = self.list_tv.item(iid,"values")[1]
            self.list_tv.item(iid, values=(i,v),
                              tags=("odd" if i%2 else "even",))

    def _list_inline_edit(self, event):
        if self.list_tv.identify("region", event.x, event.y) != "cell": return
        if self.list_tv.identify_column(event.x) != "#2": return
        iid = self.list_tv.identify_row(event.y)
        if not iid: return
        x, y, w, h = self.list_tv.bbox(iid, "#2")
        cur = self.list_tv.item(iid,"values")[1]
        ed = tk.Entry(self.list_tv, font=FONT_BODY, bg=PANEL, fg=TEXT,
                      insertbackground=ACCENT2, relief="solid", bd=1,
                      highlightthickness=1, highlightcolor=ACCENT2,
                      highlightbackground=BORDER)
        ed.insert(0, cur); ed.select_range(0, tk.END)
        ed.place(x=x, y=y, width=w, height=h); ed.focus_set()
        def commit(_=None):
            idx = self.list_tv.item(iid,"values")[0]
            tag = "odd" if int(idx)%2 else "even"
            self.list_tv.item(iid, values=(idx, ed.get().strip()), tags=(tag,))
            ed.destroy()
        ed.bind("<Return>", commit); ed.bind("<Tab>", commit)
        ed.bind("<Escape>", lambda _: ed.destroy())
        ed.bind("<FocusOut>", commit)

    def _list_import(self):
        path = filedialog.askopenfilename(
            title="Import List Values",
            filetypes=[("CSV","*.csv"),("Excel","*.xlsx *.xls"),("All","*.*")])
        if not path: return
        try:
            vals=[]; ext=path.rsplit(".",1)[-1].lower()
            if ext=="csv":
                with open(path, newline="", encoding="utf-8-sig") as f:
                    for row in csv.reader(f):
                        for cell in row:
                            try: vals.append(float(cell.strip())); break
                            except ValueError: continue
            elif ext in ("xlsx","xls"):
                try:
                    import openpyxl
                    wb=openpyxl.load_workbook(path, read_only=True, data_only=True)
                    for row in wb.active.iter_rows(values_only=True):
                        for cell in row:
                            try: vals.append(float(cell)); break
                            except: continue
                except ImportError:
                    messagebox.showerror("Missing Package","pip install openpyxl"); return
            if not vals:
                messagebox.showwarning("No Data","No numeric values found."); return
            for iid in self.list_tv.get_children(): self.list_tv.delete(iid)
            for i,v in enumerate(vals,1):
                self.list_tv.insert("","end", values=(i,f"{v:g}"),
                                    tags=("odd" if i%2 else "even",))
            self._list_import_lbl.config(text=os.path.basename(path))
        except Exception as e:
            messagebox.showerror("Import Error", str(e))

    def _get_list_values(self):
        vals=[]
        for iid in self.list_tv.get_children():
            try: vals.append(float(self.list_tv.item(iid,"values")[1]))
            except: pass
        return vals

    def _build_right_panel(self, parent):
        """Graph area with toolbar and axis selector."""
        # ── Toolbar strip ─────────────────────────────────────────────────────
        tb = tk.Frame(parent, bg=PANEL,
                      highlightbackground=BORDER, highlightthickness=1)
        tb.pack(fill="x", padx=8, pady=(8,4))

        # Progress
        prow = tk.Frame(tb, bg=PANEL); prow.pack(fill="x", padx=10, pady=(7,0))
        self.progress = ProgressRow(prow)
        self.progress.pack(fill="x")

        # Buttons
        brow = tk.Frame(tb, bg=PANEL); brow.pack(fill="x", padx=10, pady=(4,7))
        ttk.Button(brow, text="Clear Plot",   command=self._clear_plot).pack(side="left", padx=(0,4))
        ttk.Button(brow, text="Export CSV",   command=self._save_csv).pack(side="left", padx=(0,4))
        ttk.Button(brow, text="Save PNG",     command=self._save_png).pack(side="left", padx=(0,4))
        ttk.Button(brow, text="Auto Scale",   command=self._auto_scale).pack(side="left")

        # ── Chart + axis selector ─────────────────────────────────────────────
        body = tk.Frame(parent, bg=BG)
        body.pack(fill="both", expand=True, padx=8, pady=(0,8))

        chart_f = tk.Frame(body, bg=PANEL,
                           highlightbackground=BORDER, highlightthickness=1)
        chart_f.pack(side="left", fill="both", expand=True)

        self.fig = Figure(figsize=(7,5), dpi=100, facecolor=PANEL)
        self.fig.subplots_adjust(left=0.11, right=0.88, top=0.93, bottom=0.10)
        self.ax  = self.fig.add_subplot(111)
        self.ax2 = self.ax.twinx()
        self.ax2.set_visible(False)
        _style_plot(self.fig, self.ax, title="I-V Characteristic")

        self.canvas = FigureCanvasTkAgg(self.fig, master=chart_f)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(fill="both", expand=True, padx=2, pady=2)

        self._build_axis_panel(body)

    def _build_axis_panel(self, parent):
        """Right sidebar for axis channel selection."""
        sb = tk.Frame(parent, bg=PANEL, width=172,
                      highlightbackground=BORDER, highlightthickness=1)
        sb.pack(side="right", fill="y", padx=(6,0))
        sb.pack_propagate(False)

        self.ax_x  = tk.StringVar(value="Voltage (V)")
        self.ax_y1 = {ch: tk.BooleanVar(value=(ch=="Current (A)"))
                      for ch in self.CHANNELS}
        self.ax_y2 = {ch: tk.BooleanVar(value=False) for ch in self.CHANNELS}

        def section(title, col):
            tk.Frame(sb, bg=col, height=2).pack(fill="x")
            tk.Label(sb, text=title, bg=PANEL2, fg=col,
                     font=FONT_H3, anchor="w").pack(fill="x", padx=8, pady=(5,2))
            body = tk.Frame(sb, bg=PANEL)
            body.pack(fill="x", padx=4, pady=(0,4))
            tk.Frame(sb, bg=SEP, height=1).pack(fill="x")
            return body

        # X-Axis
        xb = section("X-Axis", ACCENT)
        for ch in self.CHANNELS:
            r = tk.Frame(xb, bg=PANEL); r.pack(fill="x", pady=1)
            tk.Label(r, text="◆", bg=PANEL, fg=TEXT3,
                     font=("Segoe UI",7)).pack(side="left", padx=(4,2))
            tk.Radiobutton(r, text=ch, variable=self.ax_x, value=ch,
                           bg=PANEL, fg=TEXT, selectcolor=PANEL,
                           activebackground=HOVER, activeforeground=ACCENT,
                           font=FONT_SM, command=self._replot).pack(side="left")

        # Y1-Axis
        y1b = section("Y1-Axis (left)", Y1_PAL[0])
        for i, ch in enumerate(self.CHANNELS):
            col = Y1_PAL[i % len(Y1_PAL)]
            r = tk.Frame(y1b, bg=PANEL); r.pack(fill="x", pady=1)
            tk.Label(r, text="■", bg=PANEL, fg=col,
                     font=("Segoe UI",9)).pack(side="left", padx=(4,2))
            tk.Checkbutton(r, text=ch, variable=self.ax_y1[ch],
                           bg=PANEL, fg=TEXT, selectcolor=PANEL,
                           activebackground=HOVER, activeforeground=col,
                           font=FONT_SM, relief="flat",
                           highlightthickness=0,
                           command=self._replot).pack(side="left")

        # Y2-Axis
        y2b = section("Y2-Axis (right)", Y2_PAL[0])
        for i, ch in enumerate(self.CHANNELS):
            col = Y2_PAL[i % len(Y2_PAL)]
            r = tk.Frame(y2b, bg=PANEL); r.pack(fill="x", pady=1)
            tk.Label(r, text="■", bg=PANEL, fg=col,
                     font=("Segoe UI",9)).pack(side="left", padx=(4,2))
            tk.Checkbutton(r, text=ch, variable=self.ax_y2[ch],
                           bg=PANEL, fg=TEXT, selectcolor=PANEL,
                           activebackground=HOVER, activeforeground=col,
                           font=FONT_SM, relief="flat",
                           highlightthickness=0,
                           command=self._replot).pack(side="left")

    # ── UI logic ──────────────────────────────────────────────────────────────

    def _update_ui_logic(self, _=None):
        mode = self.inputs["Mode"].get()
        is_list = "List" in mode
        if is_list:
            self.f_sweep.pack_forget()
            self.f_list.pack(fill="x", pady=(2,0))
            self._list_col_lbl.config(
                text="Voltage Value" if "Voltage" in mode else "Current Value")
        else:
            self.f_list.pack_forget()
            self.f_sweep.pack(fill="x", pady=(2,0))

        if self.inputs["PulseEn"].get():
            self.f_pulse.pack(fill="x", pady=(4,0))
        else:
            self.f_pulse.pack_forget()

        if self.inputs["Stepper"].get():
            self.f_stepper.pack(fill="x", pady=(4,0))
        else:
            self.f_stepper.pack_forget()

    def _calc_points(self, _=None):
        try:
            s = float(self.inputs["Start"].get())
            e = float(self.inputs["Stop"].get())
            st= float(self.inputs["Step"].get())
            if st == 0: return
            n = int(round(abs(e-s)/abs(st))) + 1
            self.inputs["Points"].delete(0, tk.END)
            self.inputs["Points"].insert(0, str(n))
        except: pass

    def _calc_step(self, _=None):
        try:
            s  = float(self.inputs["Start"].get())
            e  = float(self.inputs["Stop"].get())
            n  = int(self.inputs["Points"].get())
            if n <= 1: return
            st = (e-s)/(n-1)
            self.inputs["Step"].delete(0, tk.END)
            self.inputs["Step"].insert(0, f"{st:.6g}")
        except: pass

    # ── Plot ──────────────────────────────────────────────────────────────────

    def _replot(self):
        if not self.sweep_data: return
        x_ch   = self.ax_x.get()
        y1_chs = [c for c in self.CHANNELS if self.ax_y1[c].get()]
        y2_chs = [c for c in self.CHANNELS if self.ax_y2[c].get()]

        self.fig.clear()
        self.ax  = self.fig.add_subplot(111)
        self.ax2 = self.ax.twinx()
        self.fig.subplots_adjust(left=0.11, right=0.88 if y2_chs else 0.95,
                                 top=0.93, bottom=0.10)
        _style_plot(self.fig, self.ax, self.ax2 if y2_chs else None)

        self.ax.set_xlabel(x_ch, color=TEXT, fontsize=9, labelpad=5)
        if y1_chs:
            self.ax.set_ylabel(", ".join(y1_chs), color=Y1_PAL[0],
                               fontsize=9, labelpad=5)
            self.ax.tick_params(axis="y", colors=Y1_PAL[0])
            _eng_formatter(self.ax, "y")
        if y2_chs:
            self.ax2.set_ylabel(", ".join(y2_chs), color=Y2_PAL[0],
                                fontsize=9, labelpad=5)
            self.ax2.tick_params(axis="y", colors=Y2_PAL[0])
            _eng_formatter(self.ax2, "y")
        else:
            self.ax2.set_visible(False)

        lines, labels = [], []
        for curve in self.sweep_data:
            xd = curve.get(x_ch, [])
            if not xd: continue
            sv  = curve.get("step_val", 0)
            sfx = f"  [V₂={sv:.3g}V]" if sv else ""
            for j, ch in enumerate(y1_chs):
                yd = curve.get(ch, [])
                n  = min(len(xd), len(yd))
                if n == 0: continue
                col = Y1_PAL[j % len(Y1_PAL)]
                ln, = self.ax.plot(xd[:n], yd[:n], "o-",
                                   markersize=2.5, linewidth=1.6,
                                   color=col, label=f"{ch}{sfx}")
                lines.append(ln); labels.append(f"{ch}{sfx}")
            for j, ch in enumerate(y2_chs):
                yd = curve.get(ch, [])
                n  = min(len(xd), len(yd))
                if n == 0: continue
                col = Y2_PAL[j % len(Y2_PAL)]
                ln, = self.ax2.plot(xd[:n], yd[:n], "s--",
                                    markersize=2.5, linewidth=1.6,
                                    color=col, label=f"{ch}{sfx} [Y2]")
                lines.append(ln); labels.append(f"{ch}{sfx} [Y2]")

        if lines:
            self.ax.legend(lines, labels, fontsize=7.5, loc="best",
                           framealpha=0.93, edgecolor=BORDER,
                           facecolor=PANEL, labelspacing=0.3)
        self.canvas.draw_idle()

    def _clear_plot(self):
        if self.is_running: return
        if not messagebox.askyesno("Clear", "Clear all sweep data and plot?"):
            return
        self.sweep_data = []
        self.fig.clear()
        self.ax  = self.fig.add_subplot(111)
        self.ax2 = self.ax.twinx()
        self.ax2.set_visible(False)
        self.fig.subplots_adjust(left=0.11,right=0.95,top=0.93,bottom=0.10)
        _style_plot(self.fig, self.ax, title="I-V Characteristic")
        self.canvas.draw()
        self.progress.reset()
        self.status.set("ready")

    def _auto_scale(self):
        try:
            self.ax.relim(); self.ax.autoscale_view()
            if self.ax2.get_visible():
                self.ax2.relim(); self.ax2.autoscale_view()
            self.canvas.draw_idle()
        except: pass

    # ── Measurement thread ────────────────────────────────────────────────────

    def start_thread(self):
        if self.is_running: return
        self.is_running = True
        self._stop_flag = False
        self.sweep_data = []
        self.fig.clear()
        self.ax  = self.fig.add_subplot(111)
        self.ax2 = self.ax.twinx()
        self.ax2.set_visible(False)
        self.fig.subplots_adjust(left=0.11,right=0.95,top=0.93,bottom=0.10)
        _style_plot(self.fig, self.ax, title="Acquiring…")
        self.canvas.draw()
        self.status.set("running")
        self.progress.start()
        threading.Thread(target=self._run_process, daemon=True).start()

    def _run_process(self):
        p = self.inputs
        try:
            self.inst = self.rm.open_resource(self.master_addr)
            self.inst.timeout = 30000

            self.inst.write("*CLS")
            self.inst.write("reset()")
            time.sleep(0.5)
            self.inst.write("defbuffer1.clear()")
            time.sleep(0.1)

            def tsp_check(label=""):
                resp = self.inst.query(
                    "print(errorqueue.count, errorqueue.next())").strip()
                parts = resp.split(",", 1)
                try:
                    n = int(float(parts[0]))
                except (ValueError, IndexError):
                    return
                if n > 0 and len(parts) > 1:
                    raise RuntimeError(f"TSP [{label}]: {parts[1].strip()}")

            self.inst.write(
                f"smu.terminals = smu.TERMINALS_{p['Terminals'].get().upper()}")
            sense = "SENSE_4WIRE" if "4" in p["Sense"].get() else "SENSE_2WIRE"
            self.inst.write(f"smu.measure.sense = smu.{sense}")
            self.inst.write(f"smu.measure.nplc = {p['NPLC'].get()}")
            tsp_check("nplc")

            az = p["AutoZero"].get()
            if az == "Once":
                self.inst.write("smu.measure.autozero.enable = smu.OFF")
                self.inst.write("smu.measure.autozero.once()")
            elif az == "On":
                self.inst.write("smu.measure.autozero.enable = smu.ON")
            else:
                self.inst.write("smu.measure.autozero.enable = smu.OFF")

            hc = "smu.ON" if p["HighC"].get() == "On" else "smu.OFF"
            self.inst.write(f"smu.measure.highcapacitance = {hc}")
            tsp_check("highcapacitance")

            def fmt_range(v):
                return (v.replace("pA","e-12").replace("nA","e-9")
                         .replace("uA","e-6").replace("mA","e-3")
                         .replace("Ω","").replace("kΩ","e3").replace("MΩ","e6")
                         .replace("A",""))

            ri = p["RangeI"].get()
            if ri == "Auto":
                self.inst.write("smu.measure.autorange = smu.ON")
                self.inst.write(
                    f"smu.measure.autorangelow = {fmt_range(p['MinRangeI'].get())}")
            else:
                self.inst.write(f"smu.measure.range = {fmt_range(ri)}")

            mode_str   = p["Mode"].get()
            is_voltage = "Voltage" in mode_str
            if is_voltage:
                self.inst.write("smu.source.func = smu.FUNC_DC_VOLTAGE")
                self.inst.write(f"smu.source.ilimit.level = {p['Limit'].get()}")
            else:
                self.inst.write("smu.source.func = smu.FUNC_DC_CURRENT")
                self.inst.write(f"smu.source.vlimit.level = {p['Limit'].get()}")

            vr = p["RangeV"].get()
            if vr in ("Auto","Best Fixed"):
                self.inst.write("smu.source.autorange = smu.ON")
            else:
                vmap = {"200mV":"0.2","2V":"2","20V":"20","200V":"200"}
                self.inst.write(f"smu.source.range = {vmap.get(vr,'20')}")

            steps = [0.0]
            if p["Stepper"].get():
                steps = np.linspace(
                    float(p["StepStart"].get()),
                    float(p["StepStop"].get()),
                    max(int(p["StepPoints"].get()), 1)
                ).tolist()

            dual_on = p["Dual"].get()
            delay   = p["Delay"].get()

            for step_idx, step_val in enumerate(steps):
                if self._stop_flag: break

                if p["Stepper"].get():
                    self.inst.write(f"node[2].smu.source.level = {step_val}")
                    self.inst.write("node[2].smu.source.output = node[2].smu.ON")
                    self.root.after(0, lambda v=step_val:
                        self.status.msg(f"Stepping  V₂ = {v:.3f} V", WARN))

                if "List" in mode_str:
                    lv = self._get_list_values()
                    if not lv:
                        raise ValueError("List sweep: no values defined.")
                    fwd = ",".join(f"{v:g}" for v in lv)
                    rev = ",".join(f"{v:g}" for v in reversed(lv))
                    passes = [
                        (f"smu.source.sweeplist('sw',{{{fwd}}},{delay},"
                         f"1,smu.RANGE_BEST,smu.OFF)", len(lv))
                    ]
                    if dual_on:
                        passes.append((
                            f"smu.source.sweeplist('sw',{{{rev}}},{delay},"
                            f"1,smu.RANGE_BEST,smu.OFF)", len(lv)))
                else:
                    sv1 = p["Start"].get()
                    sv2 = p["Stop"].get()
                    pts = int(p["Points"].get())
                    if p["PulseEn"].get():
                        on_t  = p["OnTime"].get()
                        off_t = p["OffTime"].get()
                        passes = [
                            (f"smu.source.pulse.sweeplinear('sw',{sv1},{sv2},"
                             f"{pts},{on_t},{off_t},1,smu.RANGE_BEST,smu.OFF)", pts)
                        ]
                        if dual_on:
                            passes.append((
                                f"smu.source.pulse.sweeplinear('sw',{sv2},{sv1},"
                                f"{pts},{on_t},{off_t},1,smu.RANGE_BEST,smu.OFF)", pts))
                    else:
                        passes = [
                            (f"smu.source.sweeplinear('sw',{sv1},{sv2},"
                             f"{pts},{delay},1,smu.RANGE_BEST,smu.OFF)", pts)
                        ]
                        if dual_on:
                            passes.append((
                                f"smu.source.sweeplinear('sw',{sv2},{sv1},"
                                f"{pts},{delay},1,smu.RANGE_BEST,smu.OFF)", pts))

                pts_total = sum(n for _, n in passes)
                live = {
                    "step_val":    step_val,
                    "Voltage (V)": [],
                    "Current (A)": [],
                    "Time (s)":    [],
                    "Resistance (Ω)": [],
                    "Power (W)":   [],
                }
                self.sweep_data.append(live)

                for pass_idx, (pass_cmd, pass_pts) in enumerate(passes):
                    if self._stop_flag:
                        self._safe_abort(); break

                    self.inst.write("defbuffer1.clear()")
                    time.sleep(0.05)
                    self.inst.write(pass_cmd)
                    tsp_check(f"sweep_cmd[{pass_idx}]")
                    self.inst.write("trigger.model.initiate()")
                    tsp_check(f"initiate[{pass_idx}]")

                    last     = 0
                    t0       = time.time()
                    timeout  = pass_pts * (float(p["NPLC"].get())/50.0 +
                                           float(delay)) * 5 + 30

                    while True:
                        if self._stop_flag:
                            self._safe_abort(); break
                        if time.time() - t0 > timeout:
                            raise RuntimeError(
                                f"Pass {pass_idx+1} timed out "
                                f"(collected {last}/{pass_pts} pts)")

                        try:
                            cnt = int(float(
                                self.inst.query("print(defbuffer1.n)").strip()))
                        except ValueError:
                            self.inst.write("*CLS"); time.sleep(0.05); continue

                        if cnt > last:
                            s1, e1 = last+1, cnt
                            nsrc = [float(x) for x in self.inst.query(
                                f"printbuffer({s1},{e1},defbuffer1.sourcevalues)"
                            ).split(",")]
                            nmeas= [float(x) for x in self.inst.query(
                                f"printbuffer({s1},{e1},defbuffer1.readings)"
                            ).split(",")]
                            nts  = [float(x) for x in self.inst.query(
                                f"printbuffer({s1},{e1},defbuffer1.relativetimestamps)"
                            ).split(",")]
                            nv = nsrc  if is_voltage else nmeas
                            ni = nmeas if is_voltage else nsrc
                            live["Voltage (V)"].extend(nv)
                            live["Current (A)"].extend(ni)
                            live["Time (s)"].extend(nts)
                            live["Resistance (Ω)"].extend(
                                [abs(v/i) if i else float("inf")
                                 for v, i in zip(nv, ni)])
                            live["Power (W)"].extend(
                                [abs(v*i) for v, i in zip(nv, ni)])
                            last = cnt
                            self.root.after(0, self._replot)

                        total_so_far = len(live["Voltage (V)"])
                        pct = min(int(100 * total_so_far / pts_total), 100)
                        self.root.after(0, lambda pv=pct: self.progress.update(pv))

                        if cnt >= pass_pts:
                            break
                        time.sleep(0.05)

                self._safe_abort()

            self.inst.write("smu.source.output = smu.OFF")
            if p["Stepper"].get():
                self.inst.write("node[2].smu.source.output = node[2].smu.OFF")
            self.inst.close()

            if self._stop_flag:
                self.root.after(0, lambda: (
                    self.status.set("stopped"),
                    self.progress.stop(False)
                ))
            else:
                self.root.after(0, lambda: (
                    self.status.set("done"),
                    self.progress.stop(True)
                ))
                self.root.after(0, self._auto_save)

        except Exception as e:
            err = str(e)
            self.root.after(0, lambda: (
                messagebox.showerror("Instrument Error", err),
                self.status.set("error"),
                self.progress.stop(False)
            ))
        finally:
            self.is_running = False
            self._stop_flag = False

    # ── CSV / PNG ─────────────────────────────────────────────────────────────

    def _make_filename(self):
        sid = self.sample_id.get().strip() or "DUT"
        return f"{sid}_{time.strftime('%Y%m%d_%H%M%S')}"

    def _build_csv_rows(self):
        p  = self.inputs
        ts = p["Timestamp"].get()
        hs = p["Stepper"].get()
        mr = p["MeasR"].get()
        mp = p["MeasP"].get()
        header = []
        if ts: header.append("Time (s)")
        if hs: header.append("Step_Value (V)")
        header += ["Voltage (V)","Current (A)"]
        if mr: header.append("Resistance (Ω)")
        if mp: header.append("Power (W)")
        rows = []
        for curve in self.sweep_data:
            v  = curve.get("Voltage (V)",[])
            i  = curve.get("Current (A)",[])
            t  = curve.get("Time (s)",[])
            sv = curve.get("step_val", 0)
            for idx in range(min(len(v),len(i))):
                row = []
                if ts: row.append(f"{t[idx]:.6f}" if idx<len(t) else "")
                if hs: row.append(f"{sv:.6f}")
                row += [f"{v[idx]:.9g}", f"{i[idx]:.9g}"]
                if mr:
                    try:    row.append(f"{abs(v[idx]/i[idx]):.6g}")
                    except: row.append("inf")
                if mp: row.append(f"{abs(v[idx]*i[idx]):.9g}")
                rows.append(row)
        return header, rows

    def _write_csv(self, path):
        header, rows = self._build_csv_rows()
        p = self.inputs
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["# Keithley 2450  I-V Studio  v3.0"])
            w.writerow(["# Sample ID",  self.sample_id.get()])
            w.writerow(["# Date/Time",  time.strftime("%d/%m/%Y  %H:%M:%S")])
            w.writerow(["# Mode",       p["Mode"].get()])
            w.writerow(["# NPLC",       p["NPLC"].get()])
            w.writerow(["# Compliance", p["Limit"].get()])
            w.writerow(["# Terminals",  p["Terminals"].get()])
            w.writerow(["# Sense",      p["Sense"].get()])
            w.writerow([])
            w.writerow(header)
            w.writerows(rows)

    def _auto_save(self):
        try:
            base = self._make_filename()
            self._write_csv(os.path.join(self.save_dir, f"{base}.csv"))
            self.fig.savefig(os.path.join(self.save_dir, f"{base}.png"),
                             dpi=300, facecolor=PANEL, bbox_inches="tight")
            self.root.after(0, lambda:
                self.status.msg(f"Auto-saved  →  {self.save_dir}", SUCCESS))
        except Exception as e:
            self.root.after(0, lambda err=str(e):
                messagebox.showwarning("Auto-save Failed",
                                       f"Could not auto-save:\n{err}"))

    def _save_csv(self):
        if not self.sweep_data:
            messagebox.showwarning("No Data","Run a sweep first."); return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv", initialdir=self.save_dir,
            initialfile=f"{self._make_filename()}.csv",
            filetypes=[("CSV","*.csv"),("All","*.*")])
        if not path: return
        try:
            self._write_csv(path)
            messagebox.showinfo("Saved", f"CSV saved:\n{path}")
        except Exception as e:
            messagebox.showerror("Save Error", str(e))

    def _save_png(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".png", initialdir=self.save_dir,
            initialfile=f"{self._make_filename()}.png",
            filetypes=[("PNG","*.png"),("All","*.*")])
        if path:
            try:
                self.fig.savefig(path, dpi=300, facecolor=PANEL,
                                 bbox_inches="tight")
                self.status.msg(f"PNG saved: {os.path.basename(path)}", SUCCESS)
            except Exception as e:
                messagebox.showerror("Save Error", str(e))


# ═══════════════════════════════════════════════════════════════════════════════
#  FET CHARACTERISATION WINDOW
# ═══════════════════════════════════════════════════════════════════════════════

class FETWindow(tk.Toplevel):
    MODES     = ["Transfer Curve  (Id vs Vg)", "Output Curve  (Id vs Vds)"]
    TERMINALS = ["Gate", "Drain", "Source"]
    SMU_OPTS  = ["NONE","GNDU",
                 "SMU1 (Node 1 — Master)","SMU2 (Node 2 — Slave)"]
    OP_MODES  = ["Open","Voltage Bias","Voltage Linear Sweep",
                 "Voltage List Sweep","Voltage Log Sweep","Voltage Step",
                 "Current Bias","Current Linear Sweep","Current List Sweep",
                 "Current Log Sweep","Current Step","Common"]
    RANGE_OPTS= ["Best Fixed","Auto","200mV","2V","20V","200V"]
    OVP_OPTS  = ["OFF","20V","40V","60V","100V","200V"]

    _TSP    = {"SMU1 (Node 1 — Master)":"smu",
               "SMU2 (Node 2 — Slave)": "node[2].smu"}
    _NO_OP  = {"NONE","GNDU"}
    _SWEEP  = {"Voltage Linear Sweep","Current Linear Sweep",
               "Voltage Step","Current Step"}
    _LOG    = {"Voltage Log Sweep","Current Log Sweep"}
    _BIAS   = {"Voltage Bias","Current Bias"}
    _LIST   = {"Voltage List Sweep","Current List Sweep"}

    TERM_COL = {"Gate":"#16A34A","Drain":"#1D4ED8","Source":"#BE185D"}
    TERM_BG  = {"Gate":"#F0FDF4","Drain":"#EFF6FF","Source":"#FDF2F8"}

    def __init__(self, parent_app):
        super().__init__(parent_app.root)
        self.app        = parent_app
        self.title("Keithley 2450  —  FET Characterisation")
        self.geometry("1560x940")
        self.configure(bg=BG)
        self.resizable(True, True)
        self.is_running = False
        self._stop      = False
        self.inst       = None
        self.curves     = []
        self.v           = {t: {} for t in self.TERMINALS}
        self._pf         = {t: {} for t in self.TERMINALS}   # param frames
        self._list_vals  = {t: [] for t in self.TERMINALS}

        self._build_styles()
        self._build_ui()

    def _build_styles(self):
        s = ttk.Style()
        for name, bg, abg in [("FRun.TButton", "#16A34A","#15803D"),
                               ("FStop.TButton","#DC2626","#B91C1C")]:
            s.configure(name, background=bg, foreground="white",
                        font=FONT_H3, padding=(10,6), relief="flat")
            s.map(name, background=[("active",abg)])

    def _build_ui(self):
        # Header
        hdr = tk.Frame(self, bg=ACCENT, height=50)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        logo = tk.Frame(hdr, bg=ACCENT); logo.pack(side="left", padx=14, pady=6)
        tk.Label(logo, text="FET Characterisation", bg=ACCENT, fg="white",
                 font=FONT_H1).pack(anchor="w")
        tk.Label(logo, text="Transfer  &  Output Curve Analysis",
                 bg=ACCENT, fg="#7EB6E0", font=FONT_SM).pack(anchor="w")
        self.status_lbl = tk.Label(hdr, text="●  Ready", bg=ACCENT,
                                   fg="#4ADE80", font=FONT_MONO2)
        self.status_lbl.pack(side="right", padx=16)

        # Mode bar
        mode_bar = tk.Frame(self, bg=PANEL2,
                            highlightbackground=BORDER, highlightthickness=1)
        mode_bar.pack(fill="x", padx=8, pady=(6,0))
        tk.Label(mode_bar, text="Measurement Mode:", bg=PANEL2, fg=ACCENT,
                 font=FONT_H3).pack(side="left", padx=12, pady=7)
        self.mode_var = tk.StringVar(value=self.MODES[0])
        for m in self.MODES:
            tk.Radiobutton(mode_bar, text=m, variable=self.mode_var, value=m,
                           bg=PANEL2, fg=TEXT, selectcolor=PANEL2,
                           activebackground=PANEL2, font=FONT_BODY,
                           command=self._mode_changed).pack(side="left",
                                                            padx=8, pady=7)

        # Body
        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True, padx=8, pady=6)

        # Left: scrollable parameter table
        lf = tk.Frame(body, bg=BG, width=730)
        lf.pack(side="left", fill="y"); lf.pack_propagate(False)
        cvs = tk.Canvas(lf, bg=BG, highlightthickness=0, width=720)
        vsb = ttk.Scrollbar(lf, orient="vertical", command=cvs.yview)
        self._inner = tk.Frame(cvs, bg=BG)
        self._inner.bind("<Configure>",
                         lambda e: cvs.configure(scrollregion=cvs.bbox("all")))
        self._win = cvs.create_window((0,0), window=self._inner, anchor="nw")
        cvs.bind("<Configure>",
                 lambda e: cvs.itemconfig(self._win, width=e.width))
        cvs.configure(yscrollcommand=vsb.set)
        cvs.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        cvs.bind_all("<MouseWheel>",
                     lambda e: cvs.yview_scroll(int(-1*(e.delta/120)),"units"))
        self._build_table(self._inner)

        # Right: graph
        rf = tk.Frame(body, bg=BG)
        rf.pack(side="left", fill="both", expand=True, padx=(8,0))
        self._build_graph(rf)

    # ── Table construction ────────────────────────────────────────────────────

    def _build_table(self, p):
        LBL_W = 22

        # Column headers
        hdr = tk.Frame(p, bg=PANEL,
                       highlightbackground=BORDER, highlightthickness=1)
        hdr.pack(fill="x", pady=(0,2))
        tk.Label(hdr, text="Parameter", bg=PANEL2, fg=ACCENT,
                 font=FONT_H3, width=LBL_W, anchor="w").pack(
                 side="left", padx=(6,0), ipady=6)
        for t in self.TERMINALS:
            col = self.TERM_COL[t]
            f = tk.Frame(hdr, bg=col)
            f.pack(side="left", fill="both", expand=True, padx=2, pady=3)
            tk.Label(f, text=t, bg=col, fg="white",
                     font=FONT_H2).pack(expand=True, pady=5)

        def sec(label):
            f = tk.Frame(p, bg=ACCENT2); f.pack(fill="x", pady=(5,1))
            tk.Label(f, text=f"  {label}", bg=ACCENT2, fg="white",
                     font=FONT_H3).pack(side="left", pady=4)

        def row(label, section=False):
            r = tk.Frame(p, bg=PANEL,
                         highlightbackground=BORDER, highlightthickness=1)
            r.pack(fill="x", pady=1)
            bg = PANEL if not section else PANEL2
            tk.Label(r, text=label, bg=bg, fg=TEXT if not section else ACCENT,
                     font=FONT_BODY, width=LBL_W, anchor="w").pack(
                     side="left", ipady=4, padx=(6,0))
            cells = {}
            for t in self.TERMINALS:
                col = self.TERM_COL[t]
                bg2 = self.TERM_BG[t]
                f = tk.Frame(r, bg=PANEL,
                             highlightbackground=col, highlightthickness=1)
                f.pack(side="left", fill="both", expand=True, padx=2, pady=2)
                cells[t] = f
            return cells

        # ── Instrument ────────────────────────────────────────────────────────
        sec("Instrument Assignment")
        cr = row("Instrument", section=True)
        defaults_smu = {"Gate":"SMU1 (Node 1 — Master)",
                        "Drain":"SMU2 (Node 2 — Slave)",
                        "Source":"GNDU"}
        for t in self.TERMINALS:
            self.v[t]["smu"] = tk.StringVar(value=defaults_smu[t])
            ttk.Combobox(cr[t], textvariable=self.v[t]["smu"],
                         values=self.SMU_OPTS, state="readonly",
                         width=20, font=FONT_SM).pack(fill="x", padx=4, pady=4)
            self.v[t]["smu"].trace_add("write",
                lambda *_, term=t: self._on_smu_change(term))

        # ── Force ─────────────────────────────────────────────────────────────
        sec("Force")
        op_r = row("Operation Mode")
        op_defaults = {"Gate":"Voltage Linear Sweep",
                       "Drain":"Voltage Bias","Source":"Voltage Bias"}
        for t in self.TERMINALS:
            self.v[t]["op"] = tk.StringVar(value=op_defaults[t])
            ttk.Combobox(op_r[t], textvariable=self.v[t]["op"],
                         values=self.OP_MODES, state="readonly",
                         width=20, font=FONT_SM).pack(fill="x", padx=4, pady=4)
            self.v[t]["op"].trace_add("write",
                lambda *_, term=t: self._on_op_change(term))

        def entry_row(label, key, default_d, unit="V"):
            r = row(label)
            for t in self.TERMINALS:
                self.v[t][key] = tk.StringVar(value=default_d.get(t,"0"))
                f = tk.Frame(r[t], bg=PANEL); f.pack(fill="x", padx=4, pady=3)
                ttk.Entry(f, textvariable=self.v[t][key],
                          width=10, font=FONT_SM).pack(side="left")
                tk.Label(f, text=unit, bg=PANEL, fg=TEXT3,
                         font=FONT_SM).pack(side="left", padx=3)
            return r

        # Bias row
        bias_r = entry_row("Bias", "bias", {"Gate":"0","Drain":"10","Source":"0"})
        for t in self.TERMINALS: self._pf[t]["bias_row"] = bias_r[t]

        # List Values row
        list_r = row("List Values")
        for t in self.TERMINALS:
            cell = list_r[t]
            lbl = tk.Label(cell, text="(empty)", bg=SEP, fg=TEXT3,
                           font=FONT_SM, anchor="w")
            lbl.pack(side="left", padx=5, fill="x", expand=True)
            btn = ttk.Button(cell, text="Edit…",
                             command=lambda term=t, lb=lbl:
                                 self._open_list_editor(term, lb))
            btn.pack(side="right", padx=4, pady=4)
            self._pf[t]["list_lbl"] = lbl
            self._pf[t]["list_row"] = list_r[t]
            cell.configure(bg=SEP)
            lbl.configure(state="disabled"); btn.configure(state="disabled")

        start_r = entry_row("Start", "start", {"Gate":"0","Drain":"10","Source":"0"})
        stop_r  = entry_row("Stop",  "stop",  {"Gate":"8","Drain":"10","Source":"0"})
        step_r  = entry_row("Step",  "step",  {"Gate":"0.1","Drain":"1","Source":"0"})

        pts_r = row("Points")
        for t in self.TERMINALS:
            self.v[t]["pts"]      = tk.StringVar(value="81" if t=="Gate" else "76")
            self.v[t]["pts_auto"] = tk.BooleanVar(value=True)
            f = tk.Frame(pts_r[t], bg=PANEL); f.pack(fill="x", padx=4, pady=3)

            # Entry (editable when manual)
            ent = ttk.Entry(f, textvariable=self.v[t]["pts"], width=8, font=FONT_SM)
            ent.pack(side="left")
            self._pf[t]["pts_entry"] = ent

            # Auto label shown when auto mode active
            auto_lbl = tk.Label(f, text="auto", bg=PANEL2, fg=ACCENT2,
                                font=FONT_SM, padx=4)
            auto_lbl.pack(side="left", padx=(2,0))
            self._pf[t]["pts_auto_lbl"] = auto_lbl

            # Toggle checkbox
            def _make_toggle(term, entry_widget, lbl_widget):
                def _toggle(*_):
                    if self.v[term]["pts_auto"].get():
                        entry_widget.configure(state="disabled")
                        lbl_widget.pack(side="left", padx=(2,0))
                        self._recalc_pts_for(term)
                    else:
                        lbl_widget.pack_forget()
                        entry_widget.configure(state="normal")
                return _toggle

            toggle_fn = _make_toggle(t, ent, auto_lbl)
            self.v[t]["pts_auto"].trace_add("write", lambda *_, fn=toggle_fn: fn())

            chk = tk.Checkbutton(f, text="Auto", variable=self.v[t]["pts_auto"],
                                  bg=PANEL, fg=TEXT2, activebackground=PANEL,
                                  selectcolor=PANEL, relief="flat",
                                  highlightthickness=0, font=FONT_SM)
            chk.pack(side="left", padx=(6,0))

            # Wire start/stop/step changes → recalc when auto
            for k in ("start","stop","step"):
                self.v[t][k].trace_add("write",
                    lambda *_, term=t: self._recalc_pts_for(term))

            # Initial state
            toggle_fn()

        dual_r = row("Dual Sweep")
        for t in self.TERMINALS:
            self.v[t]["dual"] = tk.BooleanVar(value=False)
            f = tk.Frame(dual_r[t], bg=PANEL); f.pack(fill="x", padx=4, pady=3)
            tk.Checkbutton(f, variable=self.v[t]["dual"], bg=PANEL,
                           activebackground=PANEL, selectcolor=PANEL,
                           relief="flat", highlightthickness=0).pack(side="left")

        for t in self.TERMINALS:
            self._pf[t]["start_row"] = start_r[t]
            self._pf[t]["stop_row"]  = stop_r[t]
            self._pf[t]["step_row"]  = step_r[t]
            self._pf[t]["pts_row"]   = pts_r[t]

        # Range, Compliance, Delay, OVP
        rng_r = row("Source Range")
        for t in self.TERMINALS:
            self.v[t]["range"] = tk.StringVar(value="Best Fixed")
            ttk.Combobox(rng_r[t], textvariable=self.v[t]["range"],
                         values=self.RANGE_OPTS, state="readonly",
                         width=13, font=FONT_SM).pack(fill="x", padx=4, pady=4)

        comp_r = entry_row("Compliance",      "comp",
                           {"Gate":"0.01","Drain":"0.5","Source":"0.1"}, unit="A")
        delay_r= entry_row("Source Delay",    "delay",
                           {"Gate":"0.05","Drain":"0.05","Source":"0.05"}, unit="s")
        pod_r  = entry_row("Power-On Delay",  "pod",
                           {"Gate":"0","Drain":"0","Source":"0"}, unit="s")
        ovp_r  = row("Overvoltage Protect.")
        for t in self.TERMINALS:
            self.v[t]["ovp"] = tk.StringVar(value="OFF")
            ttk.Combobox(ovp_r[t], textvariable=self.v[t]["ovp"],
                         values=self.OVP_OPTS, state="readonly",
                         width=13, font=FONT_SM).pack(fill="x", padx=4, pady=4)

        # ── Measure ───────────────────────────────────────────────────────────
        sec("Measure")

        def chk_row(label, key, defaults):
            r = row(label)
            for t in self.TERMINALS:
                self.v[t][key] = tk.BooleanVar(value=defaults.get(t, True))
                f = tk.Frame(r[t], bg=PANEL); f.pack(fill="x", padx=4, pady=3)
                tk.Checkbutton(f, variable=self.v[t][key], bg=PANEL,
                               activebackground=PANEL, selectcolor=PANEL,
                               relief="flat", highlightthickness=0
                               ).pack(side="left")
            return r

        def col_name_row(label, key, defaults):
            r = row(label)
            for t in self.TERMINALS:
                self.v[t][key] = tk.StringVar(value=defaults.get(t,""))
                ttk.Entry(r[t], textvariable=self.v[t][key],
                          width=14, font=FONT_SM).pack(fill="x", padx=4, pady=3)

        chk_row("Current Measure",  "meas_i",
                {"Gate":False,"Drain":True,"Source":True})
        col_name_row("  Column Name",   "icol",
                     {"Gate":"GateI","Drain":"DrainI","Source":"SourceI"})
        meas_irange_r = row("  Current Range")
        for t in self.TERMINALS:
            self.v[t]["irange"] = tk.StringVar(value="Limited Auto")
            ttk.Combobox(meas_irange_r[t], textvariable=self.v[t]["irange"],
                         values=["Limited Auto","Auto","1nA","10nA","100nA",
                                 "1uA","10uA","100uA","1mA","10mA","100mA","1A"],
                         state="readonly", width=13,
                         font=FONT_SM).pack(fill="x", padx=4, pady=4)

        chk_row("Voltage Measure", "meas_v", {t:True for t in self.TERMINALS})
        vrep_r = row("  Report Value")
        for t in self.TERMINALS:
            self.v[t]["vrep"] = tk.StringVar(value="Programmed")
            ttk.Combobox(vrep_r[t], textvariable=self.v[t]["vrep"],
                         values=["Programmed","Actual"],
                         state="readonly", width=13,
                         font=FONT_SM).pack(fill="x", padx=4, pady=4)
        col_name_row("  Column Name", "vcol",
                     {"Gate":"GateV","Drain":"DrainV","Source":"SourceV"})

        # ── Instrument Settings ───────────────────────────────────────────────
        sec("Instrument Settings")
        nplc_r = row("NPLC")
        for t in self.TERMINALS:
            self.v[t]["nplc"] = tk.StringVar(value="1")
            ttk.Entry(nplc_r[t], textvariable=self.v[t]["nplc"],
                      width=10, font=FONT_SM).pack(fill="x", padx=4, pady=3)

        # ── Controls ─────────────────────────────────────────────────────────
        ctrl = tk.Frame(p, bg=PANEL,
                        highlightbackground=BORDER, highlightthickness=1)
        ctrl.pack(fill="x", pady=(8,0))

        row1 = tk.Frame(ctrl, bg=PANEL); row1.pack(fill="x", padx=10, pady=(8,4))
        tk.Label(row1, text="VISA Address", bg=PANEL, fg=TEXT,
                 font=FONT_BODY, width=16, anchor="w").pack(side="left")
        self.addr_var = tk.StringVar(value=self.app.master_addr)
        ttk.Entry(row1, textvariable=self.addr_var,
                  width=38, font=FONT_BODY).pack(side="left", padx=6)

        row2 = tk.Frame(ctrl, bg=PANEL); row2.pack(fill="x", padx=10, pady=(0,4))
        tk.Label(row2, text="Input Jacks", bg=PANEL, fg=TEXT,
                 font=FONT_BODY, width=16, anchor="w").pack(side="left")
        self.terminals_var = tk.StringVar(value="Rear")
        ttk.Combobox(row2, textvariable=self.terminals_var,
                     values=["Rear", "Front"],
                     state="readonly", width=8,
                     font=FONT_BODY).pack(side="left", padx=(0, 18))
        tk.Label(row2, text="Sense Mode", bg=PANEL, fg=TEXT,
                 font=FONT_BODY).pack(side="left")
        self.sense_var = tk.StringVar(value="2-Wire (Local)")
        ttk.Combobox(row2, textvariable=self.sense_var,
                     values=["2-Wire (Local)","4-Wire (Remote)"],
                     state="readonly", width=16,
                     font=FONT_BODY).pack(side="left", padx=6)
        self.nplc_all = tk.StringVar(value="1")
        tk.Label(row2, text="NPLC (all)", bg=PANEL, fg=TEXT,
                 font=FONT_BODY).pack(side="left", padx=(20,4))
        ttk.Entry(row2, textvariable=self.nplc_all,
                  width=6, font=FONT_BODY).pack(side="left")
        ttk.Button(row2, text="Apply to all",
                   command=self._apply_nplc_all).pack(side="left", padx=6)

        row3 = tk.Frame(ctrl, bg=PANEL); row3.pack(fill="x", padx=10, pady=(0,8))
        ttk.Button(row3, text="▶  RUN", style="FRun.TButton",
                   command=self._run).pack(side="left", fill="x",
                                           expand=True, ipady=5, padx=(0,6))
        ttk.Button(row3, text="■  STOP", style="FStop.TButton",
                   command=self._do_stop).pack(side="left", ipady=5, ipadx=12)
        ttk.Button(row3, text="Export CSV",
                   command=self._export_csv).pack(side="left", padx=(10,3))
        ttk.Button(row3, text="Save PNG",
                   command=self._save_png).pack(side="left")

        # Initial visibility
        for t in self.TERMINALS:
            self._on_smu_change(t)
            self._on_op_change(t)

    # ── FET visibility callbacks ──────────────────────────────────────────────

    def _on_smu_change(self, t):
        sel   = self.v[t]["smu"].get()
        no_op = sel in self._NO_OP
        for key in ("bias_row","list_row","start_row","stop_row","step_row","pts_row"):
            f = self._pf[t].get(key)
            if f:
                for w in f.winfo_children():
                    try: w.configure(state="disabled" if no_op else "normal")
                    except: pass
                f.configure(bg=SEP if no_op else PANEL)
        if not no_op:
            self._on_op_change(t)

    def _on_op_change(self, t):
        if self.v[t]["smu"].get() in self._NO_OP: return
        op = self.v[t]["op"].get()
        is_list  = op in self._LIST
        is_bias  = op in self._BIAS
        is_sweep = op in self._SWEEP
        is_log   = op in self._LOG

        vis = {
            "bias_row":  is_bias,
            "list_row":  is_list,
            "start_row": is_sweep or is_log,
            "stop_row":  is_sweep or is_log,
            "step_row":  is_sweep,
            "pts_row":   is_sweep or is_log,
        }
        for key, show in vis.items():
            f = self._pf[t].get(key)
            if not f: continue
            bg = PANEL if show else SEP
            f.configure(bg=bg)
            for w in f.winfo_children():
                try: w.configure(state="normal" if show else "disabled")
                except: pass
            if key == "list_row":
                lbl = self._pf[t].get("list_lbl")
                if lbl: lbl.configure(bg=bg)

        if t == "Gate": self._recalc_pts()

    def _recalc_pts(self, *_):
        """Legacy shim — recalculates all auto-mode terminals."""
        for t in self.TERMINALS:
            self._recalc_pts_for(t)

    def _recalc_pts_for(self, term, *_):
        """Recalculate point count for one terminal if it is in Auto mode."""
        try:
            if not self.v[term]["pts_auto"].get():
                return
            s  = float(self.v[term]["start"].get())
            e  = float(self.v[term]["stop"].get())
            st = float(self.v[term]["step"].get())
            if st == 0: return
            n = int(abs(e - s) / abs(st)) + 1
            self.v[term]["pts"].set(str(n))
        except:
            pass

    def _apply_nplc_all(self):
        val = self.nplc_all.get().strip()
        for t in self.TERMINALS:
            self.v[t]["nplc"].set(val)

    def _open_list_editor(self, terminal, lbl_widget):
        def on_save(vals):
            self._list_vals[terminal] = vals
            n = len(vals)
            lbl_widget.config(text=f"{n} value{'s' if n!=1 else ''}"
                              if n else "(empty)")
        ListEditor(self,
                   title=f"{terminal}  —  List Values",
                   initial=self._list_vals[terminal],
                   accent=self.TERM_COL[terminal],
                   on_save=on_save)

    # ── Graph ─────────────────────────────────────────────────────────────────

    def _build_graph(self, parent):
        info = tk.Frame(parent, bg=PANEL,
                        highlightbackground=BORDER, highlightthickness=1)
        info.pack(fill="x", pady=(0,4))
        self.info_lbl = tk.Label(info, text="Configure terminals and press RUN",
                                 bg=PANEL, fg=TEXT3, font=FONT_BODY, anchor="w")
        self.info_lbl.pack(side="left", padx=10, pady=5)

        self.curve_count = tk.Label(info, text="", bg=PANEL,
                                    fg=ACCENT2, font=FONT_MONO)
        self.curve_count.pack(side="right", padx=10)

        # Progress bar
        prog_f = tk.Frame(parent, bg=PANEL,
                          highlightbackground=BORDER, highlightthickness=1)
        prog_f.pack(fill="x", pady=(0,4))
        prow = tk.Frame(prog_f, bg=PANEL); prow.pack(fill="x", padx=10, pady=5)
        self.progress = ProgressRow(prow)
        self.progress.pack(fill="x")

        # Toolbar buttons
        tb = tk.Frame(parent, bg=PANEL,
                      highlightbackground=BORDER, highlightthickness=1)
        tb.pack(fill="x", pady=(0,4))
        ttk.Button(tb, text="Clear", command=self._clear_curves).pack(
            side="left", padx=6, pady=4)
        ttk.Button(tb, text="Log Y Scale",
                   command=self._toggle_log).pack(side="left", padx=(0,4))
        ttk.Button(tb, text="Auto Scale",
                   command=self._auto_scale).pack(side="left", padx=(0,4))
        self._log_y = False

        # ── Q-Point panel ─────────────────────────────────────────────────────
        self._build_qpoint_panel(parent)

        pf = tk.Frame(parent, bg=PANEL,
                      highlightbackground=BORDER, highlightthickness=1)
        pf.pack(fill="both", expand=True)

        self.fig = Figure(figsize=(7, 5.5), dpi=100, facecolor=PANEL)
        self.fig.subplots_adjust(left=0.11, right=0.95, top=0.93, bottom=0.10)
        self.ax  = self.fig.add_subplot(111)
        _style_plot(self.fig, self.ax,
                    title="Configure terminals and press RUN")

        self.canvas = FigureCanvasTkAgg(self.fig, master=pf)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(fill="both", expand=True,
                                         padx=2, pady=2)

    # ── Q-Point panel ────────────────────────────────────────────────────────

    def _build_qpoint_panel(self, parent):
        """Build the Q-point analysis bar between toolbar and chart."""
        qf = tk.Frame(parent, bg=PANEL,
                      highlightbackground=BORDER2, highlightthickness=1)
        qf.pack(fill="x", pady=(0, 4))

        # Header
        hdr = tk.Frame(qf, bg=ACCENT)
        hdr.pack(fill="x")
        tk.Label(hdr, text="  ⊙  Q-Point  Analysis",
                 bg=ACCENT, fg="white", font=FONT_H3).pack(side="left", pady=4)
        self.qp_enabled = tk.BooleanVar(value=False)
        tk.Checkbutton(hdr, text="Enable", variable=self.qp_enabled,
                       bg=ACCENT, fg="white", selectcolor=ACCENT,
                       activebackground=ACCENT, font=FONT_SM,
                       command=self._replot).pack(side="right", padx=10)

        bdy = tk.Frame(qf, bg=PANEL)
        bdy.pack(fill="x", padx=10, pady=6)

        # ── Transfer inputs (Vgs_Q) ───────────────────────────────────────────
        self._qp_xfr_f = tk.Frame(bdy, bg=PANEL)
        tk.Label(self._qp_xfr_f, text="Vgs_Q", bg=PANEL, fg=TEXT,
                 font=FONT_BODY).pack(side="left")
        self.qp_vgs = tk.StringVar(value="4.0")
        ttk.Entry(self._qp_xfr_f, textvariable=self.qp_vgs,
                  width=7, font=FONT_BODY).pack(side="left", padx=(4, 2))
        tk.Label(self._qp_xfr_f, text="V", bg=PANEL, fg=TEXT3,
                 font=FONT_SM).pack(side="left", padx=(0, 16))

        # ── Output inputs (Vdd, RL, curve select) ────────────────────────────
        self._qp_out_f = tk.Frame(bdy, bg=PANEL)

        tk.Label(self._qp_out_f, text="Vdd", bg=PANEL, fg=TEXT,
                 font=FONT_BODY).pack(side="left")
        self.qp_vdd = tk.StringVar(value="20.0")
        ttk.Entry(self._qp_out_f, textvariable=self.qp_vdd,
                  width=7, font=FONT_BODY).pack(side="left", padx=(4, 2))
        tk.Label(self._qp_out_f, text="V", bg=PANEL, fg=TEXT3,
                 font=FONT_SM).pack(side="left", padx=(0, 16))

        tk.Label(self._qp_out_f, text="RL", bg=PANEL, fg=TEXT,
                 font=FONT_BODY).pack(side="left")
        self.qp_rl = tk.StringVar(value="100.0")
        ttk.Entry(self._qp_out_f, textvariable=self.qp_rl,
                  width=7, font=FONT_BODY).pack(side="left", padx=(4, 2))
        tk.Label(self._qp_out_f, text="Ω", bg=PANEL, fg=TEXT3,
                 font=FONT_SM).pack(side="left", padx=(0, 16))

        tk.Label(self._qp_out_f, text="Vg curve", bg=PANEL, fg=TEXT,
                 font=FONT_BODY).pack(side="left")
        self.qp_curve_var = tk.StringVar(value="")
        self._qp_combo = ttk.Combobox(self._qp_out_f,
                                       textvariable=self.qp_curve_var,
                                       values=[], state="readonly",
                                       width=14, font=FONT_SM)
        self._qp_combo.pack(side="left", padx=(4, 16))

        # Apply button
        ttk.Button(bdy, text="Apply", style="Accent.TButton",
                   command=self._replot).pack(side="left", padx=(0, 12))

        # Result readout
        self.qp_result_lbl = tk.Label(bdy, text="—", bg=PANEL, fg=ACCENT2,
                                       font=FONT_MONO)
        self.qp_result_lbl.pack(side="left", fill="x")

        # Show the correct frame for the current mode
        self._qp_sync_mode()

    def _qp_sync_mode(self):
        """Show Transfer or Output Q-point input fields based on current mode."""
        is_xfr = "Transfer" in self.mode_var.get()
        if is_xfr:
            self._qp_out_f.pack_forget()
            self._qp_xfr_f.pack(side="left")
        else:
            self._qp_xfr_f.pack_forget()
            self._qp_out_f.pack(side="left")

    # ── Q-Point drawing ───────────────────────────────────────────────────────

    def _draw_qpoint(self):
        """Dispatch to Transfer or Output Q-point renderer."""
        if not self.qp_enabled.get() or not self.curves:
            self.after(0, lambda: self.qp_result_lbl.config(text="—", fg=ACCENT2))
            return
        is_xfr = "Transfer" in self.mode_var.get()
        try:
            if is_xfr:
                self._draw_qpoint_transfer()
            else:
                self._draw_qpoint_output()
        except Exception as ex:
            msg = str(ex)
            self.after(0, lambda m=msg: self.qp_result_lbl.config(
                text=f"Q-point: {m}", fg=DANGER))

    def _draw_qpoint_transfer(self):
        """
        Transfer Curve Q-point: mark (Vgs_Q, Id_Q) and compute gm.

        gm  = ΔId / ΔVgs  at the Q-point  (numerical gradient)
        """
        vgs_q = float(self.qp_vgs.get())
        c = self.curves[0]
        x = np.array(c["x"],  dtype=float)
        y = np.array([abs(v) for v in c["y"]], dtype=float)
        if len(x) < 2:
            return

        # Sort ascending for safe interpolation
        order = np.argsort(x)
        x, y  = x[order], y[order]

        id_q = float(np.interp(vgs_q, x, y))

        # gm: numerical gradient smoothed over ±2 neighbours
        dy   = np.gradient(y, x)
        gm_q = float(np.interp(vgs_q, x, dy))

        # ── Plot elements ─────────────────────────────────────────────────────
        # Crosshair dashed lines
        self.ax.axvline(vgs_q, color="#F59E0B", lw=1.2, ls="--",
                        alpha=0.75, zorder=5)
        self.ax.axhline(id_q,  color="#F59E0B", lw=1.2, ls="--",
                        alpha=0.75, zorder=5)

        # Star marker
        self.ax.plot(vgs_q, id_q, "*", color="#F59E0B", markersize=15,
                     zorder=11, markeredgecolor="#92400E", markeredgewidth=0.8,
                     label=f"Q  Vgs = {vgs_q:.3g} V,  Id = {_eng_str(id_q)}A")

        # Annotation box
        ann = (f"  Vgs = {vgs_q:.3g} V\n"
               f"  Id  = {_eng_str(id_q)}A\n"
               f"  gm  = {_eng_str(gm_q)}S")
        self.ax.annotate(
            ann, xy=(vgs_q, id_q),
            xytext=(20, 18), textcoords="offset points",
            fontsize=8, color=WARN, fontfamily="Consolas",
            bbox=dict(boxstyle="round,pad=0.5", facecolor=WARN2,
                      edgecolor="#F59E0B", alpha=0.95, linewidth=1.2),
            arrowprops=dict(arrowstyle="->", color="#F59E0B",
                            connectionstyle="arc3,rad=0.1", lw=1.2))

        result = (f"Q:  Vgs = {vgs_q:.3g} V    "
                  f"Id = {_eng_str(id_q)}A    "
                  f"gm = {_eng_str(gm_q)}S")
        self.after(0, lambda r=result: self.qp_result_lbl.config(
            text=r, fg=ACCENT2))

    def _draw_qpoint_output(self):
        """
        Output Curve Q-point: draw load line, mark intersection, show Pd hyperbola.

        Load line:   Id = (Vdd − Vds) / RL
        Q-point:     intersection of load line with selected Vg curve
        Pd hyperbola: Id = Pd_Q / Vds  (constant-power curve through Q)
        """
        vdd = float(self.qp_vdd.get())
        rl  = float(self.qp_rl.get())
        if rl <= 0:
            raise ValueError("RL must be > 0 Ω")
        if vdd <= 0:
            raise ValueError("Vdd must be > 0 V")

        eng = mticker.EngFormatter(sep="")

        # ── Load line ─────────────────────────────────────────────────────────
        id_max   = vdd / rl
        vds_line = np.array([0.0, vdd])
        id_line  = np.array([id_max, 0.0])
        self.ax.plot(vds_line, id_line, color="#F59E0B", lw=2.0, ls="-.",
                     zorder=8,
                     label=f"Load Line  Vdd = {vdd:.3g} V,  RL = {eng(rl)}Ω")
        # Intercept labels on axes
        self.ax.annotate(f"{eng(vdd)}V", xy=(vdd, 0),
                         xytext=(0, 8), textcoords="offset points",
                         fontsize=7.5, color="#92400E", ha="center")
        self.ax.annotate(f"{eng(id_max)}A", xy=(0, id_max),
                         xytext=(8, 0), textcoords="offset points",
                         fontsize=7.5, color="#92400E", va="center")

        # ── Select curve ──────────────────────────────────────────────────────
        sel   = self.qp_curve_var.get()
        curve = next((c for c in self.curves if c["label"] == sel), self.curves[0])

        x = np.array(curve["x"],  dtype=float)
        y = np.array([abs(v) for v in curve["y"]], dtype=float)
        if len(x) < 2:
            return
        order = np.argsort(x)
        x, y  = x[order], y[order]

        # ── Find intersection (load line vs curve) ────────────────────────────
        id_ll = (vdd - x) / rl
        diff  = y - id_ll
        signs = np.sign(diff)
        xings = np.where(np.diff(signs))[0]

        if len(xings) == 0:
            result = "No Q-point intersection on selected curve"
            self.after(0, lambda r=result: self.qp_result_lbl.config(
                text=r, fg=DANGER))
            return

        ci = xings[-1]                       # use last (saturation-region) crossing
        x0, x1 = x[ci], x[ci + 1]
        d0, d1  = diff[ci], diff[ci + 1]
        vds_q = float(x0 - d0 * (x1 - x0) / (d1 - d0))
        id_q  = float(np.interp(vds_q, x, y))
        p_q   = vds_q * id_q

        # ── Pd constant-power hyperbola through Q ─────────────────────────────
        vds_hyp = np.linspace(max(0.05 * vdd, 0.05), vdd * 0.98, 300)
        id_hyp  = p_q / vds_hyp
        self.ax.plot(vds_hyp, id_hyp, color="#DC2626", lw=1.2, ls=":",
                     alpha=0.65, zorder=4,
                     label=f"Pd = {eng(p_q)}W  hyperbola")

        # ── Crosshairs & marker ───────────────────────────────────────────────
        self.ax.axvline(vds_q, color="#F59E0B", lw=1.2, ls="--",
                        alpha=0.75, zorder=5)
        self.ax.axhline(id_q,  color="#F59E0B", lw=1.2, ls="--",
                        alpha=0.75, zorder=5)
        self.ax.plot(vds_q, id_q, "*", color="#F59E0B", markersize=15,
                     zorder=11, markeredgecolor="#92400E", markeredgewidth=0.8,
                     label=f"Q  Vds = {vds_q:.3g} V,  Id = {eng(id_q)}A")

        # Annotation
        ann = (f"  {curve['label']}\n"
               f"  Vds = {vds_q:.3g} V\n"
               f"  Id  = {eng(id_q)}A\n"
               f"  Pd  = {eng(p_q)}W")
        self.ax.annotate(
            ann, xy=(vds_q, id_q),
            xytext=(20, 18), textcoords="offset points",
            fontsize=8, color=WARN, fontfamily="Consolas",
            bbox=dict(boxstyle="round,pad=0.5", facecolor=WARN2,
                      edgecolor="#F59E0B", alpha=0.95, linewidth=1.2),
            arrowprops=dict(arrowstyle="->", color="#F59E0B",
                            connectionstyle="arc3,rad=-0.1", lw=1.2))

        result = (f"Q:  Vds = {vds_q:.3g} V    "
                  f"Id = {eng(id_q)}A    "
                  f"Pd = {eng(p_q)}W")
        self.after(0, lambda r=result: self.qp_result_lbl.config(
            text=r, fg=ACCENT2))

    def _replot(self):
        self.fig.clear()
        self.ax = self.fig.add_subplot(111)
        self.fig.subplots_adjust(left=0.11,right=0.95,top=0.93,bottom=0.10)
        mode   = self.mode_var.get()
        is_xfr = "Transfer" in mode
        xlabel = "Gate Voltage  Vg (V)" if is_xfr else "Drain Voltage  Vds (V)"
        ylabel = "|Drain Current  Id| (A)"
        title  = ("Transfer Curve  —  Id vs Vg" if is_xfr
                  else "Output Curve  —  Id vs Vds")
        _style_plot(self.fig, self.ax, title=title)
        self.ax.set_xlabel(xlabel, color=TEXT, fontsize=9)
        self.ax.set_ylabel(ylabel, color=TEXT, fontsize=9)
        _eng_formatter(self.ax, "x")
        if not self._log_y:
            _eng_formatter(self.ax, "y")

        # ── Update Q-point Output curve selector ──────────────────────────────
        if hasattr(self, "_qp_combo") and self.curves:
            labels = [c["label"] for c in self.curves]
            self._qp_combo["values"] = labels
            if self.qp_curve_var.get() not in labels:
                self.qp_curve_var.set(labels[0])

        # ── Plot curves ───────────────────────────────────────────────────────
        for i, c in enumerate(self.curves):
            col   = FET_PAL[i % len(FET_PAL)]
            x     = c["x"]
            y_raw = [abs(v) for v in c["y"]]
            y     = ([v if v > 1e-15 else float("nan") for v in y_raw]
                     if self._log_y else y_raw)
            n = min(len(x), len(y))
            if n == 0: continue
            self.ax.plot(x[:n], y[:n], "o-", markersize=2.5,
                         linewidth=1.8, color=col, label=c["label"])

        if self._log_y:
            try:
                self.ax.set_yscale("log")
                import matplotlib.ticker as _mticker
                self.ax.yaxis.set_major_formatter(
                    _mticker.LogFormatterSciNotation(labelOnlyBase=False))
                self.ax.yaxis.set_minor_formatter(
                    _mticker.LogFormatterSciNotation(labelOnlyBase=False,
                                                     minor_thresholds=(2, 0.4)))
            except Exception:
                pass

        # ── Q-point overlay (before legend so it appears in it) ───────────────
        self._draw_qpoint()

        if self.curves:
            self.ax.legend(fontsize=7.5, loc="best", framealpha=0.93,
                           edgecolor=BORDER, facecolor=PANEL,
                           labelspacing=0.3)
        self.canvas.draw_idle()
        n = len(self.curves)
        self.curve_count.config(text=f"{n} curve{'s' if n!=1 else ''}" if n else "")

    def _mode_changed(self):
        """When mode changes, auto-set appropriate op modes for each column."""
        mode = self.mode_var.get()
        is_xfr = "Transfer" in mode
        if is_xfr:
            self.v["Gate"]["op"].set("Voltage Linear Sweep")
            self.v["Drain"]["op"].set("Voltage Bias")
        else:
            self.v["Gate"]["op"].set("Voltage Bias")
            self.v["Drain"]["op"].set("Voltage Linear Sweep")
        for t in self.TERMINALS:
            self._on_op_change(t)
        if hasattr(self, "_qp_xfr_f"):
            self._qp_sync_mode()
        self._replot()

    def _clear_curves(self):
        if self.is_running: return
        self.curves = []
        self.fig.clear()
        self.ax = self.fig.add_subplot(111)
        self.fig.subplots_adjust(left=0.11,right=0.95,top=0.93,bottom=0.10)
        _style_plot(self.fig, self.ax,
                    title="Configure terminals and press RUN")
        self.canvas.draw()
        self.curve_count.config(text="")
        self.progress.reset()

    def _auto_scale(self):
        try:
            self.ax.relim(); self.ax.autoscale_view()
            self.canvas.draw_idle()
        except: pass

    def _toggle_log(self):
        self._log_y = not self._log_y
        self._replot()

    # ── Measurement helpers ───────────────────────────────────────────────────

    def _fv(self, t, k, fb=0.0):
        try: return float(self.v[t][k].get())
        except: return fb

    def _step_list(self, start, stop, step):
        if step == 0: return [start]
        if (stop-start)*step < 0: step = -step
        vals, v = [], start
        while ((step>0 and v<=stop+1e-9) or (step<0 and v>=stop-1e-9)):
            vals.append(round(v,9)); v += step
        return vals

    def _resolve(self, t):
        sel = self.v[t]["smu"].get()
        return self._TSP.get(sel, None), sel not in self._NO_OP

    def _set_status(self, text, col="#4ADE80"):
        self.after(0, lambda: self.status_lbl.config(
            text=f"●  {text}", fg=col))

    # ── Run / Stop ────────────────────────────────────────────────────────────

    def _run(self):
        if self.is_running: return
        is_xfr = "Transfer" in self.mode_var.get()
        pterm  = "Gate" if is_xfr else "Drain"
        _, real = self._resolve(pterm)
        if not real:
            messagebox.showerror("Configuration Error",
                f"'{pterm}' must be assigned to SMU1 or SMU2.")
            return
        self.is_running = True
        self._stop      = False
        self.curves     = []
        self.fig.clear()
        self.ax = self.fig.add_subplot(111)
        self.fig.subplots_adjust(left=0.11,right=0.95,top=0.93,bottom=0.10)
        _style_plot(self.fig, self.ax, title="Acquiring…")
        self.canvas.draw()
        self._set_status("Acquiring…", "#F59E0B")
        self.progress.start()
        threading.Thread(target=self._measure, daemon=True).start()

    def _do_stop(self):
        self._stop = True
        self._set_status("Stopping…", "#EF4444")
        self.after(0, lambda: self.progress.stop(False))
        try:
            if self.inst:
                self.inst.write(TSP_SAFE_ABORT)
                for t in self.TERMINALS:
                    tsp, real = self._resolve(t)
                    if real: self.inst.write(f"{tsp}.source.output = {tsp}.OFF")
        except: pass

    def _safe_abort(self):
        try: self.inst.write(TSP_SAFE_ABORT); time.sleep(0.15)
        except: pass

    # ── Core measurement ──────────────────────────────────────────────────────

    def _measure(self):
        """
        FET point-by-point measurement.

        WIRING (fixed):
            Gate   → SMU1  (Node 1, Master)   Rear FORCE HI / LO
            Drain  → SMU2  (Node 2, Slave)    Rear FORCE HI / LO
            Source → GNDU                      Rear GNDU post

        TRANSFER (Id vs Vg):
            Gate column  = sweep params   Drain column = bias level(s)
            Gate SMU1 steps Vg, Drain SMU2 holds Vds, Id read from SMU2

        OUTPUT (Id vs Vds):
            Drain column = sweep params   Gate column  = bias level(s)
            Drain SMU2 steps Vds, Gate SMU1 holds Vg, Id read from SMU2
        """
        sweep_tsp = bias_tsp = None
        try:
            self.inst = self.app.rm.open_resource(self.addr_var.get().strip())
            self.inst.timeout = 60000

            mode   = self.mode_var.get()
            is_xfr = "Transfer" in mode

            # ── Resolve TSP names ─────────────────────────────────────────────
            gate_tsp  = self._TSP.get(self.v["Gate"]["smu"].get())
            drain_tsp = self._TSP.get(self.v["Drain"]["smu"].get())
            if not gate_tsp:
                raise ValueError("Gate must be assigned to SMU1 or SMU2.")
            if not drain_tsp:
                raise ValueError("Drain must be assigned to SMU1 or SMU2.")

            if is_xfr:
                sweep_tsp, sweep_term = gate_tsp,  "Gate"
                bias_tsp,  bias_term  = drain_tsp, "Drain"
            else:
                sweep_tsp, sweep_term = drain_tsp, "Drain"
                bias_tsp,  bias_term  = gate_tsp,  "Gate"

            read_tsp = drain_tsp   # Id is always read from Drain (SMU2)

            # node number for waitcomplete — None means master only
            def _node(tsp):
                return 2 if "node[2]" in tsp else 1

            sweep_node = _node(sweep_tsp)
            bias_node  = _node(bias_tsp)
            read_node  = _node(read_tsp)

            # ── Parameters ────────────────────────────────────────────────────
            sw_start = self._fv(sweep_term, "start",   0.0)
            sw_stop  = self._fv(sweep_term, "stop",    8.0)
            sw_step  = self._fv(sweep_term, "step",    0.1)
            sw_comp  = self._fv(sweep_term, "comp",    0.01)
            sw_nplc  = self._fv(sweep_term, "nplc",    1.0)
            sw_delay = self._fv(sweep_term, "delay",   0.05)
            sw_dual  = self.v[sweep_term]["dual"].get()

            bi_op    = self.v[bias_term]["op"].get()
            bi_comp  = self._fv(bias_term, "comp",  0.5)
            bi_bias  = self._fv(bias_term, "bias",  10.0)
            bi_start = self._fv(bias_term, "start", 10.0)
            bi_stop  = self._fv(bias_term, "stop",   0.0)
            bi_step  = self._fv(bias_term, "step",  -1.0)

            if bi_op in self._BIAS:
                bias_levels = [bi_bias]
            elif bi_op in self._SWEEP or bi_op in self._LOG:
                bias_levels = self._step_list(bi_start, bi_stop, bi_step)
            elif bi_op in self._LIST:
                lv = self._list_vals.get(bias_term, [])
                bias_levels = lv if lv else [bi_bias]
            else:
                bias_levels = [bi_bias]

            sw_op   = self.v[sweep_term]["op"].get()
            sw_auto = self.v[sweep_term]["pts_auto"].get()
            if sw_op in self._LIST:
                lv = self._list_vals.get(sweep_term, [])
                if not lv:
                    raise ValueError(f"No list values for {sweep_term}.")
                sweep_base = list(lv)
            elif sw_auto:
                # Auto mode: derive points from start/stop/step
                if abs(sw_step) < 1e-12:
                    raise ValueError("Sweep step cannot be zero.")
                sweep_base = self._step_list(sw_start, sw_stop, sw_step)
            else:
                # Manual mode: use specified point count, linspace between start/stop
                try:
                    n_manual = max(2, int(self.v[sweep_term]["pts"].get()))
                except ValueError:
                    n_manual = 81
                sweep_base = list(np.linspace(sw_start, sw_stop, n_manual))

            sweep_vals = (list(sweep_base) + list(reversed(sweep_base))
                          if sw_dual else list(sweep_base))
            n_pts = len(sweep_vals)
            if n_pts == 0:
                raise ValueError("Sweep produced 0 points.")

            sense_c  = ("SENSE_4WIRE" if "4-Wire" in self.sense_var.get()
                        else "SENSE_2WIRE")
            term_str = self.terminals_var.get().upper()

            # ── Safe TSP write/query helpers ──────────────────────────────────
            def W(cmd):
                """Write and ignore response."""
                self.inst.write(cmd)

            def Q(cmd, default="nan"):
                """Query, return string; return default on error or nil."""
                try:
                    r = self.inst.query(cmd).strip()
                    return default if r in ("nil", "", "nil\n") else r
                except Exception:
                    return default

            def Qf(cmd, default=float("nan")):
                """Query and return float; return default on nil/error."""
                try:
                    r = self.inst.query(cmd).strip()
                    if r in ("nil", "", "nil\n"):
                        return default
                    return float(r)
                except Exception:
                    return default

            def wait(_=None):
                time.sleep(0.08)   # settle — waitcomplete() not needed for simple writes

            # ── TSP-Link / reset ──────────────────────────────────────────────
            needs_node2 = "node[2]" in (sweep_tsp + bias_tsp)
            self._set_status("Resetting…", "#F59E0B")
            W("*CLS")

            if needs_node2:
                W("tsplink.reset(1)")
                time.sleep(2.0)
                state = Q("print(tsplink.state)", "offline")
                if "online" not in state.lower():
                    raise RuntimeError(
                        f"TSP-Link not online (state={state!r}).\n\n"
                        "• Connect TSP-Link cable: rear TSP-LINK port on both 2450s\n"
                        "• Both units must be powered on\n"
                        "• Set node IDs: master=1, slave=2\n"
                        "  (front panel: MENU → System → TSP-Link Node)")
                # Reset each node cleanly — don't call reset() globally
                W("node[1].reset()"); time.sleep(0.5)
                W("node[2].reset()"); time.sleep(0.5)
                wait(1); wait(2)
            else:
                W("reset()"); time.sleep(0.6)

            # ── Configure SMUs (flat individual writes, no tsp.reset()) ───────
            self._set_status("Configuring…", "#F59E0B")

            def cfg_smu(tsp, comp, nplc, sense, term, node):
                W(f"{tsp}.terminals = {tsp}.TERMINALS_{term}")
                W(f"{tsp}.source.func = {tsp}.FUNC_DC_VOLTAGE")
                W(f"{tsp}.source.ilimit.level = {comp}")
                W(f"{tsp}.source.autorange = {tsp}.ON")
                W(f"{tsp}.source.level = 0")
                W(f"{tsp}.measure.func = {tsp}.FUNC_DC_CURRENT")
                W(f"{tsp}.measure.sense = {tsp}.{sense}")
                W(f"{tsp}.measure.nplc = {nplc}")
                W(f"{tsp}.measure.autorange = {tsp}.ON")
                W(f"{tsp}.measure.autozero.enable = {tsp}.OFF")
                W(f"{tsp}.measure.autozero.once()")
                time.sleep(0.2)
                W(f"{tsp}.source.output = {tsp}.ON")
                time.sleep(0.4)   # output settle

            cfg_smu(sweep_tsp, sw_comp, sw_nplc, sense_c, term_str, sweep_node)
            cfg_smu(bias_tsp,  bi_comp, sw_nplc, sense_c, term_str, bias_node)
            time.sleep(0.3)

            # ── Measurement loop ──────────────────────────────────────────────
            total_curves = len(bias_levels)
            total_pts    = total_curves * n_pts

            for ci, bv in enumerate(bias_levels):
                if self._stop: break

                bias_lbl = (f"Vds = {bv:.4g} V" if is_xfr else f"Vg = {bv:.4g} V")
                self._set_status(
                    f"Curve {ci+1}/{total_curves}  —  {bias_lbl}", "#F59E0B")

                W(f"{bias_tsp}.source.level = {bv:.6g}")
                time.sleep(0.05)
                time.sleep(sw_delay + 0.15)

                curve_dict = {
                    "label": bias_lbl + ("  [dual]" if sw_dual else ""),
                    "step": bv, "x": [], "y": []
                }
                self.curves.append(curve_dict)
                all_x, all_y = [], []
                interval = max(1, n_pts // 20)

                for si, sv in enumerate(sweep_vals):
                    if self._stop: break

                    # Set sweep voltage and wait for source to settle
                    W(f"{sweep_tsp}.source.level = {sv:.6g}")
                    time.sleep(max(sw_delay, 0.05))   # source settle

                    # Read drain current (always from Drain SMU2)
                    i_val = Qf(f"print({read_tsp}.measure.read())")

                    all_x.append(sv)
                    all_y.append(i_val)

                    if si % interval == 0 or si == n_pts - 1:
                        curve_dict["x"] = list(all_x)
                        curve_dict["y"] = list(all_y)
                        self.after(0, self._replot)

                    done = ci * n_pts + si + 1
                    pct  = min(int(100 * done / total_pts), 99)
                    self.after(0, lambda pv=pct: self.progress.update(pv))

                curve_dict["x"] = all_x
                curve_dict["y"] = all_y
                self.after(0, self._replot)

            # ── Ramp to 0 V then outputs off ──────────────────────────────────
            try:
                W(f"{sweep_tsp}.source.level = 0")
                W(f"{bias_tsp}.source.level = 0")
                time.sleep(0.05); time.sleep(0.05)
                time.sleep(0.15)
                W(f"{sweep_tsp}.source.output = {sweep_tsp}.OFF")
                W(f"{bias_tsp}.source.output = {bias_tsp}.OFF")
                time.sleep(0.2)
                self.inst.close()
            except Exception:
                pass

            if self._stop:
                self._set_status("Stopped", "#EF4444")
                self.after(0, lambda: self.progress.stop(False))
            else:
                n = len(self.curves)
                self._set_status(
                    f"✓  Done  —  {n} curve{'s' if n!=1 else ''}  acquired",
                    "#22C55E")
                self.after(0, lambda: self.progress.stop(True))
                self.after(0, self._auto_save)

        except Exception as e:
            err = str(e)
            self.after(0, lambda em=err: messagebox.showerror("FET Error", em))
            self._set_status("Error", "#EF4444")
            self.after(0, lambda: self.progress.stop(False))
            try:
                if sweep_tsp:
                    self.inst.write(f"{sweep_tsp}.source.output = {sweep_tsp}.OFF")
                if bias_tsp:
                    self.inst.write(f"{bias_tsp}.source.output = {bias_tsp}.OFF")
            except Exception:
                pass
        finally:
            self.is_running = False
            self._stop      = False


    # ── Export ────────────────────────────────────────────────────────────────

    def _make_fn(self):
        tag = "Transfer" if "Transfer" in self.mode_var.get() else "Output"
        return f"FET_{tag}_{time.strftime('%Y%m%d_%H%M%S')}"

    def _write_fet_csv(self, path):
        mode   = self.mode_var.get()
        is_xfr = "Transfer" in mode
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["# Keithley 2450  I-V Studio  v3.0  —  FET Characterisation"])
            w.writerow(["# Mode", mode])
            for t in self.TERMINALS:
                w.writerow([f"# {t}", self.v[t]["smu"].get(), self.v[t]["op"].get()])
            w.writerow(["# Date", time.strftime("%d/%m/%Y  %H:%M:%S")])
            w.writerow([])
            hdr = ["Vg (V)" if is_xfr else "Vds (V)"]
            for c in self.curves: hdr.append(f"Id  {c['label']} (A)")
            w.writerow(hdr)
            if not self.curves: return
            max_pts = max(len(c["x"]) for c in self.curves)
            for i in range(max_pts):
                row = [f"{self.curves[0]['x'][i]:.6g}"
                       if i < len(self.curves[0]["x"]) else ""]
                for c in self.curves:
                    row.append(f"{c['y'][i]:.9g}" if i < len(c["y"]) else "")
                w.writerow(row)

    def _export_csv(self):
        if not self.curves:
            messagebox.showwarning("No Data","Run a measurement first."); return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv", initialdir=self.app.save_dir,
            initialfile=f"{self._make_fn()}.csv",
            filetypes=[("CSV","*.csv"),("All","*.*")])
        if not path: return
        try:
            self._write_fet_csv(path)
            messagebox.showinfo("Saved", f"CSV saved:\n{path}")
        except Exception as e:
            messagebox.showerror("Export Error", str(e))

    def _save_png(self):
        if not self.curves:
            messagebox.showwarning("No Data","Run a measurement first."); return
        path = filedialog.asksaveasfilename(
            defaultextension=".png", initialdir=self.app.save_dir,
            initialfile=f"{self._make_fn()}.png",
            filetypes=[("PNG","*.png"),("All","*.*")])
        if path:
            try:
                self.fig.savefig(path, dpi=300, facecolor=PANEL,
                                 bbox_inches="tight")
                self._set_status("PNG saved")
            except Exception as e:
                messagebox.showerror("Save Error", str(e))

    def _auto_save(self):
        try:
            base = self._make_fn()
            self._write_fet_csv(
                os.path.join(self.app.save_dir, f"{base}.csv"))
            self.fig.savefig(
                os.path.join(self.app.save_dir, f"{base}.png"),
                dpi=300, facecolor=PANEL, bbox_inches="tight")
            self._set_status(f"Auto-saved  →  {self.app.save_dir}")
        except: pass


# ═══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    root = tk.Tk()
    app  = KeithleyApp(root)
    root.mainloop()