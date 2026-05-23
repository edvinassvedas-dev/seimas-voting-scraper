import io
import queue
import threading
import webbrowser
import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, messagebox

import pandas as pd

from sheets import (
    _accent_fold, normalize_tags, tag_match_keys, verify_schema,
    delete_voting_from_sheets, check_duplicate, append_to_sheets,
    load_db_summary, load_voting_detail, load_voting_meta,
    QuotaExceeded,
)

from api import fetch_voting

# ── GUI config ────────

TERMS = ["2024-2028", "2020-2024", "2016-2020"]
DEFAULT_TERM = "2024-2028"

# ─ Global font  ────────────────────────────────────────────────────────
UI_FONT_FAMILY = "Helvetica"     # body
UI_FONT_SIZE = 10               # base 
HEADING_DELTA = 1               # section headers = base + _this (bold)
TITLE_DELTA = 2                 # window/preview title = base + _this (bold)
CAPTION_DELTA = -2              # captions = base + _this (italic)

# symbols for the per-fraction symbol chart. (Can use multiple values...)
CHART_SYMBOLS = ["▓"]
CHART_BAR_WIDTH = 40  # max symbols across the widest bar

RESULT_COLORS = {
    "uz": "#1B7F2E",          # green
    "pries": "#C0392B",       # red
    "susilaike": "#7F7F7F",   # grey
}
# Distinct colours for any other result types (didn't vote, whatever else...).
FALLBACK_COLORS = ["#ffffff", "#B8860B"]


# ── sorting + cell,row copy ────────

def _coerce(value):
    try:
        return float(str(value).replace(",", ""))
    except (ValueError, TypeError):
        return str(value).lower()


def make_sortable(tree):
    """column heading: toggling asc/desc."""
    state = {"col": None, "desc": False}

    def sort_by(col):
        if state["col"] == col:
            state["desc"] = not state["desc"]
        else:
            state["col"] = col
            state["desc"] = False
        rows = [(tree.set(i, col), i) for i in tree.get_children("")]
        rows.sort(key=lambda t: _coerce(t[0]), reverse=state["desc"])
        for pos, (_, item) in enumerate(rows):
            tree.move(item, "", pos)
        for c in tree["columns"]:
            base = c.rstrip(" ▲▼")
            tree.heading(c, text=base)
        arrow = " ▼" if state["desc"] else " ▲"
        tree.heading(col, text=col + arrow)

    for c in tree["columns"]:
        tree.heading(c, text=c, command=lambda cc=c: sort_by(cc))


def make_copyable(tree):
    tree._cell = (None, None)

    def remember(event):
        if tree.identify("region", event.x, event.y) == "cell":
            tree._cell = (tree.identify_row(event.y),
                          tree.identify_column(event.x))

    def cell_text():
        item, col = tree._cell
        if not item or not col:
            return None
        idx = int(col.replace("#", "")) - 1
        vals = tree.item(item, "values")
        return str(vals[idx]) if 0 <= idx < len(vals) else None

    def to_clip(text):
        if text is None:
            return
        tree.clipboard_clear()
        tree.clipboard_append(text)
        tree.update_idletasks()

    def copy_cell(_=None):
        to_clip(cell_text())

    def copy_row(_=None):
        item = tree._cell[0] or (tree.selection()[0] if tree.selection() else None)
        if item:
            to_clip("\t".join(str(v) for v in tree.item(item, "values")))

    menu = tk.Menu(tree, tearoff=0)
    menu.add_command(label="Copy cell", command=copy_cell)
    menu.add_command(label="Copy row", command=copy_row)

    def popup(event):
        remember(event)
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    tree.bind("<Button-1>", remember, add="+")
    tree.bind("<Button-3>", popup, add="+")
    tree.bind("<Control-c>", copy_cell, add="+")
    tree.bind("<Control-C>", copy_cell, add="+")


def fill_tree(tree, rows):
    tree.delete(*tree.get_children())
    for r in rows:
        tree.insert("", "end", values=list(r))


def build_symbol_chart(d, result_order):
    """Build data for a coloured monospace symbol bar chart of record count by fraction (one symbol + one color per result value).
    """
    counts = (d.groupby(["fraction", "result"]).size()
                .unstack(fill_value=0))
    for r in result_order:
        if r not in counts.columns:
            counts[r] = 0
    counts = counts[result_order]
    counts = counts.loc[counts.sum(axis=1).sort_values(ascending=False).index]

    symbol = {r: CHART_SYMBOLS[i % len(CHART_SYMBOLS)]
              for i, r in enumerate(result_order)}

    color = {}
    fb = 0
    for r in result_order:
        known = RESULT_COLORS.get(_accent_fold(r))
        if known:
            color[r] = known
        else:
            color[r] = FALLBACK_COLORS[fb % len(FALLBACK_COLORS)]
            fb += 1

    legend = [(symbol[r], r) for r in result_order]

    max_total = int(counts.sum(axis=1).max()) if len(counts) else 0
    if max_total == 0:
        return {"symbol": symbol, "color": color, "legend": legend,
                "lines": [], "caption": ""}
    if max_total <= CHART_BAR_WIDTH:
        scale = 1.0
        caption = "1 symbol = 1 record"
    else:
        scale = CHART_BAR_WIDTH / max_total
        caption = f"bar scaled — 1 symbol \u2248 {1 / scale:.1f} records"

    name_w = min(max(len(str(f)) for f in counts.index), 24)
    lines = []
    for frac, row in counts.iterrows():
        segments = []
        for r in result_order:
            n = int(row[r])
            if n <= 0:
                continue
            seg = max(1, round(n * scale))
            segments.append((symbol[r] * seg, r))
        label = f"{str(frac)[:name_w]:<{name_w}}"
        lines.append((label, segments, int(row.sum())))
    return {"symbol": symbol, "color": color, "legend": legend,
            "lines": lines, "caption": caption}


# ── Preview window ─────────────

class PreviewWindow(tk.Toplevel):

    def __init__(self, master, voting_id, term, name):
        super().__init__(master)
        self.title(f"Preview — voting {voting_id}")
        self.geometry("640x790")
        self.transient(master)

        try:
            self.df = load_voting_detail(voting_id)
        except Exception as e:
            messagebox.showerror("Error", f"Could not load detail: {e}", parent=self)
            self.destroy()
            return

        try:
            meta = load_voting_meta(voting_id)
        except Exception:
            meta = {'name': name, 'url': '', 'term': term, 'tags': []}
        url = (meta.get('url') or '').strip()

        header = f"Voting {voting_id}"
        if name:
            header += f" — {name}"
        if term:
            header += f"  [{term}]"
        ttk.Label(self, text=header, font="AppTitle").pack(
            anchor="w", padx=10, pady=(10, 4))

        if url:
            link = ttk.Label(self, text=url, font="AppCaption",
                             foreground="#1565C0", cursor="hand2")
            link.pack(anchor="w", padx=10, pady=(0, 4))
            link.bind("<Button-1>", lambda e, u=url: webbrowser.open(u))

        if self.df.empty:
            ttk.Label(self, text="No records in the database for this voting.").pack(
                padx=10, pady=20)
            return

        # Filters
        fr = ttk.Frame(self)
        fr.pack(fill="x", padx=10, pady=4)
        ttk.Label(fr, text="Fraction:").pack(side="left")
        self.fraction_cb = ttk.Combobox(
            fr, state="readonly", width=20,
            values=["(All)"] + sorted(self.df['fraction'].dropna().unique().tolist()))
        self.fraction_cb.current(0)
        self.fraction_cb.pack(side="left", padx=(4, 16))
        ttk.Label(fr, text="Result:").pack(side="left")
        self.result_cb = ttk.Combobox(
            fr, state="readonly", width=16,
            values=["(All)"] + sorted(self.df['result'].dropna().unique().tolist()))
        self.result_cb.current(0)
        self.result_cb.pack(side="left", padx=4)
        self.fraction_cb.bind("<<ComboboxSelected>>", lambda e: self.refresh())
        self.result_cb.bind("<<ComboboxSelected>>", lambda e: self.refresh())

        # Detail table
        ttk.Label(self, text="Records", font="AppHeading").pack(
            anchor="w", padx=10, pady=(8, 2))
        cols = ["Fraction", "Member", "Result"]
        self.detail = ttk.Treeview(self, columns=cols, show="headings", height=8)
        for c in cols:
            self.detail.heading(c, text=c)
            self.detail.column(c, width=240 if c == "Member" else 150)
        self.detail.pack(fill="both", expand=True, padx=10)
        make_sortable(self.detail)
        make_copyable(self.detail)

        # Cross-tab viz
        ttk.Label(self, text="Count split — fraction x result",
                  font="AppHeading").pack(
            anchor="w", padx=10, pady=(10, 2))
        self.xtab_frame = ttk.Frame(self)
        self.xtab_frame.pack(fill="x", padx=10, pady=(0, 6))
        self.xtab = None

        ttk.Label(self, text="Count by fraction and result",
                  font="AppHeading").pack(
            anchor="w", padx=10, pady=(6, 2))
        self.chart_frame = ttk.Frame(self)
        self.chart_frame.pack(fill="x", padx=10, pady=(0, 10))
        self.chart = None

        self.refresh()

    def _filtered(self):
        d = self.df
        f = self.fraction_cb.get()
        r = self.result_cb.get()
        if f != "(All)":
            d = d[d['fraction'] == f]
        if r != "(All)":
            d = d[d['result'] == r]
        return d

    def _clear_chart(self):
        if self.chart is not None:
            self.chart.destroy()
            self.chart = None

    def refresh(self):
        d = self._filtered()
        # Display order: Fraction, Member, Result (data arrives member-first)
        disp = [c for c in ["fraction", "member", "result"] if c in d.columns]
        fill_tree(self.detail, d[disp].values.tolist())

        if self.xtab is not None:
            self.xtab.destroy()
        self._clear_chart()
        if d.empty:
            self.xtab = ttk.Label(self.xtab_frame, text="(no rows match the filter)")
            self.xtab.pack(anchor="w")
            self.chart = ttk.Label(self.chart_frame, text="(no rows match the filter)")
            self.chart.pack(anchor="w")
            return
        ct = pd.crosstab(d['fraction'], d['result'],
                         margins=True, margins_name="Total")
        result_cols = [c for c in ct.columns if c != "Total"]
        DESIRED_KEYS = ["uz", "pries", "susilaike", ""]  #
        by_key = {}
        for c in result_cols:
            by_key.setdefault(_accent_fold(c), c)
        result_order = [by_key[k] for k in DESIRED_KEYS if k in by_key]
        listed = set(result_order)
        result_order += [c for c in result_cols if c not in listed]
        ct = ct[result_order + ["Total"]]
        cols = ["Fraction"] + [str(c) for c in ct.columns]
        tv = ttk.Treeview(self.xtab_frame, columns=cols, show="headings",
                           height=min(len(ct) + 1, 12))
        for c in cols:
            tv.heading(c, text=c)
            tv.column(c, width=150 if c == "Fraction" else 90, anchor="center")
        for idx, row in ct.iterrows():
            tv.insert("", "end",
                      values=[str(idx)] + [str(int(v)) for v in row.tolist()])
        tv.pack(fill="x")
        make_copyable(tv)
        self.xtab = tv

        chart = build_symbol_chart(d, result_order)
        box = ttk.Frame(self.chart_frame)

        bg = self.cget("background")

        def tag_name(result):
            return "res_" + _accent_fold(result).replace(" ", "_")

        # Colour-coded legend
        legend = tk.Text(box, height=1, font="TkFixedFont", wrap="none",
                         borderwidth=0, background=bg)
        for sym, r in chart["legend"]:
            t = tag_name(r)
            legend.tag_configure(t, foreground=chart["color"][r])
            legend.insert("end", f"{sym} ", t)
            legend.insert("end", f"{r}    ")
        legend.configure(state="disabled")
        legend.pack(fill="x")

        body = tk.Text(box, height=min(len(chart["lines"]) + 1, 14),
                       font="TkFixedFont", wrap="none",
                       borderwidth=0, background=bg)
        for r in result_order:
            body.tag_configure(tag_name(r), foreground=chart["color"][r])
        for label, segments, total in chart["lines"]:
            body.insert("end", f"{label}  ")
            for seg_str, r in segments:
                body.insert("end", seg_str, tag_name(r))
            body.insert("end", f" ({total})\n")
        body.configure(state="disabled")
        body.pack(fill="x", pady=(2, 0))

        if chart["caption"]:
            ttk.Label(box, text=chart["caption"],
                      font="AppCaption").pack(anchor="w", pady=(2, 0))
        box.pack(fill="x")
        self.chart = box


# ── Main app ───

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Seimas Voting App")
        self.geometry("900x780")
        try:
            self.style = ttk.Style()
            self.style.theme_use("clam")
        except tk.TclError:
            self.style = ttk.Style()
        for name, color in (("Blue", "#1565C0"),
                            ("Green", "#1B7F2E"),
                            ("Red", "#C0392B")):
            self.style.configure(f"{name}.TButton", foreground=color)
            self.style.map(f"{name}.TButton",
                           foreground=[("active", color), ("pressed", color),
                                       ("disabled", "#9E9E9E")])

        self.fonts = {
            "body": tkfont.Font(name="AppBody", family=UI_FONT_FAMILY,
                                 size=UI_FONT_SIZE),
            "heading": tkfont.Font(name="AppHeading", family=UI_FONT_FAMILY,
                                    size=UI_FONT_SIZE, weight="bold"),
            "title": tkfont.Font(name="AppTitle", family=UI_FONT_FAMILY,
                                  size=UI_FONT_SIZE, weight="bold"),
            "caption": tkfont.Font(name="AppCaption", family=UI_FONT_FAMILY,
                                    size=UI_FONT_SIZE, slant="italic"),
        }

        self.result_df = pd.DataFrame()
        self.voting_id = ""
        self._summary_cache = []      
        self._all_tags = []           
        self._fetch_queue = queue.Queue()
        self._fetching = False

        self._build_inputs()
        self._build_results()
        self._build_summary()
        self.apply_font_size(UI_FONT_SIZE)
        self.refresh_summary()

    # -- Global font control --
    def apply_font_size(self, size):
        try:
            size = int(size)
        except (TypeError, ValueError):
            return

        self.fonts["body"].configure(family=UI_FONT_FAMILY, size=size)
        self.fonts["heading"].configure(family=UI_FONT_FAMILY,
                                        size=size + HEADING_DELTA, weight="bold")
        self.fonts["title"].configure(family=UI_FONT_FAMILY,
                                      size=size + TITLE_DELTA, weight="bold")
        self.fonts["caption"].configure(
            family=UI_FONT_FAMILY,
            size=max(7, size + CAPTION_DELTA), slant="italic")

        for nm in ("TkDefaultFont", "TkTextFont", "TkMenuFont",
                   "TkHeadingFont", "TkIconFont", "TkTooltipFont"):
            try:
                tkfont.nametofont(nm).configure(size=size)
            except tk.TclError:
                pass
        try:
            tkfont.nametofont("TkFixedFont").configure(size=size)
        except tk.TclError:
            pass

        rowh = self.fonts["body"].metrics("linespace") + 6
        self.style.configure("Treeview", font="AppBody", rowheight=rowh)
        self.style.configure("Treeview.Heading", font="AppBody")
        self.style.configure("TButton", font="AppBody")
        self.style.configure("TLabel", font="AppBody")
        self.style.configure("TCombobox", font="AppBody")
        self.style.configure("TEntry", font="AppBody")
        self.option_add("*TCombobox*Listbox.font", "AppBody")

    # -- Input section --
    def _build_inputs(self):
        f = ttk.Frame(self)
        f.pack(fill="x", padx=10, pady=10)

        ttk.Label(f, text="Voting ID:", width=14).grid(row=0, column=0, sticky="w", pady=2)
        self.voting_id_e = ttk.Entry(f, width=10)
        self.voting_id_e.grid(row=0, column=1, sticky="w", pady=2)

        ttk.Separator(f, orient="horizontal").grid(
            row=1, column=0, columnspan=2, sticky="ew", pady=6)

        ttk.Label(f, text="Term:", width=14).grid(row=2, column=0, sticky="w", pady=2)
        self.term_cb = ttk.Combobox(f, state="readonly", width=12, values=TERMS)
        self.term_cb.set(DEFAULT_TERM)
        self.term_cb.grid(row=2, column=1, sticky="w", pady=2)

        ttk.Label(f, text="Voting Name:", width=14).grid(row=3, column=0, sticky="w", pady=2)
        self.voting_name_e = ttk.Entry(f, width=72)
        self.voting_name_e.grid(row=3, column=1, sticky="w", pady=2)

        ttk.Label(f, text="Voting URL:", width=14).grid(row=4, column=0, sticky="w", pady=2)
        self.voting_url_e = ttk.Entry(f, width=72)
        self.voting_url_e.grid(row=4, column=1, sticky="w", pady=2)

        ttk.Label(f, text="Tags:", width=14).grid(row=5, column=0, sticky="w", pady=2)
        self.tags_e = ttk.Entry(f, width=72)
        self.tags_e.grid(row=5, column=1, sticky="w", pady=2)
        ttk.Label(f, text="comma-separated", font="AppCaption").grid(
            row=6, column=1, sticky="w")

        b = ttk.Frame(self)
        b.pack(fill="x", padx=10)
        self.get_btn = ttk.Button(b, text="Get Data", command=self.on_get_data,
                                   style="Blue.TButton")
        self.get_btn.pack(side="left")
        self.insert_btn = ttk.Button(b, text="Insert into DB",
                                     command=self.on_insert,
                                     style="Green.TButton")
        self.insert_btn.pack(side="left", padx=4)
        ttk.Button(b, text="Copy to Clipboard", command=self.on_copy_clip).pack(side="left")
        ttk.Button(b, text="Exit", command=self.destroy).pack(side="left", padx=4)
        self.status_lbl = ttk.Label(b, text="", font="AppCaption")
        self.status_lbl.pack(side="left", padx=12)

    # --- Results table --
    def _build_results(self):
        ttk.Separator(self).pack(fill="x", pady=8)
        ttk.Label(self, text="Results", font="AppHeading").pack(
            anchor="w", padx=10)
        cols = ["Voting", "Date", "Member", "Fraction", "Result"]
        widths = [70, 90, 240, 150, 110]
        self.results = ttk.Treeview(self, columns=cols, show="headings", height=10)
        for c, w in zip(cols, widths):
            self.results.heading(c, text=c)
            self.results.column(c, width=w)
        self.results.pack(fill="both", expand=True, padx=10, pady=4)
        make_sortable(self.results)
        make_copyable(self.results)

    # -- Summary table --
    def _build_summary(self):
        ttk.Separator(self).pack(fill="x", pady=8)
        bar = ttk.Frame(self)
        bar.pack(fill="x", padx=10)
        ttk.Label(bar, text="Database Summary",
                  font="AppHeading").pack(side="left")
        ttk.Button(bar, text="Refresh",
                   command=self.refresh_summary).pack(side="left", padx=8)
        self.preview_btn = ttk.Button(bar, text="Preview Selected",
                                      command=self.on_preview, state="disabled")
        self.preview_btn.pack(side="left")
        # Delete Selected
        self.delete_btn = ttk.Button(bar, text="Delete Selected",
                                     command=self.on_delete, state="disabled",
                                     style="Red.TButton")
        self.delete_btn.pack(side="right")

        filt = ttk.Frame(self)
        filt.pack(fill="x", padx=10, pady=(6, 0))
        ttk.Label(filt, text="Filter by tag:").pack(side="left")
        self.tag_filter_cb = ttk.Combobox(filt, width=28, values=["(All)"])
        self.tag_filter_cb.set("(All)")
        self.tag_filter_cb.pack(side="left", padx=(4, 4))
        self.tag_filter_cb.bind("<<ComboboxSelected>>",
                                lambda e: self._apply_tag_filter())
        self.tag_filter_cb.bind("<Return>",
                                lambda e: self._apply_tag_filter())
        ttk.Button(filt, text="Clear",
                   command=self._clear_tag_filter).pack(side="left")

        cols = ["Term", "Voting ID", "Date", "Voting Name", "Tags", "*"]
        widths = [70, 70, 70, 390, 200, 30]
        self.summary = ttk.Treeview(self, columns=cols, show="headings", height=12)
        for c, w in zip(cols, widths):
            self.summary.heading(c, text=c)
            self.summary.column(c, width=w)
        self.summary.pack(fill="both", expand=True, padx=10, pady=4)
        self.summary.bind("<<TreeviewSelect>>", self.on_summary_select)
        make_sortable(self.summary)
        make_copyable(self.summary)

    # -- Async fetch --
    def on_get_data(self):
        if self._fetching:
            return
        vid = self.voting_id_e.get().strip()
        if not vid:
            messagebox.showerror("Error", "Please enter a Voting ID.")
            return
        self.voting_id = vid
        self._fetching = True
        self.get_btn.config(state="disabled")
        self.insert_btn.config(state="disabled")
        self.config(cursor="watch")
        self.status_lbl.config(text=f"Fetching voting {vid}…")

        def worker():
            try:
                df, auto_name = fetch_voting(vid)
                self._fetch_queue.put(("ok", vid, df, auto_name))
            except Exception as e:
                self._fetch_queue.put(("err", vid, e, None))

        threading.Thread(target=worker, daemon=True).start()
        self.after(100, self._poll_fetch)

    def _poll_fetch(self):
        try:
            status, vid, payload, auto_name = self._fetch_queue.get_nowait()
        except queue.Empty:
            self.after(100, self._poll_fetch)
            return

        self._fetching = False
        self.get_btn.config(state="normal")
        self.insert_btn.config(state="normal")
        self.config(cursor="")
        self.status_lbl.config(text="")

        if status == "err":
            messagebox.showerror("Error", f"Error fetching data: {payload}")
            return

        df = payload
        if df.empty:
            messagebox.showinfo("No data",
                                f"No records found for voting ID {vid}.")
            fill_tree(self.results, [])
            self.result_df = pd.DataFrame()
            return
        self.result_df = df
        fill_tree(self.results, df.values.tolist())
        if auto_name and not self.voting_name_e.get().strip():
            self.voting_name_e.insert(0, auto_name)

    def on_insert(self):
        if self.result_df.empty:
            messagebox.showerror("Error", "Please fetch data before inserting.")
            return
        # If the duplicate check itself fails (e.g. rate limited), we must
        # NOT proceed — inserting without being able to verify could create
        # a real duplicate. Block with a clear message instead of guessing.
        try:
            is_dup = check_duplicate(self.voting_id)
        except QuotaExceeded as e:
            messagebox.showerror(
                "Rate limited",
                f"Could not verify whether voting {self.voting_id} already "
                f"exists, so the insert was not performed.\n\n{e}")
            return
        except Exception as e:
            messagebox.showerror(
                "Error",
                f"Could not check for duplicates; insert not performed.\n{e}")
            return
        if is_dup:
            if not messagebox.askyesno(
                    "Duplicate Warning",
                    f"Voting ID {self.voting_id} already exists.\nInsert anyway?"):
                return
        term = self.term_cb.get().strip()
        name = self.voting_name_e.get().strip()
        url = self.voting_url_e.get().strip()
        tags = ", ".join(normalize_tags(self.tags_e.get()))
        dim_row = [self.voting_id, term, name, url, tags]
        try:
            append_to_sheets(self.result_df, dim_row)
        except Exception as e:
            messagebox.showerror("Error", f"Insertion failed: {e}")
            return
        messagebox.showinfo("Done", "Inserted successfully.")
        self.refresh_summary()
        self.voting_id_e.delete(0, "end")
        self.voting_name_e.delete(0, "end")
        self.voting_url_e.delete(0, "end")
        self.tags_e.delete(0, "end")
        self.term_cb.set(DEFAULT_TERM)
        fill_tree(self.results, [])
        self.result_df = pd.DataFrame()

    def on_copy_clip(self):
        if self.result_df.empty:
            messagebox.showerror("Error", "No data to copy.")
            return
        buf = io.StringIO()
        self.result_df.to_csv(buf, sep="\t", index=False)
        self.clipboard_clear()
        self.clipboard_append(buf.getvalue())
        self.update_idletasks()
        messagebox.showinfo("Done", "Copied to clipboard.")

    def _selected_summary_row(self):
        sel = self.summary.selection()
        if not sel:
            return None
        return self.summary.item(sel[0], "values")

    def on_summary_select(self, _=None):
        has = bool(self.summary.selection())
        state = "normal" if has else "disabled"
        self.delete_btn.config(state=state)
        self.preview_btn.config(state=state)

    def on_delete(self):
        row = self._selected_summary_row()
        if not row:
            return
        vid, name = row[1], row[3]
        label = name or str(vid)
        if not messagebox.askyesno(
                "Confirm Delete",
                f"Delete all records for voting ID {vid} ({label})?\n"
                "This removes data from both sheets."):
            return
        try:
            delete_voting_from_sheets(vid)
        except Exception as e:
            messagebox.showerror("Error", f"Delete failed: {e}")
            return
        messagebox.showinfo("Done", f"Voting {vid} deleted.")
        self.refresh_summary()

    def on_preview(self):
        row = self._selected_summary_row()
        if not row:
            return
        # Row layout: [term(0), vid(1), date, name(3), tags, records]
        term, vid, name = row[0], row[1], row[3]
        PreviewWindow(self, vid, term, name)  # term used by PreviewWindow

    def _apply_tag_filter(self):
        sel = self.tag_filter_cb.get().strip()
        if not sel or sel == "(All)":
            rows = self._summary_cache
        else:
            sel_keys = tag_match_keys(sel)
            rows = [r for r in self._summary_cache
                    if sel_keys & tag_match_keys(r[4])]  # tags at idx 4
        fill_tree(self.summary, rows)
        self.delete_btn.config(state="disabled")
        self.preview_btn.config(state="disabled")

    def _clear_tag_filter(self):
        self.tag_filter_cb.set("(All)")
        self._apply_tag_filter()

    def refresh_summary(self):
        try:
            # Verify the spreadsheet schema once per session
            if not getattr(self, "_schema_checked", False):
                verify_schema()
                self._schema_checked = True
            self._summary_cache, self._all_tags = load_db_summary()
        except Exception as e:
            messagebox.showerror("Error", f"Error loading summary: {e}")
            self._summary_cache, self._all_tags = [], []
        self.tag_filter_cb.configure(values=["(All)"] + self._all_tags)
        cur = self.tag_filter_cb.get().strip()
        if cur and cur != "(All)" and cur in self._all_tags:
            self._apply_tag_filter()
        else:
            self.tag_filter_cb.set("(All)")
            fill_tree(self.summary, self._summary_cache)
        self.delete_btn.config(state="disabled")
        self.preview_btn.config(state="disabled")


if __name__ == "__main__":
    App().mainloop()
