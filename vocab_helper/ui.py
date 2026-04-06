from __future__ import annotations

from datetime import date, timedelta
import hashlib
import sqlite3
import tkinter as tk
import tkinter.font as tkfont
from tkinter import messagebox, simpledialog, ttk
from typing import Callable

from .db import VocabRepository
from .kana import suggest_hiragana
from .models import VocabEntry, Workbook
from .validators import ValidationError


BASE_FONT_SIZE = 12
JAPANESE_FONT_CANDIDATES = (
    "Yu Gothic",
    "Yu Gothic UI",
    "Meiryo",
    "MS Gothic",
    "Noto Sans CJK JP",
    "Hiragino Sans",
    "DejaVu Sans",
)
LATIN_FONT_CANDIDATES = (
    "Segoe UI",
    "Arial",
    "Helvetica",
    "DejaVu Sans",
)

TIER_BG_COLORS = {
    "gray": "#efefef",
    "green": "#e7f7ea",
    "yellow": "#fff5cf",
    "red": "#fde7e7",
}

LANGUAGE_NAMES = {
    "JP": "Japanese",
    "EN": "English",
}

PART_OF_SPEECH_OPTIONS = (
    "",
    "noun",
    "verb",
    "adjective",
    "adverb",
    "expression",
    "particle",
    "auxiliary",
    "other",
)

ACTIVITY_COLORS = (
    "#ebedf0",
    "#c6e48b",
    "#7bc96f",
    "#239a3b",
    "#196127",
)


def _pick_font_family(root: tk.Misc, candidates: tuple[str, ...]) -> str:
    available = {family.lower(): family for family in tkfont.families(root)}
    for candidate in candidates:
        exact = available.get(candidate.lower())
        if exact:
            return exact
    return tkfont.nametofont("TkDefaultFont").actual("family")


def _build_font_set(root: tk.Misc) -> dict[str, tkfont.Font]:
    latin_family = _pick_font_family(root, LATIN_FONT_CANDIDATES)
    japanese_family = _pick_font_family(root, JAPANESE_FONT_CANDIDATES)

    return {
        "latin": tkfont.Font(root=root, family=latin_family, size=BASE_FONT_SIZE),
        "japanese": tkfont.Font(root=root, family=japanese_family, size=BASE_FONT_SIZE),
        "tree_heading": tkfont.Font(root=root, family=latin_family, size=BASE_FONT_SIZE, weight="bold"),
    }


def _insert_markdown_inline(widget: tk.Text, text: str, tags: tuple[str, ...]) -> None:
    index = 0
    while index < len(text):
        if text.startswith("**", index):
            end_index = text.find("**", index + 2)
            if end_index != -1:
                widget.insert("end", text[index + 2 : end_index], tags + ("md_bold",))
                index = end_index + 2
                continue
        if text.startswith("*", index):
            end_index = text.find("*", index + 1)
            if end_index != -1:
                widget.insert("end", text[index + 1 : end_index], tags + ("md_italic",))
                index = end_index + 1
                continue
        if text.startswith("`", index):
            end_index = text.find("`", index + 1)
            if end_index != -1:
                widget.insert("end", text[index + 1 : end_index], tags + ("md_code_inline",))
                index = end_index + 1
                continue

        widget.insert("end", text[index], tags)
        index += 1


def _render_markdown_to_text(widget: tk.Text, markdown_text: str, base_font: tkfont.Font, monospace_font: tkfont.Font) -> None:
    heading1_font = base_font.copy()
    heading1_font.configure(size=base_font.cget("size") + 4, weight="bold")
    heading2_font = base_font.copy()
    heading2_font.configure(size=base_font.cget("size") + 2, weight="bold")
    heading3_font = base_font.copy()
    heading3_font.configure(weight="bold")

    bold_font = base_font.copy()
    bold_font.configure(weight="bold")
    italic_font = base_font.copy()
    italic_font.configure(slant="italic")

    widget.configure(state="normal")
    widget.delete("1.0", "end")
    widget.tag_configure("md_body", font=base_font)
    widget.tag_configure("md_h1", font=heading1_font, spacing1=8, spacing3=4)
    widget.tag_configure("md_h2", font=heading2_font, spacing1=6, spacing3=3)
    widget.tag_configure("md_h3", font=heading3_font, spacing1=4, spacing3=2)
    widget.tag_configure("md_bold", font=bold_font)
    widget.tag_configure("md_italic", font=italic_font)
    widget.tag_configure("md_code_inline", font=monospace_font, background="#f3f3f3")
    widget.tag_configure("md_code_block", font=monospace_font, background="#f7f7f7", lmargin1=8, lmargin2=8)

    in_code_block = False
    for raw_line in markdown_text.splitlines():
        line = raw_line.rstrip("\n")

        if line.strip().startswith("```"):
            in_code_block = not in_code_block
            continue

        if in_code_block:
            widget.insert("end", line + "\n", ("md_code_block",))
            continue

        if line.startswith("# "):
            _insert_markdown_inline(widget, line[2:], ("md_h1",))
            widget.insert("end", "\n", ("md_h1",))
            continue
        if line.startswith("## "):
            _insert_markdown_inline(widget, line[3:], ("md_h2",))
            widget.insert("end", "\n", ("md_h2",))
            continue
        if line.startswith("### "):
            _insert_markdown_inline(widget, line[4:], ("md_h3",))
            widget.insert("end", "\n", ("md_h3",))
            continue
        if line.startswith("- ") or line.startswith("* "):
            widget.insert("end", "- ", ("md_body",))
            _insert_markdown_inline(widget, line[2:], ("md_body",))
            widget.insert("end", "\n", ("md_body",))
            continue

        _insert_markdown_inline(widget, line, ("md_body",))
        widget.insert("end", "\n", ("md_body",))

    widget.configure(state="disabled")


def _open_detail_dialog_from_test(test_dialog: tk.Toplevel, entry_id: int) -> None:
    open_detail = getattr(test_dialog.master, "_open_detail_dialog", None)
    if not callable(open_detail):
        return

    test_dialog.grab_release()
    try:
        open_detail(entry_id)
    finally:
        if test_dialog.winfo_exists():
            test_dialog.grab_set()


class MainWindow(tk.Tk):
    def __init__(self, repository: VocabRepository) -> None:
        super().__init__()
        self.repository = repository
        self.fonts = _build_font_set(self)
        self._tree_entry_ids: dict[str, int] = {}
        self._entry_stats_by_id: dict[int, tuple[int, int, str]] = {}
        self.workbook_rows: list[Workbook] = self.repository.list_workbooks()
        self._workbook_display_to_id: dict[str, int] = {}
        self.current_workbook_id = self.repository.get_current_workbook_id()
        self.current_workbook = (
            self.repository.get_workbook(self.current_workbook_id)
            if self.current_workbook_id is not None
            else None
        )
        self.target_language_code = self.current_workbook.target_language_code if self.current_workbook else "JP"
        self.assistant_language_code = "EN"
        initial_settings_language = self.target_language_code if self.target_language_code in LANGUAGE_NAMES else "JP"
        self.settings_language_code_var = tk.StringVar(value=initial_settings_language)
        self._settings_workbook_display_to_id: dict[str, int] = {}
        self._settings_column_vars: dict[int, tk.BooleanVar] = {}
        self._table_column_keys: list[str] = ["target_text", "kana", "meaning"]

        self.show_tier_colors_var = tk.BooleanVar(value=True)
        self.sort_mode_var = tk.StringVar(value="time")
        self.time_order_var = tk.StringVar(value="newest")
        self.test_pick_strategy_var = tk.StringVar(value="strict")
        self.active_filter_tag_ids: list[int] = []
        self._workbook_filter_tag_ids: dict[int, list[int]] = {}
        self.tag_filter_summary_var = tk.StringVar(value="Tag filter: All")
        self.search_query_var = tk.StringVar(value="")
        self.workbook_selection_var = tk.StringVar(value="")

        self.title("Vocabulary Workbook Helper")
        self.geometry("900x580")
        self.minsize(760, 460)

        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        self._configure_styles()
        self._build_widgets()
        self.search_query_var.trace_add("write", self._on_search_query_changed)
        self._refresh_workbook_selector(select_workbook_id=self.current_workbook_id)
        self._refresh_test_button_labels()
        self._refresh_table_columns()
        self._refresh_settings_page()
        self._set_home_action_enabled(self.current_workbook_id is not None)
        self._bind_shortcuts()
        self.refresh_entries()

    def _configure_styles(self) -> None:
        style = ttk.Style(self)
        style.configure("App.TLabel", font=self.fonts["latin"])
        style.configure("Status.TLabel", font=self.fonts["latin"], foreground="#444444")
        style.configure("App.TButton", font=self.fonts["latin"])
        style.configure("App.TEntry", font=self.fonts["latin"])
        style.configure("Japanese.TEntry", font=self.fonts["japanese"])
        style.configure("Treeview", font=self.fonts["japanese"], rowheight=30)
        style.configure("Treeview.Heading", font=self.fonts["tree_heading"])
        style.configure("App.TNotebook", padding=2)
        style.configure(
            "App.TNotebook.Tab",
            font=(
                self.fonts["latin"].cget("family"),
                self.fonts["latin"].cget("size") + 1,
                "bold",
            ),
            padding=(18, 10),
        )

    def _language_display_name(self, code: str) -> str:
        return LANGUAGE_NAMES.get(code.upper(), code.upper())

    def _current_target_label(self) -> str:
        if self.current_workbook is None:
            return "Target text"
        return self.current_workbook.target_label

    def _current_meaning_label(self) -> str:
        if self.current_workbook is None:
            return "Meaning"
        return self.current_workbook.meaning_label

    def _workbook_supports_kana(self) -> bool:
        return self.current_workbook is not None and self.current_workbook.preset_key == "japanese"

    def _target_field_label(self) -> str:
        return self._current_target_label()

    def _assistant_field_label(self) -> str:
        return self._current_meaning_label()

    def _refresh_language_labels(self) -> None:
        if self.current_workbook_id is None:
            return

        self._refresh_table_columns()

    def _refresh_table_columns(self) -> None:
        if self.current_workbook_id is None:
            self.tree.configure(columns=("empty",), displaycolumns=("empty",))
            self.tree.heading("empty", text="No workbook selected")
            self.tree.column("empty", width=860, anchor="w")
            self._table_column_keys = []
            return

        try:
            property_rows = self.repository.get_workbook_visible_properties(self.current_workbook_id)
        except (LookupError, sqlite3.Error):
            property_rows = []

        visible_rows = [
            row
            for row in property_rows
            if (row[5] or row[1] == "target_text") and (row[1] != "kana" or self._workbook_supports_kana())
        ]
        if not visible_rows:
            visible_rows = [
                row
                for row in property_rows
                if row[1] in {"target_text", "meaning"}
            ]

        self._table_column_keys = [row[1] for row in visible_rows]
        column_ids = tuple(f"col_{index}" for index, _row in enumerate(visible_rows))
        self.tree.configure(columns=column_ids, displaycolumns=column_ids)

        if not visible_rows:
            return

        total_width = max(self.tree.winfo_width(), 840)
        width_per_column = max(int(total_width / len(visible_rows)) - 10, 180)

        for index, row in enumerate(visible_rows):
            column_id = column_ids[index]
            property_key = row[1]
            property_label = row[2]
            if property_key == "target_text":
                heading_text = self._target_field_label()
            elif property_key == "meaning":
                heading_text = self._assistant_field_label()
            else:
                heading_text = property_label
            self.tree.heading(column_id, text=heading_text)
            self.tree.column(column_id, width=width_per_column, anchor="w")

    def _build_widgets(self) -> None:
        container = ttk.Frame(self, padding=12)
        container.grid(row=0, column=0, sticky="nsew")
        container.columnconfigure(0, weight=1)
        container.rowconfigure(1, weight=1)

        workbook_row = ttk.Frame(container)
        workbook_row.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        workbook_row.columnconfigure(1, weight=1)

        ttk.Label(workbook_row, text="Workbook", style="App.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.workbook_combo = ttk.Combobox(
            workbook_row,
            state="readonly",
            textvariable=self.workbook_selection_var,
            width=36,
            font=self.fonts["latin"],
        )
        self.workbook_combo.grid(row=0, column=1, sticky="ew")
        self.workbook_combo.bind("<<ComboboxSelected>>", self._on_workbook_selected)
        ttk.Button(
            workbook_row,
            text="New workbook",
            command=self._open_workbook_creation_dialog,
            style="App.TButton",
        ).grid(row=0, column=2, sticky="w", padx=(8, 0))

        self.page_tabs = ttk.Notebook(container, style="App.TNotebook")
        self.page_tabs.grid(row=1, column=0, sticky="nsew")

        home_page = ttk.Frame(self.page_tabs)
        profile_page = ttk.Frame(self.page_tabs)
        settings_page = ttk.Frame(self.page_tabs)
        self.page_tabs.add(home_page, text="Home")
        self.page_tabs.add(profile_page, text="Profile")
        self.page_tabs.add(settings_page, text="Settings")

        home_page.columnconfigure(0, weight=1)
        home_page.rowconfigure(0, weight=1)

        table_frame = ttk.Frame(home_page)
        table_frame.grid(row=0, column=0, sticky="nsew")
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(
            table_frame,
            columns=("col_0", "col_1", "col_2"),
            show="headings",
            height=15,
            selectmode="extended",
        )
        self.tree.heading("col_0", text="Target text")
        self.tree.heading("col_1", text="Kana")
        self.tree.heading("col_2", text="Meaning")
        self.tree.column("col_0", width=300, anchor="w")
        self.tree.column("col_1", width=250, anchor="w")
        self.tree.column("col_2", width=300, anchor="w")

        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        self.context_menu = tk.Menu(self, tearoff=0, font=self.fonts["latin"])
        self.context_menu.add_command(label="View details", command=self._open_detail_dialog)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Edit", command=self._open_edit_dialog)
        self.context_menu.add_command(label="Delete selected", command=self._delete_selected_entry)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Increase priority", command=self._increase_selected_priority)
        self.context_menu.add_command(label="Decrease priority", command=self._decrease_selected_priority)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Assign tags to selected", command=self._assign_tags_to_selected)
        self.tree.bind("<Button-1>", self._on_tree_left_click, add="+")
        self.tree.bind("<Button-3>", self._show_context_menu)
        self.tree.bind("<Double-Button-1>", self._on_tree_double_click)

        self.tree.tag_configure("tier_gray", background=TIER_BG_COLORS["gray"])
        self.tree.tag_configure("tier_green", background=TIER_BG_COLORS["green"])
        self.tree.tag_configure("tier_yellow", background=TIER_BG_COLORS["yellow"])
        self.tree.tag_configure("tier_red", background=TIER_BG_COLORS["red"])

        button_row = ttk.Frame(home_page, padding=(0, 10, 0, 0))
        button_row.grid(row=1, column=0, sticky="ew")
        button_row.columnconfigure(0, weight=1)

        self.meaning_to_target_test_button = ttk.Button(
            button_row,
            text="Test Meaning -> Target",
            command=self._open_en_to_jp_test_dialog,
            style="App.TButton",
        )
        self.meaning_to_target_test_button.grid(row=0, column=1, padx=(0, 8), sticky="e")

        self.target_to_kana_test_button = ttk.Button(
            button_row,
            text="Test Target -> Kana",
            command=self._open_jp_to_kana_test_dialog,
            style="App.TButton",
        )
        self.target_to_kana_test_button.grid(row=0, column=2, padx=(0, 8), sticky="e")

        self.target_to_meaning_test_button = ttk.Button(
            button_row,
            text="Test Target -> Meaning",
            command=self._open_jp_to_en_test_dialog,
            style="App.TButton",
        )
        self.target_to_meaning_test_button.grid(row=0, column=3, padx=(0, 8), sticky="e")

        self.bulk_add_button = ttk.Button(
            button_row,
            text="Bulk add",
            command=self._open_bulk_add_dialog,
            style="App.TButton",
        )
        self.bulk_add_button.grid(row=0, column=4, padx=(0, 8), sticky="e")

        self.add_button = ttk.Button(button_row, text="+", width=4, command=self._open_add_dialog, style="App.TButton")
        self.add_button.grid(row=0, column=5, sticky="e")

        settings_row = ttk.Frame(home_page, padding=(0, 8, 0, 0))
        settings_row.grid(row=2, column=0, sticky="ew")
        settings_row.columnconfigure(9, weight=1)

        ttk.Checkbutton(
            settings_row,
            text="Tier colors",
            variable=self.show_tier_colors_var,
            command=self.refresh_entries,
        ).grid(row=0, column=0, sticky="w", padx=(0, 10))

        ttk.Label(settings_row, text="Sort", style="App.TLabel").grid(row=0, column=1, sticky="w")
        self.sort_mode_combo = ttk.Combobox(
            settings_row,
            values=("time", "stats", "tags"),
            state="readonly",
            width=8,
            textvariable=self.sort_mode_var,
        )
        self.sort_mode_combo.grid(row=0, column=2, sticky="w", padx=(4, 12))
        self.sort_mode_combo.bind("<<ComboboxSelected>>", self._on_sort_mode_changed)

        ttk.Label(settings_row, text="Time order", style="App.TLabel").grid(row=0, column=3, sticky="w")
        self.time_order_combo = ttk.Combobox(
            settings_row,
            values=("newest", "oldest"),
            state="readonly",
            width=8,
            textvariable=self.time_order_var,
        )
        self.time_order_combo.grid(row=0, column=4, sticky="w", padx=(4, 12))
        self.time_order_combo.bind("<<ComboboxSelected>>", self._on_sort_settings_changed)

        ttk.Label(settings_row, text="Test pick", style="App.TLabel").grid(row=0, column=5, sticky="w")
        self.test_pick_combo = ttk.Combobox(
            settings_row,
            values=("strict", "weighted"),
            state="readonly",
            width=10,
            textvariable=self.test_pick_strategy_var,
        )
        self.test_pick_combo.grid(row=0, column=6, sticky="w", padx=(4, 0))
        self.test_pick_combo.bind("<<ComboboxSelected>>", self._on_pick_strategy_changed)

        self.home_new_workbook_button = ttk.Button(
            settings_row,
            text="New workbook",
            command=self._open_workbook_creation_dialog,
            style="App.TButton",
        )
        self.home_new_workbook_button.grid(row=0, column=7, sticky="w", padx=(10, 0))

        self.home_tags_button = ttk.Button(
            settings_row,
            text="Tags",
            command=self._open_tag_manager_dialog,
            style="App.TButton",
        )
        self.home_tags_button.grid(row=0, column=8, sticky="w", padx=(8, 0))

        filter_row = ttk.Frame(home_page, padding=(0, 6, 0, 0))
        filter_row.grid(row=3, column=0, sticky="ew")
        filter_row.columnconfigure(4, weight=1)

        self.filter_tags_button = ttk.Button(
            filter_row,
            text="Filter tags",
            command=self._open_tag_filter_dialog,
            style="App.TButton",
        )
        self.filter_tags_button.grid(row=0, column=0, sticky="w")
        self.clear_filter_button = ttk.Button(
            filter_row,
            text="Clear filter",
            command=self._clear_tag_filter,
            style="App.TButton",
        )
        self.clear_filter_button.grid(row=0, column=1, sticky="w", padx=(8, 0))
        ttk.Label(filter_row, textvariable=self.tag_filter_summary_var, style="Status.TLabel").grid(
            row=0,
            column=2,
            sticky="w",
            padx=(12, 0),
        )
        ttk.Label(filter_row, text="Search", style="App.TLabel").grid(row=0, column=3, sticky="e", padx=(12, 6))
        self.search_entry = ttk.Entry(filter_row, textvariable=self.search_query_var, style="App.TEntry")
        self.search_entry.grid(row=0, column=4, sticky="ew")
        self.clear_search_button = ttk.Button(
            filter_row,
            text="Clear search",
            command=self._clear_search_query,
            style="App.TButton",
        )
        self.clear_search_button.grid(row=0, column=5, sticky="e", padx=(8, 0))

        self._update_sort_controls()
        self._refresh_language_labels()
        self._refresh_tag_filter_summary()

        status_row = ttk.Frame(home_page, padding=(0, 8, 0, 0))
        status_row.grid(row=4, column=0, sticky="ew")
        status_row.columnconfigure(0, weight=1)

        self.count_label = ttk.Label(status_row, text="Total vocabularies: 0", style="Status.TLabel")
        self.count_label.grid(row=0, column=0, sticky="w")

        profile_page.columnconfigure(0, weight=1)
        profile_page.rowconfigure(1, weight=1)

        activity_frame = ttk.LabelFrame(profile_page, text="Daily practice activity (last 180 days)", padding=8)
        activity_frame.grid(row=0, column=0, sticky="nw")
        activity_frame.columnconfigure(0, weight=1)

        self.activity_summary_label = ttk.Label(activity_frame, text="", style="Status.TLabel")
        self.activity_summary_label.grid(row=0, column=0, sticky="w", pady=(0, 6))

        self.activity_canvas = tk.Canvas(
            activity_frame,
            height=140,
            highlightthickness=0,
            background="#ffffff",
        )
        self.activity_canvas.grid(row=1, column=0, sticky="w")

        legend_frame = ttk.Frame(activity_frame)
        legend_frame.grid(row=2, column=0, sticky="w", pady=(6, 0))
        ttk.Label(legend_frame, text="Less", style="Status.TLabel").grid(row=0, column=0, padx=(0, 6))
        for index, color in enumerate(ACTIVITY_COLORS):
            swatch = tk.Canvas(legend_frame, width=10, height=10, highlightthickness=1, highlightbackground="#cccccc")
            swatch.create_rectangle(0, 0, 10, 10, fill=color, outline=color)
            swatch.grid(row=0, column=index + 1, padx=(0, 4))
        ttk.Label(legend_frame, text="More", style="Status.TLabel").grid(row=0, column=len(ACTIVITY_COLORS) + 1, padx=(2, 0))

        self._build_settings_widgets(settings_page)

    def _bind_shortcuts(self) -> None:
        self.bind("<Control-n>", self._handle_add_shortcut)
        self.bind("<Control-N>", self._handle_add_shortcut)
        self.bind("<Control-Shift-n>", self._handle_bulk_add_shortcut)
        self.bind("<Control-Shift-N>", self._handle_bulk_add_shortcut)
        self.bind("<Control-t>", self._handle_test_shortcut)
        self.bind("<Control-T>", self._handle_test_shortcut)

        # Keep Enter/Delete scoped to the list so text inputs in dialogs are unaffected.
        self.tree.bind("<Return>", self._handle_edit_shortcut)
        self.tree.bind("<KP_Enter>", self._handle_edit_shortcut)
        self.tree.bind("<Delete>", self._handle_delete_shortcut)

    def _build_settings_widgets(self, settings_page: ttk.Frame) -> None:
        settings_page.columnconfigure(0, weight=1)
        settings_page.rowconfigure(0, weight=1)

        root = ttk.Frame(settings_page, padding=10)
        root.grid(row=0, column=0, sticky="nsew")
        root.columnconfigure(0, weight=1)
        root.columnconfigure(1, weight=1)
        root.rowconfigure(0, weight=1)
        root.rowconfigure(1, weight=1)

        workbook_section = ttk.LabelFrame(root, text="Workbooks", padding=10)
        workbook_section.grid(row=0, column=0, sticky="nsew", padx=(0, 6), pady=(0, 6))
        workbook_section.columnconfigure(0, weight=1)
        workbook_section.rowconfigure(0, weight=1)

        self.settings_workbook_listbox = tk.Listbox(workbook_section, exportselection=False, font=self.fonts["latin"])
        self.settings_workbook_listbox.grid(row=0, column=0, sticky="nsew")
        workbook_scroll = ttk.Scrollbar(workbook_section, orient="vertical", command=self.settings_workbook_listbox.yview)
        workbook_scroll.grid(row=0, column=1, sticky="ns")
        self.settings_workbook_listbox.configure(yscrollcommand=workbook_scroll.set)
        self.settings_workbook_listbox.bind("<<ListboxSelect>>", self._on_settings_workbook_selected)

        workbook_actions = ttk.Frame(workbook_section)
        workbook_actions.grid(row=1, column=0, columnspan=2, sticky="e", pady=(8, 0))
        ttk.Button(workbook_actions, text="Create", command=self._open_workbook_creation_dialog, style="App.TButton").grid(
            row=0,
            column=0,
            padx=(0, 8),
        )
        ttk.Button(workbook_actions, text="Switch", command=self._switch_selected_settings_workbook, style="App.TButton").grid(
            row=0,
            column=1,
            padx=(0, 8),
        )
        ttk.Button(workbook_actions, text="Delete", command=self._delete_selected_settings_workbook, style="App.TButton").grid(
            row=0,
            column=2,
        )

        visibility_section = ttk.LabelFrame(root, text="Workbook columns", padding=10)
        visibility_section.grid(row=0, column=1, sticky="nsew", padx=(6, 0), pady=(0, 6))
        visibility_section.columnconfigure(0, weight=1)
        visibility_section.rowconfigure(1, weight=1)

        self.settings_visibility_info_var = tk.StringVar(value="Select a workbook to configure visible columns.")
        ttk.Label(
            visibility_section,
            textvariable=self.settings_visibility_info_var,
            style="Status.TLabel",
            wraplength=340,
            justify="left",
        ).grid(row=0, column=0, sticky="w", pady=(0, 8))

        self.settings_column_container = ttk.Frame(visibility_section)
        self.settings_column_container.grid(row=1, column=0, sticky="nsew")
        self.settings_column_container.columnconfigure(0, weight=1)

        ttk.Button(
            visibility_section,
            text="Save column visibility",
            command=self._save_selected_workbook_column_visibility,
            style="App.TButton",
        ).grid(row=2, column=0, sticky="e", pady=(8, 0))
        ttk.Button(
            visibility_section,
            text="Edit labels",
            command=self._edit_selected_workbook_labels,
            style="App.TButton",
        ).grid(row=3, column=0, sticky="e", pady=(8, 0))

        language_section = ttk.LabelFrame(root, text="Language schema", padding=10)
        language_section.grid(row=1, column=0, columnspan=2, sticky="nsew")
        language_section.columnconfigure(0, weight=1)
        language_section.rowconfigure(2, weight=1)

        language_row = ttk.Frame(language_section)
        language_row.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ttk.Label(language_row, text="Language", style="App.TLabel").grid(row=0, column=0, sticky="w")
        self.settings_language_combo = ttk.Combobox(
            language_row,
            state="readonly",
            values=tuple(LANGUAGE_NAMES.keys()),
            textvariable=self.settings_language_code_var,
            width=8,
            font=self.fonts["latin"],
        )
        self.settings_language_combo.grid(row=0, column=1, sticky="w", padx=(8, 0))
        self.settings_language_combo.bind("<<ComboboxSelected>>", self._on_settings_language_changed)

        property_actions = ttk.Frame(language_section)
        property_actions.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        ttk.Button(property_actions, text="Add property", command=self._add_language_property_from_settings, style="App.TButton").grid(
            row=0,
            column=0,
            padx=(0, 8),
        )
        ttk.Button(
            property_actions,
            text="Delete property",
            command=self._delete_language_property_from_settings,
            style="App.TButton",
        ).grid(row=0, column=1, padx=(0, 8))
        ttk.Button(
            property_actions,
            text="Manage tags",
            command=self._open_settings_language_tag_manager,
            style="App.TButton",
        ).grid(row=0, column=2)

        property_frame = ttk.Frame(language_section)
        property_frame.grid(row=2, column=0, sticky="nsew")
        property_frame.columnconfigure(0, weight=1)
        property_frame.rowconfigure(0, weight=1)

        self.settings_property_listbox = tk.Listbox(property_frame, exportselection=False, font=self.fonts["latin"])
        self.settings_property_listbox.grid(row=0, column=0, sticky="nsew")
        property_scroll = ttk.Scrollbar(property_frame, orient="vertical", command=self.settings_property_listbox.yview)
        property_scroll.grid(row=0, column=1, sticky="ns")
        self.settings_property_listbox.configure(yscrollcommand=property_scroll.set)

    def _set_home_action_enabled(self, enabled: bool) -> None:
        state = "normal" if enabled else "disabled"
        self.meaning_to_target_test_button.configure(state=state)
        self.target_to_meaning_test_button.configure(state=state)
        self.bulk_add_button.configure(state=state)
        self.add_button.configure(state=state)
        self.home_tags_button.configure(state=state)
        self.filter_tags_button.configure(state=state)
        self.clear_filter_button.configure(state=state)
        self.search_entry.configure(state=state)
        self.clear_search_button.configure(state=state)
        self.sort_mode_combo.configure(state="readonly" if enabled else "disabled")
        self.time_order_combo.configure(state="readonly" if enabled else "disabled")
        self.test_pick_combo.configure(state="readonly" if enabled else "disabled")

    def _refresh_settings_page(self) -> None:
        self._refresh_settings_workbook_list()
        self._refresh_settings_language_properties()
        self._refresh_settings_column_visibility()

    def _refresh_settings_workbook_list(self) -> None:
        self.workbook_rows = self.repository.list_workbooks()
        self._settings_workbook_display_to_id.clear()
        self.settings_workbook_listbox.delete(0, "end")

        selected_index = 0
        for index, workbook in enumerate(self.workbook_rows):
            display = f"{workbook.name} ({workbook.target_label})"
            self.settings_workbook_listbox.insert("end", display)
            self._settings_workbook_display_to_id[display] = workbook.id
            if workbook.id == self.current_workbook_id:
                selected_index = index

        if self.workbook_rows:
            self.settings_workbook_listbox.selection_clear(0, "end")
            self.settings_workbook_listbox.selection_set(selected_index)

    def _selected_settings_workbook_id(self) -> int | None:
        selection = self.settings_workbook_listbox.curselection()
        if not selection:
            return None
        index = selection[0]
        if index >= len(self.workbook_rows):
            return None
        return self.workbook_rows[index].id

    def _refresh_settings_column_visibility(self) -> None:
        for child in self.settings_column_container.winfo_children():
            child.destroy()
        self._settings_column_vars.clear()

        selected_workbook_id = self._selected_settings_workbook_id()
        if selected_workbook_id is None:
            ttk.Label(
                self.settings_column_container,
                text="No workbook available.",
                style="Status.TLabel",
            ).grid(row=0, column=0, sticky="w")
            self.settings_visibility_info_var.set("Create a workbook to configure columns.")
            return

        try:
            visibility_rows = self.repository.get_workbook_visible_properties(selected_workbook_id)
            workbook = self.repository.get_workbook(selected_workbook_id)
        except (LookupError, sqlite3.Error) as exc:
            self.settings_visibility_info_var.set(f"Could not load visibility settings: {exc}")
            return

        self.settings_visibility_info_var.set(
            f"Visible columns for workbook '{workbook.name}'. Target text is always visible."
        )

        for row_index, row in enumerate(visibility_rows):
            property_id, property_key, property_label, _is_predefined, _is_required, is_visible, _display_order = row
            var = tk.BooleanVar(value=bool(is_visible) or property_key == "target_text")
            self._settings_column_vars[property_id] = var
            checkbox = ttk.Checkbutton(
                self.settings_column_container,
                text=property_label,
                variable=var,
            )
            checkbox.grid(row=row_index, column=0, sticky="w", pady=(0, 4))
            if property_key == "target_text":
                checkbox.configure(state="disabled")

    def _save_selected_workbook_column_visibility(self) -> None:
        selected_workbook_id = self._selected_settings_workbook_id()
        if selected_workbook_id is None:
            messagebox.showwarning("No workbook", "Select a workbook first.", parent=self)
            return

        selected_property_ids = [
            property_id
            for property_id, var in self._settings_column_vars.items()
            if var.get()
        ]
        try:
            self.repository.set_workbook_visible_properties(selected_workbook_id, selected_property_ids)
        except (ValidationError, LookupError) as exc:
            messagebox.showerror("Visibility error", str(exc), parent=self)
            return
        except sqlite3.Error as exc:
            messagebox.showerror("Database error", f"Could not save visibility settings: {exc}", parent=self)
            return

        if self.current_workbook_id == selected_workbook_id:
            self._refresh_table_columns()
            self.refresh_entries()

    def _edit_selected_workbook_labels(self) -> None:
        selected_workbook_id = self._selected_settings_workbook_id()
        if selected_workbook_id is None:
            messagebox.showwarning("No workbook", "Select a workbook first.", parent=self)
            return

        try:
            workbook = self.repository.get_workbook(selected_workbook_id)
        except (LookupError, sqlite3.Error) as exc:
            messagebox.showerror("Workbook error", f"Could not load workbook labels: {exc}", parent=self)
            return

        dialog = WorkbookLabelsDialog(
            self,
            workbook=workbook,
            text_font=self.fonts["latin"],
        )
        self.wait_window(dialog)
        if dialog.result is None:
            return

        try:
            updated_workbook = self.repository.update_workbook_labels(
                selected_workbook_id,
                dialog.result[0],
                dialog.result[1],
            )
        except (ValidationError, LookupError) as exc:
            messagebox.showerror("Validation error", str(exc), parent=self)
            return
        except sqlite3.Error as exc:
            messagebox.showerror("Database error", f"Could not save workbook labels: {exc}", parent=self)
            return

        if self.current_workbook_id == selected_workbook_id:
            self.current_workbook = updated_workbook
            self.target_language_code = updated_workbook.target_language_code
            self._refresh_language_labels()
            self._refresh_test_button_labels()
            self._refresh_table_columns()
            self.refresh_entries()

        self._refresh_workbook_selector(select_workbook_id=self.current_workbook_id)
        self._refresh_settings_page()

    def _refresh_settings_language_properties(self) -> None:
        language_code = self.settings_language_code_var.get().strip().upper() or "JP"
        self.settings_property_listbox.delete(0, "end")
        try:
            property_rows = self.repository.list_language_properties(language_code)
        except (ValidationError, sqlite3.Error):
            return

        for _property_id, property_key, property_label, is_predefined, is_required in property_rows:
            flags: list[str] = []
            if is_predefined:
                flags.append("predefined")
            if is_required:
                flags.append("required")
            suffix = f" ({', '.join(flags)})" if flags else ""
            self.settings_property_listbox.insert("end", f"{property_label} [{property_key}]{suffix}")

    def _property_id_at_settings_selection(self) -> int | None:
        selection = self.settings_property_listbox.curselection()
        if not selection:
            return None

        language_code = self.settings_language_code_var.get().strip().upper() or "JP"
        property_rows = self.repository.list_language_properties(language_code)
        index = selection[0]
        if index >= len(property_rows):
            return None
        return property_rows[index][0]

    def _add_language_property_from_settings(self) -> None:
        language_code = self.settings_language_code_var.get().strip().upper() or "JP"
        key = simpledialog.askstring("Add property", "Property key (snake_case):", parent=self)
        if key is None:
            return
        label = simpledialog.askstring("Add property", "Property label:", parent=self)
        if label is None:
            return

        try:
            self.repository.add_language_property(language_code, key, label)
        except ValidationError as exc:
            messagebox.showerror("Validation error", str(exc), parent=self)
            return
        except sqlite3.Error as exc:
            messagebox.showerror("Database error", f"Could not add property: {exc}", parent=self)
            return

        self._refresh_settings_language_properties()
        self._refresh_settings_column_visibility()
        if self.current_workbook is not None and self.current_workbook.target_language_code == language_code:
            self._refresh_table_columns()
            self.refresh_entries()

    def _delete_language_property_from_settings(self) -> None:
        property_id = self._property_id_at_settings_selection()
        if property_id is None:
            messagebox.showwarning("No selection", "Select a property to delete.", parent=self)
            return

        if not messagebox.askyesno(
            "Delete property",
            "Delete this property from the selected language? Existing values will be removed.",
            parent=self,
        ):
            return

        try:
            self.repository.delete_language_property(property_id)
        except (LookupError, ValueError) as exc:
            messagebox.showerror("Not allowed", str(exc), parent=self)
            return
        except sqlite3.Error as exc:
            messagebox.showerror("Database error", f"Could not delete property: {exc}", parent=self)
            return

        self._refresh_settings_language_properties()
        self._refresh_settings_column_visibility()
        self._refresh_table_columns()
        self.refresh_entries()

    def _open_settings_language_tag_manager(self) -> None:
        language_code = self.settings_language_code_var.get().strip().upper() or "JP"
        dialog = TagManagerDialog(
            self,
            repository=self.repository,
            target_language_code=language_code,
            text_font=self.fonts["latin"],
        )
        self.wait_window(dialog)

        if dialog.changed and language_code == self.target_language_code:
            self._refresh_tag_filter_summary()
            self.refresh_entries()

    def _switch_selected_settings_workbook(self) -> None:
        workbook_id = self._selected_settings_workbook_id()
        if workbook_id is None:
            messagebox.showwarning("No selection", "Select a workbook to switch.", parent=self)
            return
        self._switch_workbook(workbook_id)

    def _delete_selected_settings_workbook(self) -> None:
        workbook_id = self._selected_settings_workbook_id()
        if workbook_id is None:
            messagebox.showwarning("No selection", "Select a workbook to delete.", parent=self)
            return

        workbook = next((row for row in self.workbook_rows if row.id == workbook_id), None)
        if workbook is None:
            return

        if not messagebox.askyesno(
            "Delete workbook",
            f"Delete workbook '{workbook.name}' and all its vocabulary data? This cannot be undone.",
            parent=self,
        ):
            return

        try:
            new_current_workbook_id = self.repository.delete_workbook(workbook_id)
        except LookupError as exc:
            messagebox.showerror("Not found", str(exc), parent=self)
            return
        except sqlite3.Error as exc:
            messagebox.showerror("Database error", f"Could not delete workbook: {exc}", parent=self)
            return

        self.current_workbook_id = new_current_workbook_id
        if new_current_workbook_id is None:
            self.current_workbook = None
            self.target_language_code = "JP"
            self.active_filter_tag_ids = []
            self._set_home_action_enabled(False)
            self._refresh_workbook_selector(select_workbook_id=None)
            self._refresh_language_labels()
            self._refresh_test_button_labels()
            self._refresh_table_columns()
            self.refresh_entries()
        else:
            self.current_workbook = self.repository.get_workbook(new_current_workbook_id)
            self.target_language_code = self.current_workbook.target_language_code
            self._set_home_action_enabled(True)
            self._refresh_workbook_selector(select_workbook_id=new_current_workbook_id)
            self._refresh_language_labels()
            self._refresh_test_button_labels()
            self._refresh_table_columns()
            self.refresh_entries()

        self._refresh_settings_page()

    def _on_settings_workbook_selected(self, _event: tk.Event) -> None:
        self._refresh_settings_column_visibility()

    def _on_settings_language_changed(self, _event: tk.Event) -> None:
        self._refresh_settings_language_properties()

    def _refresh_workbook_selector(self, select_workbook_id: int | None = None) -> None:
        self.workbook_rows = self.repository.list_workbooks()
        self._workbook_display_to_id = {}

        display_values: list[str] = []
        for workbook in self.workbook_rows:
            display = f"{workbook.name} ({workbook.target_label})"
            display_values.append(display)
            self._workbook_display_to_id[display] = workbook.id

        self.workbook_combo.configure(values=tuple(display_values))

        if not self.workbook_rows:
            self.workbook_selection_var.set("")
            self.workbook_combo.configure(state="disabled")
            return

        self.workbook_combo.configure(state="readonly")

        desired_id = select_workbook_id if select_workbook_id is not None else self.current_workbook_id
        selected_workbook = next((workbook for workbook in self.workbook_rows if workbook.id == desired_id), self.workbook_rows[0])
        selected_display = f"{selected_workbook.name} ({selected_workbook.target_label})"
        self.workbook_selection_var.set(selected_display)

    def _refresh_test_button_labels(self) -> None:
        if self.current_workbook_id is None:
            self.meaning_to_target_test_button.configure(text="Test Meaning -> Target")
            self.target_to_meaning_test_button.configure(text="Test Target -> Meaning")
            self.target_to_kana_test_button.grid_remove()
            return

        target_label = self._target_field_label()
        meaning_label = self._assistant_field_label()
        self.meaning_to_target_test_button.configure(text=f"Test {meaning_label} -> {target_label}")
        self.target_to_meaning_test_button.configure(text=f"Test {target_label} -> {meaning_label}")

        if self._workbook_supports_kana():
            self.target_to_kana_test_button.grid()
            self.target_to_kana_test_button.configure(text=f"Test {target_label} -> Kana", state="normal")
        else:
            self.target_to_kana_test_button.grid_remove()

    def _switch_workbook(self, workbook_id: int) -> None:
        if self.current_workbook_id:
            self._workbook_filter_tag_ids[self.current_workbook_id] = list(self.active_filter_tag_ids)

        try:
            workbook = self.repository.set_current_workbook_id(workbook_id)
        except (LookupError, ValidationError) as exc:
            messagebox.showerror("Workbook error", str(exc), parent=self)
            self._refresh_workbook_selector(select_workbook_id=self.current_workbook_id)
            return
        except sqlite3.Error as exc:
            messagebox.showerror("Database error", f"Could not switch workbook: {exc}", parent=self)
            self._refresh_workbook_selector(select_workbook_id=self.current_workbook_id)
            return

        self.current_workbook_id = workbook.id
        self.current_workbook = workbook
        self.target_language_code = workbook.target_language_code
        self.active_filter_tag_ids = list(self._workbook_filter_tag_ids.get(workbook.id, []))

        self._refresh_workbook_selector(select_workbook_id=workbook.id)
        self._refresh_language_labels()
        self._refresh_test_button_labels()
        self._refresh_table_columns()
        self._refresh_tag_filter_summary()
        self._set_home_action_enabled(True)
        self._refresh_settings_page()
        self.refresh_entries()

    def _on_workbook_selected(self, _event: tk.Event) -> None:
        selected_display = self.workbook_selection_var.get().strip()
        if not selected_display:
            return

        selected_workbook_id = self._workbook_display_to_id.get(selected_display)
        if selected_workbook_id is None or selected_workbook_id == self.current_workbook_id:
            return

        self._switch_workbook(selected_workbook_id)

    def _open_workbook_creation_dialog(self) -> None:
        dialog = WorkbookCreationDialog(self, text_font=self.fonts["latin"])
        self.wait_window(dialog)
        if dialog.result is None:
            return

        workbook_name, target_schema_code, preset_key, target_label, meaning_label = dialog.result
        try:
            workbook = self.repository.create_workbook(
                workbook_name,
                target_schema_code,
                preset_key=preset_key,
                target_label=target_label,
                meaning_label=meaning_label,
            )
        except ValidationError as exc:
            messagebox.showerror("Validation error", str(exc), parent=self)
            return
        except sqlite3.Error as exc:
            messagebox.showerror("Database error", f"Could not create workbook: {exc}", parent=self)
            return

        self._workbook_filter_tag_ids.setdefault(workbook.id, [])
        self._set_home_action_enabled(True)
        self._switch_workbook(workbook.id)

    def _on_sort_mode_changed(self, _event: tk.Event) -> None:
        self._update_sort_controls()
        self.refresh_entries()

    def _on_sort_settings_changed(self, _event: tk.Event) -> None:
        self.refresh_entries()

    def _on_pick_strategy_changed(self, _event: tk.Event) -> None:
        # Strategy is consumed when launching/starting tests.
        return

    def _update_sort_controls(self) -> None:
        if self.sort_mode_var.get() == "time":
            self.time_order_combo.configure(state="readonly")
        else:
            self.time_order_combo.configure(state="disabled")

    def _handle_add_shortcut(self, _event: tk.Event) -> str:
        self._open_add_dialog()
        return "break"

    def _handle_bulk_add_shortcut(self, _event: tk.Event) -> str:
        self._open_bulk_add_dialog()
        return "break"

    def _handle_test_shortcut(self, _event: tk.Event) -> str:
        self._open_en_to_jp_test_dialog()
        return "break"

    def _handle_edit_shortcut(self, _event: tk.Event) -> str:
        self._open_edit_dialog()
        return "break"

    def _handle_delete_shortcut(self, _event: tk.Event) -> str:
        self._delete_selected_entry()
        return "break"

    def _on_search_query_changed(self, *_args: object) -> None:
        if self.current_workbook_id is None:
            return
        self.refresh_entries()

    def _clear_search_query(self) -> None:
        if not self.search_query_var.get().strip():
            return
        self.search_query_var.set("")

    def refresh_entries(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)

        self._tree_entry_ids.clear()
        self._entry_stats_by_id.clear()

        if self.current_workbook_id is None:
            self.count_label.configure(text="No workbook. Create one in Settings.")
            self.activity_summary_label.configure(text="No workbook selected.")
            self.activity_canvas.delete("all")
            return

        entries_with_stats = self.repository.list_entries_with_stats(
            sort_mode=self.sort_mode_var.get(),
            time_order=self.time_order_var.get(),
            filter_tag_ids=self.active_filter_tag_ids,
            filter_match_mode="all",
            search_query=self.search_query_var.get(),
            target_language_code=self.target_language_code,
            workbook_id=self.current_workbook_id,
        )

        for entry, test_count, error_count, tier in entries_with_stats:
            tags: tuple[str, ...] = ()
            if self.show_tier_colors_var.get():
                tags = (f"tier_{tier}",)

            value_by_key: dict[str, str] = {
                "target_text": entry.japanese_text,
                "meaning": entry.english_text,
                "kana": entry.kana_text or "",
            }
            try:
                dynamic_values = self.repository.get_entry_property_values(entry.id)
            except (LookupError, sqlite3.Error):
                dynamic_values = {}
            value_by_key.update(dynamic_values)

            ordered_values = tuple(value_by_key.get(property_key, "") for property_key in self._table_column_keys)
            if not ordered_values:
                ordered_values = (entry.japanese_text,)

            item_id = self.tree.insert(
                "",
                "end",
                values=ordered_values,
                tags=tags,
            )
            self._tree_entry_ids[item_id] = entry.id
            self._entry_stats_by_id[entry.id] = (test_count, error_count, tier)

        has_search_query = bool(self.search_query_var.get().strip())
        if self.active_filter_tag_ids or has_search_query:
            total_count = self.repository.count_entries(workbook_id=self.current_workbook_id)
            self.count_label.configure(text=f"Visible vocabularies: {len(entries_with_stats)} / {total_count}")
        else:
            self.count_label.configure(text=f"Total vocabularies: {len(entries_with_stats)}")
        self._refresh_activity_grid()

    def _refresh_tag_filter_summary(self) -> None:
        if self.current_workbook_id is None:
            self.tag_filter_summary_var.set("Tag filter: unavailable")
            return

        try:
            available_tags = self.repository.list_tags(target_language_code=self.target_language_code)
        except sqlite3.Error:
            self.tag_filter_summary_var.set("Tag filter: unavailable")
            return

        label_by_id = {
            tag_id: f"{type_name}:{tag_name}"
            for tag_id, _type_id, type_name, tag_name, _type_predefined, _tag_predefined in available_tags
        }
        self.active_filter_tag_ids = [tag_id for tag_id in self.active_filter_tag_ids if tag_id in label_by_id]
        if not self.active_filter_tag_ids:
            self.tag_filter_summary_var.set("Tag filter: All")
            return

        selected_labels = [label_by_id[tag_id] for tag_id in self.active_filter_tag_ids]
        preview = ", ".join(selected_labels[:3])
        if len(selected_labels) > 3:
            preview += f", +{len(selected_labels) - 3} more"
        self.tag_filter_summary_var.set(f"Tag filter (ALL): {preview}")

    def _open_tag_manager_dialog(self) -> None:
        if self.current_workbook_id is None:
            messagebox.showwarning("No workbook", "Create or select a workbook first.", parent=self)
            return

        dialog = TagManagerDialog(
            self,
            repository=self.repository,
            target_language_code=self.target_language_code,
            text_font=self.fonts["latin"],
        )
        self.wait_window(dialog)

        if dialog.changed:
            self._refresh_tag_filter_summary()
            self.refresh_entries()

    def _open_tag_filter_dialog(self) -> None:
        if self.current_workbook_id is None:
            messagebox.showwarning("No workbook", "Create or select a workbook first.", parent=self)
            return

        dialog = TagSelectionDialog(
            self,
            repository=self.repository,
            target_language_code=self.target_language_code,
            selected_tag_ids=self.active_filter_tag_ids,
            include_part_of_speech=True,
            title="Filter by tags",
            text_font=self.fonts["latin"],
        )
        self.wait_window(dialog)
        if dialog.result is None:
            return

        self.active_filter_tag_ids = list(dialog.result)
        self._refresh_tag_filter_summary()
        self.refresh_entries()

    def _clear_tag_filter(self) -> None:
        if not self.active_filter_tag_ids:
            return

        self.active_filter_tag_ids = []
        self._refresh_tag_filter_summary()
        self.refresh_entries()

    @staticmethod
    def _activity_level(count: int) -> int:
        if count <= 0:
            return 0
        if count <= 10:
            return 1
        if count <= 20:
            return 2
        if count <= 30:
            return 3
        return 4

    def _refresh_activity_grid(self) -> None:
        counts_by_date = self.repository.get_daily_unique_practice_counts(
            days_back=180,
            workbook_id=self.current_workbook_id,
        )
        active_days = sum(1 for count in counts_by_date.values() if count > 0)
        unique_total = sum(counts_by_date.values())
        self.activity_summary_label.configure(
            text=f"Unique vocab practiced: {unique_total} across {active_days} active days"
        )

        today = date.today()
        start_date = today - timedelta(days=179)
        grid_start = start_date - timedelta(days=start_date.weekday())
        weeks = ((today - grid_start).days // 7) + 1

        cell_size = 14
        gap = 4
        left_label_width = 34
        top_padding = 4
        right_padding = 8
        bottom_label_height = 20
        grid_width = weeks * cell_size + (weeks - 1) * gap
        grid_height = 7 * cell_size + 6 * gap
        grid_origin_x = left_label_width
        grid_origin_y = top_padding

        width = grid_origin_x + grid_width + right_padding
        height = grid_origin_y + grid_height + bottom_label_height

        self.activity_canvas.configure(width=width, height=height)
        self.activity_canvas.delete("all")

        weekday_font = (
            self.fonts["latin"].cget("family"),
            max(self.fonts["latin"].cget("size") - 2, 9),
        )
        weekday_labels = ("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun")
        for weekday_index, label in enumerate(weekday_labels):
            y_center = grid_origin_y + weekday_index * (cell_size + gap) + (cell_size / 2)
            self.activity_canvas.create_text(
                grid_origin_x - 6,
                y_center,
                text=label,
                anchor="e",
                fill="#666666",
                font=weekday_font,
            )

        current = start_date
        while current <= today:
            week_index = (current - grid_start).days // 7
            weekday_index = current.weekday()
            count = counts_by_date.get(current.isoformat(), 0)
            color = ACTIVITY_COLORS[self._activity_level(count)]

            x1 = grid_origin_x + week_index * (cell_size + gap)
            y1 = grid_origin_y + weekday_index * (cell_size + gap)
            x2 = x1 + cell_size
            y2 = y1 + cell_size

            self.activity_canvas.create_rectangle(
                x1,
                y1,
                x2,
                y2,
                fill=color,
                outline="#d0d0d0",
            )
            current += timedelta(days=1)

        month_markers: dict[tuple[int, int], int] = {}
        current = start_date
        while current <= today:
            if current == start_date or current.day == 1:
                month_key = (current.year, current.month)
                week_index = (current - grid_start).days // 7
                if month_key not in month_markers:
                    month_markers[month_key] = week_index
            current += timedelta(days=1)

        month_y = grid_origin_y + grid_height + 10
        month_font = weekday_font
        last_label_x = -999
        minimum_gap = 24
        for year, month in sorted(month_markers):
            week_index = month_markers[(year, month)]
            label_x = grid_origin_x + week_index * (cell_size + gap)
            if label_x - last_label_x < minimum_gap:
                continue
            label_text = date(year, month, 1).strftime("%b")
            self.activity_canvas.create_text(
                label_x,
                month_y,
                text=label_text,
                anchor="w",
                fill="#666666",
                font=month_font,
            )
            last_label_x = label_x

    def _show_context_menu(self, event: tk.Event) -> None:
        item_id = self.tree.identify_row(event.y)
        if not item_id:
            return

        selected_items = set(self.tree.selection())
        if item_id not in selected_items:
            self.tree.selection_set(item_id)
        self.tree.focus(item_id)

        try:
            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()

    def _on_tree_left_click(self, event: tk.Event) -> None:
        item_id = self.tree.identify_row(event.y)
        if item_id:
            return

        region = self.tree.identify_region(event.x, event.y)
        if region in {"heading", "separator"}:
            return

        if self.tree.selection():
            self.tree.selection_remove(self.tree.selection())
        self.tree.focus("")

    def _on_tree_double_click(self, event: tk.Event) -> str:
        item_id = self.tree.identify_row(event.y)
        if not item_id:
            return "break"

        entry_id = self._tree_entry_ids.get(item_id)
        if entry_id is not None:
            self._open_detail_dialog(entry_id)
        return "break"

    def _selected_entry_id(self) -> int | None:
        focused_item = self.tree.focus()
        if focused_item:
            focused_entry_id = self._tree_entry_ids.get(focused_item)
            if focused_entry_id is not None:
                return focused_entry_id

        selected_ids = self._selected_entry_ids()
        if not selected_ids:
            return None
        return selected_ids[0]

    def _selected_entry_ids(self) -> list[int]:
        selected_ids: list[int] = []
        for item_id in self.tree.selection():
            entry_id = self._tree_entry_ids.get(item_id)
            if entry_id is not None:
                selected_ids.append(entry_id)
        return list(dict.fromkeys(selected_ids))

    def _selected_japanese_text(self) -> str:
        selection = self.tree.selection()
        if not selection:
            return "selected entry"

        values = self.tree.item(selection[0], "values")
        if not values:
            return "selected entry"

        return str(values[0])

    def _open_add_dialog(self) -> None:
        if self.current_workbook_id is None:
            messagebox.showwarning("No workbook", "Create or select a workbook first.", parent=self)
            return

        def save_with_tags(
            japanese: str,
            kana: str,
            english: str,
            part_of_speech: str,
            tag_ids: list[int],
        ) -> None:
            entry = self.repository.add_entry(japanese, kana, english, part_of_speech)
            self.repository.set_entry_tags(
                entry.id,
                tag_ids,
                target_language_code=self.target_language_code,
                include_part_of_speech=False,
            )

        dialog = EntryDialog(
            self,
            repository=self.repository,
            title="Add vocabulary",
            save_button_text="Save",
            save_handler=save_with_tags,
            on_saved=self.refresh_entries,
            initial_japanese="",
            initial_kana="",
            initial_english="",
            initial_part_of_speech="",
            initial_tag_ids=[],
            target_label=self._target_field_label(),
            assistant_label=self._assistant_field_label(),
            target_language_code=self.target_language_code,
            enable_kana_suggest=self._workbook_supports_kana(),
        )
        self.wait_window(dialog)

    def _open_bulk_add_dialog(self) -> None:
        if self.current_workbook_id is None:
            messagebox.showwarning("No workbook", "Create or select a workbook first.", parent=self)
            return

        dialog = BulkAddDialog(
            self,
            repository=self.repository,
            on_saved=self.refresh_entries,
            text_font=self.fonts["latin"],
            target_label=self._target_field_label(),
            assistant_label=self._assistant_field_label(),
            show_kana=self._workbook_supports_kana(),
        )
        self.wait_window(dialog)

    def _open_en_to_jp_test_dialog(self) -> None:
        if self.current_workbook_id is None:
            messagebox.showwarning("No workbook", "Create or select a workbook first.", parent=self)
            return

        dialog = EnglishToJapaneseTestDialog(
            self,
            repository=self.repository,
            text_font=self.fonts["latin"],
            pick_strategy=self.test_pick_strategy_var.get(),
        )
        self.wait_window(dialog)
        self.refresh_entries()

    def _open_jp_to_kana_test_dialog(self) -> None:
        if self.current_workbook_id is None:
            messagebox.showwarning("No workbook", "Create or select a workbook first.", parent=self)
            return
        if not self._workbook_supports_kana():
            messagebox.showinfo(
                "Kana test unavailable",
                "This workbook does not have kana-enabled preset features.",
                parent=self,
            )
            return

        dialog = JapaneseToKanaTestDialog(
            self,
            repository=self.repository,
            text_font=self.fonts["japanese"],
            pick_strategy=self.test_pick_strategy_var.get(),
        )
        self.wait_window(dialog)
        self.refresh_entries()

    def _open_jp_to_en_test_dialog(self) -> None:
        if self.current_workbook_id is None:
            messagebox.showwarning("No workbook", "Create or select a workbook first.", parent=self)
            return

        dialog = JapaneseToEnglishChoiceTestDialog(
            self,
            repository=self.repository,
            text_font=self.fonts["japanese"],
            pick_strategy=self.test_pick_strategy_var.get(),
        )
        self.wait_window(dialog)
        self.refresh_entries()

    def _open_edit_dialog(self) -> None:
        if self.current_workbook_id is None:
            messagebox.showwarning("No workbook", "Create or select a workbook first.", parent=self)
            return

        entry_id = self._selected_entry_id()
        if entry_id is None:
            messagebox.showwarning("No selection", "Select an entry to edit.", parent=self)
            return

        try:
            entry = self.repository.get_entry(entry_id)
        except LookupError as exc:
            messagebox.showerror("Not found", str(exc), parent=self)
            self.refresh_entries()
            return

        entry_tags = self.repository.get_entry_tags(
            entry_id,
            target_language_code=self.target_language_code,
            include_part_of_speech=False,
        )
        initial_tag_ids = [tag_id for tag_id, _type_id, _type_name, _tag_name in entry_tags]

        def save_with_tags(
            japanese: str,
            kana: str,
            english: str,
            part_of_speech: str,
            tag_ids: list[int],
        ) -> None:
            self.repository.update_entry(
                entry_id,
                japanese,
                kana,
                english,
                part_of_speech,
            )
            self.repository.set_entry_tags(
                entry_id,
                tag_ids,
                target_language_code=self.target_language_code,
                include_part_of_speech=False,
            )

        dialog = EntryDialog(
            self,
            repository=self.repository,
            title="Edit vocabulary",
            save_button_text="Save changes",
            save_handler=save_with_tags,
            on_saved=self.refresh_entries,
            initial_japanese=entry.japanese_text,
            initial_kana=entry.kana_text or "",
            initial_english=entry.english_text,
            initial_part_of_speech=entry.part_of_speech or "",
            initial_tag_ids=initial_tag_ids,
            target_label=self._target_field_label(),
            assistant_label=self._assistant_field_label(),
            target_language_code=self.target_language_code,
            enable_kana_suggest=self._workbook_supports_kana(),
        )
        self.wait_window(dialog)

    def _open_detail_dialog(self, entry_id: int | None = None) -> None:
        if self.current_workbook_id is None:
            messagebox.showwarning("No workbook", "Create or select a workbook first.", parent=self)
            return

        resolved_entry_id = entry_id if entry_id is not None else self._selected_entry_id()
        if resolved_entry_id is None:
            messagebox.showwarning("No selection", "Select an entry to view details.", parent=self)
            return

        dialog = VocabularyDetailDialog(
            self,
            repository=self.repository,
            entry_id=resolved_entry_id,
            text_fonts=self.fonts,
            target_label=self._target_field_label(),
            assistant_label=self._assistant_field_label(),
            target_language_code=self.target_language_code,
            enable_kana_suggest=self._workbook_supports_kana(),
            on_saved=self.refresh_entries,
        )
        self.wait_window(dialog)

    def _open_language_settings_dialog(self) -> None:
        dialog = LanguageSettingsDialog(
            self,
            target_language=self.target_language_code,
            assistant_language=self.assistant_language_code,
            text_font=self.fonts["latin"],
        )
        self.wait_window(dialog)

        if dialog.result is None:
            return

        try:
            target, assistant = self.repository.set_language_settings(dialog.result[0], dialog.result[1])
        except ValidationError as exc:
            messagebox.showerror("Validation error", str(exc), parent=self)
            return
        except sqlite3.Error as exc:
            messagebox.showerror("Database error", f"Could not save language settings: {exc}", parent=self)
            return

        self.target_language_code = target
        self.assistant_language_code = assistant
        self.active_filter_tag_ids = []
        self._refresh_language_labels()
        self._refresh_tag_filter_summary()
        self.refresh_entries()

    def _delete_selected_entry(self) -> None:
        if self.current_workbook_id is None:
            messagebox.showwarning("No workbook", "Create or select a workbook first.", parent=self)
            return

        entry_ids = self._selected_entry_ids()
        if not entry_ids:
            messagebox.showwarning("No selection", "Select at least one entry to delete.", parent=self)
            return

        if len(entry_ids) == 1:
            selected_text = self._selected_japanese_text()
            prompt = f"Delete '{selected_text}'? This action cannot be undone."
        else:
            prompt = f"Delete {len(entry_ids)} selected entries? This action cannot be undone."

        confirmed = messagebox.askyesno("Delete entries", prompt, parent=self)
        if not confirmed:
            return

        try:
            self.repository.delete_entries(entry_ids)
        except LookupError as exc:
            messagebox.showerror("Not found", str(exc), parent=self)
        except sqlite3.Error as exc:
            messagebox.showerror("Database error", f"Could not delete entries: {exc}", parent=self)

        self.refresh_entries()

    def _increase_selected_priority(self) -> None:
        self._adjust_selected_priority(increase=True)

    def _decrease_selected_priority(self) -> None:
        self._adjust_selected_priority(increase=False)

    @staticmethod
    def _select_part_of_speech_value(pos_values: list[str]) -> str:
        ordered_values = [value for value in PART_OF_SPEECH_OPTIONS if value in pos_values]
        if ordered_values:
            return ordered_values[0]
        return sorted(pos_values)[0]

    def _normalize_selected_tag_ids(
        self,
        tag_ids: list[int],
        tag_meta_by_id: dict[int, tuple[str, str]],
    ) -> tuple[list[int], str]:
        valid_tag_ids = sorted({tag_id for tag_id in tag_ids if tag_id in tag_meta_by_id})
        part_of_speech_tag_ids = [
            tag_id for tag_id in valid_tag_ids if tag_meta_by_id[tag_id][0].lower() == "part_of_speech"
        ]
        if not part_of_speech_tag_ids:
            return valid_tag_ids, ""

        part_of_speech_values = [tag_meta_by_id[tag_id][1] for tag_id in part_of_speech_tag_ids]
        selected_part_of_speech = self._select_part_of_speech_value(part_of_speech_values)

        chosen_part_of_speech_tag_id = next(
            (
                tag_id
                for tag_id in part_of_speech_tag_ids
                if tag_meta_by_id[tag_id][1] == selected_part_of_speech
            ),
            part_of_speech_tag_ids[0],
        )

        part_of_speech_tag_id_set = set(part_of_speech_tag_ids)
        normalized_tag_ids = [tag_id for tag_id in valid_tag_ids if tag_id not in part_of_speech_tag_id_set]
        normalized_tag_ids.append(chosen_part_of_speech_tag_id)
        return sorted(set(normalized_tag_ids)), selected_part_of_speech

    def _assign_tags_to_selected(self) -> None:
        if self.current_workbook_id is None:
            messagebox.showwarning("No workbook", "Create or select a workbook first.", parent=self)
            return

        entry_ids = self._selected_entry_ids()
        if not entry_ids:
            messagebox.showwarning("No selection", "Select at least one entry.", parent=self)
            return

        mode_dialog = BulkTagModeDialog(self, text_font=self.fonts["latin"])
        self.wait_window(mode_dialog)
        if mode_dialog.result is None:
            return

        replace_mode = mode_dialog.result == "replace"

        dialog = TagSelectionDialog(
            self,
            repository=self.repository,
            target_language_code=self.target_language_code,
            selected_tag_ids=[],
            include_part_of_speech=True,
            title="Assign tags to selected entries",
            text_font=self.fonts["latin"],
        )
        self.wait_window(dialog)
        if dialog.result is None:
            return

        selected_tag_ids = list(dialog.result)
        if replace_mode and not selected_tag_ids:
            confirm_clear = messagebox.askyesno(
                "Clear tags",
                f"No tags selected. Replace mode will clear all tags on {len(entry_ids)} entries. Continue?",
                parent=self,
            )
            if not confirm_clear:
                return

        try:
            available_tags = self.repository.list_tags(
                target_language_code=self.target_language_code,
                include_part_of_speech=True,
            )
        except sqlite3.Error as exc:
            messagebox.showerror("Database error", f"Could not load tags: {exc}", parent=self)
            return

        tag_meta_by_id = {
            tag_id: (type_name, tag_name)
            for tag_id, _type_id, type_name, tag_name, _type_predefined, _tag_predefined in available_tags
        }

        updated_count = 0
        failures: list[tuple[int, str]] = []
        for entry_id in entry_ids:
            try:
                if replace_mode:
                    candidate_tag_ids = list(selected_tag_ids)
                else:
                    existing_tag_rows = self.repository.get_entry_tags(
                        entry_id,
                        target_language_code=self.target_language_code,
                        include_part_of_speech=True,
                    )
                    existing_tag_ids = [tag_id for tag_id, _type_id, _type_name, _tag_name in existing_tag_rows]
                    candidate_tag_ids = sorted(set(existing_tag_ids).union(selected_tag_ids))

                normalized_tag_ids, selected_part_of_speech = self._normalize_selected_tag_ids(
                    candidate_tag_ids,
                    tag_meta_by_id,
                )

                entry = self.repository.get_entry(entry_id)
                if (entry.part_of_speech or "") != selected_part_of_speech:
                    self.repository.update_entry(
                        entry_id,
                        entry.japanese_text,
                        entry.kana_text or "",
                        entry.english_text,
                        selected_part_of_speech,
                    )

                self.repository.set_entry_tags(
                    entry_id,
                    normalized_tag_ids,
                    target_language_code=self.target_language_code,
                    include_part_of_speech=True,
                )
                updated_count += 1
            except (ValidationError, LookupError, sqlite3.Error) as exc:
                failures.append((entry_id, str(exc)))

        self.refresh_entries()

        if failures:
            preview = "\n".join(f"Entry {entry_id}: {error_text}" for entry_id, error_text in failures[:5])
            if len(failures) > 5:
                preview += f"\n... and {len(failures) - 5} more"
            messagebox.showwarning(
                "Bulk tag assignment",
                f"Updated {updated_count} of {len(entry_ids)} entries.\n\nFailures:\n{preview}",
                parent=self,
            )
            return

        if updated_count == 0:
            messagebox.showinfo("Bulk tag assignment", "No changes were made.", parent=self)
            return

        mode_label = "Replaced" if replace_mode else "Added"
        messagebox.showinfo("Bulk tag assignment", f"{mode_label} tags for {updated_count} entries.", parent=self)

    def _adjust_selected_priority(self, increase: bool) -> None:
        entry_ids = self._selected_entry_ids()
        if not entry_ids:
            messagebox.showwarning("No selection", "Select at least one entry.", parent=self)
            return

        updated_count = 0
        gray_skipped = 0

        for entry_id in entry_ids:
            try:
                if increase:
                    self.repository.increase_priority(entry_id)
                else:
                    self.repository.decrease_priority(entry_id)
                updated_count += 1
            except ValueError:
                gray_skipped += 1
            except LookupError as exc:
                messagebox.showerror("Not found", str(exc), parent=self)
                continue

        self.refresh_entries()

        if gray_skipped > 0:
            messagebox.showinfo(
                "Priority update",
                f"Updated {updated_count} entries. Skipped {gray_skipped} gray-tier entries.",
                parent=self,
            )


class EnglishToJapaneseTestDialog(tk.Toplevel):
    def __init__(
        self,
        parent: tk.Tk,
        repository: VocabRepository,
        text_font: tkfont.Font,
        pick_strategy: str = "strict",
    ) -> None:
        super().__init__(parent)
        self.repository = repository
        self.text_font = text_font
        target_label_getter = getattr(parent, "_target_field_label", None)
        meaning_label_getter = getattr(parent, "_assistant_field_label", None)
        supports_kana_getter = getattr(parent, "_workbook_supports_kana", None)
        self.target_label = target_label_getter() if callable(target_label_getter) else "Target text"
        self.meaning_label = meaning_label_getter() if callable(meaning_label_getter) else "Meaning"
        self._show_kana_hint = bool(supports_kana_getter()) if callable(supports_kana_getter) else False

        self.questions: list[VocabEntry] = []
        self.current_index = 0
        self.correct_count = 0
        self.current_answered = False
        self.requested_count = 15
        self.actual_count = 0
        self.in_retry_mode = False
        self.retry_cycle = 0
        self.initial_failed_questions: list[VocabEntry] = []
        self.initial_failed_entry_ids: set[int] = set()
        self.retry_questions: list[VocabEntry] = []
        self.retry_failed_questions: list[VocabEntry] = []

        self.count_var = tk.StringVar(value="15")
        self.start_info_var = tk.StringVar(value="")
        self.progress_var = tk.StringVar(value="")
        self.score_var = tk.StringVar(value="")
        self.prompt_var = tk.StringVar(value="")
        self.answer_var = tk.StringVar(value="")
        self.feedback_var = tk.StringVar(value="")
        self.result_var = tk.StringVar(value="")
        self.pick_strategy_var = tk.StringVar(value=pick_strategy if pick_strategy in {"strict", "weighted"} else "strict")

        self.title(f"Test mode: {self.meaning_label} -> {self.target_label}")
        self.transient(parent)
        self.grab_set()
        self.resizable(False, False)

        self._build_widgets()
        self._show_start_frame()

        self.bind("<Return>", self._on_return_key)
        self.bind("<Escape>", lambda _event: self.destroy())

    def _build_widgets(self) -> None:
        root = ttk.Frame(self, padding=14)
        root.grid(row=0, column=0, sticky="nsew")
        root.columnconfigure(0, weight=1)

        self.start_frame = ttk.Frame(root)
        self.start_frame.grid(row=0, column=0, sticky="nsew")
        self.start_frame.columnconfigure(0, weight=1)

        ttk.Label(
            self.start_frame,
            text=f"{self.meaning_label} -> {self.target_label} test",
            style="App.TLabel",
        ).grid(row=0, column=0, sticky="w")

        ttk.Label(
            self.start_frame,
            text="Questions per test (default 15)",
            style="App.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(10, 4))

        self.count_entry = ttk.Entry(self.start_frame, textvariable=self.count_var, width=10, style="App.TEntry")
        self.count_entry.grid(row=2, column=0, sticky="w")

        ttk.Label(
            self.start_frame,
            text="Pick preference",
            style="App.TLabel",
        ).grid(row=3, column=0, sticky="w", pady=(10, 4))

        self.pick_strategy_combo = ttk.Combobox(
            self.start_frame,
            values=("strict", "weighted"),
            state="readonly",
            width=12,
            textvariable=self.pick_strategy_var,
        )
        self.pick_strategy_combo.grid(row=4, column=0, sticky="w")

        self.available_label = ttk.Label(self.start_frame, text="", style="Status.TLabel")
        self.available_label.grid(row=5, column=0, sticky="w", pady=(8, 0))

        self.start_info_label = ttk.Label(self.start_frame, textvariable=self.start_info_var, style="Status.TLabel")
        self.start_info_label.grid(row=6, column=0, sticky="w", pady=(2, 10))

        start_actions = ttk.Frame(self.start_frame)
        start_actions.grid(row=7, column=0, sticky="e")
        ttk.Button(start_actions, text="Close", command=self.destroy, style="App.TButton").grid(
            row=0,
            column=0,
            padx=(0, 8),
        )
        ttk.Button(start_actions, text="Start", command=self._start_test, style="App.TButton").grid(row=0, column=1)

        self.test_frame = ttk.Frame(root)
        self.test_frame.grid(row=0, column=0, sticky="nsew")
        self.test_frame.columnconfigure(0, weight=1)

        header_row = ttk.Frame(self.test_frame)
        header_row.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        header_row.columnconfigure(0, weight=1)
        ttk.Label(header_row, textvariable=self.progress_var, style="App.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(header_row, textvariable=self.score_var, style="App.TLabel").grid(row=0, column=1, sticky="e")

        ttk.Label(self.test_frame, text=self.meaning_label, style="App.TLabel").grid(row=1, column=0, sticky="w")
        self.prompt_label = ttk.Label(
            self.test_frame,
            textvariable=self.prompt_var,
            style="App.TLabel",
            wraplength=600,
            font=self.text_font,
        )
        self.prompt_label.grid(row=2, column=0, sticky="w", pady=(4, 12))

        ttk.Label(
            self.test_frame,
            text=f"Your {self.target_label} answer",
            style="App.TLabel",
        ).grid(row=3, column=0, sticky="w")
        self.answer_entry = ttk.Entry(self.test_frame, textvariable=self.answer_var, width=38, style="Japanese.TEntry")
        self.answer_entry.grid(row=4, column=0, sticky="w", pady=(4, 8))

        self.feedback_label = ttk.Label(self.test_frame, textvariable=self.feedback_var, style="Status.TLabel", wraplength=600)
        self.feedback_label.grid(row=5, column=0, sticky="w", pady=(0, 10))

        test_actions = ttk.Frame(self.test_frame)
        test_actions.grid(row=6, column=0, sticky="e")
        self.submit_button = ttk.Button(test_actions, text="Submit", command=self._submit_answer, style="App.TButton")
        self.submit_button.grid(row=0, column=0, padx=(0, 8))
        self.detail_button = ttk.Button(
            test_actions,
            text="View details",
            command=self._open_current_entry_detail,
            style="App.TButton",
            state="disabled",
        )
        self.detail_button.grid(row=0, column=1, padx=(0, 8))
        self.next_button = ttk.Button(test_actions, text="Next", command=self._next_question, style="App.TButton")
        self.next_button.grid(row=0, column=2)

        self.result_frame = ttk.Frame(root)
        self.result_frame.grid(row=0, column=0, sticky="nsew")
        self.result_frame.columnconfigure(0, weight=1)

        ttk.Label(self.result_frame, text="Test complete", style="App.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            self.result_frame,
            textvariable=self.result_var,
            style="App.TLabel",
            wraplength=600,
        ).grid(row=1, column=0, sticky="w", pady=(8, 12))

        result_actions = ttk.Frame(self.result_frame)
        result_actions.grid(row=2, column=0, sticky="e")
        ttk.Button(result_actions, text="Close", command=self.destroy, style="App.TButton").grid(
            row=0,
            column=0,
            padx=(0, 8),
        )
        ttk.Button(result_actions, text="Restart", command=self._restart, style="App.TButton").grid(row=0, column=1)

    def _show_start_frame(self) -> None:
        self.test_frame.grid_remove()
        self.result_frame.grid_remove()
        self.start_frame.grid()
        available = self.repository.count_entries()
        self.available_label.configure(text=f"Available vocabularies: {available}")
        self.count_entry.focus_set()
        self.count_entry.selection_range(0, "end")

    def _show_test_frame(self) -> None:
        self.start_frame.grid_remove()
        self.result_frame.grid_remove()
        self.test_frame.grid()

    def _show_result_frame(self) -> None:
        self.start_frame.grid_remove()
        self.test_frame.grid_remove()
        self.result_frame.grid()

    def _active_questions(self) -> list[VocabEntry]:
        return self.retry_questions if self.in_retry_mode else self.questions

    def _current_question(self) -> VocabEntry | None:
        active_questions = self._active_questions()
        if not active_questions or self.current_index >= len(active_questions):
            return None
        return active_questions[self.current_index]

    def _begin_retry_cycle(self, retry_questions: list[VocabEntry]) -> None:
        if not retry_questions:
            return

        self.in_retry_mode = True
        self.retry_cycle += 1
        self.retry_questions = list(retry_questions)
        self.retry_failed_questions = []
        self.current_index = 0
        self._show_test_frame()
        self._load_question()

    def _open_current_entry_detail(self) -> None:
        current = self._current_question()
        if current is None:
            return
        _open_detail_dialog_from_test(self, current.id)

    def _start_test(self) -> None:
        try:
            requested = int(self.count_var.get().strip())
            if requested <= 0:
                raise ValidationError("Questions per test must be a positive integer.")
        except ValueError:
            messagebox.showerror("Validation error", "Questions per test must be a positive integer.", parent=self)
            return
        except ValidationError as exc:
            messagebox.showerror("Validation error", str(exc), parent=self)
            return

        available = self.repository.count_entries()
        if available <= 0:
            messagebox.showerror("No vocabularies", "Add at least one vocabulary before starting a test.", parent=self)
            return

        request_count = min(requested, available)
        strategy = self.pick_strategy_var.get().strip().lower()
        if strategy not in {"strict", "weighted"}:
            strategy = "strict"

        self.questions = self.repository.get_test_entries_by_preference(request_count, strategy)
        self.actual_count = len(self.questions)
        if self.actual_count == 0:
            messagebox.showerror("No questions", "Could not generate test questions.", parent=self)
            return

        self.requested_count = requested
        self.current_index = 0
        self.correct_count = 0
        self.in_retry_mode = False
        self.retry_cycle = 0
        self.initial_failed_questions = []
        self.initial_failed_entry_ids = set()
        self.retry_questions = []
        self.retry_failed_questions = []

        if requested > self.actual_count:
            self.start_info_var.set(f"Requested {requested}; using all {self.actual_count} available vocabularies.")
        else:
            self.start_info_var.set("")

        self._show_test_frame()
        self._load_question()

    def _load_question(self) -> None:
        active_questions = self._active_questions()
        if not active_questions:
            return

        current = active_questions[self.current_index]
        active_count = len(active_questions)
        if self.in_retry_mode:
            self.progress_var.set(f"Retry {self.retry_cycle} - Question {self.current_index + 1}/{active_count}")
        else:
            self.progress_var.set(f"Question {self.current_index + 1}/{self.actual_count}")
        self.score_var.set(f"Score: {self.correct_count}")
        self.prompt_var.set(current.english_text)
        self.answer_var.set("")
        self.feedback_var.set("")
        self.current_answered = False
        self.answer_entry.configure(state="normal")
        self.submit_button.configure(state="normal")
        self.detail_button.configure(state="disabled")
        self.next_button.configure(state="disabled")
        self.answer_entry.focus_set()

    def _submit_answer(self) -> None:
        if self.current_answered:
            return

        current = self._current_question()
        if current is None:
            return

        submitted = self.answer_var.get().strip()
        correct_answer = current.japanese_text.strip()
        is_correct = submitted == correct_answer

        try:
            self.repository.record_test_result(current.id, is_correct)
        except LookupError as exc:
            messagebox.showerror("Not found", str(exc), parent=self)
            return
        except sqlite3.Error as exc:
            messagebox.showerror("Database error", f"Could not save test result: {exc}", parent=self)
            return

        if is_correct:
            if not self.in_retry_mode:
                self.correct_count += 1
            self.feedback_var.set("Correct.")
            self.detail_button.configure(state="disabled")
        else:
            if self.in_retry_mode:
                self.retry_failed_questions.append(current)
            elif current.id not in self.initial_failed_entry_ids:
                self.initial_failed_entry_ids.add(current.id)
                self.initial_failed_questions.append(current)

            kana_hint = ""
            if self._show_kana_hint:
                kana_hint = (current.kana_text or "").strip()
                if not kana_hint:
                    suggested_kana, reliable, _message = suggest_hiragana(correct_answer)
                    if reliable and suggested_kana:
                        kana_hint = suggested_kana

            if self._show_kana_hint and kana_hint:
                self.feedback_var.set(f"Incorrect. Correct answer: {correct_answer} ({kana_hint})")
            else:
                self.feedback_var.set(f"Incorrect. Correct answer: {correct_answer}")
            self.detail_button.configure(state="normal")

        self.score_var.set(f"Score: {self.correct_count}")
        self.current_answered = True
        self.answer_entry.configure(state="disabled")
        self.submit_button.configure(state="disabled")
        self.next_button.configure(state="normal")
        self.next_button.focus_set()

    def _next_question(self) -> None:
        if not self.current_answered:
            return

        active_count = len(self._active_questions())
        if self.current_index + 1 >= active_count:
            self._finish_test()
            return

        self.current_index += 1
        self._load_question()

    def _finish_test(self) -> None:
        if self.actual_count <= 0:
            self.result_var.set("No questions were completed.")
            self._show_result_frame()
            return

        if not self.in_retry_mode and self.initial_failed_questions:
            self._begin_retry_cycle(self.initial_failed_questions)
            return

        if self.in_retry_mode and self.retry_failed_questions:
            self._begin_retry_cycle(self.retry_failed_questions)
            return

        accuracy = (self.correct_count / self.actual_count) * 100
        result_text = f"Score: {self.correct_count}/{self.actual_count} ({accuracy:.1f}%)."
        if self.initial_failed_questions:
            result_text += (
                f" Initially missed: {len(self.initial_failed_questions)}. "
                f"Retried until correct in {self.retry_cycle} cycle(s)."
            )
        self.result_var.set(result_text)
        self._show_result_frame()

    def _restart(self) -> None:
        self.questions = []
        self.current_index = 0
        self.correct_count = 0
        self.actual_count = 0
        self.current_answered = False
        self.in_retry_mode = False
        self.retry_cycle = 0
        self.initial_failed_questions = []
        self.initial_failed_entry_ids = set()
        self.retry_questions = []
        self.retry_failed_questions = []
        self.progress_var.set("")
        self.score_var.set("")
        self.prompt_var.set("")
        self.answer_var.set("")
        self.feedback_var.set("")
        self.result_var.set("")
        self.detail_button.configure(state="disabled")
        self._show_start_frame()

    def _on_return_key(self, _event: tk.Event) -> str:
        if self.start_frame.winfo_ismapped():
            self._start_test()
            return "break"

        if self.test_frame.winfo_ismapped():
            if self.current_answered:
                self._next_question()
            else:
                self._submit_answer()
            return "break"

        if self.result_frame.winfo_ismapped():
            self._restart()
            return "break"

        return "break"


class JapaneseToKanaTestDialog(tk.Toplevel):
    def __init__(
        self,
        parent: tk.Tk,
        repository: VocabRepository,
        text_font: tkfont.Font,
        pick_strategy: str = "strict",
    ) -> None:
        super().__init__(parent)
        self.repository = repository
        self.text_font = text_font
        target_label_getter = getattr(parent, "_target_field_label", None)
        self.target_label = target_label_getter() if callable(target_label_getter) else "Target text"

        self.questions: list[VocabEntry] = []
        self.current_index = 0
        self.correct_count = 0
        self.current_answered = False
        self.requested_count = 15
        self.actual_count = 0
        self.in_retry_mode = False
        self.retry_cycle = 0
        self.initial_failed_questions: list[VocabEntry] = []
        self.initial_failed_entry_ids: set[int] = set()
        self.retry_questions: list[VocabEntry] = []
        self.retry_failed_questions: list[VocabEntry] = []

        self.count_var = tk.StringVar(value="15")
        self.start_info_var = tk.StringVar(value="")
        self.progress_var = tk.StringVar(value="")
        self.score_var = tk.StringVar(value="")
        self.prompt_var = tk.StringVar(value="")
        self.answer_var = tk.StringVar(value="")
        self.feedback_var = tk.StringVar(value="")
        self.result_var = tk.StringVar(value="")
        self.pick_strategy_var = tk.StringVar(value=pick_strategy if pick_strategy in {"strict", "weighted"} else "strict")

        self.title(f"Test mode: {self.target_label} -> Kana")
        self.transient(parent)
        self.grab_set()
        self.resizable(False, False)

        self._build_widgets()
        self._show_start_frame()

        self.bind("<Return>", self._on_return_key)
        self.bind("<Escape>", lambda _event: self.destroy())

    def _build_widgets(self) -> None:
        root = ttk.Frame(self, padding=14)
        root.grid(row=0, column=0, sticky="nsew")
        root.columnconfigure(0, weight=1)

        self.start_frame = ttk.Frame(root)
        self.start_frame.grid(row=0, column=0, sticky="nsew")
        self.start_frame.columnconfigure(0, weight=1)

        ttk.Label(
            self.start_frame,
            text=f"{self.target_label} -> Kana test",
            style="App.TLabel",
        ).grid(row=0, column=0, sticky="w")

        ttk.Label(
            self.start_frame,
            text="Questions per test (default 15)",
            style="App.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(10, 4))

        self.count_entry = ttk.Entry(self.start_frame, textvariable=self.count_var, width=10, style="App.TEntry")
        self.count_entry.grid(row=2, column=0, sticky="w")

        ttk.Label(
            self.start_frame,
            text="Pick preference",
            style="App.TLabel",
        ).grid(row=3, column=0, sticky="w", pady=(10, 4))

        self.pick_strategy_combo = ttk.Combobox(
            self.start_frame,
            values=("strict", "weighted"),
            state="readonly",
            width=12,
            textvariable=self.pick_strategy_var,
        )
        self.pick_strategy_combo.grid(row=4, column=0, sticky="w")

        self.available_label = ttk.Label(self.start_frame, text="", style="Status.TLabel")
        self.available_label.grid(row=5, column=0, sticky="w", pady=(8, 0))

        self.start_info_label = ttk.Label(self.start_frame, textvariable=self.start_info_var, style="Status.TLabel")
        self.start_info_label.grid(row=6, column=0, sticky="w", pady=(2, 10))

        start_actions = ttk.Frame(self.start_frame)
        start_actions.grid(row=7, column=0, sticky="e")
        ttk.Button(start_actions, text="Close", command=self.destroy, style="App.TButton").grid(
            row=0,
            column=0,
            padx=(0, 8),
        )
        ttk.Button(start_actions, text="Start", command=self._start_test, style="App.TButton").grid(row=0, column=1)

        self.test_frame = ttk.Frame(root)
        self.test_frame.grid(row=0, column=0, sticky="nsew")
        self.test_frame.columnconfigure(0, weight=1)

        header_row = ttk.Frame(self.test_frame)
        header_row.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        header_row.columnconfigure(0, weight=1)
        ttk.Label(header_row, textvariable=self.progress_var, style="App.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(header_row, textvariable=self.score_var, style="App.TLabel").grid(row=0, column=1, sticky="e")

        ttk.Label(self.test_frame, text=self.target_label, style="App.TLabel").grid(row=1, column=0, sticky="w")
        self.prompt_label = ttk.Label(
            self.test_frame,
            textvariable=self.prompt_var,
            style="App.TLabel",
            wraplength=600,
            font=self.text_font,
        )
        self.prompt_label.grid(row=2, column=0, sticky="w", pady=(4, 12))

        ttk.Label(self.test_frame, text="Your kana answer", style="App.TLabel").grid(row=3, column=0, sticky="w")
        self.answer_entry = ttk.Entry(self.test_frame, textvariable=self.answer_var, width=38, style="Japanese.TEntry")
        self.answer_entry.grid(row=4, column=0, sticky="w", pady=(4, 8))

        self.feedback_label = ttk.Label(self.test_frame, textvariable=self.feedback_var, style="Status.TLabel", wraplength=600)
        self.feedback_label.grid(row=5, column=0, sticky="w", pady=(0, 10))

        test_actions = ttk.Frame(self.test_frame)
        test_actions.grid(row=6, column=0, sticky="e")
        self.submit_button = ttk.Button(test_actions, text="Submit", command=self._submit_answer, style="App.TButton")
        self.submit_button.grid(row=0, column=0, padx=(0, 8))
        self.detail_button = ttk.Button(
            test_actions,
            text="View details",
            command=self._open_current_entry_detail,
            style="App.TButton",
            state="disabled",
        )
        self.detail_button.grid(row=0, column=1, padx=(0, 8))
        self.next_button = ttk.Button(test_actions, text="Next", command=self._next_question, style="App.TButton")
        self.next_button.grid(row=0, column=2)

        self.result_frame = ttk.Frame(root)
        self.result_frame.grid(row=0, column=0, sticky="nsew")
        self.result_frame.columnconfigure(0, weight=1)

        ttk.Label(self.result_frame, text="Test complete", style="App.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            self.result_frame,
            textvariable=self.result_var,
            style="App.TLabel",
            wraplength=600,
        ).grid(row=1, column=0, sticky="w", pady=(8, 12))

        result_actions = ttk.Frame(self.result_frame)
        result_actions.grid(row=2, column=0, sticky="e")
        ttk.Button(result_actions, text="Close", command=self.destroy, style="App.TButton").grid(
            row=0,
            column=0,
            padx=(0, 8),
        )
        ttk.Button(result_actions, text="Restart", command=self._restart, style="App.TButton").grid(row=0, column=1)

    @staticmethod
    def _entry_has_kana(entry: VocabEntry) -> bool:
        return bool((entry.kana_text or "").strip())

    def _build_eligible_questions(self, strategy: str) -> list[VocabEntry]:
        available = self.repository.count_entries()
        if available <= 0:
            return []

        ordered_entries = self.repository.get_test_entries_by_preference(available, strategy)
        return [entry for entry in ordered_entries if self._entry_has_kana(entry)]

    def _show_start_frame(self) -> None:
        self.test_frame.grid_remove()
        self.result_frame.grid_remove()
        self.start_frame.grid()

        strategy = self.pick_strategy_var.get().strip().lower()
        if strategy not in {"strict", "weighted"}:
            strategy = "strict"

        available = len(self._build_eligible_questions(strategy))
        self.available_label.configure(text=f"Available vocabularies with kana: {available}")
        self.count_entry.focus_set()
        self.count_entry.selection_range(0, "end")

    def _show_test_frame(self) -> None:
        self.start_frame.grid_remove()
        self.result_frame.grid_remove()
        self.test_frame.grid()

    def _show_result_frame(self) -> None:
        self.start_frame.grid_remove()
        self.test_frame.grid_remove()
        self.result_frame.grid()

    def _active_questions(self) -> list[VocabEntry]:
        return self.retry_questions if self.in_retry_mode else self.questions

    def _current_question(self) -> VocabEntry | None:
        active_questions = self._active_questions()
        if not active_questions or self.current_index >= len(active_questions):
            return None
        return active_questions[self.current_index]

    def _begin_retry_cycle(self, retry_questions: list[VocabEntry]) -> None:
        eligible_retry_questions = [entry for entry in retry_questions if self._entry_has_kana(entry)]
        if not eligible_retry_questions:
            return

        self.in_retry_mode = True
        self.retry_cycle += 1
        self.retry_questions = eligible_retry_questions
        self.retry_failed_questions = []
        self.current_index = 0
        self._show_test_frame()
        self._load_question()

    def _open_current_entry_detail(self) -> None:
        current = self._current_question()
        if current is None:
            return
        _open_detail_dialog_from_test(self, current.id)

    def _start_test(self) -> None:
        try:
            requested = int(self.count_var.get().strip())
            if requested <= 0:
                raise ValidationError("Questions per test must be a positive integer.")
        except ValueError:
            messagebox.showerror("Validation error", "Questions per test must be a positive integer.", parent=self)
            return
        except ValidationError as exc:
            messagebox.showerror("Validation error", str(exc), parent=self)
            return

        strategy = self.pick_strategy_var.get().strip().lower()
        if strategy not in {"strict", "weighted"}:
            strategy = "strict"

        eligible_questions = self._build_eligible_questions(strategy)
        if not eligible_questions:
            messagebox.showerror(
                "No eligible vocabularies",
                f"Add vocabularies with kana before starting a {self.target_label}->Kana test.",
                parent=self,
            )
            return

        request_count = min(requested, len(eligible_questions))
        self.questions = eligible_questions[:request_count]
        self.actual_count = len(self.questions)
        if self.actual_count == 0:
            messagebox.showerror("No questions", "Could not generate test questions.", parent=self)
            return

        self.requested_count = requested
        self.current_index = 0
        self.correct_count = 0
        self.in_retry_mode = False
        self.retry_cycle = 0
        self.initial_failed_questions = []
        self.initial_failed_entry_ids = set()
        self.retry_questions = []
        self.retry_failed_questions = []

        if requested > self.actual_count:
            self.start_info_var.set(f"Requested {requested}; using all {self.actual_count} eligible vocabularies.")
        else:
            self.start_info_var.set("")

        self._show_test_frame()
        self._load_question()

    def _load_question(self) -> None:
        active_questions = self._active_questions()
        if not active_questions:
            return

        current = active_questions[self.current_index]
        active_count = len(active_questions)
        if self.in_retry_mode:
            self.progress_var.set(f"Retry {self.retry_cycle} - Question {self.current_index + 1}/{active_count}")
        else:
            self.progress_var.set(f"Question {self.current_index + 1}/{self.actual_count}")
        self.score_var.set(f"Score: {self.correct_count}")
        self.prompt_var.set(current.japanese_text)
        self.answer_var.set("")
        self.feedback_var.set("")
        self.current_answered = False
        self.answer_entry.configure(state="normal")
        self.submit_button.configure(state="normal")
        self.detail_button.configure(state="disabled")
        self.next_button.configure(state="disabled")
        self.answer_entry.focus_set()

    def _submit_answer(self) -> None:
        if self.current_answered:
            return

        current = self._current_question()
        if current is None:
            return

        submitted = self.answer_var.get().strip()
        correct_answer = (current.kana_text or "").strip()
        is_correct = submitted == correct_answer

        try:
            self.repository.record_test_result(current.id, is_correct)
        except LookupError as exc:
            messagebox.showerror("Not found", str(exc), parent=self)
            return
        except sqlite3.Error as exc:
            messagebox.showerror("Database error", f"Could not save test result: {exc}", parent=self)
            return

        if is_correct:
            if not self.in_retry_mode:
                self.correct_count += 1
            self.feedback_var.set("Correct.")
            self.detail_button.configure(state="disabled")
        else:
            if self.in_retry_mode:
                self.retry_failed_questions.append(current)
            elif current.id not in self.initial_failed_entry_ids:
                self.initial_failed_entry_ids.add(current.id)
                self.initial_failed_questions.append(current)
            self.feedback_var.set(f"Incorrect. Correct answer: {correct_answer}")
            self.detail_button.configure(state="normal")

        self.score_var.set(f"Score: {self.correct_count}")
        self.current_answered = True
        self.answer_entry.configure(state="disabled")
        self.submit_button.configure(state="disabled")
        self.next_button.configure(state="normal")
        self.next_button.focus_set()

    def _next_question(self) -> None:
        if not self.current_answered:
            return

        active_count = len(self._active_questions())
        if self.current_index + 1 >= active_count:
            self._finish_test()
            return

        self.current_index += 1
        self._load_question()

    def _finish_test(self) -> None:
        if self.actual_count <= 0:
            self.result_var.set("No questions were completed.")
            self._show_result_frame()
            return

        if not self.in_retry_mode and self.initial_failed_questions:
            self._begin_retry_cycle(self.initial_failed_questions)
            return

        if self.in_retry_mode and self.retry_failed_questions:
            self._begin_retry_cycle(self.retry_failed_questions)
            return

        accuracy = (self.correct_count / self.actual_count) * 100
        result_text = f"Score: {self.correct_count}/{self.actual_count} ({accuracy:.1f}%)."
        if self.initial_failed_questions:
            result_text += (
                f" Initially missed: {len(self.initial_failed_questions)}. "
                f"Retried until correct in {self.retry_cycle} cycle(s)."
            )
        self.result_var.set(result_text)
        self._show_result_frame()

    def _restart(self) -> None:
        self.questions = []
        self.current_index = 0
        self.correct_count = 0
        self.actual_count = 0
        self.current_answered = False
        self.in_retry_mode = False
        self.retry_cycle = 0
        self.initial_failed_questions = []
        self.initial_failed_entry_ids = set()
        self.retry_questions = []
        self.retry_failed_questions = []
        self.progress_var.set("")
        self.score_var.set("")
        self.prompt_var.set("")
        self.answer_var.set("")
        self.feedback_var.set("")
        self.result_var.set("")
        self.detail_button.configure(state="disabled")
        self._show_start_frame()

    def _on_return_key(self, _event: tk.Event) -> str:
        if self.start_frame.winfo_ismapped():
            self._start_test()
            return "break"

        if self.test_frame.winfo_ismapped():
            if self.current_answered:
                self._next_question()
            else:
                self._submit_answer()
            return "break"

        if self.result_frame.winfo_ismapped():
            self._restart()
            return "break"

        return "break"


class JapaneseToEnglishChoiceTestDialog(tk.Toplevel):
    def __init__(
        self,
        parent: tk.Tk,
        repository: VocabRepository,
        text_font: tkfont.Font,
        pick_strategy: str = "strict",
    ) -> None:
        super().__init__(parent)
        self.repository = repository
        self.text_font = text_font
        target_label_getter = getattr(parent, "_target_field_label", None)
        meaning_label_getter = getattr(parent, "_assistant_field_label", None)
        self.target_label = target_label_getter() if callable(target_label_getter) else "Target text"
        self.meaning_label = meaning_label_getter() if callable(meaning_label_getter) else "Meaning"

        self.questions: list[VocabEntry] = []
        self.options_by_question: list[list[str]] = []
        self.current_index = 0
        self.correct_count = 0
        self.current_answered = False
        self.requested_count = 15
        self.actual_count = 0
        self.in_retry_mode = False
        self.retry_cycle = 0
        self.initial_failed_questions: list[VocabEntry] = []
        self.initial_failed_entry_ids: set[int] = set()
        self.retry_questions: list[VocabEntry] = []
        self.retry_failed_questions: list[VocabEntry] = []

        self.count_var = tk.StringVar(value="15")
        self.start_info_var = tk.StringVar(value="")
        self.progress_var = tk.StringVar(value="")
        self.score_var = tk.StringVar(value="")
        self.prompt_var = tk.StringVar(value="")
        self.feedback_var = tk.StringVar(value="")
        self.result_var = tk.StringVar(value="")
        self.pick_strategy_var = tk.StringVar(value=pick_strategy if pick_strategy in {"strict", "weighted"} else "strict")

        self.title(f"Test mode: {self.target_label} -> {self.meaning_label}")
        self.transient(parent)
        self.grab_set()
        self.resizable(False, False)

        self._build_widgets()
        self._show_start_frame()

        self.bind("<Return>", self._on_return_key)
        self.bind("<Escape>", lambda _event: self.destroy())

    def _build_widgets(self) -> None:
        root = ttk.Frame(self, padding=14)
        root.grid(row=0, column=0, sticky="nsew")
        root.columnconfigure(0, weight=1)

        self.start_frame = ttk.Frame(root)
        self.start_frame.grid(row=0, column=0, sticky="nsew")
        self.start_frame.columnconfigure(0, weight=1)

        ttk.Label(
            self.start_frame,
            text=f"{self.target_label} -> {self.meaning_label} test (single choice)",
            style="App.TLabel",
        ).grid(row=0, column=0, sticky="w")

        ttk.Label(
            self.start_frame,
            text="Questions per test (default 15)",
            style="App.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(10, 4))

        self.count_entry = ttk.Entry(self.start_frame, textvariable=self.count_var, width=10, style="App.TEntry")
        self.count_entry.grid(row=2, column=0, sticky="w")

        ttk.Label(
            self.start_frame,
            text="Pick preference",
            style="App.TLabel",
        ).grid(row=3, column=0, sticky="w", pady=(10, 4))

        self.pick_strategy_combo = ttk.Combobox(
            self.start_frame,
            values=("strict", "weighted"),
            state="readonly",
            width=12,
            textvariable=self.pick_strategy_var,
        )
        self.pick_strategy_combo.grid(row=4, column=0, sticky="w")

        self.available_label = ttk.Label(self.start_frame, text="", style="Status.TLabel")
        self.available_label.grid(row=5, column=0, sticky="w", pady=(8, 0))

        self.start_info_label = ttk.Label(self.start_frame, textvariable=self.start_info_var, style="Status.TLabel")
        self.start_info_label.grid(row=6, column=0, sticky="w", pady=(2, 10))

        start_actions = ttk.Frame(self.start_frame)
        start_actions.grid(row=7, column=0, sticky="e")
        ttk.Button(start_actions, text="Close", command=self.destroy, style="App.TButton").grid(
            row=0,
            column=0,
            padx=(0, 8),
        )
        ttk.Button(start_actions, text="Start", command=self._start_test, style="App.TButton").grid(row=0, column=1)

        self.test_frame = ttk.Frame(root)
        self.test_frame.grid(row=0, column=0, sticky="nsew")
        self.test_frame.columnconfigure(0, weight=1)

        header_row = ttk.Frame(self.test_frame)
        header_row.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        header_row.columnconfigure(0, weight=1)
        ttk.Label(header_row, textvariable=self.progress_var, style="App.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(header_row, textvariable=self.score_var, style="App.TLabel").grid(row=0, column=1, sticky="e")

        ttk.Label(
            self.test_frame,
            text=self.target_label,
            style="App.TLabel",
        ).grid(row=1, column=0, sticky="w")
        self.prompt_label = ttk.Label(
            self.test_frame,
            textvariable=self.prompt_var,
            style="App.TLabel",
            wraplength=600,
            font=self.text_font,
        )
        self.prompt_label.grid(row=2, column=0, sticky="w", pady=(4, 12))

        ttk.Label(self.test_frame, text=f"Choose the correct {self.meaning_label.lower()}", style="App.TLabel").grid(
            row=3,
            column=0,
            sticky="w",
        )

        options_frame = ttk.Frame(self.test_frame)
        options_frame.grid(row=4, column=0, sticky="ew", pady=(6, 8))
        options_frame.columnconfigure(0, weight=1)

        self.option_buttons: list[ttk.Button] = []
        for index in range(4):
            button = ttk.Button(options_frame, text="", style="App.TButton")
            button.grid(row=index, column=0, sticky="ew", pady=(0, 6))
            self.option_buttons.append(button)

        self.feedback_label = ttk.Label(self.test_frame, textvariable=self.feedback_var, style="Status.TLabel", wraplength=600)
        self.feedback_label.grid(row=5, column=0, sticky="w", pady=(0, 10))

        test_actions = ttk.Frame(self.test_frame)
        test_actions.grid(row=6, column=0, sticky="e")
        ttk.Button(test_actions, text="Close", command=self.destroy, style="App.TButton").grid(
            row=0,
            column=0,
            padx=(0, 8),
        )
        self.detail_button = ttk.Button(
            test_actions,
            text="View details",
            command=self._open_current_entry_detail,
            style="App.TButton",
            state="disabled",
        )
        self.detail_button.grid(row=0, column=1, padx=(0, 8))
        self.next_button = ttk.Button(test_actions, text="Next", command=self._next_question, style="App.TButton")
        self.next_button.grid(row=0, column=2)

        self.result_frame = ttk.Frame(root)
        self.result_frame.grid(row=0, column=0, sticky="nsew")
        self.result_frame.columnconfigure(0, weight=1)

        ttk.Label(self.result_frame, text="Test complete", style="App.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            self.result_frame,
            textvariable=self.result_var,
            style="App.TLabel",
            wraplength=600,
        ).grid(row=1, column=0, sticky="w", pady=(8, 12))

        result_actions = ttk.Frame(self.result_frame)
        result_actions.grid(row=2, column=0, sticky="e")
        ttk.Button(result_actions, text="Close", command=self.destroy, style="App.TButton").grid(
            row=0,
            column=0,
            padx=(0, 8),
        )
        ttk.Button(result_actions, text="Restart", command=self._restart, style="App.TButton").grid(row=0, column=1)

    def _build_eligible_questions(self, strategy: str) -> tuple[list[VocabEntry], list[list[str]]]:
        available = self.repository.count_entries()
        if available <= 0:
            return [], []

        ordered_entries = self.repository.get_test_entries_by_preference(available, strategy)

        eligible_questions: list[VocabEntry] = []
        eligible_options: list[list[str]] = []
        for entry in ordered_entries:
            try:
                options = self.repository.get_english_options_for_entry(entry.id, max_options=4)
            except LookupError:
                continue

            if len(options) < 2:
                continue

            eligible_questions.append(entry)
            eligible_options.append(options)

        return eligible_questions, eligible_options

    def _show_start_frame(self) -> None:
        self.test_frame.grid_remove()
        self.result_frame.grid_remove()
        self.start_frame.grid()

        strategy = self.pick_strategy_var.get().strip().lower()
        if strategy not in {"strict", "weighted"}:
            strategy = "strict"

        available_questions, _ = self._build_eligible_questions(strategy)
        self.available_label.configure(
            text=f"Available vocabularies with at least 2 choices: {len(available_questions)}"
        )

        self.count_entry.focus_set()
        self.count_entry.selection_range(0, "end")

    def _show_test_frame(self) -> None:
        self.start_frame.grid_remove()
        self.result_frame.grid_remove()
        self.test_frame.grid()

    def _show_result_frame(self) -> None:
        self.start_frame.grid_remove()
        self.test_frame.grid_remove()
        self.result_frame.grid()

    def _active_questions(self) -> list[VocabEntry]:
        return self.retry_questions if self.in_retry_mode else self.questions

    def _current_question(self) -> VocabEntry | None:
        active_questions = self._active_questions()
        if not active_questions or self.current_index >= len(active_questions):
            return None
        return active_questions[self.current_index]

    def _options_for_current_question(self, current: VocabEntry) -> list[str]:
        if not self.in_retry_mode:
            if self.current_index < len(self.options_by_question):
                return self.options_by_question[self.current_index]
            return [current.english_text]

        try:
            options = self.repository.get_english_options_for_entry(current.id, max_options=4)
        except LookupError:
            options = [current.english_text]

        if current.english_text not in options:
            options = [current.english_text, *options]
        return options[:4]

    def _begin_retry_cycle(self, retry_questions: list[VocabEntry]) -> None:
        if not retry_questions:
            return

        self.in_retry_mode = True
        self.retry_cycle += 1
        self.retry_questions = list(retry_questions)
        self.retry_failed_questions = []
        self.current_index = 0
        self._show_test_frame()
        self._load_question()

    def _open_current_entry_detail(self) -> None:
        current = self._current_question()
        if current is None:
            return
        _open_detail_dialog_from_test(self, current.id)

    def _start_test(self) -> None:
        try:
            requested = int(self.count_var.get().strip())
            if requested <= 0:
                raise ValidationError("Questions per test must be a positive integer.")
        except ValueError:
            messagebox.showerror("Validation error", "Questions per test must be a positive integer.", parent=self)
            return
        except ValidationError as exc:
            messagebox.showerror("Validation error", str(exc), parent=self)
            return

        strategy = self.pick_strategy_var.get().strip().lower()
        if strategy not in {"strict", "weighted"}:
            strategy = "strict"

        eligible_questions, eligible_options = self._build_eligible_questions(strategy)
        if not eligible_questions:
            messagebox.showerror(
                "No eligible vocabularies",
                "Add at least two vocabularies with different meanings before starting this test.",
                parent=self,
            )
            return

        request_count = min(requested, len(eligible_questions))
        self.questions = eligible_questions[:request_count]
        self.options_by_question = eligible_options[:request_count]
        self.actual_count = len(self.questions)

        if self.actual_count == 0:
            messagebox.showerror("No questions", "Could not generate test questions.", parent=self)
            return

        self.requested_count = requested
        self.current_index = 0
        self.correct_count = 0
        self.in_retry_mode = False
        self.retry_cycle = 0
        self.initial_failed_questions = []
        self.initial_failed_entry_ids = set()
        self.retry_questions = []
        self.retry_failed_questions = []

        if requested > self.actual_count:
            self.start_info_var.set(f"Requested {requested}; using all {self.actual_count} eligible vocabularies.")
        else:
            self.start_info_var.set("")

        self._show_test_frame()
        self._load_question()

    def _load_question(self) -> None:
        active_questions = self._active_questions()
        if not active_questions:
            return

        current = active_questions[self.current_index]
        options = self._options_for_current_question(current)
        active_count = len(active_questions)

        if self.in_retry_mode:
            self.progress_var.set(f"Retry {self.retry_cycle} - Question {self.current_index + 1}/{active_count}")
        else:
            self.progress_var.set(f"Question {self.current_index + 1}/{self.actual_count}")
        self.score_var.set(f"Score: {self.correct_count}")
        self.prompt_var.set(current.japanese_text)
        self.feedback_var.set("")
        self.current_answered = False
        self.detail_button.configure(state="disabled")

        for index, button in enumerate(self.option_buttons):
            if index < len(options):
                option_text = options[index]
                button.configure(
                    text=option_text,
                    command=lambda selected=option_text: self._submit_choice(selected),
                    state="normal",
                )
                button.grid()
            else:
                button.grid_remove()

        self.next_button.configure(state="disabled")

    def _submit_choice(self, selected_option: str) -> None:
        if self.current_answered:
            return

        current = self._current_question()
        if current is None:
            return

        correct_answer = current.english_text.strip()
        is_correct = selected_option.strip() == correct_answer

        try:
            self.repository.record_test_result(current.id, is_correct)
        except LookupError as exc:
            messagebox.showerror("Not found", str(exc), parent=self)
            return
        except sqlite3.Error as exc:
            messagebox.showerror("Database error", f"Could not save test result: {exc}", parent=self)
            return

        if is_correct:
            if not self.in_retry_mode:
                self.correct_count += 1
            self.feedback_var.set("Correct.")
            self.detail_button.configure(state="disabled")
        else:
            if self.in_retry_mode:
                self.retry_failed_questions.append(current)
            elif current.id not in self.initial_failed_entry_ids:
                self.initial_failed_entry_ids.add(current.id)
                self.initial_failed_questions.append(current)
            self.feedback_var.set(f"Incorrect. Correct answer: {correct_answer}")
            self.detail_button.configure(state="normal")

        self.score_var.set(f"Score: {self.correct_count}")
        self.current_answered = True

        for button in self.option_buttons:
            if button.winfo_ismapped():
                button.configure(state="disabled")

        self.next_button.configure(state="normal")
        self.next_button.focus_set()

    def _next_question(self) -> None:
        if not self.current_answered:
            return

        active_count = len(self._active_questions())
        if self.current_index + 1 >= active_count:
            self._finish_test()
            return

        self.current_index += 1
        self._load_question()

    def _finish_test(self) -> None:
        if self.actual_count <= 0:
            self.result_var.set("No questions were completed.")
            self._show_result_frame()
            return

        if not self.in_retry_mode and self.initial_failed_questions:
            self._begin_retry_cycle(self.initial_failed_questions)
            return

        if self.in_retry_mode and self.retry_failed_questions:
            self._begin_retry_cycle(self.retry_failed_questions)
            return

        accuracy = (self.correct_count / self.actual_count) * 100
        result_text = f"Score: {self.correct_count}/{self.actual_count} ({accuracy:.1f}%)."
        if self.initial_failed_questions:
            result_text += (
                f" Initially missed: {len(self.initial_failed_questions)}. "
                f"Retried until correct in {self.retry_cycle} cycle(s)."
            )
        self.result_var.set(result_text)
        self._show_result_frame()

    def _restart(self) -> None:
        self.questions = []
        self.options_by_question = []
        self.current_index = 0
        self.correct_count = 0
        self.actual_count = 0
        self.current_answered = False
        self.in_retry_mode = False
        self.retry_cycle = 0
        self.initial_failed_questions = []
        self.initial_failed_entry_ids = set()
        self.retry_questions = []
        self.retry_failed_questions = []
        self.progress_var.set("")
        self.score_var.set("")
        self.prompt_var.set("")
        self.feedback_var.set("")
        self.result_var.set("")
        self.detail_button.configure(state="disabled")
        self._show_start_frame()

    def _on_return_key(self, _event: tk.Event) -> str:
        if self.start_frame.winfo_ismapped():
            self._start_test()
            return "break"

        if self.test_frame.winfo_ismapped():
            if self.current_answered:
                self._next_question()
            return "break"

        if self.result_frame.winfo_ismapped():
            self._restart()
            return "break"

        return "break"


class BulkAddDialog(tk.Toplevel):
    def __init__(
        self,
        parent: tk.Tk,
        repository: VocabRepository,
        on_saved: Callable[[], None],
        text_font: tkfont.Font,
        target_label: str,
        assistant_label: str,
        show_kana: bool,
    ) -> None:
        super().__init__(parent)
        self.repository = repository
        self.on_saved = on_saved
        self.target_label = target_label
        self.assistant_label = assistant_label
        self._show_kana = show_kana

        self.title("Bulk add vocabulary")
        self.transient(parent)
        self.grab_set()
        self.resizable(True, True)
        self.minsize(860, 420)

        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        frame = ttk.Frame(self, padding=12)
        frame.grid(row=0, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(2, weight=1)

        if self._show_kana:
            help_text = (
                f"One entry per line, aligned across 3 columns: {self.target_label}, "
                f"Kana (optional), {self.assistant_label}."
            )
        else:
            help_text = f"One entry per line, aligned across 2 columns: {self.target_label} and {self.assistant_label}."
        ttk.Label(frame, text=help_text, style="App.TLabel", wraplength=820).grid(
            row=0,
            column=0,
            sticky="w",
            pady=(0, 8),
        )

        labels_row = ttk.Frame(frame)
        labels_row.grid(row=1, column=0, sticky="ew", pady=(0, 4))
        labels_row.columnconfigure(0, weight=1)
        labels_row.columnconfigure(1, weight=1)
        labels_row.columnconfigure(2, weight=1)
        labels_row.columnconfigure(3, weight=0)

        ttk.Label(labels_row, text=f"{self.target_label} *", style="App.TLabel").grid(
            row=0,
            column=0,
            sticky="w",
            padx=(0, 8),
        )
        if self._show_kana:
            ttk.Label(labels_row, text="Kana (optional)", style="App.TLabel").grid(
                row=0,
                column=1,
                sticky="w",
                padx=(0, 8),
            )
        ttk.Label(labels_row, text=f"{self.assistant_label} *", style="App.TLabel").grid(
            row=0,
            column=2 if self._show_kana else 1,
            sticky="w",
            padx=(0, 8),
        )

        columns_frame = ttk.Frame(frame)
        columns_frame.grid(row=2, column=0, sticky="nsew")
        columns_frame.columnconfigure(0, weight=1)
        columns_frame.columnconfigure(1, weight=1)
        columns_frame.columnconfigure(2, weight=1)
        columns_frame.columnconfigure(3, weight=0)
        columns_frame.rowconfigure(0, weight=1)

        self.jp_text = tk.Text(columns_frame, width=26, height=14, wrap="none", font=text_font)
        self.kana_text = tk.Text(columns_frame, width=26, height=14, wrap="none", font=text_font)
        self.en_text = tk.Text(columns_frame, width=26, height=14, wrap="none", font=text_font)

        self.jp_text.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        if self._show_kana:
            self.kana_text.grid(row=0, column=1, sticky="nsew", padx=(0, 8))
            self.en_text.grid(row=0, column=2, sticky="nsew", padx=(0, 8))
        else:
            self.en_text.grid(row=0, column=1, sticky="nsew", padx=(0, 8))

        self.scrollbar = ttk.Scrollbar(columns_frame, orient="vertical", command=self._on_scrollbar)
        self.scrollbar.grid(row=0, column=3, sticky="ns")

        self.jp_text.configure(yscrollcommand=self._on_text_yscroll)
        self.kana_text.configure(yscrollcommand=self._on_text_yscroll)
        self.en_text.configure(yscrollcommand=self._on_text_yscroll)

        for text_widget in (self.jp_text, self.en_text):
            text_widget.bind("<MouseWheel>", self._on_mousewheel)
            text_widget.bind("<Button-4>", self._on_mousewheel_linux_up)
            text_widget.bind("<Button-5>", self._on_mousewheel_linux_down)
        if self._show_kana:
            self.kana_text.bind("<MouseWheel>", self._on_mousewheel)
            self.kana_text.bind("<Button-4>", self._on_mousewheel_linux_up)
            self.kana_text.bind("<Button-5>", self._on_mousewheel_linux_down)

        actions = ttk.Frame(frame, padding=(0, 10, 0, 0))
        actions.grid(row=3, column=0, sticky="e")

        action_column = 0
        if self._show_kana:
            ttk.Button(
                actions,
                text="Fill kana",
                command=self._auto_fill_missing_kana,
                style="App.TButton",
            ).grid(row=0, column=action_column, padx=(0, 8))
            action_column += 1

        ttk.Button(actions, text="Cancel", command=self.destroy, style="App.TButton").grid(
            row=0,
            column=action_column,
            padx=(0, 8),
        )
        action_column += 1
        ttk.Button(actions, text="Add entries", command=self._save, style="App.TButton").grid(
            row=0,
            column=action_column,
        )

        self.bind("<Escape>", lambda _event: self.destroy())
        self.jp_text.focus_set()

    def _on_scrollbar(self, *args: str) -> None:
        widgets = [self.jp_text, self.en_text]
        if self._show_kana:
            widgets.append(self.kana_text)
        for text_widget in widgets:
            text_widget.yview(*args)

    def _on_text_yscroll(self, first: str, last: str) -> None:
        self.scrollbar.set(first, last)

    def _on_mousewheel(self, event: tk.Event) -> str:
        delta = -1 if event.delta > 0 else 1
        widgets = [self.jp_text, self.en_text]
        if self._show_kana:
            widgets.append(self.kana_text)
        for text_widget in widgets:
            text_widget.yview_scroll(delta, "units")
        return "break"

    def _on_mousewheel_linux_up(self, _event: tk.Event) -> str:
        widgets = [self.jp_text, self.en_text]
        if self._show_kana:
            widgets.append(self.kana_text)
        for text_widget in widgets:
            text_widget.yview_scroll(-1, "units")
        return "break"

    def _on_mousewheel_linux_down(self, _event: tk.Event) -> str:
        widgets = [self.jp_text, self.en_text]
        if self._show_kana:
            widgets.append(self.kana_text)
        for text_widget in widgets:
            text_widget.yview_scroll(1, "units")
        return "break"

    def _parse_entries(self, target_raw: str, kana_raw: str, meaning_raw: str) -> list[tuple[str, str, str]]:
        target_lines = target_raw.splitlines()
        kana_lines = kana_raw.splitlines() if self._show_kana else []
        meaning_lines = meaning_raw.splitlines()

        line_count = max(len(target_lines), len(kana_lines), len(meaning_lines))
        parsed_entries: list[tuple[str, str, str]] = []

        for index in range(line_count):
            line_number = index + 1
            target_text = target_lines[index].strip() if index < len(target_lines) else ""
            kana_text = kana_lines[index].strip() if index < len(kana_lines) else ""
            meaning_text = meaning_lines[index].strip() if index < len(meaning_lines) else ""

            if not target_text and not kana_text and not meaning_text:
                continue

            if not target_text or not meaning_text:
                raise ValidationError(
                    f"Line {line_number}: {self.target_label} and {self.assistant_label} are required."
                )

            parsed_entries.append((target_text, kana_text, meaning_text))

        if not parsed_entries:
            raise ValidationError("Add at least one non-empty line.")

        return parsed_entries

    def _auto_fill_missing_kana(self) -> None:
        if not self._show_kana:
            return

        target_lines = self.jp_text.get("1.0", "end-1c").splitlines()
        kana_lines = self.kana_text.get("1.0", "end-1c").splitlines()
        line_count = max(len(target_lines), len(kana_lines))
        if line_count <= 0:
            messagebox.showinfo("Kana fill", f"Enter {self.target_label} lines first.", parent=self)
            return

        updated_count = 0
        unresolved_count = 0
        merged_kana_lines: list[str] = []

        for index in range(line_count):
            target_text = target_lines[index].strip() if index < len(target_lines) else ""
            existing_kana = kana_lines[index].strip() if index < len(kana_lines) else ""

            if existing_kana or not target_text:
                merged_kana_lines.append(existing_kana)
                continue

            suggested_kana, reliable, _message = suggest_hiragana(target_text)
            if reliable and suggested_kana:
                merged_kana_lines.append(suggested_kana)
                updated_count += 1
            else:
                merged_kana_lines.append("")
                unresolved_count += 1

        self.kana_text.delete("1.0", "end")
        self.kana_text.insert("1.0", "\n".join(merged_kana_lines))

        if updated_count == 0 and unresolved_count == 0:
            messagebox.showinfo("Kana fill", "No empty kana rows were found.", parent=self)
            return
        if unresolved_count == 0:
            messagebox.showinfo("Kana fill", f"Filled kana for {updated_count} rows.", parent=self)
            return

        messagebox.showwarning(
            "Kana fill",
            f"Filled kana for {updated_count} rows. Could not infer kana for {unresolved_count} rows.",
            parent=self,
        )

    def _save(self) -> None:
        try:
            rows = self._parse_entries(
                self.jp_text.get("1.0", "end-1c"),
                self.kana_text.get("1.0", "end-1c") if self._show_kana else "",
                self.en_text.get("1.0", "end-1c"),
            )
            created = self.repository.add_entries(rows)
        except ValidationError as exc:
            messagebox.showerror("Validation error", str(exc), parent=self)
            return
        except sqlite3.Error as exc:
            messagebox.showerror("Database error", f"Could not add entries: {exc}", parent=self)
            return

        self.on_saved()
        messagebox.showinfo("Bulk add complete", f"Added {len(created)} entries.", parent=self)
        self.destroy()


class BulkTagModeDialog(tk.Toplevel):
    def __init__(self, parent: tk.Misc, text_font: tkfont.Font) -> None:
        super().__init__(parent)
        self.result: str | None = None

        self.title("Bulk tag assignment mode")
        self.transient(parent)
        self.grab_set()
        self.resizable(False, False)

        frame = ttk.Frame(self, padding=14)
        frame.grid(row=0, column=0, sticky="nsew")

        ttk.Label(
            frame,
            text="How should selected tags be applied?",
            style="App.TLabel",
        ).grid(row=0, column=0, columnspan=3, sticky="w")

        ttk.Label(
            frame,
            text="Replace: overwrite existing tags. Add: keep existing tags and append new tags.",
            style="Status.TLabel",
            wraplength=420,
            justify="left",
            font=text_font,
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(6, 12))

        ttk.Button(
            frame,
            text="Replace",
            command=lambda: self._close_with_result("replace"),
            style="App.TButton",
        ).grid(row=2, column=0, padx=(0, 8))
        ttk.Button(
            frame,
            text="Add",
            command=lambda: self._close_with_result("add"),
            style="App.TButton",
        ).grid(row=2, column=1, padx=(0, 8))
        ttk.Button(
            frame,
            text="Cancel",
            command=self.destroy,
            style="App.TButton",
        ).grid(row=2, column=2)

        self.bind("<Escape>", lambda _event: self.destroy())

    def _close_with_result(self, value: str) -> None:
        self.result = value
        self.destroy()


class TagSelectionDialog(tk.Toplevel):
    def __init__(
        self,
        parent: tk.Misc,
        repository: VocabRepository,
        target_language_code: str,
        selected_tag_ids: list[int],
        include_part_of_speech: bool,
        title: str,
        text_font: tkfont.Font,
    ) -> None:
        super().__init__(parent)
        self.repository = repository
        self.target_language_code = target_language_code
        self.include_part_of_speech = include_part_of_speech
        self.text_font = text_font
        self.result: list[int] | None = None
        self._selected_tag_ids: set[int] = set(selected_tag_ids)
        self._chip_vars: dict[int, tk.BooleanVar] = {}
        self._chip_buttons: dict[int, tk.Checkbutton] = {}
        self._chip_colors_by_tag_id: dict[int, tuple[str, str]] = {}
        self._tags_by_type: dict[str, list[tuple[int, str]]] = {}
        self._last_layout_width = 0

        self.title(title)
        self.transient(parent)
        self.grab_set()
        self.geometry("920x580")
        self.minsize(680, 400)

        frame = ttk.Frame(self, padding=12)
        frame.grid(row=0, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)

        ttk.Label(
            frame,
            text=f"Select tags ({LANGUAGE_NAMES.get(self.target_language_code, self.target_language_code)})",
            style="App.TLabel",
        ).grid(row=0, column=0, sticky="w", pady=(0, 6))

        list_frame = ttk.Frame(frame)
        list_frame.grid(row=1, column=0, sticky="nsew")
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        self.tag_canvas = tk.Canvas(
            list_frame,
            highlightthickness=1,
            highlightbackground="#d0d0d0",
            background="#ffffff",
        )
        self.tag_canvas.grid(row=0, column=0, sticky="nsew")
        list_scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.tag_canvas.yview)
        list_scroll.grid(row=0, column=1, sticky="ns")
        self.tag_canvas.configure(yscrollcommand=list_scroll.set)

        self.tag_content_frame = tk.Frame(self.tag_canvas, background="#ffffff")
        self._tag_content_window = self.tag_canvas.create_window(
            (0, 0),
            window=self.tag_content_frame,
            anchor="nw",
        )
        self.tag_canvas.bind("<Configure>", self._on_canvas_configure)
        self.tag_content_frame.bind("<Configure>", self._on_content_configure)
        self.tag_canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.tag_canvas.bind("<Button-4>", self._on_mousewheel)
        self.tag_canvas.bind("<Button-5>", self._on_mousewheel)

        actions = ttk.Frame(frame, padding=(0, 10, 0, 0))
        actions.grid(row=2, column=0, sticky="e")
        ttk.Button(actions, text="Clear", command=self._clear_selection, style="App.TButton").grid(
            row=0,
            column=0,
            padx=(0, 8),
        )
        ttk.Button(actions, text="Cancel", command=self.destroy, style="App.TButton").grid(
            row=0,
            column=1,
            padx=(0, 8),
        )
        ttk.Button(actions, text="Apply", command=self._apply, style="App.TButton").grid(row=0, column=2)

        self._load_tags(selected_tag_ids)

        self.bind("<Escape>", lambda _event: self.destroy())
        self.bind("<Return>", lambda _event: self._apply())

    def _on_canvas_configure(self, event: tk.Event) -> None:
        self.tag_canvas.itemconfigure(self._tag_content_window, width=event.width)
        resolved_width = max(int(event.width), 1)
        if abs(resolved_width - self._last_layout_width) >= 20:
            self._last_layout_width = resolved_width
            self._render_tag_sections()

    def _on_content_configure(self, _event: tk.Event) -> None:
        self.tag_canvas.configure(scrollregion=self.tag_canvas.bbox("all"))

    def _on_mousewheel(self, event: tk.Event) -> None:
        if event.num == 5 or event.delta < 0:
            self.tag_canvas.yview_scroll(1, "units")
            return
        if event.num == 4 or event.delta > 0:
            self.tag_canvas.yview_scroll(-1, "units")

    @staticmethod
    def _display_type_name(type_name: str) -> str:
        return type_name.replace("_", " ")

    def _estimate_chip_width(self, label: str) -> int:
        return max(self.text_font.measure(label) + 38, 86)

    def _on_chip_toggled(self, tag_id: int) -> None:
        if tag_id not in self._chip_vars:
            return

        if self._chip_vars[tag_id].get():
            self._selected_tag_ids.add(tag_id)
        else:
            self._selected_tag_ids.discard(tag_id)
        self._refresh_chip_style(tag_id)

    def _refresh_chip_style(self, tag_id: int) -> None:
        button = self._chip_buttons.get(tag_id)
        var = self._chip_vars.get(tag_id)
        if button is None or var is None:
            return

        selected_fill, selected_active_fill = self._chip_colors_by_tag_id.get(
            tag_id,
            ("#e6f2ff", "#d8ebff"),
        )

        if var.get():
            button.configure(
                background=selected_fill,
                activebackground=selected_active_fill,
                fg="#1f2933",
                activeforeground="#1f2933",
            )
        else:
            button.configure(
                background="#e3e3e3",
                activebackground="#d5d5d5",
                fg="#1f2933",
                activeforeground="#1f2933",
            )

    def _render_tag_sections(self) -> None:
        for widget in self.tag_content_frame.winfo_children():
            widget.destroy()

        self._chip_vars.clear()
        self._chip_buttons.clear()
        self._chip_colors_by_tag_id.clear()

        if not self._tags_by_type:
            tk.Label(
                self.tag_content_frame,
                text="No tags available.",
                background="#ffffff",
                anchor="w",
                justify="left",
                font=self.text_font,
            ).pack(anchor="w", padx=6, pady=6)
            return

        heading_font = self.text_font.copy()
        heading_font.configure(weight="bold")

        type_palette: tuple[tuple[str, str], ...] = (
            ("#dff4ff", "#cdeaff"),
            ("#e8f8e1", "#d8efcf"),
            ("#fff1dc", "#ffe5c4"),
            ("#fde8ef", "#f8d8e4"),
            ("#f1e8ff", "#e7d9ff"),
            ("#e7f7f4", "#d7f0eb"),
            ("#f4efdf", "#ece4cd"),
        )

        sorted_type_names = sorted(self._tags_by_type)
        available_width = max(self.tag_canvas.winfo_width() - 34, 340)

        for type_index, type_name in enumerate(sorted_type_names):
            section_frame = tk.Frame(self.tag_content_frame, background="#ffffff")
            section_frame.pack(fill="x", anchor="w", padx=6, pady=(8, 2))

            selected_fill, selected_active_fill = type_palette[type_index % len(type_palette)]

            tk.Label(
                section_frame,
                text=self._display_type_name(type_name),
                background="#ffffff",
                anchor="w",
                justify="left",
                font=heading_font,
            ).pack(anchor="w")

            chips_frame = tk.Frame(section_frame, background="#ffffff")
            chips_frame.pack(fill="x", anchor="w", pady=(4, 0))

            sorted_tags = sorted(self._tags_by_type[type_name], key=lambda item: item[1].lower())
            row = 0
            column = 0
            current_row_width = 0
            for tag_id, tag_name in sorted_tags:
                chip_width = self._estimate_chip_width(tag_name)
                if current_row_width > 0 and current_row_width + chip_width + 6 > available_width:
                    row += 1
                    column = 0
                    current_row_width = 0

                var = tk.BooleanVar(value=tag_id in self._selected_tag_ids)
                self._chip_vars[tag_id] = var

                button = tk.Checkbutton(
                    chips_frame,
                    text=tag_name,
                    variable=var,
                    onvalue=True,
                    offvalue=False,
                    indicatoron=False,
                    borderwidth=1,
                    relief="solid",
                    offrelief="solid",
                    highlightthickness=0,
                    selectcolor=selected_fill,
                    padx=8,
                    pady=3,
                    anchor="center",
                    font=self.text_font,
                    command=lambda current_tag_id=tag_id: self._on_chip_toggled(current_tag_id),
                )
                button.grid(row=row, column=column, sticky="w", padx=(0, 6), pady=(0, 6))
                self._chip_buttons[tag_id] = button
                self._chip_colors_by_tag_id[tag_id] = (selected_fill, selected_active_fill)
                self._refresh_chip_style(tag_id)

                current_row_width += chip_width + 6
                column += 1

            if type_index < len(sorted_type_names) - 1:
                ttk.Separator(self.tag_content_frame, orient="horizontal").pack(fill="x", padx=8, pady=(4, 2))

    def _load_tags(self, selected_tag_ids: list[int]) -> None:
        try:
            tags = self.repository.list_tags(
                target_language_code=self.target_language_code,
                include_part_of_speech=self.include_part_of_speech,
            )
        except sqlite3.Error as exc:
            messagebox.showerror("Database error", f"Could not load tags: {exc}", parent=self)
            self.destroy()
            return

        self._selected_tag_ids = set(selected_tag_ids)

        tags_by_type: dict[str, list[tuple[int, str]]] = {}
        for tag_id, _type_id, type_name, tag_name, _type_predefined, _tag_predefined in tags:
            tags_by_type.setdefault(type_name, []).append((tag_id, tag_name))

        self._tags_by_type = tags_by_type
        self._last_layout_width = max(self.tag_canvas.winfo_width(), 1)
        self._render_tag_sections()

    def _clear_selection(self) -> None:
        self._selected_tag_ids.clear()
        for tag_id, var in self._chip_vars.items():
            var.set(False)
            self._refresh_chip_style(tag_id)

    def _apply(self) -> None:
        self.result = sorted(self._selected_tag_ids)
        self.destroy()


class TagManagerDialog(tk.Toplevel):
    def __init__(
        self,
        parent: tk.Misc,
        repository: VocabRepository,
        target_language_code: str,
        text_font: tkfont.Font,
    ) -> None:
        super().__init__(parent)
        self.repository = repository
        self.target_language_code = target_language_code
        self.text_font = text_font
        self.changed = False

        self._type_rows: list[tuple[int, str, bool]] = []
        self._tag_rows: list[tuple[int, int, str, str, bool, bool]] = []

        self.title(f"Tag manager ({LANGUAGE_NAMES.get(target_language_code, target_language_code)})")
        self.transient(parent)
        self.grab_set()
        self.geometry("860x520")
        self.minsize(700, 420)

        frame = ttk.Frame(self, padding=12)
        frame.grid(row=0, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(1, weight=1)

        ttk.Label(frame, text="Tag types", style="App.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 6))
        ttk.Label(frame, text="Tags", style="App.TLabel").grid(row=0, column=1, sticky="w", pady=(0, 6), padx=(12, 0))

        type_frame = ttk.Frame(frame)
        type_frame.grid(row=1, column=0, sticky="nsew")
        type_frame.columnconfigure(0, weight=1)
        type_frame.rowconfigure(0, weight=1)

        self.type_listbox = tk.Listbox(type_frame, exportselection=False, font=self.text_font)
        self.type_listbox.grid(row=0, column=0, sticky="nsew")
        type_scroll = ttk.Scrollbar(type_frame, orient="vertical", command=self.type_listbox.yview)
        type_scroll.grid(row=0, column=1, sticky="ns")
        self.type_listbox.configure(yscrollcommand=type_scroll.set)
        self.type_listbox.bind("<<ListboxSelect>>", self._on_type_selected)

        tag_frame = ttk.Frame(frame)
        tag_frame.grid(row=1, column=1, sticky="nsew", padx=(12, 0))
        tag_frame.columnconfigure(0, weight=1)
        tag_frame.rowconfigure(0, weight=1)

        self.tag_listbox = tk.Listbox(tag_frame, exportselection=False, font=self.text_font)
        self.tag_listbox.grid(row=0, column=0, sticky="nsew")
        tag_scroll = ttk.Scrollbar(tag_frame, orient="vertical", command=self.tag_listbox.yview)
        tag_scroll.grid(row=0, column=1, sticky="ns")
        self.tag_listbox.configure(yscrollcommand=tag_scroll.set)

        actions = ttk.Frame(frame, padding=(0, 10, 0, 0))
        actions.grid(row=2, column=0, columnspan=2, sticky="ew")

        ttk.Button(actions, text="Add type", command=self._add_type, style="App.TButton").grid(row=0, column=0, padx=(0, 8))
        ttk.Button(actions, text="Delete type", command=self._delete_type, style="App.TButton").grid(row=0, column=1, padx=(0, 16))
        ttk.Button(actions, text="Add tag", command=self._add_tag, style="App.TButton").grid(row=0, column=2, padx=(0, 8))
        ttk.Button(actions, text="Delete tag", command=self._delete_tag, style="App.TButton").grid(row=0, column=3, padx=(0, 16))
        ttk.Button(actions, text="Close", command=self.destroy, style="App.TButton").grid(row=0, column=4)

        self._refresh_type_list(select_first=True)

        self.bind("<Escape>", lambda _event: self.destroy())

    def _selected_type(self) -> tuple[int, str, bool] | None:
        selection = self.type_listbox.curselection()
        if not selection:
            return None
        index = selection[0]
        if index >= len(self._type_rows):
            return None
        return self._type_rows[index]

    def _selected_tag(self) -> tuple[int, int, str, str, bool, bool] | None:
        selection = self.tag_listbox.curselection()
        if not selection:
            return None
        index = selection[0]
        if index >= len(self._tag_rows):
            return None
        return self._tag_rows[index]

    def _refresh_type_list(self, select_first: bool = False) -> None:
        rows = self.repository.list_tag_types(target_language_code=self.target_language_code)
        if self.target_language_code != "JP":
            rows = [
                row
                for row in rows
                if row[1].lower() != "difficulty"
            ]
        self._type_rows = rows
        self.type_listbox.delete(0, "end")
        for _type_id, type_name, is_predefined in self._type_rows:
            suffix = " (predefined)" if is_predefined else ""
            self.type_listbox.insert("end", f"{type_name}{suffix}")

        if self._type_rows:
            if select_first:
                self.type_listbox.selection_set(0)
            elif not self.type_listbox.curselection():
                self.type_listbox.selection_set(0)
        self._refresh_tag_list()

    def _refresh_tag_list(self) -> None:
        selected_type = self._selected_type()
        self.tag_listbox.delete(0, "end")
        self._tag_rows = []
        if selected_type is None:
            return

        type_id, _type_name, _is_predefined = selected_type
        self._tag_rows = self.repository.list_tags(target_language_code=self.target_language_code, tag_type_id=type_id)
        for _tag_id, _type_id, _type_name, tag_name, _type_predefined, tag_predefined in self._tag_rows:
            suffix = " (predefined)" if tag_predefined else ""
            self.tag_listbox.insert("end", f"{tag_name}{suffix}")

    def _on_type_selected(self, _event: tk.Event) -> None:
        self._refresh_tag_list()

    def _add_type(self) -> None:
        entered = simpledialog.askstring("Add tag type", "Type name:", parent=self)
        if entered is None:
            return
        if self.target_language_code != "JP" and entered.strip().lower() == "difficulty":
            messagebox.showerror(
                "Validation error",
                "Difficulty tags are only available for JP workbooks.",
                parent=self,
            )
            return
        try:
            self.repository.add_tag_type(entered, target_language_code=self.target_language_code)
            self.changed = True
        except ValidationError as exc:
            messagebox.showerror("Validation error", str(exc), parent=self)
            return
        except sqlite3.Error as exc:
            messagebox.showerror("Database error", f"Could not add tag type: {exc}", parent=self)
            return

        self._refresh_type_list()

    def _delete_type(self) -> None:
        selected_type = self._selected_type()
        if selected_type is None:
            messagebox.showwarning("No selection", "Select a type to delete.", parent=self)
            return

        type_id, type_name, _is_predefined = selected_type
        if not messagebox.askyesno("Delete type", f"Delete type '{type_name}' and all of its tags?", parent=self):
            return

        try:
            self.repository.delete_tag_type(type_id)
            self.changed = True
        except ValueError as exc:
            messagebox.showerror("Not allowed", str(exc), parent=self)
            return
        except LookupError as exc:
            messagebox.showerror("Not found", str(exc), parent=self)
            return
        except sqlite3.Error as exc:
            messagebox.showerror("Database error", f"Could not delete type: {exc}", parent=self)
            return

        self._refresh_type_list(select_first=True)

    def _add_tag(self) -> None:
        selected_type = self._selected_type()
        if selected_type is None:
            messagebox.showwarning("No selection", "Select a type first.", parent=self)
            return

        type_id, type_name, _is_predefined = selected_type
        entered = simpledialog.askstring("Add tag", f"Tag name for '{type_name}':", parent=self)
        if entered is None:
            return

        try:
            self.repository.add_tag(type_id, entered)
            self.changed = True
        except ValidationError as exc:
            messagebox.showerror("Validation error", str(exc), parent=self)
            return
        except LookupError as exc:
            messagebox.showerror("Not found", str(exc), parent=self)
            return
        except sqlite3.Error as exc:
            messagebox.showerror("Database error", f"Could not add tag: {exc}", parent=self)
            return

        self._refresh_tag_list()

    def _delete_tag(self) -> None:
        selected_tag = self._selected_tag()
        if selected_tag is None:
            messagebox.showwarning("No selection", "Select a tag to delete.", parent=self)
            return

        tag_id, _type_id, _type_name, tag_name, _type_predefined, _tag_predefined = selected_tag
        if not messagebox.askyesno("Delete tag", f"Delete tag '{tag_name}'?", parent=self):
            return

        try:
            self.repository.delete_tag(tag_id)
            self.changed = True
        except ValueError as exc:
            messagebox.showerror("Not allowed", str(exc), parent=self)
            return
        except LookupError as exc:
            messagebox.showerror("Not found", str(exc), parent=self)
            return
        except sqlite3.Error as exc:
            messagebox.showerror("Database error", f"Could not delete tag: {exc}", parent=self)
            return

        self._refresh_tag_list()


class EntryDialog(tk.Toplevel):
    def __init__(
        self,
        parent: tk.Tk,
        repository: VocabRepository,
        title: str,
        save_button_text: str,
        save_handler: Callable[[str, str, str, str, list[int]], object],
        on_saved: Callable[[], None],
        initial_japanese: str,
        initial_kana: str,
        initial_english: str,
        initial_part_of_speech: str,
        initial_tag_ids: list[int],
        target_label: str,
        assistant_label: str,
        target_language_code: str,
        enable_kana_suggest: bool,
    ) -> None:
        super().__init__(parent)
        self.repository = repository
        self.save_handler = save_handler
        self.on_saved = on_saved
        self.target_label = target_label
        self.assistant_label = assistant_label
        self.target_language_code = target_language_code
        self.enable_kana_suggest = enable_kana_suggest
        self.selected_tag_ids = list(initial_tag_ids)

        self._auto_suggest_job: str | None = None
        self._updating_kana = False
        self._kana_user_override = bool(initial_kana.strip())

        self.japanese_var = tk.StringVar(value=initial_japanese)
        self.kana_var = tk.StringVar(value=initial_kana)
        self.english_var = tk.StringVar(value=initial_english)
        self.part_of_speech_var = tk.StringVar(value=initial_part_of_speech)
        self.tags_summary_var = tk.StringVar(value="")
        default_status = "Kana is optional. You can edit any suggestion."
        if not self.enable_kana_suggest:
            default_status = "Kana is optional. Automatic kana suggestion is available only when the workbook preset enables kana."
        self.status_var = tk.StringVar(value=default_status)

        self.title(title)
        self.transient(parent)
        self.grab_set()
        self.resizable(False, False)

        self._build_widgets(save_button_text)

        self.japanese_var.trace_add("write", self._on_japanese_text_change)
        self.kana_var.trace_add("write", self._on_kana_text_change)

        self.bind("<Return>", lambda _event: self._save())
        self.bind("<Escape>", lambda _event: self.destroy())

        self.japanese_entry.focus_set()
        self._refresh_tag_summary()

    def _build_widgets(self, save_button_text: str) -> None:
        frame = ttk.Frame(self, padding=14)
        frame.grid(row=0, column=0, sticky="nsew")

        ttk.Label(frame, text=f"{self.target_label} *", style="App.TLabel").grid(row=0, column=0, sticky="w")
        self.japanese_entry = ttk.Entry(frame, textvariable=self.japanese_var, width=48, style="Japanese.TEntry")
        self.japanese_entry.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 10))

        if self.enable_kana_suggest:
            ttk.Label(frame, text="Kana (optional)", style="App.TLabel").grid(row=2, column=0, sticky="w")
            self.kana_entry = ttk.Entry(frame, textvariable=self.kana_var, width=48, style="Japanese.TEntry")
            self.kana_entry.grid(row=3, column=0, sticky="ew", pady=(0, 10))

            suggest_button = ttk.Button(frame, text="Suggest kana", command=self._suggest_kana_manual, style="App.TButton")
            suggest_button.grid(row=3, column=1, sticky="e", padx=(8, 0), pady=(0, 10))
            if not self.enable_kana_suggest:
                suggest_button.configure(state="disabled")

        ttk.Label(frame, text=f"{self.assistant_label} *", style="App.TLabel").grid(row=4, column=0, sticky="w")
        self.english_entry = ttk.Entry(frame, textvariable=self.english_var, width=48, style="App.TEntry")
        self.english_entry.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(0, 8))

        ttk.Label(frame, text="Part of speech (optional)", style="App.TLabel").grid(row=6, column=0, sticky="w")
        self.part_of_speech_combo = ttk.Combobox(
            frame,
            values=PART_OF_SPEECH_OPTIONS,
            state="normal",
            width=46,
            textvariable=self.part_of_speech_var,
        )
        self.part_of_speech_combo.grid(row=7, column=0, columnspan=2, sticky="ew", pady=(0, 8))

        ttk.Label(frame, text="Tags (optional)", style="App.TLabel").grid(row=8, column=0, sticky="w")
        tag_actions = ttk.Frame(frame)
        tag_actions.grid(row=9, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        tag_actions.columnconfigure(1, weight=1)
        ttk.Button(tag_actions, text="Select tags", command=self._open_tag_selector, style="App.TButton").grid(
            row=0,
            column=0,
            sticky="w",
        )
        ttk.Label(tag_actions, textvariable=self.tags_summary_var, style="Status.TLabel").grid(
            row=0,
            column=1,
            sticky="w",
            padx=(8, 0),
        )

        status_label = ttk.Label(frame, textvariable=self.status_var, style="Status.TLabel", wraplength=460)
        status_label.grid(row=10, column=0, columnspan=2, sticky="w", pady=(2, 12))

        actions = ttk.Frame(frame)
        actions.grid(row=11, column=0, columnspan=2, sticky="e")

        cancel_button = ttk.Button(actions, text="Cancel", command=self.destroy, style="App.TButton")
        save_button = ttk.Button(actions, text=save_button_text, command=self._save, style="App.TButton")
        cancel_button.grid(row=0, column=0, padx=(0, 8))
        save_button.grid(row=0, column=1)

    def _on_japanese_text_change(self, *_: object) -> None:
        if not self.enable_kana_suggest:
            return
        if self._auto_suggest_job is not None:
            self.after_cancel(self._auto_suggest_job)
        self._auto_suggest_job = self.after(300, self._suggest_kana_automatic)

    def _on_kana_text_change(self, *_: object) -> None:
        if self._updating_kana:
            return
        self._kana_user_override = bool(self.kana_var.get().strip())

    def _suggest_kana_automatic(self) -> None:
        if not self.enable_kana_suggest:
            return
        self._auto_suggest_job = None
        if self._kana_user_override and self.kana_var.get().strip():
            return
        self._suggest_kana()

    def _suggest_kana_manual(self) -> None:
        if not self.enable_kana_suggest:
            self.status_var.set("Kana suggestion is disabled for the current target language.")
            return
        self._suggest_kana(force_message=True)

    def _suggest_kana(self, force_message: bool = False) -> None:
        suggestion, reliable, message = suggest_hiragana(self.japanese_var.get())

        if reliable and suggestion:
            self._updating_kana = True
            self.kana_var.set(suggestion)
            self._updating_kana = False
            self._kana_user_override = False
            self.status_var.set(message)
            return

        if force_message or not self.kana_var.get().strip():
            self.status_var.set(message)

    def _open_tag_selector(self) -> None:
        dialog = TagSelectionDialog(
            self,
            repository=self.repository,
            target_language_code=self.target_language_code,
            selected_tag_ids=self.selected_tag_ids,
            include_part_of_speech=False,
            title="Select tags",
            text_font=tkfont.nametofont("TkDefaultFont"),
        )
        self.wait_window(dialog)
        if dialog.result is None:
            return

        self.selected_tag_ids = list(dialog.result)
        self._refresh_tag_summary()

    def _refresh_tag_summary(self) -> None:
        if not self.selected_tag_ids:
            self.tags_summary_var.set("No tags selected")
            return

        try:
            tags = self.repository.list_tags(
                target_language_code=self.target_language_code,
                include_part_of_speech=False,
            )
        except sqlite3.Error:
            self.tags_summary_var.set(f"{len(self.selected_tag_ids)} tags selected")
            return

        label_by_id = {
            tag_id: f"{type_name}:{tag_name}"
            for tag_id, _type_id, type_name, tag_name, _type_predefined, _tag_predefined in tags
        }
        selected_labels = [label_by_id[tag_id] for tag_id in self.selected_tag_ids if tag_id in label_by_id]
        if not selected_labels:
            self.tags_summary_var.set(f"{len(self.selected_tag_ids)} tags selected")
            return

        summary = ", ".join(selected_labels[:2])
        if len(selected_labels) > 2:
            summary += f", +{len(selected_labels) - 2} more"
        self.tags_summary_var.set(summary)

    def _save(self) -> None:
        try:
            self.save_handler(
                self.japanese_var.get(),
                self.kana_var.get(),
                self.english_var.get(),
                self.part_of_speech_var.get(),
                list(self.selected_tag_ids),
            )
        except ValidationError as exc:
            messagebox.showerror("Validation error", str(exc), parent=self)
            return
        except LookupError as exc:
            messagebox.showerror("Not found", str(exc), parent=self)
            return
        except sqlite3.Error as exc:
            messagebox.showerror("Database error", f"Could not save entry: {exc}", parent=self)
            return

        self.on_saved()
        self.destroy()


class WorkbookCreationDialog(tk.Toplevel):
    def __init__(self, parent: tk.Tk, text_font: tkfont.Font) -> None:
        super().__init__(parent)
        self.result: tuple[str, str, str, str, str] | None = None

        self.name_var = tk.StringVar(value="")
        self.target_mode_var = tk.StringVar(value="supported")
        self.target_language_var = tk.StringVar(value="JP")
        self.target_custom_label_var = tk.StringVar(value="")

        self.meaning_mode_var = tk.StringVar(value="default")
        self.meaning_language_var = tk.StringVar(value="EN")
        self.meaning_custom_label_var = tk.StringVar(value="")

        self.enable_preset_var = tk.BooleanVar(value=True)
        self.preset_info_var = tk.StringVar(
            value="Japanese preset adds kana support and predefined JP difficulty tags."
        )

        self.title("Create workbook")
        self.transient(parent)
        self.grab_set()
        self.resizable(False, False)

        frame = ttk.Frame(self, padding=14)
        frame.grid(row=0, column=0, sticky="nsew")
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="Workbook name *", style="App.TLabel").grid(row=0, column=0, sticky="w")
        self.name_entry = ttk.Entry(frame, textvariable=self.name_var, style="App.TEntry", width=32)
        self.name_entry.grid(row=0, column=1, sticky="ew", padx=(8, 0), pady=(0, 8))

        ttk.Label(frame, text="Target label source", style="App.TLabel").grid(row=1, column=0, sticky="w")
        target_mode_row = ttk.Frame(frame)
        target_mode_row.grid(row=1, column=1, sticky="w", padx=(8, 0), pady=(0, 8))
        ttk.Radiobutton(
            target_mode_row,
            text="Supported language",
            value="supported",
            variable=self.target_mode_var,
            command=self._refresh_mode_rows,
        ).grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(
            target_mode_row,
            text="Custom label",
            value="custom",
            variable=self.target_mode_var,
            command=self._refresh_mode_rows,
        ).grid(row=0, column=1, sticky="w", padx=(8, 0))

        self.target_supported_row = ttk.Frame(frame)
        self.target_supported_row.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        ttk.Label(self.target_supported_row, text="Target language", style="App.TLabel").grid(row=0, column=0, sticky="w")
        self.target_combo = ttk.Combobox(
            self.target_supported_row,
            state="readonly",
            values=("JP", "EN"),
            textvariable=self.target_language_var,
            width=8,
            font=text_font,
        )
        self.target_combo.grid(row=0, column=1, sticky="w", padx=(8, 0))
        self.target_combo.bind("<<ComboboxSelected>>", lambda _event: self._refresh_mode_rows())

        self.target_custom_row = ttk.Frame(frame)
        self.target_custom_row.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        self.target_custom_row.columnconfigure(1, weight=1)
        ttk.Label(self.target_custom_row, text="Target label *", style="App.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Entry(
            self.target_custom_row,
            textvariable=self.target_custom_label_var,
            style="App.TEntry",
            width=28,
        ).grid(row=0, column=1, sticky="ew", padx=(8, 0))

        ttk.Separator(frame, orient="horizontal").grid(row=4, column=0, columnspan=2, sticky="ew", pady=(2, 10))

        ttk.Label(frame, text="Meaning label", style="App.TLabel").grid(row=5, column=0, sticky="w")
        meaning_mode_row = ttk.Frame(frame)
        meaning_mode_row.grid(row=5, column=1, sticky="w", padx=(8, 0), pady=(0, 8))
        ttk.Radiobutton(
            meaning_mode_row,
            text="Default",
            value="default",
            variable=self.meaning_mode_var,
            command=self._refresh_mode_rows,
        ).grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(
            meaning_mode_row,
            text="Supported language",
            value="supported",
            variable=self.meaning_mode_var,
            command=self._refresh_mode_rows,
        ).grid(row=0, column=1, sticky="w", padx=(8, 0))
        ttk.Radiobutton(
            meaning_mode_row,
            text="Custom label",
            value="custom",
            variable=self.meaning_mode_var,
            command=self._refresh_mode_rows,
        ).grid(row=0, column=2, sticky="w", padx=(8, 0))

        self.meaning_supported_row = ttk.Frame(frame)
        self.meaning_supported_row.grid(row=6, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        ttk.Label(self.meaning_supported_row, text="Meaning language", style="App.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Combobox(
            self.meaning_supported_row,
            state="readonly",
            values=("JP", "EN"),
            textvariable=self.meaning_language_var,
            width=8,
            font=text_font,
        ).grid(row=0, column=1, sticky="w", padx=(8, 0))

        self.meaning_custom_row = ttk.Frame(frame)
        self.meaning_custom_row.grid(row=7, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        self.meaning_custom_row.columnconfigure(1, weight=1)
        ttk.Label(self.meaning_custom_row, text="Meaning label *", style="App.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Entry(
            self.meaning_custom_row,
            textvariable=self.meaning_custom_label_var,
            style="App.TEntry",
            width=28,
        ).grid(row=0, column=1, sticky="ew", padx=(8, 0))

        self.preset_section = ttk.Frame(frame)
        self.preset_section.grid(row=8, column=0, columnspan=2, sticky="ew", pady=(0, 12))
        ttk.Checkbutton(
            self.preset_section,
            text="Enable preset for selected target language",
            variable=self.enable_preset_var,
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            self.preset_section,
            textvariable=self.preset_info_var,
            style="Status.TLabel",
            wraplength=430,
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))

        actions = ttk.Frame(frame)
        actions.grid(row=9, column=0, columnspan=2, sticky="e")
        ttk.Button(actions, text="Cancel", command=self.destroy, style="App.TButton").grid(
            row=0,
            column=0,
            padx=(0, 8),
        )
        ttk.Button(actions, text="Create", command=self._save, style="App.TButton").grid(row=0, column=1)

        self.bind("<Escape>", lambda _event: self.destroy())
        self.bind("<Return>", lambda _event: self._save())
        self.name_entry.focus_set()

        self._refresh_mode_rows()

    def _target_has_available_preset(self) -> bool:
        return self.target_mode_var.get() == "supported" and self.target_language_var.get().strip().upper() == "JP"

    def _refresh_mode_rows(self) -> None:
        if self.target_mode_var.get() == "supported":
            self.target_supported_row.grid()
            self.target_custom_row.grid_remove()
        else:
            self.target_supported_row.grid_remove()
            self.target_custom_row.grid()

        meaning_mode = self.meaning_mode_var.get()
        if meaning_mode == "supported":
            self.meaning_supported_row.grid()
            self.meaning_custom_row.grid_remove()
        elif meaning_mode == "custom":
            self.meaning_supported_row.grid_remove()
            self.meaning_custom_row.grid()
        else:
            self.meaning_supported_row.grid_remove()
            self.meaning_custom_row.grid_remove()

        if self._target_has_available_preset():
            self.preset_section.grid()
            self.preset_info_var.set("Japanese preset adds kana support and predefined JP difficulty tags.")
        else:
            self.preset_section.grid_remove()

    def _save(self) -> None:
        workbook_name = self.name_var.get().strip()
        if not workbook_name:
            messagebox.showerror("Validation error", "Workbook name is required.", parent=self)
            return

        if self.target_mode_var.get() == "supported":
            selected_target_language = self.target_language_var.get().strip().upper()
            if selected_target_language not in LANGUAGE_NAMES:
                messagebox.showerror("Validation error", "Select a supported target language.", parent=self)
                return

            target_label = LANGUAGE_NAMES[selected_target_language]
            target_schema_code = selected_target_language
            if selected_target_language == "JP" and not self.enable_preset_var.get():
                target_schema_code = "JP_GENERIC"
            preset_key = "japanese" if self._target_has_available_preset() and self.enable_preset_var.get() else "generic"
        else:
            target_label = self.target_custom_label_var.get().strip()
            if not target_label:
                messagebox.showerror("Validation error", "Target label is required.", parent=self)
                return
            custom_hash = hashlib.sha1(target_label.encode("utf-8")).hexdigest()[:10].upper()
            target_schema_code = f"CUSTOM_{custom_hash}"
            preset_key = "generic"

        meaning_mode = self.meaning_mode_var.get()
        if meaning_mode == "supported":
            selected_meaning_language = self.meaning_language_var.get().strip().upper()
            if selected_meaning_language not in LANGUAGE_NAMES:
                messagebox.showerror("Validation error", "Select a supported meaning language.", parent=self)
                return
            meaning_label = LANGUAGE_NAMES[selected_meaning_language]
        elif meaning_mode == "custom":
            meaning_label = self.meaning_custom_label_var.get().strip()
            if not meaning_label:
                messagebox.showerror("Validation error", "Meaning label is required.", parent=self)
                return
        else:
            meaning_label = "Meaning"

        self.result = (workbook_name, target_schema_code, preset_key, target_label, meaning_label)
        self.destroy()


class WorkbookLabelsDialog(tk.Toplevel):
    def __init__(self, parent: tk.Tk, workbook: Workbook, text_font: tkfont.Font) -> None:
        super().__init__(parent)
        self.result: tuple[str, str] | None = None

        self.target_label_var = tk.StringVar(value=workbook.target_label)
        self.meaning_label_var = tk.StringVar(value=workbook.meaning_label)

        self.title("Edit workbook labels")
        self.transient(parent)
        self.grab_set()
        self.resizable(False, False)

        frame = ttk.Frame(self, padding=14)
        frame.grid(row=0, column=0, sticky="nsew")
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="Workbook", style="Status.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(frame, text=workbook.name, style="App.TLabel").grid(row=0, column=1, sticky="w", padx=(8, 0), pady=(0, 8))

        ttk.Label(frame, text="Target label *", style="App.TLabel").grid(row=1, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.target_label_var, style="App.TEntry", width=32, font=text_font).grid(
            row=1,
            column=1,
            sticky="ew",
            padx=(8, 0),
            pady=(0, 8),
        )

        ttk.Label(frame, text="Meaning label", style="App.TLabel").grid(row=2, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.meaning_label_var, style="App.TEntry", width=32, font=text_font).grid(
            row=2,
            column=1,
            sticky="ew",
            padx=(8, 0),
            pady=(0, 10),
        )

        ttk.Label(
            frame,
            text="Leave meaning label empty to use the default 'Meaning'.",
            style="Status.TLabel",
            wraplength=420,
            justify="left",
        ).grid(row=3, column=0, columnspan=2, sticky="w", pady=(0, 12))

        actions = ttk.Frame(frame)
        actions.grid(row=4, column=0, columnspan=2, sticky="e")
        ttk.Button(actions, text="Cancel", command=self.destroy, style="App.TButton").grid(row=0, column=0, padx=(0, 8))
        ttk.Button(actions, text="Save", command=self._save, style="App.TButton").grid(row=0, column=1)

        self.bind("<Escape>", lambda _event: self.destroy())
        self.bind("<Return>", lambda _event: self._save())

    def _save(self) -> None:
        target_label = self.target_label_var.get().strip()
        if not target_label:
            messagebox.showerror("Validation error", "Target label is required.", parent=self)
            return

        meaning_label = self.meaning_label_var.get().strip()
        if not meaning_label:
            meaning_label = "Meaning"

        self.result = (target_label, meaning_label)
        self.destroy()


class LanguageSettingsDialog(tk.Toplevel):
    def __init__(
        self,
        parent: tk.Tk,
        target_language: str,
        assistant_language: str,
        text_font: tkfont.Font,
    ) -> None:
        super().__init__(parent)
        self.result: tuple[str, str] | None = None

        self.target_var = tk.StringVar(value=target_language)
        self.assistant_var = tk.StringVar(value=assistant_language)

        self.title("Language settings")
        self.transient(parent)
        self.grab_set()
        self.resizable(False, False)

        frame = ttk.Frame(self, padding=14)
        frame.grid(row=0, column=0, sticky="nsew")

        ttk.Label(frame, text="Target language", style="App.TLabel").grid(row=0, column=0, sticky="w")
        target_combo = ttk.Combobox(
            frame,
            values=tuple(LANGUAGE_NAMES.keys()),
            state="readonly",
            width=10,
            textvariable=self.target_var,
            font=text_font,
        )
        target_combo.grid(row=1, column=0, sticky="w", pady=(4, 10))

        ttk.Label(frame, text="Assistant language", style="App.TLabel").grid(row=2, column=0, sticky="w")
        assistant_combo = ttk.Combobox(
            frame,
            values=tuple(LANGUAGE_NAMES.keys()),
            state="readonly",
            width=10,
            textvariable=self.assistant_var,
            font=text_font,
        )
        assistant_combo.grid(row=3, column=0, sticky="w", pady=(4, 12))

        actions = ttk.Frame(frame)
        actions.grid(row=4, column=0, sticky="e")

        ttk.Button(actions, text="Cancel", command=self.destroy, style="App.TButton").grid(
            row=0,
            column=0,
            padx=(0, 8),
        )
        ttk.Button(actions, text="Save", command=self._save, style="App.TButton").grid(row=0, column=1)

        self.bind("<Escape>", lambda _event: self.destroy())
        self.bind("<Return>", lambda _event: self._save())

    def _save(self) -> None:
        target = self.target_var.get().strip().upper()
        assistant = self.assistant_var.get().strip().upper()
        if target == assistant:
            messagebox.showerror("Validation error", "Target and assistant languages must be different.", parent=self)
            return

        self.result = (target, assistant)
        self.destroy()


class VocabularyDetailDialog(tk.Toplevel):
    def __init__(
        self,
        parent: tk.Tk,
        repository: VocabRepository,
        entry_id: int,
        text_fonts: dict[str, tkfont.Font],
        target_label: str,
        assistant_label: str,
        target_language_code: str,
        enable_kana_suggest: bool,
        on_saved: Callable[[], None],
    ) -> None:
        super().__init__(parent)
        self.repository = repository
        self.entry_id = entry_id
        self.text_fonts = text_fonts
        self.target_label = target_label
        self.assistant_label = assistant_label
        self.target_language_code = target_language_code
        self.enable_kana_suggest = enable_kana_suggest
        self.on_saved = on_saved

        self._auto_suggest_job: str | None = None
        self._updating_kana = False
        self._kana_user_override = False
        self._is_editing_markdown = False

        self.target_var = tk.StringVar(value="")
        self.kana_var = tk.StringVar(value="")
        self.assistant_var = tk.StringVar(value="")
        self.stats_var = tk.StringVar(value="")
        self.created_var = tk.StringVar(value="")
        self.latest_practice_var = tk.StringVar(value="")
        self.details_markdown = ""
        self.selected_tag_ids: list[int] = []
        self._tag_chip_labels: list[tk.Label] = []
        self._tags_chip_area: tk.Frame | None = None

        self.title("Vocabulary details")
        self.transient(parent)
        self.grab_set()
        self.geometry("900x680")
        self.minsize(760, 560)

        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        self._build_widgets()
        self._load_entry()

        self.target_var.trace_add("write", self._on_target_text_change)
        self.kana_var.trace_add("write", self._on_kana_text_change)

        self.bind("<Escape>", lambda _event: self.destroy())

    def _build_widgets(self) -> None:
        frame = ttk.Frame(self, padding=14)
        frame.grid(row=0, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)

        header = ttk.LabelFrame(frame, text="Vocabulary", padding=10)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        header.columnconfigure(1, weight=1)

        ttk.Label(header, text=f"{self.target_label} *", style="App.TLabel").grid(row=0, column=0, sticky="w")
        self.target_entry = ttk.Entry(header, textvariable=self.target_var, style="Japanese.TEntry")
        self.target_entry.grid(row=0, column=1, sticky="ew", padx=(8, 0), pady=(0, 6))

        ttk.Label(header, text="Kana (optional)", style="App.TLabel").grid(row=1, column=0, sticky="w")
        self.kana_entry = ttk.Entry(header, textvariable=self.kana_var, style="Japanese.TEntry")
        self.kana_entry.grid(row=1, column=1, sticky="ew", padx=(8, 0), pady=(0, 6))

        suggest_kana_button = ttk.Button(
            header,
            text="Suggest kana",
            command=self._suggest_kana_manual,
            style="App.TButton",
        )
        suggest_kana_button.grid(row=1, column=2, sticky="w", padx=(8, 0), pady=(0, 6))
        if not self.enable_kana_suggest:
            suggest_kana_button.configure(state="disabled")

        ttk.Label(header, text=f"{self.assistant_label} *", style="App.TLabel").grid(row=2, column=0, sticky="w")
        self.assistant_entry = ttk.Entry(header, textvariable=self.assistant_var, style="App.TEntry")
        self.assistant_entry.grid(row=2, column=1, sticky="ew", padx=(8, 0), pady=(0, 6))

        ttk.Label(header, text="Tags", style="App.TLabel").grid(row=3, column=0, sticky="w")
        tag_actions = ttk.Frame(header)
        tag_actions.grid(row=3, column=1, columnspan=2, sticky="ew", padx=(8, 0), pady=(0, 6))
        tag_actions.columnconfigure(1, weight=1)
        ttk.Button(tag_actions, text="Select tags", command=self._open_tag_selector, style="App.TButton").grid(
            row=0,
            column=0,
            sticky="w",
        )
        self._tags_chip_area = tk.Frame(tag_actions, background="#ffffff", highlightthickness=0, borderwidth=0)
        self._tags_chip_area.grid(row=0, column=1, sticky="w", padx=(8, 0))

        ttk.Label(header, textvariable=self.stats_var, style="Status.TLabel").grid(
            row=4,
            column=0,
            columnspan=3,
            sticky="w",
            pady=(4, 0),
        )
        ttk.Label(header, textvariable=self.created_var, style="Status.TLabel").grid(
            row=5,
            column=0,
            columnspan=3,
            sticky="w",
        )
        ttk.Label(header, textvariable=self.latest_practice_var, style="Status.TLabel").grid(
            row=6,
            column=0,
            columnspan=3,
            sticky="w",
        )

        details_section = ttk.LabelFrame(frame, text="Details (Markdown)", padding=10)
        details_section.grid(row=1, column=0, sticky="nsew")
        details_section.columnconfigure(0, weight=1)
        details_section.rowconfigure(1, weight=1)

        details_actions = ttk.Frame(details_section)
        details_actions.grid(row=0, column=0, sticky="e", pady=(0, 8))
        self.toggle_details_button = ttk.Button(
            details_actions,
            text="Edit markdown",
            command=self._toggle_markdown_mode,
            style="App.TButton",
        )
        self.toggle_details_button.grid(row=0, column=0)

        preview_frame = ttk.Frame(details_section)
        preview_frame.grid(row=1, column=0, sticky="nsew")
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)

        self.details_preview = tk.Text(
            preview_frame,
            wrap="word",
            font=self.text_fonts["latin"],
            state="disabled",
        )
        self.details_preview.grid(row=0, column=0, sticky="nsew")
        preview_scroll = ttk.Scrollbar(preview_frame, orient="vertical", command=self.details_preview.yview)
        preview_scroll.grid(row=0, column=1, sticky="ns")
        self.details_preview.configure(yscrollcommand=preview_scroll.set)

        editor_frame = ttk.Frame(details_section)
        editor_frame.grid(row=1, column=0, sticky="nsew")
        editor_frame.columnconfigure(0, weight=1)
        editor_frame.rowconfigure(0, weight=1)

        self.details_editor = tk.Text(editor_frame, wrap="word", font=self.text_fonts["latin"])
        self.details_editor.grid(row=0, column=0, sticky="nsew")
        editor_scroll = ttk.Scrollbar(editor_frame, orient="vertical", command=self.details_editor.yview)
        editor_scroll.grid(row=0, column=1, sticky="ns")
        self.details_editor.configure(yscrollcommand=editor_scroll.set)
        editor_frame.grid_remove()

        self._details_preview_frame = preview_frame
        self._details_editor_frame = editor_frame

        footer = ttk.Frame(frame, padding=(0, 10, 0, 0))
        footer.grid(row=2, column=0, sticky="e")

        ttk.Button(footer, text="Cancel", command=self.destroy, style="App.TButton").grid(
            row=0,
            column=0,
            padx=(0, 8),
        )
        ttk.Button(footer, text="Save", command=self._save, style="App.TButton").grid(row=0, column=1)

    def _load_entry(self) -> None:
        try:
            entry = self.repository.get_entry(self.entry_id)
            test_count, error_count, tier = self.repository.get_entry_stats(self.entry_id)
            latest_practiced = self.repository.get_entry_last_practiced(self.entry_id)
        except LookupError as exc:
            messagebox.showerror("Not found", str(exc), parent=self)
            self.destroy()
            return

        self.target_var.set(entry.japanese_text)
        self.kana_var.set(entry.kana_text or "")
        self.assistant_var.set(entry.english_text)
        self.details_markdown = entry.details_markdown or ""

        try:
            tag_rows = self.repository.get_entry_tags(
                self.entry_id,
                target_language_code=self.target_language_code,
                include_part_of_speech=True,
            )
        except sqlite3.Error:
            tag_rows = []
        self.selected_tag_ids = [tag_id for tag_id, _type_id, _type_name, _tag_name in tag_rows]
        self._refresh_tag_summary()

        self.stats_var.set(f"Tests: {test_count} | Errors: {error_count} | Tier: {tier}")
        self.created_var.set(f"Created at: {entry.created_at}")
        self.latest_practice_var.set(f"Latest practice: {latest_practiced or 'Never'}")

        self.details_editor.delete("1.0", "end")
        self.details_editor.insert("1.0", self.details_markdown)
        _render_markdown_to_text(
            self.details_preview,
            self.details_markdown,
            self.text_fonts["latin"],
            tkfont.Font(self, family="Consolas", size=BASE_FONT_SIZE),
        )

    def _toggle_markdown_mode(self) -> None:
        self._is_editing_markdown = not self._is_editing_markdown
        if self._is_editing_markdown:
            self._details_preview_frame.grid_remove()
            self._details_editor_frame.grid()
            self.toggle_details_button.configure(text="Done editing")
            self.details_editor.focus_set()
            return

        self.details_markdown = self.details_editor.get("1.0", "end-1c")
        _render_markdown_to_text(
            self.details_preview,
            self.details_markdown,
            self.text_fonts["latin"],
            tkfont.Font(self, family="Consolas", size=BASE_FONT_SIZE),
        )
        self._details_editor_frame.grid_remove()
        self._details_preview_frame.grid()
        self.toggle_details_button.configure(text="Edit markdown")

    def _on_target_text_change(self, *_: object) -> None:
        if not self.enable_kana_suggest:
            return
        if self._auto_suggest_job is not None:
            self.after_cancel(self._auto_suggest_job)
        self._auto_suggest_job = self.after(300, self._suggest_kana_automatic)

    def _on_kana_text_change(self, *_: object) -> None:
        if self._updating_kana:
            return
        self._kana_user_override = bool(self.kana_var.get().strip())

    def _suggest_kana_automatic(self) -> None:
        self._auto_suggest_job = None
        if not self.enable_kana_suggest:
            return
        if self._kana_user_override and self.kana_var.get().strip():
            return
        self._suggest_kana()

    def _suggest_kana_manual(self) -> None:
        if not self.enable_kana_suggest:
            return
        self._suggest_kana(force_message=True)

    def _suggest_kana(self, force_message: bool = False) -> None:
        suggestion, reliable, message = suggest_hiragana(self.target_var.get())

        if reliable and suggestion:
            self._updating_kana = True
            self.kana_var.set(suggestion)
            self._updating_kana = False
            self._kana_user_override = False
            return

        if force_message:
            messagebox.showinfo("Kana suggestion", message, parent=self)

    def _open_tag_selector(self) -> None:
        dialog = TagSelectionDialog(
            self,
            repository=self.repository,
            target_language_code=self.target_language_code,
            selected_tag_ids=self.selected_tag_ids,
            include_part_of_speech=True,
            title="Select tags",
            text_font=self.text_fonts["latin"],
        )
        self.wait_window(dialog)
        if dialog.result is None:
            return

        self.selected_tag_ids = list(dialog.result)
        self._refresh_tag_summary()

    def _refresh_tag_summary(self) -> None:
        if self._tags_chip_area is None:
            return

        for chip in self._tag_chip_labels:
            chip.destroy()
        self._tag_chip_labels = []

        type_palette: tuple[tuple[str, str], ...] = (
            ("#dff4ff", "#cdeaff"),
            ("#e8f8e1", "#d8efcf"),
            ("#fff1dc", "#ffe5c4"),
            ("#fde8ef", "#f8d8e4"),
            ("#f1e8ff", "#e7d9ff"),
            ("#e7f7f4", "#d7f0eb"),
            ("#f4efdf", "#ece4cd"),
        )

        if not self.selected_tag_ids:
            empty_label = tk.Label(
                self._tags_chip_area,
                text="No tags selected",
                background="#ffffff",
                foreground="#667085",
                anchor="w",
                justify="left",
                font=self.text_fonts["latin"],
            )
            empty_label.grid(row=0, column=0, sticky="w")
            self._tag_chip_labels.append(empty_label)
            return

        try:
            tags = self.repository.list_tags(
                target_language_code=self.target_language_code,
                include_part_of_speech=True,
            )
        except sqlite3.Error:
            fallback_label = tk.Label(
                self._tags_chip_area,
                text=f"{len(self.selected_tag_ids)} tags selected",
                background="#ffffff",
                foreground="#667085",
                anchor="w",
                justify="left",
                font=self.text_fonts["latin"],
            )
            fallback_label.grid(row=0, column=0, sticky="w")
            self._tag_chip_labels.append(fallback_label)
            return

        selected_tag_rows = {
            tag_id: (type_name, tag_name)
            for tag_id, _type_id, type_name, tag_name, _type_predefined, _tag_predefined in tags
        }
        selected_rows = [selected_tag_rows[tag_id] for tag_id in self.selected_tag_ids if tag_id in selected_tag_rows]
        if not selected_rows:
            fallback_label = tk.Label(
                self._tags_chip_area,
                text=f"{len(self.selected_tag_ids)} tags selected",
                background="#ffffff",
                foreground="#667085",
                anchor="w",
                justify="left",
                font=self.text_fonts["latin"],
            )
            fallback_label.grid(row=0, column=0, sticky="w")
            self._tag_chip_labels.append(fallback_label)
            return

        sorted_type_names = sorted({type_name for type_name, _tag_name in selected_rows})
        color_by_type = {
            type_name: type_palette[index % len(type_palette)]
            for index, type_name in enumerate(sorted_type_names)
        }

        max_columns = 4
        for index, (type_name, tag_name) in enumerate(selected_rows):
            row = index // max_columns
            column = index % max_columns
            selected_fill, _selected_active_fill = color_by_type[type_name]
            chip_label = tk.Label(
                self._tags_chip_area,
                text=f"{self._display_type_name(type_name)}: {tag_name}",
                background=selected_fill,
                foreground="#1f2933",
                relief="solid",
                borderwidth=1,
                padx=8,
                pady=3,
                anchor="center",
                justify="center",
                font=self.text_fonts["latin"],
            )
            chip_label.grid(row=row, column=column, sticky="w", padx=(0, 6), pady=(0, 6))
            self._tag_chip_labels.append(chip_label)

    @staticmethod
    def _display_type_name(type_name: str) -> str:
        return type_name.replace("_", " ")

    @staticmethod
    def _select_part_of_speech_value(pos_values: list[str]) -> str:
        ordered_values = [value for value in PART_OF_SPEECH_OPTIONS if value in pos_values]
        if ordered_values:
            return ordered_values[0]
        return sorted(pos_values)[0]

    def _normalize_selected_tags_for_save(self) -> tuple[list[int], str]:
        normalized_tag_ids = list(self.selected_tag_ids)
        selected_part_of_speech_tag_ids: list[int] = []
        selected_part_of_speech_values: list[str] = []

        try:
            tags = self.repository.list_tags(
                target_language_code=self.target_language_code,
                include_part_of_speech=True,
            )
        except sqlite3.Error:
            return normalized_tag_ids, ""

        selected_lookup = set(self.selected_tag_ids)
        for tag_id, _type_id, type_name, tag_name, _type_predefined, _tag_predefined in tags:
            if tag_id not in selected_lookup:
                continue
            if type_name.lower() != "part_of_speech":
                continue
            selected_part_of_speech_tag_ids.append(tag_id)
            selected_part_of_speech_values.append(tag_name)

        if not selected_part_of_speech_values:
            return normalized_tag_ids, ""

        selected_part_of_speech = self._select_part_of_speech_value(selected_part_of_speech_values)
        chosen_tag_id: int | None = None
        for tag_id, _type_id, type_name, tag_name, _type_predefined, _tag_predefined in tags:
            if type_name.lower() == "part_of_speech" and tag_name == selected_part_of_speech:
                chosen_tag_id = tag_id
                break

        pruned_tag_ids = [tag_id for tag_id in normalized_tag_ids if tag_id not in set(selected_part_of_speech_tag_ids)]
        if chosen_tag_id is not None:
            pruned_tag_ids.append(chosen_tag_id)

        return sorted(set(pruned_tag_ids)), selected_part_of_speech

    def _save(self) -> None:
        details_value = self.details_editor.get("1.0", "end-1c") if self._is_editing_markdown else self.details_markdown
        normalized_tag_ids, part_of_speech_value = self._normalize_selected_tags_for_save()

        try:
            self.repository.update_entry(
                self.entry_id,
                self.target_var.get(),
                self.kana_var.get(),
                self.assistant_var.get(),
                part_of_speech_value,
            )
            self.repository.set_entry_tags(
                self.entry_id,
                normalized_tag_ids,
                target_language_code=self.target_language_code,
                include_part_of_speech=True,
            )
            self.repository.update_entry_details(self.entry_id, details_value)
        except ValidationError as exc:
            messagebox.showerror("Validation error", str(exc), parent=self)
            return
        except LookupError as exc:
            messagebox.showerror("Not found", str(exc), parent=self)
            return
        except sqlite3.Error as exc:
            messagebox.showerror("Database error", f"Could not save details: {exc}", parent=self)
            return

        self.on_saved()
        self.destroy()
