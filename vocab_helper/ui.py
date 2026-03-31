from __future__ import annotations

import sqlite3
import tkinter as tk
import tkinter.font as tkfont
from tkinter import messagebox, ttk
from typing import Callable

from .db import VocabRepository
from .kana import suggest_hiragana
from .models import VocabEntry
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


class MainWindow(tk.Tk):
    def __init__(self, repository: VocabRepository) -> None:
        super().__init__()
        self.repository = repository
        self.fonts = _build_font_set(self)
        self._tree_entry_ids: dict[str, int] = {}
        self._entry_stats_by_id: dict[int, tuple[int, int, str]] = {}
        self.target_language_code, self.assistant_language_code = self.repository.get_language_settings()

        self.show_tier_colors_var = tk.BooleanVar(value=True)
        self.sort_mode_var = tk.StringVar(value="time")
        self.time_order_var = tk.StringVar(value="newest")
        self.test_pick_strategy_var = tk.StringVar(value="strict")

        self.title("JP <-> EN Vocabulary Helper")
        self.geometry("900x580")
        self.minsize(760, 460)

        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        self._configure_styles()
        self._build_widgets()
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

    def _language_display_name(self, code: str) -> str:
        return LANGUAGE_NAMES.get(code.upper(), code.upper())

    def _target_field_label(self) -> str:
        return f"Target text ({self._language_display_name(self.target_language_code)})"

    def _assistant_field_label(self) -> str:
        return f"Assistant meaning ({self._language_display_name(self.assistant_language_code)})"

    def _refresh_language_labels(self) -> None:
        self.tree.heading("jp", text=self._target_field_label())
        self.tree.heading("kana", text="Kana (optional)")
        self.tree.heading("en", text=self._assistant_field_label())

    def _build_widgets(self) -> None:
        container = ttk.Frame(self, padding=12)
        container.grid(row=0, column=0, sticky="nsew")
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)

        table_frame = ttk.Frame(container)
        table_frame.grid(row=0, column=0, sticky="nsew")
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(
            table_frame,
            columns=("jp", "kana", "en"),
            show="headings",
            height=15,
            selectmode="extended",
        )
        self.tree.heading("jp", text="Target text")
        self.tree.heading("kana", text="Kana (optional)")
        self.tree.heading("en", text="Assistant meaning")
        self.tree.column("jp", width=300, anchor="w")
        self.tree.column("kana", width=250, anchor="w")
        self.tree.column("en", width=300, anchor="w")

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
        self.tree.bind("<Button-1>", self._on_tree_left_click, add="+")
        self.tree.bind("<Button-3>", self._show_context_menu)
        self.tree.bind("<Double-Button-1>", self._on_tree_double_click)

        self.tree.tag_configure("tier_gray", background=TIER_BG_COLORS["gray"])
        self.tree.tag_configure("tier_green", background=TIER_BG_COLORS["green"])
        self.tree.tag_configure("tier_yellow", background=TIER_BG_COLORS["yellow"])
        self.tree.tag_configure("tier_red", background=TIER_BG_COLORS["red"])

        button_row = ttk.Frame(container, padding=(0, 10, 0, 0))
        button_row.grid(row=1, column=0, sticky="ew")
        button_row.columnconfigure(0, weight=1)

        test_button = ttk.Button(
            button_row,
            text="Test EN->JP",
            command=self._open_en_to_jp_test_dialog,
            style="App.TButton",
        )
        test_button.grid(row=0, column=1, padx=(0, 8), sticky="e")

        test_kana_button = ttk.Button(
            button_row,
            text="Test JP->Kana",
            command=self._open_jp_to_kana_test_dialog,
            style="App.TButton",
        )
        test_kana_button.grid(row=0, column=2, padx=(0, 8), sticky="e")

        test_en_choice_button = ttk.Button(
            button_row,
            text="Test JP->EN",
            command=self._open_jp_to_en_test_dialog,
            style="App.TButton",
        )
        test_en_choice_button.grid(row=0, column=3, padx=(0, 8), sticky="e")

        bulk_add_button = ttk.Button(
            button_row,
            text="Bulk add",
            command=self._open_bulk_add_dialog,
            style="App.TButton",
        )
        bulk_add_button.grid(row=0, column=4, padx=(0, 8), sticky="e")

        add_button = ttk.Button(button_row, text="+", width=4, command=self._open_add_dialog, style="App.TButton")
        add_button.grid(row=0, column=5, sticky="e")

        settings_row = ttk.Frame(container, padding=(0, 8, 0, 0))
        settings_row.grid(row=2, column=0, sticky="ew")
        settings_row.columnconfigure(8, weight=1)

        ttk.Checkbutton(
            settings_row,
            text="Tier colors",
            variable=self.show_tier_colors_var,
            command=self.refresh_entries,
        ).grid(row=0, column=0, sticky="w", padx=(0, 10))

        ttk.Label(settings_row, text="Sort", style="App.TLabel").grid(row=0, column=1, sticky="w")
        self.sort_mode_combo = ttk.Combobox(
            settings_row,
            values=("time", "stats"),
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

        ttk.Button(
            settings_row,
            text="Languages",
            command=self._open_language_settings_dialog,
            style="App.TButton",
        ).grid(row=0, column=7, sticky="w", padx=(10, 0))

        self._update_sort_controls()
        self._refresh_language_labels()

        status_row = ttk.Frame(container, padding=(0, 8, 0, 0))
        status_row.grid(row=3, column=0, sticky="ew")
        status_row.columnconfigure(0, weight=1)

        self.count_label = ttk.Label(status_row, text="Total vocabularies: 0", style="Status.TLabel")
        self.count_label.grid(row=0, column=0, sticky="w")

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

    def refresh_entries(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)

        self._tree_entry_ids.clear()
        self._entry_stats_by_id.clear()
        entries_with_stats = self.repository.list_entries_with_stats(
            sort_mode=self.sort_mode_var.get(),
            time_order=self.time_order_var.get(),
        )

        for entry, test_count, error_count, tier in entries_with_stats:
            tags: tuple[str, ...] = ()
            if self.show_tier_colors_var.get():
                tags = (f"tier_{tier}",)

            item_id = self.tree.insert(
                "",
                "end",
                values=(entry.japanese_text, entry.kana_text or "", entry.english_text),
                tags=tags,
            )
            self._tree_entry_ids[item_id] = entry.id
            self._entry_stats_by_id[entry.id] = (test_count, error_count, tier)

        self.count_label.configure(text=f"Total vocabularies: {len(entries_with_stats)}")

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
        dialog = EntryDialog(
            self,
            title="Add vocabulary",
            save_button_text="Save",
            save_handler=self.repository.add_entry,
            on_saved=self.refresh_entries,
            initial_japanese="",
            initial_kana="",
            initial_english="",
            initial_part_of_speech="",
            target_label=self._target_field_label(),
            assistant_label=self._assistant_field_label(),
            enable_kana_suggest=self.target_language_code == "JP",
        )
        self.wait_window(dialog)

    def _open_bulk_add_dialog(self) -> None:
        dialog = BulkAddDialog(
            self,
            repository=self.repository,
            on_saved=self.refresh_entries,
            text_font=self.fonts["latin"],
        )
        self.wait_window(dialog)

    def _open_en_to_jp_test_dialog(self) -> None:
        dialog = EnglishToJapaneseTestDialog(
            self,
            repository=self.repository,
            text_font=self.fonts["latin"],
            pick_strategy=self.test_pick_strategy_var.get(),
        )
        self.wait_window(dialog)
        self.refresh_entries()

    def _open_jp_to_kana_test_dialog(self) -> None:
        dialog = JapaneseToKanaTestDialog(
            self,
            repository=self.repository,
            text_font=self.fonts["japanese"],
            pick_strategy=self.test_pick_strategy_var.get(),
        )
        self.wait_window(dialog)
        self.refresh_entries()

    def _open_jp_to_en_test_dialog(self) -> None:
        dialog = JapaneseToEnglishChoiceTestDialog(
            self,
            repository=self.repository,
            text_font=self.fonts["japanese"],
            pick_strategy=self.test_pick_strategy_var.get(),
        )
        self.wait_window(dialog)
        self.refresh_entries()

    def _open_edit_dialog(self) -> None:
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

        dialog = EntryDialog(
            self,
            title="Edit vocabulary",
            save_button_text="Save changes",
            save_handler=lambda japanese, kana, english, part_of_speech: self.repository.update_entry(
                entry_id,
                japanese,
                kana,
                english,
                part_of_speech,
            ),
            on_saved=self.refresh_entries,
            initial_japanese=entry.japanese_text,
            initial_kana=entry.kana_text or "",
            initial_english=entry.english_text,
            initial_part_of_speech=entry.part_of_speech or "",
            target_label=self._target_field_label(),
            assistant_label=self._assistant_field_label(),
            enable_kana_suggest=self.target_language_code == "JP",
        )
        self.wait_window(dialog)

    def _open_detail_dialog(self, entry_id: int | None = None) -> None:
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
            enable_kana_suggest=self.target_language_code == "JP",
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
        self._refresh_language_labels()
        self.refresh_entries()

    def _delete_selected_entry(self) -> None:
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

        self.questions: list[VocabEntry] = []
        self.current_index = 0
        self.correct_count = 0
        self.current_answered = False
        self.requested_count = 15
        self.actual_count = 0

        self.count_var = tk.StringVar(value="15")
        self.start_info_var = tk.StringVar(value="")
        self.progress_var = tk.StringVar(value="")
        self.score_var = tk.StringVar(value="")
        self.prompt_var = tk.StringVar(value="")
        self.answer_var = tk.StringVar(value="")
        self.feedback_var = tk.StringVar(value="")
        self.result_var = tk.StringVar(value="")
        self.pick_strategy_var = tk.StringVar(value=pick_strategy if pick_strategy in {"strict", "weighted"} else "strict")

        self.title("Test mode: English -> Japanese")
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
            text="English -> Japanese test",
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

        ttk.Label(self.test_frame, text="English meaning", style="App.TLabel").grid(row=1, column=0, sticky="w")
        self.prompt_label = ttk.Label(
            self.test_frame,
            textvariable=self.prompt_var,
            style="App.TLabel",
            wraplength=600,
            font=self.text_font,
        )
        self.prompt_label.grid(row=2, column=0, sticky="w", pady=(4, 12))

        ttk.Label(self.test_frame, text="Your Japanese answer", style="App.TLabel").grid(row=3, column=0, sticky="w")
        self.answer_entry = ttk.Entry(self.test_frame, textvariable=self.answer_var, width=38, style="Japanese.TEntry")
        self.answer_entry.grid(row=4, column=0, sticky="w", pady=(4, 8))

        self.feedback_label = ttk.Label(self.test_frame, textvariable=self.feedback_var, style="Status.TLabel", wraplength=600)
        self.feedback_label.grid(row=5, column=0, sticky="w", pady=(0, 10))

        test_actions = ttk.Frame(self.test_frame)
        test_actions.grid(row=6, column=0, sticky="e")
        self.submit_button = ttk.Button(test_actions, text="Submit", command=self._submit_answer, style="App.TButton")
        self.submit_button.grid(row=0, column=0, padx=(0, 8))
        self.next_button = ttk.Button(test_actions, text="Next", command=self._next_question, style="App.TButton")
        self.next_button.grid(row=0, column=1)

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

        if requested > self.actual_count:
            self.start_info_var.set(f"Requested {requested}; using all {self.actual_count} available vocabularies.")
        else:
            self.start_info_var.set("")

        self._show_test_frame()
        self._load_question()

    def _load_question(self) -> None:
        if not self.questions:
            return

        current = self.questions[self.current_index]
        self.progress_var.set(f"Question {self.current_index + 1}/{self.actual_count}")
        self.score_var.set(f"Score: {self.correct_count}")
        self.prompt_var.set(current.english_text)
        self.answer_var.set("")
        self.feedback_var.set("")
        self.current_answered = False
        self.answer_entry.configure(state="normal")
        self.submit_button.configure(state="normal")
        self.next_button.configure(state="disabled")
        self.answer_entry.focus_set()

    def _submit_answer(self) -> None:
        if self.current_answered or not self.questions:
            return

        current = self.questions[self.current_index]
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
            self.correct_count += 1
            self.feedback_var.set("Correct.")
        else:
            kana_hint = (current.kana_text or "").strip()
            if not kana_hint:
                suggested_kana, reliable, _message = suggest_hiragana(correct_answer)
                if reliable and suggested_kana:
                    kana_hint = suggested_kana

            if kana_hint:
                self.feedback_var.set(f"Incorrect. Correct answer: {correct_answer} ({kana_hint})")
            else:
                self.feedback_var.set(f"Incorrect. Correct answer: {correct_answer}")

        self.score_var.set(f"Score: {self.correct_count}")
        self.current_answered = True
        self.answer_entry.configure(state="disabled")
        self.submit_button.configure(state="disabled")
        self.next_button.configure(state="normal")
        self.next_button.focus_set()

    def _next_question(self) -> None:
        if not self.current_answered:
            return

        if self.current_index + 1 >= self.actual_count:
            self._finish_test()
            return

        self.current_index += 1
        self._load_question()

    def _finish_test(self) -> None:
        if self.actual_count <= 0:
            self.result_var.set("No questions were completed.")
            self._show_result_frame()
            return

        accuracy = (self.correct_count / self.actual_count) * 100
        self.result_var.set(
            f"Score: {self.correct_count}/{self.actual_count} ({accuracy:.1f}%)."
        )
        self._show_result_frame()

    def _restart(self) -> None:
        self.questions = []
        self.current_index = 0
        self.correct_count = 0
        self.actual_count = 0
        self.current_answered = False
        self.progress_var.set("")
        self.score_var.set("")
        self.prompt_var.set("")
        self.answer_var.set("")
        self.feedback_var.set("")
        self.result_var.set("")
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

        self.questions: list[VocabEntry] = []
        self.current_index = 0
        self.correct_count = 0
        self.current_answered = False
        self.requested_count = 15
        self.actual_count = 0

        self.count_var = tk.StringVar(value="15")
        self.start_info_var = tk.StringVar(value="")
        self.progress_var = tk.StringVar(value="")
        self.score_var = tk.StringVar(value="")
        self.prompt_var = tk.StringVar(value="")
        self.answer_var = tk.StringVar(value="")
        self.feedback_var = tk.StringVar(value="")
        self.result_var = tk.StringVar(value="")
        self.pick_strategy_var = tk.StringVar(value=pick_strategy if pick_strategy in {"strict", "weighted"} else "strict")

        self.title("Test mode: Japanese -> Kana")
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
            text="Japanese -> Kana test",
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

        ttk.Label(self.test_frame, text="Japanese writing", style="App.TLabel").grid(row=1, column=0, sticky="w")
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
        self.next_button = ttk.Button(test_actions, text="Next", command=self._next_question, style="App.TButton")
        self.next_button.grid(row=0, column=1)

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
                "Add vocabularies with kana before starting a JP->Kana test.",
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

        if requested > self.actual_count:
            self.start_info_var.set(f"Requested {requested}; using all {self.actual_count} eligible vocabularies.")
        else:
            self.start_info_var.set("")

        self._show_test_frame()
        self._load_question()

    def _load_question(self) -> None:
        if not self.questions:
            return

        current = self.questions[self.current_index]
        self.progress_var.set(f"Question {self.current_index + 1}/{self.actual_count}")
        self.score_var.set(f"Score: {self.correct_count}")
        self.prompt_var.set(current.japanese_text)
        self.answer_var.set("")
        self.feedback_var.set("")
        self.current_answered = False
        self.answer_entry.configure(state="normal")
        self.submit_button.configure(state="normal")
        self.next_button.configure(state="disabled")
        self.answer_entry.focus_set()

    def _submit_answer(self) -> None:
        if self.current_answered or not self.questions:
            return

        current = self.questions[self.current_index]
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
            self.correct_count += 1
            self.feedback_var.set("Correct.")
        else:
            self.feedback_var.set(f"Incorrect. Correct answer: {correct_answer}")

        self.score_var.set(f"Score: {self.correct_count}")
        self.current_answered = True
        self.answer_entry.configure(state="disabled")
        self.submit_button.configure(state="disabled")
        self.next_button.configure(state="normal")
        self.next_button.focus_set()

    def _next_question(self) -> None:
        if not self.current_answered:
            return

        if self.current_index + 1 >= self.actual_count:
            self._finish_test()
            return

        self.current_index += 1
        self._load_question()

    def _finish_test(self) -> None:
        if self.actual_count <= 0:
            self.result_var.set("No questions were completed.")
            self._show_result_frame()
            return

        accuracy = (self.correct_count / self.actual_count) * 100
        self.result_var.set(
            f"Score: {self.correct_count}/{self.actual_count} ({accuracy:.1f}%)."
        )
        self._show_result_frame()

    def _restart(self) -> None:
        self.questions = []
        self.current_index = 0
        self.correct_count = 0
        self.actual_count = 0
        self.current_answered = False
        self.progress_var.set("")
        self.score_var.set("")
        self.prompt_var.set("")
        self.answer_var.set("")
        self.feedback_var.set("")
        self.result_var.set("")
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

        self.questions: list[VocabEntry] = []
        self.options_by_question: list[list[str]] = []
        self.current_index = 0
        self.correct_count = 0
        self.current_answered = False
        self.requested_count = 15
        self.actual_count = 0

        self.count_var = tk.StringVar(value="15")
        self.start_info_var = tk.StringVar(value="")
        self.progress_var = tk.StringVar(value="")
        self.score_var = tk.StringVar(value="")
        self.prompt_var = tk.StringVar(value="")
        self.feedback_var = tk.StringVar(value="")
        self.result_var = tk.StringVar(value="")
        self.pick_strategy_var = tk.StringVar(value=pick_strategy if pick_strategy in {"strict", "weighted"} else "strict")

        self.title("Test mode: Japanese -> English")
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
            text="Japanese -> English test (single choice)",
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

        ttk.Label(self.test_frame, text="Japanese writing", style="App.TLabel").grid(row=1, column=0, sticky="w")
        self.prompt_label = ttk.Label(
            self.test_frame,
            textvariable=self.prompt_var,
            style="App.TLabel",
            wraplength=600,
            font=self.text_font,
        )
        self.prompt_label.grid(row=2, column=0, sticky="w", pady=(4, 12))

        ttk.Label(self.test_frame, text="Choose the correct English meaning", style="App.TLabel").grid(
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
        self.next_button = ttk.Button(test_actions, text="Next", command=self._next_question, style="App.TButton")
        self.next_button.grid(row=0, column=1)

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
                "Add at least two vocabularies with different English meanings before starting JP->EN test.",
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

        if requested > self.actual_count:
            self.start_info_var.set(f"Requested {requested}; using all {self.actual_count} eligible vocabularies.")
        else:
            self.start_info_var.set("")

        self._show_test_frame()
        self._load_question()

    def _load_question(self) -> None:
        if not self.questions:
            return

        current = self.questions[self.current_index]
        options = self.options_by_question[self.current_index]

        self.progress_var.set(f"Question {self.current_index + 1}/{self.actual_count}")
        self.score_var.set(f"Score: {self.correct_count}")
        self.prompt_var.set(current.japanese_text)
        self.feedback_var.set("")
        self.current_answered = False

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
        if self.current_answered or not self.questions:
            return

        current = self.questions[self.current_index]
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
            self.correct_count += 1
            self.feedback_var.set("Correct.")
        else:
            self.feedback_var.set(f"Incorrect. Correct answer: {correct_answer}")

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

        if self.current_index + 1 >= self.actual_count:
            self._finish_test()
            return

        self.current_index += 1
        self._load_question()

    def _finish_test(self) -> None:
        if self.actual_count <= 0:
            self.result_var.set("No questions were completed.")
            self._show_result_frame()
            return

        accuracy = (self.correct_count / self.actual_count) * 100
        self.result_var.set(
            f"Score: {self.correct_count}/{self.actual_count} ({accuracy:.1f}%)."
        )
        self._show_result_frame()

    def _restart(self) -> None:
        self.questions = []
        self.options_by_question = []
        self.current_index = 0
        self.correct_count = 0
        self.actual_count = 0
        self.current_answered = False
        self.progress_var.set("")
        self.score_var.set("")
        self.prompt_var.set("")
        self.feedback_var.set("")
        self.result_var.set("")
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
    ) -> None:
        super().__init__(parent)
        self.repository = repository
        self.on_saved = on_saved

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

        help_text = "One entry per line, aligned across 3 columns: Japanese, Kana (optional), English."
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

        ttk.Label(labels_row, text="Japanese writing *", style="App.TLabel").grid(
            row=0,
            column=0,
            sticky="w",
            padx=(0, 8),
        )
        ttk.Label(labels_row, text="Kana (optional)", style="App.TLabel").grid(
            row=0,
            column=1,
            sticky="w",
            padx=(0, 8),
        )
        ttk.Label(labels_row, text="English meaning *", style="App.TLabel").grid(
            row=0,
            column=2,
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
        self.kana_text.grid(row=0, column=1, sticky="nsew", padx=(0, 8))
        self.en_text.grid(row=0, column=2, sticky="nsew", padx=(0, 8))

        self.scrollbar = ttk.Scrollbar(columns_frame, orient="vertical", command=self._on_scrollbar)
        self.scrollbar.grid(row=0, column=3, sticky="ns")

        self.jp_text.configure(yscrollcommand=self._on_text_yscroll)
        self.kana_text.configure(yscrollcommand=self._on_text_yscroll)
        self.en_text.configure(yscrollcommand=self._on_text_yscroll)

        for text_widget in (self.jp_text, self.kana_text, self.en_text):
            text_widget.bind("<MouseWheel>", self._on_mousewheel)
            text_widget.bind("<Button-4>", self._on_mousewheel_linux_up)
            text_widget.bind("<Button-5>", self._on_mousewheel_linux_down)

        actions = ttk.Frame(frame, padding=(0, 10, 0, 0))
        actions.grid(row=3, column=0, sticky="e")

        ttk.Button(actions, text="Cancel", command=self.destroy, style="App.TButton").grid(
            row=0,
            column=0,
            padx=(0, 8),
        )
        ttk.Button(actions, text="Add entries", command=self._save, style="App.TButton").grid(
            row=0,
            column=1,
        )

        self.bind("<Escape>", lambda _event: self.destroy())
        self.jp_text.focus_set()

    def _on_scrollbar(self, *args: str) -> None:
        for text_widget in (self.jp_text, self.kana_text, self.en_text):
            text_widget.yview(*args)

    def _on_text_yscroll(self, first: str, last: str) -> None:
        self.scrollbar.set(first, last)

    def _on_mousewheel(self, event: tk.Event) -> str:
        delta = -1 if event.delta > 0 else 1
        for text_widget in (self.jp_text, self.kana_text, self.en_text):
            text_widget.yview_scroll(delta, "units")
        return "break"

    def _on_mousewheel_linux_up(self, _event: tk.Event) -> str:
        for text_widget in (self.jp_text, self.kana_text, self.en_text):
            text_widget.yview_scroll(-1, "units")
        return "break"

    def _on_mousewheel_linux_down(self, _event: tk.Event) -> str:
        for text_widget in (self.jp_text, self.kana_text, self.en_text):
            text_widget.yview_scroll(1, "units")
        return "break"

    def _parse_entries(self, japanese_raw: str, kana_raw: str, english_raw: str) -> list[tuple[str, str, str]]:
        japanese_lines = japanese_raw.splitlines()
        kana_lines = kana_raw.splitlines()
        english_lines = english_raw.splitlines()

        line_count = max(len(japanese_lines), len(kana_lines), len(english_lines))
        parsed_entries: list[tuple[str, str, str]] = []

        for index in range(line_count):
            line_number = index + 1
            japanese_text = japanese_lines[index].strip() if index < len(japanese_lines) else ""
            kana_text = kana_lines[index].strip() if index < len(kana_lines) else ""
            english_text = english_lines[index].strip() if index < len(english_lines) else ""

            if not japanese_text and not kana_text and not english_text:
                continue

            if not japanese_text or not english_text:
                raise ValidationError(
                    f"Line {line_number}: Japanese writing and English meaning are required."
                )

            parsed_entries.append((japanese_text, kana_text, english_text))

        if not parsed_entries:
            raise ValidationError("Add at least one non-empty line.")

        return parsed_entries

    def _save(self) -> None:
        try:
            rows = self._parse_entries(
                self.jp_text.get("1.0", "end-1c"),
                self.kana_text.get("1.0", "end-1c"),
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


class EntryDialog(tk.Toplevel):
    def __init__(
        self,
        parent: tk.Tk,
        title: str,
        save_button_text: str,
        save_handler: Callable[[str, str, str, str], object],
        on_saved: Callable[[], None],
        initial_japanese: str,
        initial_kana: str,
        initial_english: str,
        initial_part_of_speech: str,
        target_label: str,
        assistant_label: str,
        enable_kana_suggest: bool,
    ) -> None:
        super().__init__(parent)
        self.save_handler = save_handler
        self.on_saved = on_saved
        self.target_label = target_label
        self.assistant_label = assistant_label
        self.enable_kana_suggest = enable_kana_suggest

        self._auto_suggest_job: str | None = None
        self._updating_kana = False
        self._kana_user_override = bool(initial_kana.strip())

        self.japanese_var = tk.StringVar(value=initial_japanese)
        self.kana_var = tk.StringVar(value=initial_kana)
        self.english_var = tk.StringVar(value=initial_english)
        self.part_of_speech_var = tk.StringVar(value=initial_part_of_speech)
        default_status = "Kana is optional. You can edit any suggestion."
        if not self.enable_kana_suggest:
            default_status = "Kana is optional. Automatic kana suggestion is only available when target language is Japanese."
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

    def _build_widgets(self, save_button_text: str) -> None:
        frame = ttk.Frame(self, padding=14)
        frame.grid(row=0, column=0, sticky="nsew")

        ttk.Label(frame, text=f"{self.target_label} *", style="App.TLabel").grid(row=0, column=0, sticky="w")
        self.japanese_entry = ttk.Entry(frame, textvariable=self.japanese_var, width=48, style="Japanese.TEntry")
        self.japanese_entry.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 10))

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

        status_label = ttk.Label(frame, textvariable=self.status_var, style="Status.TLabel", wraplength=460)
        status_label.grid(row=8, column=0, columnspan=2, sticky="w", pady=(2, 12))

        actions = ttk.Frame(frame)
        actions.grid(row=9, column=0, columnspan=2, sticky="e")

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

    def _save(self) -> None:
        try:
            self.save_handler(
                self.japanese_var.get(),
                self.kana_var.get(),
                self.english_var.get(),
                self.part_of_speech_var.get(),
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
        enable_kana_suggest: bool,
        on_saved: Callable[[], None],
    ) -> None:
        super().__init__(parent)
        self.repository = repository
        self.entry_id = entry_id
        self.text_fonts = text_fonts
        self.target_label = target_label
        self.assistant_label = assistant_label
        self.enable_kana_suggest = enable_kana_suggest
        self.on_saved = on_saved

        self._auto_suggest_job: str | None = None
        self._updating_kana = False
        self._kana_user_override = False
        self._is_editing_markdown = False

        self.target_var = tk.StringVar(value="")
        self.kana_var = tk.StringVar(value="")
        self.assistant_var = tk.StringVar(value="")
        self.part_of_speech_var = tk.StringVar(value="")
        self.stats_var = tk.StringVar(value="")
        self.created_var = tk.StringVar(value="")
        self.details_markdown = ""

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

        ttk.Label(header, text="Part of speech", style="App.TLabel").grid(row=3, column=0, sticky="w")
        self.part_of_speech_combo = ttk.Combobox(
            header,
            values=PART_OF_SPEECH_OPTIONS,
            state="normal",
            textvariable=self.part_of_speech_var,
        )
        self.part_of_speech_combo.grid(row=3, column=1, sticky="ew", padx=(8, 0), pady=(0, 6))

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
        except LookupError as exc:
            messagebox.showerror("Not found", str(exc), parent=self)
            self.destroy()
            return

        self.target_var.set(entry.japanese_text)
        self.kana_var.set(entry.kana_text or "")
        self.assistant_var.set(entry.english_text)
        self.part_of_speech_var.set(entry.part_of_speech or "")
        self.details_markdown = entry.details_markdown or ""

        self.stats_var.set(f"Tests: {test_count} | Errors: {error_count} | Tier: {tier}")
        self.created_var.set(f"Created at: {entry.created_at}")

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

    def _save(self) -> None:
        details_value = self.details_editor.get("1.0", "end-1c") if self._is_editing_markdown else self.details_markdown

        try:
            self.repository.update_entry(
                self.entry_id,
                self.target_var.get(),
                self.kana_var.get(),
                self.assistant_var.get(),
                self.part_of_speech_var.get(),
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
