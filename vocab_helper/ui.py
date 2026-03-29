from __future__ import annotations

import sqlite3
import tkinter as tk
import tkinter.font as tkfont
from tkinter import messagebox, ttk
from typing import Callable

from .db import VocabRepository
from .kana import suggest_hiragana
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


class MainWindow(tk.Tk):
    def __init__(self, repository: VocabRepository) -> None:
        super().__init__()
        self.repository = repository
        self.fonts = _build_font_set(self)
        self._tree_entry_ids: dict[str, int] = {}

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
        )
        self.tree.heading("jp", text="Japanese writing")
        self.tree.heading("kana", text="Kana (hiragana)")
        self.tree.heading("en", text="English meaning")
        self.tree.column("jp", width=300, anchor="w")
        self.tree.column("kana", width=250, anchor="w")
        self.tree.column("en", width=300, anchor="w")

        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        self.context_menu = tk.Menu(self, tearoff=0, font=self.fonts["latin"])
        self.context_menu.add_command(label="Edit", command=self._open_edit_dialog)
        self.context_menu.add_command(label="Delete", command=self._delete_selected_entry)
        self.tree.bind("<Button-3>", self._show_context_menu)

        button_row = ttk.Frame(container, padding=(0, 10, 0, 0))
        button_row.grid(row=1, column=0, sticky="ew")
        button_row.columnconfigure(0, weight=1)

        add_button = ttk.Button(button_row, text="+", width=4, command=self._open_add_dialog, style="App.TButton")
        add_button.grid(row=0, column=1, sticky="e")

        status_row = ttk.Frame(container, padding=(0, 8, 0, 0))
        status_row.grid(row=2, column=0, sticky="ew")
        status_row.columnconfigure(0, weight=1)

        self.count_label = ttk.Label(status_row, text="Total vocabularies: 0", style="Status.TLabel")
        self.count_label.grid(row=0, column=0, sticky="w")

    def _bind_shortcuts(self) -> None:
        self.bind("<Control-n>", self._handle_add_shortcut)
        self.bind("<Control-N>", self._handle_add_shortcut)

        # Keep Enter/Delete scoped to the list so text inputs in dialogs are unaffected.
        self.tree.bind("<Return>", self._handle_edit_shortcut)
        self.tree.bind("<KP_Enter>", self._handle_edit_shortcut)
        self.tree.bind("<Delete>", self._handle_delete_shortcut)

    def _handle_add_shortcut(self, _event: tk.Event) -> str:
        self._open_add_dialog()
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
        entries = self.repository.list_entries()

        for entry in entries:
            item_id = self.tree.insert(
                "",
                "end",
                values=(entry.japanese_text, entry.kana_text or "", entry.english_text),
            )
            self._tree_entry_ids[item_id] = entry.id

        self.count_label.configure(text=f"Total vocabularies: {len(entries)}")

    def _show_context_menu(self, event: tk.Event) -> None:
        item_id = self.tree.identify_row(event.y)
        if not item_id:
            return

        self.tree.selection_set(item_id)
        self.tree.focus(item_id)

        try:
            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()

    def _selected_entry_id(self) -> int | None:
        selection = self.tree.selection()
        if not selection:
            return None
        return self._tree_entry_ids.get(selection[0])

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
        )
        self.wait_window(dialog)

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
            save_handler=lambda japanese, kana, english: self.repository.update_entry(
                entry_id,
                japanese,
                kana,
                english,
            ),
            on_saved=self.refresh_entries,
            initial_japanese=entry.japanese_text,
            initial_kana=entry.kana_text or "",
            initial_english=entry.english_text,
        )
        self.wait_window(dialog)

    def _delete_selected_entry(self) -> None:
        entry_id = self._selected_entry_id()
        if entry_id is None:
            messagebox.showwarning("No selection", "Select an entry to delete.", parent=self)
            return

        selected_text = self._selected_japanese_text()
        confirmed = messagebox.askyesno(
            "Delete entry",
            f"Delete '{selected_text}'? This action cannot be undone.",
            parent=self,
        )
        if not confirmed:
            return

        try:
            self.repository.delete_entry(entry_id)
        except LookupError as exc:
            messagebox.showerror("Not found", str(exc), parent=self)
        except sqlite3.Error as exc:
            messagebox.showerror("Database error", f"Could not delete entry: {exc}", parent=self)

        self.refresh_entries()


class EntryDialog(tk.Toplevel):
    def __init__(
        self,
        parent: tk.Tk,
        title: str,
        save_button_text: str,
        save_handler: Callable[[str, str, str], object],
        on_saved: Callable[[], None],
        initial_japanese: str,
        initial_kana: str,
        initial_english: str,
    ) -> None:
        super().__init__(parent)
        self.save_handler = save_handler
        self.on_saved = on_saved

        self._auto_suggest_job: str | None = None
        self._updating_kana = False
        self._kana_user_override = bool(initial_kana.strip())

        self.japanese_var = tk.StringVar(value=initial_japanese)
        self.kana_var = tk.StringVar(value=initial_kana)
        self.english_var = tk.StringVar(value=initial_english)
        self.status_var = tk.StringVar(value="Kana is optional. You can edit any suggestion.")

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

        ttk.Label(frame, text="Japanese writing *", style="App.TLabel").grid(row=0, column=0, sticky="w")
        self.japanese_entry = ttk.Entry(frame, textvariable=self.japanese_var, width=48, style="Japanese.TEntry")
        self.japanese_entry.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 10))

        ttk.Label(frame, text="Kana (hiragana, optional)", style="App.TLabel").grid(row=2, column=0, sticky="w")
        self.kana_entry = ttk.Entry(frame, textvariable=self.kana_var, width=48, style="Japanese.TEntry")
        self.kana_entry.grid(row=3, column=0, sticky="ew", pady=(0, 10))

        suggest_button = ttk.Button(frame, text="Suggest kana", command=self._suggest_kana_manual, style="App.TButton")
        suggest_button.grid(row=3, column=1, sticky="e", padx=(8, 0), pady=(0, 10))

        ttk.Label(frame, text="English meaning *", style="App.TLabel").grid(row=4, column=0, sticky="w")
        self.english_entry = ttk.Entry(frame, textvariable=self.english_var, width=48, style="App.TEntry")
        self.english_entry.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(0, 8))

        status_label = ttk.Label(frame, textvariable=self.status_var, style="Status.TLabel", wraplength=460)
        status_label.grid(row=6, column=0, columnspan=2, sticky="w", pady=(2, 12))

        actions = ttk.Frame(frame)
        actions.grid(row=7, column=0, columnspan=2, sticky="e")

        cancel_button = ttk.Button(actions, text="Cancel", command=self.destroy, style="App.TButton")
        save_button = ttk.Button(actions, text=save_button_text, command=self._save, style="App.TButton")
        cancel_button.grid(row=0, column=0, padx=(0, 8))
        save_button.grid(row=0, column=1)

    def _on_japanese_text_change(self, *_: object) -> None:
        if self._auto_suggest_job is not None:
            self.after_cancel(self._auto_suggest_job)
        self._auto_suggest_job = self.after(300, self._suggest_kana_automatic)

    def _on_kana_text_change(self, *_: object) -> None:
        if self._updating_kana:
            return
        self._kana_user_override = bool(self.kana_var.get().strip())

    def _suggest_kana_automatic(self) -> None:
        self._auto_suggest_job = None
        if self._kana_user_override and self.kana_var.get().strip():
            return
        self._suggest_kana()

    def _suggest_kana_manual(self) -> None:
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
