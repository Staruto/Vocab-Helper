from __future__ import annotations

import sqlite3
import tkinter as tk
from tkinter import messagebox, ttk
from typing import Callable

from .db import VocabRepository
from .kana import suggest_hiragana
from .validators import ValidationError


class MainWindow(tk.Tk):
    def __init__(self, repository: VocabRepository) -> None:
        super().__init__()
        self.repository = repository

        self.title("JP <-> EN Vocabulary Helper")
        self.geometry("820x520")
        self.minsize(680, 420)

        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        self._build_widgets()
        self.refresh_entries()

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

        self.tree.column("jp", width=260, anchor="w")
        self.tree.column("kana", width=220, anchor="w")
        self.tree.column("en", width=260, anchor="w")

        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        button_row = ttk.Frame(container, padding=(0, 10, 0, 0))
        button_row.grid(row=1, column=0, sticky="ew")
        button_row.columnconfigure(0, weight=1)

        add_button = ttk.Button(button_row, text="+", width=4, command=self._open_add_dialog)
        add_button.grid(row=0, column=1, sticky="e")

    def refresh_entries(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)

        for entry in self.repository.list_entries():
            self.tree.insert(
                "",
                "end",
                values=(entry.japanese_text, entry.kana_text or "", entry.english_text),
            )

    def _open_add_dialog(self) -> None:
        dialog = AddEntryDialog(self, self.repository, on_saved=self.refresh_entries)
        self.wait_window(dialog)


class AddEntryDialog(tk.Toplevel):
    def __init__(
        self,
        parent: tk.Tk,
        repository: VocabRepository,
        on_saved: Callable[[], None],
    ) -> None:
        super().__init__(parent)
        self.parent = parent
        self.repository = repository
        self.on_saved = on_saved

        self._auto_suggest_job: str | None = None
        self._updating_kana = False
        self._kana_user_override = False

        self.japanese_var = tk.StringVar()
        self.kana_var = tk.StringVar()
        self.english_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Kana is optional. You can edit any suggestion.")

        self.title("Add vocabulary")
        self.transient(parent)
        self.grab_set()
        self.resizable(False, False)

        self._build_widgets()

        self.japanese_var.trace_add("write", self._on_japanese_text_change)
        self.kana_var.trace_add("write", self._on_kana_text_change)

        self.japanese_entry.focus_set()

    def _build_widgets(self) -> None:
        frame = ttk.Frame(self, padding=14)
        frame.grid(row=0, column=0, sticky="nsew")

        ttk.Label(frame, text="Japanese writing *").grid(row=0, column=0, sticky="w")
        self.japanese_entry = ttk.Entry(frame, textvariable=self.japanese_var, width=48)
        self.japanese_entry.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 10))

        ttk.Label(frame, text="Kana (hiragana, optional)").grid(row=2, column=0, sticky="w")
        self.kana_entry = ttk.Entry(frame, textvariable=self.kana_var, width=48)
        self.kana_entry.grid(row=3, column=0, sticky="ew", pady=(0, 10))

        suggest_button = ttk.Button(frame, text="Suggest kana", command=self._suggest_kana_manual)
        suggest_button.grid(row=3, column=1, sticky="e", padx=(8, 0), pady=(0, 10))

        ttk.Label(frame, text="English meaning *").grid(row=4, column=0, sticky="w")
        self.english_entry = ttk.Entry(frame, textvariable=self.english_var, width=48)
        self.english_entry.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(0, 8))

        status_label = ttk.Label(frame, textvariable=self.status_var, foreground="#444444", wraplength=460)
        status_label.grid(row=6, column=0, columnspan=2, sticky="w", pady=(2, 12))

        actions = ttk.Frame(frame)
        actions.grid(row=7, column=0, columnspan=2, sticky="e")

        cancel_button = ttk.Button(actions, text="Cancel", command=self.destroy)
        save_button = ttk.Button(actions, text="Save", command=self._save)
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
            self.repository.add_entry(
                self.japanese_var.get(),
                self.kana_var.get(),
                self.english_var.get(),
            )
        except ValidationError as exc:
            messagebox.showerror("Validation error", str(exc), parent=self)
            return
        except sqlite3.Error as exc:
            messagebox.showerror("Database error", f"Could not save entry: {exc}", parent=self)
            return

        self.on_saved()
        self.destroy()
