from __future__ import annotations

import json
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .archicad_adapter import ArchicadAdapter
from .dbn import load_catalog, open_document
from .excel_service import demo_rows, export_rows, import_rows
from .storage import NoteStore

ROOT = Path(__file__).resolve().parents[1]
STATE_FILE = ROOT / "data" / "ui_state.json"
THEMES = {
    "dark": {"outer": "#131b22", "shadow": "#0a1015", "bg": "#1b232c", "sidebar": "#172028", "panel": "#222c37", "border": "#364451", "text": "#edf2f7", "muted": "#a2afbd", "input": "#1b2530", "button": "#2b3744", "active": "#344454", "accent": "#3b82f6", "selected": "#294a6e"},
    "light": {"outer": "#eaf3f7", "shadow": "#cbd6dc", "bg": "#ffffff", "sidebar": "#edf5f8", "panel": "#ffffff", "border": "#d9e0e4", "text": "#22272e", "muted": "#69747c", "input": "#f3f5f6", "button": "#f0f3f4", "active": "#e2e8eb", "accent": "#2f93e6", "selected": "#dfe8eb"},
}


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("CoDA")
        self.geometry("1280x800")
        self.minsize(980, 640)
        self.store = NoteStore(ROOT / "data")
        self.catalog = load_catalog(ROOT / "data" / "dbn_catalog.xlsx")
        self.adapter = ArchicadAdapter()
        self.state = self._load_state()
        self.theme_name = self.state.get("theme", "dark")
        self.current_page = "dbn"
        self.current_note: dict | None = None
        self.note_attachments: list[dict] = []
        self.note_pinned = False
        self.draft_note_id: str | None = None
        self.pending_note_id: str | None = None
        self._build()

    def _load_state(self) -> dict:
        try:
            return json.loads(STATE_FILE.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return {"theme": "light", "pinned_dbn": [], "recent_dbn": [], "recent_notes": []}

    def _save_state(self) -> None:
        STATE_FILE.parent.mkdir(parents=True, exist_ok=True)
        STATE_FILE.write_text(json.dumps(self.state, ensure_ascii=False, indent=2), encoding="utf-8")

    @property
    def c(self) -> dict:
        return THEMES[self.theme_name]

    def _build(self) -> None:
        for widget in self.winfo_children():
            widget.destroy()
        self.configure(bg=self.c["outer"])
        self._styles()
        self._shell()
        {"dbn": self.show_dbn, "notes": self.show_notes, "excel": self.show_excel}[self.current_page]()

    def _styles(self) -> None:
        c = self.c
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure("TFrame", background=c["bg"])
        s.configure("TLabel", background=c["bg"], foreground=c["text"], font=("Segoe UI", 10))
        s.configure("Muted.TLabel", foreground=c["muted"])
        s.configure("Title.TLabel", font=("Segoe UI Semibold", 22))
        s.configure("Card.TLabel", background=c["panel"], foreground=c["text"])
        s.configure("TButton", background=c["button"], foreground=c["text"], bordercolor=c["border"], padding=(12, 8))
        s.map("TButton", background=[("active", c["active"])])
        s.configure("Accent.TButton", background=c["accent"], foreground="white", bordercolor=c["accent"])
        s.map("Accent.TButton", background=[("active", c["accent"])])
        s.configure("Treeview", background=c["panel"], fieldbackground=c["panel"], foreground=c["text"], rowheight=32, bordercolor=c["border"])
        s.configure("Treeview.Heading", background=c["button"], foreground=c["text"])
        s.map("Treeview", background=[("selected", c["selected"])])

    def _shell(self) -> None:
        c = self.c
        workspace = tk.Frame(self, bg=c["outer"])
        workspace.pack(fill="both", expand=True)
        sidebar = tk.Frame(workspace, bg=c["sidebar"], width=260)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)
        tk.Label(sidebar, text="CoDA  ⌄", bg=c["sidebar"], fg=c["text"], font=("Segoe UI Semibold", 18), anchor="w").pack(fill="x", padx=20, pady=(20, 25))
        for icon, text, page in [("▣", "Довідник ДБН", "dbn"), ("✎", "Нотатки", "notes"), ("⇄", "Archicad / Excel", "excel")]:
            tk.Button(sidebar, text=f"{icon}   {text}", command=lambda p=page: self._go(p), bg=c["sidebar"], fg=c["text"], activebackground=c["active"], activeforeground=c["text"], bd=0, anchor="w", padx=20, pady=9, font=("Segoe UI", 10)).pack(fill="x")
        self._sidebar_section(sidebar, "Закріплені", self._pinned_items())
        self._sidebar_section(sidebar, "Нещодавні", self._recent_items())
        ttk.Button(sidebar, text="☀  Світла тема" if self.theme_name == "dark" else "◐  Темна тема", command=self._toggle_theme).pack(side="bottom", fill="x", padx=14, pady=(0, 12))
        right = tk.Frame(workspace, bg=c["outer"])
        right.pack(side="left", fill="both", expand=True)
        shadow = tk.Frame(right, bg=c["shadow"])
        shadow.pack(fill="both", expand=True, padx=(0, 18), pady=(16, 12))
        self.main = tk.Frame(shadow, bg=c["bg"], highlightbackground=c["border"], highlightthickness=1)
        self.main.place(x=0, y=0, relwidth=1, relheight=1, width=-4, height=-4)

    def _sidebar_section(self, parent: tk.Widget, title: str, items: list[tuple[str, object]]) -> None:
        c = self.c
        if not items:
            return
        tk.Label(parent, text=title, bg=c["sidebar"], fg="#a2aaae" if self.theme_name == "light" else c["muted"], font=("Segoe UI Semibold", 10), anchor="w").pack(fill="x", padx=20, pady=(24, 7))
        for label, command in items[:6]:
            tk.Button(parent, text=label, command=command, bg=c["sidebar"], fg=c["text"], activebackground=c["active"], activeforeground=c["text"], bd=0, anchor="w", padx=20, pady=5, font=("Segoe UI", 10)).pack(fill="x")

    def _pinned_items(self) -> list[tuple[str, object]]:
        items: list[tuple[str, object]] = []
        for note in self.store.list_notes():
            if note.get("pinned"):
                items.append((f"✎  {note['title']}", lambda n=note: self._open_sidebar_note(n)))
        pinned = set(self.state.get("pinned_dbn", []))
        for item in self.catalog:
            if item.get("path") in pinned:
                items.append((f"▣  {item.get('number', 'ДБН')}", lambda x=item: self._open_sidebar_dbn(x)))
        return items

    def _recent_items(self) -> list[tuple[str, object]]:
        items: list[tuple[str, object]] = []
        notes = {note["id"]: note for note in self.store.list_notes()}
        for note_id in self.state.get("recent_notes", []):
            if note_id in notes:
                items.append((f"✎  {notes[note_id]['title']}", lambda n=notes[note_id]: self._open_sidebar_note(n)))
        dbn = {item.get("path"): item for item in self.catalog}
        for path in self.state.get("recent_dbn", []):
            if path in dbn:
                items.append((f"▣  {dbn[path].get('number', 'ДБН')}", lambda x=dbn[path]: self._open_sidebar_dbn(x)))
        return items

    def _open_sidebar_note(self, note: dict) -> None:
        self.pending_note_id = note["id"]
        self._go("notes")

    def _open_sidebar_dbn(self, item: dict) -> None:
        self._mark_recent("recent_dbn", item["path"])
        path = ROOT / item["path"]
        if path.exists():
            open_document(path)

    def _mark_recent(self, key: str, value: str) -> None:
        values = [item for item in self.state.get(key, []) if item != value]
        self.state[key] = [value, *values][:8]
        self._save_state()

    def _go(self, page: str) -> None:
        self.current_page = page
        self._build()

    def _toggle_theme(self) -> None:
        self.theme_name = "light" if self.theme_name == "dark" else "dark"
        self.state["theme"] = self.theme_name
        self._save_state()
        self._build()

    def clear(self) -> None:
        for widget in self.main.winfo_children():
            widget.destroy()

    def header(self, title: str, subtitle: str) -> None:
        box = tk.Frame(self.main, bg=self.c["bg"])
        box.pack(fill="x", padx=34, pady=(28, 18))
        ttk.Label(box, text=title, style="Title.TLabel").pack(anchor="w")
        ttk.Label(box, text=subtitle, style="Muted.TLabel").pack(anchor="w", pady=(5, 0))

    def show_dbn(self) -> None:
        self.clear()
        self.header("Державні будівельні норми", f"Локальна бібліотека • {len(self.catalog)} документів")
        tools = tk.Frame(self.main, bg=self.c["bg"])
        tools.pack(fill="x", padx=34, pady=(0, 12))
        query, category = tk.StringVar(), tk.StringVar(value="Усі категорії")
        tk.Entry(tools, textvariable=query, bg=self.c["input"], fg=self.c["text"], insertbackground=self.c["text"], relief="flat", font=("Segoe UI", 11)).pack(side="left", fill="x", expand=True, ipady=10, padx=(0, 10))
        cats = ["Усі категорії"] + sorted({x.get("category", "Інше") for x in self.catalog})
        ttk.Combobox(tools, textvariable=category, values=cats, state="readonly", width=24).pack(side="right")
        columns = ("pin", "number", "title", "category", "effective", "changes")
        tree = ttk.Treeview(self.main, columns=columns, show="headings")
        for col, title, width in zip(columns, ["", "Номер", "Назва", "Категорія", "Чинний з", "Зміни"], [38, 175, 400, 175, 100, 80]):
            tree.heading(col, text=title)
            tree.column(col, width=width, anchor="w", stretch=col == "title")
        tree.pack(fill="both", expand=True, padx=34, pady=(0, 10))
        bar = tk.Frame(self.main, bg=self.c["bg"])
        bar.pack(fill="x", padx=34, pady=(0, 22))
        ttk.Label(bar, text="Подвійний клік відкриває документ", style="Muted.TLabel").pack(side="left")
        pin_button = ttk.Button(bar, text="Закріпити / відкріпити", command=lambda: self._toggle_dbn_pin(tree))
        pin_button.pack(side="right")

        def refresh(*_args: object) -> None:
            tree.delete(*tree.get_children())
            pinned = set(self.state.get("pinned_dbn", []))
            items = sorted(self.catalog, key=lambda item: (item.get("path") not in pinned, item.get("number", "")))
            for item in items:
                hay = f"{item.get('number', '')} {item.get('title', '')}".lower()
                if query.get().lower() in hay and (category.get() == "Усі категорії" or item.get("category") == category.get()):
                    item_id = self.catalog.index(item)
                    tree.insert("", "end", iid=str(item_id), values=("★" if item.get("path") in pinned else "", item.get("number", "—"), item.get("title", "Без назви"), item.get("category", "Інше"), item.get("effective_date", "Не визначено"), "Так" if item.get("has_changes") else "Ні"))

        def open_selected(_event: object = None) -> None:
            selected = tree.selection()
            if selected:
                path = ROOT / self.catalog[int(selected[0])]["path"]
                if path.exists():
                    self._mark_recent("recent_dbn", self.catalog[int(selected[0])]["path"])
                    open_document(path)
                else:
                    messagebox.showerror("Файл не знайдено", str(path))

        query.trace_add("write", refresh)
        category.trace_add("write", refresh)
        tree.bind("<Double-1>", open_selected)
        self._dbn_refresh = refresh
        refresh()

    def _toggle_dbn_pin(self, tree: ttk.Treeview) -> None:
        selected = tree.selection()
        if not selected:
            messagebox.showinfo("Оберіть ДБН", "Спершу оберіть документ у списку.")
            return
        path = self.catalog[int(selected[0])]["path"]
        pinned = set(self.state.get("pinned_dbn", []))
        if path in pinned:
            pinned.remove(path)
        else:
            pinned.add(path)
        self.state["pinned_dbn"] = sorted(pinned)
        self._save_state()
        self._build()

    def show_notes(self) -> None:
        self.clear()
        self.header("Нотатки", "Текст і вкладення зберігаються локально")
        body = tk.Frame(self.main, bg=self.c["bg"])
        body.pack(fill="both", expand=True, padx=34, pady=(0, 28))
        left = tk.Frame(body, bg=self.c["sidebar"], width=265)
        left.pack(side="left", fill="y", padx=(0, 14))
        left.pack_propagate(False)
        ttk.Button(left, text="＋ Нова нотатка", style="Accent.TButton", command=lambda: self._load_note(None)).pack(fill="x", padx=10, pady=10)
        self.notes_list = tk.Listbox(left, bg=self.c["sidebar"], fg=self.c["text"], selectbackground=self.c["selected"], selectforeground=self.c["text"], bd=0, highlightthickness=0, font=("Segoe UI", 10))
        self.notes_list.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        editor = tk.Frame(body, bg=self.c["panel"], highlightbackground=self.c["border"], highlightthickness=1)
        editor.pack(side="left", fill="both", expand=True)
        self.note_title = tk.Entry(editor, bg=self.c["panel"], fg=self.c["text"], insertbackground=self.c["text"], bd=0, font=("Segoe UI Semibold", 20))
        self.note_title.pack(fill="x", padx=22, pady=(20, 8))
        self.note_body = tk.Text(editor, bg=self.c["panel"], fg=self.c["text"], insertbackground=self.c["text"], bd=0, wrap="word", font=("Segoe UI", 11), undo=True)
        self.note_body.pack(fill="both", expand=True, padx=18, pady=8)
        self.attachments_box = tk.Frame(editor, bg=self.c["panel"])
        self.attachments_box.pack(fill="x", padx=18, pady=(0, 6))
        bar = tk.Frame(editor, bg=self.c["panel"])
        bar.pack(fill="x", padx=18, pady=14)
        ttk.Button(bar, text="Прикріпити файл", command=self._attach).pack(side="left")
        self.pin_note_button = ttk.Button(bar, text="☆ Закріпити", command=self._toggle_note_pin)
        self.pin_note_button.pack(side="left", padx=8)
        ttk.Button(bar, text="Зберегти", style="Accent.TButton", command=self._save_note).pack(side="right")
        ttk.Button(bar, text="Видалити", command=self._delete_note).pack(side="right", padx=8)
        self._refresh_notes()
        self.notes_list.bind("<<ListboxSelect>>", self._select_note)
        selected = next((note for note in self.notes if note["id"] == self.pending_note_id), None)
        self.pending_note_id = None
        self._load_note(selected)

    def _refresh_notes(self) -> None:
        self.notes = self.store.list_notes()
        self.notes_list.delete(0, "end")
        for note in self.notes:
            self.notes_list.insert("end", f"★  {note['title']}" if note.get("pinned") else note["title"])

    def _select_note(self, _event: object) -> None:
        selected = self.notes_list.curselection()
        if selected:
            note = self.notes[selected[0]]
            self._mark_recent("recent_notes", note["id"])
            self._load_note(note)

    def _load_note(self, note: dict | None) -> None:
        self.current_note = note
        self.draft_note_id = note.get("id") if note else None
        self.note_attachments = list(note.get("attachments", [])) if note else []
        self.note_pinned = note.get("pinned", False) if note else False
        self.note_title.delete(0, "end")
        self.note_title.insert(0, note.get("title", "") if note else "")
        self.note_body.delete("1.0", "end")
        self.note_body.insert("1.0", note.get("body", "") if note else "")
        self.pin_note_button.config(text="★ Закріплено" if self.note_pinned else "☆ Закріпити")
        self._render_attachments()

    def _render_attachments(self) -> None:
        for widget in self.attachments_box.winfo_children():
            widget.destroy()
        if not self.note_attachments:
            ttk.Label(self.attachments_box, text="Вкладень немає", style="Muted.TLabel").pack(anchor="w")
            return
        ttk.Label(self.attachments_box, text="Вкладення — натисніть, щоб відкрити", style="Muted.TLabel").pack(anchor="w", pady=(0, 4))
        for attachment in self.note_attachments:
            ttk.Button(self.attachments_box, text=f"↗  {attachment.get('name', 'Файл')}", command=lambda a=attachment: self._open_attachment(a)).pack(side="left", padx=(0, 6), pady=2)

    def _open_attachment(self, attachment: dict) -> None:
        path = ROOT / "data" / attachment["path"]
        if path.exists():
            open_document(path)
        else:
            messagebox.showerror("Файл не знайдено", str(path))

    def _attach(self) -> None:
        paths = filedialog.askopenfilenames()
        for name in paths:
            if not self.draft_note_id:
                self.draft_note_id = __import__("uuid").uuid4().hex
            self.note_attachments.append(self.store.attach(self.draft_note_id, Path(name)))
        self._render_attachments()

    def _toggle_note_pin(self) -> None:
        self.note_pinned = not self.note_pinned
        self.pin_note_button.config(text="★ Закріплено" if self.note_pinned else "☆ Закріпити")

    def _save_note(self) -> None:
        saved = self.store.save(self.draft_note_id, self.note_title.get(), self.note_body.get("1.0", "end-1c"), self.note_attachments, self.note_pinned)
        self.current_note = saved
        self.draft_note_id = saved["id"]
        self._mark_recent("recent_notes", saved["id"])
        self._refresh_notes()
        messagebox.showinfo("Збережено", "Нотатку збережено локально")

    def _delete_note(self) -> None:
        if self.current_note and messagebox.askyesno("Видалити нотатку?", "Цю дію неможливо скасувати."):
            self.store.delete(self.current_note["id"])
            self._refresh_notes()
            self._load_note(None)

    def show_excel(self) -> None:
        self.clear()
        self.header("Archicad / Excel", "Обмін елементами та специфікаціями через XLSX")
        box = tk.Frame(self.main, bg=self.c["panel"], highlightbackground=self.c["border"], highlightthickness=1)
        box.pack(fill="x", padx=34, pady=10)
        tk.Label(box, text="Підключення Archicad", bg=self.c["panel"], fg=self.c["text"], font=("Segoe UI Semibold", 14)).pack(anchor="w", padx=22, pady=(18, 4))
        self.ac_status = tk.Label(box, text="● Не перевірено", bg=self.c["panel"], fg=self.c["muted"], font=("Segoe UI", 10))
        self.ac_status.pack(anchor="w", padx=22, pady=(0, 12))
        ttk.Button(box, text="Перевірити підключення", command=self._connect).pack(anchor="w", padx=22, pady=(0, 18))
        grid = tk.Frame(self.main, bg=self.c["bg"])
        grid.pack(fill="both", expand=True, padx=34, pady=10)
        cards = [("Усі елементи", "all", "Базова таблиця елементів проєкту"), ("Відомість дверей", "doors", "Розміри, ID та властивості дверей"), ("Відомість вікон", "windows", "Розміри, ID та властивості вікон")]
        for i, (title, kind, description) in enumerate(cards):
            card = tk.Frame(grid, bg=self.c["panel"], highlightbackground=self.c["border"], highlightthickness=1)
            card.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else 7, 0 if i == 2 else 7))
            tk.Label(card, text=title, bg=self.c["panel"], fg=self.c["text"], font=("Segoe UI Semibold", 14)).pack(anchor="w", padx=18, pady=(20, 6))
            tk.Label(card, text=description, bg=self.c["panel"], fg=self.c["muted"], wraplength=240, justify="left").pack(anchor="w", padx=18, pady=(0, 20))
            ttk.Button(card, text="Експорт XLSX", style="Accent.TButton", command=lambda k=kind, t=title: self._export(k, t)).pack(fill="x", padx=18, pady=5)
            ttk.Button(card, text="Імпорт XLSX", command=self._import).pack(fill="x", padx=18, pady=(5, 20))
            grid.columnconfigure(i, weight=1)

    def _connect(self) -> None:
        ok = self.adapter.connect()
        self.ac_status.config(text="● Підключено" if ok else "● Archicad недоступний — демо-режим", fg="#34a853" if ok else "#d99228")

    def _export(self, kind: str, title: str) -> None:
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")], initialfile=f"{kind}.xlsx")
        if path:
            try:
                rows = self.adapter.get_elements(kind) if self.adapter.connection else demo_rows(kind)
                export_rows(Path(path), rows, title)
                messagebox.showinfo("Готово", f"Експортовано рядків: {len(rows)}")
            except Exception as error:
                messagebox.showerror("Помилка експорту", str(error))

    def _import(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if path:
            try:
                rows = import_rows(Path(path))
                count = self.adapter.update_elements(rows) if self.adapter.connection else len(rows)
                messagebox.showinfo("Готово", f"Оброблено елементів: {count}")
            except Exception as error:
                messagebox.showerror("Помилка імпорту", str(error))


def main() -> None:
    App().mainloop()
