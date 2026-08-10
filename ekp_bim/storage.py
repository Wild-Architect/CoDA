from __future__ import annotations

import json
import shutil
import uuid
from datetime import datetime
from pathlib import Path


class NoteStore:
    def __init__(self, data_dir: Path):
        self.data_dir = data_dir
        self.notes_dir = data_dir / "notes"
        self.attachments_dir = data_dir / "attachments"
        self.notes_dir.mkdir(parents=True, exist_ok=True)
        self.attachments_dir.mkdir(parents=True, exist_ok=True)

    def list_notes(self) -> list[dict]:
        notes = []
        for path in self.notes_dir.glob("*.json"):
            try:
                notes.append(json.loads(path.read_text(encoding="utf-8")))
            except (OSError, json.JSONDecodeError):
                continue
        return sorted(notes, key=lambda n: (n.get("pinned", False), n.get("updated_at", "")), reverse=True)

    def save(
        self, note_id: str | None, title: str, body: str, attachments: list[dict], pinned: bool = False
    ) -> dict:
        note_id = note_id or uuid.uuid4().hex
        now = datetime.now().isoformat(timespec="seconds")
        path = self.notes_dir / f"{note_id}.json"
        created = now
        if path.exists():
            try:
                created = json.loads(path.read_text(encoding="utf-8")).get("created_at", now)
            except (OSError, json.JSONDecodeError):
                pass
        note = {"id": note_id, "title": title.strip() or "Без назви", "body": body,
                "attachments": attachments, "pinned": pinned, "created_at": created, "updated_at": now}
        path.write_text(json.dumps(note, ensure_ascii=False, indent=2), encoding="utf-8")
        return note

    def attach(self, note_id: str | None, source: Path) -> dict:
        note_id = note_id or uuid.uuid4().hex
        folder = self.attachments_dir / note_id
        folder.mkdir(parents=True, exist_ok=True)
        target = folder / source.name
        if target.exists():
            target = folder / f"{target.stem}_{uuid.uuid4().hex[:6]}{target.suffix}"
        shutil.copy2(source, target)
        return {"name": source.name, "path": str(target.relative_to(self.data_dir))}

    def delete(self, note_id: str) -> None:
        path = self.notes_dir / f"{note_id}.json"
        if path.exists():
            path.unlink()
