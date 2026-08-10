"""Boundary for the real Archicad Python API integration.

Drop existing project-specific API code into this adapter and implement the three
methods below. The UI and Excel service deliberately do not depend on Archicad.
"""
from __future__ import annotations


class ArchicadAdapter:
    def __init__(self):
        self.connection = None

    def connect(self) -> bool:
        try:
            from archicad import ACConnection
            self.connection = ACConnection.connect()
        except (ImportError, RuntimeError):
            self.connection = None
        return self.connection is not None

    def get_elements(self, kind: str = "all") -> list[dict]:
        if not self.connection:
            raise RuntimeError("Archicad не підключено")
        raise NotImplementedError("Підключіть наявне зіставлення властивостей Archicad у archicad_adapter.py")

    def update_elements(self, rows: list[dict]) -> int:
        if not self.connection:
            raise RuntimeError("Archicad не підключено")
        raise NotImplementedError("Підключіть запис властивостей Archicad у archicad_adapter.py")

