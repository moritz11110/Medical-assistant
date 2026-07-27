"""Backup-Helfer fuer Word-Dateien."""

from __future__ import annotations

import shutil
import sys
from datetime import datetime
from pathlib import Path

projekt_ordner = Path(__file__).resolve().parent.parent
if str(projekt_ordner) not in sys.path:
    sys.path.insert(0, str(projekt_ordner))

from config import lade_config


def _projekt_ordner() -> Path:
    return Path(__file__).resolve().parent.parent


def _ist_im_projekt(pfad: Path) -> bool:
    try:
        pfad.resolve().relative_to(_projekt_ordner())
        return True
    except ValueError:
        return False


def _hole_backup_ordner() -> Path:
    config = lade_config()
    pfade = config.get("pfade", {})
    backup_ordner = str(pfade.get("backup_ordner", "") or "").strip()
    if backup_ordner:
        return Path(backup_ordner)
    return _projekt_ordner() / "backups"


def _baue_backup_dateiname(datei: Path) -> str:
    zeitstempel = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"{datei.stem}_backup_{zeitstempel}{datei.suffix}"


def create_backup(filepath: str) -> str:
    try:
        quelle = Path(filepath)
        if not quelle.exists() or not quelle.is_file():
            return ""
        ziel_ordner = _hole_backup_ordner()
        ziel_ordner.mkdir(parents=True, exist_ok=True)
        ziel = ziel_ordner / _baue_backup_dateiname(quelle)
        shutil.copy2(quelle, ziel)
        return str(ziel)
    except Exception:
        return ""


def _ist_docx_backup(datei: Path) -> bool:
    if not datei.is_file() or datei.suffix.lower() != ".docx":
        return False
    return "_backup_" in datei.stem.lower()


def loesche_docx_backups() -> dict:
    try:
        ordner = _hole_backup_ordner()
        if not ordner.exists() or not ordner.is_dir():
            return {"erfolg": True, "geloescht": 0}
        geloescht = 0
        for datei in ordner.iterdir():
            if not _ist_docx_backup(datei):
                continue
            datei.unlink()
            geloescht += 1
        return {"erfolg": True, "geloescht": geloescht}
    except Exception:
        return {"fehler": "Backup-Ordner konnte nicht bereinigt werden"}


def loesche_alle_docx_dateien() -> dict:
    try:
        ordner = _hole_backup_ordner()
        if not ordner.exists() or not ordner.is_dir():
            return {"erfolg": True, "geloescht": 0}
        geloescht = 0
        for datei in ordner.iterdir():
            if not datei.is_file() or datei.suffix.lower() != ".docx":
                continue
            datei.unlink()
            geloescht += 1
        return {"erfolg": True, "geloescht": geloescht}
    except Exception:
        return {"fehler": "Word-Dateien konnten nicht geloescht werden"}
