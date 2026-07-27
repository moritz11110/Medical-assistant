"""Datei-Synchronisierung aus OneDrive vor dem App-Start."""

from __future__ import annotations

import fnmatch
import hashlib
import json
import os
import shutil
import sys
from datetime import datetime
from pathlib import Path

from config import lade_config


def _standard_updater_config() -> dict:
    return {
        "aktiv": True,
        "onedrive_quelle": "",
        "lokales_ziel": "",
        "status_datei": "logs/updater_state.json",
        "whitelist": ["main.py", "config.py", "updater.py", "start_praxis.bat", "agent/", "gui/", "tools/", "assets/", "start_praxis.vbs"],
        "ausschluesse": [
            "config.json",
            "logs/*",
            "backups/*",
            ".pytest_cache/*",
            "agent/word_search_index.json",
            "test*.py",
        ],
    }


def _hole_updater_config(config: dict) -> dict:
    standard = _standard_updater_config()
    roh = config.get("updater", {})
    if not isinstance(roh, dict):
        return standard
    gemischt = dict(standard)
    gemischt.update(roh)
    return gemischt


def _normalisiere_relativ(relativ: Path) -> str:
    return relativ.as_posix().lstrip("./")


def _ist_relativ_gueltig(eintrag: str) -> bool:
    text = str(eintrag or "").replace("\\", "/").strip().lstrip("./")
    if not text or ":" in text or text.startswith("/"):
        return False
    teile = [teil for teil in text.split("/") if teil]
    return all(teil not in {"..", "."} for teil in teile)


def _ist_ausgeschlossen(relativ: str, ausschluesse: list[str]) -> bool:
    wert = relativ.replace("\\", "/")
    for muster in ausschluesse:
        m = str(muster or "").replace("\\", "/")
        if not m:
            continue
        if m.endswith("/") and (wert == m[:-1] or wert.startswith(m)):
            return True
        if fnmatch.fnmatch(wert, m):
            return True
    return False


def _sammle_quell_dateien(quellbasis: Path, whitelist: list[str], ausschluesse: list[str]) -> list[Path]:
    dateien: list[Path] = []
    for eintrag in whitelist:
        if not _ist_relativ_gueltig(str(eintrag or "")):
            continue
        rel_text = str(eintrag).replace("\\", "/").lstrip("./")
        rel_pfad = Path(rel_text.rstrip("/"))
        start = quellbasis / rel_pfad
        if start.is_file():
            if not _ist_ausgeschlossen(_normalisiere_relativ(rel_pfad), ausschluesse):
                dateien.append(start)
            continue
        if not start.is_dir():
            continue
        for kandidat in start.rglob("*"):
            if not kandidat.is_file():
                continue
            rel = _normalisiere_relativ(kandidat.relative_to(quellbasis))
            if not _ist_ausgeschlossen(rel, ausschluesse):
                dateien.append(kandidat)
    return sorted(set(dateien))


def _sha256_datei(datei: Path) -> str:
    hasher = hashlib.sha256()
    with datei.open("rb") as handle:
        while True:
            block = handle.read(1024 * 1024)
            if not block:
                break
            hasher.update(block)
    return hasher.hexdigest()


def _lade_status(status_datei: Path) -> dict:
    try:
        text = status_datei.read_text(encoding="utf-8")
        daten = json.loads(text)
    except Exception:
        return {"dateien": {}}
    if isinstance(daten, dict) and isinstance(daten.get("dateien"), dict):
        return daten
    return {"dateien": {}}


def _speichere_status(status_datei: Path, datei_hashes: dict[str, str]) -> dict:
    try:
        status_datei.parent.mkdir(parents=True, exist_ok=True)
        payload = {
            "aktualisiert_um": datetime.now().isoformat(timespec="seconds"),
            "dateien": datei_hashes,
        }
        status_datei.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        return {"erfolg": True}
    except Exception:
        return {"fehler": "status_datei_konnte_nicht_gespeichert_werden"}


def _schreibe_laufstatus(status_datei: Path, payload: dict) -> dict:
    try:
        status_datei.parent.mkdir(parents=True, exist_ok=True)
        alt = _lade_status(status_datei)
        neu = dict(alt)
        neu.update(payload)
        neu["aktualisiert_um"] = datetime.now().isoformat(timespec="seconds")
        status_datei.write_text(json.dumps(neu, ensure_ascii=False, indent=2), encoding="utf-8")
        return {"erfolg": True}
    except Exception:
        return {"fehler": "laufstatus_konnte_nicht_gespeichert_werden"}


def _status_datei_pfad(cfg: dict, ziel: Path) -> Path:
    basis = Path(__file__).resolve().parent
    relativ = str(cfg.get("status_datei", "logs/updater_state.json")).replace("\\", "/").strip()
    if _ist_relativ_gueltig(relativ) and ziel.is_dir():
        return ziel / relativ
    if _ist_relativ_gueltig(relativ):
        return basis / relativ
    return basis / "logs" / "updater_state.json"


def _setup_check(cfg: dict) -> dict:
    quelle = Path(str(cfg.get("onedrive_quelle", ""))).expanduser()
    ziel = Path(str(cfg.get("lokales_ziel", ""))).expanduser()
    return {
        "aktiv": bool(cfg.get("aktiv", True)),
        "onedrive_quelle": str(quelle),
        "lokales_ziel": str(ziel),
        "quelle_ok": quelle.is_dir(),
        "ziel_ok": ziel.is_dir(),
        "whitelist_anzahl": len(list(cfg.get("whitelist", []))),
        "ausschluesse_anzahl": len(list(cfg.get("ausschluesse", []))),
    }


def _kopiere_sicher(quell: Path, ziel: Path) -> dict:
    temp = ziel.with_suffix(ziel.suffix + ".tmp_sync")
    try:
        ziel.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(quell, temp)
        os.replace(temp, ziel)
        return {"erfolg": True}
    except Exception:
        try:
            if temp.exists():
                temp.unlink()
        except Exception:
            pass
        return {"fehler": "datei_konnte_nicht_kopiert_werden"}


def _sync_dateien(quellbasis: Path, zielbasis: Path, cfg: dict) -> dict:
    ausschluesse = list(cfg.get("ausschluesse", []))
    whitelist = list(cfg.get("whitelist", []))
    status_datei = zielbasis / str(cfg.get("status_datei", "logs/updater_state.json"))
    status_alt = _lade_status(status_datei).get("dateien", {})
    status_neu: dict[str, str] = {}
    kopiert, uebersprungen, fehler = 0, 0, 0
    for quelle in _sammle_quell_dateien(quellbasis, whitelist, ausschluesse):
        rel = _normalisiere_relativ(quelle.relative_to(quellbasis))
        hash_quelle = _sha256_datei(quelle)
        status_neu[rel] = hash_quelle
        ziel = zielbasis / rel
        if status_alt.get(rel) == hash_quelle and ziel.exists():
            uebersprungen += 1
            continue
        if ziel.exists() and _sha256_datei(ziel) == hash_quelle:
            uebersprungen += 1
            continue
        if _kopiere_sicher(quelle, ziel).get("erfolg"):
            kopiert += 1
        else:
            fehler += 1
    _speichere_status(status_datei, status_neu)
    return {"kopiert": kopiert, "uebersprungen": uebersprungen, "fehler": fehler}


def run_updater() -> dict:
    try:
        config = lade_config()
        cfg = _hole_updater_config(config)
        check = _setup_check(cfg)
        ziel = Path(str(cfg.get("lokales_ziel", ""))).expanduser()
        status_datei = _status_datei_pfad(cfg, ziel)
        if not cfg.get("aktiv", True):
            _schreibe_laufstatus(status_datei, {"status": "deaktiviert", "setup_check": check})
            return {"erfolg": True, "status": "deaktiviert", "setup_check": check}
        if not check.get("quelle_ok") or not check.get("ziel_ok"):
            _schreibe_laufstatus(status_datei, {"status": "uebersprungen", "grund": "quelle_oder_ziel_unzulaessig", "setup_check": check})
            return {"erfolg": True, "status": "uebersprungen", "grund": "quelle_oder_ziel_unzulaessig", "setup_check": check}
        quelle = Path(str(cfg.get("onedrive_quelle", ""))).expanduser()
        ergebnis = _sync_dateien(quelle, ziel, cfg)
        _schreibe_laufstatus(status_datei, {"status": "fertig", **ergebnis, "setup_check": check})
        return {"erfolg": True, "status": "fertig", **ergebnis, "setup_check": check}
    except Exception:
        return {"erfolg": True, "status": "uebersprungen", "grund": "unerwarteter_updater_fehler"}


def main() -> int:
    cfg = _hole_updater_config(lade_config())
    if len(sys.argv) > 1 and sys.argv[1].lower() == "--check":
        print(json.dumps(_setup_check(cfg), ensure_ascii=False, indent=2))
        return 0
    run_updater()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
