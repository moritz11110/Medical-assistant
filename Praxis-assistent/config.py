"""Konfigurationszugriff fuer den PraxisAssistenten."""

from __future__ import annotations

import json
import os
from pathlib import Path

KEYRING_SERVICE = "PraxisAssistent"
KEYRING_USERNAME = "mistral_api_key"

STANDARD_CONFIG = {
    "app_title": "PraxisAssistent",
    "hotkey": "ctrl+alt+p",
    "feature_flags": {
        "word_search_index": True,
    },
    "hintergrundbetrieb": {
        "autostart_aktiv": False,
        "manuelle_updates": True,
    },
    "ki_freigabe": {
        "modus": "konfiguration",
        "allgemein": True,
        "termin": True,
        "mail": True,
    },
    "updater": {
        "aktiv": True,
        "onedrive_quelle": "",
        "lokales_ziel": "",
        "status_datei": "logs/updater_state.json",
        "whitelist": [
            "main.py",
            "config.py",
            "updater.py",
            "start_praxis.bat",
            "start_praxis.vbs",
            "agent/",
            "gui/",
            "tools/",
            "assets/",
        ],
        "ausschluesse": [
            "config.json",
            "logs/*",
            "backups/*",
            ".pytest_cache/*",
            "agent/word_search_index.json",
            "test*.py",
        ],
    },
}


def _lade_json_datei(datei_pfad: Path) -> dict:
    try:
        inhalt = datei_pfad.read_text(encoding="utf-8")
        daten = json.loads(inhalt)
    except (OSError, json.JSONDecodeError):
        return {}
    if isinstance(daten, dict):
        return daten
    return {}


def _kombiniere_standardwerte(rohe_config: dict) -> dict:
    config = dict(STANDARD_CONFIG)
    config.update(rohe_config)
    standard_flags = STANDARD_CONFIG.get("feature_flags", {})
    rohe_flags = rohe_config.get("feature_flags", {})
    if isinstance(standard_flags, dict) and isinstance(rohe_flags, dict):
        gemergte_flags = dict(standard_flags)
        gemergte_flags.update(rohe_flags)
        config["feature_flags"] = gemergte_flags
    elif isinstance(standard_flags, dict):
        config["feature_flags"] = dict(standard_flags)
    _merge_updater_config(config, rohe_config)
    _merge_hintergrund_config(config, rohe_config)
    _merge_ki_freigabe(config, rohe_config)
    return config


def _merge_updater_config(config: dict, rohe_config: dict) -> None:
    standard = STANDARD_CONFIG.get("updater", {})
    roh = rohe_config.get("updater", {})
    if not isinstance(standard, dict):
        return
    if not isinstance(roh, dict):
        config["updater"] = dict(standard)
        return
    gemischt = dict(standard)
    gemischt.update(roh)
    config["updater"] = gemischt


def _merge_hintergrund_config(config: dict, rohe_config: dict) -> None:
    standard = STANDARD_CONFIG.get("hintergrundbetrieb", {})
    roh = rohe_config.get("hintergrundbetrieb", {})
    if not isinstance(standard, dict):
        return
    if not isinstance(roh, dict):
        config["hintergrundbetrieb"] = dict(standard)
        return
    gemischt = dict(standard)
    gemischt.update(roh)
    config["hintergrundbetrieb"] = gemischt


def _merge_ki_freigabe(config: dict, rohe_config: dict) -> None:
    standard = STANDARD_CONFIG.get("ki_freigabe", {})
    roh = rohe_config.get("ki_freigabe", {})
    if not isinstance(standard, dict):
        return
    if not isinstance(roh, dict):
        config["ki_freigabe"] = dict(standard)
        return
    gemischt = dict(standard)
    gemischt.update(roh)
    config["ki_freigabe"] = gemischt


def lade_config() -> dict:
    datei_pfad = Path(__file__).with_name("config.json")
    rohe_config = _lade_json_datei(datei_pfad)
    config = _kombiniere_standardwerte(rohe_config)
    if rohe_config:
        config["_fehler"] = ""
    else:
        config["_fehler"] = "config.json fehlt, ist ungueltig oder leer"
    return config


def speichere_config(config: dict) -> dict:
    datei_pfad = Path(__file__).with_name("config.json")
    daten = dict(config)
    daten.pop("_fehler", None)
    daten.pop("mistral_api_key", None)
    try:
        text = json.dumps(daten, ensure_ascii=False, indent=2)
        datei_pfad.write_text(text, encoding="utf-8")
        return {"erfolg": True}
    except OSError:
        return {"fehler": "config.json konnte nicht gespeichert werden"}


def _hole_keyring_modul():
    try:
        import keyring  # type: ignore

        return keyring
    except Exception:
        return None


def _lade_api_key_aus_keyring() -> str:
    keyring_modul = _hole_keyring_modul()
    if keyring_modul is None:
        return ""
    try:
        wert = keyring_modul.get_password(KEYRING_SERVICE, KEYRING_USERNAME)
        return str(wert or "").strip()
    except Exception:
        return ""


def speichere_mistral_api_key(api_key: str) -> dict:
    keyring_modul = _hole_keyring_modul()
    if keyring_modul is None:
        return {"fehler": "keyring ist nicht verfuegbar"}
    key_text = str(api_key or "").strip()
    if not key_text:
        return {"fehler": "API-Key ist leer"}
    try:
        keyring_modul.set_password(KEYRING_SERVICE, KEYRING_USERNAME, key_text)
        return {"erfolg": True}
    except Exception:
        return {"fehler": "API-Key konnte nicht sicher gespeichert werden"}


def loesche_mistral_api_key() -> dict:
    keyring_modul = _hole_keyring_modul()
    if keyring_modul is None:
        return {"erfolg": True}
    try:
        keyring_modul.delete_password(KEYRING_SERVICE, KEYRING_USERNAME)
        return {"erfolg": True}
    except Exception:
        return {"erfolg": True}


def hole_mistral_api_key(config: dict) -> str:
    env_key = str(os.environ.get("PRAXIS_MISTRAL_API_KEY", "")).strip()
    if env_key:
        return env_key
    return _lade_api_key_aus_keyring()
