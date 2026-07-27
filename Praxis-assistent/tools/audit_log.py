"""Audit-Logging ohne Patiententext fuer Nachweiszwecke."""

from __future__ import annotations

import getpass
import hashlib
import json
import os
from datetime import datetime
from pathlib import Path
from uuid import uuid4

PROJEKT_ORDNER = Path(__file__).resolve().parent.parent
LOG_ORDNER = PROJEKT_ORDNER / "logs"
SESSION_ID = uuid4().hex


def _zeitstempel() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def _audit_pfad() -> Path:
    datum = datetime.now().strftime("%Y-%m-%d")
    return LOG_ORDNER / f"audit-{datum}.log"


def _state_pfad() -> Path:
    return LOG_ORDNER / "audit_state.json"


def _sha256_text(text: str) -> str:
    roh = str(text or "").encode("utf-8")
    return hashlib.sha256(roh).hexdigest()


def neue_request_id() -> str:
    return uuid4().hex


def hash_file_ref(datei_ref: str) -> str:
    return _sha256_text(str(datei_ref or ""))


def query_fingerprint(text: str) -> str:
    return _sha256_text(str(text or ""))


def _windows_user() -> str:
    domain = str(os.environ.get("USERDOMAIN", "")).strip()
    user = str(getpass.getuser() or "").strip()
    if domain and user:
        return f"{domain}\\{user}"
    return user


def _lade_state() -> dict:
    try:
        return json.loads(_state_pfad().read_text(encoding="utf-8"))
    except Exception:
        return {}


def _lade_prev_hash() -> str:
    return str(_lade_state().get("last_hash", ""))


def _speichere_prev_hash(hashwert: str) -> None:
    state = {"last_hash": hashwert, "updated_at": _zeitstempel()}
    text = json.dumps(state, ensure_ascii=False)
    _state_pfad().write_text(text, encoding="utf-8")


def _entry_hash(prev_hash: str, eintrag: dict) -> str:
    payload = {"prev_hash": prev_hash, "entry": eintrag}
    roh = json.dumps(payload, ensure_ascii=False, sort_keys=True)
    return _sha256_text(roh)


def _meta_whitelist(meta: dict | None) -> dict:
    if not isinstance(meta, dict):
        return {}
    erlaubt = [
        "tool_name",
        "file_ref_hash",
        "stage",
        "match_detail",
        "exception",
        "mail_count",
        "neue_mails",
        "treffer",
        "roh_treffer",
        "ungueltig_treffer",
        "quelle",
        "policy_decision",
        "block_reason",
        "data_classes",
        "modell",
        "intensiv_modus",
        "timeout_s",
        "versuch",
        "retry",
        "ersetzen",
        "write_error_code",
        "tool_error_code",
        "tool_arg_keys",
        "live_error",
        "getobj_error",
        "temp_error",
        "com_error",
    ]
    sauber = {}
    zahl_keys = {"mail_count", "neue_mails", "treffer", "roh_treffer", "ungueltig_treffer"}
    for key in erlaubt:
        if key not in meta:
            continue
        wert = meta.get(key)
        if key in zahl_keys:
            try:
                sauber[key] = int(wert)
            except Exception:
                continue
            continue
        if key == "data_classes":
            if not isinstance(wert, list):
                continue
            klassen = []
            for eintrag in wert:
                text = str(eintrag or "").strip().lower()
                if text and text not in klassen:
                    klassen.append(text[:40])
            if klassen:
                sauber[key] = klassen[:10]
            continue
        if isinstance(wert, bool):
            sauber[key] = str(wert).lower()
            continue
        text = str(wert).strip()
        if text:
            sauber[key] = text[:160]
    return sauber


def _basis_eintrag(
    action: str,
    request_id: str,
    result: str,
    error_code: str,
    duration_ms: int,
    meta: dict,
    query_hash: str,
) -> dict:
    eintrag = {
        "timestamp": _zeitstempel(),
        "request_id": str(request_id or ""),
        "windows_user": _windows_user(),
        "session_id": SESSION_ID,
        "action": str(action or "unbekannt"),
        "result": "error" if str(result or "ok").lower() == "error" else "ok",
        "error_code": str(error_code or ""),
        "duration_ms": max(0, int(duration_ms or 0)),
        "meta": meta,
    }
    if query_hash:
        eintrag["query_fingerprint"] = query_hash
    return eintrag


def schreibe_audit_event(
    action: str,
    request_id: str = "",
    result: str = "ok",
    error_code: str = "",
    duration_ms: int = 0,
    meta: dict | None = None,
    query_text: str = "",
) -> dict:
    try:
        LOG_ORDNER.mkdir(parents=True, exist_ok=True)
        meta_block = _meta_whitelist(meta)
        query_hash = query_fingerprint(query_text) if query_text else ""
        eintrag = _basis_eintrag(
            action,
            request_id,
            result,
            error_code,
            duration_ms,
            meta_block,
            query_hash,
        )
        prev_hash = _lade_prev_hash()
        eintrag["prev_hash"] = prev_hash
        eintrag["entry_hash"] = _entry_hash(prev_hash, eintrag)
        zeile = json.dumps(eintrag, ensure_ascii=False)
        with _audit_pfad().open("a", encoding="utf-8") as datei:
            datei.write(zeile + "\n")
        _speichere_prev_hash(str(eintrag.get("entry_hash", "")))
        return {"erfolg": True, "pfad": str(_audit_pfad())}
    except Exception:
        return {"fehler": "Audit-Log konnte nicht geschrieben werden"}


def lese_audit_log(limit: int = 500) -> dict:
    try:
        pfad = _audit_pfad()
        if not pfad.exists():
            return {"eintraege": [], "pfad": str(pfad)}
        zeilen = pfad.read_text(encoding="utf-8").splitlines()
        roh_liste = zeilen[-max(1, int(limit)):]
        eintraege = []
        for zeile in roh_liste:
            try:
                daten = json.loads(zeile)
            except Exception:
                continue
            if isinstance(daten, dict):
                eintraege.append(daten)
        return {"eintraege": eintraege, "pfad": str(pfad)}
    except Exception:
        return {"fehler": "Audit-Log konnte nicht gelesen werden"}
