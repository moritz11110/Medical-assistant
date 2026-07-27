"""Agent-Logik fuer Tool-Use mit Mistral Function Calling."""

from __future__ import annotations

import json
import re
import sys
import time
import unicodedata
from datetime import datetime
from pathlib import Path

projekt_ordner = Path(__file__).resolve().parent.parent
if str(projekt_ordner) not in sys.path:
    sys.path.insert(0, str(projekt_ordner))

from config import hole_mistral_api_key, lade_config
from tools.outlook_tools import get_recent_mails, search_mails
from tools.audit_log import hash_file_ref, schreibe_audit_event
from tools.word_tools import (
    find_file_by_date,
    find_free_slots,
    list_files_in_folder,
    read_word_file,
    search_appointments,
    search_docx_files,
    write_appointment,
)

SYSTEM_PROMPT = (
    "Du bist ein Praxis-Assistent. Arbeite praezise und knapp. "
    "Nutze Tools, um Fakten zu pruefen. Schreibe niemals autonom in Dateien. "
    "Bei write_appointment nur Vorschlag liefern und bestaetigungspflichtig markieren. "
    "Wenn der Nutzer ausdruecklich ersetzen/ueberschreiben sagt, setze ersetzen=true. "
    "Wenn ein Datum mit Tag, Monat und Jahr erkennbar ist, normalisiere es intern und frage nicht nach einem strengen Datumsformat. "
    "Nutze fuer relative Zeitangaben bevorzugt get_current_datetime. "
    "Bei Suchen ueber mehrere Word-Dateien nutze search_appointments. "
    "Bei Aenderungen bestehender Termine nur explizit genannte Felder aendern und alle anderen Felder unveraendert lassen. "
    "Leere Felder niemals stillschweigend loeschen; vor dem Loeschen immer explizit rueckfragen."
)

DATUM_PATTERN = re.compile(r"\b\d{1,2}\s*[.\-_/ ]\s*\d{1,2}\s*[.\-_/ ]\s*\d{4}\b")
ERSETZ_PATTERN = re.compile(
    r"\b(ersetz\w*|ueberschreib\w*|überschreib\w*|aender\w*|änder\w*|umbenenn\w*|benenn\w*)\b"
)
NEGATION_PATTERN = re.compile(r"\b(nicht|kein(?:e|en|er|em|es)?|ohne)\b")
LESEN_VERB_PATTERN = re.compile(r"\b(oeffn\w*|öffn\w*|zeig\w*|lies\w*|lese\w*|anzeig\w*)\b")
DATEI_BEZUG_PATTERN = re.compile(r"\b(datei\w*|docx|terminliste\w*|inhalt\w*)\b")
DOCX_NAME_PATTERN = re.compile(r"[\w .-]+\.docx\b", flags=re.IGNORECASE)
SCHREIB_FELDER = [
    "filepath",
    "uhrzeit",
    "name",
    "svnr",
    "geburtsdatum",
    "adresse",
    "firma",
    "untersuchungsart",
    "vgue",
    "ersetzen",
]

TOOL_SCHEMAS = [
    {"type": "function", "function": {"name": "list_files_in_folder", "description": "Listet Terminlisten-Dateien", "parameters": {"type": "object", "properties": {}, "additionalProperties": False}}},
    {"type": "function", "function": {"name": "search_docx_files", "description": "Sucht .docx Dateien in erlaubten Ordnern", "parameters": {"type": "object", "properties": {"dateiname": {"type": "string"}, "datum": {"type": "string"}}, "additionalProperties": False}}},
    {"type": "function", "function": {"name": "search_appointments", "description": "Sucht Termine in allen erlaubten Word-Dateien", "parameters": {"type": "object", "properties": {"suchtext": {"type": "string"}, "uhrzeit": {"type": "string"}, "datum": {"type": "string"}}, "additionalProperties": False}}},
    {"type": "function", "function": {"name": "read_word_file", "description": "Liest Word-Datei", "parameters": {"type": "object", "properties": {"filepath": {"type": "string"}}, "required": ["filepath"], "additionalProperties": False}}},
    {"type": "function", "function": {"name": "find_free_slots", "description": "Findet freie Slots", "parameters": {"type": "object", "properties": {"filepath": {"type": "string"}}, "required": ["filepath"], "additionalProperties": False}}},
    {"type": "function", "function": {"name": "find_file_by_date", "description": "Sucht Datei nach Datum", "parameters": {"type": "object", "properties": {"datum": {"type": "string"}}, "required": ["datum"], "additionalProperties": False}}},
    {"type": "function", "function": {"name": "write_appointment", "description": "Schreibvorschlag fuer Termin", "parameters": {"type": "object", "properties": {"filepath": {"type": "string"}, "uhrzeit": {"type": "string"}, "name": {"type": "string"}, "svnr": {"type": "string"}, "geburtsdatum": {"type": "string"}, "adresse": {"type": "string"}, "firma": {"type": "string"}, "untersuchungsart": {"type": "string"}, "vgue": {"type": "string"}, "ersetzen": {"type": "boolean"}}, "required": ["filepath", "uhrzeit"], "additionalProperties": False}}},
    {"type": "function", "function": {"name": "get_recent_mails", "description": "Holt letzte Mails", "parameters": {"type": "object", "properties": {"anzahl": {"type": "integer", "minimum": 1}}, "additionalProperties": False}}},
    {"type": "function", "function": {"name": "search_mails", "description": "Sucht Mails", "parameters": {"type": "object", "properties": {"stichwort": {"type": "string"}}, "required": ["stichwort"], "additionalProperties": False}}},
    {"type": "function", "function": {"name": "get_current_datetime", "description": "Liefert aktuelle lokale Zeit und Datum", "parameters": {"type": "object", "properties": {}, "additionalProperties": False}}},
]

LLM_TIMEOUT_STANDARD_S = 35
LLM_TIMEOUT_LARGE_S = 35
LLM_RETRIES = 2
LANGSAMER_LLM_CALL_MS = 15000


def _audit(
    aktion: str,
    request_id: str,
    result: str = "ok",
    error_code: str = "",
    duration_ms: int = 0,
    meta: dict | None = None,
    query_text: str = "",
) -> None:
    schreibe_audit_event(
        action=aktion,
        request_id=request_id,
        result=result,
        error_code=error_code,
        duration_ms=duration_ms,
        meta=meta or {},
        query_text=query_text,
    )


def _meta_fuer_tool(tool_name: str, ergebnis: dict) -> dict:
    meta = {"tool_name": str(tool_name or "")}
    if not isinstance(ergebnis, dict):
        return meta
    pfad = str(ergebnis.get("filepath", "") or ergebnis.get("backup_pfad", "") or "").strip()
    if pfad:
        meta["file_ref_hash"] = hash_file_ref(pfad)
    mails = ergebnis.get("mails")
    if isinstance(mails, list):
        meta["mail_count"] = len(mails)
    return meta


def _data_classes_fuer_tool(tool_name: str) -> list[str]:
    name = str(tool_name or "").strip().lower()
    if "mail" in name:
        return ["mail"]
    if "appointment" in name or "word" in name or "docx" in name or "slot" in name:
        return ["termin"]
    return ["allgemein"]


def _meta_compliance(tool_name: str, policy_decision: str, block_reason: str = "") -> dict:
    meta = {
        "tool_name": str(tool_name or ""),
        "policy_decision": str(policy_decision or "allow").strip().lower(),
        "data_classes": _data_classes_fuer_tool(tool_name),
    }
    grund = str(block_reason or "").strip().lower()
    if grund:
        meta["block_reason"] = grund
    return meta


def _meta_fuer_run_agent(frage: str, policy_decision: str, block_reason: str = "") -> dict:
    text = str(frage or "").strip().lower()
    klassen = ["allgemein"]
    if any(token in text for token in ["mail", "e-mail", "email"]):
        klassen = ["mail"]
    if any(token in text for token in ["termin", "uhrzeit", "docx", "datei"]):
        klassen = ["termin"]
    meta = {
        "tool_name": "run_agent",
        "policy_decision": str(policy_decision or "allow").strip().lower(),
        "data_classes": klassen,
    }
    grund = str(block_reason or "").strip().lower()
    if grund:
        meta["block_reason"] = grund
    return meta


def _modell_fuer_modus(intensiv_modus: bool) -> str:
    if bool(intensiv_modus):
        return "mistral-large-latest"
    return "mistral-medium-latest"


def _mistral_client_und_modell(intensiv_modus: bool):
    config = lade_config()
    api_key = hole_mistral_api_key(config)
    modell = _modell_fuer_modus(intensiv_modus)
    if not api_key:
        return None, "", "Mistral API-Key fehlt in config.json"
    try:
        from openai import OpenAI

        return OpenAI(api_key=api_key, base_url="https://api.mistral.ai/v1", max_retries=0), modell, ""
    except Exception:
        return None, "", "Mistral SDK konnte nicht initialisiert werden"


def _llm_timeout_s(modell: str) -> int:
    name = str(modell or "").strip().lower()
    if "large" in name:
        return LLM_TIMEOUT_LARGE_S
    return LLM_TIMEOUT_STANDARD_S


def _max_retries_fuer_modell(modell: str) -> int:
    name = str(modell or "").strip().lower()
    if "large" in name:
        return 1
    return LLM_RETRIES


def _parse_argumente(argumente_raw: str) -> dict:
    try:
        daten = json.loads(argumente_raw or "{}")
    except Exception:
        return {}
    return daten if isinstance(daten, dict) else {}


def _bool_wert(wert) -> bool:
    if isinstance(wert, bool):
        return wert
    text = str(wert or "").strip().lower()
    return text in {"1", "true", "ja", "yes", "y", "an"}


def _schreibdaten_aus_dict(daten: dict) -> dict:
    schreibdaten = {}
    for feld in SCHREIB_FELDER:
        if feld == "ersetzen":
            schreibdaten[feld] = _bool_wert(daten.get(feld, False))
            continue
        schreibdaten[feld] = str(daten.get(feld, "") or "").strip()
    return schreibdaten


def _schreibdaten_aus_raw(argumente_raw: str) -> dict:
    return _schreibdaten_aus_dict(_parse_argumente(argumente_raw))


def _aktuelle_zeitdaten() -> dict:
    jetzt = datetime.now().astimezone()
    return {
        "lokal": jetzt.strftime("%Y-%m-%d %H:%M:%S"),
        "datum": jetzt.strftime("%Y-%m-%d"),
        "uhrzeit": jetzt.strftime("%H:%M:%S"),
        "wochentag": jetzt.strftime("%A"),
        "zeitzone": str(jetzt.tzinfo or ""),
    }


def _zeit_systemnachricht() -> dict:
    zeit = _aktuelle_zeitdaten()
    inhalt = (
        f"Aktuelle lokale Zeit: {zeit.get('lokal','')}, "
        f"Datum: {zeit.get('datum','')}, Uhrzeit: {zeit.get('uhrzeit','')}, "
        f"Zeitzone: {zeit.get('zeitzone','')}"
    )
    return {"role": "system", "content": inhalt}


def _schreibvorschlag(argumente: dict) -> dict:
    daten = _schreibdaten_aus_dict(argumente if isinstance(argumente, dict) else {})
    leer_hinweis = ""
    if daten.get("ersetzen"):
        leer_hinweis = "Leere Angaben loeschen keine bestehenden Werte; vor Loeschungen wird rueckgefragt."
    teile = [
        f"Datei: {daten.get('filepath', '')}",
        f"Uhrzeit: {daten.get('uhrzeit', '')}",
        f"Name: {daten.get('name', '')}",
        f"SVNr: {daten.get('svnr', '')}",
        f"Geburtsdatum: {daten.get('geburtsdatum', '')}",
        f"Adresse: {daten.get('adresse', '')}",
        f"Firma: {daten.get('firma', '')}",
        f"Untersuchung: {daten.get('untersuchungsart', '')}",
        f"VGUE: {daten.get('vgue', '')}",
        f"Ersetzen: {daten.get('ersetzen', False)}",
    ]
    return {
        "bestaetigung_noetig": True,
        "hinweis": " ".join(["Kein autonomes Schreiben", leer_hinweis]).strip(),
        "vorschlag": "\n".join(teile),
    }


def _tool_funktionen() -> dict:
    return {
        "list_files_in_folder": lambda args: list_files_in_folder(),
        "search_docx_files": lambda args: search_docx_files(str(args.get("dateiname", "")), str(args.get("datum", ""))),
        "search_appointments": lambda args: search_appointments(str(args.get("suchtext", "")), str(args.get("uhrzeit", "")), str(args.get("datum", ""))),
        "read_word_file": lambda args: read_word_file(str(args.get("filepath", ""))),
        "find_free_slots": lambda args: find_free_slots(str(args.get("filepath", ""))),
        "find_file_by_date": lambda args: find_file_by_date(str(args.get("datum", ""))),
        "write_appointment": lambda args: _schreibvorschlag(args if isinstance(args, dict) else {}),
        "get_recent_mails": lambda args: get_recent_mails(args.get("anzahl")),
        "search_mails": lambda args: search_mails(str(args.get("stichwort", ""))),
        "get_current_datetime": lambda args: _aktuelle_zeitdaten(),
    }


def _ist_verbotene_mail_aktion(name: str) -> bool:
    text = str(name or "").strip().lower()
    return any(token in text for token in ["send", "reply", "forward", "create_mail", "compose_mail"])


def _normalisiere_sicherheitstext(text: str) -> str:
    roh = unicodedata.normalize("NFKC", str(text or ""))
    roh = "".join(ch for ch in roh if unicodedata.category(ch) not in {"Cf", "Cc"})
    roh = roh.lower()
    roh = roh.translate(str.maketrans({"0": "o", "1": "i", "3": "e", "4": "a", "5": "s", "7": "t"}))
    roh = roh.replace("ä", "ae").replace("ö", "oe").replace("ü", "ue").replace("ß", "ss")
    roh = re.sub(r"[_\-./\\]+", " ", roh)
    return re.sub(r"\s+", " ", roh).strip()


def _phase_a_tool_risiko(name: str) -> dict:
    normalisiert = _normalisiere_sicherheitstext(name)
    kompakt = re.sub(r"[^a-z0-9]", "", normalisiert)
    tokens = ["send", "reply", "forward", "createmail", "composemail", "antwort", "weiterleit"]
    treffer = [token for token in tokens if token in kompakt]
    return {
        "phase_a_tool_shadow_block": bool(treffer),
        "phase_a_tool_shadow_tokens": ",".join(treffer[:6]),
    }


def _phase_a_nutzer_risiko(frage: str) -> dict:
    normalisiert = _normalisiere_sicherheitstext(frage)
    kompakt = re.sub(r"[^a-z0-9]", "", normalisiert)
    mail_bezug = any(token in normalisiert for token in ["mail", "e mail", "email"])
    send_tokens = ["send", "sende", "schick", "verschick", "versend", "reply", "antwort"]
    konto_bezug = any(token in normalisiert for token in ["konto", "account", "profil"])
    konto_tokens = ["wechsel", "aender", "ander", "setze", "umstell", "switch"]
    send_treffer = [token for token in send_tokens if token in kompakt]
    konto_treffer = [token for token in konto_tokens if token in kompakt]
    return {
        "phase_a_shadow_send_mail": bool(mail_bezug and send_treffer),
        "phase_a_shadow_konto_wechsel": bool(konto_bezug and konto_treffer),
        "phase_a_shadow_send_tokens": ",".join(send_treffer[:6]),
        "phase_a_shadow_konto_tokens": ",".join(konto_treffer[:6]),
    }


def _enthaelt_verbotene_nutzeraktion(frage: str) -> bool:
    text = str(frage or "").strip().lower()
    send_mail = ("mail" in text or "e-mail" in text or "email" in text) and any(
        token in text for token in ["send", "sende", "schick", "verschick"]
    )
    konto_wechsel = ("konto" in text or "account" in text) and any(
        token in text for token in ["wechsel", "aender", "?nder", "setze", "umstell"]
    )
    return send_mail or konto_wechsel


def _tool_arg_keys(argumente_raw: str) -> str:
    daten = _parse_argumente(argumente_raw)
    if not isinstance(daten, dict):
        return ""
    keys = [str(key).strip().lower() for key in daten.keys() if str(key).strip()]
    keys = sorted(list(dict.fromkeys(keys)))
    return ",".join(keys[:12])


def _tool_meta_debug(name: str, argumente_raw: str, ergebnis: dict | None = None) -> dict:
    meta = _meta_compliance(name, "allow")
    keys = _tool_arg_keys(argumente_raw)
    if keys:
        meta["tool_arg_keys"] = keys
    if isinstance(ergebnis, dict):
        code = str(ergebnis.get("fehler_code", "") or "").strip().lower()
        if code:
            meta["tool_error_code"] = code
    return meta


def _ist_temporaerer_word_fehler(ergebnis: dict) -> bool:
    if not isinstance(ergebnis, dict):
        return False
    code = str(ergebnis.get("fehler_code", "") or "").strip().lower()
    text = str(ergebnis.get("fehler", "") or "").strip().lower()
    if code in {"open_retry_failed", "open_exception"}:
        return True
    return "gesperrt" in text or "nicht geoeffnet" in text


def _retry_word_tool(name: str, argumente_raw: str, funktion) -> dict:
    ergebnis = funktion(_parse_argumente(argumente_raw))
    if "word" not in str(name or "").lower() and "appointment" not in str(name or "").lower():
        return ergebnis
    if not _ist_temporaerer_word_fehler(ergebnis if isinstance(ergebnis, dict) else {}):
        return ergebnis
    time.sleep(0.35)
    zweiter_versuch = funktion(_parse_argumente(argumente_raw))
    if isinstance(zweiter_versuch, dict):
        zweiter_versuch["_retry_once"] = True
    return zweiter_versuch


def _fuehre_tool_aus(name: str, argumente_raw: str, request_id: str) -> dict:
    start_ms = int(time.time() * 1000)
    phase_a_risiko = _phase_a_tool_risiko(name)
    if _ist_verbotene_mail_aktion(name):
        meta = _meta_compliance(name, "block", "mail_block")
        meta.update(phase_a_risiko)
        _audit("agent.tool_call", request_id, "error", "mail_block", 0, meta)
        return {"fehler": "Mail-Sendeaktionen sind deaktiviert"}
    funktion = _tool_funktionen().get(name)
    if funktion is None:
        meta = _meta_compliance(name, "block", "unknown_tool")
        meta.update(phase_a_risiko)
        _audit("agent.tool_call", request_id, "error", "unknown_tool", 0, meta)
        return {"fehler": f"Unbekanntes Tool: {name}"}
    try:
        ergebnis = _retry_word_tool(name, argumente_raw, funktion)
    except Exception:
        dauer = max(0, int(time.time() * 1000) - start_ms)
        meta = _tool_meta_debug(name, argumente_raw)
        meta["tool_error_code"] = "tool_exception"
        meta.update(phase_a_risiko)
        _audit("agent.tool_call", request_id, "error", "tool_exception", dauer, meta)
        return {"fehler": f"Tool fehlgeschlagen: {name}"}
    dauer = max(0, int(time.time() * 1000) - start_ms)
    meta = _meta_fuer_tool(name, ergebnis if isinstance(ergebnis, dict) else {})
    meta.update(_tool_meta_debug(name, argumente_raw, ergebnis if isinstance(ergebnis, dict) else None))
    meta.update(phase_a_risiko)
    result = "error" if isinstance(ergebnis, dict) and ergebnis.get("fehler") else "ok"
    code = "tool_return_error" if result == "error" else ""
    if result == "error" and isinstance(ergebnis, dict):
        tool_code = str(ergebnis.get("fehler_code", "") or "").strip().lower()
        if tool_code:
            code = tool_code
    _audit("agent.tool_call", request_id, result, code, dauer, meta)
    return ergebnis if isinstance(ergebnis, dict) else {"ergebnis": ergebnis}


def _assistant_tool_nachricht(nachricht) -> dict:
    def _wert(objekt, feld, standard=None):
        if isinstance(objekt, dict):
            return objekt.get(feld, standard)
        return getattr(objekt, feld, standard)

    def _tool_call_kompakt(call) -> dict:
        funktion = _wert(call, "function", {}) or {}
        name = str(_wert(funktion, "name", "") or "").strip()
        argumente = _wert(funktion, "arguments", "")
        if isinstance(argumente, (dict, list)):
            argumente_text = json.dumps(argumente, ensure_ascii=False)
        else:
            argumente_text = str(argumente or "")
        call_id = str(_wert(call, "id", "") or "").strip()
        return {
            "id": call_id,
            "name": name,
            "arguments": argumente_text,
        }

    inhalt = _wert(nachricht, "content", "")
    if isinstance(inhalt, list):
        inhalt_text = json.dumps(inhalt, ensure_ascii=False)
    else:
        inhalt_text = str(inhalt or "")
    aufrufe = _wert(nachricht, "tool_calls", []) or []
    tool_calls = []
    for index, call in enumerate(aufrufe, start=1):
        kompakt = _tool_call_kompakt(call)
        call_id = kompakt.get("id") or f"tool_call_{index}"
        tool_calls.append({
            "id": call_id,
            "type": "function",
            "function": {
                "name": kompakt.get("name", ""),
                "arguments": kompakt.get("arguments", ""),
            },
        })
    return {"role": "assistant", "content": inhalt_text, "tool_calls": tool_calls}


def _tool_call_werte(tool_call, index: int) -> tuple[str, str, str]:
    def _wert(objekt, feld, standard=None):
        if isinstance(objekt, dict):
            return objekt.get(feld, standard)
        return getattr(objekt, feld, standard)

    funktion = _wert(tool_call, "function", {}) or {}
    tool_name = str(_wert(funktion, "name", "") or "").strip()
    argumente = _wert(funktion, "arguments", "")
    if isinstance(argumente, (dict, list)):
        argumente_raw = json.dumps(argumente, ensure_ascii=False)
    else:
        argumente_raw = str(argumente or "")
    tool_call_id = str(_wert(tool_call, "id", "") or "").strip() or f"tool_call_{index}"
    return tool_name, argumente_raw, tool_call_id


def _hat_tool_calls(nachricht) -> bool:
    tool_calls = getattr(nachricht, "tool_calls", None)
    if tool_calls is None and isinstance(nachricht, dict):
        tool_calls = nachricht.get("tool_calls")
    return bool(tool_calls)


def _inhalt_aus_nachricht(nachricht) -> str:
    inhalt = getattr(nachricht, "content", None)
    if inhalt is None and isinstance(nachricht, dict):
        inhalt = nachricht.get("content")
    if isinstance(inhalt, list):
        return json.dumps(inhalt, ensure_ascii=False)
    return str(inhalt or "")


def _fehlercode_aus_exception(exc: Exception) -> str:
    name = exc.__class__.__name__.lower()
    text = str(exc or "").strip().lower()
    if "timeout" in name or "timeout" in text:
        return "timeout"
    if "ratelimit" in name or "429" in text or "rate limit" in text:
        return "rate_limit"
    if "connection" in name or "connect" in text:
        return "connection"
    if "auth" in name or "401" in text or "unauthorized" in text:
        return "auth"
    return "llm_exception"


def _fehler_meta(exc: Exception, versuch: int) -> dict:
    return {
        "tool_name": "chat_completion",
        "policy_decision": "allow",
        "data_classes": ["allgemein"],
        "block_reason": f"try_{max(1, int(versuch or 1))}",
        "error_type": exc.__class__.__name__,
    }


def _llm_debug_meta(modell: str, timeout_s: int, versuch: int, retry: int, intensiv_modus: bool) -> dict:
    return {
        "tool_name": "chat_completion",
        "modell": str(modell or "").strip().lower(),
        "timeout_s": max(1, int(timeout_s or 1)),
        "versuch": max(1, int(versuch or 1)),
        "retry": max(1, int(retry or 1)),
        "intensiv_modus": bool(intensiv_modus),
    }


def _ist_retry_fehler(code: str) -> bool:
    return str(code or "").strip().lower() in {"timeout", "rate_limit", "connection"}


def _backoff_ms(code: str, retry_nummer: int) -> int:
    retry = max(1, int(retry_nummer or 1))
    faktor = 1300 if str(code or "").strip().lower() == "rate_limit" else 800
    return min(6000, retry * faktor)


def _fehlertext_fuer_code(code: str) -> str:
    mapping = {
        "timeout": "Mistral Anfrage dauerte zu lange (Timeout)",
        "rate_limit": "Mistral Rate-Limit erreicht. Bitte kurz warten und erneut versuchen",
        "connection": "Keine Verbindung zur Mistral API",
        "auth": "Authentifizierung bei Mistral fehlgeschlagen (API-Key prüfen)",
    }
    schluessel = str(code or "").strip().lower()
    return mapping.get(schluessel, "Mistral Anfrage fehlgeschlagen")


def _chat_completion(
    client,
    modell: str,
    nachrichten: list[dict],
    request_id: str,
    versuch: int,
    timeout_s: int,
    retry: int,
    intensiv_modus: bool,
):
    start_ms = int(time.time() * 1000)
    start_meta = _llm_debug_meta(modell, timeout_s, versuch, retry, intensiv_modus)
    _audit("agent.llm_call.start", request_id, "ok", "", 0, start_meta)
    try:
        antwort = client.chat.completions.create(
            model=modell,
            messages=nachrichten,
            tools=TOOL_SCHEMAS,
            tool_choice="auto",
            temperature=0,
            timeout=max(1, int(timeout_s)),
        )
    except Exception as exc:
        dauer = max(0, int(time.time() * 1000) - start_ms)
        code = _fehlercode_aus_exception(exc)
        meta = _fehler_meta(exc, versuch)
        meta.update(_llm_debug_meta(modell, timeout_s, versuch, retry, intensiv_modus))
        _audit("agent.llm_call.end", request_id, "error", code, dauer, meta)
        raise
    dauer = max(0, int(time.time() * 1000) - start_ms)
    if dauer >= LANGSAMER_LLM_CALL_MS:
        langsam_meta = _llm_debug_meta(modell, timeout_s, versuch, retry, intensiv_modus)
        _audit(
            "agent.llm_call.slow",
            request_id,
            "ok",
            "",
            dauer,
            langsam_meta,
        )
    end_meta = _llm_debug_meta(modell, timeout_s, versuch, retry, intensiv_modus)
    _audit("agent.llm_call.end", request_id, "ok", "", dauer, end_meta)
    return antwort


def _verlauf_zu_nachrichten(verlauf: list[dict]) -> list[dict]:
    nachrichten = []
    for eintrag in verlauf:
        if not isinstance(eintrag, dict):
            continue
        rolle = str(eintrag.get("rolle", "")).strip().lower()
        inhalt = str(eintrag.get("inhalt", "") or "").strip()
        if not inhalt:
            continue
        if rolle in {"assistant", "user"}:
            nachrichten.append({"role": rolle, "content": inhalt})
    return nachrichten


def _hole_nachricht(
    client,
    modell: str,
    nachrichten: list[dict],
    request_id: str,
    versuch: int,
    timeout_s: int,
    intensiv_modus: bool,
):
    max_retries = _max_retries_fuer_modell(modell)
    for retry in range(1, max_retries + 1):
        try:
            antwort = _chat_completion(
                client,
                modell,
                nachrichten,
                request_id,
                versuch,
                timeout_s,
                retry,
                intensiv_modus,
            )
            return antwort.choices[0].message, "", ""
        except Exception as exc:
            code = _fehlercode_aus_exception(exc)
            if retry < max_retries and _ist_retry_fehler(code):
                warte_ms = _backoff_ms(code, retry)
                meta = _llm_debug_meta(modell, timeout_s, versuch, retry, intensiv_modus)
                meta["next_wait_ms"] = warte_ms
                meta["error_type"] = exc.__class__.__name__
                _audit("agent.llm_call.retry", request_id, "error", code, 0, meta)
                time.sleep(warte_ms / 1000)
                continue
            meta = _llm_debug_meta(modell, timeout_s, versuch, retry, intensiv_modus)
            meta["policy_decision"] = "allow"
            meta["data_classes"] = ["allgemein"]
            meta["block_reason"] = f"try_{max(1, int(versuch or 1))}"
            _audit("agent.llm_call.error", request_id, "error", code, 0, meta)
            return None, _fehlertext_fuer_code(code), code
    return None, _fehlertext_fuer_code("llm_exception"), "llm_exception"


def _frage_will_ersetzen(frage: str) -> bool:
    text = str(frage or "").strip().lower()
    if not text:
        return False
    return _hat_ersetzsignal_ohne_negation(text)


def _hat_ersetzsignal_ohne_negation(text: str) -> bool:
    for treffer in ERSETZ_PATTERN.finditer(str(text or "")):
        if _ist_negiert_vor_treffer(text, treffer.start()):
            continue
        return True
    return False


def _ist_negiert_vor_treffer(text: str, start_index: int) -> bool:
    links = max(0, int(start_index) - 24)
    ausschnitt = str(text or "")[links:int(start_index)]
    return NEGATION_PATTERN.search(ausschnitt) is not None


def _normalisiere_datumstext(text: str) -> str:
    treffer = re.search(r"(\d{1,2})\s*[.\-_/ ]\s*(\d{1,2})\s*[.\-_/ ]\s*(\d{4})", str(text or ""))
    if treffer is None:
        return ""
    tag, monat, jahr = treffer.groups()
    if not (1 <= int(tag) <= 31 and 1 <= int(monat) <= 12):
        return ""
    return f"{int(tag):02d}.{int(monat):02d}.{jahr}"


def _normalisiere_frage_datum(frage: str) -> str:
    text = str(frage or "")
    norm = _normalisiere_datumstext(text)
    if not norm:
        return text
    return re.sub(r"\b\d{1,2}\s*[.\-_/ ]\s*\d{1,2}\s*[.\-_/ ]\s*\d{4}\b", norm, text, count=1)


def _letztes_datum_aus_verlauf(frage: str, verlauf: list[dict]) -> str:
    quellen = [str(frage or "")]
    for eintrag in verlauf:
        inhalt = str(eintrag.get("inhalt", "") or "").strip()
        if inhalt:
            quellen.append(inhalt)
    for text in reversed(quellen):
        treffer = DATUM_PATTERN.findall(text)
        if treffer:
            return treffer[-1]
    return ""


def _kontext_nachricht_fuer_dateibezug(frage: str, verlauf: list[dict]):
    text = str(frage or "").lower()
    if not any(token in text for token in ["dieser datei", "diese datei", "in der datei", "die datei"]):
        return None
    if DATUM_PATTERN.search(text) or ".docx" in text or "\\" in text or "/" in text:
        return None
    datum = _letztes_datum_aus_verlauf(frage, verlauf)
    if not datum:
        return None
    pfad = str(find_file_by_date(datum).get("filepath", "") or "").strip()
    if not pfad:
        return None
    return {"role": "system", "content": f"Kontext: Mit 'dieser Datei' ist Datei {datum} gemeint. Pfad: {pfad}"}



def _frage_will_datei_lesen(frage: str) -> bool:
    text = str(frage or "").strip().lower()
    if not text:
        return False
    if _enthaelt_negation_global(text):
        return False
    if _enthaelt_docx_name(text):
        return True
    return _hat_lesesignal_ohne_negation(text) and _hat_dateibezug(text)


def _hat_lesesignal_ohne_negation(text: str) -> bool:
    for treffer in LESEN_VERB_PATTERN.finditer(str(text or "")):
        if _ist_negiert_vor_treffer(text, treffer.start()):
            continue
        return True
    return False


def _hat_dateibezug(text: str) -> bool:
    if DATEI_BEZUG_PATTERN.search(str(text or "")) is not None:
        return True
    if DATUM_PATTERN.search(str(text or "")) is not None:
        return True
    return _enthaelt_docx_name(text)


def _enthaelt_docx_name(text: str) -> bool:
    return DOCX_NAME_PATTERN.search(str(text or "")) is not None


def _enthaelt_negation_global(text: str) -> bool:
    return NEGATION_PATTERN.search(str(text or "")) is not None


def _dateiname_aus_text(text: str) -> str:
    treffer = re.search(r"([\w .-]+\.docx)", str(text or ""), flags=re.IGNORECASE)
    if treffer is None:
        return ""
    return str(treffer.group(1) or "").strip()


def _pfad_aus_frage_oder_verlauf(frage: str, verlauf: list[dict]) -> str:
    dateiname = _dateiname_aus_text(frage)
    if dateiname:
        return _pfad_aus_dateiname(dateiname)
    datum = _normalisiere_datumstext(frage)
    if datum:
        return str(find_file_by_date(datum).get("filepath", "") or "").strip()
    letztes = _letztes_datum_aus_verlauf(frage, verlauf)
    if letztes:
        return str(find_file_by_date(letztes).get("filepath", "") or "").strip()
    return ""


def _kontext_nachricht_fuer_dateilesen(frage: str, verlauf: list[dict]):
    if not _frage_will_datei_lesen(frage):
        return None
    pfad = _pfad_aus_frage_oder_verlauf(frage, verlauf)
    if not pfad:
        return None
    inhalt = (
        "Wenn der Nutzer eine Datei oeffnen/anzeigen will, "
        f"nutze read_word_file mit filepath='{pfad}' und antworte anhand des Tool-Ergebnisses."
    )
    return {"role": "system", "content": inhalt}


def _tool_cache_key(tool_name: str, argumente_raw: str) -> str:
    name = str(tool_name or "").strip().lower()
    args = str(argumente_raw or "").strip()
    return f"{name}::{args}"


def _verarbeite_tool_calls(
    nachricht,
    nachrichten: list[dict],
    schreibdaten: dict,
    vorschlag: str,
    ersetzen_erzwingen: bool,
    request_id: str,
    tool_cache: dict,
):
    neue_schreibdaten = schreibdaten
    neuer_vorschlag = vorschlag
    nachrichten.append(_assistant_tool_nachricht(nachricht))
    tool_calls = getattr(nachricht, "tool_calls", None)
    if tool_calls is None and isinstance(nachricht, dict):
        tool_calls = nachricht.get("tool_calls")
    for index, tool_call in enumerate(tool_calls or [], start=1):
        tool_name, argumente_raw, tool_call_id = _tool_call_werte(tool_call, index)
        if not tool_name:
            fehler = {"fehler": "Unvollstaendiger Tool-Call: Name fehlt", "fehler_code": "invalid_tool_call"}
            nachrichten.append({
                "role": "tool",
                "tool_call_id": tool_call_id,
                "name": "unknown_tool",
                "content": json.dumps(fehler, ensure_ascii=False),
            })
            continue
        cache_key = _tool_cache_key(tool_name, argumente_raw)
        if tool_name == "write_appointment":
            neue_schreibdaten = _schreibdaten_aus_raw(argumente_raw)
            if ersetzen_erzwingen:
                neue_schreibdaten["ersetzen"] = True
        if cache_key in tool_cache:
            ergebnis = tool_cache.get(cache_key, {})
        else:
            ergebnis = _fuehre_tool_aus(tool_name, argumente_raw, request_id)
            tool_cache[cache_key] = ergebnis
        if tool_name == "write_appointment":
            neuer_vorschlag = str(ergebnis.get("vorschlag", "") or "")
        nachrichten.append({
            "role": "tool",
            "tool_call_id": tool_call_id,
            "name": tool_name,
            "content": json.dumps(ergebnis, ensure_ascii=False),
        })
    return neue_schreibdaten, neuer_vorschlag


def _pfad_aus_dateiname(dateiname: str) -> str:
    name = str(dateiname or "").strip()
    if not name:
        _audit("agent.path.resolve_name", "", "error", "name_empty", meta={"tool_name": "write_appointment"})
        return ""
    ergebnis = search_docx_files(dateiname=name)
    dateien = ergebnis.get("dateien", []) if isinstance(ergebnis, dict) else []
    if not isinstance(dateien, list) or not dateien:
        _audit("agent.path.resolve_name", "", "error", "name_not_found", meta={"tool_name": "write_appointment"})
        return ""
    exakt = [eintrag for eintrag in dateien if str(eintrag.get("dateiname", "")).lower() == name.lower()]
    kandidat = exakt[0] if exakt else dateien[0]
    pfad = str(kandidat.get("pfad", "") or "").strip()
    if pfad:
        _audit(
            "agent.path.resolve_name",
            "",
            meta={
                "tool_name": "write_appointment",
                "file_ref_hash": hash_file_ref(pfad),
                "block_reason": "name_resolution_ok",
            },
        )
    return pfad


def _pfad_aus_datum_text(text: str) -> str:
    treffer = DATUM_PATTERN.search(str(text or ""))
    if treffer is None:
        _audit("agent.path.resolve_date", "", "error", "date_not_found_in_text", meta={"tool_name": "write_appointment"})
        return ""
    pfad = str(find_file_by_date(treffer.group(0)).get("filepath", "") or "").strip()
    if pfad:
        _audit(
            "agent.path.resolve_date",
            "",
            meta={
                "tool_name": "write_appointment",
                "file_ref_hash": hash_file_ref(pfad),
                "block_reason": "date_resolution_ok",
            },
        )
    return pfad


def _normalisiere_schreibpfad(daten: dict) -> dict:
    pfadtext = str(daten.get("filepath", "") or "").strip()
    if not pfadtext:
        _audit("agent.path.normalize", "", "error", "filepath_empty", meta={"tool_name": "write_appointment"})
        return daten
    pfad_obj = Path(pfadtext)
    if pfad_obj.is_absolute() and pfad_obj.suffix.lower() == ".docx":
        _audit(
            "agent.path.normalize",
            "",
            meta={"tool_name": "write_appointment", "block_reason": "absolute_docx_path"},
        )
        return daten
    aufgeloest = _pfad_aus_dateiname(pfad_obj.name) or _pfad_aus_datum_text(pfadtext)
    if aufgeloest:
        daten["filepath"] = aufgeloest
        _audit(
            "agent.path.normalize",
            "",
            meta={
                "tool_name": "write_appointment",
                "file_ref_hash": hash_file_ref(aufgeloest),
                "block_reason": "path_normalized",
            },
        )
        return daten
    _audit(
        "agent.path.normalize",
        "",
        "error",
        "path_not_resolved",
        meta={"tool_name": "write_appointment", "block_reason": "name_and_date_resolution_empty"},
    )
    return daten


def _baue_finale_antwort(text: str, schreibdaten: dict, vorschlag: str) -> dict:
    antwort = {"antwort": text or "Keine Antwort vom Modell"}
    if schreibdaten and vorschlag:
        antwort["bestaetigung_noetig"] = True
        antwort["vorschlag"] = vorschlag
        antwort["schreibdaten"] = schreibdaten
    return antwort


def _meta_fuer_confirm_schreiben(daten: dict, ergebnis: dict) -> dict:
    meta = _meta_fuer_tool("write_appointment", ergebnis)
    meta["ersetzen"] = bool(daten.get("ersetzen", False))
    code = str(ergebnis.get("fehler_code", "") or "").strip()
    if code:
        meta["write_error_code"] = code
    return meta


def bestaetige_schreibvorschlag(schreibdaten: dict, request_id: str = "") -> dict:
    start_ms = int(time.time() * 1000)
    try:
        daten = _normalisiere_schreibpfad(_schreibdaten_aus_dict(schreibdaten))
        if not daten.get("filepath") or not daten.get("uhrzeit"):
            dauer = max(0, int(time.time() * 1000) - start_ms)
            _audit("confirm.write_appointment", request_id, "error", "data_incomplete", dauer)
            return {"fehler": "Schreibdaten unvollstaendig"}
        if not daten.get("ersetzen") and not daten.get("name"):
            dauer = max(0, int(time.time() * 1000) - start_ms)
            _audit("confirm.write_appointment", request_id, "error", "name_missing_for_new", dauer)
            return {"fehler": "Name fuer neuen Termin erforderlich"}
        ergebnis = write_appointment(**daten, bestaetigt=True)
        dauer = max(0, int(time.time() * 1000) - start_ms)
        result = "ok" if ergebnis.get("erfolg") is True else "error"
        code = "" if result == "ok" else str(ergebnis.get("fehler_code", "write_failed") or "write_failed")
        meta = _meta_fuer_confirm_schreiben(daten, ergebnis)
        _audit("confirm.write_appointment", request_id, result, code, dauer, meta)
        return ergebnis
    except Exception:
        dauer = max(0, int(time.time() * 1000) - start_ms)
        _audit("confirm.write_appointment", request_id, "error", "exception", dauer)
        return {"fehler": "Schreibvorgang fehlgeschlagen"}


def _freigabe_receipt(datenklasse: str, zweck: str, erteilt: bool) -> dict:
    status = "erteilt" if erteilt else "blockiert"
    return {
        "status": status,
        "datenklasse": str(datenklasse or "allgemein"),
        "zweck": str(zweck or "Anfrage bearbeiten"),
        "anbieter": "Mistral API (EU-Konfiguration pruefen)",
    }


def _quellen_aus_tool_cache(tool_cache: dict) -> list[str]:
    namen = []
    for schluessel in tool_cache:
        name = str(schluessel or "").split("::", 1)[0].strip()
        if name and name not in namen:
            namen.append(name)
    return namen[:8]


def _datenklasse_fuer_frage(frage: str) -> str:
    text = str(frage or "").lower()
    if any(token in text for token in ["mail", "e-mail", "email"]):
        return "mail"
    if any(token in text for token in ["termin", "word", "docx", "datei"]):
        return "termin"
    return "allgemein"


def _ki_freigabe_erlaubt(datenklasse: str) -> bool:
    config = lade_config()
    regel = config.get("ki_freigabe", {})
    if not isinstance(regel, dict):
        return False
    if str(regel.get("modus", "")).strip().lower() != "konfiguration":
        return False
    klasse = str(datenklasse or "allgemein").strip().lower()
    return bool(regel.get(klasse, False))


def run_agent(
    frage: str,
    verlauf: list[dict] | None = None,
    request_id: str = "",
    intensiv_modus: bool = False,
) -> dict:
    start_ms = int(time.time() * 1000)
    phase_a_risiko = _phase_a_nutzer_risiko(frage)
    if _enthaelt_verbotene_nutzeraktion(frage):
        dauer = max(0, int(time.time() * 1000) - start_ms)
        meta = _meta_fuer_run_agent(frage, "block", "blocked_action")
        meta.update(phase_a_risiko)
        _audit("run_agent", request_id, "error", "blocked_action", dauer, meta, frage)
        return {"antwort": "Diese Aktion ist aus Sicherheitsgruenden deaktiviert."}
    datenklasse = _datenklasse_fuer_frage(frage)
    if not _ki_freigabe_erlaubt(datenklasse):
        dauer = max(0, int(time.time() * 1000) - start_ms)
        meta = _meta_fuer_run_agent(frage, "block", "freigabe_erforderlich")
        _audit("run_agent", request_id, "error", "freigabe_erforderlich", dauer, meta, frage)
        return {
            "fehler": "Diese Datenklasse ist in config.json fuer externe KI gesperrt.",
            "fehler_code": "ki_freigabe_gesperrt",
            "freigabe_receipt": _freigabe_receipt(datenklasse, "Anfrage und benoetigte Quellen auswerten", False),
        }
    client, modell, fehler = _mistral_client_und_modell(intensiv_modus)
    if fehler:
        dauer = max(0, int(time.time() * 1000) - start_ms)
        meta = _meta_fuer_run_agent(frage, "allow")
        meta.update(phase_a_risiko)
        _audit("run_agent", request_id, "error", "client_init", dauer, meta, frage)
        return {"fehler": fehler}
    bereinigte_frage = _normalisiere_frage_datum(frage)
    timeout_s = _llm_timeout_s(modell)
    verlauf_liste = list(verlauf or [])
    kontext = _verlauf_zu_nachrichten(verlauf_liste)
    datei_kontext = _kontext_nachricht_fuer_dateibezug(bereinigte_frage, verlauf_liste)
    lese_kontext = _kontext_nachricht_fuer_dateilesen(bereinigte_frage, verlauf_liste)
    nachrichten = [{"role": "system", "content": SYSTEM_PROMPT}, _zeit_systemnachricht(), *kontext]
    if datei_kontext is not None:
        nachrichten.append(datei_kontext)
    if lese_kontext is not None:
        nachrichten.append(lese_kontext)
    nachrichten.append({"role": "user", "content": bereinigte_frage})
    letzte_schreibdaten = {}
    letzter_vorschlag = ""
    tool_cache = {}
    ersetzen_erzwingen = _frage_will_ersetzen(bereinigte_frage)
    for versuch in range(1, 9):
        nachricht, loop_fehler, loop_fehler_code = _hole_nachricht(
            client, modell, nachrichten, request_id, versuch, timeout_s, intensiv_modus
        )
        if loop_fehler:
            dauer = max(0, int(time.time() * 1000) - start_ms)
            meta = _meta_fuer_run_agent(frage, "allow")
            meta["block_reason"] = loop_fehler_code or f"try_{versuch}"
            meta.update(phase_a_risiko)
            _audit("run_agent", request_id, "error", "llm_request", dauer, meta, frage)
            return {"fehler": loop_fehler, "fehler_code": loop_fehler_code}
        if not _hat_tool_calls(nachricht):
            text = _inhalt_aus_nachricht(nachricht).strip()
            dauer = max(0, int(time.time() * 1000) - start_ms)
            meta = _meta_fuer_run_agent(frage, "allow")
            meta.update(phase_a_risiko)
            _audit("run_agent", request_id, "ok", "", dauer, meta, frage)
            antwort = _baue_finale_antwort(text, letzte_schreibdaten, letzter_vorschlag)
            antwort["freigabe_receipt"] = _freigabe_receipt(datenklasse, "Anfrage und benoetigte Quellen auswerten", True)
            antwort["quellen"] = _quellen_aus_tool_cache(tool_cache)
            return antwort
        letzte_schreibdaten, letzter_vorschlag = _verarbeite_tool_calls(
            nachricht, nachrichten, letzte_schreibdaten, letzter_vorschlag, ersetzen_erzwingen, request_id, tool_cache
        )
    dauer = max(0, int(time.time() * 1000) - start_ms)
    meta = _meta_fuer_run_agent(frage, "allow")
    meta.update(phase_a_risiko)
    _audit("run_agent", request_id, "error", "loop_limit", dauer, meta, frage)
    return {"fehler": "Tool-Loop hat das Limit erreicht", "bestaetigung_noetig": bool(letzte_schreibdaten)}


def _debug_frage(frage: str) -> None:
    ergebnis = run_agent(frage)
    print(json.dumps(ergebnis, ensure_ascii=False, indent=2))


def _json_text_aus_antwort(text: str) -> str:
    roh = str(text or "").strip()
    if not roh:
        return ""
    if roh.startswith("```"):
        teile = roh.split("```")
        for teil in teile:
            kandidat = teil.strip()
            if kandidat.startswith("json"):
                kandidat = kandidat[4:].strip()
            if kandidat.startswith("{") and kandidat.endswith("}"):
                return kandidat
    return roh


def _kategorie_norm(text: str) -> str:
    wert = str(text or "").strip().lower()
    kompakt = re.sub(r"[^a-zäöüß]", "", wert)
    if not kompakt:
        return ""
    if (
        "storno" in kompakt
        or "absage" in kompakt
        or "absagen" in kompakt
        or "terminabsage" in kompakt
    ):
        return "storno"
    if (
        "änder" in wert
        or "aender" in wert
        or "verschieb" in wert
        or "umlegung" in wert
        or "umbuch" in wert
    ):
        return "aenderung"
    if (
        "termin" in kompakt
        or "anfrage" in kompakt
        or "neutermin" in kompakt
        or "wunsch" in kompakt
    ):
        return "terminwunsch"
    return ""


def _triage_eintraege_aus_daten(daten: dict) -> list[dict]:
    eintraege = daten.get("eintraege", []) if isinstance(daten, dict) else []
    if not isinstance(eintraege, list):
        return []
    sauber = []
    for eintrag in eintraege:
        if not isinstance(eintrag, dict):
            continue
        mail_id = str(eintrag.get("mail_id", "") or "").strip()
        kategorie = _kategorie_norm(str(eintrag.get("kategorie", "") or ""))
        if not mail_id or not kategorie:
            continue
        kurztext = str(eintrag.get("kurztext", "") or "").strip()[:220]
        begruendung = str(eintrag.get("begruendung", "") or "").strip()[:220]
        sauber.append({"mail_id": mail_id, "kategorie": kategorie, "kurztext": kurztext, "begruendung": begruendung})
    return sauber


def _triage_debug_zaehler(daten: dict, treffer: list[dict]) -> dict:
    eintraege = daten.get("eintraege", []) if isinstance(daten, dict) else []
    roh_anzahl = len(eintraege) if isinstance(eintraege, list) else 0
    gueltig_anzahl = len(treffer) if isinstance(treffer, list) else 0
    ungueltig_anzahl = max(0, roh_anzahl - gueltig_anzahl)
    return {
        "roh_anzahl": roh_anzahl,
        "gueltig_anzahl": gueltig_anzahl,
        "ungueltig_anzahl": ungueltig_anzahl,
    }


def _triage_prompt_fuer_mails(mails: list[dict]) -> str:
    komprimiert = []
    for mail in mails:
        if not isinstance(mail, dict):
            continue
        text = str(mail.get("text", "") or "")[:1200]
        komprimiert.append(
            {
                "mail_id": str(mail.get("mail_id", "") or ""),
                "betreff": str(mail.get("betreff", "") or ""),
                "absender": str(mail.get("absender", "") or ""),
                "received_iso": str(mail.get("received_iso", "") or ""),
                "text": text,
            }
        )
    return json.dumps({"mails": komprimiert}, ensure_ascii=False)


def klassifiziere_mails_fuer_terminliste(mails: list[dict], request_id: str = "") -> dict:
    if not isinstance(mails, list) or not mails:
        return {"eintraege": []}
    if not _ki_freigabe_erlaubt("mail"):
        meta = _meta_compliance("mail", "block", "ki_freigabe_gesperrt")
        _audit("mail.triage", request_id, "error", "ki_freigabe_gesperrt", meta=meta)
        return {"fehler": "Mail-Daten sind in config.json fuer externe KI gesperrt", "fehler_code": "ki_freigabe_gesperrt"}
    client, _modell, fehler = _mistral_client_und_modell(True)
    if fehler:
        return {"fehler": fehler}
    system = (
        "Klassifiziere E-Mails fuer eine Arztpraxis. "
        "Gib NUR JSON zurueck: {'eintraege':[{'mail_id':'','kategorie':'terminwunsch|storno|aenderung','kurztext':'','begruendung':''}]}. "
        "Nur relevante Mails aufnehmen. Keine weiteren Keys."
    )
    nachrichten = [{"role": "system", "content": system}, {"role": "user", "content": _triage_prompt_fuer_mails(mails)}]
    try:
        antwort = _chat_completion(client, "mistral-large-latest", nachrichten, request_id, 1, LLM_TIMEOUT_LARGE_S, 1, True)
        inhalt = str(antwort.choices[0].message.content or "")
        daten = json.loads(_json_text_aus_antwort(inhalt) or "{}")
        treffer = _triage_eintraege_aus_daten(daten)
        return {"eintraege": treffer, "debug": _triage_debug_zaehler(daten, treffer)}
    except Exception:
        return {"fehler": "Mail-Klassifikation fehlgeschlagen"}


if __name__ == "__main__":
    _debug_frage("Ist am 26.07.2024 um 09:00 noch was frei?")






