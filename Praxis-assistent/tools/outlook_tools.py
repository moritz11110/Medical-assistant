"""Outlook-Tools fuer den PraxisAssistenten."""

from __future__ import annotations

import sys
import time
from datetime import datetime
import hashlib
from pathlib import Path

projekt_ordner = Path(__file__).resolve().parent.parent
if str(projekt_ordner) not in sys.path:
    sys.path.insert(0, str(projekt_ordner))

from config import lade_config

MAIL_ITEM_KLASSE = 43
POSTEINGANG_ID = 6
STANDARD_RETRY = 3
STANDARD_WARTEZEIT = 0.2
CACHE_TTL_STANDARD_S = 12.0
OUTLOOK_CACHE: dict[str, dict] = {}


def _lade_win32_client():
    try:
        import win32com.client as win32  # type: ignore

        return win32, ""
    except Exception:
        return None, "win32com.client ist nicht verfuegbar"


def _lade_pythoncom():
    try:
        import pythoncom  # type: ignore

        return pythoncom
    except Exception:
        return None


def _initialisiere_com() -> bool:
    pythoncom = _lade_pythoncom()
    if pythoncom is None:
        return False
    try:
        pythoncom.CoInitialize()
        return True
    except Exception:
        return False


def _beende_com(com_aktiv: bool) -> None:
    if not com_aktiv:
        return
    pythoncom = _lade_pythoncom()
    if pythoncom is None:
        return
    try:
        pythoncom.CoUninitialize()
    except Exception:
        return


def _outlook_namespace():
    win32, fehler = _lade_win32_client()
    if fehler:
        return None, fehler
    try:
        anwendung = win32.Dispatch("Outlook.Application")
        return anwendung.GetNamespace("MAPI"), ""
    except Exception:
        return None, "Outlook konnte nicht geoeffnet werden"


def _mailbox_stores(namespace) -> list:
    stores = []
    for index in range(1, namespace.Folders.Count + 1):
        stores.append(namespace.Folders.Item(index))
    return stores


def _suche_ordner_rekursiv(ordner, ordnername: str):
    if str(ordner.Name).strip().lower() == ordnername.lower():
        return ordner
    for index in range(1, ordner.Folders.Count + 1):
        gefunden = _suche_ordner_rekursiv(ordner.Folders.Item(index), ordnername)
        if gefunden is not None:
            return gefunden
    return None


def _text_aus_config(config: dict, schluessel: str, default: str) -> str:
    wert = str(config.get(schluessel, default) or "").strip()
    return wert or default


def _float_aus_config(config: dict, schluessel: str, default: float) -> float:
    try:
        wert = float(config.get(schluessel, default))
    except Exception:
        return default
    return max(0.0, wert)


def _max_mails_aus_config(config: dict) -> int:
    wert = config.get("max_mails", 20)
    if isinstance(wert, int) and wert > 0:
        return wert
    return 20


def _int_aus_config(config: dict, schluessel: str, default: int, minimum: int, maximum: int) -> int:
    try:
        wert = int(config.get(schluessel, default))
    except Exception:
        return default
    return max(minimum, min(maximum, wert))


def _optionen_aus_config() -> dict:
    config = lade_config()
    return {
        "konto": _text_aus_config(config, "outlook_konto", ""),
        "ordner": _text_aus_config(config, "outlook_ordner", "Posteingang"),
        "start_warte": _float_aus_config(config, "outlook_lese_verzoegerung_s", 0.8),
        "zweitversuch": _float_aus_config(config, "outlook_zweitversuch_s", 1.5),
        "max_mails": _max_mails_aus_config(config),
        "mail_suche_max": _int_aus_config(config, "mail_suche_max", 200, 20, 1000),
        "cache_ttl_s": _float_aus_config(config, "outlook_cache_ttl_s", CACHE_TTL_STANDARD_S),
    }


def _cache_key(optionen: dict, limit: int) -> str:
    konto = str(optionen.get("konto", "") or "").strip().lower()
    ordner = str(optionen.get("ordner", "Posteingang") or "Posteingang").strip().lower()
    return f"{konto}|{ordner}|{int(limit)}"


def _cache_ttl(optionen: dict) -> float:
    ttl = float(optionen.get("cache_ttl_s", CACHE_TTL_STANDARD_S) or CACHE_TTL_STANDARD_S)
    return max(0.0, ttl)


def _recent_cache_name(basis_key: str, seit_iso: str) -> str:
    seit = str(seit_iso or "").strip()
    if not seit:
        return f"recent::{basis_key}"
    kurz = hashlib.sha256(seit.encode("utf-8")).hexdigest()[:12]
    return f"recent::{basis_key}::{kurz}"


def _mails_kopie(mails: list) -> list:
    kopie = []
    for mail in mails:
        if isinstance(mail, dict):
            kopie.append(dict(mail))
    return kopie


def _ergebnis_kopie(ergebnis: dict) -> dict:
    daten = dict(ergebnis)
    daten["mails"] = _mails_kopie(daten.get("mails", []))
    daten["quelle"] = dict(daten.get("quelle", {}))
    return daten


def _cache_hole(cache_name: str, ttl_s: float):
    eintrag = OUTLOOK_CACHE.get(cache_name)
    if not isinstance(eintrag, dict):
        return None
    zeitstempel = float(eintrag.get("zeit", 0.0) or 0.0)
    if ttl_s <= 0 or (time.time() - zeitstempel) > ttl_s:
        OUTLOOK_CACHE.pop(cache_name, None)
        return None
    daten = eintrag.get("daten")
    if not isinstance(daten, dict):
        return None
    return _ergebnis_kopie(daten)


def _cache_setze(cache_name: str, daten: dict) -> None:
    OUTLOOK_CACHE[cache_name] = {"zeit": time.time(), "daten": _ergebnis_kopie(daten)}


def _finde_store_fuer_konto(namespace, konto: str):
    suchtext = str(konto or "").strip().lower()
    if not suchtext:
        return None
    for store in _mailbox_stores(namespace):
        name = str(getattr(store, "Name", "") or "").strip().lower()
        if suchtext in name:
            return store
    return None


def _hole_mail_ordner(namespace, optionen: dict):
    ordnername = str(optionen.get("ordner", "Posteingang") or "").strip()
    konto = str(optionen.get("konto", "") or "").strip()
    basis = None
    try:
        inbox = namespace.GetDefaultFolder(POSTEINGANG_ID)
    except Exception:
        return None, "", "Outlook-Posteingang konnte nicht geladen werden"
    if konto:
        basis = _finde_store_fuer_konto(namespace, konto)
        if basis is None:
            return None, "", "Outlook-Konto nicht gefunden"
    if basis is None:
        if not ordnername or ordnername.lower() == str(inbox.Name).strip().lower():
            return inbox, str(inbox.Name), ""
        gefunden = _suche_ordner_rekursiv(inbox.Parent, ordnername)
        if gefunden is not None:
            return gefunden, str(gefunden.Name), ""
        return inbox, str(inbox.Name), ""
    if not ordnername:
        ordnername = "Posteingang"
    gefunden = _suche_ordner_rekursiv(basis, ordnername)
    if gefunden is None:
        return None, "", "Outlook-Ordner im Konto nicht gefunden"
    return gefunden, str(gefunden.Name), ""


def _hole_item_mit_retry(items, index: int):
    for versuch in range(STANDARD_RETRY):
        try:
            return items.Item(index)
        except Exception:
            if versuch == STANDARD_RETRY - 1:
                return None
            time.sleep(STANDARD_WARTEZEIT)
    return None


def _hole_feld_mit_retry(item, feldname: str):
    for versuch in range(STANDARD_RETRY):
        try:
            return getattr(item, feldname, None)
        except Exception:
            if versuch == STANDARD_RETRY - 1:
                return None
            time.sleep(STANDARD_WARTEZEIT)
    return None


def _mail_item_zu_dict(item) -> dict:
    betreff = str(_hole_feld_mit_retry(item, "Subject") or "").strip()
    absender = str(_hole_feld_mit_retry(item, "SenderName") or "").strip()
    datum_obj = _hole_feld_mit_retry(item, "ReceivedTime")
    datum = str(datum_obj) if datum_obj is not None else ""
    text = str(_hole_feld_mit_retry(item, "Body") or "").strip()
    received_iso = _received_iso(datum_obj, datum)
    entry_id = str(_hole_feld_mit_retry(item, "EntryID") or "").strip()
    if not entry_id:
        entry_id = _mail_hash(absender, betreff, received_iso, text)
    return {
        "mail_id": entry_id,
        "betreff": betreff,
        "absender": absender,
        "datum": datum,
        "received_iso": received_iso,
        "text": text,
    }


def _mail_hash(absender: str, betreff: str, received_iso: str, text: str) -> str:
    roh = "|".join([str(absender or ""), str(betreff or ""), str(received_iso or ""), str(text[:120] or "")])
    return hashlib.sha256(roh.encode("utf-8")).hexdigest()


def _received_iso(datum_obj, datum_text: str) -> str:
    try:
        if datum_obj is not None and hasattr(datum_obj, "strftime"):
            return datum_obj.strftime("%Y-%m-%dT%H:%M:%S")
    except Exception:
        pass
    text = str(datum_text or "").strip()
    if not text:
        return ""
    for fmt in ["%d.%m.%Y %H:%M:%S", "%Y-%m-%d %H:%M:%S", "%m/%d/%Y %H:%M:%S"]:
        try:
            return datetime.strptime(text, fmt).strftime("%Y-%m-%dT%H:%M:%S")
        except Exception:
            continue
    return ""


def _mail_ist_neuer_als(mail: dict, seit_iso: str) -> bool:
    grenze = str(seit_iso or "").strip()
    if not grenze:
        return True
    empfangen = str(mail.get("received_iso", "") or "").strip()
    if not empfangen:
        return True
    return empfangen > grenze


def _filtere_mails_ab(mails: list, seit_iso: str) -> list:
    if not str(seit_iso or "").strip():
        return list(mails)
    return [mail for mail in mails if _mail_ist_neuer_als(mail, seit_iso)]


def _sortierte_mail_items(ordner):
    items = ordner.Items
    items.Sort("[ReceivedTime]", True)
    return items


def _lese_neueste_mails(ordner, limit: int) -> list:
    mails = []
    items = _sortierte_mail_items(ordner)
    for index in range(1, items.Count + 1):
        item = _hole_item_mit_retry(items, index)
        if item is None:
            continue
        if int(getattr(item, "Class", 0)) != MAIL_ITEM_KLASSE:
            continue
        mails.append(_mail_item_zu_dict(item))
        if len(mails) >= limit:
            break
    return mails


def _liste_score(mails: list) -> tuple:
    if not mails:
        return (0, "")
    erstes_datum = str(mails[0].get("datum", "") or "")
    return (len(mails), erstes_datum)


def _lese_mails_stabil(ordner, limit: int, optionen: dict):
    start_warte = float(optionen.get("start_warte", 0.8))
    zweitversuch = float(optionen.get("zweitversuch", 1.5))
    if start_warte > 0:
        time.sleep(start_warte)
    erste_liste = _lese_neueste_mails(ordner, limit)
    versuche = 1
    if zweitversuch <= 0:
        return erste_liste, versuche
    time.sleep(zweitversuch)
    zweite_liste = _lese_neueste_mails(ordner, limit)
    versuche = 2
    if _liste_score(zweite_liste) >= _liste_score(erste_liste):
        return zweite_liste, versuche
    return erste_liste, versuche


def _mail_enthaelt_stichwort(mail: dict, stichwort: str) -> bool:
    prueftext = " ".join(
        [mail.get("betreff", ""), mail.get("absender", ""), mail.get("datum", ""), mail.get("text", "")]
    )
    return stichwort.lower() in prueftext.lower()


def _tokenisiere_suchtext(stichwort: str) -> list[str]:
    stopwoerter = {"der", "die", "das", "und", "oder", "eine", "einen", "mail", "mails", "nach", "von", "mit"}
    teile = [teil.strip().lower() for teil in str(stichwort or "").replace(",", " ").split()]
    return [teil for teil in teile if len(teil) >= 3 and teil not in stopwoerter]


def _mail_passt_auf_suche(mail: dict, stichwort: str) -> bool:
    if _mail_enthaelt_stichwort(mail, stichwort):
        return True
    prueftext = " ".join(
        [mail.get("betreff", ""), mail.get("absender", ""), mail.get("datum", ""), mail.get("text", "")]
    ).lower()
    tokens = _tokenisiere_suchtext(stichwort)
    if not tokens:
        return False
    treffer = sum(1 for token in tokens if token in prueftext)
    return treffer >= 1


def get_recent_mails(anzahl: int | None = None, seit_iso: str = "") -> dict:
    com_aktiv = _initialisiere_com()
    try:
        optionen = _optionen_aus_config()
        limit = anzahl if isinstance(anzahl, int) and anzahl > 0 else int(optionen.get("max_mails", 20))
        basis_key = _cache_key(optionen, limit)
        ttl_s = _cache_ttl(optionen)
        recent_name = _recent_cache_name(basis_key, seit_iso)
        cache_treffer = _cache_hole(recent_name, ttl_s)
        if isinstance(cache_treffer, dict):
            return cache_treffer
        namespace, fehler = _outlook_namespace()
        if fehler:
            return {"fehler": fehler}
        ordner, ordnername, fehler = _hole_mail_ordner(namespace, optionen)
        if fehler:
            return {"fehler": fehler}
        mails, versuche = _lese_mails_stabil(ordner, limit, optionen)
        stand = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        quelle = {"konto": str(optionen.get("konto", "")), "ordner": ordnername}
        pool_daten = {"mails": mails, "stand": stand, "versuche": versuche, "quelle": quelle}
        gefiltert = _filtere_mails_ab(mails, seit_iso)
        ergebnis = {"mails": gefiltert, "stand": stand, "versuche": versuche, "quelle": quelle}
        _cache_setze(recent_name, ergebnis)
        _cache_setze(f"pool::{basis_key}", pool_daten)
        return ergebnis
    except Exception:
        return {"fehler": "Mails konnten nicht gelesen werden"}
    finally:
        _beende_com(com_aktiv)


def search_mails(stichwort: str) -> dict:
    com_aktiv = _initialisiere_com()
    try:
        suchtext = str(stichwort or "").strip()
        if not suchtext:
            return {"fehler": "Stichwort ist leer"}
        optionen = _optionen_aus_config()
        limit = int(optionen.get("mail_suche_max", 200))
        basis_key = _cache_key(optionen, limit)
        ttl_s = _cache_ttl(optionen)
        pool = _cache_hole(f"pool::{basis_key}", ttl_s)
        if not isinstance(pool, dict):
            namespace, fehler = _outlook_namespace()
            if fehler:
                return {"fehler": fehler}
            ordner, ordnername, fehler = _hole_mail_ordner(namespace, optionen)
            if fehler:
                return {"fehler": fehler}
            mails, versuche = _lese_mails_stabil(ordner, limit, optionen)
            stand = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            quelle = {"konto": str(optionen.get("konto", "")), "ordner": ordnername}
            pool = {"mails": mails, "stand": stand, "versuche": versuche, "quelle": quelle}
            _cache_setze(f"pool::{basis_key}", pool)
        treffer = [mail for mail in pool.get("mails", []) if _mail_passt_auf_suche(mail, suchtext)]
        stand = str(pool.get("stand", "") or datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        quelle = dict(pool.get("quelle", {}))
        versuche = int(pool.get("versuche", 1) or 1)
        return {"mails": treffer, "stand": stand, "versuche": versuche, "quelle": quelle}
    except Exception:
        return {"fehler": "Mailsuche fehlgeschlagen"}
    finally:
        _beende_com(com_aktiv)

