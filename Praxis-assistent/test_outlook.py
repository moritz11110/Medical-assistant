"""Manueller Test fuer Outlook-Tools (Phase 3)."""

from __future__ import annotations

import re

from tools.outlook_tools import get_recent_mails, search_mails


def _assert_true(bedingung: bool, meldung: str) -> None:
    if not bedingung:
        raise AssertionError(meldung)


def _jahr_aus_mail(mail: dict) -> str:
    datum = str(mail.get("datum", ""))
    treffer = re.search(r"(20\d{2})", datum)
    if treffer:
        return treffer.group(1)
    return "2026"


def _test_get_recent_mails() -> dict:
    ergebnis = get_recent_mails(anzahl=5)
    _assert_true("fehler" not in ergebnis, f"get_recent_mails Fehler: {ergebnis.get('fehler')}")
    _assert_true("mails" in ergebnis, "get_recent_mails liefert keine 'mails'")
    _assert_true(isinstance(ergebnis["mails"], list), "'mails' ist keine Liste")
    for mail in ergebnis["mails"]:
        _assert_true("mail_id" in mail, "mail_id fehlt")
        _assert_true("received_iso" in mail, "received_iso fehlt")
    print(f"get_recent_mails: OK, {len(ergebnis['mails'])} Mail(s) gelesen")
    return ergebnis


def _test_recent_filter() -> None:
    zukunft = "2999-01-01T00:00:00"
    ergebnis = get_recent_mails(anzahl=20, seit_iso=zukunft)
    _assert_true("fehler" not in ergebnis, f"recent_filter Fehler: {ergebnis.get('fehler')}")
    _assert_true(isinstance(ergebnis.get("mails", []), list), "recent_filter liefert keine Liste")
    _assert_true(len(ergebnis.get("mails", [])) == 0, "seit_iso-Filter greift nicht wie erwartet")
    print("recent_filter: OK")


def _test_search_mails(recent: dict) -> None:
    mails = recent.get("mails", [])
    stichwort = _jahr_aus_mail(mails[0]) if mails else "2026"
    ergebnis = search_mails(stichwort)
    _assert_true("fehler" not in ergebnis, f"search_mails Fehler: {ergebnis.get('fehler')}")
    _assert_true("mails" in ergebnis, "search_mails liefert keine 'mails'")
    _assert_true(isinstance(ergebnis["mails"], list), "'mails' ist keine Liste")
    print(f"search_mails: OK, {len(ergebnis['mails'])} Treffer fuer Jahres-Stichwort")


def main() -> None:
    recent = _test_get_recent_mails()
    _test_recent_filter()
    _test_search_mails(recent)
    print("Outlook-Tests erfolgreich.")


if __name__ == "__main__":
    main()
