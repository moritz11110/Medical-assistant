"""Einfacher Testlauf fuer Phase-2-Word-Tools."""

from __future__ import annotations

import shutil
from pathlib import Path

from tools.word_tools import (
    find_file_by_date,
    find_free_slots,
    list_files_in_folder,
    read_word_file,
    write_appointment,
)


def _assert_true(bedingung: bool, meldung: str) -> None:
    if not bedingung:
        raise AssertionError(meldung)


def _test_list_files() -> None:
    ergebnis = list_files_in_folder()
    _assert_true("dateien" in ergebnis, "list_files_in_folder liefert keine dateien")
    dateien = set(ergebnis["dateien"])
    hat_format = "26_07_2024.docx" in dateien or "26.07.2024.docx" in dateien
    _assert_true(hat_format, "Beispieldatei fehlt in Dateiliste")


def _test_read_word_file(datei: Path) -> None:
    ergebnis = read_word_file(str(datei))
    _assert_true("tabellen" in ergebnis, "read_word_file liefert keine Tabellen")
    _assert_true(len(ergebnis["tabellen"]) >= 1, "Es wurde keine Tabelle gelesen")


def _test_find_free_slots(datei: Path) -> None:
    ergebnis = find_free_slots(str(datei))
    _assert_true("freie_slots" in ergebnis, "find_free_slots liefert keine freie_slots")
    _assert_true(isinstance(ergebnis["freie_slots"], list), "freie_slots ist keine Liste")


def _test_find_file_by_date() -> None:
    ergebnis = find_file_by_date("26.07.2024")
    _assert_true("filepath" in ergebnis, "find_file_by_date liefert keinen filepath")
    pfad = str(ergebnis["filepath"])
    gueltig = pfad.endswith("26_07_2024.docx") or pfad.endswith("26.07.2024.docx")
    _assert_true(gueltig, "Falscher Dateiname gefunden")


def _test_write_appointment(datei: Path) -> None:
    testdatei = datei.with_name(f"{datei.stem}_test.docx")
    shutil.copy2(datei, testdatei)
    freie = find_free_slots(str(testdatei)).get("freie_slots", [])
    if freie:
        ziel_uhrzeit = str(freie[0])
        ersetzen = False
    else:
        daten = read_word_file(str(testdatei)).get("tabellen", [])
        erste_tabelle = daten[0] if daten else []
        erste_datenzeile = erste_tabelle[1] if len(erste_tabelle) > 1 else ["07:30"]
        ziel_uhrzeit = str(erste_datenzeile[0] or "07:30")
        ersetzen = True
    ergebnis = write_appointment(
        str(testdatei),
        ziel_uhrzeit,
        "Test Person",
        svnr="1234",
        vgue="Ja",
        ersetzen=ersetzen,
        bestaetigt=True,
    )
    _assert_true(ergebnis.get("erfolg") is True, "write_appointment war nicht erfolgreich")
    slots_nachher = find_free_slots(str(testdatei)).get("freie_slots", [])
    if not ersetzen:
        _assert_true(ziel_uhrzeit not in slots_nachher, "Gewaehlter freier Slot wurde nicht belegt")
    testdatei.unlink(missing_ok=True)


def main() -> None:
    datei = Path("beispiele/26_07_2024.docx")
    if not datei.exists():
        datei = Path("beispiele/26.07.2024.docx")
    _assert_true(datei.exists(), "Beispieldatei fehlt")
    _test_list_files()
    _test_read_word_file(datei)
    _test_find_free_slots(datei)
    _test_find_file_by_date()
    _test_write_appointment(datei)
    print("Alle Tests erfolgreich.")


if __name__ == "__main__":
    main()
