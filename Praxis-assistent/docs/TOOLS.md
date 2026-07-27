# Tool-Definitionen

Alle Tools liegen in `/tools/` und werden in `agent/agent.py` als
JSON-Schema für Mistral Function Calling registriert.

---

## Dateistruktur Terminlisten (wichtig!)

Terminlisten sind Word-Dateien deren **Dateiname das Datum ist**: `26_07_2024.docx`

**Tabelle 0** ist die Hauptterminliste mit 9 Spalten (Index 0–8):
| Index | Inhalt |
|-------|--------|
| 0 | Uhrzeit (z.B. "07:30") |
| 1 | Name des Patienten/Probanden |
| 2 | SVNr |
| 3 | Geburtsdatum |
| 4 | Adresse privat + Telefon |
| 5 | Firma + Kontakt-E-Mail |
| 6 | Art der Untersuchung / Tätigkeit |
| 7 | Datum letzte Untersuchung (NB = Neu, F = Folge, / = keine Angabe) |
| 8 | Formular VGÜ ja/nein |

**Zeile 0** = Kopfzeile (überspringen)
**Freier Slot** = Zeile wo Spalte 1 (Name) leer ist
**Tabellen 1 und 2** = Abrechnungsdetails (vorerst ignorieren)

---

## Schreib-Sicherheit (wichtig!)

Der Agent schreibt **nie autonom**. Bei jedem Schreibvorgang:
1. Agent schlägt vor was er eintragen würde (zeigt Vorschau im Popup)
2. Mitarbeiterin klickt "Bestätigen"
3. Erst dann wird `write_appointment()` ausgeführt

---

## Sicherheitsregeln für alle Tools

- Jeder Dateipfad wird gegen `config.json → pfade.erlaubte_ordner` geprüft
- Keine freien Pfadeingaben ohne Whitelist-Prüfung
- Keine Dateiinhalte oder Patientendaten in Logs
- Nur `.docx` Dateien werden geöffnet oder zurückgegeben

---

## list_files_in_folder
**Datei:** `tools/word_tools.py`
**Zweck:** Listet alle .docx Dateien in den erlaubten Ordnern auf
**Parameter:** keine
**Rückgabe:** `{"dateien": ["23_04_2026.docx", "24_04_2026.docx", ...]}` oder `{"fehler": "..."}`

---

## search_appointments
**Datei:** `tools/word_tools.py`
**Zweck:** Durchsucht alle erlaubten Terminlisten nach Eintraegen
**Parameter:**
- `suchtext` (str, optional) ? z.B. Name oder Begriff
- `uhrzeit` (str, optional) ? z.B. `11:00`
- `datum` (str, optional) ? `DD.MM.YYYY` oder `DD_MM_YYYY`

**Rueckgabe:** `{'treffer': [{'datum': 'DD_MM_YYYY', 'dateiname': '...', 'filepath': '...', 'uhrzeit': 'HH:MM', 'name': '...'}]}` oder `{'fehler': '...'}`

---

## read_word_file
**Datei:** `tools/word_tools.py`
**Zweck:** Liest den gesamten Text und alle Tabellen einer .docx Datei
**Parameter:** `filepath` (str) – absoluter Pfad zur Datei
**Sicherheit:** Pfad wird gegen erlaubte_ordner geprüft
**Rückgabe:** `{"freitext": "...", "tabellen": [[["Zelle", ...], ...], ...]}` oder `{"fehler": "..."}`

---

## find_free_slots
**Datei:** `tools/word_tools.py`
**Zweck:** Gibt alle freien Zeitslots eines Terminlisten-Dokuments zurück
**Parameter:** `filepath` (str)
**Logik:** Tabelle 0 durchgehen, Zeile 0 überspringen, alle Zeilen wo Spalte 1 (Name) leer ist = freier Slot. Uhrzeit aus Spalte 0 zurückgeben.
**Rückgabe:** `{"freie_slots": ["09:00", "11:30", ...], "datei": "26_07_2024.docx"}` oder `{"fehler": "..."}`

---

## find_file_by_date
**Datei:** `tools/word_tools.py`
**Zweck:** Sucht in den erlaubten Ordnern nach der Datei für ein bestimmtes Datum
**Parameter:** `datum` (str) – Format "DD.MM.YYYY" oder "DD_MM_YYYY"
**Logik:** Konvertiert Datum in Dateiname-Format (DD_MM_YYYY.docx) und sucht in erlaubten Ordnern aus config.json
**Rückgabe:** `{"filepath": "C:/Ordination/Terminlisten/26_07_2024.docx"}` oder `{"fehler": "Keine Datei für dieses Datum gefunden"}`

---

## search_docx_files
**Datei:** `tools/word_tools.py`
**Zweck:** Sucht rekursiv nach .docx Dateien in den konfigurierten erlaubten Ordnern
**Parameter:**
- `dateiname` (str, optional) – Filter auf Teilstring im Dateinamen
- `datum` (str, optional) – Datum im Format "DD_MM_YYYY" oder "DD.MM.YYYY"

**Sicherheitslogik:**
- Durchsucht nur `config.json → pfade.erlaubte_ordner`
- Gibt nur Metadaten zurück, niemals Dateiinhalt
- Prüft jeden Treffer gegen die Whitelist

**Rückgabe:** `{"dateien": [{"dateiname": "...", "pfad": "...", "datum": "DD_MM_YYYY"}]}` oder `{"fehler": "..."}`

---

## write_appointment
**Datei:** `tools/word_tools.py`
**Zweck:** Trägt einen Probanden in einen freien Slot ein
**Parameter:**
- `filepath` (str) – Pfad zur Terminlisten-Datei
- `uhrzeit` (str) – z.B. "09:00"
- `name` (str)
- `svnr` (str, optional)
- `geburtsdatum` (str, optional)
- `adresse` (str, optional)
- `firma` (str, optional)
- `untersuchungsart` (str, optional)
- `vgue` (str, optional) – "Ja" oder "Nein"

**Logik:** Backup anlegen → Tabelle 0 durchsuchen → Zeile mit passender Uhrzeit in Spalte 0 finden → Felder befüllen
**Wichtig:** Wird nur ausgeführt wenn `bestaetigt=True` übergeben wird
**Rückgabe:** `{"erfolg": true, "backup_pfad": "..."}` oder `{"fehler": "Slot nicht gefunden oder bereits belegt"}`

---

## get_recent_mails
**Datei:** `tools/outlook_tools.py`
**Zweck:** Holt die letzten N Mails aus dem konfigurierten Outlook-Ordner
**Parameter:** `anzahl` (int, default: Wert aus config.json)
**Rückgabe:** `[{"betreff": "...", "absender": "...", "datum": "...", "text": "..."}]` oder `{"fehler": "..."}`

**Hinweis (Stand aktuell):** `config.json -> max_mails` steht auf `60`. Damit werden in der Praxis die **60 neuesten Mails** berücksichtigt. Bei mehr als 60 eingegangenen Mails werden ältere Mails außerhalb dieses Fensters nicht erfasst.

---

## search_mails
**Datei:** `tools/outlook_tools.py`
**Zweck:** Durchsucht Outlook nach Mails die ein Stichwort enthalten (Name, Firma, Datum)
**Parameter:** `stichwort` (str)
**Rückgabe:** `[{"betreff": "...", "absender": "...", "datum": "...", "text": "..."}]` oder `{"fehler": "..."}`
