# Projektplan – PraxisAssistent

## Ziel
Windows-Desktop-Tool für eine Arztpraxis-Mitarbeiterin.
Läuft im Hintergrund, öffnet per Hotkey ein Overlay-Fenster,
beantwortet Fragen zu Terminlisten und Outlook-Mails über die Mistral API.

---

## Phase 1 – Shell (kein LLM)
**Ziel:** Popup öffnet sich per Hotkey, nimmt Text entgegen, zeigt "Antwort kommt..." an.

Aufgaben:
- main.py mit Tray-Icon (pystray) und Hotkey (keyboard)
- gui/popup.py mit PyQt: Eingabefeld, Button, Ausgabebereich, "Bestätigen"-Button (zunächst versteckt)
- konfiguration.py lädt config.json
- Threading-Modell aufsetzen (GUI im Main Thread via PyQt, Tray in Thread 2, Agent als QThread)

Testkriterium: Hotkey drücken → Fenster öffnet sich. Text eingeben, Button klicken → "Platzhalter-Antwort" erscheint.

---

## Phase 2 – Word-Tools
**Ziel:** Word-Dateien zuverlässig lesen und schreiben.

Aufgaben:
- `list_files_in_folder()` implementieren (Ordner aus config.json durchsuchen)
- `search_docx_files()` implementieren (rekursive Suche in erlaubten Ordnern, nur .docx)
- `read_word_file()` implementieren
- `find_file_by_date()` implementieren (Dateiname = Datum)
- `find_free_slots()` implementieren (Tabelle 0, Spalte 1 leer = frei)
- `write_appointment()` implementieren (mit Backup-Pflicht, nur nach Bestätigung)
- `backup.create_backup()` implementieren
- Jeden Dateipfad gegen erlaubte_ordner aus config.json prüfen
- Jede Funktion einzeln mit `test_word.py` gegen Beispieldatei testen

Testkriterium: `find_free_slots("beispiele/26_07_2024.docx")` gibt korrekte freie Zeiten zurück.

---

## Phase 3 – Outlook-Tools
**Ziel:** Mails aus Outlook lesen und durchsuchen.

Aufgaben:
- `get_recent_mails()` via win32com implementieren
- `search_mails()` implementieren
- Mit echter Outlook-Instanz testen

Testkriterium: `get_recent_mails(5)` gibt 5 Mails mit Betreff, Absender, Datum, Text zurück.

---

## Phase 4 – Agent
**Ziel:** LLM ruft Tools selbstständig auf und beantwortet Fragen.

Aufgaben:
- Mistral API Key in config.json eintragen
- agent.py mit Tool-Use-Loop implementieren (openai SDK gegen Mistral API, Function Calling)
- Alle Tools aus Phase 2 und 3 als JSON-Schema registrieren
- System-Prompt auf Arztpraxis-Kontext zuschneiden (Schutz gegen Prompt Injection)
- Schreib-Sicherheit: Agent gibt Vorschau zurück, schreibt erst nach Bestätigung
- Mit echten Fragen testen: "Ist am 23.04.2026 um 14:00 noch was frei?"

Testkriterium: Agent findet die richtige Datei über search_docx_files, liest die Tabelle, antwortet korrekt.

---

## Phase 5 – Integration
**Ziel:** Alles zusammenstecken, GUI zeigt echte Agenten-Antworten.

Aufgaben:
- Agent-Thread in GUI einbinden (QThread + Qt Signals)
- Ladeanimation während Agent arbeitet
- Fehler sauber anzeigen (kein Absturz bei falschem Pfad etc.)
- config.json mit echten Praxis-Pfaden befüllen

---

## Phase 5b – Hauptfenster
**Ziel:** Einstellungs-UI für die Mitarbeiterin.

Aufgaben:
- gui/hauptfenster.py mit PyQt implementieren
- Erlaubte Ordner verwalten (hinzufügen/entfernen, mindestens 1 Ordner)
- Mistral API-Key eingeben und Verbindung testen
- Backup-Ordner öffnen Button
- Log-Bereich (keine Patientendaten)
- Minimieren in Tray statt Schließen

---

## Phase 6 – Deployment
**Ziel:** Eine .exe die die Mitarbeiterin einfach starten kann.

Aufgaben:
- pyinstaller → einzige .exe bauen
- Windows Autostart einrichten
- config.json mit echtem Mistral API-Key und Praxis-Pfaden befüllen
- Kurze Bedienungsanleitung (1 Seite) schreiben


Eventuell ein Firewall einrichten, der Die verbindung zu mistral komplett isoliertS
---

## Datenschutz-Checkliste
- [ ] Mistral DPA automatisch aktiv bei Nutzungsbedingungen – Training-Opt-Out aktivieren
- [ ] Mistral Server-Standort EU (Standardeinstellung, keinen US-Endpunkt verwenden)
- [ ] Keine Patientendaten in Log-Dateien
- [ ] Backup vor jedem Schreibvorgang
- [ ] Kein autonomes Schreiben – immer Bestätigung durch Mitarbeiterin
- [ ] Nur Whitelist-Ordner aus config.json erlaubt
- [ ] Praxis hat Datenschutzbeauftragten informiert

---

## Model A Setup auf Ziel-PC (Python + venv)

1. Python 3.11 installieren (Option "Add Python to PATH" aktivieren).
2. Projekt nach `C:/PraxisAssistent` kopieren.
3. Terminal in `C:/PraxisAssistent` öffnen.
4. Virtuelle Umgebung erstellen: `python -m venv .venv`
5. Umgebung aktivieren: `.venv\Scripts\activate`
6. Abhängigkeiten installieren: `python -m pip install -r requirements.txt`
7. In `config.json` prüfen:
   - `updater.onedrive_quelle` zeigt auf den Shared OneDrive-Ordner
   - `updater.lokales_ziel` ist `C:/PraxisAssistent`
8. Desktop-Verknüpfung auf `start_praxis.bat` erstellen.

Setup-Check pro Laptop:
- `python updater.py --check`
- `quelle_ok` und `ziel_ok` müssen beide `true` sein.

Manueller Kurztest:
- `start_praxis.bat` per Doppelklick starten.
- Prüfen, ob das Tool im Tray erscheint und Hotkey funktioniert.
- Testweise eine Datei im OneDrive-Quellordner ändern und neu starten.
