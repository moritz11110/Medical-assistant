# Architektur

## Coding-Regeln (für Codex verbindlich)
- **Maximale Funktionslänge: 30 Zeilen.** Ist eine Funktion länger, muss sie in kleinere Hilfsfunktionen aufgeteilt werden.
- **Vor jeder neuen Datei nachfragen:** "Soll ich das in [dateiname] unter [pfad] anlegen?" – nie einfach eine neue Datei erstellen ohne Bestätigung.

---

## Ordnerstruktur

```
praxis-assistent/
├── AGENTS.md
├── config.json
├── main.py
├── docs/
│   ├── TOOLS.md
│   ├── ARCHITECTURE.md
│   └── PLANNING.md
├── gui/
│   └── popup.py
├── agent/
│   └── agent.py
├── tools/
│   ├── word_tools.py
│   ├── outlook_tools.py
│   └── backup.py
└── beispiele/
    └── 26_07_2024.docx
```

---

## Module

### main.py
Einstiegspunkt für den ereignisbasierten Hintergrund-Agenten. Liest `config.json`,
synchronisiert den benutzerspezifischen Windows-Autostart, startet Tray und globalen
Hotkey. Es wird keine technische Oberfläche beim Start geöffnet.

### config.py
Lädt `config.json` und stellt alle Einstellungen als Dict bereit.
Wird von allen anderen Modulen importiert – nie direkt config.json lesen.

### gui/popup.py
Kompakte KI-Palette, die ausschließlich per Hotkey oder Tray geöffnet wird.
Die Palette wird erst bei Bedarf erzeugt, zeigt Ergebnis und Quellen kompakt und
blendet den **Bestätigen**-Button nur für vorbereitete Word-Änderungen ein.

### agent/agent.py
Herzstück der Anwendung. Enthält die Funktion `run_agent(user_query) -> str`.
- Definiert alle Tools als JSON-Schema für Mistral Function Calling
- Schickt Query + Tool-Definitionen an Mistral API (openai SDK mit Mistral base_url, API-Key aus config)
- Tool-Use-Loop: solange das Modell ein Tool aufruft → ausführen → Ergebnis zurückschicken
- Gibt finale Textantwort als String zurück
- Schreibvorgänge: Agent gibt nur Vorschlag zurück, schreibt erst nach Bestätigung

### tools/word_tools.py
Alle Funktionen für Word-Dateien. Siehe docs/TOOLS.md für genaue Signaturen.
- Jede Funktion gibt ein Dict zurück, wirft keine Exceptions
- Jeder Dateipfad wird gegen erlaubte_ordner aus config.json geprüft
- Keine Dateiinhalte oder Patientendaten in Logs

### tools/outlook_tools.py
Outlook-Zugriff via `win32com.client`. Verbindet sich mit der laufenden
Outlook-Instanz (nicht via IMAP/SMTP). Jede Funktion gibt Dict zurück.

### tools/backup.py
Enthält `create_backup(filepath) -> str`.
Kopiert die Datei in den Backup-Ordner aus config.json mit Timestamp im Namen.
Beispiel: `26_07_2024_backup_20240726_143022.docx`
Gibt den Pfad des Backups zurück.

---

## Threading-Modell

```
Main Thread     → PyQt GUI (Pflicht: GUI nur im Main Thread)
Thread 2        → pystray Tray-Icon
Thread 3        → keyboard Hotkey-Listener
Agent-Thread    → QThread, nur während einer KI-Anfrage aktiv
```

---

## Mistral-Anbindung

```python
from openai import OpenAI
client = OpenAI(api_key=api_key, base_url="https://api.mistral.ai/v1")
```

Der API-Key liegt im Windows-Keyring oder in `PRAXIS_MISTRAL_API_KEY`.
Die externe Datenfreigabe wird ausschließlich über `config.json → ki_freigabe`
gesteuert. Keine Patientendaten in Logs schreiben.
