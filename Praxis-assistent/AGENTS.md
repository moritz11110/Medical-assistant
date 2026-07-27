# PraxisAssistent – AGENTS.md

- Vor dem Erstellen einer neuen Datei immer fragen: "Soll ich das in [dateiname] anlegen?"
- Keine Funktion darf länger als 30 Zeilen sein – bei Bedarf in Hilfsfunktionen aufteilen

## Projektkontext
Windows-Desktop-Tool für eine Arztpraxis. System-Tray-App mit
Hotkey-Overlay und Api-Key.

## Stack
- Python 3.11
- PYQT, keyboard, pystray
- python-docx, win32com.client
- openai SDK gegen Mistral API (base_url: https://api.mistral.ai/v1)

## Verhalten beim Coden
- Kommentare und Variablennamen auf Deutsch
- Vor jedem Schreibvorgang in Word: Backup anlegen
- Keine Patientendaten in Logs oder print()-Ausgaben
- Jede Tool-Funktion gibt Dict zurück, wirft keine Exceptions

## Autonome Aktionen erlaubt
- Dateien erstellen und bearbeiten
- pip install ausführen
- Pytest-Tests ausführen

## Autonome Aktionen VERBOTEN
- Keine Änderungen an config.json ohne Rückfrage
- Keine Dateien außerhalb des Projektordners anfassen

## Reihenfolge beim Implementieren
1. Erst docs/TOOLS.md lesen bevor du ein Tool implementierst
2. Erst docs/ARCHITECTURE.md lesen bevor du ein neues Modul erstellst
3. Nach jeder Implementierung einen kurzen manuellen Test vorschlagen