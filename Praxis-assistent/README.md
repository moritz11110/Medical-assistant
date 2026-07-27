# PraxisCopilot

PraxisCopilot is a local Windows assistant for handling appointment-related work in medical practices. It runs as a lightweight system-tray agent and opens a compact AI palette with `Ctrl + Alt + P`.

The assistant can prepare answers, inspect configured Outlook and Word sources, and propose Word appointment-list changes. It never writes to Word or sends email without an explicit user confirmation.

> [!WARNING]
> This is a pilot project, not a certified medical device and not a substitute for clinical, organisational, legal, or data-protection review. Test it with synthetic or approved test data before using any real patient data.

## What it does

- Opens on demand with `Ctrl + Alt + P`; there is no dashboard at startup.
- Stays idle in the Windows notification area until a user opens the palette.
- Uses Outlook and local Word appointment lists only when a request requires them.
- Creates a backup before every confirmed Word change.
- Records local audit events without storing patient data in application logs.
- Checks the configured AI data-sharing rules before a Mistral request is sent.

## Requirements

- Windows 10 or Windows 11
- Python 3.11
- Microsoft Word desktop
- Microsoft Outlook desktop, if you want to use mail functions
- A Mistral API key

The application is Windows-only because it uses Windows Tray APIs, Outlook COM automation, Word-related local workflows, and Windows Credential Manager.

## Install from GitHub

1. Download the repository as a ZIP and extract it, or clone it with Git.
2. Open PowerShell in the extracted project folder.
3. Create and activate a Python virtual environment:

   ```powershell
   py -3.11 -m venv .venv
   .\.venv\Scripts\Activate.ps1
   ```

   If PowerShell blocks activation, run the remaining commands with `.\.venv\Scripts\python.exe` instead.

4. Install the required packages:

   ```powershell
   .\.venv\Scripts\python.exe -m pip install --upgrade pip
   .\.venv\Scripts\python.exe -m pip install -r requirements.txt
   ```

5. Create a new file called `config.json` in the project folder and paste the configuration below. Replace every placeholder path with folders on your own PC.

   ```json
   {
     "app_title": "PraxisCopilot",
     "hotkey": "ctrl+alt+p",
     "feature_flags": {
       "word_search_index": true
     },
     "hintergrundbetrieb": {
       "autostart_aktiv": false,
       "manuelle_updates": true
     },
     "ki_freigabe": {
       "modus": "konfiguration",
       "allgemein": true,
       "termin": false,
       "mail": false
     },
     "updater": {
       "aktiv": false,
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
         "assets/"
       ],
       "ausschluesse": [
         "config.json",
         "logs/*",
         "backups/*",
         ".pytest_cache/*",
         "agent/word_search_index.json",
         "test*.py"
       ]
     },
     "mistral_model": "mistral-large-latest",
     "pfade": {
       "backup_ordner": "C:/PraxisCopilot/backups",
       "vgue_ordner": "C:/Praxis/Terminlisten",
       "erlaubte_ordner": [
         "C:/Praxis/Terminlisten"
       ]
     },
     "outlook_ordner": "Inbox",
     "max_mails": 60
   }
   ```

   `erlaubte_ordner` is the strict allowlist for Word files. Only add folders that the practice has approved. Keep `termin` and `mail` set to `false` until the practice has explicitly approved sending those data classes to Mistral.

6. Store your Mistral key in Windows Credential Manager. This command asks for the key without displaying it and does not put it in `config.json`:

   ```powershell
   .\.venv\Scripts\python.exe -c "from getpass import getpass; from config import speichere_mistral_api_key; print(speichere_mistral_api_key(getpass('Mistral API key: ')))"
   ```

   For short-lived troubleshooting only, you can use an environment variable in the current PowerShell session instead. Do not save a key in a script, repository, or shared configuration file:

   ```powershell
   $env:PRAXIS_MISTRAL_API_KEY = "your_mistral_key"
   ```

## Start and use

- Double-click `start_praxis.vbs` for a silent start, or run `start_praxis.bat` from a terminal.
- For visible startup errors, run `start_praxis_debug.bat`.
- Look for the PraxisCopilot icon in the Windows notification area.
- Press `Ctrl + Alt + P`, enter a request, and select the send arrow in the input card. `Ctrl + Enter` also submits the current request.
- The result stays a draft. A Word change is executed only after **Aenderung bestaetigen** is clicked.
- To stop the app, right-click the tray icon and select **Hintergrund-Agent beenden**.

## Autostart

Set `hintergrundbetrieb.autostart_aktiv` to `true` in `config.json`, then start PraxisCopilot once. It registers a per-user Windows autostart entry; no administrator permission and no Windows service are used.

To disable it, set the value back to `false`, start PraxisCopilot once so it removes the entry, then exit the tray agent. Restart the agent after every `config.json` change.

## Safety and privacy

- Mistral is the only external service used by the application. Outlook, Word, backups, audit records, and indexes stay local.
- `ki_freigabe` controls whether general, appointment, or mail data may be sent to Mistral. A blocked class is not sent.
- The application is designed to minimise persistent work-list data. Do not place patient data, API keys, or document copies in source files, logs, or GitHub issues.
- Check proposed slots, patient matching, source documents, and Word previews before confirming any change.
- Keep backups and audit records on a practice-approved local location with appropriate access controls.

## Troubleshooting

| Problem 
`py` is not recognised 
| What to check: 
Install Python 3.11 from python.org and enable the Python launcher during installation. 
| Problem 
The palette does not open 
| What to check: 
Confirm that PraxisCopilot is running in the notification area and that `hotkey` is not used by another application. 
| Problem 
No Outlook results 
| What to check: 
Open Outlook desktop, verify `outlook_ordner`, and confirm that mail sharing is enabled only if approved. 
| Problem 
Word request is blocked 
| What to check: 
Check `pfade.erlaubte_ordner`, the document location, and the relevant `ki_freigabe` value. 
enabled only if approved. 
| Problem 
Mistral request fails  
| What to check: 
Re-run the keyring command, verify internet access, and check that the selected model is available for your Mistral account. |

## License

No open-source license is currently published for this repository. Viewing the repository does not grant permission to copy, redistribute, or use it beyond rights that may apply by law.
