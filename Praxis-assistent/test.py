import os
import time

import requests

start = time.time()

api_key = os.environ.get("PRAXIS_MISTRAL_API_KEY", "").strip()
if not api_key:
    raise SystemExit("PRAXIS_MISTRAL_API_KEY ist nicht gesetzt.")

response = requests.post(
    "https://api.mistral.ai/v1/chat/completions",
    headers={"Authorization": f"Bearer {api_key}"},
    json={
        "model": "mistral-medium-latest",
        "messages": [{"role": "user", "content": "Hallo!"}]
    }
)

dauer = time.time() - start

response.raise_for_status()
print(f"Verbindung erfolgreich. Antwortzeit: {dauer:.2f} Sekunden")
