"""Leichter Hintergrund-Agent fuer den PraxisCopilot."""

from __future__ import annotations

import sys
import threading
from pathlib import Path
from typing import Optional

import keyboard
import pystray
from PIL import Image, ImageDraw
from PyQt5.QtCore import QObject, pyqtSignal
from PyQt5.QtWidgets import QApplication

projekt_ordner = Path(__file__).resolve().parent
if str(projekt_ordner) not in sys.path:
    sys.path.insert(0, str(projekt_ordner))

from config import lade_config
from gui.popup import PopupFenster
from tools.audit_log import schreibe_audit_event

AUTOSTART_PFAD = r"Software\Microsoft\Windows\CurrentVersion\Run"
AUTOSTART_NAME = "PraxisCopilot"


class EreignisBruecke(QObject):
    popup_anzeigen = pyqtSignal()
    app_beenden = pyqtSignal()


def _erstelle_tray_bild() -> Image.Image:
    bild = Image.new("RGB", (64, 64), color=(17, 43, 70))
    zeichner = ImageDraw.Draw(bild)
    zeichner.ellipse((8, 8, 56, 56), fill=(255, 255, 255))
    zeichner.rectangle((30, 18, 34, 46), fill=(22, 122, 155))
    zeichner.rectangle((18, 30, 46, 34), fill=(22, 122, 155))
    return bild


def _autostart_befehl() -> str:
    skript = projekt_ordner / "start_praxis.vbs"
    return f'wscript.exe "{skript}"'


def _autostart_status() -> dict:
    try:
        import winreg

        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, AUTOSTART_PFAD) as schluessel:
            wert, _typ = winreg.QueryValueEx(schluessel, AUTOSTART_NAME)
        return {"erfolg": True, "aktiv": str(wert) == _autostart_befehl()}
    except FileNotFoundError:
        return {"erfolg": True, "aktiv": False}
    except OSError:
        return {"fehler": "Windows-Autostart konnte nicht gelesen werden"}


def _setze_autostart(aktiv: bool) -> dict:
    try:
        import winreg

        with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, AUTOSTART_PFAD) as schluessel:
            if aktiv:
                winreg.SetValueEx(schluessel, AUTOSTART_NAME, 0, winreg.REG_SZ, _autostart_befehl())
            else:
                try:
                    winreg.DeleteValue(schluessel, AUTOSTART_NAME)
                except FileNotFoundError:
                    pass
        return _autostart_status()
    except OSError:
        return {"fehler": "Windows-Autostart konnte nicht gespeichert werden"}


def _autostart_gewollt(config: dict) -> bool:
    hintergrund = config.get("hintergrundbetrieb", {})
    return bool(hintergrund.get("autostart_aktiv", False)) if isinstance(hintergrund, dict) else False


def _synchronisiere_autostart(config: dict) -> dict:
    gewollt = _autostart_gewollt(config)
    status = _autostart_status()
    if "fehler" in status or bool(status.get("aktiv", False)) == gewollt:
        return status
    ergebnis = _setze_autostart(gewollt)
    result = "error" if "fehler" in ergebnis else "ok"
    schreibe_audit_event("autostart.synchronisieren", result=result, error_code="registry_failed" if result == "error" else "")
    return ergebnis


def _oeffnen_aus_tray(_icon, _item, bruecke: EreignisBruecke) -> None:
    bruecke.popup_anzeigen.emit()


def _beenden_aus_tray(icon, _item, bruecke: EreignisBruecke, stop_event: threading.Event) -> None:
    stop_event.set()
    bruecke.app_beenden.emit()
    icon.stop()


def _starte_tray_thread(bruecke: EreignisBruecke, stop_event: threading.Event, icon_box: dict) -> None:
    icon = pystray.Icon("praxis_copilot", _erstelle_tray_bild(), "PraxisCopilot")
    oeffnen = lambda i, m: _oeffnen_aus_tray(i, m, bruecke)
    beenden = lambda i, m: _beenden_aus_tray(i, m, bruecke, stop_event)
    icon.menu = pystray.Menu(
        pystray.MenuItem("KI-Assistent oeffnen", oeffnen),
        pystray.MenuItem("Hintergrund-Agent beenden", beenden),
    )
    icon_box["icon"] = icon
    icon.run()
    icon_box["icon"] = None


def _starte_hotkey_listener(hotkey: str, bruecke: EreignisBruecke, stop_event: threading.Event) -> None:
    hotkey_id: Optional[int] = None
    try:
        hotkey_id = keyboard.add_hotkey(hotkey, bruecke.popup_anzeigen.emit)
        stop_event.wait()
    except Exception:
        stop_event.wait()
    finally:
        if hotkey_id is not None:
            keyboard.remove_hotkey(hotkey_id)


def _beende_tray_icon(icon_box: dict) -> None:
    icon = icon_box.get("icon")
    if icon is not None:
        icon.stop()


def _starte_thread(name: str, ziel, args: tuple) -> threading.Thread:
    thread = threading.Thread(target=ziel, args=args, name=name, daemon=True)
    thread.start()
    return thread


def _beende_app(app: QApplication, stop_event: threading.Event, icon_box: dict) -> None:
    stop_event.set()
    _beende_tray_icon(icon_box)
    app.quit()


def _leere_popup_box(box: dict) -> None:
    box["fenster"] = None


def _oeffne_popup(config: dict, popup_box: dict) -> None:
    popup = popup_box.get("fenster")
    if popup is None:
        popup = PopupFenster(config)
        popup.destroyed.connect(lambda: _leere_popup_box(popup_box))
        popup_box["fenster"] = popup
    popup.zeige_popup()


def _warte_auf_threads(threads: list[threading.Thread]) -> None:
    for thread in threads:
        thread.join(timeout=1.0)


def main() -> int:
    config = lade_config()
    _synchronisiere_autostart(config)
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)
    bruecke = EreignisBruecke()
    stop_event = threading.Event()
    icon_box = {"icon": None}
    popup_box = {"fenster": None}
    bruecke.popup_anzeigen.connect(lambda: _oeffne_popup(config, popup_box))
    bruecke.app_beenden.connect(lambda: _beende_app(app, stop_event, icon_box))
    tray_thread = _starte_thread("tray", _starte_tray_thread, (bruecke, stop_event, icon_box))
    hotkey_thread = _starte_thread("hotkey", _starte_hotkey_listener, (config.get("hotkey", "ctrl+alt+p"), bruecke, stop_event))
    exit_code = app.exec_()
    stop_event.set()
    _beende_tray_icon(icon_box)
    _warte_auf_threads([tray_thread, hotkey_thread])
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
