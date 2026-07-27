"""Kompakte KI-Palette fuer den PraxisCopilot."""

from __future__ import annotations

import time
from math import ceil
from pathlib import Path

from PyQt5.QtCore import QEasingCurve, QParallelAnimationGroup, QPropertyAnimation, QRect, QObject, QThread, QTimer, Qt, pyqtSignal, pyqtSlot
from PyQt5.QtGui import QKeySequence, QPixmap
from PyQt5.QtWidgets import QApplication, QDialog, QFrame, QHBoxLayout, QLabel, QPushButton, QShortcut, QTextEdit, QVBoxLayout

from agent.agent import bestaetige_schreibvorschlag, run_agent
from tools.audit_log import neue_request_id, schreibe_audit_event


class AgentWorker(QObject):
    fertig = pyqtSignal(dict)

    def __init__(self, frage: str, request_id: str):
        super().__init__()
        self.frage = frage
        self.request_id = request_id

    @pyqtSlot()
    def ausfuehren(self) -> None:
        try:
            self.fertig.emit(run_agent(self.frage, request_id=self.request_id))
        except Exception:
            self.fertig.emit({"fehler": "Agentenlauf ist unerwartet fehlgeschlagen.", "fehler_code": "agent_exception"})


class SchreibWorker(QObject):
    fertig = pyqtSignal(dict)

    def __init__(self, schreibdaten: dict, request_id: str):
        super().__init__()
        self.schreibdaten = schreibdaten
        self.request_id = request_id

    @pyqtSlot()
    def ausfuehren(self) -> None:
        self.fertig.emit(bestaetige_schreibvorschlag(self.schreibdaten, self.request_id))


class PopupFenster(QDialog):
    MIN_ANTWORT_HOEHE = 112
    ANIMATIONS_DAUER_MS = 200

    def __init__(self, config: dict):
        super().__init__()
        self.config = dict(config)
        self.agent_thread = None
        self.schreib_thread = None
        self.agent_worker = None
        self.schreib_worker = None
        self.schreibdaten_pending = {}
        self.request_id_aktiv = ""
        self.anfrage_start_ms = 0
        self.ladepunkt_index = 0
        self.groessen_animation = None
        self.lade_timer = QTimer(self)
        self.lade_timer.setInterval(250)
        self.setAttribute(Qt.WA_DeleteOnClose, True)
        self.setWindowTitle("PraxisCopilot")
        self.setWindowFlag(Qt.WindowContextHelpButtonHint, False)
        self.setWindowFlag(Qt.WindowStaysOnTopHint, True)
        self.setWindowFlag(Qt.MSWindowsFixedSizeDialogHint, True)
        self.setFixedWidth(660)
        self.resize(660, 350)
        self._baue_ui()
        self._setze_stil()
        self._verbinde_signale()

    def _baue_ui(self) -> None:
        self._initialisiere_elemente()
        layout = QVBoxLayout()
        layout.setContentsMargins(18, 16, 18, 18)
        layout.setSpacing(10)
        layout.addWidget(self._baue_kopfzeile())
        layout.addWidget(self.hinweis_label)
        layout.addWidget(self.eingabe_karte)
        layout.addWidget(self.ergebnis_rahmen)
        layout.addWidget(self.status_label)
        self.setLayout(layout)

    def _initialisiere_elemente(self) -> None:
        self.eingabe_feld = QTextEdit()
        self.eingabe_feld.setObjectName("EingabeFeld")
        self.eingabe_feld.setPlaceholderText("Anfrage eingeben, z. B. freie Termine fuer morgen pruefen ...")
        self.eingabe_feld.setFixedHeight(62)
        self.hinweis_label = QLabel("Antworten werden vorbereitet. Aenderungen bleiben bis zur Bestaetigung ein Entwurf.")
        self.hinweis_label.setObjectName("HinweisLabel")
        self.fragen_button = QPushButton("↑")
        self.fragen_button.setObjectName("SendenButton")
        self.fragen_button.setFixedSize(36, 36)
        self.fragen_button.setToolTip("Anfrage senden (Strg + Enter)")
        self.fragen_button.setAccessibleName("Anfrage senden")
        self.eingabe_karte = self._baue_eingabekarte()
        self.status_label = QLabel("")
        self.status_label.setObjectName("StatusLabel")
        self.ergebnis_rahmen = self._baue_ergebnis()
        self.ergebnis_rahmen.hide()

    def _baue_kopfzeile(self) -> QFrame:
        rahmen = QFrame()
        rahmen.setObjectName("Kopfzeile")
        layout = QHBoxLayout(rahmen)
        layout.setContentsMargins(14, 12, 14, 12)
        avatar = self._baue_mascott_avatar()
        details = QVBoxLayout()
        titel = QLabel("PraxisCopilot")
        titel.setObjectName("TitelLabel")
        untertitel = QLabel("Sichere Assistenz fuer Terminablaeufe")
        untertitel.setObjectName("UntertitelLabel")
        badge = QLabel("KI-REGEL: CONFIG")
        badge.setObjectName("StatusBadge")
        details.setSpacing(1)
        details.addWidget(titel)
        details.addWidget(untertitel)
        layout.addWidget(avatar)
        layout.addLayout(details)
        layout.addStretch(1)
        layout.addWidget(badge, alignment=Qt.AlignTop)
        return rahmen

    def _baue_mascott_avatar(self) -> QLabel:
        avatar = QLabel()
        avatar.setObjectName("MascottAvatar")
        avatar.setFixedSize(38, 38)
        pixmap = QPixmap(str(self._mascott_pfad()))
        if pixmap.isNull():
            avatar.hide()
            return avatar
        avatar.setPixmap(pixmap.scaled(32, 32, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        avatar.setAlignment(Qt.AlignCenter)
        return avatar

    def _mascott_pfad(self) -> Path:
        return Path(__file__).resolve().parents[1] / "assets" / "mascott.png"

    def _baue_eingabekarte(self) -> QFrame:
        rahmen = QFrame()
        rahmen.setObjectName("EingabeKarte")
        rahmen.setFixedHeight(112)
        layout = QVBoxLayout(rahmen)
        layout.setContentsMargins(6, 5, 6, 5)
        layout.setSpacing(0)
        aktionen = QHBoxLayout()
        aktionen.addStretch(1)
        aktionen.addWidget(self.fragen_button)
        layout.addWidget(self.eingabe_feld)
        layout.addLayout(aktionen)
        return rahmen

    def _baue_ergebnis(self) -> QFrame:
        rahmen = QFrame()
        rahmen.setObjectName("ErgebnisKarte")
        layout = QVBoxLayout(rahmen)
        layout.setContentsMargins(13, 12, 13, 12)
        titel = QLabel("Vorbereitete Antwort")
        titel.setObjectName("KartenTitel")
        self.ausgabe_feld = QTextEdit()
        self.ausgabe_feld.setReadOnly(True)
        self.ausgabe_feld.setObjectName("AusgabeFeld")
        self.ausgabe_feld.setMinimumHeight(self.MIN_ANTWORT_HOEHE)
        self.ausgabe_feld.setMaximumHeight(self.MIN_ANTWORT_HOEHE)
        self.ausgabe_feld.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.ausgabe_feld.document().setDefaultStyleSheet(self._antwort_stil())
        self.nachweis_label = QLabel("")
        self.nachweis_label.setObjectName("NachweisLabel")
        self.bestaetigen_button = QPushButton("Aenderung bestaetigen")
        self.bestaetigen_button.setObjectName("SekundaerButton")
        self.bestaetigen_button.hide()
        layout.addWidget(titel)
        layout.addWidget(self.ausgabe_feld)
        layout.addWidget(self.nachweis_label)
        layout.addWidget(self.bestaetigen_button)
        return rahmen

    def _setze_stil(self) -> None:
        self.setObjectName("PopupFenster")
        self.setStyleSheet(
            """
            #PopupFenster { background: #f7f9fc; border: 1px solid #d8e1eb; border-radius: 16px; }
            #Kopfzeile { background: #102a43; border-radius: 11px; }
            #MascottAvatar { background: #e7f7fa; border-radius: 19px; }
            #TitelLabel { color: #ffffff; font-size: 19px; font-weight: 700; }
            #UntertitelLabel { color: #c9d7e5; font-size: 11px; }
            #StatusBadge { background: #d8f4ec; color: #075c4f; border-radius: 9px; padding: 5px 8px; font-size: 10px; font-weight: 700; }
            #HinweisLabel, #NachweisLabel, #StatusLabel { color: #52687d; font-size: 12px; }
            #KartenTitel { color: #102a43; font-size: 14px; font-weight: 700; }
            #ErgebnisKarte { background: #ffffff; border: 1px solid #dbe4ee; border-radius: 10px; }
            #EingabeKarte { background: #ffffff; border: 1px solid #b9c9d8; border-radius: 10px; }
            #EingabeKarte:focus { border: 2px solid #087e8b; }
            QTextEdit#EingabeFeld { background: transparent; border: 0; padding: 6px 8px; color: #102a43; font-size: 14px; selection-background-color: #b9edf4; }
            QTextEdit#AusgabeFeld { background: #fbfdff; border: 1px solid #dbe4ee; border-radius: 10px; padding: 10px; color: #102a43; font-size: 14px; selection-background-color: #b9edf4; }
            QPushButton#SendenButton { background: #087e8b; color: #ffffff; border: 1px solid #066873; border-radius: 18px; font-size: 19px; font-weight: 700; }
            QPushButton#SendenButton:hover { background: #066873; }
            QPushButton#SendenButton:focus { border: 2px solid #102a43; }
            QPushButton#SendenButton:disabled { background: #9ebbc0; border-color: #9ebbc0; color: #f7f9fc; }
            QPushButton#SekundaerButton { background: #e9f7f8; color: #075c4f; border: 1px solid #8fc9ce; border-radius: 8px; padding: 8px 13px; font-weight: 700; }
            QPushButton#SekundaerButton:hover { background: #d9f0f2; }
            QPushButton#SekundaerButton:focus { border: 2px solid #087e8b; }
            QPushButton#SekundaerButton:disabled { color: #7f9699; background: #edf3f4; border-color: #cedbdd; }
            """
        )

    def _antwort_stil(self) -> str:
        return (
            "h1, h2, h3 { color: #102a43; font-weight: 700; margin-top: 12px; } "
            "h1 { font-size: 18px; } h2 { font-size: 16px; } h3 { font-size: 14px; } "
            "p { margin: 0 0 8px 0; } li { margin-bottom: 4px; } "
            "pre { background: #edf4f7; color: #102a43; padding: 8px; }"
        )

    def _setze_antworttext(self, text: str) -> None:
        inhalt = str(text or "")
        try:
            self.ausgabe_feld.setMarkdown(inhalt)
        except AttributeError:
            self.ausgabe_feld.setPlainText(inhalt)

    def _verbinde_signale(self) -> None:
        self.fragen_button.clicked.connect(self._frage_absenden)
        self.bestaetigen_button.clicked.connect(self._bestaetigen_absenden)
        self.lade_timer.timeout.connect(self._aktualisiere_ladeanzeige)
        self.absende_shortcut = QShortcut(QKeySequence("Ctrl+Return"), self)
        self.absende_shortcut.activated.connect(self._frage_absenden)

    def _audit(self, aktion: str, result: str = "ok", error_code: str = "", dauer: int = 0) -> None:
        schreibe_audit_event(aktion, self.request_id_aktiv, result, error_code, dauer)

    def _setze_busy_zustand(self, aktiv: bool) -> None:
        self.eingabe_feld.setEnabled(not aktiv)
        self.fragen_button.setEnabled(not aktiv)
        self.bestaetigen_button.setEnabled(not aktiv)
        if aktiv:
            self.lade_timer.start()
        else:
            self.lade_timer.stop()

    def _aktualisiere_ladeanzeige(self) -> None:
        self.ladepunkt_index = (self.ladepunkt_index + 1) % 4
        self.status_label.setText(f"Antwort wird vorbereitet{'.' * self.ladepunkt_index}")

    def _starte_thread(self, worker: QObject, art: str) -> None:
        thread = QThread(self)
        worker.moveToThread(thread)
        thread.started.connect(worker.ausfuehren)
        worker.fertig.connect(self._ergebnis_fertig if art == "agent" else self._schreibvorgang_fertig)
        worker.fertig.connect(thread.quit)
        thread.finished.connect(worker.deleteLater)
        thread.finished.connect(thread.deleteLater)
        thread.finished.connect(lambda: self._thread_beendet(art))
        if art == "agent":
            self.agent_thread = thread
            self.agent_worker = worker
        else:
            self.schreib_thread = thread
            self.schreib_worker = worker
        thread.start()

    def _thread_beendet(self, art: str) -> None:
        if art == "agent":
            self.agent_thread = None
            self.agent_worker = None
        else:
            self.schreib_thread = None
            self.schreib_worker = None

    def _verfuegbare_geometrie(self) -> QRect:
        screen = self.screen() or QApplication.primaryScreen()
        return screen.availableGeometry() if screen is not None else QRect(0, 0, 1280, 800)

    def _maximale_dialoghoehe(self) -> int:
        return max(360, int(self._verfuegbare_geometrie().height() * 0.6))

    def _setze_antworthoehe(self, hoehe: int) -> None:
        wert = max(self.MIN_ANTWORT_HOEHE, int(hoehe))
        self.ausgabe_feld.setMinimumHeight(wert)
        self.ausgabe_feld.setMaximumHeight(wert)

    def _zielhoehe_antwort(self) -> int:
        breite = max(1, self.ausgabe_feld.viewport().width())
        dokument = self.ausgabe_feld.document()
        dokument.setTextWidth(breite)
        inhalt_hoehe = int(ceil(dokument.size().height())) + 24
        self._setze_antworthoehe(self.MIN_ANTWORT_HOEHE)
        basis_hoehe = self.sizeHint().height()
        frei = self._maximale_dialoghoehe() - basis_hoehe
        maximum = self.MIN_ANTWORT_HOEHE + max(0, frei)
        return min(max(self.MIN_ANTWORT_HOEHE, inhalt_hoehe), maximum)

    def _ziel_geometrie(self, start: QRect, zielhoehe: int) -> QRect:
        verfuegbar = self._verfuegbare_geometrie()
        y_max = verfuegbar.bottom() - zielhoehe + 1
        ziel_y = max(verfuegbar.y(), min(start.y(), y_max))
        return QRect(start.x(), ziel_y, start.width(), zielhoehe)

    def _baue_animation(self, objekt, eigenschaft: bytes, start, ende) -> QPropertyAnimation:
        animation = QPropertyAnimation(objekt, eigenschaft, self)
        animation.setDuration(self.ANIMATIONS_DAUER_MS)
        animation.setEasingCurve(QEasingCurve.OutCubic)
        animation.setStartValue(start)
        animation.setEndValue(ende)
        return animation

    def _stoppe_groessen_animation(self) -> None:
        if self.groessen_animation is None:
            return
        self.groessen_animation.stop()
        self.groessen_animation.deleteLater()
        self.groessen_animation = None

    def _animation_beendet(self, animation, zielhoehe: int) -> None:
        if self.groessen_animation is not animation:
            return
        self._setze_antworthoehe(zielhoehe)
        self.groessen_animation = None

    def _animierte_antwortgroesse(self) -> None:
        self._stoppe_groessen_animation()
        start_geometrie = self.geometry()
        start_hoehe = self.ausgabe_feld.height()
        ziel_antworthoehe = self._zielhoehe_antwort()
        self._setze_antworthoehe(ziel_antworthoehe)
        ziel_dialoghoehe = min(self.sizeHint().height(), self._maximale_dialoghoehe())
        ziel_geometrie = self._ziel_geometrie(start_geometrie, ziel_dialoghoehe)
        self._setze_antworthoehe(start_hoehe)
        if ziel_geometrie == start_geometrie and ziel_antworthoehe == start_hoehe:
            self._setze_antworthoehe(ziel_antworthoehe)
            return
        animation = QParallelAnimationGroup(self)
        animation.addAnimation(self._baue_animation(self, b"geometry", start_geometrie, ziel_geometrie))
        animation.addAnimation(self._baue_animation(self.ausgabe_feld, b"minimumHeight", start_hoehe, ziel_antworthoehe))
        animation.addAnimation(self._baue_animation(self.ausgabe_feld, b"maximumHeight", start_hoehe, ziel_antworthoehe))
        animation.finished.connect(lambda: self._animation_beendet(animation, ziel_antworthoehe))
        self.groessen_animation = animation
        animation.start()

    def _klappe_ergebnis_ein(self) -> None:
        self._stoppe_groessen_animation()
        self.ergebnis_rahmen.hide()
        self._setze_antworthoehe(self.MIN_ANTWORT_HOEHE)
        self.resize(660, 350)

    def _frage_absenden(self) -> None:
        frage = self.eingabe_feld.toPlainText().strip()
        if not frage or self.agent_thread is not None or self.schreib_thread is not None:
            return
        self.request_id_aktiv = neue_request_id()
        self.anfrage_start_ms = int(time.time() * 1000)
        self.schreibdaten_pending = {}
        self._klappe_ergebnis_ein()
        self._audit("popup.frage_absenden")
        self._setze_busy_zustand(True)
        self._starte_thread(AgentWorker(frage, self.request_id_aktiv), "agent")

    def _ergebnis_fertig(self, ergebnis: dict) -> None:
        self._setze_busy_zustand(False)
        self.ergebnis_rahmen.show()
        if "fehler" in ergebnis:
            self._zeige_fehler(ergebnis)
        else:
            self._zeige_ergebnis(ergebnis)
        dauer = max(0, int(time.time() * 1000) - self.anfrage_start_ms)
        self._audit("popup.agent_antwort", "error" if "fehler" in ergebnis else "ok", dauer=dauer)
        self._animierte_antwortgroesse()

    def _zeige_fehler(self, ergebnis: dict) -> None:
        code = str(ergebnis.get("fehler_code", "") or "")
        self.ausgabe_feld.setPlainText(str(ergebnis.get("fehler", "Unbekannter Fehler")))
        self.nachweis_label.setText(f"Status: {code or 'Fehler'}")
        self.status_label.setText("Anfrage wurde nicht ausgefuehrt.")
        self.bestaetigen_button.hide()

    def _zeige_ergebnis(self, ergebnis: dict) -> None:
        text = str(ergebnis.get("antwort", "") or "")
        vorschlag = str(ergebnis.get("vorschlag", "") or "")
        self._setze_antworttext(f"{text}\n\n{vorschlag}".strip())
        self.schreibdaten_pending = ergebnis.get("schreibdaten", {}) if isinstance(ergebnis.get("schreibdaten"), dict) else {}
        self._aktualisiere_nachweis(ergebnis)
        self.bestaetigen_button.setVisible(bool(ergebnis.get("bestaetigung_noetig")))
        self.status_label.setText("Antwort bereit.")

    def _aktualisiere_nachweis(self, ergebnis: dict) -> None:
        quellen = ergebnis.get("quellen", [])
        quelle_text = ", ".join(str(q) for q in quellen) if isinstance(quellen, list) and quellen else "keine lokalen Tools"
        if bool(ergebnis.get("bestaetigung_noetig")):
            self.nachweis_label.setText(f"Aenderung nur nach Bestaetigung · Quellen: {quelle_text}")
        else:
            self.nachweis_label.setText(f"Quellen: {quelle_text}")

    def _bestaetigen_absenden(self) -> None:
        if not self.schreibdaten_pending or self.schreib_thread is not None:
            return
        self._setze_busy_zustand(True)
        self.status_label.setText("Aenderung wird ausgefuehrt...")
        self._starte_thread(SchreibWorker(dict(self.schreibdaten_pending), self.request_id_aktiv), "schreiben")

    def _schreibvorgang_fertig(self, ergebnis: dict) -> None:
        self._setze_busy_zustand(False)
        if ergebnis.get("erfolg") is True:
            self.schreibdaten_pending = {}
            self.bestaetigen_button.hide()
            self.nachweis_label.setText("Aenderung bestaetigt · Backup wurde angelegt.")
            self.status_label.setText("Aenderung erfolgreich ausgefuehrt.")
            self._audit("popup.schreibvorgang")
            return
        self.status_label.setText("Aenderung fehlgeschlagen.")
        self.ausgabe_feld.setPlainText(str(ergebnis.get("fehler", "Unbekannter Fehler")))
        self._audit("popup.schreibvorgang", "error", "write_failed")

    def zeige_popup(self) -> None:
        self.showNormal()
        self.raise_()
        self.activateWindow()
        self.eingabe_feld.setFocus()

    def closeEvent(self, event) -> None:  # noqa: N802
        self._stoppe_groessen_animation()
        for thread in [self.agent_thread, self.schreib_thread]:
            if thread is not None and thread.isRunning():
                thread.quit()
                thread.wait(2000)
        super().closeEvent(event)
