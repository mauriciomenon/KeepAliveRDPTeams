
# pylint: disable=E0602,E0102,E1101

"""
Keep Alive Manager - RDP & Teams
Versão corrigida e funcional para Python 3.13.
Implementa:
- Verificação de sessão RDP ativa
- Prevenção de bloqueio de tela
- Simulação de atividade do mouse
- GUI funcional via PyQt6
"""

import sys
import os
import time
import ctypes
from datetime import datetime
import pyautogui
import win32ts
from PyQt6.QtCore import Qt, QTimer, QTime, QCoreApplication
from PyQt6.QtGui import QIcon, QAction
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QSpinBox, QCheckBox, QPushButton, QSystemTrayIcon,
    QMenu, QStyle, QFrame, QTimeEdit
)

ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001
ES_DISPLAY_REQUIRED = 0x00000002

def is_rdp_active():
    session_id = win32ts.WTSGetActiveConsoleSessionId()
    return session_id != 0xFFFFFFFF

class KeepAliveApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Keep Alive Manager")
        self.setFixedSize(450, 450)

        self.default_interval = 120
        self.default_start_time = QTime(8, 45)
        self.default_end_time = QTime(17, 15)
        self.is_running = False
        self.activity_count = 0

        self.activity_timer = QTimer()
        self.activity_timer.timeout.connect(self.perform_activity)

        self.schedule_timer = QTimer()
        self.schedule_timer.timeout.connect(self.check_schedule)
        self.schedule_timer.start(60000)

        self.setup_ui()
        self.setup_tray()

    def setup_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        self.status_label = QLabel("Aguardando início do serviço")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.status_label)

        config_frame = QFrame()
        config_frame.setFrameShape(QFrame.Shape.StyledPanel)
        config_layout = QVBoxLayout(config_frame)
        layout.addWidget(config_frame)

        interval_layout = QHBoxLayout()
        interval_label = QLabel("Intervalo (segundos):")
        self.interval_spin = QSpinBox()
        self.interval_spin.setRange(30, 300)
        self.interval_spin.setValue(self.default_interval)
        interval_layout.addWidget(interval_label)
        interval_layout.addWidget(self.interval_spin)
        config_layout.addLayout(interval_layout)

        time_layout = QHBoxLayout()
        start_label = QLabel("Início:")
        self.start_time_edit = QTimeEdit()
        self.start_time_edit.setTime(self.default_start_time)
        end_label = QLabel("Término:")
        self.end_time_edit = QTimeEdit()
        self.end_time_edit.setTime(self.default_end_time)
        time_layout.addWidget(start_label)
        time_layout.addWidget(self.start_time_edit)
        time_layout.addWidget(end_label)
        time_layout.addWidget(self.end_time_edit)
        config_layout.addLayout(time_layout)

        self.minimize_to_tray_cb = QCheckBox("Minimizar para bandeja ao fechar")
        self.minimize_to_tray_cb.setChecked(True)
        config_layout.addWidget(self.minimize_to_tray_cb)

        button_layout = QHBoxLayout()
        self.toggle_button = QPushButton("Iniciar")
        self.toggle_button.clicked.connect(self.toggle_service)
        self.minimize_button = QPushButton("Minimizar")
        self.minimize_button.clicked.connect(self.hide)
        button_layout.addWidget(self.toggle_button)
        button_layout.addWidget(self.minimize_button)
        layout.addLayout(button_layout)

    def setup_tray(self):
        icon = self.style().standardIcon(QStyle.StandardPixmap.SP_ComputerIcon)
        self.tray_icon = QSystemTrayIcon(icon, self)
        self.tray_icon.setToolTip("Keep Alive Manager")

        menu = QMenu()
        show_action = QAction("Mostrar", self)
        show_action.triggered.connect(self.show)
        self.toggle_tray_action = QAction("Iniciar", self)
        self.toggle_tray_action.triggered.connect(self.toggle_service)
        quit_action = QAction("Sair", self)
        quit_action.triggered.connect(self.quit_application)

        menu.addAction(show_action)
        menu.addAction(self.toggle_tray_action)
        menu.addSeparator()
        menu.addAction(quit_action)

        self.tray_icon.setContextMenu(menu)
        self.tray_icon.activated.connect(lambda reason: self.show() if reason == QSystemTrayIcon.ActivationReason.DoubleClick else None)
        self.tray_icon.show()

    def closeEvent(self, event):
        if self.minimize_to_tray_cb.isChecked():
            event.ignore()
            self.hide()
            self.tray_icon.showMessage("Keep Alive", "Rodando em segundo plano")
        else:
            self.quit_application()

    def toggle_service(self):
        if not self.is_running:
            self.activity_timer.start(self.interval_spin.value() * 1000)
            self.is_running = True
            self.toggle_button.setText("Parar")
            self.toggle_tray_action.setText("Parar")
            self.status_label.setText("Serviço iniciado")
        else:
            self.activity_timer.stop()
            self.is_running = False
            self.toggle_button.setText("Iniciar")
            self.toggle_tray_action.setText("Iniciar")
            self.status_label.setText("Serviço parado")
            ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)

    def perform_activity(self):
        try:
            if is_rdp_active():
                ctypes.windll.kernel32.SetThreadExecutionState(
                    ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED
                )
                pyautogui.moveRel(1, 0)
                pyautogui.moveRel(-1, 0)
                self.activity_count += 1
                now = datetime.now().strftime("%H:%M:%S")
                self.status_label.setText(f"Última atividade: {now} (#{self.activity_count})")
            else:
                self.status_label.setText("Sessão RDP inativa. Sem atividade.")
        except Exception as e:
            self.status_label.setText(f"Erro: {str(e)}")

    def check_schedule(self):
        now = QTime.currentTime()
        if self.start_time_edit.time() <= now <= self.end_time_edit.time():
            if not self.is_running:
                self.toggle_service()
        else:
            if self.is_running:
                self.toggle_service()

    def quit_application(self):
        self.activity_timer.stop()
        self.tray_icon.hide()
        QApplication.quit()

def main():
    app = QApplication(sys.argv)
    if hasattr(Qt.ApplicationAttribute, "AA_EnableHighDpiScaling"):
        QCoreApplication.setAttribute(Qt.ApplicationAttribute.AA_EnableHighDpiScaling, True)
    if hasattr(Qt.ApplicationAttribute, "AA_UseHighDpiPixmaps"):
        QCoreApplication.setAttribute(Qt.ApplicationAttribute.AA_UseHighDpiPixmaps, True)
    os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"
    app.setStyle("Windows")
    window = KeepAliveApp()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
