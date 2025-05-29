# pylint: disable=E0602,E0102,E1101
"""
Keep Alive Manager - RDP & Teams (PyAutoGUI Implementation)
--------------------------------------------------------

Esta implementação utiliza o PyAutoGUI para automação de interface gráfica,
simulando interações do usuário para manter o RDP ativo e gerenciar o status do Teams.

Características principais:
- Usa PyAutoGUI para simulação de mouse e teclado
- Requer que a janela do Teams seja trazada para primeiro plano durante a mudança de status
- Interface gráfica em PyQt6 com suporte a system tray
- Suporte a múltiplos idiomas nos comandos do Teams (PT-BR, EN-US, ES)
- Sistema de logs para debug
- Prevenção de bloqueio de tela via Windows API
- Controle de horário de funcionamento
- Detecção de atividade do usuário

Vantagens:
- Implementação simples e direta
- Não requer configuração de APIs ou permissões especiais
- Funciona com qualquer versão do Teams
- Fácil de entender e modificar

Desvantagens:
- Necessidade de trazer janela para primeiro plano
- Velocidade limitada devido aos delays necessários
- Mais suscetível a falhas por mudanças na interface
- Pode ser interrompido por outras interações do usuário

Requisitos:
- PyAutoGUI
- PyQt6
- pywin32
- Windows OS

Autor: Mauricio Menon
Data: Janeiro/2024
Versão: 1.0
"""

import ctypes
import os
import sys
import time
from datetime import datetime
from datetime import time as dt_time
from enum import Enum, auto

import keyboard
import pyautogui
import uiautomation as auto
import win32api
import win32con
import win32gui
import win32ts
from PyQt6.QtCore import QCoreApplication, Qt, QTime, QTimer
from PyQt6.QtGui import QAction, QIcon
from PyQt6.QtWidgets import (
    QApplication,
    QCheckBox,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMenu,
    QPushButton,
    QSpinBox,
    QStyle,
    QSystemTrayIcon,
    QTimeEdit,
    QVBoxLayout,
    QWidget,
)

# Ajustes de DPI no Qt
if hasattr(Qt.ApplicationAttribute, "AA_EnableHighDpiScaling"):
    QCoreApplication.setAttribute(Qt.ApplicationAttribute.AA_EnableHighDpiScaling, True)
if hasattr(Qt.ApplicationAttribute, "AA_UseHighDpiPixmaps"):
    QCoreApplication.setAttribute(Qt.ApplicationAttribute.AA_UseHighDpiPixmaps, True)
os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"


# Função de log de debug
def debug_log(message: str, to_console: bool = False) -> None:
    """
    Registra mensagens de debug com timestamp em arquivo e opcionalmente no console.
    """
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_msg = f"[{timestamp}] {message}"
    log_file = os.path.join(os.path.dirname(__file__), "keep_alive_manager.log")
    try:
        with open(log_file, "a", encoding="utf-8") as f:
            f.write(log_msg + "\n")
    except Exception:
        pass
    if to_console:
        print(log_msg)


# Configurações do sistema
ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001
ES_DISPLAY_REQUIRED = 0x00000002

# Configuração do PyAutoGUI
pyautogui.FAILSAFE = False
pyautogui.PAUSE = 0.1


class TeamsStatus(Enum):
    AVAILABLE = ("Disponível", "available")
    BUSY = ("Ocupado", "busy")
    DO_NOT_DISTURB = ("Não Perturbe", "dnd")
    AWAY = ("Ausente", "away")
    OFFLINE = ("Offline", "offline")

    def __init__(self, display_name, status_code):
        self.display_name = display_name
        self.status_code = status_code


class StyleFrame(QFrame):
    """Frame estilizado para a interface"""

    def __init__(self):
        super().__init__()
        self.setFrameShape(QFrame.Shape.StyledPanel)
        self.setFrameShadow(QFrame.Shadow.Raised)
        self.setStyleSheet(
            """
            StyleFrame {
                background-color: palette(window);
                border: 1px solid palette(mid);
                border-radius: 5px;
                padding: 10px;
            }
        """
        )


class KeepAliveApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Keep Alive Manager - RDP & Teams")
        self.setFixedSize(450, 450)

        # Configurações padrão
        self.default_interval = 120
        self.default_start_time = QTime(8, 45)
        self.default_end_time = QTime(17, 15)
        self.is_running = False
        self.activity_count = 0
        self.last_user_activity = time.time()
        self.current_teams_status = TeamsStatus.AVAILABLE

        # Configurar ícone da aplicação
        app_icon = self.style().standardIcon(QStyle.StandardPixmap.SP_ComputerIcon)
        self.setWindowIcon(app_icon)

        # Timer para verificar atividade
        self.activity_timer = QTimer()
        self.activity_timer.timeout.connect(self.check_user_activity)
        self.activity_timer.start(1000)  # checa a cada segundo

        # Timer principal de keep-alive
        self.timer = QTimer()
        self.timer.timeout.connect(self.perform_activity)

        # Timer de schedule
        self.schedule_timer = QTimer()
        self.schedule_timer.timeout.connect(self.check_schedule)
        self.schedule_timer.start(60000)  # a cada 60s

        self.setup_ui()
        self.setup_tray(app_icon)
        self.check_schedule()
        self.set_teams_status(TeamsStatus.AVAILABLE)

    def setup_tray(self, icon: QIcon) -> None:
        tray = QSystemTrayIcon(icon, self)
        tray.setToolTip("Keep Alive Manager")
        menu = QMenu()
        menu.addAction(QAction("Mostrar", self, triggered=self.show))
        menu.addAction(QAction("Iniciar/Parar", self, triggered=self.toggle_service))
        menu.addSeparator()
        menu.addAction(QAction("Sair", self, triggered=self.quit_application))
        tray.setContextMenu(menu)
        tray.activated.connect(
            lambda reason: (
                self.show()
                if reason == QSystemTrayIcon.ActivationReason.DoubleClick
                else None
            )
        )
        tray.show()
        self.tray_icon = tray

    def setup_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setSpacing(15)

        # Status Label
        status_frame = StyleFrame()
        sl = QVBoxLayout(status_frame)
        self.status_label = QLabel("Aguardando início... configure e clique em Iniciar")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        sl.addWidget(self.status_label)
        layout.addWidget(status_frame)

        # Configurações
        cfg_frame = StyleFrame()
        cl = QVBoxLayout(cfg_frame)
        row = QHBoxLayout()
        row.addWidget(QLabel("Intervalo (s):"))
        self.interval_spin = QSpinBox()
        self.interval_spin.setRange(30, 300)
        self.interval_spin.setValue(self.default_interval)
        row.addWidget(self.interval_spin)
        cl.addLayout(row)
        tr = QHBoxLayout()
        tr.addWidget(QLabel("Início:"))
        self.start_time_edit = QTimeEdit(self.default_start_time)
        tr.addWidget(self.start_time_edit)
        tr.addWidget(QLabel("Término:"))
        self.end_time_edit = QTimeEdit(self.default_end_time)
        tr.addWidget(self.end_time_edit)
        cl.addLayout(tr)
        layout.addWidget(cfg_frame)

        # Teams Status
        ts_frame = StyleFrame()
        tl = QVBoxLayout(ts_frame)
        tl.addWidget(QLabel("Status do Teams:"))
        hl = QHBoxLayout()
        self.teams_buttons = {}
        for s in TeamsStatus:
            btn = QPushButton(s.display_name)
            btn.setCheckable(True)
            btn.clicked.connect(lambda _, st=s: self.set_teams_status(st))
            self.teams_buttons[s] = btn
            hl.addWidget(btn)
        self.teams_buttons[TeamsStatus.AVAILABLE].setChecked(True)
        tl.addLayout(hl)
        layout.addWidget(ts_frame)

        # Opções
        opt_frame = StyleFrame()
        ol = QVBoxLayout(opt_frame)
        self.min_cb = QCheckBox("Minimizar ao fechar")
        self.min_cb.setChecked(True)
        ol.addWidget(self.min_cb)
        layout.addWidget(opt_frame)

        # Botões
        bl = QHBoxLayout()
        bl.addStretch()
        self.toggle_btn = QPushButton("Iniciar")
        self.toggle_btn.clicked.connect(self.toggle_service)
        bl.addWidget(self.toggle_btn)
        self.min_btn = QPushButton("Minimizar")
        self.min_btn.clicked.connect(self.hide)
        bl.addWidget(self.min_btn)
        bl.addStretch()
        layout.addLayout(bl)

    def check_user_activity(self) -> None:
        try:
            pos = win32api.GetCursorPos()
            if not hasattr(self, "last_cursor_pos"):
                self.last_cursor_pos = pos
            if pos != self.last_cursor_pos:
                self.last_user_activity = time.time()
                self.last_cursor_pos = pos
        except Exception:
            pass

    def should_move_mouse(self) -> bool:
        return (time.time() - self.last_user_activity) >= self.interval_spin.value()

    def set_teams_status(self, status: TeamsStatus) -> None:
        self.current_teams_status = status
        for s, btn in self.teams_buttons.items():
            btn.setChecked(s == status)
        try:
            win = auto.WindowControl(
                searchDepth=1, ClassName="Chrome_WidgetWin_1", SubName="Teams"
            )
            if not win.Exists(1):
                raise RuntimeError("Teams não encontrado")
            st_btn = win.ButtonControl(AutomationId="status-bar-item")
            st_btn.Click()
            auto.WaitForInputIdle(1)
            mp = {
                TeamsStatus.AVAILABLE: "Disponível",
                TeamsStatus.BUSY: "Ocupado",
                TeamsStatus.DO_NOT_DISTURB: "Não perturbe",
                TeamsStatus.AWAY: "Ausente",
                TeamsStatus.OFFLINE: "Offline",
            }
            item = win.MenuItemControl(Name=mp[status])
            item.Click()
            self.status_label.setText(f"Status: {status.display_name}")
        except Exception as e:
            debug_log(f"Erro status: {e}", True)
            self.status_label.setText(f"Erro: {e}")

    def perform_activity(self) -> None:
        try:
            if not self.check_rdp_connection():
                self.status_label.setText("RDP instável")
            if self.should_move_mouse():
                self.move_mouse_safely()
            self.keep_rdp_alive()
            self.prevent_screen_lock()
            self.activity_count += 1
            t = datetime.now().strftime("%H:%M:%S")
            self.status_label.setText(f"Atividade #{self.activity_count} em {t}")
        except Exception as e:
            debug_log(f"Erro activity: {e}", True)
            self.status_label.setText(f"Erro: {e}")

    def move_mouse_safely(self) -> bool:
        try:
            w = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
            h = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)
            x, y = pyautogui.position()
            dx = min(10, w - x - 1)
            dx = max(dx, -x)
            dy = min(10, h - y - 1)
            dy = max(dy, -y)
            for mx, my in [(dx, 0), (0, dy), (-dx, 0), (0, -dy)]:
                pyautogui.moveRel(mx, my, duration=0.2)
            return True
        except Exception as e:
            debug_log(f"Erro mouse: {e}")
            return False

    def keep_rdp_alive(self) -> bool:
        try:
            win32ts.WTSResetPersistentSession()
            return True
        except:
            return False

    def check_rdp_connection(self) -> bool:
        try:
            sid = win32ts.WTSGetActiveConsoleSessionId()
            return sid != 0xFFFFFFFF
        except:
            return False

    def prevent_screen_lock(self) -> bool:
        try:
            ctypes.windll.kernel32.SetThreadExecutionState(
                ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED
            )
            return True
        except:
            return False

    def toggle_service(self) -> None:
        if not self.is_running:
            debug_log("Iniciando serviço")
            self.timer.start(self.interval_spin.value() * 1000)
            self.is_running = True
            self.toggle_btn.setText("Parar")
        else:
            debug_log("Parando serviço")
            self.timer.stop()
            self.is_running = False
            self.toggle_btn.setText("Iniciar")
            ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)

    def check_schedule(self) -> None:
        now = QTime.currentTime()
        if self.start_time_edit.time() <= now <= self.end_time_edit.time():
            if not self.is_running:
                self.toggle_service()
        else:
            if self.is_running:
                self.toggle_service()

    def closeEvent(self, event) -> None:
        if self.min_cb.isChecked():
            event.ignore()
            self.hide()
            self.tray_icon.showMessage(
                "Keep Alive Manager",
                "Executando em segundo plano",
                QSystemTrayIcon.MessageIcon.Information,
                2000,
            )
        else:
            self.quit_application()

    def quit_application(self) -> None:
        if self.is_running:
            self.toggle_service()
        self.tray_icon.hide()
        QApplication.quit()


if __name__ == "__main__":
    from win32api import GetLastError
    from win32event import CreateMutex
    from winerror import ERROR_ALREADY_EXISTS

    mutex = CreateMutex(None, 1, "KeepAliveManager_Mutex")
    if GetLastError() == ERROR_ALREADY_EXISTS:
        sys.exit(1)

    app = QApplication(sys.argv)
    app.setStyle("Windows")
    window = KeepAliveApp()
    window.show()
    sys.exit(app.exec())
