"""
Keep Alive Manager - RDP & Teams (Electron IPC Implementation - Dev)
------------------------------------------------------------------

Esta versão está em desenvolvimento e implementa:
- Interface base com PyQt6
- Sistema de gestão de status via IPC (simulado)
- Detecção do processo do Teams
- Gerenciamento de RDP e tela
- System tray e agendamento

TO-DO:
- Implementar comunicação real via IPC com o Teams
- Melhorar feedback de erros
- Adicionar sistema de retry
- Implementar logs detalhados
- Investigar uso do Microsoft Graph como fallback

Autor: Maurício Menon
Data: Janeiro/2024
Versão: 1.0-dev
"""

import sys
import os
import json
import ctypes
import psutil
import win32gui
import win32con
import win32api
import win32ts
import win32process
import win32security
from datetime import datetime
import time
from pathlib import Path
from enum import Enum

# Importações Qt - todas após criação do QApplication
from PyQt6.QtCore import Qt, QTimer, QTime, QCoreApplication
from PyQt6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QSpinBox,
    QTimeEdit,
    QCheckBox,
    QPushButton,
    QSystemTrayIcon,
    QMenu,
    QStyle,
    QFrame,
    QButtonGroup,
    QGridLayout,
)
from PyQt6.QtGui import QIcon, QAction

# Configurações do sistema
ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001
ES_DISPLAY_REQUIRED = 0x00000002


class TeamsStatus(Enum):
    """Enumeração dos status possíveis do Teams"""

    AVAILABLE = ("Disponível", "Available")
    BUSY = ("Ocupado", "Busy")
    DO_NOT_DISTURB = ("Não incomodar", "DoNotDisturb")
    AWAY = ("Ausente", "Away")
    OFFLINE = ("Offline", "Offline")
    BE_RIGHT_BACK = ("Volto logo", "BeRightBack")

    def __init__(self, display_name, ipc_status):
        self.display_name = display_name
        self.ipc_status = ipc_status


class TeamsElectronManager:
    """Gerenciador de comunicação com o processo Electron do Teams"""

    def __init__(self):
        self.teams_process = None
        self.ipc_connected = False
        self.connection_timer = QTimer()
        self.connection_timer.timeout.connect(self._try_connect)
        self.connection_attempts = 0
        self.max_attempts = 3

    def start_connection(self):
        """Inicia tentativas de conexão com o Teams"""
        self.connection_timer.start(1000)

    def _try_connect(self):
        """Tenta estabelecer conexão com o Teams"""
        try:
            self.connection_attempts += 1
            print(
                f"Tentativa de conexão {self.connection_attempts}/{self.max_attempts}"
            )

            self.teams_process = self._find_teams_process()
            if not self.teams_process:
                print("Processo do Teams não encontrado")
                if self.connection_attempts >= self.max_attempts:
                    self.connection_timer.stop()
                return False

            print("Processo do Teams encontrado")
            self.ipc_connected = True
            self.connection_timer.stop()
            return True

        except Exception as e:
            print(f"Erro ao conectar com Teams: {str(e)}")
            if self.connection_attempts >= self.max_attempts:
                self.connection_timer.stop()
            return False

    def _find_teams_process(self):
        """Localiza o processo principal do Teams"""
        for proc in psutil.process_iter(["pid", "name", "cmdline"]):
            try:
                if "teams.exe" in proc.info["name"].lower():
                    if any(
                        "--type=renderer" not in str(cmd).lower()
                        for cmd in proc.info["cmdline"]
                    ):
                        return proc
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
        return None

    def set_status(self, status):
        """Simula a alteração de status via IPC"""
        if not self.ipc_connected:
            return False

        try:
            print(f"Alterando status para: {status}")
            return True
        except Exception as e:
            print(f"Erro ao definir status: {str(e)}")
            return False

    def close(self):
        """Fecha a conexão com o Teams"""
        self.connection_timer.stop()
        self.ipc_connected = False


def main():
    # Prevenir múltiplas instâncias
    from win32event import CreateMutex
    from win32api import GetLastError
    from winerror import ERROR_ALREADY_EXISTS

    handle = CreateMutex(None, 1, "KeepAliveManager_Mutex")
    if GetLastError() == ERROR_ALREADY_EXISTS:
        sys.exit(1)

    # Criar aplicação Qt primeiro
    app = QApplication(sys.argv)

    # Configurações de DPI
    if hasattr(Qt.ApplicationAttribute, "AA_EnableHighDpiScaling"):
        QCoreApplication.setAttribute(
            Qt.ApplicationAttribute.AA_EnableHighDpiScaling, True
        )
    if hasattr(Qt.ApplicationAttribute, "AA_UseHighDpiPixmaps"):
        QCoreApplication.setAttribute(
            Qt.ApplicationAttribute.AA_UseHighDpiPixmaps, True
        )
    os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"

    # Só depois de criar QApplication, importamos e definimos classes que usam widgets
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
            self.setWindowTitle("Keep Alive Manager - RDP & Teams (Electron)")
            self.setFixedSize(450, 450)

            # Configurações padrão
            self.default_interval = 120
            self.default_start_time = QTime(8, 45)
            self.default_end_time = QTime(17, 15)
            self.is_running = False
            self.activity_count = 0
            self.last_user_activity = time.time()
            self.current_teams_status = TeamsStatus.AVAILABLE

            # Inicializar gerenciador do Teams
            self.teams_manager = TeamsElectronManager()

            # Configurar ícone da aplicação
            app_icon = self.style().standardIcon(QStyle.StandardPixmap.SP_ComputerIcon)
            self.setWindowIcon(app_icon)

            # Timer para verificar atividade
            self.activity_timer = QTimer()
            self.activity_timer.timeout.connect(self.check_user_activity)
            self.activity_timer.start(1000)  # Verifica a cada segundo

            # Timer principal
            self.timer = QTimer()
            self.timer.timeout.connect(self.perform_activity)

            # Timer para horário
            self.schedule_timer = QTimer()
            self.schedule_timer.timeout.connect(self.check_schedule)
            self.schedule_timer.start(60000)

            self.setup_ui()
            self.setup_tray(app_icon)
            self.check_schedule()

        def setup_ui(self):
            """Configura a interface do usuário"""
            central_widget = QWidget()
            self.setCentralWidget(central_widget)
            layout = QVBoxLayout(central_widget)
            layout.setSpacing(15)

            # Status Frame
            status_frame = StyleFrame()
            status_layout = QVBoxLayout(status_frame)
            self.status_label = QLabel(
                "Aguardando início do serviço\n"
                "Configure os parâmetros abaixo e clique em 'Iniciar'\n"
                "O programa funcionará em segundo plano"
            )
            self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            status_layout.addWidget(self.status_label)
            layout.addWidget(status_frame)

            # Configurações Frame
            config_frame = StyleFrame()
            config_layout = QVBoxLayout(config_frame)

            # Intervalo
            interval_layout = QHBoxLayout()
            interval_layout.addWidget(QLabel("Intervalo entre ações (segundos):"))
            self.interval_spin = QSpinBox()
            self.interval_spin.setRange(30, 300)
            self.interval_spin.setValue(self.default_interval)
            interval_layout.addWidget(self.interval_spin)
            config_layout.addLayout(interval_layout)

            # Horários
            time_layout = QHBoxLayout()
            time_layout.addWidget(QLabel("Início:"))
            self.start_time_edit = QTimeEdit()
            self.start_time_edit.setTime(self.default_start_time)
            time_layout.addWidget(self.start_time_edit)

            time_layout.addWidget(QLabel("Término:"))
            self.end_time_edit = QTimeEdit()
            self.end_time_edit.setTime(self.default_end_time)
            time_layout.addWidget(self.end_time_edit)
            config_layout.addLayout(time_layout)

            layout.addWidget(config_frame)

            # Teams Status Frame
            teams_frame = StyleFrame()
            teams_layout = QVBoxLayout(teams_frame)
            teams_layout.addWidget(QLabel("Status do Teams:"))

            status_grid = QGridLayout()
            self.teams_buttons = {}

            status_positions = {
                TeamsStatus.AVAILABLE: (0, 0),
                TeamsStatus.BUSY: (0, 1),
                TeamsStatus.DO_NOT_DISTURB: (0, 2),
                TeamsStatus.AWAY: (1, 0),
                TeamsStatus.BE_RIGHT_BACK: (1, 1),
                TeamsStatus.OFFLINE: (1, 2),
            }

            for status, pos in status_positions.items():
                btn = QPushButton(status.display_name)
                btn.setCheckable(True)
                btn.setMinimumWidth(150)
                btn.clicked.connect(lambda checked, s=status: self.set_teams_status(s))
                self.teams_buttons[status] = btn
                status_grid.addWidget(btn, pos[0], pos[1])

            teams_layout.addLayout(status_grid)
            layout.addWidget(teams_frame)

            # Opções Frame
            options_frame = StyleFrame()
            options_layout = QVBoxLayout(options_frame)
            self.minimize_to_tray_cb = QCheckBox(
                "Minimizar para bandeja do sistema ao fechar"
            )
            self.minimize_to_tray_cb.setChecked(True)
            options_layout.addWidget(self.minimize_to_tray_cb)
            layout.addWidget(options_frame)

            # Botões de controle
            button_layout = QHBoxLayout()
            button_layout.addStretch()

            self.toggle_button = QPushButton("Iniciar")
            self.toggle_button.clicked.connect(self.toggle_service)
            self.toggle_button.setFixedWidth(100)
            button_layout.addWidget(self.toggle_button)

            self.minimize_button = QPushButton("Minimizar")
            self.minimize_button.clicked.connect(self.hide)
            self.minimize_button.setFixedWidth(100)
            button_layout.addWidget(self.minimize_button)

            button_layout.addStretch()
            layout.addLayout(button_layout)

        def setup_tray(self, icon):
            """Configura o ícone na bandeja do sistema"""
            self.tray_icon = QSystemTrayIcon(icon, self)
            self.tray_icon.setToolTip("Keep Alive Manager")

            tray_menu = QMenu()
            show_action = QAction("Mostrar", self)
            quit_action = QAction("Sair", self)
            toggle_action = QAction("Iniciar", self)

            show_action.triggered.connect(self.show)
            quit_action.triggered.connect(self.quit_application)
            toggle_action.triggered.connect(self.toggle_service)

            tray_menu.addAction(show_action)
            tray_menu.addAction(toggle_action)
            tray_menu.addSeparator()
            tray_menu.addAction(quit_action)

            self.tray_icon.setContextMenu(tray_menu)
            self.tray_icon.show()

            self.tray_icon.activated.connect(
                lambda reason: (
                    self.show()
                    if reason == QSystemTrayIcon.ActivationReason.DoubleClick
                    else None
                )
            )

        def check_user_activity(self):
            """Verifica se houve atividade do usuário"""
            try:
                current_pos = win32api.GetCursorPos()
                if not hasattr(self, "last_cursor_pos"):
                    self.last_cursor_pos = current_pos
                    return

                if current_pos != self.last_cursor_pos:
                    self.last_user_activity = time.time()
                    self.last_cursor_pos = current_pos
                    if (
                        not hasattr(self, "last_movement_log")
                        or time.time() - self.last_movement_log > 60
                    ):
                        print("Atividade do usuário detectada")
                        self.last_movement_log = time.time()
            except Exception as e:
                print(f"Erro ao verificar atividade do usuário: {str(e)}")

        def should_update_status(self):
            """Verifica se deve atualizar o status baseado na última atividade"""
            if not hasattr(self, "last_user_activity"):
                self.last_user_activity = time.time()
                return False
            return time.time() - self.last_user_activity >= self.interval_spin.value()

        def perform_activity(self):
            """Executa as atividades de keep-alive periodicamente"""
            try:
                print("Iniciando ciclo de atividade")

                # Verificar conexão RDP
                if not self.check_rdp_connection():
                    print("Aviso: Conexão RDP pode estar instável")
                    self.status_label.setText(
                        "Aviso: Conexão RDP pode estar instável\n"
                        "Verificando métodos alternativos..."
                    )

                # Atualizar status se necessário
                if self.should_update_status():
                    print("Atualizando status do Teams")
                    if self.teams_manager.set_status(
                        self.current_teams_status.ipc_status
                    ):
                        print("Status atualizado com sucesso")

                # Manter RDP e tela ativos
                self.keep_rdp_alive()
                self.prevent_screen_lock()

                # Atualizar status na interface
                self.activity_count += 1
                current_time = datetime.now().strftime("%H:%M:%S")

                status_text = (
                    f"Mantendo conexões ativas\n"
                    f"Status atual: {self.current_teams_status.display_name}\n"
                    f"Última atividade: {current_time} (#{self.activity_count})\n"
                    f"Próxima ação em {self.interval_spin.value()} segundos"
                )

                print(f"Ciclo completado")
                self.status_label.setText(status_text)

            except Exception as e:
                error_msg = f"Erro ao executar atividade: {str(e)}"
                print(error_msg)
                self.status_label.setText(
                    f"{error_msg}\n" "Tentando métodos alternativos..."
                )

        def keep_rdp_alive(self):
            """Manter sessão RDP ativa"""
            try:
                print("Tentando manter sessão RDP ativa")
                win32ts.WTSResetPersistentSession()
                return True
            except Exception as e:
                print(f"Erro ao manter RDP: {str(e)}")
                return False

        def check_rdp_connection(self):
            """Verificar estado da conexão RDP"""
            try:
                print("Verificando conexão RDP")
                session_id = win32ts.WTSGetActiveConsoleSessionId()
                return session_id != 0xFFFFFFFF
            except Exception as e:
                print(f"Erro ao verificar RDP: {str(e)}")
                return False

        def prevent_screen_lock(self):
            """Prevenir bloqueio de tela usando API do Windows"""
            try:
                print("Prevenindo bloqueio de tela")
                ctypes.windll.kernel32.SetThreadExecutionState(
                    ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED
                )
                return True
            except Exception as e:
                print(f"Erro ao prevenir bloqueio: {str(e)}")
                return False

        def toggle_service(self):
            """Inicia ou para o serviço de keep-alive"""
            if not self.is_running:
                print("Iniciando serviço")
                self.teams_manager.start_connection()
                self.timer.start(self.interval_spin.value() * 1000)
                self.is_running = True
                self.toggle_button.setText("Parar")
                self.status_label.setText(
                    "Serviço iniciado\n"
                    "Mantendo suas conexões ativas\n"
                    "O programa está funcionando em segundo plano"
                )
            else:
                print("Parando serviço")
                self.timer.stop()
                self.is_running = False
                self.toggle_button.setText("Iniciar")
                self.status_label.setText(
                    "Serviço parado\n"
                    "Suas conexões não estão sendo mantidas ativas\n"
                    "Clique em 'Iniciar' para retomar"
                )
                ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)

        def check_schedule(self):
            """Verifica se deve estar rodando baseado no horário configurado"""
            current_time = QTime.currentTime()
            start_time = self.start_time_edit.time()
            end_time = self.end_time_edit.time()

            print(f"Verificando horário: {current_time.toString()}")
            if start_time <= current_time <= end_time:
                if not self.is_running and self.toggle_button.text() == "Iniciar":
                    print("Dentro do horário programado - iniciando")
                    self.toggle_service()
            else:
                if self.is_running:
                    print("Fora do horário programado - parando")
                    self.toggle_service()

        def closeEvent(self, event):
            """Trata o evento de fechamento da janela"""
            if self.minimize_to_tray_cb.isChecked():
                print("Minimizando para a bandeja")
                event.ignore()
                self.hide()
                self.tray_icon.showMessage(
                    "Keep Alive Manager",
                    "O programa continua rodando em segundo plano",
                    QSystemTrayIcon.MessageIcon.Information,
                    2000,
                )
            else:
                print("Fechando aplicação")
                self.quit_application()

        def set_teams_status(self, status):
            """Define o status do Teams"""
            try:
                print(f"Alterando status para: {status.display_name}")
                if self.teams_manager.set_status(status.ipc_status):
                    self.current_teams_status = status
                    # Atualizar botões da UI
                    for s, btn in self.teams_buttons.items():
                        btn.setChecked(s == status)
                    return True
                return False
            except Exception as e:
                print(f"Erro ao definir status: {str(e)}")
                return False

        def quit_application(self):
            """Encerra a aplicação adequadamente"""
            print("Encerrando aplicação")
            if self.is_running:
                self.toggle_service()
            self.teams_manager.close()
            self.tray_icon.hide()
            QApplication.quit()

    app.setStyle("Windows")
    window = KeepAliveApp()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
