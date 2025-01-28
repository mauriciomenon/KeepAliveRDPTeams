"""
Keep Alive Manager - RDP & Teams (PyAutoGUI Implementation)
--------------------------------------------------------

Esta implementação utiliza o PyAutoGUI para automação de interface gráfica,
simulando interações do usuário para manter o RDP ativo e gerenciar o status do Teams.

Características principais:
- Usa PyAutoGUI para simulação de mouse e teclado
- Requer que a janela do Teams seja trazida para primeiro plano durante a mudança de status
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

import sys
import os
import subprocess
from datetime import datetime, time as dt_time
import time
import keyboard  # Adicionar no topo do arquivo junto com outros imports
import pyautogui
import win32gui
import win32con
import win32api
import win32ts
import win32com.client
import psutil
import ctypes
from ctypes import wintypes
from enum import Enum
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

# Configuração do PyAutoGUI
pyautogui.FAILSAFE = False
pyautogui.PAUSE = 0.1

# Configurar PyAutoGUI para usar codificação UTF-8
import locale

locale.setlocale(locale.LC_ALL, "pt_BR.UTF-8")


def debug_log(message, force_print=False):
    """
    Função global para logging com controle de verbosidade

    Args:
        message (str): Mensagem a ser registrada
        force_print (bool): Se True, força impressão mesmo para mensagens não críticas
    """
    # Lista de termos que indicam mensagens importantes/críticas
    critical_terms = [
        "ERRO",
        "Iniciando",
        "Parando",
        "Status alterado",
        "Teams não encontrado",
        "Falha",
        "CRÍTICO",
    ]

    is_critical = any(term in message for term in critical_terms)

    if is_critical or force_print:
        print(f"DEBUG: {message}")
        with open("keepalive_debug.log", "a", encoding="utf-8") as f:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            f.write(f"{timestamp} - {message}\n")


class TeamsStatus(Enum):
    """
    Enumeração dos status possíveis do Teams com suporte multi-idioma.
    Mantém a acentuação correta em todos os comandos e nomes.
    """

    AVAILABLE = ("Disponível", ["/Disponível", "/Available", "/Disponible"])
    BUSY = ("Ocupado", ["/Ocupado", "/Busy", "/Ocupado"])
    DO_NOT_DISTURB = (
        "Não incomodar",
        ["/Não incomodar", "/DoNotDisturb", "/NoMolestar"],
    )
    AWAY = ("Ausente", ["/Ausente", "/Away", "/Ausente"])
    OFFLINE = ("Offline", ["/Offline", "/Offline", "/Desconectado"])
    BE_RIGHT_BACK = ("Volto logo", ["/Volto logo", "/BeRightBack", "/Vuelvo"])

    def __init__(self, display_name, commands):
        """
        Inicializa um status do Teams.

        Args:
            display_name (str): Nome de exibição do status na interface (com acentuação)
            commands (list): Lista de comandos para cada idioma (com acentuação)
        """
        self.display_name = display_name
        self.commands = commands


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

        # Duplo clique no ícone mostra a janela
        self.tray_icon.activated.connect(
            lambda reason: (
                self.show()
                if reason == QSystemTrayIcon.ActivationReason.DoubleClick
                else None
            )
        )

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

        # Grid layout para os botões de status
        status_grid = QGridLayout()
        self.teams_buttons = {}

        # Organizar botões em grade 2x3
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
            btn.setMinimumWidth(150)  # Largura mínima para acomodar o texto
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
                # Log de movimento apenas a cada 60 segundos para reduzir verbosidade
                if (
                    not hasattr(self, "last_movement_log")
                    or time.time() - self.last_movement_log > 60
                ):
                    debug_log("Atividade do usuário detectada")
                    self.last_movement_log = time.time()
        except Exception as e:
            debug_log(f"Erro ao verificar atividade do usuário: {str(e)}", True)

    def should_move_mouse(self):
        """Verifica se deve mover o mouse baseado na última atividade"""
        try:
            if not hasattr(self, "last_user_activity"):
                self.last_user_activity = time.time()
                return False

            return time.time() - self.last_user_activity >= self.interval_spin.value()
        except Exception as e:
            debug_log(f"Erro em should_move_mouse: {str(e)}")
            return False

    def perform_activity(self):
        """
        Executa as atividades de keep-alive periodicamente.
        Mantém a sessão RDP ativa e previne mudança automática de status no Teams.
        """
        try:
            debug_log("Iniciando ciclo de atividade")

            # Verificar conexão RDP
            if not self.check_rdp_connection():
                debug_log("Aviso: Conexão RDP pode estar instável", True)
                self.status_label.setText(
                    "Aviso: Conexão RDP pode estar instável\n"
                    "Verificando métodos alternativos..."
                )

            # Mover mouse se necessário
            if self.should_move_mouse():
                debug_log("Movendo mouse para manter status ativo")
                if self.move_mouse_safely():
                    # Se moveu o mouse com sucesso, reforçar o status atual
                    if hasattr(self, "current_teams_status"):
                        self.set_teams_status(self.current_teams_status)

            # Manter RDP e tela ativos
            self.keep_rdp_alive()
            self.prevent_screen_lock()

            # Atualizar status
            self.activity_count += 1
            current_time = datetime.now().strftime("%H:%M:%S")

            status_text = (
                f"Mantendo conexões ativas\n"
                f"Status atual: {self.current_teams_status.display_name}\n"
                f"Última atividade: {current_time} (#{self.activity_count})\n"
                f"Próxima ação em {self.interval_spin.value()} segundos"
            )

            debug_log(f"Ciclo completado", True)
            self.status_label.setText(status_text)

        except Exception as e:
            error_msg = f"Erro ao executar atividade: {str(e)}"
            debug_log(error_msg, True)
            self.status_label.setText(
                f"{error_msg}\n" "Tentando métodos alternativos..."
            )

    def move_mouse_safely(self):
        """Movimento do mouse em um pequeno quadrado"""
        try:
            debug_log("Iniciando movimento seguro do mouse")
            screen_width = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
            screen_height = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)
            current_pos = pyautogui.position()

            # Definir área segura de movimento
            max_movement = 10
            x_move = min(max_movement, screen_width - current_pos[0] - 1)
            x_move = max(x_move, -current_pos[0])
            y_move = min(max_movement, screen_height - current_pos[1] - 1)
            y_move = max(y_move, -current_pos[1])

            # Executar movimento
            moves = [(x_move, 0), (0, y_move), (-x_move, 0), (0, -y_move)]
            for dx, dy in moves:
                debug_log(f"Movendo mouse: ({dx}, {dy})")
                pyautogui.moveRel(dx, dy, duration=0.2)

            return True
        except Exception as e:
            debug_log(f"Erro ao mover mouse: {str(e)}")
            return False

    def keep_rdp_alive(self):
        """Manter sessão RDP ativa"""
        try:
            debug_log("Tentando manter sessão RDP ativa")
            win32ts.WTSResetPersistentSession()
            return True
        except Exception as e:
            debug_log(f"Erro ao manter RDP: {str(e)}")
            return False

    def check_rdp_connection(self):
        """Verificar estado da conexão RDP"""
        try:
            debug_log("Verificando conexão RDP")
            session_id = win32ts.WTSGetActiveConsoleSessionId()
            return session_id != 0xFFFFFFFF
        except Exception as e:
            debug_log(f"Erro ao verificar RDP: {str(e)}")
            return False

    def prevent_screen_lock(self):
        """Prevenir bloqueio de tela usando API do Windows"""
        try:
            debug_log("Prevenindo bloqueio de tela")
            ctypes.windll.kernel32.SetThreadExecutionState(
                ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED
            )
            return True
        except Exception as e:
            debug_log(f"Erro ao prevenir bloqueio: {str(e)}")
            return False

    def toggle_service(self):
        """Inicia ou para o serviço de keep-alive"""
        if not self.is_running:
            debug_log("Iniciando serviço")
            self.timer.start(self.interval_spin.value() * 1000)
            self.is_running = True
            self.toggle_button.setText("Parar")
            self.status_label.setText(
                "Serviço iniciado\n"
                "Mantendo suas conexões ativas\n"
                "O programa está funcionando em segundo plano"
            )
        else:
            debug_log("Parando serviço")
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

        debug_log(f"Verificando horário: {current_time.toString()}")
        if start_time <= current_time <= end_time:
            if not self.is_running and self.toggle_button.text() == "Iniciar":
                debug_log("Dentro do horário programado - iniciando")
                self.toggle_service()
        else:
            if self.is_running:
                debug_log("Fora do horário programado - parando")
                self.toggle_service()

    def set_teams_status(self, status):
        """
        Define o status do Teams usando comandos nativos via caixa de pesquisa.
        Tenta comandos em português, inglês e espanhol.

        Args:
            status (TeamsStatus): Novo status a ser definido

        Returns:
            bool: True se o status foi alterado com sucesso, False caso contrário
        """
        self.current_teams_status = status
        self.debug_info = []

        def log_debug(message):
            debug_log(message)
            self.debug_info.append(message)
            self.status_label.setText(
                f"Status: {status.display_name}\n" + self.debug_info[-1]
            )

        try:
            # Atualizar botões da UI
            log_debug("1. Atualizando botões da interface...")
            for s, btn in self.teams_buttons.items():
                btn.setChecked(s == status)

            # Procurar janela do Teams
            log_debug("2. Procurando janela do Teams...")
            teams_windows = []

            def find_teams_window(hwnd, _):
                if win32gui.IsWindowVisible(hwnd):
                    window_text = win32gui.GetWindowText(hwnd)
                    if "Microsoft Teams" in window_text:
                        teams_windows.append((hwnd, window_text))
                        log_debug(f"   Encontrada janela: {window_text} (hwnd: {hwnd})")
                return True

            win32gui.EnumWindows(find_teams_window, None)

            if not teams_windows:
                raise Exception("Nenhuma janela do Teams encontrada")

            # Tentar alterar status
            for teams_hwnd, window_text in teams_windows:
                try:
                    log_debug(f"3. Tentando janela: {window_text}")

                    # Verificar e focar janela
                    if not win32gui.IsWindow(teams_hwnd):
                        continue

                    # Guardar janela atual
                    current_foreground = win32gui.GetForegroundWindow()

                    # Focar Teams
                    log_debug("4. Focando janela do Teams...")
                    win32gui.ShowWindow(teams_hwnd, win32con.SW_RESTORE)
                    win32gui.SetForegroundWindow(teams_hwnd)
                    time.sleep(0.5)

                    # Abrir caixa de pesquisa
                    log_debug("5. Abrindo caixa de pesquisa (Ctrl+E)...")
                    pyautogui.hotkey("ctrl", "e")
                    time.sleep(0.5)

                    # Tentar comandos em diferentes idiomas
                    log_debug(f"Alterando status para: {status.display_name}")
                    success = False
                    last_error = None
                    languages = ["PT-BR", "EN-US", "ES"]

                    for idx, command in enumerate(status.commands):
                        try:
                            # Limpar qualquer texto existente
                            pyautogui.hotkey("ctrl", "a")
                            pyautogui.press("delete")
                            time.sleep(0.2)

                            # Enviar comando
                            log_debug(f"Tentando comando {command} ({languages[idx]})")
                            keyboard.write(command)
                            time.sleep(0.2)
                            keyboard.press_and_release("enter")
                            time.sleep(0.5)

                            # Se não houve erro, consideramos que o comando funcionou
                            success = True
                            log_debug(
                                f"Status alterado com sucesso usando {languages[idx]}"
                            )
                            break
                        except Exception as e:
                            last_error = (
                                f"Erro com {command} ({languages[idx]}): {str(e)}"
                            )
                            debug_log(last_error, True)
                            continue
                        finally:
                            # Garantir que a caixa de pesquisa seja fechada
                            pyautogui.press("esc")
                            time.sleep(0.2)

                    if not success:
                        raise Exception(
                            f"Falha ao alterar status. Último erro: {last_error}"
                        )

                    # Pressionar Esc para fechar qualquer menu aberto
                    pyautogui.press("esc")
                    time.sleep(0.2)

                    # Voltar para a janela original
                    log_debug("7. Retornando à janela original...")
                    if current_foreground:
                        win32gui.SetForegroundWindow(current_foreground)

                    log_debug("Status alterado com sucesso!")
                    return True

                except Exception as e:
                    log_debug(f"ERRO ao tentar janela específica: {str(e)}")
                    continue

            raise Exception("Nenhuma janela do Teams respondeu corretamente")

        except Exception as e:
            error_msg = f"ERRO CRÍTICO ao alterar status: {str(e)}"
            log_debug(error_msg)
            return False

    def closeEvent(self, event):
        """Trata o evento de fechamento da janela"""
        if self.minimize_to_tray_cb.isChecked():
            debug_log("Minimizando para a bandeja")
            event.ignore()
            self.hide()
            self.tray_icon.showMessage(
                "Keep Alive Manager",
                "O programa continua rodando em segundo plano",
                QSystemTrayIcon.MessageIcon.Information,
                2000,
            )
        else:
            debug_log("Fechando aplicação")
            self.quit_application()

    def quit_application(self):
        """Encerra a aplicação adequadamente"""
        debug_log("Encerrando aplicação")
        if self.is_running:
            self.toggle_service()
        self.tray_icon.hide()
        QApplication.quit()


if __name__ == "__main__":
    # Prevenir múltiplas instâncias
    from win32event import CreateMutex
    from win32api import GetLastError
    from winerror import ERROR_ALREADY_EXISTS

    handle = CreateMutex(None, 1, "KeepAliveManager_Mutex")
    if GetLastError() == ERROR_ALREADY_EXISTS:
        sys.exit(1)

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

    # Iniciar aplicação
    app = QApplication(sys.argv)
    app.setStyle("Windows")
    window = KeepAliveApp()
    window.show()
    sys.exit(app.exec())
