# pylint: disable=E0602,E0102,E1101

"""
Keep Alive RDP Connection 2.1.5
Maurício Menon
Foz do Iguaçu 04/06/2025
https://github.com/mauriciomenon/KeepAliveRDPTeams
"""

import ctypes
import logging
import os
import random
import sys
import time
from datetime import datetime, timedelta

import pyautogui
import win32api
import win32con
import win32event
import win32gui
from PyQt6.QtCore import QSettings, Qt, QTime, QTimer
from PyQt6.QtGui import QAction
from PyQt6.QtWidgets import (
    QApplication,
    QCheckBox,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPushButton,
    QSlider,
    QSpinBox,
    QStyle,
    QSystemTrayIcon,
    QTabWidget,
    QTextEdit,
    QTimeEdit,
    QVBoxLayout,
    QWidget,
)

# =============================================================================
# CONFIGURAÇÕES PADRÃO - DEFINIDAS NO INÍCIO
# =============================================================================
DEFAULT_INTERVAL = 60  # Segundos entre cada "despertar" do programa
DEFAULT_USER_TIMEOUT = 60  # Segundos de inatividade necessária para executar
DEFAULT_AUTOSTART_DELAY = 5000  # Milissegundos para auto-start (5 segundos)
TEAMS_CHECK_INTERVAL = 60000  # Milissegundos entre verificações do Teams (60s)
DEFAULT_START_TIME_HOUR = 8  # Hora de início padrão
DEFAULT_START_TIME_MINUTE = 0  # Minuto de início padrão
DEFAULT_END_TIME_HOUR = 18  # Hora de término padrão
DEFAULT_END_TIME_MINUTE = 0  # Minuto de término padrão
MUTEX_NAME = "KeepAlive_RDP_Unique_Instance_2025"

# Configurações para prevenir bloqueios de tela
ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001
ES_DISPLAY_REQUIRED = 0x00000002
ES_AWAYMODE_REQUIRED = 0x00000040

# Desabilita o fail-safe do PyAutoGUI para evitar problemas
pyautogui.FAILSAFE = False

# Configuração básica de logging (apenas console para debug)
logging.basicConfig(
    level=logging.WARNING,
    format="%(asctime)s - %(levelname)s - %(message)s",
)

# Verificação de instância única com melhoria e fallback
mutex = None


def is_already_running():
    """Verifica se outra instância do aplicativo já está em execução"""
    global mutex
    try:
        # Método 1: Mutex
        mutex = win32event.CreateMutex(None, False, MUTEX_NAME)
        if win32api.GetLastError() == 183:  # ERROR_ALREADY_EXISTS = 183
            return True

        # Método 2: Fallback - verificar por janela
        try:

            def enum_windows_callback(hwnd, windows):
                if win32gui.IsWindowVisible(hwnd):
                    window_text = win32gui.GetWindowText(hwnd)
                    if "Keep Alive RDP Connection" in window_text:
                        windows.append(hwnd)
                return True

            windows = []
            win32gui.EnumWindows(enum_windows_callback, windows)
            return len(windows) > 0

        except Exception:
            pass

        return False

    except Exception as e:
        logging.warning(f"Erro na verificação de instância: {str(e)}")
        # Método 3: Fallback final - verificar por arquivo de lock
        try:
            lock_file = os.path.join(os.path.expanduser("~"), ".keepalive_running")
            if os.path.exists(lock_file):
                # Verifica se o arquivo é muito antigo (mais de 1 hora)
                if time.time() - os.path.getmtime(lock_file) > 3600:
                    os.remove(lock_file)
                    return False
                return True
            else:
                # Cria arquivo de lock
                with open(lock_file, "w") as f:
                    f.write(str(os.getpid()))
                return False
        except Exception:
            return False


def cleanup_lock():
    """Remove arquivo de lock ao sair"""
    try:
        lock_file = os.path.join(os.path.expanduser("~"), ".keepalive_running")
        if os.path.exists(lock_file):
            os.remove(lock_file)
    except Exception:
        pass


def prevent_system_lock():
    """Previne bloqueio de tela e hibernação"""
    try:
        ctypes.windll.kernel32.SetThreadExecutionState(
            ES_CONTINUOUS
            | ES_SYSTEM_REQUIRED
            | ES_DISPLAY_REQUIRED
            | ES_AWAYMODE_REQUIRED
        )
        return True
    except Exception:
        return False


def get_user_activity_timeout():
    """Obtém o tempo de inatividade do usuário em segundos"""
    try:

        class LASTINPUTINFO(ctypes.Structure):
            _fields_ = [("cbSize", ctypes.c_uint), ("dwTime", ctypes.c_uint)]

        last_input_info = LASTINPUTINFO()
        last_input_info.cbSize = ctypes.sizeof(LASTINPUTINFO)
        ctypes.windll.user32.GetLastInputInfo(ctypes.byref(last_input_info))

        current_time = win32api.GetTickCount()
        idle_time = (current_time - last_input_info.dwTime) / 1000.0
        return idle_time
    except Exception:
        return 0


def simulate_safe_activity():
    """
    Simula atividade SEGURA com verificação de usuário
    """
    try:
        # Movimento SEGURO no CENTRO da tela
        screen_width, screen_height = pyautogui.size()

        # Define área CENTRAL SEGURA (10% ao redor do centro)
        center_x = screen_width // 2
        center_y = screen_height // 2
        safe_zone = min(screen_width, screen_height) // 20  # 5% da menor dimensão

        # Posição aleatória na área central
        target_x = center_x + random.randint(-safe_zone, safe_zone)
        target_y = center_y + random.randint(-safe_zone, safe_zone)

        # Movimento suave para posição segura
        pyautogui.moveTo(target_x, target_y, duration=0.2)

        # Pequeno movimento adicional
        move_x = random.randint(-3, 3)
        move_y = random.randint(-3, 3)
        pyautogui.moveRel(move_x, move_y, duration=0.1)

        # Eventos de teclado seguros
        safe_keys = ["numlock", "scrolllock", "capslock"]
        selected_key = random.choice(safe_keys)

        # Pressiona a tecla duas vezes para garantir que volta ao estado original
        pyautogui.press(selected_key)
        time.sleep(0.1)
        pyautogui.press(selected_key)

        return True, "Atividade simulada"

    except Exception as e:
        return False, f"Erro na simulação: {str(e)}"


def get_teams_status():
    """Verifica status real do Teams (disponível/ausente/ocupado)"""
    try:
        import subprocess

        # Método 1: Verificar processo do Teams
        try:
            # Verifica se Teams está rodando
            teams_processes = []
            for proc in subprocess.run(
                ["tasklist", "/FI", "IMAGENAME eq ms-teams.exe", "/FO", "CSV"],
                capture_output=True,
                text=True,
                shell=True,
            ).stdout.split("\n"):
                if "ms-teams.exe" in proc.lower():
                    teams_processes.append(proc)

            # Se não tem processo do Teams, não está ativo
            if not teams_processes:
                # Tenta o Teams clássico
                for proc in subprocess.run(
                    ["tasklist", "/FI", "IMAGENAME eq Teams.exe", "/FO", "CSV"],
                    capture_output=True,
                    text=True,
                    shell=True,
                ).stdout.split("\n"):
                    if "Teams.exe" in proc.lower():
                        teams_processes.append(proc)

                if not teams_processes:
                    return "INATIVO", "Teams não está em execução"

        except Exception:
            pass

        # Método 2: Verificar janelas do Teams e título para status
        try:

            def enum_windows_callback(hwnd, results):
                try:
                    if win32gui.IsWindowVisible(hwnd):
                        window_text = win32gui.GetWindowText(hwnd)
                        if "teams" in window_text.lower() and len(window_text) > 5:
                            # Verifica se a janela tem tamanho razoável
                            rect = win32gui.GetWindowRect(hwnd)
                            width = rect[2] - rect[0]
                            height = rect[3] - rect[1]
                            if width > 100 and height > 100:
                                results.append(window_text)

                                # Tenta detectar status no título da janela
                                title_lower = window_text.lower()
                                if any(
                                    word in title_lower
                                    for word in ["away", "ausente", "busy", "ocupado"]
                                ):
                                    results.append("STATUS_AUSENTE")
                                elif any(
                                    word in title_lower
                                    for word in ["available", "disponível", "online"]
                                ):
                                    results.append("STATUS_DISPONIVEL")
                except Exception:
                    pass
                return True

            windows_found = []
            win32gui.EnumWindows(enum_windows_callback, windows_found)

            if windows_found:
                if "STATUS_AUSENTE" in windows_found:
                    return "AUSENTE", "Teams detectado com status ausente"
                elif "STATUS_DISPONIVEL" in windows_found:
                    return "DISPONÍVEL", "Teams detectado com status disponível"
                else:
                    return (
                        "ATIVO",
                        f"Teams detectado mas status indeterminado - {len(windows_found)} janela(s)",
                    )

        except Exception:
            pass

        # Método 3: Fallback simples - só verifica se está rodando
        try:
            teams_found = False
            for window in pyautogui.getAllWindows():
                if hasattr(window, "title") and window.title:
                    if (
                        "teams" in window.title.lower()
                        and window.width > 100
                        and window.height > 100
                    ):
                        teams_found = True
                        break

            if teams_found:
                return "DETECTADO", "Teams detectado (método fallback)"

        except Exception:
            pass

        return "INDETERMINADO", "Não foi possível determinar status do Teams"

    except Exception as e:
        return None, f"Erro na verificação: {str(e)}"


class LogTab(QWidget):
    """Aba para exibir logs de atividade"""

    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        layout.addWidget(self.log_text)

        button_layout = QHBoxLayout()
        clear_button = QPushButton("Limpar Log")
        clear_button.clicked.connect(self.clear_log)
        save_button = QPushButton("Salvar Log")
        save_button.clicked.connect(self.save_log)
        button_layout.addWidget(clear_button)
        button_layout.addWidget(save_button)
        layout.addLayout(button_layout)

    def add_log(self, message):
        """Adiciona mensagem ao log com timestamp - APENAS na interface"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.log_text.append(f"[{timestamp}] {message}")

        # Rola para o final do log
        scrollbar = self.log_text.verticalScrollBar()
        if scrollbar:
            scrollbar.setValue(scrollbar.maximum())

    def clear_log(self):
        """Limpa o log visual"""
        self.log_text.clear()

    def save_log(self):
        """Salva o log em um arquivo na área de trabalho"""
        try:
            desktop = os.path.join(os.path.expanduser("~"), "Desktop")
            filename = os.path.join(
                desktop,
                f"keep_alive_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
            )
            with open(filename, "w", encoding="utf-8") as f:
                f.write(self.log_text.toPlainText())
            self.add_log(f"Log salvo em: {filename}")
        except Exception as e:
            self.add_log(f"Erro ao salvar log: {str(e)}")


class AdvancedTab(QWidget):
    """Aba para configurações avançadas"""

    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)

        # Grupo de opções básicas
        options_group = QGroupBox("Opções de Atividade")
        options_layout = QVBoxLayout(options_group)

        self.enable_mouse = QCheckBox("Ativar movimentos do mouse")
        self.enable_mouse.setChecked(True)
        options_layout.addWidget(self.enable_mouse)

        self.enable_keyboard = QCheckBox("Ativar eventos de teclado (NumLock)")
        self.enable_keyboard.setChecked(True)
        options_layout.addWidget(self.enable_keyboard)

        self.random_intervals = QCheckBox("Usar intervalos variáveis (±30%)")
        self.random_intervals.setChecked(True)
        options_layout.addWidget(self.random_intervals)

        self.minimize_to_tray_cb = QCheckBox("Minimizar para bandeja ao fechar")
        self.minimize_to_tray_cb.setChecked(True)
        options_layout.addWidget(self.minimize_to_tray_cb)

        layout.addWidget(options_group)

        # Botões de teste
        test_group = QGroupBox("Testes")
        test_layout = QVBoxLayout(test_group)

        test_button = QPushButton("Testar Simulação Agora")
        test_button.clicked.connect(self.test_simulation)
        test_layout.addWidget(test_button)

        test_teams_button = QPushButton("Testar Detecção do Teams")
        test_teams_button.clicked.connect(self.test_teams_detection)
        test_layout.addWidget(test_teams_button)

        layout.addWidget(test_group)
        layout.addStretch()

    def test_teams_detection(self):
        """Testa detecção do Teams"""
        try:
            teams_status, teams_message = get_teams_status()
            if teams_status:
                result_msg = f"Teams Status: {teams_status} - {teams_message}"
                # Print simples para mostrar status
                print(f"DEBUG TEAMS: {teams_status} - {teams_message}")
            else:
                result_msg = f"Teams Status: ERRO - {teams_message}"
                print(f"DEBUG TEAMS ERRO: {teams_message}")

            main_window = self.window()
            if hasattr(main_window, "log_tab") and main_window.log_tab:
                main_window.log_tab.add_log("Teste de detecção do Teams")
                main_window.log_tab.add_log(result_msg)

        except Exception as e:
            error_msg = f"Erro no teste: {str(e)}"
            print(f"DEBUG TEAMS EXCEÇÃO: {error_msg}")
            main_window = self.window()
            if hasattr(main_window, "log_tab") and main_window.log_tab:
                main_window.log_tab.add_log(error_msg)

    def test_simulation(self):
        """Testa simulação"""
        try:
            # Verifica se usuário está ativo
            idle_time = get_user_activity_timeout()

            main_window = self.window()
            if hasattr(main_window, "log_tab") and main_window.log_tab:
                main_window.log_tab.add_log("Teste de simulação iniciado")
                main_window.log_tab.add_log(f"Inatividade atual: {idle_time:.1f}s")

            success, message = simulate_safe_activity()

            if success:
                result_msg = "Teste executado com sucesso"
            else:
                result_msg = message

            if hasattr(main_window, "log_tab") and main_window.log_tab:
                main_window.log_tab.add_log(result_msg)

        except Exception as e:
            error_msg = f"Erro no teste: {str(e)}"
            main_window = self.window()
            if hasattr(main_window, "log_tab") and main_window.log_tab:
                main_window.log_tab.add_log(error_msg)


class AboutTab(QWidget):
    """Aba 'About' com informações do projeto"""

    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        about_text = QLabel()
        about_text.setText(
            "Keep Alive RDP Connection 2.1.5\n\n"
            "- Correção temporização.\n"
            "- Mudança no layout de abas.\n"
            "Maurício Menon\n"
            "https://github.com/mauriciomenon/KeepAliveRDPTeams\n"
            "Foz do Iguaçu 05/06/2025\n"
        )
        about_text.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(about_text)
        layout.addStretch()


class KeepAliveApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Keep Alive RDP Connection")
        self.setFixedSize(500, 600)

        # Verifica se já está rodando - MELHORADO
        if is_already_running():
            QMessageBox.warning(
                None,
                "Aviso",
                "O Keep Alive já está em execução!\n"
                "Verifique a bandeja do sistema ou o gerenciador de tarefas.\n"
                "Se o problema persistir, reinicie o computador.",
            )
            sys.exit(1)

        # Configurações básicas
        self.settings = QSettings("KeepAliveTools", "KeepAliveManager")
        self.default_interval = self.settings.value("interval", DEFAULT_INTERVAL, int)
        self.default_user_timeout = self.settings.value(
            "user_timeout", DEFAULT_USER_TIMEOUT, int
        )
        self.default_start_time = self.settings.value(
            "start_time",
            QTime(DEFAULT_START_TIME_HOUR, DEFAULT_START_TIME_MINUTE),
            QTime,
        )
        self.default_end_time = self.settings.value(
            "end_time", QTime(DEFAULT_END_TIME_HOUR, DEFAULT_END_TIME_MINUTE), QTime
        )

        self.is_running = False
        self.activity_count = 0
        self.teams_check_count = 0

        # Timers
        self.activity_timer = QTimer()
        self.activity_timer.timeout.connect(self.perform_activity)

        # Timer para verificar Teams a cada 60 segundos
        self.teams_timer = QTimer()
        self.teams_timer.timeout.connect(self.check_teams_status)
        self.teams_timer.start(TEAMS_CHECK_INTERVAL)

        # Configura interface PRIMEIRO
        self.setup_ui()
        self.setup_tray()
        self.load_settings()

        # AUTO-START em 5 segundos como solicitado
        QTimer.singleShot(DEFAULT_AUTOSTART_DELAY, self.start_service)

    def setup_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # Status principal - SEM status do Teams
        status_frame = QFrame()
        status_frame.setFrameShape(QFrame.Shape.StyledPanel)
        status_layout = QVBoxLayout(status_frame)

        self.status_label = QLabel("Aguardando início do serviço")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        status_layout.addWidget(self.status_label)

        layout.addWidget(status_frame)

        # Abas
        self.tab_widget = QTabWidget()
        layout.addWidget(self.tab_widget)

        # Aba principal
        main_tab = QWidget()
        main_layout = QVBoxLayout(main_tab)

        # Espaço de uma linha
        main_layout.addWidget(QLabel(""))

        # Intervalo
        interval_layout = QHBoxLayout()
        interval_label = QLabel("Intervalo entre tentativas (segundos):")
        self.interval_spin = QSpinBox()
        self.interval_spin.setRange(30, 300)  # Entre 30s e 5min
        self.interval_spin.setValue(self.default_interval)
        interval_layout.addWidget(interval_label)
        interval_layout.addWidget(self.interval_spin)
        interval_layout.addStretch()
        main_layout.addLayout(interval_layout)

        # Espaço de uma linha
        main_layout.addWidget(QLabel(""))

        # Horários com tab entre textos
        time_layout = QHBoxLayout()
        start_label = QLabel("Início:")
        self.start_time_edit = QTimeEdit()
        self.start_time_edit.setTime(self.default_start_time)
        self.start_time_edit.setFixedWidth(80)
        # Tab space
        tab_label = QLabel("\t\t")
        end_label = QLabel("Término:")
        self.end_time_edit = QTimeEdit()
        self.end_time_edit.setTime(self.default_end_time)
        self.end_time_edit.setFixedWidth(80)
        time_layout.addWidget(start_label)
        time_layout.addWidget(self.start_time_edit)
        time_layout.addWidget(tab_label)
        time_layout.addWidget(end_label)
        time_layout.addWidget(self.end_time_edit)
        time_layout.addStretch()
        main_layout.addLayout(time_layout)

        # Espaço de uma linha
        main_layout.addWidget(QLabel(""))

        # Verificação de usuário ativo
        user_group = QGroupBox("Verificação de Atividade do Usuário")
        user_layout = QVBoxLayout(user_group)

        user_info = QLabel("\nInatividade mínima necessária para simular:")
        user_layout.addWidget(user_info)

        # Barra imediatamente abaixo
        slider_layout = QHBoxLayout()
        slider_layout.setSpacing(5)
        self.user_timeout_slider = QSlider(Qt.Orientation.Horizontal)
        self.user_timeout_slider.setMinimum(5)  # 5 segundos mínimo
        self.user_timeout_slider.setMaximum(300)  # 5 minutos máximo
        self.user_timeout_slider.setValue(self.default_user_timeout)
        self.user_timeout_slider.valueChanged.connect(self.update_timeout_label)

        self.timeout_label = QLabel(f"{self.default_user_timeout} segundos\n")
        self.timeout_label.setMinimumWidth(80)

        slider_layout.addWidget(self.user_timeout_slider)
        slider_layout.addWidget(self.timeout_label)
        slider_layout.addStretch()
        user_layout.addLayout(slider_layout)

        # Explicação clara dos tempos - caixa maior
        explanation_text = QTextEdit()
        explanation_text.setMaximumHeight(160)
        explanation_text.setReadOnly(True)
        explanation_text.setText(
            "Intervalos\n\n"
            "• Intervalo: O programa 'acorda' e verifica se deve simular.'\n"
            "• Inatividade: Quando acorda só simula se usuário está parado há X segundos\n\n"
            "Exemplo: Intervalo=120s, Inatividade=60s\n"
            "12:00 → Acorda 12:02\n"
            "Se usuário parado há 60s+ = Simula, senão Cancela"
        )
        explanation_text.setStyleSheet(
            "background-color: #000000; color: #ffffff; border: 1px solid #000000;"
        )
        user_layout.addWidget(explanation_text)

        main_layout.addWidget(user_group)
        main_layout.addStretch()

        self.tab_widget.addTab(main_tab, "Principal")

        # Abas
        self.advanced_tab = AdvancedTab(self.tab_widget)
        self.tab_widget.addTab(self.advanced_tab, "Opções")

        self.log_tab = LogTab(self.tab_widget)
        self.tab_widget.addTab(self.log_tab, "Log")

        about_tab = AboutTab(self.tab_widget)
        self.tab_widget.addTab(about_tab, "About")

        # Botões
        button_layout = QHBoxLayout()
        self.toggle_button = QPushButton("Iniciar")
        self.toggle_button.clicked.connect(self.toggle_service)

        self.close_button = QPushButton("Fechar Completamente")
        self.close_button.clicked.connect(self.quit_application)

        self.minimize_button = QPushButton("Minimizar")
        self.minimize_button.clicked.connect(self.hide)

        button_layout.addWidget(self.toggle_button)
        button_layout.addWidget(self.close_button)
        button_layout.addWidget(self.minimize_button)
        layout.addLayout(button_layout)

    def update_timeout_label(self, value):
        """Atualiza label do timeout"""
        self.timeout_label.setText(f"{value} segundos")

    def setup_tray(self):
        """Configura o ícone na bandeja do sistema"""
        try:
            app_style = QApplication.style()
            icon = app_style.standardIcon(QStyle.StandardPixmap.SP_ComputerIcon)
            self.tray_icon = QSystemTrayIcon(icon, self)
            self.tray_icon.setToolTip("Keep Alive RDP Connection")

            menu = QMenu()
            show_action = QAction("Mostrar", self)
            show_action.triggered.connect(self.show)
            self.toggle_tray_action = QAction("Iniciar", self)
            self.toggle_tray_action.triggered.connect(self.toggle_service)

            close_action = QAction("Fechar", self)
            close_action.triggered.connect(self.quit_application)

            menu.addAction(show_action)
            menu.addAction(self.toggle_tray_action)
            menu.addSeparator()
            menu.addAction(close_action)

            self.tray_icon.setContextMenu(menu)
            self.tray_icon.activated.connect(
                lambda reason: (
                    self.show()
                    if reason == QSystemTrayIcon.ActivationReason.DoubleClick
                    else None
                )
            )
            self.tray_icon.show()
        except Exception as e:
            logging.error(f"Erro ao configurar bandeja: {str(e)}")

    def load_settings(self):
        """Carrega configurações salvas"""
        self.advanced_tab.minimize_to_tray_cb.setChecked(
            self.settings.value("minimize_to_tray", True, bool)
        )

        # Carrega configurações avançadas
        self.advanced_tab.enable_mouse.setChecked(
            self.settings.value("enable_mouse", True, bool)
        )
        self.advanced_tab.enable_keyboard.setChecked(
            self.settings.value("enable_keyboard", True, bool)
        )
        self.advanced_tab.random_intervals.setChecked(
            self.settings.value("random_intervals", True, bool)
        )

    def save_settings(self):
        """Salva configurações"""
        self.settings.setValue("interval", self.interval_spin.value())
        self.settings.setValue("user_timeout", self.user_timeout_slider.value())
        self.settings.setValue("start_time", self.start_time_edit.time())
        self.settings.setValue("end_time", self.end_time_edit.time())
        self.settings.setValue(
            "minimize_to_tray", self.advanced_tab.minimize_to_tray_cb.isChecked()
        )

        # Configurações avançadas
        self.settings.setValue(
            "enable_mouse", self.advanced_tab.enable_mouse.isChecked()
        )
        self.settings.setValue(
            "enable_keyboard", self.advanced_tab.enable_keyboard.isChecked()
        )
        self.settings.setValue(
            "random_intervals", self.advanced_tab.random_intervals.isChecked()
        )

        self.log_tab.add_log("Configurações salvas")

    def check_teams_status(self):
        """Verifica status do Teams e loga se necessário"""
        try:
            teams_status, teams_message = get_teams_status()
            self.teams_check_count += 1

            # Print simples para debug (posteriormente comentar)
            print(f"DEBUG TEAMS: {teams_status} - {teams_message}")

            # Só loga se tiver informação confiável
            if teams_status and teams_status != "INDETERMINADO":
                # Para status importantes, loga sempre
                if teams_status in ["AUSENTE", "OCUPADO", "INATIVO"]:
                    self.log_tab.add_log(
                        f"Status Teams: {teams_status} - {teams_message}"
                    )
                # Para status normal (disponível/ativo), loga a cada 4 verificações
                elif (
                    teams_status in ["DISPONÍVEL", "ATIVO", "DETECTADO"]
                    and self.teams_check_count % 4 == 0
                ):
                    self.log_tab.add_log(
                        f"Status Teams: {teams_status} - {teams_message}"
                    )
            # Se teams_status é None ou INDETERMINADO, não loga pois não é confiável

        except Exception:
            # Em caso de erro, não loga status não confiável
            pass

    def closeEvent(self, event):
        self.save_settings()
        if self.advanced_tab.minimize_to_tray_cb.isChecked():
            event.ignore()
            self.hide()
            self.tray_icon.showMessage(
                "Keep Alive RDP Connection",
                "Rodando em segundo plano",
                QSystemTrayIcon.MessageIcon.Information,
                2000,
            )
        else:
            self.quit_application()

    def toggle_service(self):
        if not self.is_running:
            self.start_service()
        else:
            self.stop_service()

    def start_service(self):
        """Inicia o serviço"""
        base_interval = self.interval_spin.value()
        base_interval_ms = base_interval * 1000

        # Aplica variação se habilitado
        if self.advanced_tab.random_intervals.isChecked():
            variation = base_interval_ms * 0.3
            next_interval = random.randint(
                int(base_interval_ms - variation), int(base_interval_ms + variation)
            )
            log_msg = f"Serviço iniciado - intervalo: {base_interval}s (±30%) = {next_interval/1000:.1f}s"
        else:
            next_interval = base_interval_ms
            log_msg = f"Serviço iniciado - intervalo fixo: {base_interval}s"

        # Calcula próxima execução
        next_time = datetime.now() + timedelta(milliseconds=next_interval)
        log_msg += f" - Próxima atividade: {next_time.strftime('%H:%M:%S')}"

        self.activity_timer.start(next_interval)
        self.is_running = True
        self.toggle_button.setText("Parar")
        self.toggle_tray_action.setText("Parar")
        self.status_label.setText("Serviço rodando")

        self.log_tab.add_log(log_msg)

    def stop_service(self):
        """Para o serviço"""
        self.activity_timer.stop()
        self.is_running = False
        self.toggle_button.setText("Iniciar")
        self.toggle_tray_action.setText("Iniciar")
        self.status_label.setText("Serviço parado")
        ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)
        self.log_tab.add_log("Serviço parado")

    def perform_activity(self):
        try:
            # Verificar se usuário está ativo
            idle_time = get_user_activity_timeout()
            timeout_limit = self.user_timeout_slider.value()

            if idle_time < timeout_limit:
                # Usuário ativo - cancela simulação
                self.activity_count += 1
                self.log_tab.add_log(
                    f"CANCELADA: Usuário ativo (inatividade: {idle_time:.1f}s < {timeout_limit}s)"
                )

                # Calcula próximo intervalo mesmo assim
                if self.is_running:
                    self.schedule_next_activity()
                return

            # Prevenir bloqueio sempre
            prevent_system_lock()

            # Simular atividade
            activity_success, activity_message = simulate_safe_activity()

            # Atualizar contadores
            self.activity_count += 1
            now = datetime.now().strftime("%H:%M:%S")
            status_msg = f"Atividade #{self.activity_count} em {now}"

            if activity_success:
                self.status_label.setText(status_msg + " (Sucesso)")
                self.log_tab.add_log("Simulação executada com sucesso")
            else:
                self.status_label.setText(status_msg + f" ({activity_message})")
                self.log_tab.add_log(f"ERRO: {activity_message}")

            # Agendar próxima atividade
            if self.is_running:
                self.schedule_next_activity()

        except Exception as e:
            error_msg = f"Erro na atividade #{self.activity_count}: {str(e)}"
            self.status_label.setText(error_msg)
            self.log_tab.add_log(f"ERRO CRÍTICO: {error_msg}")

    def schedule_next_activity(self):
        """Agenda próxima atividade com log do horário"""
        self.activity_timer.stop()  # Para o timer atual

        base_interval = self.interval_spin.value() * 1000

        if self.advanced_tab.random_intervals.isChecked():
            variation = base_interval * 0.3
            next_interval = random.randint(
                int(base_interval - variation), int(base_interval + variation)
            )
            interval_text = (
                f"{self.interval_spin.value()}s (±30%) = {next_interval/1000:.1f}s"
            )
        else:
            next_interval = base_interval
            interval_text = f"{self.interval_spin.value()}s fixo"

        self.activity_timer.setInterval(next_interval)
        self.activity_timer.start()  # Reinicia com novo intervalo

        # Calcula e loga próxima execução
        next_time = datetime.now() + timedelta(milliseconds=next_interval)
        self.log_tab.add_log(
            f"Próxima atividade agendada para: {next_time.strftime('%H:%M:%S')} ({interval_text})"
        )

    def quit_application(self):
        try:
            self.save_settings()
            self.log_tab.add_log("Aplicativo encerrado")
        except Exception:
            pass

        try:
            self.activity_timer.stop()
            self.teams_timer.stop()
            self.tray_icon.hide()
            cleanup_lock()
        except Exception:
            pass

        QApplication.quit()


def main():
    """Função principal"""
    app = QApplication(sys.argv)

    # Evita mensagens de DPI - REMOVIDO configurações que davam erro
    try:
        app.setStyle("Windows")
    except Exception:
        pass

    window = KeepAliveApp()
    window.show()

    # Mensagem de boas-vindas
    window.log_tab.add_log("Keep Alive RDP Connection iniciado")
    window.log_tab.add_log("Auto-start configurado para 5 segundos")

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
