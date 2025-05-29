# pylint: disable=E0602,E0102,E1101

"""
Keep Alive Manager - RDP & Teams 2.0.1
Foco principal em manter conexão RDP ativa e status do Teams como "Disponível"
"""

import ctypes
import logging
import os
import random
import sys
import time
from datetime import datetime

import pyautogui
import win32api
import win32con
import win32event
from PyQt6.QtCore import QSettings, Qt, QTime, QTimer
from PyQt6.QtGui import QAction, QIcon
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
    QSpinBox,
    QStyle,
    QSystemTrayIcon,
    QTabWidget,
    QTextEdit,
    QTimeEdit,
    QVBoxLayout,
    QWidget,
)

# Configurações para prevenir bloqueios de tela
ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001
ES_DISPLAY_REQUIRED = 0x00000002
ES_AWAYMODE_REQUIRED = 0x00000040

# Configuração de logging
log_file = os.path.join(os.path.expanduser("~"), "keep_alive_manager.log")
logging.basicConfig(
    filename=log_file,
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)

# Verificação de instância única
mutex = None


def is_already_running():
    global mutex
    app_name = "KeepAliveManager_RDP_Teams"
    try:
        mutex = win32event.CreateMutex(None, False, app_name)
        return win32api.GetLastError() == win32con.ERROR_ALREADY_EXISTS
    except:
        return False


def prevent_system_lock():
    """Previne bloqueio de tela e hibernação de forma mais eficaz"""
    try:
        ctypes.windll.kernel32.SetThreadExecutionState(
            ES_CONTINUOUS
            | ES_SYSTEM_REQUIRED
            | ES_DISPLAY_REQUIRED
            | ES_AWAYMODE_REQUIRED
        )
        return True
    except Exception as e:
        logging.error(f"Erro ao prevenir bloqueio: {str(e)}")
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
    except Exception as e:
        logging.error(f"Erro ao verificar atividade do usuário: {str(e)}")
        return 0


def simulate_effective_activity():
    """
    Simula atividade eficaz para manter RDP e Teams ativos,
    evitando interferência com atividades reais do usuário
    """
    try:
        # Verifica se o usuário está ativo recentemente
        idle_time = get_user_activity_timeout()
        if idle_time < 30:  # Usuário ativo nos últimos 30 segundos
            logging.info(
                f"Usuário ativo (inativo por {idle_time:.1f}s) - pulando simulação"
            )
            return False, "Usuário ativo - pulando simulação"

        # 1. Movimento de mouse significativo mas discreto
        screen_width, screen_height = pyautogui.size()
        target_x = random.randint(100, screen_width - 100)
        target_y = random.randint(100, screen_height - 100)
        pyautogui.moveTo(target_x, target_y, duration=0.3)

        # 2. Eventos de teclado seguros (não interferem com o usuário)
        keys = ["shift", "ctrl", "numlock"]
        key = random.choice(keys)
        pyautogui.press(key)

        # 3. Pequeno movimento adicional
        pyautogui.moveRel(random.randint(-3, 3), random.randint(-3, 3), duration=0.1)

        return True, "Atividade simulada com sucesso"
    except Exception as e:
        logging.error(f"Erro na simulação de atividade: {str(e)}")
        return False, f"Erro: {str(e)}"


def is_teams_active():
    """Verifica se o Teams está com status 'Disponível'"""
    try:
        from win32gui import GetForegroundWindow, GetWindowText

        active_window = GetWindowText(GetForegroundWindow()).lower()

        # Verifica se o Teams está em foco (indicativo de atividade)
        if "teams" in active_window:
            return True

        # Verificação adicional: se o mouse está sobre a janela do Teams
        teams_rect = None
        for window in pyautogui.getAllWindows():
            if "teams" in window.title.lower():
                teams_rect = (window.left, window.top, window.width, window.height)
                break

        if teams_rect:
            x, y = pyautogui.position()
            if (
                teams_rect[0] <= x <= teams_rect[0] + teams_rect[2]
                and teams_rect[1] <= y <= teams_rect[1] + teams_rect[3]
            ):
                return True

        return False
    except Exception as e:
        logging.warning(f"Erro ao verificar status do Teams: {str(e)}")
        return True


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
        """Adiciona mensagem ao log"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.log_text.append(f"[{timestamp}] {message}")
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )
        logging.info(message)

    def clear_log(self):
        """Limpa o log visual"""
        self.log_text.clear()

    def save_log(self):
        """Salva o log em um arquivo"""
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
    """Aba para configurações avançadas de atividade"""

    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)

        # Grupo de opções de simulação
        sim_group = QGroupBox("Otimizações de Atividade")
        sim_layout = QVBoxLayout(sim_group)

        self.enable_mouse = QCheckBox("Ativar movimentos do mouse")
        self.enable_mouse.setChecked(True)
        sim_layout.addWidget(self.enable_mouse)

        self.enable_keyboard = QCheckBox("Ativar eventos de teclado")
        self.enable_keyboard.setChecked(True)
        sim_layout.addWidget(self.enable_keyboard)

        self.avoid_user_activity = QCheckBox("Evitar simulação quando usuário ativo")
        self.avoid_user_activity.setChecked(True)
        sim_layout.addWidget(self.avoid_user_activity)

        self.random_intervals = QCheckBox("Usar intervalos variáveis (±30%)")
        self.random_intervals.setChecked(True)
        sim_layout.addWidget(self.random_intervals)

        layout.addWidget(sim_group)

        # Botão de teste
        test_button = QPushButton("Testar Simulação Agora")
        test_button.clicked.connect(self.test_simulation)
        layout.addWidget(test_button)

        # Espaçador
        layout.addStretch()

    def test_simulation(self):
        """Testa a simulação de atividade uma vez"""
        try:
            success, message = simulate_effective_activity()
            if success:
                logging.info("Teste de simulação executado com sucesso")
                if self.parent().parent().log_tab:
                    self.parent().parent().log_tab.add_log(
                        "Teste de simulação executado com sucesso"
                    )
            else:
                logging.warning(f"Teste de simulação pulado: {message}")
                if self.parent().parent().log_tab:
                    self.parent().parent().log_tab.add_log(f"Teste pulado: {message}")
        except Exception as e:
            logging.error(f"Erro no teste de simulação: {str(e)}")
            if self.parent().parent().log_tab:
                self.parent().parent().log_tab.add_log(f"Erro no teste: {str(e)}")


class KeepAliveApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Keep Alive Manager - RDP & Teams")
        self.setFixedSize(500, 500)

        # Verifica se já está rodando
        if is_already_running():
            QMessageBox.warning(
                None,
                "Aviso",
                "O Keep Alive Manager já está em execução!\n"
                "Verifique a bandeja do sistema ou o gerenciador de tarefas.",
            )
            sys.exit(1)

        # Configurações padrão
        self.settings = QSettings("KeepAliveTools", "KeepAliveManager")
        self.default_interval = self.settings.value("interval", 30, int)  # 30s padrão
        self.default_start_time = self.settings.value("start_time", QTime(8, 45), QTime)
        self.default_end_time = self.settings.value("end_time", QTime(17, 15), QTime)

        self.is_running = False
        self.activity_count = 0
        self.completely_stopped = self.settings.value("completely_stopped", False, bool)

        # Timers
        self.activity_timer = QTimer()
        self.activity_timer.timeout.connect(self.perform_activity)

        self.schedule_timer = QTimer()
        self.schedule_timer.timeout.connect(self.check_schedule)
        self.schedule_timer.start(60000)  # Verifica agendamento a cada minuto

        # Status check timer
        self.status_timer = QTimer()
        self.status_timer.timeout.connect(self.update_status)
        self.status_timer.start(5000)

        logging.info("Keep Alive Manager iniciado")

        self.setup_ui()
        self.setup_tray()
        self.load_settings()

    def setup_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # Status
        status_frame = QFrame()
        status_frame.setFrameShape(QFrame.Shape.StyledPanel)
        status_layout = QVBoxLayout(status_frame)

        self.status_label = QLabel("Aguardando início do serviço")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        status_layout.addWidget(self.status_label)

        self.teams_status = QLabel("Status Teams: Verificando...")
        status_layout.addWidget(self.teams_status)

        self.idle_status = QLabel("Inatividade do usuário: Verificando...")
        status_layout.addWidget(self.idle_status)

        layout.addWidget(status_frame)

        # Abas
        self.tab_widget = QTabWidget()
        layout.addWidget(self.tab_widget)

        # Aba principal
        main_tab = QWidget()
        main_layout = QVBoxLayout(main_tab)

        # Configurações básicas
        config_frame = QFrame()
        config_frame.setFrameShape(QFrame.Shape.StyledPanel)
        config_layout = QVBoxLayout(config_frame)

        interval_layout = QHBoxLayout()
        interval_label = QLabel("Intervalo (segundos):")
        self.interval_spin = QSpinBox()
        self.interval_spin.setRange(15, 180)  # Intervalo mais curto
        self.interval_spin.setValue(self.default_interval)  # Valor padrão 30s
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

        self.auto_start_cb = QCheckBox("Iniciar automaticamente na abertura")
        self.auto_start_cb.setChecked(True)  # Ativado por padrão
        config_layout.addWidget(self.auto_start_cb)

        main_layout.addWidget(config_frame)

        self.tab_widget.addTab(main_tab, "Principal")

        # Aba de configurações avançadas
        self.advanced_tab = AdvancedTab(self.tab_widget)
        self.tab_widget.addTab(self.advanced_tab, "Otimizações")

        # Aba de log
        self.log_tab = LogTab(self.tab_widget)
        self.tab_widget.addTab(self.log_tab, "Log")

        # Botões
        button_layout = QHBoxLayout()
        self.toggle_button = QPushButton("Iniciar")
        self.toggle_button.clicked.connect(self.toggle_service)

        self.minimize_button = QPushButton("Minimizar")
        self.minimize_button.clicked.connect(self.hide)

        self.stop_completely_button = QPushButton("Parar Completamente")
        self.stop_completely_button.clicked.connect(self.stop_completely)

        button_layout.addWidget(self.toggle_button)
        button_layout.addWidget(self.minimize_button)
        button_layout.addWidget(self.stop_completely_button)
        layout.addLayout(button_layout)

    def setup_tray(self):
        icon = self.style().standardIcon(QStyle.StandardPixmap.SP_ComputerIcon)
        self.tray_icon = QSystemTrayIcon(icon, self)
        self.tray_icon.setToolTip("Keep Alive Manager - RDP & Teams")

        menu = QMenu()
        show_action = QAction("Mostrar", self)
        show_action.triggered.connect(self.show)
        self.toggle_tray_action = QAction("Iniciar", self)
        self.toggle_tray_action.triggered.connect(self.toggle_service)

        stop_completely_action = QAction("Parar Completamente", self)
        stop_completely_action.triggered.connect(self.stop_completely)

        quit_action = QAction("Sair", self)
        quit_action.triggered.connect(self.quit_application)

        menu.addAction(show_action)
        menu.addAction(self.toggle_tray_action)
        menu.addAction(stop_completely_action)
        menu.addSeparator()
        menu.addAction(quit_action)

        self.tray_icon.setContextMenu(menu)
        self.tray_icon.activated.connect(
            lambda reason: (
                self.show()
                if reason == QSystemTrayIcon.ActivationReason.DoubleClick
                else None
            )
        )
        self.tray_icon.show()

    def load_settings(self):
        """Carrega configurações salvas"""
        self.minimize_to_tray_cb.setChecked(
            self.settings.value("minimize_to_tray", True, bool)
        )
        self.auto_start_cb.setChecked(self.settings.value("auto_start", True, bool))
        self.completely_stopped = self.settings.value("completely_stopped", False, bool)

        # Carrega configurações avançadas
        self.advanced_tab.enable_mouse.setChecked(
            self.settings.value("enable_mouse", True, bool)
        )
        self.advanced_tab.enable_keyboard.setChecked(
            self.settings.value("enable_keyboard", True, bool)
        )
        self.advanced_tab.avoid_user_activity.setChecked(
            self.settings.value("avoid_user_activity", True, bool)
        )
        self.advanced_tab.random_intervals.setChecked(
            self.settings.value("random_intervals", True, bool)
        )

        # Aplica parada completa se estiver marcada
        if self.completely_stopped:
            self.status_label.setText("Serviço completamente parado até amanhã")
            # Não desativamos os botões para permitir reinício manual
        # Inicia automaticamente se configurado e dentro do horário
        elif self.auto_start_cb.isChecked():
            # Verifica se está dentro do horário agendado
            now = QTime.currentTime()
            start_time = self.start_time_edit.time()
            end_time = self.end_time_edit.time()

            if start_time <= now <= end_time:
                QTimer.singleShot(1000, self.toggle_service)
            else:
                self.log_tab.add_log(
                    "Fora do horário agendado - aguardando início automático"
                )

    def save_settings(self):
        """Salva configurações atuais"""
        self.settings.setValue("interval", self.interval_spin.value())
        self.settings.setValue("start_time", self.start_time_edit.time())
        self.settings.setValue("end_time", self.end_time_edit.time())
        self.settings.setValue("minimize_to_tray", self.minimize_to_tray_cb.isChecked())
        self.settings.setValue("auto_start", self.auto_start_cb.isChecked())
        self.settings.setValue("completely_stopped", self.completely_stopped)

        # Salva configurações avançadas
        self.settings.setValue(
            "enable_mouse", self.advanced_tab.enable_mouse.isChecked()
        )
        self.settings.setValue(
            "enable_keyboard", self.advanced_tab.enable_keyboard.isChecked()
        )
        self.settings.setValue(
            "avoid_user_activity", self.advanced_tab.avoid_user_activity.isChecked()
        )
        self.settings.setValue(
            "random_intervals", self.advanced_tab.random_intervals.isChecked()
        )

        self.log_tab.add_log("Configurações salvas com sucesso")

    def closeEvent(self, event):
        self.save_settings()
        if self.minimize_to_tray_cb.isChecked():
            event.ignore()
            self.hide()
            self.tray_icon.showMessage(
                "Keep Alive - RDP & Teams",
                "Rodando em segundo plano",
                QSystemTrayIcon.MessageIcon.Information,
                2000,
            )
        else:
            self.quit_application()

    def toggle_service(self):
        if not self.is_running:
            # Verifica se está dentro do horário agendado
            now = QTime.currentTime()
            start_time = self.start_time_edit.time()
            end_time = self.end_time_edit.time()

            if start_time <= now <= end_time:
                self.start_service()
            else:
                # Fora do horário - pede confirmação
                reply = QMessageBox.question(
                    self,
                    "Fora do horário agendado",
                    f"Você está tentando iniciar fora do horário agendado ({start_time.toString('HH:mm')} - {end_time.toString('HH:mm')}).\nDeseja iniciar mesmo assim?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.No,
                )

                if reply == QMessageBox.StandardButton.Yes:
                    self.start_service()
                else:
                    self.log_tab.add_log(
                        "Início fora do horário cancelado pelo usuário"
                    )
                    self.status_label.setText("Início cancelado (fora do horário)")
        else:
            self.stop_service()

    def start_service(self):
        """Inicia o serviço de keep-alive"""
        interval = self.interval_spin.value() * 1000

        # Aplica intervalo randômico se habilitado
        if self.advanced_tab.random_intervals.isChecked():
            variation = interval * 0.3  # 30% de variação
            interval = random.randint(
                int(interval - variation), int(interval + variation)
            )

        self.activity_timer.start(interval)
        self.is_running = True
        self.toggle_button.setText("Parar")
        self.toggle_tray_action.setText("Parar")
        self.status_label.setText("Serviço iniciado")
        self.log_tab.add_log(
            f"Serviço iniciado com intervalo de {interval/1000:.1f} segundos"
        )

    def stop_service(self):
        """Para o serviço de keep-alive"""
        self.activity_timer.stop()
        self.is_running = False
        self.toggle_button.setText("Iniciar")
        self.toggle_tray_action.setText("Iniciar")
        self.status_label.setText("Serviço parado")
        ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)
        self.log_tab.add_log("Serviço parado")

    def stop_completely(self):
        """Para completamente o serviço até o próximo dia"""
        if self.is_running:
            self.stop_service()  # Para o serviço se estiver rodando

        self.completely_stopped = True
        self.status_label.setText("Serviço completamente parado até amanhã")

        # Salva a hora atual para resetar amanhã
        self.settings.setValue("completely_stopped_time", datetime.now().isoformat())

        self.log_tab.add_log("Serviço completamente parado até amanhã")
        self.save_settings()

        # Mostra mensagem no tray
        self.tray_icon.showMessage(
            "Keep Alive - Parado",
            "Serviço completamente parado até amanhã",
            QSystemTrayIcon.MessageIcon.Information,
            3000,
        )

    def perform_activity(self):
        try:
            # 1. Prevenir bloqueio do sistema (sempre ativo)
            prevent_system_lock()

            # 2. Simular atividade eficaz
            activity_success, activity_message = simulate_effective_activity()

            # 3. Atualizar status e contadores
            self.activity_count += 1
            now = datetime.now().strftime("%H:%M:%S")
            status_msg = f"Atividade #{self.activity_count} em {now}"

            if activity_success:
                self.status_label.setText(status_msg + " (Sucesso)")
                if self.activity_count % 5 == 0:
                    self.log_tab.add_log(status_msg)
            else:
                self.status_label.setText(status_msg + f" ({activity_message})")
                if "pulando" not in activity_message:
                    self.log_tab.add_log(f"ATENÇÃO: {activity_message}")

            # 4. Programar próxima atividade com variação
            if self.is_running and self.advanced_tab.random_intervals.isChecked():
                interval = self.interval_spin.value() * 1000
                variation = interval * 0.3
                next_interval = random.randint(
                    int(interval - variation), int(interval + variation)
                )
                self.activity_timer.setInterval(next_interval)

        except Exception as e:
            error_msg = f"Erro na atividade #{self.activity_count}: {str(e)}"
            self.status_label.setText(error_msg)
            self.log_tab.add_log(f"ERRO: {error_msg}")
            logging.error(error_msg)

    def update_status(self):
        """Atualiza status do Teams e inatividade do usuário"""
        try:
            teams_active = is_teams_active()
            status = (
                "Ativo (Disponível)" if teams_active else "Inativo (Pode estar ausente)"
            )
            self.teams_status.setText(f"Status Teams: {status}")

            # Atualiza status de inatividade
            idle_time = get_user_activity_timeout()
            self.idle_status.setText(
                f"Inatividade do usuário: {idle_time:.1f} segundos"
            )
        except:
            self.teams_status.setText("Status Teams: Verificação indisponível")
            self.idle_status.setText("Inatividade: Verificação indisponível")

    def check_schedule(self):
        """Verifica se está dentro do horário agendado e se a parada completa já expirou"""
        # Verifica se a parada completa deve ser resetada (novo dia)
        if self.completely_stopped:
            stopped_time_str = self.settings.value("completely_stopped_time", "")
            if stopped_time_str:
                try:
                    stopped_time = datetime.fromisoformat(stopped_time_str)
                    if datetime.now().date() > stopped_time.date():
                        self.completely_stopped = False
                        self.log_tab.add_log("Parada completa resetada - novo dia")
                        self.save_settings()
                except ValueError:
                    pass

        # Se estiver completamente parado, não inicia automaticamente
        if self.completely_stopped:
            return

        # Verifica o agendamento normal
        now = QTime.currentTime()
        start_time = self.start_time_edit.time()
        end_time = self.end_time_edit.time()

        within_schedule = start_time <= now <= end_time

        if within_schedule and not self.is_running:
            self.log_tab.add_log("Iniciando serviço conforme agendamento")
            self.toggle_service()  # Vai iniciar (dentro do horário, sem confirmação)
        elif not within_schedule and self.is_running:
            self.log_tab.add_log("Parando serviço conforme agendamento")
            self.stop_service()

    def quit_application(self):
        self.save_settings()
        self.activity_timer.stop()
        self.status_timer.stop()
        self.schedule_timer.stop()
        self.tray_icon.hide()
        self.log_tab.add_log("Aplicativo encerrado")
        logging.info("Keep Alive Manager encerrado")
        QApplication.quit()


def main():
    app = QApplication(sys.argv)
    if hasattr(Qt.ApplicationAttribute, "AA_EnableHighDpiScaling"):
        app.setAttribute(Qt.ApplicationAttribute.AA_EnableHighDpiScaling, True)
    if hasattr(Qt.ApplicationAttribute, "AA_UseHighDpiPixmaps"):
        app.setAttribute(Qt.ApplicationAttribute.AA_UseHighDpiPixmaps, True)
    os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"
    app.setStyle("Windows")

    # Tenta criar o diretório de logs
    try:
        log_dir = os.path.dirname(log_file)
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
    except:
        pass

    window = KeepAliveApp()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
