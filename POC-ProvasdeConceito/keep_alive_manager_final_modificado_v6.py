# pylint: disable=E0602,E0102,E1101

"""
Keep Alive Manager - RDP & Teams
Versão aprimorada e funcional para Python 3.13.
Implementa:
- Verificação de sessão RDP ativa
- Prevenção de bloqueio de tela
- Simulação de atividade do mouse e teclado
- Movimentos randômicos para parecer mais natural
- Log de atividades
- Persistência de configurações
- Verificação específica para Teams
- GUI funcional via PyQt6
"""

import sys
import os
import time
import ctypes
import random
import json
import logging
from datetime import datetime
import pyautogui
import win32ts
import win32api

# Adicione psutil às dependências - fácil de instalar com:
# pip install psutil
try:
    import psutil
except ImportError:
    logging.warning("psutil não encontrado. Algumas funções podem não funcionar.")
    psutil = None
from PyQt6.QtCore import Qt, QTimer, QTime, QCoreApplication, QSettings
from PyQt6.QtGui import QIcon, QAction
from PyQt6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QSpinBox,
    QCheckBox,
    QPushButton,
    QSystemTrayIcon,
    QMenu,
    QStyle,
    QFrame,
    QTimeEdit,
    QTabWidget,
    QTextEdit,
    QComboBox,
    QGroupBox,
    QRadioButton,
)

# Configurações para prevenir bloqueios de tela
ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001
ES_DISPLAY_REQUIRED = 0x00000002
ES_AWAYMODE_REQUIRED = 0x00000040  # Previne hibernação

# Constantes para simular entrada de teclado
VK_SHIFT = 0x10
KEYEVENTF_KEYUP = 0x0002

# Configuração de logging
log_file = os.path.join(os.path.expanduser("~"), "keep_alive_manager.log")
logging.basicConfig(
    filename=log_file,
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)


def is_rdp_active():
    """
    Verifica se a sessão RDP está ativa - versão simplificada
    Não é 100% confiável, então é melhor manter ativo sempre
    """
    try:
        # Verifica se o serviço de terminal server está em execução
        import win32serviceutil

        status = win32serviceutil.QueryServiceStatus("TermService")[1]
        is_running = status == 4  # 4 = running

        # Tenta verificar a existência de conexões RDP (apenas para logging)
        # Isso não deve bloquear a funcionalidade principal
        try:
            import subprocess

            result = subprocess.run(
                "query session", capture_output=True, text=True, shell=True
            )
            rdp_connections = "rdp-tcp" in result.stdout.lower()
            logging.info(f"RDP connections detected: {rdp_connections}")
        except:
            pass

        return is_running
    except Exception as e:
        logging.warning(f"Não foi possível verificar status RDP: {str(e)}")
        return True  # Assume ativo em caso de erro para evitar problemas


def is_teams_running():
    """
    Verificação simplificada para o Teams que não quebra o código
    Apenas verifica se o processo está em execução, sem tentar acessar janelas
    """
    try:
        import psutil

        teams_processes = [
            p
            for p in psutil.process_iter(["name"])
            if p.info["name"] and "teams" in p.info["name"].lower()
        ]
        return len(teams_processes) > 0
    except Exception as e:
        logging.warning(f"Erro ao verificar Teams (não crítico): {str(e)}")
        return True  # Assume que está executando para não bloquear funcionalidade


def simulate_random_activity():
    """Simula atividade aleatória mais natural, mas não intrusiva"""
    actions = [
        # Movimentos de mouse sutis
        lambda: pyautogui.moveRel(random.randint(-3, 3), random.randint(-3, 3)),
        lambda: pyautogui.moveRel(random.randint(-1, 1), random.randint(-1, 1)),
        # Pressionar Shift (invisível)
        lambda: simulate_key_press(),
    ]

    # Executa apenas 1-2 ações aleatórias para não ser intrusivo
    for _ in range(random.randint(1, 2)):
        random.choice(actions)()
        time.sleep(random.uniform(0.1, 0.2))


def simulate_key_press():
    """Simula pressionamento de tecla shift (invisível)"""
    try:
        # Pressiona Shift (invisível para o usuário)
        win32api.keybd_event(VK_SHIFT, 0, 0, 0)
        time.sleep(0.05)
        win32api.keybd_event(VK_SHIFT, 0, KEYEVENTF_KEYUP, 0)
    except Exception as e:
        logging.error(f"Erro ao simular tecla: {str(e)}")


def simulate_alt_tab():
    """Simula Alt+Tab rapidamente (para alternar aplicativo em foco)"""
    try:
        pyautogui.keyDown("alt")
        pyautogui.press("tab")
        time.sleep(0.05)  # Muito rápido para ser visível
        pyautogui.press("tab")  # Volta para o aplicativo original
        pyautogui.keyUp("alt")
    except Exception as e:
        logging.error(f"Erro ao simular Alt+Tab: {str(e)}")


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
        # Rola para o final
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )
        # Também registra no arquivo de log
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
    """Aba para configurações avançadas"""

    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)

        # Grupo de opções de simulação
        sim_group = QGroupBox("Opções de Simulação")
        sim_layout = QVBoxLayout(sim_group)

        self.simulate_mouse = QCheckBox("Simular movimentos do mouse")
        self.simulate_mouse.setChecked(True)
        sim_layout.addWidget(self.simulate_mouse)

        self.simulate_keyboard = QCheckBox("Simular pressionamentos de teclas")
        self.simulate_keyboard.setChecked(True)
        sim_layout.addWidget(self.simulate_keyboard)

        self.simulate_alt_tab = QCheckBox("Simular Alt+Tab ocasionalmente")
        self.simulate_alt_tab.setChecked(True)
        sim_layout.addWidget(self.simulate_alt_tab)

        self.random_intervals = QCheckBox("Usar intervalos randômicos (±20%)")
        self.random_intervals.setChecked(True)
        sim_layout.addWidget(self.random_intervals)

        layout.addWidget(sim_group)

        # Grupo de opções de verificação
        check_group = QGroupBox("Verificações")
        check_layout = QVBoxLayout(check_group)

        self.check_rdp = QCheckBox("Verificar sessão RDP ativa")
        self.check_rdp.setChecked(True)
        check_layout.addWidget(self.check_rdp)

        self.check_teams = QCheckBox("Verificar Teams em execução")
        self.check_teams.setChecked(True)
        check_layout.addWidget(self.check_teams)

        layout.addWidget(check_group)

        # Botão de teste
        test_button = QPushButton("Testar Simulação (Uma vez)")
        test_button.clicked.connect(self.test_simulation)
        layout.addWidget(test_button)

        # Espaçador
        layout.addStretch()

    def test_simulation(self):
        """Testa a simulação de atividade uma vez"""
        try:
            simulate_random_activity()
            logging.info("Teste de simulação executado com sucesso")
            if self.parent().parent().log_tab:
                self.parent().parent().log_tab.add_log(
                    "Teste de simulação executado com sucesso"
                )
        except Exception as e:
            logging.error(f"Erro no teste de simulação: {str(e)}")
            if self.parent().parent().log_tab:
                self.parent().parent().log_tab.add_log(f"Erro no teste: {str(e)}")


class KeepAliveApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Keep Alive Manager")
        self.setFixedSize(500, 500)

        # Configurações padrão
        self.settings = QSettings("KeepAliveTools", "KeepAliveManager")
        self.default_interval = self.settings.value("interval", 120, int)
        self.default_start_time = self.settings.value("start_time", QTime(8, 45), QTime)
        self.default_end_time = self.settings.value("end_time", QTime(17, 15), QTime)

        self.is_running = False
        self.activity_count = 0
        self.last_teams_check = False

        # Timers
        self.activity_timer = QTimer()
        self.activity_timer.timeout.connect(self.perform_activity)

        self.schedule_timer = QTimer()
        self.schedule_timer.timeout.connect(self.check_schedule)
        self.schedule_timer.start(60000)  # Verifica agendamento a cada minuto

        # Status check timer (para atualizações na UI)
        self.status_timer = QTimer()
        self.status_timer.timeout.connect(self.update_status)
        self.status_timer.start(5000)  # Atualiza status a cada 5 segundos

        # Registro de inicialização
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

        self.rdp_status = QLabel("Status RDP: Verificando...")
        self.teams_status = QLabel("Status Teams: Verificando...")
        status_layout.addWidget(self.rdp_status)
        status_layout.addWidget(self.teams_status)

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

        self.auto_start_cb = QCheckBox("Iniciar automaticamente na abertura")
        self.auto_start_cb.setChecked(False)
        config_layout.addWidget(self.auto_start_cb)

        main_layout.addWidget(config_frame)

        self.tab_widget.addTab(main_tab, "Principal")

        # Aba de configurações avançadas
        self.advanced_tab = AdvancedTab(self.tab_widget)
        self.tab_widget.addTab(self.advanced_tab, "Avançado")

        # Configura valores padrão seguros para opções avançadas
        self.advanced_tab.simulate_mouse.setChecked(True)
        self.advanced_tab.simulate_keyboard.setChecked(True)
        self.advanced_tab.simulate_alt_tab.setChecked(
            False
        )  # Alt+Tab desativado por padrão
        self.advanced_tab.random_intervals.setChecked(True)

        # Aba de log
        self.log_tab = LogTab(self.tab_widget)
        self.tab_widget.addTab(self.log_tab, "Log")

        # Botões
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
        self.auto_start_cb.setChecked(self.settings.value("auto_start", False, bool))

        # Carrega configurações avançadas
        self.advanced_tab.simulate_mouse.setChecked(
            self.settings.value("simulate_mouse", True, bool)
        )
        self.advanced_tab.simulate_keyboard.setChecked(
            self.settings.value("simulate_keyboard", True, bool)
        )
        self.advanced_tab.simulate_alt_tab.setChecked(
            self.settings.value("simulate_alt_tab", True, bool)
        )
        self.advanced_tab.random_intervals.setChecked(
            self.settings.value("random_intervals", True, bool)
        )
        self.advanced_tab.check_rdp.setChecked(
            self.settings.value("check_rdp", True, bool)
        )
        self.advanced_tab.check_teams.setChecked(
            self.settings.value("check_teams", True, bool)
        )

        # Inicia automaticamente se configurado
        if self.auto_start_cb.isChecked():
            QTimer.singleShot(1000, self.toggle_service)

    def save_settings(self):
        """Salva configurações atuais"""
        self.settings.setValue("interval", self.interval_spin.value())
        self.settings.setValue("start_time", self.start_time_edit.time())
        self.settings.setValue("end_time", self.end_time_edit.time())
        self.settings.setValue("minimize_to_tray", self.minimize_to_tray_cb.isChecked())
        self.settings.setValue("auto_start", self.auto_start_cb.isChecked())

        # Salva configurações avançadas
        self.settings.setValue(
            "simulate_mouse", self.advanced_tab.simulate_mouse.isChecked()
        )
        self.settings.setValue(
            "simulate_keyboard", self.advanced_tab.simulate_keyboard.isChecked()
        )
        self.settings.setValue(
            "simulate_alt_tab", self.advanced_tab.simulate_alt_tab.isChecked()
        )
        self.settings.setValue(
            "random_intervals", self.advanced_tab.random_intervals.isChecked()
        )
        self.settings.setValue("check_rdp", self.advanced_tab.check_rdp.isChecked())
        self.settings.setValue("check_teams", self.advanced_tab.check_teams.isChecked())

        self.log_tab.add_log("Configurações salvas com sucesso")

    def closeEvent(self, event):
        self.save_settings()
        if self.minimize_to_tray_cb.isChecked():
            event.ignore()
            self.hide()
            self.tray_icon.showMessage("Keep Alive", "Rodando em segundo plano")
        else:
            self.quit_application()

    def toggle_service(self):
        if not self.is_running:
            # Inicia o serviço
            interval = self.interval_spin.value() * 1000  # Converte para milissegundos

            # Aplica intervalo randômico se habilitado
            if self.advanced_tab.random_intervals.isChecked():
                variation = interval * 0.2  # 20% de variação
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
        else:
            # Para o serviço
            self.activity_timer.stop()
            self.is_running = False
            self.toggle_button.setText("Iniciar")
            self.toggle_tray_action.setText("Iniciar")
            self.status_label.setText("Serviço parado")
            ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)
            self.log_tab.add_log("Serviço parado")

    def perform_activity(self):
        try:
            # Sempre ativar, independente de verificações
            # Prevenir bloqueio de tela e hibernação com opções mais robustas
            ctypes.windll.kernel32.SetThreadExecutionState(
                ES_CONTINUOUS
                | ES_SYSTEM_REQUIRED
                | ES_DISPLAY_REQUIRED
                | ES_AWAYMODE_REQUIRED
            )

            # Faz verificações apenas para exibição/log (não afeta funcionalidade principal)
            rdp_status = True
            teams_status = True

            if self.advanced_tab.check_rdp.isChecked():
                rdp_status = is_rdp_active()

            if self.advanced_tab.check_teams.isChecked() and psutil:
                teams_status = is_teams_running()
                if teams_status != self.last_teams_check:
                    self.last_teams_check = teams_status
                    self.log_tab.add_log(
                        f"Status do Teams alterado: {'Em execução' if teams_status else 'Não detectado'}"
                    )

            # Simular atividade SEMPRE (independente das verificações)
            # Isso garante que o sistema permanecerá ativo mesmo se as verificações falharem
            if (
                self.advanced_tab.simulate_mouse.isChecked()
                or self.advanced_tab.simulate_keyboard.isChecked()
            ):
                simulate_random_activity()
            else:
                # Movimento mínimo para manter ativo
                pyautogui.moveRel(1, 0)
                pyautogui.moveRel(-1, 0)

            # Atualiza contadores e status
            self.activity_count += 1
            now = datetime.now().strftime("%H:%M:%S")

            # Status diferenciado apenas para feedback visual
            status_suffix = ""
            if not rdp_status:
                status_suffix = " (RDP não detectado, mas continuando ativo)"
            elif not teams_status and self.advanced_tab.check_teams.isChecked():
                status_suffix = " (Teams não detectado, mas continuando ativo)"

            self.status_label.setText(
                f"Última atividade: {now} (#{self.activity_count}){status_suffix}"
            )

            # Log a cada 10 atividades para não sobrecarregar
            if self.activity_count % 10 == 0:
                self.log_tab.add_log(
                    f"Atividade #{self.activity_count} executada com sucesso"
                )

            # Programar próxima atividade com variação se habilitado
            if self.is_running and self.advanced_tab.random_intervals.isChecked():
                interval = self.interval_spin.value() * 1000
                variation = interval * 0.2
                next_interval = random.randint(
                    int(interval - variation), int(interval + variation)
                )
                self.activity_timer.setInterval(next_interval)

        except Exception as e:
            error_msg = f"Erro: {str(e)}"
            self.status_label.setText(error_msg)
            self.log_tab.add_log(f"ERRO: {error_msg}")
            logging.error(error_msg)

            # Mesmo com erro, tentar manter o sistema ativo
            try:
                ctypes.windll.kernel32.SetThreadExecutionState(
                    ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED
                )
                pyautogui.moveRel(1, 0)
                pyautogui.moveRel(-1, 0)
            except:
                pass

    def update_status(self):
        """Atualiza indicadores de status na interface"""
        if self.advanced_tab.check_rdp.isChecked():
            rdp_status = is_rdp_active()
            self.rdp_status.setText(
                f"Status RDP: {'Ativo' if rdp_status else 'Inativo'}"
            )
        else:
            self.rdp_status.setText("Status RDP: Verificação desativada")

        if self.advanced_tab.check_teams.isChecked():
            teams_status = is_teams_running()
            self.teams_status.setText(
                f"Status Teams: {'Em execução' if teams_status else 'Não detectado'}"
            )
        else:
            self.teams_status.setText("Status Teams: Verificação desativada")

    def check_schedule(self):
        """Verifica se está dentro do horário agendado"""
        now = QTime.currentTime()
        start_time = self.start_time_edit.time()
        end_time = self.end_time_edit.time()

        within_schedule = start_time <= now <= end_time

        if within_schedule and not self.is_running:
            self.log_tab.add_log(
                f"Iniciando serviço automaticamente conforme agendamento ({start_time.toString()} - {end_time.toString()})"
            )
            self.toggle_service()
        elif not within_schedule and self.is_running:
            self.log_tab.add_log(
                f"Parando serviço automaticamente conforme agendamento ({start_time.toString()} - {end_time.toString()})"
            )
            self.toggle_service()

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
        QCoreApplication.setAttribute(
            Qt.ApplicationAttribute.AA_EnableHighDpiScaling, True
        )
    if hasattr(Qt.ApplicationAttribute, "AA_UseHighDpiPixmaps"):
        QCoreApplication.setAttribute(
            Qt.ApplicationAttribute.AA_UseHighDpiPixmaps, True
        )
    os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"
    app.setStyle("Windows")

    # Tenta criar o diretório de logs se não existir
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
