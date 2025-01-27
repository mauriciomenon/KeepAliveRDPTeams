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
from datetime import datetime, time as dt_time
import time
import keyboard  # Adicionar no topo do arquivo junto com outros imports
import pyautogui
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                           QHBoxLayout, QLabel, QSpinBox, QTimeEdit, QCheckBox, 
                           QPushButton, QSystemTrayIcon, QMenu, QStyle, QFrame)
from PyQt6.QtCore import QTimer, Qt, QTime
from PyQt6.QtGui import QIcon, QAction
import win32gui
import win32con
import win32api
import ctypes
from ctypes import wintypes
import random

# Configurações do sistema
ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001
ES_DISPLAY_REQUIRED = 0x00000002

# Configuração do PyAutoGUI
pyautogui.FAILSAFE = False
pyautogui.PAUSE = 0.1

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
            lambda reason: self.show() if reason == QSystemTrayIcon.ActivationReason.DoubleClick else None
        )

    def check_schedule(self):
        current_time = QTime.currentTime()
        start_time = self.start_time_edit.time()
        end_time = self.end_time_edit.time()
        
        if start_time <= current_time <= end_time:
            if not self.is_running and self.toggle_button.text() == "Iniciar":
                self.toggle_service()
        else:
            if self.is_running:
                self.toggle_service()

    def perform_activity(self):
        try:
            # Rotacionar entre diferentes métodos de keep-alive
            methods = [
                self.move_mouse_safely,
                self.simulate_shift_key,
                self.prevent_screen_lock
            ]
            
            # Usar o método atual e mover para o próximo
            current_method = methods[self.method_index]
            success = current_method()
            
            # Se o método atual falhar, tentar o próximo
            if not success:
                self.method_index = (self.method_index + 1) % len(methods)
                backup_method = methods[self.method_index]
                success = backup_method()
            
            self.activity_count += 1
            current_time = datetime.now().strftime("%H:%M:%S")
            
            if success:
                self.status_label.setText(
                    f"Mantendo conexões ativas\n"
                    f"Última atividade: {current_time} (#{self.activity_count})\n"
                    f"Próxima ação em {self.interval_spin.value()} segundos"
                )
            else:
                self.status_label.setText(
                    f"Aviso: Dificuldade em executar ações\n"
                    f"Tentando métodos alternativos\n"
                    f"Última tentativa: {current_time}"
                )
            
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
            
            # Definir um quadrado de 20x20 pixels centrado na posição atual
            max_movement = 10  # 10 pixels para cada lado = quadrado de 20x20
            
            # Calcular movimentos seguros dentro dos limites da tela
            x_move = min(max_movement, screen_width - current_pos[0] - 1)
            x_move = max(x_move, -current_pos[0])
            y_move = min(max_movement, screen_height - current_pos[1] - 1)
            y_move = max(y_move, -current_pos[1])
            
            # Movimento suave em pequeno quadrado
            moves = [
                (x_move, 0),
                (0, y_move),
                (-x_move, 0),
                (0, -y_move)
            ]
            
            for dx, dy in moves:
                debug_log(f"Movendo mouse: ({dx}, {dy})")
                pyautogui.moveRel(dx, dy, duration=0.2)

            return True
        except Exception as e:
            debug_log(f"Erro ao mover mouse: {str(e)}")
            return False

    def simulate_shift_key(self):
        """Simular pressionamento da tecla Shift para Teams"""
        try:
            pyautogui.press('shift')
            return True
        except Exception:
            return False

    def prevent_screen_lock(self):
        """Prevenir bloqueio de tela usando API do Windows"""
        try:
            debug_log("Prevenindo bloqueio de tela")
            ctypes.windll.kernel32.SetThreadExecutionState(
                ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED
            )
            return True
        except Exception:
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
    
    app = QApplication(sys.argv)
    app.setStyle("Windows")
    window = KeepAliveApp()
    window.show()
    sys.exit(app.exec())
