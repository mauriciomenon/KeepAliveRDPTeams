# pylint: disable=E0602,E0102,E1101
"""
Keep Alive Manager - RDP & Teams (Electron IPC Implementation)
------------------------------------------------------------------

Esta versão está em desenvolvimento e necessita das seguintes implementações:

1. Interface e Sistema Base
   - Debugar inicialização do QApplication
     (possível problema no QTimer thread-safety)
   - Implementar singleton pattern corretamente usando win32event
   - Adicionar logging com rotação de arquivos (logging.handlers.RotatingFileHandler)
   - Melhorar tratamento de exceções com contexto específico para cada operação
   - Implementar mecanismo de retry com backoff exponencial para reconexões
   - Corrigir problema de inicialização da UI principal

2. Comunicação Teams (IPC/WebSocket)
   - Implementar protocolo WebSocket do Teams:
     * Handshake inicial com autenticação
     * Handling de mensagens com formato específico do Teams
     * Keep-alive via ping/pong a cada 30s
     * Reconexão automática em caso de disconnect
   - Investigar necessidade de interceptação de credenciais SSO
   - Validar comportamento com diferentes versões do Teams (New/Classic)
   - Implementar fallback para Graph API quando necessário

3. Gerenciamento de Estado
   - Implementar FSM (Finite State Machine) para controle de estados
   - Adicionar fila de mensagens para retry em caso de falha
   - Persistir configurações em arquivo JSON com encryption
   - Implementar mecanismo de cache para status
   - Adicionar validação de estado da sessão RDP

4. Performance e Recursos
   - Otimizar uso de threads/processos
   - Implementar rate limiting para chamadas IPC
   - Adicionar monitoramento de uso de memória
   - Implementar cleanup de recursos não utilizados
   - Melhorar gestão de timers Qt

5. Segurança
   - Adicionar validação de integridade do processo Teams
   - Implementar proteção contra DLL injection
   - Adicionar checksums nas mensagens IPC
   - Implementar sanitização de inputs
   - Validar permissões de acesso

6. Pontos Críticos
   - Resolver race conditions na thread do WebSocket
   - Implementar mecanismo de deadlock prevention
   - Melhorar handling de exceções assíncronas
   - Resolver memory leaks na UI
   - Implementar graceful shutdown

7. Debugging/Logs
   - Adicionar logs detalhados de IPC/WebSocket
   - Implementar modo debug com packet inspection
   - Adicionar métricas de performance
   - Implementar sistema de diagnóstico
   - Criar dumps de memória em caso de crash

8. Próximos Passos (Prioridade)
   - Adicionar logging detalhado da inicialização
   - Verificar problemas de thread-safety nos QTimers
   - Implementar reconexão automática do WebSocket
   - Adicionar tratamento para Teams não encontrado
   - Melhorar feedback visual de erros

Dependências necessárias:
- PyQt6
- websockets
- aiohttp
- cryptography
- psutil
- pywin32

Autor: Maurício Menon
Data: Janeiro/2024
Versão: 1.0-dev
"""

import sys
import os
import json
import ctypes
import threading
import uuid
import logging
from datetime import datetime
import time
from enum import Enum
from typing import Dict, Any
import asyncio

import psutil
import win32ts
import websockets
import aiohttp

from PyQt6.QtCore import (
    Qt,
    QTimer,
    QTime,
    QCoreApplication,
    pyqtSignal,
    QObject,
)
from PyQt6.QtGui import QIcon
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
    QGridLayout,
    QTimeEdit,
    QAction,
)

# Configuração de logging básico
logging.basicConfig(
    level=logging.DEBUG, format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)

# Configurações do sistema
ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001
ES_DISPLAY_REQUIRED = 0x00000002

# Configurações do Teams
TEAMS_WS_PORT_START = 8001
TEAMS_WS_PORT_END = 8999
TEAMS_WS_TIMEOUT = 5  # segundos
TEAMS_WS_RETRY_DELAY = 1000  # ms

# Templates de mensagens para o Teams
TEAMS_STATUS_MSG = {
    "id": "{uuid}",
    "method": "setUserPresence",
    "params": {"status": "{status}", "expiry": None, "deviceId": None},
}


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


class TeamsError(Exception):
    """Exceção customizada para erros do Teams"""

    pass


class TeamsSignals(QObject):
    """Sinais para comunicação assíncrona do Teams"""

    connection_status = pyqtSignal(bool, str)
    status_changed = pyqtSignal(bool, str)


class TeamsElectronManager:
    """Gerenciador de comunicação com o processo Electron do Teams"""

    def __init__(self):
        self.teams_process = None
        self.ipc_connected = False
        self.websocket = None
        self.connection_timer = QTimer()
        self.connection_timer.timeout.connect(self._try_connect)
        self.connection_attempts = 0
        self.max_attempts = 3
        self.ipc_port = None
        self.signals = TeamsSignals()
        self.ws_lock = threading.Lock()

        # Configurar loop assíncrono em thread separada
        self.loop = asyncio.new_event_loop()
        self.thread = threading.Thread(target=self._run_async_loop, daemon=True)
        self.thread.start()

    def _run_async_loop(self):
        """Executa o loop assíncrono em thread separada"""
        asyncio.set_event_loop(self.loop)
        self.loop.run_forever()

    async def _find_teams_ws_port(self):
        """Procura a porta WebSocket do Teams"""
        teams_path = os.path.join(os.getenv("LOCALAPPDATA"), "Microsoft", "Teams")
        config_file = os.path.join(teams_path, "desktop-config.json")

        try:
            if os.path.exists(config_file):
                with open(config_file, "r") as f:
                    config = json.load(f)
                    if "webSocketPort" in config:
                        port = config["webSocketPort"]
                        logger.info(f"Porta WebSocket encontrada no config: {port}")
                        if await self._test_ws_connection(port):
                            return port
        except Exception as e:
            logger.error(f"Erro ao ler arquivo de configuração: {str(e)}")

        logger.info("Procurando porta WebSocket ativa...")
        for port in range(TEAMS_WS_PORT_START, TEAMS_WS_PORT_END):
            if await self._test_ws_connection(port):
                logger.info(f"Porta WebSocket encontrada: {port}")
                return port

        return None

    async def _test_ws_connection(self, port: int) -> bool:
        """Testa conexão WebSocket em uma porta específica"""
        try:
            async with aiohttp.ClientSession() as session:
                async with session.ws_connect(
                    f"ws://127.0.0.1:{port}", timeout=0.5
                ) as ws:
                    await ws.ping()
                    return True
        except Exception:
            return False

    async def _connect_websocket(self) -> None:
        """Estabelece conexão WebSocket com o Teams"""
        if not self.ipc_port:
            raise TeamsError("Porta WebSocket não configurada")

        try:
            self.websocket = await websockets.connect(
                f"ws://127.0.0.1:{self.ipc_port}",
                ping_interval=None,
                ping_timeout=TEAMS_WS_TIMEOUT,
            )
            logger.info("Conexão WebSocket estabelecida")
        except Exception as e:
            raise TeamsError(f"Erro ao conectar WebSocket: {str(e)}")

    def start_connection(self):
        """Inicia tentativas de conexão com o Teams"""
        self.connection_timer.start(TEAMS_WS_RETRY_DELAY)

    def _try_connect(self):
        """Tenta estabelecer conexão com o Teams"""
        try:
            self.connection_attempts += 1
            logger.info(
                f"Tentativa de conexão {self.connection_attempts}/{self.max_attempts}"
            )

            self.teams_process = self._find_teams_process()
            if not self.teams_process:
                logger.warning("Processo do Teams não encontrado")
                self.signals.connection_status.emit(False, "Teams não encontrado")
                if self.connection_attempts >= self.max_attempts:
                    self.connection_timer.stop()
                return False

            logger.info("Processo do Teams encontrado")

            # Procurar porta WebSocket
            future = asyncio.run_coroutine_threadsafe(
                self._find_teams_ws_port(), self.loop
            )
            self.ipc_port = future.result()

            if not self.ipc_port:
                logger.warning("Porta WebSocket não encontrada")
                self.signals.connection_status.emit(False, "Porta não encontrada")
                if self.connection_attempts >= self.max_attempts:
                    self.connection_timer.stop()
                return False

            # Estabelecer conexão WebSocket
            future = asyncio.run_coroutine_threadsafe(
                self._connect_websocket(), self.loop
            )
            future.result()

            logger.info(f"Porta WebSocket encontrada: {self.ipc_port}")
            self.ipc_connected = True
            self.signals.connection_status.emit(True, "Conectado")
            self.connection_timer.stop()
            return True

        except Exception as e:
            logger.error(f"Erro ao conectar com Teams: {str(e)}")
            self.signals.connection_status.emit(False, str(e))
            if self.connection_attempts >= self.max_attempts:
                self.connection_timer.stop()
            return False

    def _find_teams_process(self):
        """Localiza o processo principal do Teams"""
        teams_path = os.path.join(os.getenv("LOCALAPPDATA"), "Microsoft", "Teams")

        for proc in psutil.process_iter(["pid", "name", "cmdline", "exe"]):
            try:
                if "ms-teams.exe" in proc.info["name"].lower():
                    proc_path = proc.info.get("exe", "")
                    # Verificar se é o processo principal do Teams
                    if teams_path.lower() in proc_path.lower():
                        logger.info(f"Teams encontrado em: {proc_path}")
                        return proc
            except (psutil.NoSuchProcess, psutil.AccessDenied, KeyError):
                continue
        return None

    async def _send_ws_message(self, message: Dict[str, Any]) -> bool:
        """Envia mensagem WebSocket para o Teams"""
        if not self.websocket:
            return False

        try:
            message_str = json.dumps(message)
            await self.websocket.send(message_str)
            response = await self.websocket.recv()
            return "error" not in response
        except Exception as e:
            logger.error(f"Erro ao enviar mensagem WebSocket: {str(e)}")
            return False

    def set_status(self, status: str) -> bool:
        """Altera o status via WebSocket"""
        if not self.ipc_connected or not self.websocket:
            self.signals.status_changed.emit(False, "Não conectado")
            return False

        try:
            # Preparar mensagem de status
            message = TEAMS_STATUS_MSG.copy()
            message["id"] = str(uuid.uuid4())
            message["params"]["status"] = status

            # Enviar mensagem
            future = asyncio.run_coroutine_threadsafe(
                self._send_ws_message(message), self.loop
            )
            success = future.result()

            if success:
                self.signals.status_changed.emit(True, "Status alterado")
                return True
            else:
                self.signals.status_changed.emit(False, "Erro ao alterar status")
                return False

        except Exception as e:
            logger.error(f"Erro ao definir status: {str(e)}")
            self.signals.status_changed.emit(False, str(e))
            return False

    def close(self):
        """Fecha a conexão com o Teams"""
        logger.info("Fechando conexões do Teams")
        self.connection_timer.stop()
        if self.websocket:
            future = asyncio.run_coroutine_threadsafe(self.websocket.close(), self.loop)
            future.result()
        self.ipc_connected = False
        self.loop.call_soon_threadsafe(self.loop.stop)
        self.thread.join(timeout=1.0)


class StyleFrame(QFrame):
    """Frame estilizado para a interface"""

    def __init__(self):
        super().__init__()
        self.setFrameShape(QFrame.Shape.StyledPanel)
        self.setFrameShadow(QFrame.Shape.Raised)
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
        self.teams_manager.signals.connection_status.connect(self.on_connection_status)
        self.teams_manager.signals.status_changed.connect(self.on_status_changed)

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

    def on_connection_status(self, connected, message):
        """Callback para mudanças no status de conexão"""
        if connected:
            logger.info(f"Teams conectado: {message}")
            self.status_label.setText(f"Teams conectado\n{message}")
        else:
            logger.warning(f"Erro na conexão com Teams: {message}")
            self.status_label.setText(f"Erro na conexão com Teams\n{message}")

    def on_status_changed(self, success, message):
        """Callback para mudanças no status do Teams"""
        if success:
            logger.info(f"Status atualizado: {message}")
            self.status_label.setText(f"Status atualizado\n{message}")
        else:
            logger.warning(f"Erro ao atualizar status: {message}")
            self.status_label.setText(f"Erro ao atualizar status\n{message}")

    def setup_ui(self) -> None:
        """Configura a interface do usuário"""
        logger.info("Configurando interface do usuário")

        # Widget central
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        layout.setSpacing(15)

        # Status Frame
        status_frame = StyleFrame()
        status_layout = QVBoxLayout()
        status_frame.setLayout(status_layout)

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
        config_layout = QVBoxLayout()
        config_frame.setLayout(config_layout)

        # Intervalo
        interval_layout = QHBoxLayout()
        interval_label = QLabel("Intervalo entre ações (segundos):")
        self.interval_spin = QSpinBox()
        self.interval_spin.setRange(30, 300)
        self.interval_spin.setValue(self.default_interval)

        interval_layout.addWidget(interval_label)
        interval_layout.addWidget(self.interval_spin)
        config_layout.addLayout(interval_layout)

        # Horários
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
        layout.addWidget(config_frame)

        # Teams Status Frame
        teams_frame = StyleFrame()
        teams_layout = QVBoxLayout()
        teams_frame.setLayout(teams_layout)

        status_label = QLabel("Status do Teams:")
        teams_layout.addWidget(status_label)

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
        options_layout = QVBoxLayout()
        options_frame.setLayout(options_layout)

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

        self.minimize_button = QPushButton("Minimizar")
        self.minimize_button.clicked.connect(self.hide)
        self.minimize_button.setFixedWidth(100)

        button_layout.addWidget(self.toggle_button)
        button_layout.addWidget(self.minimize_button)
        button_layout.addStretch()

        layout.addLayout(button_layout)
        logger.info("Interface configurada com sucesso")

    def setup_tray(self, icon: QIcon) -> None:
        """
        Configura o ícone na bandeja do sistema
        Args:
            icon (QIcon): Ícone a ser usado na bandeja
        """
        try:
            logger.info("Configurando ícone na bandeja")
            self.tray_icon = QSystemTrayIcon(icon, self)
            self.tray_icon.setToolTip("Keep Alive Manager")

            # Criar menu de contexto
            tray_menu = QMenu()

            # Ação Mostrar
            show_action = QAction("Mostrar", self)
            show_action.triggered.connect(self.show)

            # Ação Iniciar/Parar
            self.toggle_tray_action = QAction("Iniciar", self)
            self.toggle_tray_action.triggered.connect(self.toggle_service)

            # Ação Sair
            quit_action = QAction("Sair", self)
            quit_action.triggered.connect(self.quit_application)

            # Adicionar ações ao menu
            tray_menu.addAction(show_action)
            tray_menu.addAction(self.toggle_tray_action)
            tray_menu.addSeparator()
            tray_menu.addAction(quit_action)

            # Configurar menu
            self.tray_icon.setContextMenu(tray_menu)
            self.tray_icon.show()

            # Conectar sinal de duplo clique
            self.tray_icon.activated.connect(self._handle_tray_activation)

            logger.info("Ícone da bandeja configurado com sucesso")

        except Exception as e:
            logger.error(f"Erro ao configurar ícone na bandeja: {str(e)}")
            raise

    def _handle_tray_activation(self, reason: int) -> None:
        """
        Trata eventos de ativação do ícone na bandeja
        Args:
            reason (int): Razão da ativação
        """
        try:
            if reason == QSystemTrayIcon.ActivationReason.DoubleClick:
                if self.isHidden():
                    self.show()
                    self.activateWindow()
                else:
                    self.hide()
        except Exception as e:
            logger.error(f"Erro ao tratar ativação da bandeja: {str(e)}")

    def update_tray_status(self, is_running: bool) -> None:
        """
        Atualiza o status no menu da bandeja
        Args:
            is_running (bool): Se o serviço está rodando
        """
        try:
            if hasattr(self, "toggle_tray_action"):
                self.toggle_tray_action.setText("Parar" if is_running else "Iniciar")
        except Exception as e:
            logger.error(f"Erro ao atualizar status na bandeja: {str(e)}")

    def show_tray_message(
        self,
        title: str,
        message: str,
        icon_type: int = QSystemTrayIcon.MessageIcon.Information,
        duration: int = 2000,
    ) -> None:
        """
        Mostra mensagem na bandeja do sistema
        Args:
            title: Título da mensagem
            message: Conteúdo da mensagem
            icon_type: Tipo do ícone da mensagem
            duration: Duração em milissegundos
        """
        try:
            if hasattr(self, "tray_icon") and self.tray_icon.isVisible():
                self.tray_icon.showMessage(title, message, icon_type, duration)
        except Exception as e:
            logger.error(f"Erro ao mostrar mensagem na bandeja: {str(e)}")

    def closeEvent(self, event) -> None:
        """
        Trata o evento de fechamento da janela
        Args:
            event: Evento de fechamento
        """
        try:
            if self.minimize_to_tray_cb.isChecked():
                logger.info("Minimizando para a bandeja")
                event.ignore()
                self.hide()
                self.show_tray_message(
                    "Keep Alive Manager", "O programa continua rodando em segundo plano"
                )
            else:
                logger.info("Fechando aplicação")
                self.quit_application()
        except Exception as e:
            logger.error(f"Erro ao tratar fechamento da janela: {str(e)}")
            event.accept()

    def quit_application(self) -> None:
        """Encerra a aplicação adequadamente"""
        try:
            logger.info("Encerrando aplicação")

            # Parar serviço se estiver rodando
            if self.is_running:
                self.toggle_service()

            # Fechar gerenciador do Teams
            if hasattr(self, "teams_manager"):
                self.teams_manager.close()

            # Remover ícone da bandeja
            if hasattr(self, "tray_icon"):
                self.tray_icon.hide()
                self.tray_icon = None

            # Encerrar aplicação
            QApplication.quit()

        except Exception as e:
            logger.error(f"Erro ao encerrar aplicação: {str(e)}")
            QApplication.quit()


# Removed leftover code


def main():
    """Função principal da aplicação"""
    try:
        logger.info("Iniciando aplicação")

        # Prevenir múltiplas instâncias
        from win32event import CreateMutex
        from win32api import GetLastError
        from winerror import ERROR_ALREADY_EXISTS

        _ = CreateMutex(
            None, 1, "KeepAliveManager_Mutex"
        )  # Suprime warning de variável não usada
        if GetLastError() == ERROR_ALREADY_EXISTS:
            logger.warning("Aplicação já está em execução")
            sys.exit(1)

        # Criar aplicação Qt primeiro
        app = QApplication(sys.argv)

        # Configurações de DPI (sem necessidade de hasattr, PyQt6 sempre define)
        QCoreApplication.setAttribute(
            Qt.ApplicationAttribute.AA_EnableHighDpiScaling, True
        )
        QCoreApplication.setAttribute(
            Qt.ApplicationAttribute.AA_UseHighDpiPixmaps, True
        )
        os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"

        app.setStyle("Windows")
        window = KeepAliveApp()
        window.show()

        logger.info("Aplicação iniciada com sucesso")
        return app.exec()

    except Exception as e:
        logger.critical(f"Erro fatal na aplicação: {str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    status = main()
    sys.exit(status)

    try:
        self.current_teams_status.ipc_status
        logger.info("Status atualizado com sucesso")

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

        logger.debug("Ciclo completado")
        self.status_label.setText(status_text)

    except Exception as e:
        error_msg = f"Erro ao executar atividade: {str(e)}"
        logger.error(error_msg)
        self.status_label.setText(f"{error_msg}\n" "Tentando métodos alternativos...")

    def keep_rdp_alive(self):
        """Manter sessão RDP ativa"""
        try:
            logger.debug("Mantendo sessão RDP ativa")
            win32ts.WTSResetPersistentSession()
            return True
        except Exception as e:
            logger.error(f"Erro ao manter RDP: {str(e)}")
            return False

    def check_rdp_connection(self):
        """Verificar estado da conexão RDP"""
        try:
            logger.debug("Verificando conexão RDP")
            session_id = win32ts.WTSGetActiveConsoleSessionId()
            return session_id != 0xFFFFFFFF
        except Exception as e:
            logger.error(f"Erro ao verificar RDP: {str(e)}")
            return False

    def prevent_screen_lock(self):
        """Prevenir bloqueio de tela usando API do Windows"""
        try:
            logger.debug("Prevenindo bloqueio de tela")
            ctypes.windll.kernel32.SetThreadExecutionState(
                ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED
            )
            return True
        except Exception as e:
            logger.error(f"Erro ao prevenir bloqueio: {str(e)}")
            return False

    def toggle_service(self):
        """Inicia ou para o serviço de keep-alive"""
        if not self.is_running:
            logger.info("Iniciando serviço")
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
            logger.info("Parando serviço")
            self.timer.stop()
            self.is_running = False
            self.toggle_button.setText("Iniciar")
            self.status_label.setText(
                "Serviço parado\n"
                "Suas conexões não estão sendo mantidas ativas\n"
                "Clique em 'Iniciar' para retomar"
            )
            ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)

    def check_schedule(self) -> None:
        """Verifica se deve estar rodando baseado no horário configurado"""
        current_time = QTime.currentTime()
        start_time = self.start_time_edit.time()
        end_time = self.end_time_edit.time()

        logger.debug(f"Verificando horário: {current_time.toString()}")

        if start_time <= current_time <= end_time:
            if not self.is_running and self.toggle_button.text() == "Iniciar":
                logger.info("Dentro do horário programado - iniciando")
                self.toggle_service()
        else:
            if self.is_running:
                logger.info("Fora do horário programado - parando")
                self.toggle_service()

    def set_teams_status(self, status: TeamsStatus) -> bool:
        """
        Define o status do Teams
        Args:
            status: Novo status a ser definido
        Returns:
            bool: True se sucesso, False caso contrário
        """
        try:
            logger.info(f"Alterando status para: {status.display_name}")
            if self.teams_manager.set_status(status.ipc_status):
                self.current_teams_status = status
                # Atualizar botões da UI
                for s, btn in self.teams_buttons.items():
                    btn.setChecked(s == status)
                return True
            return False
        except Exception as e:
            logger.error(f"Erro ao definir status: {str(e)}")
            return False
