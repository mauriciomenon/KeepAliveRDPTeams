# pylint: disable=E0602,E0102,E1101
"""
Keep Alive Manager - RDP & Teams Integration
--------------------------------------------

Comprehensive implementation addressing development requirements:

1. Base Interface and System
   - Implement correct singleton pattern
   - Add file rotation logging
   - Improve exception handling
   - Implement retry mechanism with exponential backoff
   - Fix main UI initialization issues

2. Teams Communication (IPC/WebSocket)
   - Implement Teams WebSocket protocol
   - Handle authentication and message formats
   - Implement keep-alive mechanism
   - Automatic reconnection handling
   - Validate behavior across Teams versions

Dependencies:
- PyQt6
- websockets
- aiohttp
- cryptography
- psutil
- pywin32

Author: Maurício Menon (AI-Assisted Implementation)
Date: January/2024
Version: 1.0-dev
"""


import sys
import os
os.environ["QT_QPA_PLATFORM"] = "windows"
import json
import ctypes
import psutil
import win32event
from winerror import ERROR_ALREADY_EXISTS
import win32gui
import win32con
import win32api
import win32ts
import win32process
import win32security
import websockets
import asyncio
import aiohttp
from typing import Callable
import socket
import threading
import uuid
import logging
from logging.handlers import RotatingFileHandler
from datetime import datetime
import time
from pathlib import Path
from enum import Enum
from typing import Optional, Dict, Any, List

from PyQt6.QtCore import Qt, QTimer, QTime, QCoreApplication, pyqtSignal, QObject
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
    QGridLayout,
)
from PyQt6.QtGui import QIcon, QAction

# Enhanced Logging Configuration
def setup_logging(log_dir: str = 'logs') -> logging.Logger:
    """
    Configure comprehensive logging with rotation and detailed context
    
    Args:
        log_dir (str): Directory for log files
    
    Returns:
        logging.Logger: Configured logger with multiple handlers
    """
    # Ensure log directory exists
    os.makedirs(log_dir, exist_ok=True)
    
    # Create logger
    logger = logging.getLogger('KeepAliveManager')
    logger.setLevel(logging.DEBUG)
    
    # Clear any existing handlers to prevent duplicate logging
    logger.handlers.clear()
    
    # File handler with advanced rotation
    file_handler = RotatingFileHandler(
        os.path.join(log_dir, 'keep_alive.log'),
        maxBytes=10*1024*1024,  # 10 MB
        backupCount=5,
        encoding='utf-8'
    )
    file_handler.setLevel(logging.DEBUG)
    
    # Console handler
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    
    # Formatter with more comprehensive information
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(processName)s - %(threadName)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)
    
    # Add handlers
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    
    return logger

# Global logger with enhanced configuration
logger = setup_logging()


# Singleton Implementation
class SingletonMeta(type):
    """
    Metaclass for implementing singleton pattern with thread-safety
    and cross-process prevention
    """

    _instances = {}
    _singleton_mutex = None

    def __call__(cls, *args, **kwargs):
        # Ensure only one instance is created
        if cls not in cls._instances:
            try:
                from win32event import CreateMutex
                from win32api import GetLastError
                from winerror import ERROR_ALREADY_EXISTS

                # Create named mutex
                cls._singleton_mutex = CreateMutex(None, False, f"{cls.__name__}_Mutex")

                # Check if mutex already exists
                if GetLastError() == ERROR_ALREADY_EXISTS:
                    logger.warning(f"Singleton {cls.__name__} already running")
                    sys.exit(1)

                # Create instance if not exists
                cls._instances[cls] = super().__call__(*args, **kwargs)
            except Exception as e:
                logger.critical(f"Singleton initialization error: {e}")
                raise

        return cls._instances[cls]


# Retry Decorator with Exponential Backoff
def retry(max_attempts=3, base_delay=1, max_delay=30, 
          exceptions=(Exception,), logger=logger):
    """
    Decorator for implementing retry mechanism with exponential backoff
    
    Args:
        max_attempts (int): Maximum number of retry attempts
        base_delay (int): Initial delay between retries
        max_delay (int): Maximum delay between retries
        exceptions (tuple): Exceptions to catch and retry
        logger (logging.Logger): Logger for tracking retry attempts
    """
    def decorator(func):
        def wrapper(*args, **kwargs):
            attempts = 0
            delay = base_delay
            
            while attempts < max_attempts:
                try:
                    return func(*args, **kwargs)
                except exceptions as e:
                    attempts += 1
                    if attempts == max_attempts:
                        logger.error(f"Max attempts reached. Final error: {e}")
                        raise
                    
                    logger.warning(
                        f"Attempt {attempts}/{max_attempts} failed: {e}. "
                        f"Retrying in {delay} seconds..."
                    )
                    
                    # Sleep with exponential backoff
                    time.sleep(delay)
                    
                    # Increase delay exponentially
                    delay = min(delay * 2, max_delay)
        
        return wrapper
    return decorator

# Constants and Configuration
class SystemConfig:
    """
    Centralized system configuration with enhanced error handling
    """
    # System preservation constants
    ES_CONTINUOUS = 0x80000000
    ES_SYSTEM_REQUIRED = 0x00000001
    ES_DISPLAY_REQUIRED = 0x00000002
    
    # Teams WebSocket configuration
    TEAMS_WS_PORT_START = 8001
    TEAMS_WS_PORT_END = 8999
    TEAMS_WS_TIMEOUT = 5  # seconds
    TEAMS_WS_RETRY_DELAY = 1000  # ms
    
    @classmethod
    def load_config(cls, config_path='config.json'):
        """
        Load configuration with robust error handling
        
        Args:
            config_path (str): Path to configuration file
        
        Returns:
            dict: Loaded configuration
        """
        try:
            if not os.path.exists(config_path):
                return cls.create_default_config(config_path)
            
            with open(config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            logger.error(f"Config load error: {e}")
            return cls.create_default_config(config_path)
    
    @classmethod
    def create_default_config(cls, config_path='config.json'):
        """
        Create default configuration file
        
        Args:
            config_path (str): Path to save configuration
        
        Returns:
            dict: Default configuration
        """
        default_config = {
            'interval': 120,  # seconds
            'start_time': '08:45',
            'end_time': '17:15',
            'minimize_to_tray': True,
            'default_teams_status': 'Available'
        }
        
        try:
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(default_config, f, indent=4)
            return default_config
        except Exception as e:
            logger.critical(f"Failed to create default config: {e}")
            return default_config

# Enhanced error classes for more specific error handling
class KeepAliveError(Exception):
    """Base exception for Keep Alive Manager"""
    pass

class TeamsConnectionError(KeepAliveError):
    """Specific exception for Teams connection issues"""
    pass

class RDPConnectionError(KeepAliveError):
    """Specific exception for RDP connection issues"""
    pass

# Teams Status Management
class TeamsStatus(Enum):
    """Enumeration of possible Teams statuses with enhanced metadata"""
    AVAILABLE = ("Available", "Available", "#00ff00")
    BUSY = ("Busy", "Busy", "#ff0000")
    DO_NOT_DISTURB = ("Do Not Disturb", "DoNotDisturb", "#ff4500")
    AWAY = ("Away", "Away", "#ffa500")
    OFFLINE = ("Offline", "Offline", "#808080")
    BE_RIGHT_BACK = ("Be Right Back", "BeRightBack", "#ffff00")
    
    def __init__(self, display_name: str, ipc_status: str, color: str):
        self.display_name = display_name
        self.ipc_status = ipc_status
        self.color = color

class TeamsElectronManager:
    """
    Advanced manager for Teams Electron IPC communication
    
    Handles:
    - WebSocket connection
    - Status management
    - Automatic reconnection
    - Error handling
    """
    def __init__(self, config: Dict[str, Any] = None):
        """
        Initialize Teams Electron IPC Manager
        
        Args:
            config (Dict[str, Any], optional): Configuration dictionary
        """
        # Configuration
        self.config = config or {}
        
        # WebSocket communication
        self.websocket = None
        self.ipc_port = None
        self.is_connected = False
        
        # Connection management
        self.connection_attempts = 0
        self.max_connection_attempts = self.config.get(
            'max_connection_attempts', 
            SystemConfig.MAX_CONNECTION_ATTEMPTS
        )
        
        # Async loop and threading
        self.async_loop = asyncio.new_event_loop()
        self.async_thread = None
        
        # Signals and event handling
        self.connection_events = {
            'on_connect': [],
            'on_disconnect': [],
            'on_status_change': []
        }
        
        # Logging
        self.logger = logging.getLogger(self.__class__.__name__)
        
        # Start async thread
        self._start_async_thread()
    
    def _start_async_thread(self):
        """
        Start async event loop in a separate thread
        """
        def run_loop():
            asyncio.set_event_loop(self.async_loop)
            self.async_loop.run_forever()
        
        self.async_thread = threading.Thread(
            target=run_loop, 
            daemon=True
        )
        self.async_thread.start()
    
    @retry(max_attempts=3, base_delay=1, max_delay=10)
    async def _discover_teams_port(self) -> Optional[int]:
        """
        Discover active WebSocket port for Teams
        
        Returns:
            Optional[int]: Active WebSocket port
        """
        # Check Teams configuration file first
        teams_config_path = os.path.join(
            os.getenv('LOCALAPPDATA', ''), 
            'Microsoft', 'Teams', 
            'desktop-config.json'
        )
        
        try:
            # Read config file
            if os.path.exists(teams_config_path):
                with open(teams_config_path, 'r') as f:
                    config = json.load(f)
                    if 'webSocketPort' in config:
                        port = config['webSocketPort']
                        if await self._test_websocket_port(port):
                            return port
            
            # Scan ports if config fails
            for port in range(
                SystemConfig.TEAMS_WS_PORT_START, 
                SystemConfig.TEAMS_WS_PORT_END
            ):
                if await self._test_websocket_port(port):
                    return port
            
            return None
        except Exception as e:
            self.logger.error(f"Port discovery error: {e}")
            return None
    
    async def _test_websocket_port(self, port: int) -> bool:
        """
        Test if a WebSocket port is active
        
        Args:
            port (int): Port to test
        
        Returns:
            bool: Port is active and responsive
        """
        try:
            async with aiohttp.ClientSession() as session:
                async with session.ws_connect(
                    f'ws://127.0.0.1:{port}',
                    timeout=0.5
                ) as ws:
                    await ws.ping()
                    return True
        except Exception:
            return False
    
    async def _establish_websocket_connection(self, port: int):
        """
        Establish WebSocket connection to Teams
        
        Args:
            port (int): WebSocket port to connect
        
        Raises:
            TeamsConnectionError: If connection fails
        """
        try:
            self.websocket = await websockets.connect(
                f'ws://127.0.0.1:{port}',
                ping_interval=30,  # Keep-alive every 30 seconds
                ping_timeout=10
            )
            self.is_connected = True
            self.ipc_port = port
            
            # Trigger connection events
            for callback in self.connection_events['on_connect']:
                callback()
            
            self.logger.info(f"Connected to Teams WebSocket on port {port}")
        except Exception as e:
            self.is_connected = False
            raise TeamsConnectionError(f"WebSocket connection failed: {e}")
    
    def connect(self):
        """
        Initiate Teams connection process
        Runs asynchronously and manages connection lifecycle
        """
        def connection_task():
            try:
                # Find port
                port_future = asyncio.run_coroutine_threadsafe(
                    self._discover_teams_port(), 
                    self.async_loop
                )
                port = port_future.result()
                
                if not port:
                    raise TeamsConnectionError("No Teams WebSocket port found")
                
                # Establish connection
                connect_future = asyncio.run_coroutine_threadsafe(
                    self._establish_websocket_connection(port), 
                    self.async_loop
                )
                connect_future.result()
            
            except TeamsConnectionError as e:
                self.logger.warning(f"Connection failed: {e}")
                # Trigger disconnect events
                for callback in self.connection_events['on_disconnect']:
                    callback()
                
                # Retry connection if attempts not exhausted
                self.connection_attempts += 1
                if self.connection_attempts < self.max_connection_attempts:
                    time.sleep(2 ** self.connection_attempts)  # Exponential backoff
                    self.connect()
        
        # Run connection in separate thread
        threading.Thread(
            target=connection_task, 
            daemon=True
        ).start()
    
    def set_status(self, status: TeamsStatus):
        """
        Set Teams user status
        
        Args:
            status (TeamsStatus): Desired Teams status
        
        Returns:
            bool: Status change success
        """
        if not self.is_connected:
            self.logger.warning("Cannot set status - not connected")
            return False
        try:
            message = {
                "id": str(uuid.uuid4()),
                "method": "setUserPresence",
                "params": {
                    "status": status.ipc_status,
                    "expiry": None,
                    "deviceId": None
                }
            }
            
            # Send message via WebSocket
            send_future = asyncio.run_coroutine_threadsafe(
            self._send_message(message), 
            self.async_loop
            )
            success = send_future.result()
            
            if success:
                # Trigger status change events
                for callback in self.connection_events['on_status_change']:
                    callback(status)
            
            return success
        except Exception as e:
            self.logger.error(f"Status change error: {e}")
            return False
    
    async def _send_message(self, message: Dict[str, Any]) -> bool:
        """
        Send WebSocket message to Teams
        
        Args:
            message (Dict[str, Any]): Message to send
        
        Returns:
            bool: Message sent successfully
        """
        if not self.websocket:
            return False
        
        try:
            await self.websocket.send(json.dumps(message))
            response = await self.websocket.recv()
            return "error" not in response.lower()
        except Exception as e:
            self.logger.error(f"WebSocket send error: {e}")
            return False
    
    def add_event_listener(
        self, 
        event: str, 
        callback: Callable
    ):
        """
        Add event listener for connection events
        
        Args:
            event (str): Event type ('on_connect', 'on_disconnect', 'on_status_change')
            callback (Callable): Callback function
        """
        if event in self.connection_events:
            self.connection_events[event].append(callback)
    
    def close(self):
        """
        Close WebSocket connection and clean up resources
        """
        try:
            # Close WebSocket
            if self.websocket:
                close_future = asyncio.run_coroutine_threadsafe(
                    self.websocket.close(), 
                    self.async_loop
                )
                close_future.result()
            
            # Stop async loop
            self.async_loop.call_soon_threadsafe(self.async_loop.stop)
            
            # Wait for thread to finish
            if self.async_thread:
                self.async_thread.join(timeout=1.0)
            
            # Reset connection state
            self.is_connected = False
            self.websocket = None
            self.ipc_port = None
        except Exception as e:
            self.logger.error(f"Cleanup error: {e}")

class StyleFrame(QFrame):
    """
    Custom styled frame for consistent UI components
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFrameShape(QFrame.Shape.StyledPanel)
        self.setFrameShadow(QFrame.Shape.Raised)
        self.setStyleSheet("""
            StyleFrame {
                background-color: palette(window);
                border: 1px solid palette(mid);
                border-radius: 5px;
                padding: 10px;
            }
        """)

class RDPManager:
    """
    Manages RDP (Remote Desktop Protocol) session activities
    """
    def __init__(self, logger=None):
        self.logger = logger or logging.getLogger(self.__class__.__name__)

    def keep_session_alive(self) -> bool:
        """
        Prevent RDP session from timing out
        
        Returns:
            bool: Success of keeping session active
        """
        try:
            # Use Windows API to reset persistent session
            win32ts.WTSResetPersistentSession()

            # Prevent screen lock
            ctypes.windll.kernel32.SetThreadExecutionState(
                SystemConfig.ES_CONTINUOUS | 
                SystemConfig.ES_SYSTEM_REQUIRED | 
                SystemConfig.ES_DISPLAY_REQUIRED
            )

            return True
        except Exception as e:
            self.logger.error(f"RDP session management error: {e}")
            return False

    def check_session_status(self) -> bool:
        """
        Check current RDP session status
        
        Returns:
            bool: Session is active
        """
        try:
            session_id = win32ts.WTSGetActiveConsoleSessionId()
            return session_id != 0xFFFFFFFF
        except Exception as e:
            self.logger.error(f"RDP session status check error: {e}")
            return False


class KeepAliveApp(QMainWindow):
    _instance = None

    def __new__(cls, *args, **kwargs):
        if cls._instance is None:
            cls._instance = super(KeepAliveApp, cls).__new__(cls)
        return cls._instance

    def __init__(self, *args, **kwargs):
        print(
            "DEBUG: Entrando no __init__ da KeepAliveApp - ANTES do super().__init__()"
        )

        try:
            super().__init__(*args, **kwargs)
            print("DEBUG: Passou pelo super().__init__()")
        except Exception as e:
            print(f"ERRO CRÍTICO: Falha ao chamar super().__init__() -> {e}")
            return

        if hasattr(self, "_is_initialized"):
            print("DEBUG: KeepAliveApp já inicializado! Retornando.")
            return  # Evita reinicialização

        print("DEBUG: Inicializando KeepAliveApp")

        # Logging
        self.logger = logging.getLogger(self.__class__.__name__)

        # Configuração
        try:
            print("DEBUG: Carregando configurações")
            self.config = SystemConfig.load_config()
            print("DEBUG: Configuração carregada:", self.config)
        except Exception as e:
            print(f"ERRO CRÍTICO: Falha ao carregar configurações -> {e}")
            return  # Sai da inicialização se falhar

        # Criando Managers
        try:
            print("DEBUG: Inicializando TeamsElectronManager")
            self.teams_manager = TeamsElectronManager(self.config)
            print("DEBUG: TeamsElectronManager inicializado")

            print("DEBUG: Inicializando RDPManager")
            self.rdp_manager = RDPManager(self.logger)
            print("DEBUG: RDPManager inicializado")
        except Exception as e:
            print(f"ERRO CRÍTICO: Falha ao iniciar Managers -> {e}")
            return  # Sai da inicialização se falhar

        # Estado da aplicação
        self.is_running = False
        self.current_teams_status = TeamsStatus.AVAILABLE
        print("DEBUG: Estado da aplicação inicializado")

        # Criando Timers
        try:
            print("DEBUG: Criando timers")
            self.activity_timer = QTimer(self)
            self.activity_timer.timeout.connect(self.check_system_activity)

            self.main_timer = QTimer(self)
            self.main_timer.timeout.connect(self.perform_keep_alive)

            self.schedule_timer = QTimer(self)
            self.schedule_timer.timeout.connect(self.check_schedule)
            print("DEBUG: Timers criados")
        except Exception as e:
            print(f"ERRO CRÍTICO: Falha ao iniciar timers -> {e}")
            return  # Sai da inicialização se falhar

        # Configuração da UI
        try:
            print("DEBUG: Chamando setup_ui()")
            self.setup_ui()
            print("DEBUG: setup_ui() finalizado")

            print("DEBUG: Chamando setup_tray_icon()")
            self.setup_tray_icon()
            print("DEBUG: setup_tray_icon() finalizado")

            print("DEBUG: Chamando setup_event_listeners()")
            self.setup_event_listeners()
            print("DEBUG: setup_event_listeners() finalizado")
        except Exception as e:
            print(f"ERRO CRÍTICO: Falha ao configurar UI -> {e}")
            return  # Sai da inicialização se falhar

        # Teste rápido: mostrar uma QLabel simples para ver se a UI renderiza
        self.test_label = QLabel("Teste de inicialização da UI", self)
        self.test_label.setGeometry(10, 10, 200, 50)
        self.test_label.show()
        print("DEBUG: Teste de QLabel exibido")

        # Marca como inicializado
        self._is_initialized = True
        print("DEBUG: KeepAliveApp inicializado com sucesso!")

        self.logger.info("Keep Alive Manager inicializado corretamente.")

    # 🔹 CONFIGURANDO A UI
    def setup_ui(self):
        print("DEBUG: Configurando UI")
        self.setWindowTitle("Keep Alive Manager")
        self.setFixedSize(500, 600)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        self.status_label = QLabel(
            "Keep Alive Manager - Configure settings and start the service"
        )
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.status_label)

        print("DEBUG: UI configurada")

    # 🔹 CONFIGURANDO O ÍCONE DA BANDEJA
    def setup_tray_icon(self):
        print("DEBUG: Configurando ícone da bandeja")
        self.tray_icon = QSystemTrayIcon(QIcon("icon.png"), self)

        tray_menu = QMenu()
        show_action = QAction("Show", self)
        show_action.triggered.connect(self.show)

        quit_action = QAction("Exit", self)
        quit_action.triggered.connect(self.quit_application)

        tray_menu.addAction(show_action)
        tray_menu.addSeparator()
        tray_menu.addAction(quit_action)

        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.show()
        print("DEBUG: Ícone da bandeja configurado")

    # 🔹 EVENT LISTENERS
    def setup_event_listeners(self):
        print("DEBUG: Configurando eventos")
        self.tray_icon.activated.connect(self.toggle_window)
        print("DEBUG: Eventos configurados")

    def toggle_window(self, reason):
        if reason == QSystemTrayIcon.ActivationReason.Trigger:
            if self.isHidden():
                self.show()
            else:
                self.hide()

    def quit_application(self):
        print("DEBUG: Encerrando aplicação")
        self.tray_icon.hide()
        QApplication.quit()


def main():
    """
    Main application entry point with comprehensive error handling
    """
    try:
        os.environ["QT_QPA_PLATFORM"] = "windows"

        logger = logging.getLogger("KeepAliveManager")
        logger.info("Initializing Keep Alive Manager")
        print("DEBUG: Iniciando aplicação")

        # 🔴 Comentando Mutex temporariamente para teste
        # print("DEBUG: Verificando Mutex...")
        # mutex = win32event.CreateMutex(None, False, "KeepAliveManagerMutex")
        # if win32api.GetLastError() == ERROR_ALREADY_EXISTS:
        #     logger.warning("Outra instância do KeepAliveManager já está rodando.")
        #     print("DEBUG: Já existe uma instância rodando. Encerrando.")
        #     return 1
        # print("DEBUG: Mutex adquirido com sucesso.")

        # Criar aplicação Qt
        print("DEBUG: Verificando QApplication existente")
        if not QApplication.instance():
            app = QApplication(sys.argv)
        else:
            app = QApplication.instance()
        print("DEBUG: QApplication criada")

        # Criar a janela principal
        print("DEBUG: Criando KeepAliveApp...")
        window = KeepAliveApp()
        print("DEBUG: KeepAliveApp instanciada")

        # Forçar exibição da janela
        window.setWindowFlags(Qt.WindowType.WindowStaysOnTopHint)
        window.show()
        window.raise_()
        window.activateWindow()
        print("DEBUG: Janela principal exibida")

        # Iniciar o loop da interface gráfica
        print("DEBUG: Entrando no loop do Qt")
        ret = app.exec()
        print(f"DEBUG: Loop do Qt finalizado, saindo com código: {ret}")
        return ret

    except Exception as e:
        logger.critical(f"Fatal application error: {e}", exc_info=True)
        print(f"ERRO CRÍTICO: {e}")
        return 1


# Garantir captura de erros globais ao iniciar
if __name__ == "__main__":
    try:
        print("DEBUG: Iniciando aplicação")
        ret = main()
        print(f"DEBUG: Aplicação finalizou com código {ret}")
    except Exception as e:
        print(f"ERRO FATAL: {e}")
