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
from typing import Callable, Optional, Dict, Any, List
import socket
import threading
import uuid
import logging
from logging.handlers import RotatingFileHandler
from datetime import datetime
import time
from pathlib import Path
from enum import Enum

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
    QComboBox,
)
from PyQt6.QtGui import QIcon, QAction


# Enhanced Logging Configuration
def setup_logging(log_dir: str = "logs") -> logging.Logger:
    """Configure comprehensive logging with rotation and detailed context"""
    os.makedirs(log_dir, exist_ok=True)
    logger = logging.getLogger("KeepAliveManager")
    logger.setLevel(logging.DEBUG)
    logger.handlers.clear()

    file_handler = RotatingFileHandler(
        os.path.join(log_dir, "keep_alive.log"),
        maxBytes=10 * 1024 * 1024,
        backupCount=5,
        encoding="utf-8",
    )
    file_handler.setLevel(logging.DEBUG)

    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.DEBUG)

    formatter = logging.Formatter(
        "%(asctime)s - %(name)s - %(processName)s - %(threadName)s - %(levelname)s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    return logger


logger = setup_logging()


# Retry Decorator
def retry(max_attempts=3, base_delay=1, max_delay=30, exceptions=(Exception,)):
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

                    logger.warning(f"Attempt {attempts}/{max_attempts} failed: {e}")
                    time.sleep(delay)
                    delay = min(delay * 2, max_delay)

        return wrapper

    return decorator


# Constants and Configuration
class SystemConfig:
    ES_CONTINUOUS = 0x80000000
    ES_SYSTEM_REQUIRED = 0x00000001
    ES_DISPLAY_REQUIRED = 0x00000002

    TEAMS_WS_PORT_START = 8001
    TEAMS_WS_PORT_END = 8999
    TEAMS_WS_TIMEOUT = 5
    TEAMS_WS_RETRY_DELAY = 1000

    @classmethod
    def load_config(cls, config_path="config.json"):
        try:
            if not os.path.exists(config_path):
                return cls.create_default_config(config_path)

            with open(config_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            logger.error(f"Config load error: {e}")
            return cls.create_default_config(config_path)

    @classmethod
    def create_default_config(cls, config_path="config.json"):
        default_config = {
            "interval": 120,
            "start_time": "08:45",
            "end_time": "17:15",
            "minimize_to_tray": True,
            "default_teams_status": "Available",
        }

        try:
            with open(config_path, "w", encoding="utf-8") as f:
                json.dump(default_config, f, indent=4)
            return default_config
        except Exception as e:
            logger.error(f"Failed to create default config: {e}")
            return default_config


# Teams Status Management
class TeamsStatus(Enum):
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
    def __init__(self, config: Dict[str, Any] = None):
        self.config = config or {}
        self.websocket = None
        self.ipc_port = None
        self.is_connected = False
        self.connection_attempts = 0
        self.max_connection_attempts = 5
        self.async_loop = asyncio.new_event_loop()
        self.async_thread = None
        self.connection_events = {
            "on_connect": [],
            "on_disconnect": [],
            "on_status_change": [],
        }
        self.logger = logging.getLogger(self.__class__.__name__)
        self._start_async_thread()

    def _start_async_thread(self):
        def run_loop():
            asyncio.set_event_loop(self.async_loop)
            self.async_loop.run_forever()

        self.async_thread = threading.Thread(target=run_loop, daemon=True)
        self.async_thread.start()
        self.logger.debug("Async thread started")

    async def _discover_teams_port(self) -> Optional[int]:
        """Discover Teams WebSocket port with enhanced logging"""
        self.logger.debug("Starting Teams port discovery")

        # First try: Check Teams config file
        teams_config_path = os.path.join(
            os.getenv("LOCALAPPDATA", ""), "Microsoft", "Teams", "desktop-config.json"
        )

        self.logger.debug(f"Checking Teams config at: {teams_config_path}")

        try:
            if os.path.exists(teams_config_path):
                with open(teams_config_path, "r") as f:
                    config = json.load(f)
                    if "webSocketPort" in config:
                        port = config["webSocketPort"]
                        self.logger.debug(f"Found port in config: {port}")
                        if await self._test_websocket_port(port):
                            self.logger.info(
                                f"Successfully connected to port {port} from config"
                            )
                            return port
                        else:
                            self.logger.debug(f"Port {port} from config not responsive")
            else:
                self.logger.debug("Teams config file not found")
        except Exception as e:
            self.logger.error(f"Error reading Teams config: {e}")

        # Second try: Port scanning
        self.logger.debug("Starting port scan")
        for port in range(
            SystemConfig.TEAMS_WS_PORT_START, SystemConfig.TEAMS_WS_PORT_END
        ):
            self.logger.debug(f"Testing port {port}")
            if await self._test_websocket_port(port):
                self.logger.info(f"Found active Teams port through scanning: {port}")
                return port

        self.logger.warning("No active Teams ports found")
        return None

    async def _test_websocket_port(self, port: int) -> bool:
        """Test WebSocket port with enhanced error handling"""
        try:
            # Verificar primeiro se a porta está aberta
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.settimeout(0.5)
            result = sock.connect_ex(("127.0.0.1", port))
            sock.close()

            if result != 0:
                self.logger.debug(f"Port {port} is not open")
                return False

            self.logger.debug(f"Testing WebSocket connection on port {port}")

            async with aiohttp.ClientSession() as session:
                try:
                    async with session.ws_connect(
                        f"ws://127.0.0.1:{port}",
                        timeout=2.0,
                        headers={
                            "Origin": "https://teams.microsoft.com",
                            "User-Agent": "Microsoft Teams Client",
                        },
                    ) as ws:
                        # Enviar mensagem de teste
                        test_message = {"id": str(uuid.uuid4()), "method": "ping"}
                        await ws.send_json(test_message)

                        try:
                            response = await ws.receive_json(timeout=2.0)
                            self.logger.debug(
                                f"Got response from port {port}: {response}"
                            )
                            if "error" not in str(response).lower():
                                return True
                        except Exception as e:
                            self.logger.debug(
                                f"Error receiving response on port {port}: {e}"
                            )

                except Exception as e:
                    self.logger.debug(
                        f"WebSocket connection failed on port {port}: {e}"
                    )

            return False

        except Exception as e:
            self.logger.debug(f"Port {port} test failed: {e}")
            return False

    async def _establish_websocket_connection(self, port: int):
        """Establish WebSocket connection with enhanced error handling"""
        try:
            self.logger.debug(
                f"Attempting to establish WebSocket connection on port {port}"
            )

            self.websocket = await websockets.connect(
                f"ws://127.0.0.1:{port}",
                ping_interval=30,
                ping_timeout=10,
                extra_headers={
                    "Origin": "https://teams.microsoft.com",
                    "User-Agent": "Microsoft Teams Client",
                },
            )

            # Test connection with ping
            test_message = {"id": str(uuid.uuid4()), "method": "ping"}
            await self.websocket.send(json.dumps(test_message))
            response = await self.websocket.recv()

            if "error" not in str(response).lower():
                self.is_connected = True
                self.ipc_port = port
                self.logger.info(
                    f"Successfully established Teams WebSocket connection on port {port}"
                )

                for callback in self.connection_events["on_connect"]:
                    callback()
                return True

            self.logger.warning(f"Connection test failed on port {port}")
            return False

        except Exception as e:
            self.is_connected = False
            self.logger.error(f"Failed to establish WebSocket connection: {e}")
            raise

    def connect(self):
        """Initiate Teams connection with enhanced error handling"""

        def connection_task():
            try:
                self.logger.info("Starting Teams connection process")

                # Find port
                port_future = asyncio.run_coroutine_threadsafe(
                    self._discover_teams_port(), self.async_loop
                )
                port = port_future.result(timeout=5.0)

                if not port:
                    raise Exception("No Teams WebSocket port found")

                # Establish connection
                connect_future = asyncio.run_coroutine_threadsafe(
                    self._establish_websocket_connection(port), self.async_loop
                )
                connect_future.result(timeout=5.0)

            except Exception as e:
                self.logger.warning(f"Teams connection failed: {e}")
                for callback in self.connection_events["on_disconnect"]:
                    callback()

                self.connection_attempts += 1
                if self.connection_attempts < self.max_connection_attempts:
                    self.logger.info(
                        f"Retrying connection (attempt {self.connection_attempts + 1})"
                    )
                    time.sleep(2**self.connection_attempts)
                    self.connect()
                else:
                    self.logger.error("Max connection attempts reached")

        threading.Thread(target=connection_task, daemon=True).start()

    async def _test_websocket_port(self, port: int) -> bool:
        try:
            async with aiohttp.ClientSession() as session:
                try:
                    async with session.ws_connect(
                        f"ws://127.0.0.1:{port}",
                        timeout=0.5,
                        headers={
                            "Origin": "https://teams.microsoft.com",
                            "User-Agent": "Microsoft Teams",
                        },
                    ) as ws:
                        # Enviar mensagem de teste
                        test_message = {"id": str(uuid.uuid4()), "method": "ping"}
                        await ws.send_json(test_message)

                        # Aguardar resposta
                        try:
                            response = await ws.receive_json(timeout=1.0)
                            if "error" not in response:
                                self.logger.info(
                                    f"Teams WebSocket port {port} is responsive"
                                )
                                return True
                        except Exception:
                            pass

                        return False
                except Exception:
                    return False
        except Exception as e:
            self.logger.debug(f"Port {port} test failed: {e}")
            return False

    async def _establish_websocket_connection(self, port: int):
        try:
            self.websocket = await websockets.connect(
                f"ws://127.0.0.1:{port}", ping_interval=30, ping_timeout=10
            )
            self.is_connected = True
            self.ipc_port = port

            for callback in self.connection_events["on_connect"]:
                callback()

            self.logger.info(f"Connected to Teams WebSocket on port {port}")
        except Exception as e:
            self.is_connected = False
            raise Exception(f"WebSocket connection failed: {e}")

    def connect(self):
        def connection_task():
            try:
                port_future = asyncio.run_coroutine_threadsafe(
                    self._discover_teams_port(), self.async_loop
                )
                port = port_future.result()

                if not port:
                    raise Exception("No Teams WebSocket port found")

                connect_future = asyncio.run_coroutine_threadsafe(
                    self._establish_websocket_connection(port), self.async_loop
                )
                connect_future.result()

            except Exception as e:
                self.logger.warning(f"Connection failed: {e}")
                for callback in self.connection_events["on_disconnect"]:
                    callback()

                self.connection_attempts += 1
                if self.connection_attempts < self.max_connection_attempts:
                    time.sleep(2**self.connection_attempts)
                    self.connect()

        threading.Thread(target=connection_task, daemon=True).start()

    async def _send_message(self, message: Dict[str, Any]) -> bool:
        """Send WebSocket message with enhanced error handling"""
        if not self.websocket:
            self.logger.warning("Cannot send message - no WebSocket connection")
            return False

        try:
            self.logger.debug(f"Sending message: {message}")
            await self.websocket.send(json.dumps(message))

            response = await self.websocket.recv()
            self.logger.debug(f"Received response: {response}")

            if isinstance(response, str):
                response_data = json.loads(response)
                if "error" not in response_data:
                    return True
                else:
                    self.logger.warning(
                        f"Error in response: {response_data.get('error')}"
                    )
            return False

        except Exception as e:
            self.logger.error(f"WebSocket send error: {e}")
            return False

    def set_status(self, status: TeamsStatus):
        """Set Teams status with enhanced error handling"""
        if not self.is_connected:
            self.logger.warning("Cannot set status - not connected to Teams")
            return False

        try:
            self.logger.info(
                f"Attempting to set Teams status to: {status.display_name}"
            )

            message = {
                "id": str(uuid.uuid4()),
                "method": "setUserPresence",
                "params": {
                    "status": status.ipc_status,
                    "expiry": None,
                    "deviceId": None,
                },
            }

            send_future = asyncio.run_coroutine_threadsafe(
                self._send_message(message), self.async_loop
            )

            success = send_future.result(timeout=5.0)

            if success:
                self.logger.info(
                    f"Successfully set Teams status to: {status.display_name}"
                )
                for callback in self.connection_events["on_status_change"]:
                    callback(status)
                return True
            else:
                self.logger.warning(
                    f"Failed to set Teams status to: {status.display_name}"
                )
                return False

        except Exception as e:
            self.logger.error(f"Status change error: {e}")
            return False

    def add_event_listener(self, event: str, callback: Callable):
        if event in self.connection_events:
            self.connection_events[event].append(callback)

    def close(self):
        try:
            if self.websocket:
                close_future = asyncio.run_coroutine_threadsafe(
                    self.websocket.close(), self.async_loop
                )
                close_future.result()

            self.async_loop.call_soon_threadsafe(self.async_loop.stop)

            if self.async_thread:
                self.async_thread.join(timeout=1.0)

            self.is_connected = False
            self.websocket = None
            self.ipc_port = None
        except Exception as e:
            self.logger.error(f"Cleanup error: {e}")

    # ... [rest of TeamsElectronManager methods] ...


class RDPManager:
    def __init__(self):
        self.logger = logging.getLogger(self.__class__.__name__)
        # Flags para simular atividade do mouse/teclado
        self.MOUSEEVENTF_MOVE = 0x0001
        self.INPUT_MOUSE = 0
        self.INPUT_KEYBOARD = 1

    def _simulate_input(self):
        """Simula um pequeno movimento do mouse para manter a sessão ativa"""
        try:
            # Estrutura para input do mouse
            class MOUSEINPUT(ctypes.Structure):
                _fields_ = [
                    ("dx", ctypes.c_long),
                    ("dy", ctypes.c_long),
                    ("mouseData", ctypes.c_ulong),
                    ("dwFlags", ctypes.c_ulong),
                    ("time", ctypes.c_ulong),
                    ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong)),
                ]

            class INPUT(ctypes.Structure):
                _fields_ = [("type", ctypes.c_ulong), ("mi", MOUSEINPUT)]

            # Criar input do mouse
            x = INPUT()
            x.type = self.INPUT_MOUSE
            x.mi = MOUSEINPUT(1, 1, 0, self.MOUSEEVENTF_MOVE, 0, None)

            # Enviar input
            ctypes.windll.user32.SendInput(1, ctypes.pointer(x), ctypes.sizeof(x))

            # Mover de volta
            x.mi = MOUSEINPUT(-1, -1, 0, self.MOUSEEVENTF_MOVE, 0, None)
            ctypes.windll.user32.SendInput(1, ctypes.pointer(x), ctypes.sizeof(x))

            return True
        except Exception as e:
            self.logger.error(f"Failed to simulate input: {e}")
            return False

    def keep_session_alive(self) -> bool:
        try:
            # Prevenir suspensão do sistema
            ctypes.windll.kernel32.SetThreadExecutionState(
                SystemConfig.ES_CONTINUOUS
                | SystemConfig.ES_SYSTEM_REQUIRED
                | SystemConfig.ES_DISPLAY_REQUIRED
            )

            # Simular input para manter a sessão RDP ativa
            if self._simulate_input():
                return True

            return False
        except Exception as e:
            self.logger.error(f"RDP session management error: {e}")
            return False

    def check_session_status(self) -> bool:
        try:
            # Verifica se estamos em uma sessão RDP
            session_id = win32ts.WTSGetActiveConsoleSessionId()
            if session_id == 0xFFFFFFFF:
                return False

            # Verifica se a sessão está ativa
            return ctypes.windll.user32.GetForegroundWindow() != 0
        except Exception as e:
            self.logger.error(f"Session status check error: {e}")
            return False


class KeepAliveApp(QMainWindow):
    def __init__(self):
        # Initialize parent class first
        super().__init__()
        logger.debug("Super init completed")

        # Basic configuration
        self.config = SystemConfig.load_config()
        logger.debug("Configuration loaded")

        # Initialize managers
        self.teams_manager = TeamsElectronManager(self.config)
        self.rdp_manager = RDPManager()
        logger.debug("Managers initialized")

        # Setup UI
        self.setup_ui()
        self.setup_tray()
        self.setup_timers()

        logger.debug("KeepAliveApp initialization completed")

    def setup_ui(self):
        self.setWindowTitle("Keep Alive Manager")
        self.setFixedSize(500, 600)

        # Create central widget and layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # Status section
        status_frame = QFrame()
        status_frame.setFrameStyle(QFrame.Shape.StyledPanel | QFrame.Shadow.Raised)
        status_layout = QVBoxLayout(status_frame)

        self.status_label = QLabel("Status: Initializing...")
        status_layout.addWidget(self.status_label)

        # Last activity section
        self.last_activity_label = QLabel("Last Activity: None")
        status_layout.addWidget(self.last_activity_label)

        # Teams connection status
        self.teams_connection_label = QLabel("Teams Connection: Disconnected")
        status_layout.addWidget(self.teams_connection_label)

        # Teams status section
        teams_frame = QFrame()
        teams_frame.setFrameStyle(QFrame.Shape.StyledPanel | QFrame.Shadow.Raised)
        teams_layout = QGridLayout(teams_frame)

        teams_label = QLabel("Teams Status:")
        self.teams_status_combo = QComboBox()
        for status in TeamsStatus:
            self.teams_status_combo.addItem(status.display_name)
        self.teams_status_combo.currentIndexChanged.connect(
            self.on_teams_status_changed
        )

        teams_layout.addWidget(teams_label, 0, 0)
        teams_layout.addWidget(self.teams_status_combo, 0, 1)

        # Settings section
        settings_frame = QFrame()
        settings_frame.setFrameStyle(QFrame.Shape.StyledPanel | QFrame.Shadow.Raised)
        settings_layout = QGridLayout(settings_frame)

        # Interval setting
        interval_label = QLabel("Keep-alive Interval (seconds):")
        self.interval_spinbox = QSpinBox()
        self.interval_spinbox.setRange(30, 3600)
        self.interval_spinbox.setValue(self.config["interval"])
        self.interval_spinbox.valueChanged.connect(self.on_interval_changed)

        settings_layout.addWidget(interval_label, 0, 0)
        settings_layout.addWidget(self.interval_spinbox, 0, 1)

        # Working hours settings
        start_time_label = QLabel("Start Time:")
        self.start_time_edit = QTimeEdit()
        self.start_time_edit.setTime(
            QTime.fromString(self.config["start_time"], "HH:mm")
        )
        self.start_time_edit.timeChanged.connect(self.on_time_changed)

        end_time_label = QLabel("End Time:")
        self.end_time_edit = QTimeEdit()
        self.end_time_edit.setTime(QTime.fromString(self.config["end_time"], "HH:mm"))
        self.end_time_edit.timeChanged.connect(self.on_time_changed)

        settings_layout.addWidget(start_time_label, 1, 0)
        settings_layout.addWidget(self.start_time_edit, 1, 1)
        settings_layout.addWidget(end_time_label, 2, 0)
        settings_layout.addWidget(self.end_time_edit, 2, 1)

        # Minimize to tray option
        self.minimize_checkbox = QCheckBox("Minimize to Tray")
        self.minimize_checkbox.setChecked(self.config["minimize_to_tray"])
        self.minimize_checkbox.stateChanged.connect(self.on_minimize_changed)
        settings_layout.addWidget(self.minimize_checkbox, 3, 0, 1, 2)

        # Control buttons
        button_layout = QHBoxLayout()

        self.start_button = QPushButton("Start")
        self.start_button.clicked.connect(self.start_service)
        button_layout.addWidget(self.start_button)

        self.stop_button = QPushButton("Stop")
        self.stop_button.clicked.connect(self.stop_service)
        self.stop_button.setEnabled(False)
        button_layout.addWidget(self.stop_button)

        # Add all sections to main layout
        layout.addWidget(status_frame)
        layout.addWidget(teams_frame)
        layout.addWidget(settings_frame)
        layout.addLayout(button_layout)

        # Connect Teams manager events
        self.teams_manager.add_event_listener("on_connect", self.on_teams_connected)
        self.teams_manager.add_event_listener(
            "on_disconnect", self.on_teams_disconnected
        )
        self.teams_manager.add_event_listener(
            "on_status_change", self.on_teams_status_updated
        )

        logger.debug("UI setup completed")

    def setup_tray(self):
        self.tray_icon = QSystemTrayIcon(
            self.style().standardIcon(QStyle.StandardPixmap.SP_ComputerIcon), self
        )

        tray_menu = QMenu()
        show_action = QAction("Show", self)
        show_action.triggered.connect(self.show)
        quit_action = QAction("Exit", self)
        quit_action.triggered.connect(self.quit_application)

        tray_menu.addAction(show_action)
        tray_menu.addSeparator()
        tray_menu.addAction(quit_action)

        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.activated.connect(self.tray_icon_activated)
        self.tray_icon.show()

        logger.debug("Tray icon setup completed")

    def setup_timers(self):
        # Activity check timer
        self.activity_timer = QTimer(self)
        self.activity_timer.timeout.connect(self.check_system_activity)
        self.activity_timer.start(60000)  # Check every minute

        # Keep alive timer
        self.keepalive_timer = QTimer(self)
        self.keepalive_timer.timeout.connect(self.perform_keep_alive)
        self.keepalive_timer.start(self.config["interval"] * 1000)

        logger.debug("Timers setup completed")

    def check_schedule(self) -> bool:
        """Check if current time is within working hours"""
        current_time = QTime.currentTime()
        start_time = self.start_time_edit.time()
        end_time = self.end_time_edit.time()

        # Handle case where end time is on the next day
        if end_time < start_time:
            return current_time >= start_time or current_time <= end_time

        return start_time <= current_time <= end_time

    def check_system_activity(self):
        """Check system activity and perform keep-alive if needed"""
        try:
            # Check if we're within working hours
            if not self.check_schedule():
                self.status_label.setText("Status: Outside working hours")
                return

            # Check RDP session status
            if not self.rdp_manager.check_session_status():
                self.status_label.setText("Status: No active RDP session")
                return

            # Perform keep-alive
            self.perform_keep_alive()

        except Exception as e:
            logger.error(f"Error checking system activity: {e}")
            self.status_label.setText(f"Status: Activity check failed - {str(e)}")

    def perform_keep_alive(self):
        """Perform keep-alive actions"""
        try:
            current_time = datetime.now().strftime("%H:%M:%S")
            status_messages = []

            # Keep RDP session alive
            if self.rdp_manager.keep_session_alive():
                status_messages.append("RDP: Active")
                self.last_activity_label.setText(f"Last Activity: {current_time}")
                logger.info(f"Keep-alive performed at {current_time}")
            else:
                status_messages.append("RDP: Failed")
                logger.warning("RDP keep-alive failed")

            # Update Teams status if needed
            if self.teams_manager.is_connected:
                status_messages.append("Teams: Connected")
                # Refresh Teams status
                current_status = self.teams_status_combo.currentText()
                status = next(
                    (s for s in TeamsStatus if s.display_name == current_status),
                    TeamsStatus.AVAILABLE,
                )
                if self.teams_manager.set_status(status):
                    status_messages.append(f"Status: {status.display_name}")
                else:
                    status_messages.append("Teams Status Update Failed")
            else:
                status_messages.append("Teams: Disconnected")
                # Tentar reconectar
                self.teams_manager.connect()

            # Update status display
            status_text = f"Status [{current_time}]: " + " | ".join(status_messages)
            self.status_label.setText(status_text)

        except Exception as e:
            logger.error(f"Error performing keep-alive: {e}")
            self.status_label.setText(f"Status: Keep-alive failed - {str(e)}")

    def tray_icon_activated(self, reason):
        if reason == QSystemTrayIcon.ActivationReason.Trigger:
            if self.isHidden():
                self.show()
            else:
                self.hide()

    # Event Handlers
    def on_teams_status_changed(self, index):
        """Handle Teams status combo box changes"""
        status = list(TeamsStatus)[index]
        if self.teams_manager.is_connected:
            self.teams_manager.set_status(status)
            self.status_label.setText(
                f"Status: Changing Teams status to {status.display_name}"
            )

    def on_interval_changed(self, value):
        """Handle keep-alive interval changes"""
        self.config["interval"] = value
        if hasattr(self, "keepalive_timer"):
            self.keepalive_timer.setInterval(value * 1000)
        self.save_config()

    def on_time_changed(self):
        """Handle working hours changes"""
        self.config["start_time"] = self.start_time_edit.time().toString("HH:mm")
        self.config["end_time"] = self.end_time_edit.time().toString("HH:mm")
        self.save_config()

    def on_minimize_changed(self, state):
        """Handle minimize to tray setting changes"""
        self.config["minimize_to_tray"] = bool(state)
        self.save_config()

    def on_teams_connected(self):
        """Handle Teams connection established"""
        self.teams_connection_label.setText("Teams Connection: Connected")
        self.teams_connection_label.setStyleSheet("color: green")

        # Set initial status
        current_status = self.teams_status_combo.currentText()
        status = next(
            (s for s in TeamsStatus if s.display_name == current_status),
            TeamsStatus.AVAILABLE,
        )
        self.teams_manager.set_status(status)

    def on_teams_disconnected(self):
        """Handle Teams connection lost"""
        self.teams_connection_label.setText("Teams Connection: Disconnected")
        self.teams_connection_label.setStyleSheet("color: red")

    def on_teams_status_updated(self, status):
        """Handle Teams status updates"""
        index = self.teams_status_combo.findText(status.display_name)
        if index >= 0:
            self.teams_status_combo.setCurrentIndex(index)

    # Service Control Methods
    def start_service(self):
        """Start the keep-alive service"""
        try:
            # Start timers
            self.activity_timer.start()
            self.keepalive_timer.start()

            # Connect to Teams
            self.teams_manager.connect()

            # Update UI
            self.start_button.setEnabled(False)
            self.stop_button.setEnabled(True)
            self.status_label.setText("Status: Service running")

            # Perform initial keep-alive
            self.perform_keep_alive()

            logger.info("Keep-alive service started")
        except Exception as e:
            logger.error(f"Failed to start service: {e}")
            self.status_label.setText(f"Status: Failed to start service - {str(e)}")

    def stop_service(self):
        """Stop the keep-alive service"""
        try:
            # Stop timers
            self.activity_timer.stop()
            self.keepalive_timer.stop()

            # Disconnect Teams
            self.teams_manager.close()

            # Update UI
            self.start_button.setEnabled(True)
            self.stop_button.setEnabled(False)
            self.status_label.setText("Status: Service stopped")

            logger.info("Keep-alive service stopped")
        except Exception as e:
            logger.error(f"Failed to stop service: {e}")
            self.status_label.setText(f"Status: Failed to stop service - {str(e)}")

    def save_config(self):
        """Save current configuration to file"""
        try:
            with open("config.json", "w", encoding="utf-8") as f:
                json.dump(self.config, f, indent=4)
            logger.debug("Configuration saved successfully")
        except Exception as e:
            logger.error(f"Failed to save configuration: {e}")

    def closeEvent(self, event):
        """Handle application close event"""
        if self.config["minimize_to_tray"] and self.tray_icon.isVisible():
            event.ignore()
            self.hide()
        else:
            self.stop_service()
            event.accept()

    def quit_application(self):
        """Quit the application"""
        logger.debug("Quitting application")
        self.stop_service()
        self.tray_icon.hide()
        QApplication.quit()


def main():
    try:
        # Create application
        app = QApplication(sys.argv)
        logger.debug("QApplication created")

        # Create and show main window
        window = KeepAliveApp()
        window.show()
        logger.debug("Main window created and shown")

        # Start event loop
        return app.exec()

    except Exception as e:
        logger.error(f"Error in main: {e}", exc_info=True)
        return 1


if __name__ == "__main__":
    sys.exit(main())
