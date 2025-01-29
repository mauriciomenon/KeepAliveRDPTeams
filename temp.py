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
import json
import ctypes
import psutil
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
        # Chamar a superclasse ANTES de qualquer verificação
        super().__init__(
            *args, **kwargs
        )  # Chama corretamente a inicialização do QMainWindow

        if hasattr(self, "_is_initialized"):
            return  # Evita reinicialização indesejada

        # Logging
        self.logger = logging.getLogger(self.__class__.__name__)

        # Configuração
        self.config = SystemConfig.load_config()

        # Managers
        self.teams_manager = TeamsElectronManager(self.config)
        self.rdp_manager = RDPManager(self.logger)

        # Estado da aplicação
        self.is_running = False
        self.current_teams_status = TeamsStatus.AVAILABLE

        # Timers
        self.activity_timer = QTimer(self)
        self.activity_timer.timeout.connect(self.check_system_activity)

        self.main_timer = QTimer(self)
        self.main_timer.timeout.connect(self.perform_keep_alive)

        self.schedule_timer = QTimer(self)
        self.schedule_timer.timeout.connect(self.check_schedule)

        # Configuração da UI
        self.setup_ui()
        self.setup_tray_icon()

        # Listeners de eventos
        self.setup_event_listeners()

        # Marca como inicializado
        self._is_initialized = True

        self.logger.info("Keep Alive Manager inicializado corretamente.")

    def setup_ui(self):
        """
        Configure main application user interface
        """
        # Window properties
        self.setWindowTitle("Keep Alive Manager")
        self.setFixedSize(500, 600)

        # Central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # Status Frame
        status_frame = StyleFrame()
        status_layout = QVBoxLayout(status_frame)

        self.status_label = QLabel(
            "Keep Alive Manager\n"
            "Configure settings and start the service"
        )
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        status_layout.addWidget(self.status_label)
        main_layout.addWidget(status_frame)

        # Teams Status Frame
        teams_frame = StyleFrame()
        teams_layout = QVBoxLayout(teams_frame)

        teams_label = QLabel("Teams Status:")
        teams_layout.addWidget(teams_label)

        # Teams status buttons
        self.teams_status_buttons = {}
        teams_status_grid = QGridLayout()

        status_grid_config = [
            (TeamsStatus.AVAILABLE, (0, 0)),
            (TeamsStatus.BUSY, (0, 1)),
            (TeamsStatus.DO_NOT_DISTURB, (0, 2)),
            (TeamsStatus.AWAY, (1, 0)),
            (TeamsStatus.BE_RIGHT_BACK, (1, 1)),
            (TeamsStatus.OFFLINE, (1, 2))
        ]

        for status, (row, col) in status_grid_config:
            btn = QPushButton(status.display_name)
            btn.setCheckable(True)
            btn.setStyleSheet(f"background-color: {status.color};")
            btn.clicked.connect(lambda _, s=status: self.set_teams_status(s))

            self.teams_status_buttons[status] = btn
            teams_status_grid.addWidget(btn, row, col)

        teams_layout.addLayout(teams_status_grid)
        main_layout.addWidget(teams_frame)

        # Configuration Frame
        config_frame = StyleFrame()
        config_layout = QVBoxLayout(config_frame)

        # Interval configuration
        interval_layout = QHBoxLayout()
        interval_label = QLabel("Keep Alive Interval (seconds):")
        self.interval_spin = QSpinBox()
        self.interval_spin.setRange(30, 300)
        self.interval_spin.setValue(self.config.get('interval', 120))

        interval_layout.addWidget(interval_label)
        interval_layout.addWidget(self.interval_spin)
        config_layout.addLayout(interval_layout)

        # Time range configuration
        time_layout = QHBoxLayout()
        start_label = QLabel("Active Period:")

        self.start_time_edit = QTimeEdit()
        self.start_time_edit.setTime(QTime.fromString(
            self.config.get('start_time', '08:45'), 
            'HH:mm'
        ))

        end_label = QLabel("to")

        self.end_time_edit = QTimeEdit()
        self.end_time_edit.setTime(QTime.fromString(
            self.config.get('end_time', '17:15'), 
            'HH:mm'
        ))

        time_layout.addWidget(start_label)
        time_layout.addWidget(self.start_time_edit)
        time_layout.addWidget(end_label)
        time_layout.addWidget(self.end_time_edit)

        config_layout.addLayout(time_layout)
        main_layout.addWidget(config_frame)

        # Options Frame
        options_frame = StyleFrame()
        options_layout = QVBoxLayout(options_frame)

        self.minimize_to_tray_cb = QCheckBox("Minimize to System Tray")
        self.minimize_to_tray_cb.setChecked(
            self.config.get('minimize_to_tray', True)
        )
        options_layout.addWidget(self.minimize_to_tray_cb)
        main_layout.addWidget(options_frame)

        # Control Buttons
        button_layout = QHBoxLayout()

        self.toggle_button = QPushButton("Start")
        self.toggle_button.clicked.connect(self.toggle_service)

        self.minimize_button = QPushButton("Minimize")
        self.minimize_button.clicked.connect(self.hide)

        button_layout.addWidget(self.toggle_button)
        button_layout.addWidget(self.minimize_button)

        main_layout.addLayout(button_layout)

    def setup_tray_icon(self):
        """
        Configure system tray icon and menu
        """
        # Tray icon
        self.tray_icon = QSystemTrayIcon(
            self.style().standardIcon(QStyle.StandardPixmap.SP_ComputerIcon),
            self
        )

        # Tray menu
        tray_menu = QMenu()

        # Actions
        show_action = QAction("Show", self)
        show_action.triggered.connect(self.show)

        self.toggle_tray_action = QAction("Start", self)
        self.toggle_tray_action.triggered.connect(self.toggle_service)

        quit_action = QAction("Exit", self)
        quit_action.triggered.connect(self.quit_application)

        # Add actions to menu
        tray_menu.addAction(show_action)
        tray_menu.addAction(self.toggle_tray_action)
        tray_menu.addSeparator()
        tray_menu.addAction(quit_action)

        # Set menu and show tray icon
        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.show()

    def setup_event_listeners(self):
        """
        Set up event listeners for Teams manager
        """
        # Connection events
        self.teams_manager.add_event_listener(
            'on_connect', 
            self.on_teams_connected
        )
        self.teams_manager.add_event_listener(
            'on_disconnect', 
            self.on_teams_disconnected
        )
        self.teams_manager.add_event_listener(
            'on_status_change', 
            self.on_teams_status_changed
        )

    def on_teams_connected(self):
        """
        Handle Teams connection event
        """
        self.logger.info("Teams connected")
        self.status_label.setText("Teams connected successfully")

    def on_teams_disconnected(self):
        """
        Handle Teams disconnection event
        """
        self.logger.warning("Teams disconnected")
        self.status_label.setText("Teams disconnected. Attempting to reconnect...")

    def on_teams_status_changed(self, status):
        """
        Handle Teams status change event
        
        Args:
            status (TeamsStatus): New Teams status
        """
        self.logger.info(f"Teams status changed to {status.display_name}")

        # Update UI buttons
        for btn_status, btn in self.teams_status_buttons.items():
            btn.setChecked(btn_status == status)

    def check_system_activity(self):
        """
        Monitor system activity to prevent unnecessary keep-alive actions
        """
        try:
            current_pos = win32api.GetCursorPos()
            if not hasattr(self, 'last_cursor_pos'):
                self.last_cursor_pos = current_pos
                return

            # Check for cursor movement
            if current_pos != self.last_cursor_pos:
                self.last_user_activity = time.time()
                self.last_cursor_pos = current_pos
        except Exception as e:
            self.logger.error(f"System activity check error: {e}")

    def perform_keep_alive(self):
        """
        Perform keep-alive actions for RDP and Teams
        """
        try:
            # Check if we should perform keep-alive
            if not self.should_perform_keep_alive():
                return

            # Keep RDP session alive
            rdp_success = self.rdp_manager.keep_session_alive()

            # Update Teams status
            teams_success = self.teams_manager.set_status(self.current_teams_status)

            # Update status label
            status_text = (
                f"Keep Alive Status:\n"
                f"RDP: {'Maintained' if rdp_success else 'Failed'}\n"
                f"Teams: {'Status Updated' if teams_success else 'Update Failed'}"
            )
            self.status_label.setText(status_text)

            self.logger.info("Keep alive cycle completed")
        except Exception as e:
            self.logger.error(f"Keep alive cycle error: {e}")

    def should_perform_keep_alive(self) -> bool:
        """
        Determine if keep-alive should be performed
        
        Returns:
            bool: Whether to perform keep-alive
        """
        # Check schedule
        current_time = QTime.currentTime()
        start_time = self.start_time_edit.time()
        end_time = self.end_time_edit.time()

        # Check if current time is within scheduled period
        if not (start_time <= current_time <= end_time):
            if self.is_running:
                self.toggle_service()
            return False

        # Check user activity
        if hasattr(self, 'last_user_activity'):
            inactivity_duration = time.time() - self.last_user_activity
            if inactivity_duration < self.interval_spin.value():
                return False

        return self.is_running

    def set_teams_status(self, status: TeamsStatus):
        """
        Set Teams status
        
        Args:
            status (TeamsStatus): Desired Teams status
        """
        try:
            if self.teams_manager.set_status(status):
                self.current_teams_status = status
                self.logger.info(f"Teams status set to {status.display_name}")
            else:
                self.logger.warning(f"Failed to set Teams status to {status.display_name}")
        except Exception as e:
            self.logger.error(f"Error setting Teams status: {e}")

    def toggle_service(self):
        """
        Toggle keep-alive service on/off
        """
        self.is_running = not self.is_running

        if self.is_running:
            # Start service
            self.toggle_button.setText("Stop")
            self.toggle_tray_action.setText("Stop")

            # Start timers
            self.activity_timer.start(1000)  # Check activity every second
            self.main_timer.start(self.interval_spin.value() * 1000)
            self.schedule_timer.start(60000)  # Check schedule every minute

            # Start Teams connection
            self.teams_manager.connect()

            self.status_label.setText("Keep Alive service started")
            self.logger.info("Keep Alive service started")
        else:
            # Stop service
            self.toggle_button.setText("Start")
            self.toggle_tray_action.setText("Start")

            # Stop timers
            self.activity_timer.stop()
            self.main_timer.stop()
            self.schedule_timer.stop()

            # Close Teams connection
            self.teams_manager.close()

            self.status_label.setText("Keep Alive service stopped")
            self.logger.info("Keep Alive service stopped")

    def closeEvent(self, event):
        """
        Handle window close event
        
        Args:
            event (QCloseEvent): Close event
        """
        if self.minimize_to_tray_cb.isChecked():
            event.ignore()
            self.hide()
            self.tray_icon.showMessage(
                "Keep Alive Manager",
                "Application is running in background",
                QSystemTrayIcon.MessageIcon.Information,
                2000
            )
        else:
            self.quit_application()

    def quit_application(self):
        """
        Gracefully quit the application
        """
        try:
            # Stop service if running
            if self.is_running:
                self.toggle_service()

            # Close Teams manager
            if hasattr(self, 'teams_manager'):
                self.teams_manager.close()

            # Remove tray icon
            if hasattr(self, 'tray_icon'):
                self.tray_icon.hide()

            # Save current configuration
            config_update = {
                'interval': self.interval_spin.value(),
                'start_time': self.start_time_edit.time().toString('HH:mm'),
                'end_time': self.end_time_edit.time().toString('HH:mm'),
                'minimize_to_tray': self.minimize_to_tray_cb.isChecked(),
                'default_teams_status': self.current_teams_status.ipc_status
            }

            try:
                with open('config.json', 'w', encoding='utf-8') as f:
                    json.dump(config_update, f, indent=4)
            except Exception as e:
                self.logger.error(f"Failed to save configuration: {e}")

            # Exit application
            QApplication.quit()

        except Exception as e:
            self.logger.error(f"Error during application quit: {e}")
            QApplication.quit()


def main():
    """
    Main application entry point with comprehensive error handling
    
    Returns:
        int: Application exit status
    """
    try:
        # Configure logging
        logger = logging.getLogger('KeepAliveManager')
        logger.info("Initializing Keep Alive Manager")
        
        # Ensure single instance
        import win32event
        import win32api
        from winerror import ERROR_ALREADY_EXISTS
        
        # Create mutex to prevent multiple instances
        mutex = win32event.CreateMutex(None, False, "KeepAliveManagerMutex")
        if win32api.GetLastError() == ERROR_ALREADY_EXISTS:
            logger.warning("Another instance of the application is already running")
            return 1
        
        # Create Qt application
        app = QApplication(sys.argv)
        
        # Set application attributes for high DPI support
        if hasattr(Qt.ApplicationAttribute, "AA_EnableHighDpiScaling"):
            QCoreApplication.setAttribute(Qt.ApplicationAttribute.AA_EnableHighDpiScaling, True)
        if hasattr(Qt.ApplicationAttribute, "AA_UseHighDpiPixmaps"):
            QCoreApplication.setAttribute(Qt.ApplicationAttribute.AA_UseHighDpiPixmaps, True)
        os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"
        
        # Set Windows-specific style
        app.setStyle("Windows")
        
        # Create and show main window
        window = KeepAliveApp()
        window.show()
        
        logger.info("Application initialized successfully")
        return app.exec()
    
    except Exception as e:
        # Log critical errors
        logger = logging.getLogger('KeepAliveManager')
        logger.critical(f"Fatal application error: {e}", exc_info=True)
        return 1

if __name__ == "__main__":
    # Run the application and exit with its status
    sys.exit(main())
