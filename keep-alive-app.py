# pylint: disable=E0602,E0102,E1101


"""
Keep Alive RDP Connection 2.2.1
Maurício Menon
Foz do Iguaçu 11/06/2025
https://github.com/mauriciomenon/KeepAliveRDP
"""

import ctypes
import logging
import os
import random
import re
import socket
import subprocess
import sys
import time
from datetime import datetime, timedelta

import pyautogui
import win32api
import win32con
import win32event
import win32gui
from PyQt6.QtCore import QLoggingCategory, QSettings, Qt, QTime, QTimer
from PyQt6.QtGui import QAction, QFont
from PyQt6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPushButton,
    QSizePolicy,
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
# CONSTANTES E CONFIGURAÇÕES
# =============================================================================
# Informações do programa
APP_NAME = "Keep Alive RDP Connection"
APP_VERSION = "2.2.1"
APP_AUTHOR = "Maurício Menon"
APP_DATE = "Foz do Iguaçu 11/06/2025"
APP_URL = "https://github.com/mauriciomenon/KeepAliveRDP"

# Windows
SPI_GETSCREENSAVETIMEOUT = 0x000E  # SystemParametersInfo action code
DEFAULT_SCREENSAVER_TIMEOUT = 900  # 15 min (padrão Windows)

# Configurações padrão
DEFAULT_INTERVAL = 60  # Segundos entre cada "despertar"
DEFAULT_USER_TIMEOUT = 60  # Segundos de inatividade necessária
INACTIVE_LOG_INTERVAL = 900000  # Milissegundos para log inativo (15 min)

# Ranges dos sliders
INTERVAL_MIN = 30  # Mínimo para intervalo entre tentativas
INTERVAL_MAX = 300  # Máximo para intervalo entre tentativas
USER_TIMEOUT_MIN = 30  # Mínimo para inatividade do usuário
USER_TIMEOUT_MAX = 300  # Máximo para inatividade do usuário

# Horários padrão
DEFAULT_START_TIME_HOUR = 8  # Hora de início
DEFAULT_START_TIME_MINUTE = 0  # Minuto de início
DEFAULT_END_TIME_HOUR = 18  # Hora de término
DEFAULT_END_TIME_MINUTE = 0  # Minuto de término

MUTEX_NAME = "KeepAlive_RDP_Unique_Instance_2025"

# Strings constantes
STR_SERVICE_RUNNING = "Serviço em Execução"
STR_SERVICE_STOPPED = "Serviço Parado"
STR_SCHEDULE_RUNNING = "Execução no Intervalo {} - {}"
STR_CONTINUOUS_RUNNING = "Execução sem Interrupção"
STR_START_CONTINUOUS = "Iniciar Contínuo"
STR_START_SCHEDULED = "Iniciar Agendado"
STR_STOP_SERVICE = "Pausar"
STR_RESTART_APP = "Reiniciar Programa"
STR_MINIMIZE = "Minimizar"
STR_QUIT = "Fechar Completamente"
STR_NEXT_ACTIVITY = "Próxima Atividade: {} ({}s)"
STR_SIMULATION_SUCCESS = "Simulação Executada"
STR_SERVICE_STARTED = "Serviço Iniciado"
STR_SERVICE_STOPPED_LOG = "Serviço Parado"
STR_USER_ACTIVE = "Usuário Ativo (inatividade: {:.1f}s < {}s)"
STR_SETTINGS_SAVED = "Configurações Salvas"

# Configurações para prevenir bloqueios
ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001
ES_DISPLAY_REQUIRED = 0x00000002
ES_AWAYMODE_REQUIRED = 0x00000040

# Desabilita fail-safe do PyAutoGUI
pyautogui.FAILSAFE = False

# Configuração básica de logging
logging.basicConfig(
    level=logging.WARNING,
    format="%(asctime)s - %(levelname)s - %(message)s",
)

# Verificação de instância única
mutex = None


def is_already_running():
    """Verifica se outra instância está em execução"""
    global mutex
    try:
        # Método 1: Mutex
        mutex = win32event.CreateMutex(None, False, MUTEX_NAME)
        if win32api.GetLastError() == 183:  # ERROR_ALREADY_EXISTS
            return True

        # Método 2: Fallback - verificar por janela
        try:

            def enum_windows_callback(hwnd, windows):
                if win32gui.IsWindowVisible(hwnd):
                    window_text = win32gui.GetWindowText(hwnd)
                    if APP_NAME in window_text:
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
        # Método 3: Fallback final - arquivo de lock
        try:
            lock_file = os.path.join(os.path.expanduser("~"), ".keepalive_running")
            if os.path.exists(lock_file):
                if time.time() - os.path.getmtime(lock_file) > 3600:
                    os.remove(lock_file)
                    return False
                return True
            else:
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
    """Obtém tempo de inatividade do usuário em segundos"""
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


def get_screen_saver_timeout() -> int:
    """
    Fallback:
    1. SystemParametersInfo → valor efetivo (respeita GPO/AD).
    2. Registro HKCU\Control Panel\Desktop\ScreenSaveTimeOut.
    3. DEFAULT_SCREENSAVER_TIMEOUT (avisa em stderr; remova o print depois).
    0 indica protetor de tela desativado.
    """

    try:
        timeout = ctypes.c_int()
        if ctypes.windll.user32.SystemParametersInfoW(
            SPI_GETSCREENSAVETIMEOUT, 0, ctypes.byref(timeout), 0
        ):
            return timeout.value
    except Exception:
        pass
    try:
        key = win32api.RegOpenKeyEx(
            win32con.HKEY_CURRENT_USER,
            r"Control Panel\Desktop",
            0,
            win32con.KEY_READ,
        )
        try:
            value, _ = win32api.RegQueryValueEx(key, "ScreenSaveTimeOut")
        finally:
            win32api.RegCloseKey(key)

        if isinstance(value, (bytes, bytearray)):
            value = value.decode("ascii", errors="ignore")

        return int(value.split("\0", 1)[0].strip())
    except Exception as err:
        # debug – apague depois de validar
        print(
            f"[fallback] ScreenSaveTimeOut não lido, usando "
            f"{DEFAULT_SCREENSAVER_TIMEOUT}s: {err!r}",
            file=sys.stderr,
        )
    return DEFAULT_SCREENSAVER_TIMEOUT


def get_rdp_disconnect_time():
    """Obtém tempo de desconexão/timeout RDP do sistema de forma robusta
    1. HKLM Policies (GPO aplicado via domínio)
    2. HKCU Policies (políticas do usuário)
    3. Configurações locais do Terminal Services

    - MaxIdleTime: Timeout por inatividade
    - MaxDisconnectionTime: Timeout de desconexão de sessão
    - MaxConnectionTime: Tempo máximo de conexão contínua
    - MaxSessionTime: Tempo máximo de sessão ativa
    """
    try:
        # Chaves do registro para timeout RDP - em ordem de prioridade
        base_key = "SYSTEM\\CurrentControlSet\\Control\\Terminal Server"
        policy_key_machine = r"SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services"
        policy_key_user = r"SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services"

        # Lista de chaves para verificar (prioridade: Policies primeiro)
        keys_to_check = [
            # Políticas de Grupo - prioridade máxima
            (win32con.HKEY_LOCAL_MACHINE, policy_key_machine),
            (win32con.HKEY_CURRENT_USER, policy_key_user),
            # Configurações locais do Terminal Services
            (win32con.HKEY_LOCAL_MACHINE, f"{base_key}\\WinStations\\RDP-Tcp"),
            (win32con.HKEY_LOCAL_MACHINE, base_key),
        ]

        # Tipos de timeout para verificar (em ordem de prioridade)
        timeout_values = [
            "MaxIdleTime",  # Timeout por inatividade
            "MaxDisconnectionTime",  # Timeout de desconexão
            "MaxConnectionTime",  # Tempo máximo de conexão
            "MaxSessionTime",  # Tempo máximo de sessão
        ]

        for hive, key_path in keys_to_check:
            try:
                key = win32api.RegOpenKeyEx(
                    hive,
                    key_path,
                    0,
                    win32con.KEY_READ,
                )

                # Verificar cada tipo de timeout nesta chave
                # Coleta TODOS os valores desta chave para encontrar o menor
                found_timeouts = []
                for timeout_type in timeout_values:
                    try:
                        value, _ = win32api.RegQueryValueEx(key, timeout_type)
                        if value > 0:
                            found_timeouts.append(value)
                    except Exception:
                        # Este tipo de timeout não existe nesta chave, continua
                        continue

                # Se encontrou timeouts, retorna o MENOR (o que desconecta primeiro)
                if found_timeouts:
                    win32api.RegCloseKey(key)
                    return int(min(found_timeouts) / 1000)  # Converte ms para segundos

                win32api.RegCloseKey(key)

            except Exception:
                continue

        return 0  # Nenhum timeout encontrado

    except Exception:
        return 0


def adjust_user_timeout():
    """
    Ajusta timeout de inatividade baseado em proteção de tela e RDP
    """
    try:
        # Obtém tempos do sistema
        screen_saver_time = get_screen_saver_timeout()
        rdp_time = get_rdp_disconnect_time()

        print(
            f"[DEBUG timeout] Screen Saver: {screen_saver_time}s ({screen_saver_time//60} min)"
        )
        print(
            f"[DEBUG timeout] RDP Timeout: {rdp_time}s ({rdp_time//60} min)"
            if rdp_time > 0
            else "[DEBUG timeout] RDP Timeout: não definido"
        )

        # Lista para timeouts calculados
        calculated_timeouts = []

        # Processa Screen Saver timeout
        if screen_saver_time > 0:
            screen_timeout = max(int(screen_saver_time / 5.0), USER_TIMEOUT_MIN)
            # Limita screen timeout a no máximo 120s (2 minutos)
            screen_timeout = min(screen_timeout, 120)
            calculated_timeouts.append(screen_timeout)
            # print(f"[DEBUG timeout] Screen timeout: {screen_timeout}s (1/5 de {screen_saver_time}s)")

        # Processa RDP timeout
        if rdp_time > 0:
            rdp_timeout = max(int(rdp_time / 4.0), USER_TIMEOUT_MIN)
            # Limita RDP timeout a no máximo 90s (1.5 minutos)
            rdp_timeout = min(rdp_timeout, 90)
            calculated_timeouts.append(rdp_timeout)
            # print(f"[DEBUG timeout] RDP timeout: {rdp_timeout}s (1/4 de {rdp_time}s)")

        # Determina timeout final
        if calculated_timeouts:
            # Usa o MENOR dos timeouts
            final_timeout = min(calculated_timeouts)
            # print(f"[DEBUG timeout] Escolhido: {final_timeout}s (menor entre {calculated_timeouts})")
        else:
            # Nenhum timeout encontrado, usa padrão
            final_timeout = DEFAULT_USER_TIMEOUT
            # print(f"[DEBUG timeout] Nenhum timeout do sistema, usando padrão: {final_timeout}s")

        # Aplica limites finais mais restritivos
        # Mínimo: 30s, Máximo: 120s
        TIMEOUT_MAX_AJUSTADO = 120  # 2 minutos no máximo
        limited_timeout = max(
            min(final_timeout, TIMEOUT_MAX_AJUSTADO), USER_TIMEOUT_MIN
        )

        if limited_timeout != final_timeout:
            print(
                f"[DEBUG timeout] Limitado: {final_timeout}s → {limited_timeout}s (máx {TIMEOUT_MAX_AJUSTADO}s)"
            )

        print(
            f"[DEBUG timeout] ✅ RESULTADO FINAL: {limited_timeout}s ({limited_timeout//60}min {limited_timeout%60}s)"
        )
        print("-" * 50)

        return limited_timeout

    except Exception as e:
        print(f"[ERROR timeout] Erro no cálculo: {e}")
        return DEFAULT_USER_TIMEOUT


def format_time_intelligent(seconds):
    """Formata tempo: < 60s: "45s" | >= 60s: "10 min 15s" """
    if seconds < 60:
        return f"{seconds}s"
    minutes = seconds // 60
    remaining_seconds = seconds % 60
    if remaining_seconds == 0:
        return f"{minutes} min"
    else:
        return f"{minutes} min {remaining_seconds}s"


def get_network_info():
    """Obtém gateway, IP e nome REAL da interface principal"""
    try:
        gateway_ip = "192.168.1.1"
        my_ip = "127.0.0.1"
        interface_name = "Local"

        if sys.platform == "win32":
            # Detecta gateway principal
            result = subprocess.run(
                ["route", "print", "0.0.0.0"], capture_output=True, text=True, timeout=5
            )
            if result.returncode == 0:
                lines = result.stdout.split("\n")
                for line in lines:
                    if "0.0.0.0" in line and "On-link" not in line:
                        parts = line.split()
                        if len(parts) >= 4:
                            potential_gateway = parts[2]
                            potential_interface_ip = parts[3]
                            if re.match(
                                r"^(\d{1,3}\.){3}\d{1,3}$", potential_gateway
                            ) and re.match(
                                r"^(\d{1,3}\.){3}\d{1,3}$", potential_interface_ip
                            ):
                                gateway_ip = potential_gateway
                                my_ip = potential_interface_ip
                                break

            # NOVA LÓGICA: Captura nome REAL da interface do ipconfig
            try:
                result = subprocess.run(
                    ["ipconfig"], capture_output=True, text=True, timeout=3
                )
                if result.returncode == 0:
                    lines = result.stdout.split("\n")
                    current_adapter = ""

                    for line in lines:
                        # Detecta linha de adaptador
                        if "Adaptador" in line and ":" in line:
                            # Extrai nome real do adaptador
                            adapter_match = re.search(r"Adaptador\s+\w+\s+(.+?):", line)
                            if adapter_match:
                                current_adapter = adapter_match.group(1).strip()

                        # Verifica se é o IP atual
                        elif my_ip in line and current_adapter:
                            # Processa nome da interface para ficar compacto
                            interface_name = current_adapter

                            # Crop de interfaces longas (baseado nos seus exemplos)
                            if len(interface_name) > 12:
                                if "VirtualBox" in interface_name:
                                    interface_name = "VirtualBox"
                                elif "VMware Network Adapter VMnet" in interface_name:
                                    vmnet_match = re.search(
                                        r"VMnet(\d+)", interface_name
                                    )
                                    if vmnet_match:
                                        interface_name = f"VMnet{vmnet_match.group(1)}"
                                    else:
                                        interface_name = "VMware"
                                elif "vEthernet" in interface_name:
                                    # vEthernet (WSL) -> vEth(WSL)
                                    veth_match = re.search(
                                        r"vEthernet\s*\((.+?)\)", interface_name
                                    )
                                    if veth_match:
                                        inner = veth_match.group(1)
                                        if len(inner) > 8:
                                            inner = inner[:6] + ".."
                                        interface_name = f"vEth({inner})"
                                    else:
                                        interface_name = "vEthernet"
                                elif "Conexão de Rede Bluetooth" in interface_name:
                                    interface_name = "Bluetooth"
                                else:
                                    # Crop genérico para outros casos
                                    interface_name = interface_name[:10] + ".."
                            break
            except Exception:
                # Fallback para lógica anterior simplificada
                if "Ethernet" in current_adapter:
                    interface_name = "Ethernet"
                elif "Wi-Fi" in current_adapter or "Wireless" in current_adapter:
                    interface_name = "Wi-Fi"
                elif "VPN" in current_adapter:
                    interface_name = "VPN"
                else:
                    interface_name = "Rede"

        return gateway_ip, my_ip, interface_name
    except Exception:
        return "192.168.1.1", "127.0.0.1", "Local"


def get_rdp_interface_ip():
    """Detecta IP da interface usada para RDP"""
    try:
        result = subprocess.run(
            ["netstat", "-an"], capture_output=True, text=True, timeout=5
        )
        if result.returncode == 0:
            lines = result.stdout.split("\n")
            for line in lines:
                if ":3389" in line and "ESTABLISHED" in line:
                    match = re.search(
                        r"(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}):3389", line
                    )
                    if match:
                        local_ip = match.group(1)
                        if local_ip not in ["127.0.0.1", "0.0.0.0"]:
                            return local_ip
        _, my_ip, _ = get_network_info()
        return my_ip
    except Exception:
        _, my_ip, _ = get_network_info()
        return my_ip


def ping_host(host, timeout=3):
    """Ping host e retorna tempo em ms (-1 se falhou)"""
    try:
        if sys.platform == "win32":
            cmd = ["ping", "-n", "1", "-w", str(timeout * 1000), host]
        else:
            cmd = ["ping", "-c", "1", "-W", str(timeout), host]

        result = subprocess.run(
            cmd, capture_output=True, text=True, timeout=timeout + 1
        )

        if result.returncode == 0:
            if sys.platform == "win32":
                match = re.search(r"tempo[<=](\d+)ms|time[<=](\d+)ms", result.stdout)
                if match:
                    return int(match.group(1) or match.group(2))
            else:
                match = re.search(r"time=(\d+\.?\d*).*ms", result.stdout)
                if match:
                    return int(float(match.group(1)))
        return -1
    except Exception:
        return -1


def detectar_conexoes_rdp():
    """Detecta conexões RDP ativas"""
    try:
        conexoes_ativas = []

        if sys.platform == "win32":
            try:
                result = subprocess.run(
                    ["qwinsta"], capture_output=True, text=True, timeout=5
                )
                if result.returncode == 0:
                    lines = result.stdout.split("\n")
                    for line in lines:
                        if "rdp" in line.lower() and "ativo" in line.lower():
                            conexoes_ativas.append("Local-RDP")
            except Exception:
                pass

        try:
            result = subprocess.run(
                ["netstat", "-an"], capture_output=True, text=True, timeout=5
            )
            if result.returncode == 0:
                lines = result.stdout.split("\n")
                for line in lines:
                    if ":3389" in line and "ESTABLISHED" in line:
                        parts = line.split()
                        if len(parts) >= 2:
                            foreign_addr = parts[1]
                            match = re.search(
                                r"(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}):", foreign_addr
                            )
                            if match:
                                remote_ip = match.group(1)
                                if remote_ip not in ["127.0.0.1", "0.0.0.0"]:
                                    last_octet = remote_ip.split(".")[-1]
                                    conexoes_ativas.append(f"RDP-{last_octet}")
        except Exception:
            pass

        conexoes_ativas = list(dict.fromkeys(conexoes_ativas))
        tem_conexao = len(conexoes_ativas) > 0
        conexao_principal = conexoes_ativas[0] if conexoes_ativas else "Nenhuma"

        return tem_conexao, conexoes_ativas, conexao_principal
    except Exception:
        return False, [], "Erro"


def ping_site_brasileiro():
    """Ping para site brasileiro confiável"""
    sites_brasileiros = [
        ("itaipu.gov.br", "Itaipu"),
        ("uol.com.br", "UOL"),
    ]

    for site_url, site_name in sites_brasileiros:
        tempo = ping_host(site_url, timeout=2)
        if tempo > 0:
            return tempo, site_name

    return -1, "Brasil"


def get_computer_info():
    """Obtém informações do computador (nome e usuário)"""
    try:
        import getpass
        import socket

        computer_name = socket.gethostname()

        # Limita o nome se muito longo
        if len(computer_name) > 12:
            computer_name = computer_name[:9] + "..."

        return computer_name

    except Exception:
        return "Local"


def simulate_safe_activity():
    """Simula atividade segura com verificação"""
    try:
        # Movimento seguro no centro da tela
        screen_width, screen_height = pyautogui.size()
        center_x = screen_width // 2
        center_y = screen_height // 2
        safe_zone = min(screen_width, screen_height) // 20

        # Posição aleatória na área central
        target_x = center_x + random.randint(-safe_zone, safe_zone)
        target_y = center_y + random.randint(-safe_zone, safe_zone)

        # Movimento suave
        pyautogui.moveTo(target_x, target_y, duration=0.2)

        # Pequeno movimento adicional
        move_x = random.randint(-3, 3)
        move_y = random.randint(-3, 3)
        pyautogui.moveRel(move_x, move_y, duration=0.1)

        # Eventos de teclado seguros
        safe_keys = ["numlock", "scrolllock", "capslock"]
        selected_key = random.choice(safe_keys)
        pyautogui.press(selected_key)
        time.sleep(0.1)
        pyautogui.press(selected_key)

        return True, "Atividade simulada"
    except Exception as e:
        return False, f"Erro na simulação: {str(e)}"


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

    def add_log(self, message, is_orientation=False):
        """Adiciona mensagem ao log com timestamp"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_message = f"[{timestamp}] {message}"

        if is_orientation:
            # Apenas negrito sem fundo azul
            formatted_text = f'<span style="font-weight: bold;">{log_message}</span>'
            self.log_text.append(formatted_text)
        else:
            self.log_text.append(log_message)

        # Limite de 1000 linhas
        text_content = self.log_text.toPlainText()
        lines = text_content.split("\n")
        if len(lines) > 1000:
            lines = lines[-1000:]
            self.log_text.setPlainText("\n".join(lines))

        # Rola para o final
        scrollbar = self.log_text.verticalScrollBar()
        if scrollbar:
            scrollbar.setValue(scrollbar.maximum())

    def clear_log(self):
        """Limpa o log visual"""
        self.log_text.clear()

    def save_log(self):
        """Salva o log em arquivo usando seletor nativo"""
        try:
            # Gerar timestamp para nome do arquivo
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            default_filename = f"keep_alive_log_{timestamp}.txt"

            # Obter nome do arquivo usando diálogo
            filename, _ = QFileDialog.getSaveFileName(
                self,
                "Salvar Log",
                os.path.join(os.path.expanduser("~"), "Desktop", default_filename),
                "Arquivos de Texto (*.txt);;Todos os Arquivos (*)",
            )

            if not filename:
                return  # Usuário cancelou

            # Adicionar extensão se necessário
            if not filename.lower().endswith(".txt"):
                filename += ".txt"

            # Salvar arquivo
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

        self.random_intervals = QCheckBox("Usar intervalos variáveis (Δ%)")
        self.random_intervals.setChecked(True)
        options_layout.addWidget(self.random_intervals)

        self.minimize_to_tray_cb = QCheckBox("Minimizar para bandeja ao fechar")
        self.minimize_to_tray_cb.setChecked(True)
        options_layout.addWidget(self.minimize_to_tray_cb)

        layout.addWidget(options_group)
        layout.addSpacing(20)  # Espaço dobrado

        # Grupo de opções avançadas (sliders)
        advanced_options_group = QGroupBox("Opções Avançadas")
        advanced_layout = QVBoxLayout(advanced_options_group)

        # Slider de intervalo
        interval_container = QVBoxLayout()
        interval_label = QLabel("Intervalo entre tentativas (segundos):")
        advanced_layout.addWidget(interval_label)

        slider_layout = QHBoxLayout()
        self.interval_slider = QSlider(Qt.Orientation.Horizontal)
        self.interval_slider.setMinimum(INTERVAL_MIN)
        self.interval_slider.setMaximum(INTERVAL_MAX)
        self.interval_slider.setValue(DEFAULT_INTERVAL)
        self.interval_slider.setSingleStep(5)
        slider_layout.addWidget(self.interval_slider)

        self.interval_value = QLabel(str(DEFAULT_INTERVAL))
        self.interval_value.setFixedWidth(40)
        slider_layout.addWidget(self.interval_value)

        advanced_layout.addLayout(slider_layout)
        self.interval_slider.valueChanged.connect(
            lambda v: self.interval_value.setText(str(v))
        )
        advanced_layout.addSpacing(10)  # Espaço adicional

        # Slider de inatividade
        timeout_container = QVBoxLayout()
        timeout_label = QLabel("Inatividade mínima (segundos):")
        advanced_layout.addWidget(timeout_label)

        timeout_slider_layout = QHBoxLayout()
        self.timeout_slider = QSlider(Qt.Orientation.Horizontal)
        self.timeout_slider.setMinimum(USER_TIMEOUT_MIN)
        self.timeout_slider.setMaximum(USER_TIMEOUT_MAX)
        self.timeout_slider.setValue(DEFAULT_USER_TIMEOUT)
        self.timeout_slider.setSingleStep(5)
        timeout_slider_layout.addWidget(self.timeout_slider)

        self.timeout_value = QLabel(str(DEFAULT_USER_TIMEOUT))
        self.timeout_value.setFixedWidth(40)
        timeout_slider_layout.addWidget(self.timeout_value)

        advanced_layout.addLayout(timeout_slider_layout)
        self.timeout_slider.valueChanged.connect(
            lambda v: self.timeout_value.setText(str(v))
        )

        layout.addWidget(advanced_options_group)
        layout.addSpacing(20)  # Espaço dobrado

        # Botões de teste
        test_group = QGroupBox("Testes")
        test_layout = QVBoxLayout(test_group)

        test_button = QPushButton("Testar Simulação Agora")
        test_button.clicked.connect(self.test_simulation)
        test_layout.addWidget(test_button)

        layout.addWidget(test_group)
        layout.addStretch()

    def test_simulation(self):
        """Testa simulação"""
        try:
            # Verifica inatividade do usuário
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
    """Aba 'Sobre' com informações do projeto"""

    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.addStretch(1)

        # Caixa com informações
        about_frame = QFrame()
        about_frame.setFrameShape(QFrame.Shape.StyledPanel)
        about_layout = QVBoxLayout(about_frame)

        about_text = QLabel()
        about_text.setText(
            f"{APP_NAME} {APP_VERSION}\n\n"
            "- Correção temporização.\n"
            "- Mudança no layout de abas.\n"
            f"{APP_AUTHOR}\n"
        )
        about_text.setAlignment(Qt.AlignmentFlag.AlignCenter)
        about_layout.addWidget(about_text)

        # Link clicável
        link_label = QLabel(f'<a href="{APP_URL}">{APP_URL}</a>')
        link_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        link_label.setOpenExternalLinks(True)
        about_layout.addWidget(link_label)

        date_text = QLabel(APP_DATE)
        date_text.setAlignment(Qt.AlignmentFlag.AlignCenter)
        about_layout.addWidget(date_text)

        layout.addWidget(about_frame)
        layout.addSpacing(40)  # Espaço dobrado

        # Caixa para "Como usar"
        help_frame = QFrame()
        help_frame.setFrameShape(QFrame.Shape.StyledPanel)
        help_layout = QVBoxLayout(help_frame)

        help_title = QLabel("Como usar:")
        help_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        help_title.setStyleSheet("font-weight: bold;")
        help_layout.addWidget(help_title)

        help_text = QLabel(
            "• Utilize 'Iniciar Contínuo' para execução imediata contínua, sem interrupções\n"
            "• Utilize 'Iniciar Agendado' para execução nos horários definidos\n"
            "• Configure intervalos e inatividade mínima na aba Opções\n"
            "• Ajuste opções avançadas na mesma aba"
        )
        help_text.setAlignment(Qt.AlignmentFlag.AlignLeft)
        help_layout.addWidget(help_text)

        layout.addWidget(help_frame)
        layout.addStretch(2)


class KeepAliveApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.setFixedSize(500, 650)

        # Verifica instância única
        if is_already_running():
            QMessageBox.warning(
                None,
                "Aviso",
                f"O {APP_NAME} já está em execução!\n"
                "Verifique a bandeja do sistema ou o gerenciador de tarefas.\n"
                "Se o problema persistir, reinicie o computador.",
            )
            sys.exit(1)

        # Configurações
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

        # Ajusta timeout automaticamente
        self.screen_saver_time = get_screen_saver_timeout()
        self.rdp_time = get_rdp_disconnect_time()
        adjusted_timeout = adjust_user_timeout()

        if adjusted_timeout != self.default_user_timeout:
            self.default_user_timeout = adjusted_timeout
            self.settings.setValue("user_timeout", adjusted_timeout)

        self.is_running = False
        self.use_schedule = True
        self.activity_count = 0

        # Timers
        self.activity_timer = QTimer()
        self.activity_timer.timeout.connect(self.perform_activity)
        self.inactive_log_timer = QTimer()
        self.inactive_log_timer.timeout.connect(self.log_inactive_status)
        self.inactive_log_timer.start(INACTIVE_LOG_INTERVAL)
        self.help_timer = QTimer()
        self.help_timer.timeout.connect(self.log_help_message_if_inactive)
        self.help_timer.start(INACTIVE_LOG_INTERVAL)

        # Configura interface
        self.setup_ui()
        self.setup_tray()

        # Configurar conectividade
        self.setup_connectivity_timer()

    def setup_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # Status principal
        status_frame = QFrame()
        status_frame.setFrameShape(QFrame.Shape.StyledPanel)
        status_layout = QVBoxLayout(status_frame)

        self.status_label = QLabel("⭕ Aguardando início do serviço")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        status_layout.addWidget(self.status_label)

        self.execution_type_label = QLabel("")
        self.execution_type_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        status_layout.addWidget(self.execution_type_label)

        layout.addWidget(status_frame)

        # Abas
        self.tab_widget = QTabWidget()
        layout.addWidget(self.tab_widget)

        # ─────────── Aba principal ────────────────────────────────────────────────
        main_tab = QWidget()
        main_layout = QVBoxLayout(main_tab)

        # Agendamento (com horário atual)
        schedule_frame = QFrame()
        schedule_frame.setFrameShape(QFrame.Shape.StyledPanel)
        schedule_layout = QVBoxLayout(schedule_frame)

        schedule_title = QLabel("Agendamento de Execução")
        schedule_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        schedule_title.setStyleSheet("font-weight: bold; padding: 2px;")
        schedule_layout.addWidget(schedule_title)

        # Layout do agendamento - VERSÃO LIMPA SEM SOLAPAMENTO
        time_layout = QHBoxLayout()
        time_layout.addStretch()

        # Início
        time_layout.addWidget(QLabel("Início:"))
        self.start_time_edit = QTimeEdit()
        self.start_time_edit.setTime(self.default_start_time)
        self.start_time_edit.setFixedWidth(70)
        time_layout.addWidget(self.start_time_edit)

        time_layout.addSpacing(15)

        # Até
        spacer_label = QLabel("até")
        spacer_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        time_layout.addWidget(spacer_label)

        time_layout.addSpacing(15)

        # Término
        time_layout.addWidget(QLabel("Término:"))
        self.end_time_edit = QTimeEdit()
        self.end_time_edit.setTime(self.default_end_time)
        self.end_time_edit.setFixedWidth(70)
        time_layout.addWidget(self.end_time_edit)

        time_layout.addSpacing(40)

        # Horário Atual - FONTE BRANCA
        time_layout.addWidget(QLabel("Atual:"))
        self.current_time_label = QLabel("<b>--:--:--</b>")
        self.current_time_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.current_time_label.setStyleSheet(
            "font-size: 110%; color: white; font-weight: bold;"
        )
        time_layout.addWidget(self.current_time_label)

        time_layout.addStretch()

        schedule_layout.addLayout(time_layout)
        main_layout.addWidget(schedule_frame)
        main_layout.addSpacing(8)

        # Configuração do Sistema (4 colunas alinhadas)
        system_frame = QFrame()
        system_frame.setFrameShape(QFrame.Shape.StyledPanel)
        system_layout = QVBoxLayout(system_frame)

        system_title = QLabel("Configuração do Sistema")
        system_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        system_title.setStyleSheet("font-weight: bold; padding: 2px;")
        system_layout.addWidget(system_title)

        # Layout das 4 colunas - PERFEITAMENTE ALINHADO E CENTRALIZADO
        timeouts_layout = QHBoxLayout()
        timeouts_layout.addStretch()  # Espaço antes

        # Screen Saver - COLUNA 1
        ss_layout = QVBoxLayout()
        ss_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        ss_label = QLabel("Proteção de Tela:")
        ss_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        ss_time_formatted = (
            format_time_intelligent(self.screen_saver_time)
            if self.screen_saver_time > 0
            else "Desativado"
        )
        self.ss_value_label = QLabel(f"<b>{ss_time_formatted}</b>")
        self.ss_value_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        ss_layout.addWidget(ss_label)
        ss_layout.addWidget(self.ss_value_label)
        timeouts_layout.addLayout(ss_layout)

        timeouts_layout.addSpacing(30)  # Espaço uniforme

        # RDP Timeout - COLUNA 2
        rdp_layout = QVBoxLayout()
        rdp_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        rdp_label = QLabel("RDP Timeout:")
        rdp_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        rdp_time_formatted = (
            format_time_intelligent(self.rdp_time) if self.rdp_time > 0 else "N/D"
        )
        self.rdp_value_label = QLabel(f"<b>{rdp_time_formatted}</b>")
        self.rdp_value_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        rdp_layout.addWidget(rdp_label)
        rdp_layout.addWidget(self.rdp_value_label)
        timeouts_layout.addLayout(rdp_layout)

        timeouts_layout.addSpacing(30)  # Espaço uniforme

        # Meu IP - COLUNA 3
        ip_layout = QVBoxLayout()
        ip_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.ip_label = QLabel("IP:")  # ← SELF adicionado
        self.ip_label.setAlignment(Qt.AlignmentFlag.AlignCenter)  # ← SELF adicionado
        self.my_ip_label = QLabel("<b>Detectando...</b>")
        self.my_ip_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        ip_layout.addWidget(self.ip_label)  # ← SELF adicionado
        ip_layout.addWidget(self.my_ip_label)
        timeouts_layout.addLayout(ip_layout)

        timeouts_layout.addSpacing(30)  # Espaço uniforme

        # Computador - COLUNA 4 (CENTRALIZADA)
        computer_layout = QVBoxLayout()
        computer_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        computer_label = QLabel("Computador:")
        computer_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.computer_name_label = QLabel("<b>Detectando...</b>")
        self.computer_name_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        computer_layout.addWidget(computer_label)
        computer_layout.addWidget(self.computer_name_label)
        timeouts_layout.addLayout(computer_layout)

        timeouts_layout.addStretch()  # Espaço depois

        system_layout.addLayout(timeouts_layout)
        main_layout.addWidget(system_frame)

        # Conectividade (layout otimizado)
        connectivity_frame = QFrame()
        connectivity_frame.setFrameShape(QFrame.Shape.StyledPanel)
        connectivity_layout = QVBoxLayout(connectivity_frame)

        connectivity_title = QLabel("Conectividade de Rede")
        connectivity_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        connectivity_title.setStyleSheet("font-weight: bold; padding: 2px;")
        connectivity_layout.addWidget(connectivity_title)

        # Layout principal da conectividade (2 linhas organizadas)
        connectivity_main_layout = QVBoxLayout()

        # Linha 1: Status RDP (centralizado)
        rdp_row = QHBoxLayout()
        rdp_row.addStretch()

        rdp_status_label_title = QLabel("Status RDP:")
        self.rdp_status_label = QLabel("<b>Verificando...</b>")
        self.rdp_combo = QComboBox()
        self.rdp_combo.setMaximumWidth(150)
        self.rdp_combo.setVisible(False)

        rdp_row.addWidget(rdp_status_label_title)
        rdp_row.addSpacing(8)
        rdp_row.addWidget(self.rdp_status_label)
        rdp_row.addSpacing(15)
        rdp_row.addWidget(self.rdp_combo)
        rdp_row.addStretch()

        # Linha 2: Latências (distribuídas uniformemente)
        ping_row = QHBoxLayout()
        ping_row.addStretch()

        # Gateway
        gateway_container = QHBoxLayout()
        gateway_container.addWidget(QLabel("Gateway:"))
        self.gateway_ping_label = QLabel("<b>...</b>")
        gateway_container.addWidget(self.gateway_ping_label)

        # Brasil
        brasil_container = QHBoxLayout()
        brasil_container.addWidget(QLabel("Itaipu:"))
        self.brasil_ping_label = QLabel("<b>...</b>")
        brasil_container.addWidget(self.brasil_ping_label)

        # Sistema (dinâmico)
        sistema_container = QHBoxLayout()
        sistema_label = QLabel("Sistema:")
        self.sistema_ping_label = QLabel("<b>...</b>")
        sistema_container.addWidget(sistema_label)
        sistema_container.addWidget(self.sistema_ping_label)

        # Widget container para sistema (para controlar visibilidade)
        self.sistema_widget = QWidget()
        sistema_widget_layout = QHBoxLayout(self.sistema_widget)
        sistema_widget_layout.setContentsMargins(0, 0, 0, 0)
        sistema_widget_layout.addLayout(sistema_container)
        self.sistema_widget.setVisible(False)

        # Adiciona pings à linha
        ping_row.addLayout(gateway_container)
        ping_row.addSpacing(25)
        ping_row.addLayout(brasil_container)
        ping_row.addSpacing(25)
        ping_row.addWidget(self.sistema_widget)
        ping_row.addStretch()

        # Montagem das linhas
        connectivity_main_layout.addLayout(rdp_row)
        connectivity_main_layout.addSpacing(5)
        connectivity_main_layout.addLayout(ping_row)

        connectivity_layout.addLayout(connectivity_main_layout)
        main_layout.addWidget(connectivity_frame)
        main_layout.addSpacing(8)

        # Log de Atividade (altura reduzida)
        user_group = QGroupBox("Verificação de Atividade do Usuário")
        user_group_layout = QVBoxLayout(user_group)
        self.main_log_text = QTextEdit()
        self.main_log_text.setReadOnly(True)
        self.main_log_text.setMaximumHeight(110)
        self.main_log_text.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed
        )
        user_group_layout.addWidget(self.main_log_text)
        main_layout.addWidget(user_group)
        main_layout.addSpacing(12)

        # Botões Reiniciar/Fechar (PRESERVADOS)
        restart_quit_layout = QHBoxLayout()
        restart_quit_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.restart_button = QPushButton(STR_RESTART_APP)
        self.restart_button.clicked.connect(self.restart_application)
        self.restart_button.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed
        )
        restart_quit_layout.addWidget(self.restart_button)

        self.quit_button = QPushButton(STR_QUIT)
        self.quit_button.clicked.connect(self.quit_application)
        self.quit_button.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed
        )
        restart_quit_layout.addWidget(self.quit_button)

        main_layout.addLayout(restart_quit_layout)

        self.tab_widget.addTab(main_tab, "Principal")
        # Final Aba Principal
        # ─────────────────────────────────────────────────────────────────────────

        # Abas restantes
        self.advanced_tab = AdvancedTab(self.tab_widget)
        self.tab_widget.addTab(self.advanced_tab, "Opções")

        self.log_tab = LogTab(self.tab_widget)
        self.tab_widget.addTab(self.log_tab, "Log")

        about_tab = AboutTab(self.tab_widget)
        self.tab_widget.addTab(about_tab, "Sobre")

        # Botões inferiores (mesmo tamanho) - ABAIXO DAS ABAS
        button_layout = QHBoxLayout()
        button_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.toggle_button = QPushButton(STR_START_CONTINUOUS)
        self.toggle_button.setFixedSize(120, 40)
        self.toggle_button.setStyleSheet("font-weight: bold;")
        self.toggle_button.clicked.connect(self.toggle_service_no_schedule)

        self.schedule_button = QPushButton(STR_START_SCHEDULED)
        self.schedule_button.setFixedSize(120, 40)
        self.schedule_button.clicked.connect(self.toggle_service_with_schedule)

        self.stop_button = QPushButton(STR_STOP_SERVICE)
        self.stop_button.setFixedSize(120, 40)
        self.stop_button.clicked.connect(self.stop_service)

        self.minimize_button = QPushButton(STR_MINIMIZE)
        self.minimize_button.setFixedSize(120, 40)
        self.minimize_button.clicked.connect(self.hide)

        button_layout.addWidget(self.toggle_button)
        button_layout.addWidget(self.schedule_button)
        button_layout.addWidget(self.stop_button)
        button_layout.addWidget(self.minimize_button)
        layout.addLayout(button_layout)

    def add_main_log(self, message, is_orientation=False):
        """Adiciona mensagem ao log principal (15 linhas)"""
        # timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        timestamp = datetime.now().strftime("%H:%M:%S")  # Só hora:min:seg
        log_message = f"[{timestamp}] {message}"

        if is_orientation:
            # Apenas negrito sem fundo azul
            formatted_text = f'<span style="font-weight: bold;">{log_message}</span>'
            self.main_log_text.append(formatted_text)
        else:
            self.main_log_text.append(log_message)

        # Limite de 15 linhas
        text_content = self.main_log_text.toPlainText()
        lines = text_content.split("\n")
        if len(lines) > 15:
            lines = lines[-15:]
            self.main_log_text.setPlainText("\n".join(lines))

        # Rola para o final
        scrollbar = self.main_log_text.verticalScrollBar()
        if scrollbar:
            scrollbar.setValue(scrollbar.maximum())

    def update_execution_type_label(self):
        """Atualiza o label do tipo de execução"""
        if not self.is_running:
            self.execution_type_label.setText("")
        elif not self.use_schedule:
            self.execution_type_label.setText(STR_CONTINUOUS_RUNNING)
        else:
            start_time = self.start_time_edit.time().toString("HH:mm")
            end_time = self.end_time_edit.time().toString("HH:mm")
            self.execution_type_label.setText(
                STR_SCHEDULE_RUNNING.format(start_time, end_time)
            )

    def setup_tray(self):
        """Configura o ícone na bandeja"""
        try:
            app_style = QApplication.style()
            icon = app_style.standardIcon(QStyle.StandardPixmap.SP_ComputerIcon)
            self.tray_icon = QSystemTrayIcon(icon, self)
            self.tray_icon.setToolTip(APP_NAME)

            menu = QMenu()
            show_action = QAction("Mostrar", self)
            show_action.triggered.connect(self.show)
            self.toggle_tray_action = QAction(STR_START_CONTINUOUS, self)
            self.toggle_tray_action.triggered.connect(self.toggle_service_no_schedule)

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
        self.advanced_tab.enable_mouse.setChecked(
            self.settings.value("enable_mouse", True, bool)
        )
        self.advanced_tab.enable_keyboard.setChecked(
            self.settings.value("enable_keyboard", True, bool)
        )
        self.advanced_tab.random_intervals.setChecked(
            self.settings.value("random_intervals", True, bool)
        )
        self.advanced_tab.interval_slider.setValue(
            self.settings.value("interval", DEFAULT_INTERVAL, int)
        )
        self.advanced_tab.timeout_slider.setValue(self.default_user_timeout)

    def save_settings(self):
        """Salva configurações"""
        self.settings.setValue("interval", self.advanced_tab.interval_slider.value())
        self.settings.setValue("user_timeout", self.advanced_tab.timeout_slider.value())
        self.settings.setValue("start_time", self.start_time_edit.time())
        self.settings.setValue("end_time", self.end_time_edit.time())
        self.settings.setValue(
            "minimize_to_tray", self.advanced_tab.minimize_to_tray_cb.isChecked()
        )
        self.settings.setValue(
            "enable_mouse", self.advanced_tab.enable_mouse.isChecked()
        )
        self.settings.setValue(
            "enable_keyboard", self.advanced_tab.enable_keyboard.isChecked()
        )
        self.settings.setValue(
            "random_intervals", self.advanced_tab.random_intervals.isChecked()
        )

        self.add_filtered_log(STR_SETTINGS_SAVED)
        # self.log_tab.add_log(STR_SETTINGS_SAVED)
        # self.add_main_log(STR_SETTINGS_SAVED)

    def check_schedule(self):
        """Verifica horário de funcionamento"""
        if not self.use_schedule:
            return True

        current_time = QTime.currentTime()
        start_time = self.start_time_edit.time()
        end_time = self.end_time_edit.time()

        if start_time <= end_time:
            return start_time <= current_time <= end_time
        else:
            return current_time >= start_time or current_time <= end_time

    def log_inactive_status(self):
        """Loga status de inatividade"""
        if not self.is_running:
            status_msg = STR_SERVICE_STOPPED
            help_msg = "<b>Inicie o programa clicando na opção desejada</b>"
            self.log_tab.log_text.append(help_msg)
            self.main_log_text.append(help_msg)

            # self.add_filtered_log(status_msg)
            # self.log_tab.add_log(status_msg)
            # self.add_main_log(status_msg)

            # Adiciona mensagem de ajuda
            # timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            # formatted_help = f"[{timestamp}] {help_msg}"
            # self.log_tab.log_text.append(formatted_help)
            # self.main_log_text.append(formatted_help)

    def log_help_message(self):
        """Loga mensagem de orientação"""
        # help_msg = "Utilize 'Iniciar Contínuo' para execução imediata contínua ou 'Iniciar Agendado' para execução no horário selecionado."
        # self.log_tab.add_log(help_msg)
        # self.log_tab.add_log(help_msg, is_orientation=True)  para negrito
        # self.add_main_log(help_msg)
        # self.add_filtered_log(help_msg)
        help_msg_full = "Utilize 'Iniciar Contínuo' para execução imediata contínua ou 'Iniciar Agendado' para execução exclusivamente no horário selecionado."
        help_msg_short = "Utilize 'Iniciar Contínuo' para execução imediata ou 'Iniciar Agendado' para execução exclusivamente no horário selecionado."

        # Log completo (com timestamp)
        self.log_tab.add_log(help_msg_full)

        # Log principal (SEM timestamp)
        self.main_log_text.append(help_msg_short)

    def log_help_message_if_inactive(self):
        """Loga orientação quando inativo"""
        if not self.is_running:
            self.log_help_message()

    def setup_connectivity_timer(self):
        """Configura timer para conectividade"""
        self.connectivity_timer = QTimer()
        self.connectivity_timer.timeout.connect(self.update_connectivity_info)
        self.connectivity_timer.start(15000)  # 15 segundos
        self.connectivity_log_counter = 0
        self.last_gateway = ""
        self.last_my_ip = ""
        self.last_interface = ""

        # Timer para horário atual (atualiza a cada segundo)
        self.current_time_timer = QTimer()
        self.current_time_timer.timeout.connect(self.update_current_time)
        self.current_time_timer.start(1000)
        self.update_current_time()  # Atualização inicial

        QTimer.singleShot(2000, self.update_connectivity_info)

    def update_connectivity_info(self):
        """Atualiza informações de conectividade"""
        try:
            gateway, my_ip, interface = get_network_info()
            tem_rdp, lista_rdp, conexao_principal = detectar_conexoes_rdp()

            ip_usado = get_rdp_interface_ip() if tem_rdp else my_ip
            ping_gw = ping_host(gateway, timeout=2)
            ping_brasil, site_brasil = ping_site_brasileiro()

            ping_sistema = -1
            sistema_nome = ""
            if tem_rdp and conexao_principal not in ["Local-RDP", "Nenhuma", "Erro"]:
                if conexao_principal.startswith("RDP-"):
                    last_octet = conexao_principal.replace("RDP-", "")
                    gateway_base = ".".join(gateway.split(".")[:-1])
                    sistema_ip = f"{gateway_base}.{last_octet}"
                    ping_sistema = ping_host(sistema_ip, timeout=2)
                    sistema_nome = conexao_principal

            self.update_connectivity_display(
                gateway,
                my_ip,
                ip_usado,
                interface,
                tem_rdp,
                lista_rdp,
                conexao_principal,
                ping_gw,
                ping_brasil,
                ping_sistema,
                site_brasil,
                sistema_nome,
            )

            self.connectivity_log_counter += 1
            if self.connectivity_log_counter >= 4:  # Log a cada 1 minuto
                self.connectivity_log_counter = 0
                self.log_connectivity_info(
                    tem_rdp,
                    conexao_principal,
                    ping_gw,
                    ping_brasil,
                    site_brasil,
                    ping_sistema,
                    sistema_nome,
                )

            self.last_gateway = gateway
            self.last_my_ip = ip_usado
            self.last_interface = interface

        except Exception as e:
            print(f"[DEBUG] Erro conectividade: {e}")

    def update_current_time(self):
        """Atualiza horário atual na interface"""
        try:
            current_time = QTime.currentTime()
            time_str = current_time.toString("HH:mm:ss")
            self.current_time_label.setText(f"<b>{time_str}</b>")
        except Exception:
            pass

    def update_connectivity_display(
        self,
        gateway,
        my_ip_original,
        my_ip_usado,
        interface,
        tem_rdp,
        lista_rdp,
        conexao_principal,
        ping_gw,
        ping_brasil,
        ping_sistema,
        site_brasil,
        sistema_nome,
    ):
        """Atualiza display de conectividade"""
        try:
            if not hasattr(self, "_computer_name_set"):
                computer_name = get_computer_info()
                self.computer_name_label.setText(f"<b>{computer_name}</b>")
                self._computer_name_set = True
            # Meu IP
            # Atualiza LABEL "IP:" com nome da interface
            self.ip_label.setText(f"IP ({interface}):")

            # IP simples - SEM interface no valor (já está no label)
            if my_ip_usado != my_ip_original:
                ip_display = (
                    f"<b style='color: orange;'>{my_ip_usado}</b> <small>→RDP</small>"
                )
            else:
                ip_display = f"<b>{my_ip_usado}</b>"
            self.my_ip_label.setText(ip_display)

            # Status RDP
            if tem_rdp:
                rdp_text = f"<b style='color: green;'>Ativo</b>"
                if len(lista_rdp) > 1:
                    rdp_text += f" <small>({len(lista_rdp)})</small>"
            else:
                rdp_text = "<b style='color: gray;'>Inativo</b>"
            self.rdp_status_label.setText(rdp_text)

            # Dropdown RDP
            self.rdp_combo.clear()
            if lista_rdp and len(lista_rdp) > 1:
                for conexao in lista_rdp:
                    display_name = conexao.replace("Local-RDP", "Local").replace(
                        "RDP-", "Sistema ."
                    )
                    self.rdp_combo.addItem(display_name)
                self.rdp_combo.setVisible(True)
            else:
                self.rdp_combo.setVisible(False)

            # Formatação de ping
            def format_ping(ping_time):
                if ping_time > 0:
                    if ping_time < 30:
                        return f"<b style='color: green;'>{ping_time}ms</b>"
                    elif ping_time < 100:
                        return f"<b style='color: orange;'>{ping_time}ms</b>"
                    else:
                        return f"<b style='color: red;'>{ping_time}ms</b>"
                else:
                    return "<b style='color: red;'>---</b>"

            # Labels de ping
            self.gateway_ping_label.setText(format_ping(ping_gw))
            self.brasil_ping_label.setText(format_ping(ping_brasil))

            if ping_sistema > 0 and sistema_nome:
                sistema_display = sistema_nome.replace("RDP-", "")
                # self.brasil_ping_label.setText(f"Itaipu: {format_ping(ping_brasil)}")
                self.sistema_ping_label.setText(f"{format_ping(ping_sistema)}")
                self.sistema_widget.setVisible(True)
            else:
                self.sistema_widget.setVisible(False)

        except Exception as e:
            print(f"[DEBUG] Erro display: {e}")

    def log_connectivity_info(
        self,
        tem_rdp,
        conexao_principal,
        ping_gw,
        ping_brasil,
        site_brasil,
        ping_sistema,
        sistema_nome,
    ):
        """Loga conectividade de forma compacta"""
        try:
            status_parts = []

            if tem_rdp:
                if conexao_principal == "Local-RDP":
                    status_parts.append("RDP:Local")
                else:
                    sistema_short = conexao_principal.replace("RDP-", ".")
                    status_parts.append(f"RDP:{sistema_short}")
            else:
                status_parts.append("RDP:Off")

            if ping_gw > 0:
                status_parts.append(f"GW:{ping_gw}ms")
            if ping_brasil > 0:
                status_parts.append(f"BR:{ping_brasil}ms")
            if ping_sistema > 0:
                sistema_short = sistema_nome.replace("RDP-", "S")
                status_parts.append(f"{sistema_short}:{ping_sistema}ms")

            log_msg = "Net: " + " | ".join(status_parts)
            self.add_filtered_log(log_msg)
            # self.log_tab.add_log(log_msg)
            # self.add_main_log(log_msg)

        except Exception as e:
            print(f"[DEBUG] Erro log: {e}")

    def closeEvent(self, event):
        self.save_settings()
        if self.advanced_tab.minimize_to_tray_cb.isChecked():
            event.ignore()
            self.hide()
            self.tray_icon.showMessage(
                APP_NAME,
                "Executando em segundo plano",
                QSystemTrayIcon.MessageIcon.Information,
                2000,
            )
        else:
            self.quit_application()

    def toggle_service_no_schedule(self):
        """Iniciar sem agendamento"""
        # Para qualquer serviço em execução
        if self.is_running:
            self.stop_service()

        # Inicia modo contínuo
        self.use_schedule = False
        self.start_service()

    def toggle_service_with_schedule(self):
        """Iniciar com agendamento"""
        # Para qualquer serviço em execução
        if self.is_running:
            self.stop_service()

        # Inicia modo agendado
        self.use_schedule = True
        self.start_service_with_schedule()

    def start_service_with_schedule(self):
        """Inicia serviço com verificação de agendamento"""
        self.add_filtered_log("Iniciado agendamento")
        # self.log_tab.add_log("Iniciado agendamento")
        # self.add_main_log("Iniciado agendamento")

        if self.check_schedule():
            self.start_service()
        else:
            start_time = self.start_time_edit.time().toString("HH:mm")
            end_time = self.end_time_edit.time().toString("HH:mm")
            self.status_label.setText(STR_SERVICE_STOPPED)
            self.update_execution_type_label()
            self.add_filtered_log("Serviço parado - Fora do horário de agendamento")
            # self.log_tab.add_log("Serviço parado - Fora do horário de agendamento")
            # self.add_main_log("Serviço parado - Fora do horário de agendamento")

    # ────────────────────────── start_service ──────────────────────────────
    def start_service(self):
        """Inicia o serviço"""
        base_interval = self.advanced_tab.interval_slider.value()  # em segundos

        if self.advanced_tab.random_intervals.isChecked():
            variation = base_interval * 0.30
            rand_secs = random.uniform(
                base_interval - variation, base_interval + variation
            )
        else:
            rand_secs = float(base_interval)

        next_interval_ms = int(rand_secs * 1000)
        next_time = datetime.now() + timedelta(seconds=rand_secs)

        # Log completo (detalhado)
        log_msg_full = (
            f"{STR_SERVICE_STARTED}."
            f" Simulação: {base_interval}s ±δ → {rand_secs:.1f}s"
            f" → {next_time.strftime('%H:%M:%S')}"
        )

        # Log principal (compacto)
        log_msg_short = (
            f"Próxima atividade ({rand_secs:.1f}s): {next_time.strftime('%H:%M:%S')}"
        )

        self.activity_timer.start(next_interval_ms)
        self.is_running = True
        self.status_label.setText(STR_SERVICE_RUNNING)
        self.update_execution_type_label()

        # Adiciona aos logs separadamente
        self.log_tab.add_log(log_msg_full)  # Log completo
        self.add_main_log(log_msg_short)  # Log principal customizado
        # self.log_tab.add_log(log_msg)
        # self.add_main_log(log_msg)

    def stop_service(self):
        """Para o serviço"""
        if not self.is_running:
            return

        self.activity_timer.stop()
        self.is_running = False
        self.status_label.setText(STR_SERVICE_STOPPED)
        self.update_execution_type_label()
        ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)
        self.add_filtered_log(STR_SERVICE_STOPPED_LOG)
        # self.log_tab.add_log(STR_SERVICE_STOPPED_LOG)
        # self.add_main_log(STR_SERVICE_STOPPED_LOG)

    def restart_application(self):
        """Reinicia completamente a aplicação"""
        self.save_settings()
        python = sys.executable
        os.execl(python, python, *sys.argv)

    def perform_activity(self):
        try:
            # Verifica agendamento
            if self.use_schedule and not self.check_schedule():
                self.stop_service()
                return

            # Verifica inatividade do usuário
            idle_time = get_user_activity_timeout()
            timeout_limit = self.advanced_tab.timeout_slider.value()

            if idle_time < timeout_limit:
                self.activity_count += 1
                cancel_message = STR_USER_ACTIVE.format(idle_time, timeout_limit)
                self.add_filtered_log(cancel_message)
                # self.log_tab.add_log(cancel_message)
                # self.add_main_log(cancel_message)

                if self.is_running:
                    self.schedule_next_activity()
                return

            # Prevenir bloqueio
            prevent_system_lock()

            # Simular atividade
            activity_success, activity_message = simulate_safe_activity()
            self.activity_count += 1
            now = datetime.now().strftime("%H:%M:%S")
            status_msg = f"Atividade #{self.activity_count} em {now}"

            if activity_success:
                self.status_label.setText(status_msg + " (Sucesso)")
                self.add_filtered_log(STR_SIMULATION_SUCCESS)
                # self.log_tab.add_log(STR_SIMULATION_SUCCESS)
                # self.add_main_log(STR_SIMULATION_SUCCESS)
            else:
                self.status_label.setText(status_msg + f" ({activity_message})")
                error_message = f"ERRO: {activity_message}"
                self.log_tab.add_log(error_message)
                self.add_main_log(error_message)

            # Agendar próxima
            if self.is_running:
                self.schedule_next_activity()

        except Exception as e:
            error_msg = f"Erro na atividade #{self.activity_count}: {str(e)}"
            self.status_label.setText(error_msg)
            critical_error = f"ERRO CRÍTICO: {error_msg}"
            self.log_tab.add_log(critical_error)
            self.add_main_log(critical_error)

    def schedule_next_activity(self):
        """Agenda próxima atividade"""
        self.activity_timer.stop()
        base_interval = self.advanced_tab.interval_slider.value()  # em segundos

        if self.advanced_tab.random_intervals.isChecked():
            variation = base_interval * 0.30
            rand_secs = random.uniform(
                base_interval - variation, base_interval + variation
            )
            interval_text = f"{base_interval}s ±30% → {rand_secs:.1f}s"
        else:
            rand_secs = float(base_interval)
            interval_text = f"{base_interval}s fixo"

        next_interval_ms = int(rand_secs * 1000)
        self.activity_timer.setInterval(next_interval_ms)
        self.activity_timer.start()

        next_time = datetime.now() + timedelta(seconds=rand_secs)
        next_message = STR_NEXT_ACTIVITY.format(
            next_time.strftime("%H:%M:%S"), f"{rand_secs:.1f}"
        )

        self.add_filtered_log(next_message)
        # self.log_tab.add_log(next_message)
        # self.add_main_log(next_message)

    def filter_log_message(self, message):
        """
        NOVA FUNÇÃO: Determina se mensagem é importante para log principal
        """
        # Mensagens IMPORTANTES (aparecem no log principal)
        important_keywords = [
            "Atividade Executada",
            "Usuário Ativo",
            "Próxima Atividade",
            "Serviço Iniciado",
            "Serviço Parado",
            "Keep Alive RDP Connection iniciado",
            "iniciado",
            "agendamento",
            "Erro",
            "Falha",
            "ERRO",
        ]

        # Mensagens FILTRADAS (só no log completo)
        filtered_keywords = [
            "Net: RDP:Off",
            "Net: RDP:Local",
            "GW:",
            "BR:",
            "|",  # Linhas de conectividade têm "|"
            "Keep Alive RDP Connection iniciado",
            "Usuário Ativo",
        ]

        # Se contém palavra-chave importante
        for keyword in important_keywords:
            if keyword in message:
                return True

        # Se contém palavra-chave filtrada
        for keyword in filtered_keywords:
            if keyword in message:
                return False

        # Por padrão, considera importante
        return True

    def add_filtered_log(self, message, is_orientation=False):
        """
        NOVA FUNÇÃO: Adiciona log com filtro inteligente
        """
        # SEMPRE adiciona ao log completo (aba Log)
        self.log_tab.add_log(message, is_orientation)

        # Só adiciona ao log principal se for importante
        if self.filter_log_message(message):
            self.add_main_log(message, is_orientation)

    def quit_application(self):
        try:
            self.save_settings()
            self.log_tab.add_log("Aplicativo encerrado")
            self.add_main_log("Aplicativo encerrado")
        except Exception:
            pass

        try:
            self.activity_timer.stop()
            self.inactive_log_timer.stop()
            self.help_timer.stop()
            self.connectivity_timer.stop()
            self.current_time_timer.stop()
            self.tray_icon.hide()
            cleanup_lock()
        except Exception:
            pass

        QApplication.quit()


def main():
    """Função principal"""
    QLoggingCategory.setFilterRules("qt.qpa.paint.debug=false")
    app = QApplication(sys.argv)

    try:
        app.setStyle("Windows")
    except Exception:
        pass

    window = KeepAliveApp()
    window.show()

    # Mensagem inicial
    # window.log_tab.add_log(f"{APP_NAME} iniciado")
    # window.add_main_log(f"{APP_NAME} iniciado")
    # window.add_filtered_log(f"{APP_NAME} iniciado")
    # window.log_inactive_status()
    window.log_tab.add_log(f"{APP_NAME} iniciado")  # Só no log completo
    window.log_help_message()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
