# =============================================================================
# 🔥 GERAÇÃO DE EXE OTIMIZADO - KeepAlive RDP Connection
# =============================================================================

import os
import shutil
import subprocess
import sys
from pathlib import Path


def detect_python_versions():
    """Detecta versões do Python disponíveis"""
    versions = []

    # Verificar Python padrão
    try:
        result = subprocess.run(
            [sys.executable, "--version"], capture_output=True, text=True
        )
        if result.returncode == 0:
            version = result.stdout.strip().replace("Python ", "")
            versions.append((sys.executable, version, "atual"))
    except:
        pass

    # Verificar outras versões comuns
    common_paths = [
        r"C:\Python310\python.exe",
        r"C:\Python311\python.exe",
        r"C:\Python312\python.exe",
        r"C:\Python39\python.exe",
        "python",
        "python3",
        "python3.10",
        "python3.11",
        "python3.12",
    ]

    for path in common_paths:
        try:
            result = subprocess.run([path, "--version"], capture_output=True, text=True)
            if result.returncode == 0:
                version = result.stdout.strip().replace("Python ", "")
                if (path, version, "") not in [(v[0], v[1], "") for v in versions]:
                    versions.append((path, version, ""))
        except:
            continue

    return versions


def select_python_version():
    """Permite selecionar versão do Python"""
    versions = detect_python_versions()

    if not versions:
        print("❌ Nenhuma versão do Python encontrada!")
        return None

    print("\n🐍 Versões do Python encontradas:")
    for i, (path, version, note) in enumerate(versions):
        note_str = f" ({note})" if note else ""
        print(f"{i+1}. Python {version}{note_str} - {path}")

    while True:
        try:
            choice = int(input(f"\nEscolha a versão (1-{len(versions)}): ")) - 1
            if 0 <= choice < len(versions):
                selected = versions[choice]
                print(f"✅ Selecionado: Python {selected[1]} - {selected[0]}")
                return selected[0]
            else:
                print("❌ Opção inválida!")
        except ValueError:
            print("❌ Digite um número válido!")


def install_dependencies(python_exe):
    """Instala dependências necessárias"""
    print("\n📦 Verificando e instalando dependências...")

    dependencies = [
        "pyinstaller>=6.0.0",
        "PyQt6>=6.0.0",
        "pyautogui>=0.9.50",
        "pywin32>=300",
    ]

    for dep in dependencies:
        print(f"📦 Instalando {dep}...")
        try:
            result = subprocess.run(
                [python_exe, "-m", "pip", "install", dep],
                capture_output=True,
                text=True,
            )
            if result.returncode == 0:
                print(f"✅ {dep} instalado com sucesso")
            else:
                print(f"⚠️ Aviso em {dep}: {result.stderr.strip()}")
        except Exception as e:
            print(f"❌ Erro instalando {dep}: {e}")

    print("✅ Verificação de dependências concluída!")


def build_optimized_exe(python_exe=None):
    """Gera EXE otimizado com menor tamanho possível"""

    if python_exe is None:
        python_exe = sys.executable

    print(f"🔥 Iniciando build otimizado com Python: {python_exe}")

    # Nome do arquivo Python principal
    main_file = "keep-alive-app.py"

    # Verificar se arquivo existe
    if not os.path.exists(main_file):
        print(f"❌ Arquivo {main_file} não encontrado!")
        print("📁 Arquivos .py encontrados:")
        for file in os.listdir("."):
            if file.endswith(".py"):
                print(f"   - {file}")
        return False

    # Comando PyInstaller ULTRA OTIMIZADO
    cmd = [
        python_exe,
        "-m",
        "PyInstaller",  # Usar Python selecionado
        "--onefile",  # Um único arquivo EXE
        "--windowed",  # Sem console (GUI)
        "--optimize=2",  # Otimização máxima Python
        "--strip",  # Remove símbolos de debug
        "--noupx",  # Desabilita UPX (pode dar problemas)
        # 🎯 EXCLUSÕES AGRESSIVAS para reduzir tamanho
        "--exclude-module=tkinter",
        "--exclude-module=matplotlib",
        "--exclude-module=numpy",
        "--exclude-module=pandas",
        "--exclude-module=scipy",
        "--exclude-module=PIL",
        "--exclude-module=cv2",
        "--exclude-module=tensorflow",
        "--exclude-module=torch",
        "--exclude-module=django",
        "--exclude-module=flask",
        "--exclude-module=requests",
        "--exclude-module=urllib3",
        "--exclude-module=cryptography",
        "--exclude-module=sqlite3",
        "--exclude-module=xml",
        "--exclude-module=html",
        "--exclude-module=email",
        "--exclude-module=http",
        "--exclude-module=json",
        "--exclude-module=pickle",
        "--exclude-module=zipfile",
        "--exclude-module=tarfile",
        "--exclude-module=gzip",
        "--exclude-module=bz2",
        "--exclude-module=lzma",
        "--exclude-module=unittest",
        "--exclude-module=doctest",
        "--exclude-module=pdb",
        "--exclude-module=profile",
        "--exclude-module=pstats",
        "--exclude-module=timeit",
        "--exclude-module=trace",
        "--exclude-module=turtle",
        "--exclude-module=idlelib",
        "--exclude-module=lib2to3",
        # 🚫 Exclusões específicas PyQt6 desnecessárias
        "--exclude-module=PyQt6.QtWebEngine",
        "--exclude-module=PyQt6.QtWebEngineWidgets",
        "--exclude-module=PyQt6.QtWebEngineCore",
        "--exclude-module=PyQt6.QtMultimedia",
        "--exclude-module=PyQt6.QtMultimediaWidgets",
        "--exclude-module=PyQt6.QtOpenGL",
        "--exclude-module=PyQt6.QtOpenGLWidgets",
        "--exclude-module=PyQt6.QtSql",
        "--exclude-module=PyQt6.QtTest",
        "--exclude-module=PyQt6.QtDesigner",
        "--exclude-module=PyQt6.QtHelp",
        "--exclude-module=PyQt6.QtNetwork",
        "--exclude-module=PyQt6.QtPrintSupport",
        "--exclude-module=PyQt6.QtQml",
        "--exclude-module=PyQt6.QtQuick",
        "--exclude-module=PyQt6.QtSvg",
        "--exclude-module=PyQt6.QtSvgWidgets",
        "--exclude-module=PyQt6.Qt3D",
        "--exclude-module=PyQt6.QtCharts",
        "--exclude-module=PyQt6.QtDataVisualization",
        # 📂 Nome e localização do EXE
        "--name=KeepAliveRDP",
        "--distpath=dist_optimized",  # Pasta de saída
        "--workpath=build_temp",  # Pasta temporária
        "--specpath=specs",  # Pasta dos specs
        # 🎯 ÍCONE (usa o mesmo da bandeja do sistema)
        "--icon=NONE",  # Vamos usar ícone interno do Windows
        main_file,
    ]

    try:
        print("⚡ Executando PyInstaller...")
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        print("✅ Build concluído com sucesso!")

        # Verificar tamanho do arquivo gerado
        exe_path = Path("dist_optimized/KeepAliveRDP.exe")
        if exe_path.exists():
            size_mb = exe_path.stat().st_size / (1024 * 1024)
            print(f"📦 Tamanho do EXE: {size_mb:.1f} MB")
            print(f"📁 Localização: {exe_path.absolute()}")

        return True

    except subprocess.CalledProcessError as e:
        print(f"❌ Erro no build: {e}")
        print("📋 Output do erro:")
        print(e.stdout)
        print(e.stderr)
        return False


# =============================================================================
# 🎯 **5. VERSÃO ALTERNATIVA - COM ÍCONE EXTRAÍDO**
# =============================================================================


def extract_system_icon():
    """Extrai ícone de computador do Windows para usar no EXE"""
    try:
        import win32api
        import win32con
        import win32gui

        # Obter handle do ícone padrão de computador
        icon_handle = win32gui.LoadIcon(0, win32con.IDI_APPLICATION)

        # Salvar como arquivo .ico (necessário para PyInstaller)
        # Método simples: usar shell32.dll
        shell32_path = os.path.join(os.environ["WINDIR"], "System32", "shell32.dll")

        # Extrair ícone índice 15 (computador) da shell32.dll
        cmd_extract = [
            "powershell",
            "-Command",
            f"""
            [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') | Out-Null
            $icon = [System.Drawing.Icon]::ExtractAssociatedIcon('{shell32_path}')
            $stream = [System.IO.File]::Create('computer_icon.ico')
            $icon.Save($stream)
            $stream.Close()
            """,
        ]

        subprocess.run(cmd_extract, check=True)
        print("✅ Ícone extraído: computer_icon.ico")
        return "computer_icon.ico"

    except Exception as e:
        print(f"⚠️ Erro extraindo ícone: {e}")
        return None


# =============================================================================
# 🎯 **6. BUILD COM ÍCONE PERSONALIZADO**
# =============================================================================


def build_with_custom_icon(python_exe=None):
    """Build com ícone extraído do sistema"""

    if python_exe is None:
        python_exe = sys.executable

    print("🎨 Extraindo ícone do sistema...")
    icon_file = extract_system_icon()

    # Comando com ícone personalizado
    cmd = [
        python_exe,
        "-m",
        "PyInstaller",
        "--onefile",
        "--windowed",
        "--optimize=2",
        "--strip",
        # Todas as exclusões da versão anterior...
        "--exclude-module=tkinter",
        "--exclude-module=matplotlib",
        # ... (todas as outras exclusões)
        "--name=KeepAliveRDP",
        "--distpath=dist_with_icon",
    ]

    # Adicionar ícone se foi extraído com sucesso
    if icon_file and os.path.exists(icon_file):
        cmd.append(f"--icon={icon_file}")
        print(f"🎨 Usando ícone: {icon_file}")

    cmd.append("keep-alive-app.py")

    try:
        subprocess.run(cmd, check=True)
        print("✅ Build com ícone concluído!")

        exe_path = Path("dist_with_icon/KeepAliveRDP.exe")
        if exe_path.exists():
            size_mb = exe_path.stat().st_size / (1024 * 1024)
            print(f"📦 Tamanho: {size_mb:.1f} MB")

    except Exception as e:
        print(f"❌ Erro: {e}")


# =============================================================================
# 🚀 **7. EXECUÇÃO PRINCIPAL**
# =============================================================================

if __name__ == "__main__":
    print("🔥 KeepAlive RDP - Gerador de EXE Otimizado")
    print("=" * 50)

    choice = input(
        """
Escolha o tipo de build:
1 - Build básico otimizado (sem ícone personalizado)
2 - Build com ícone extraído do sistema
3 - Ambos
Digite sua escolha (1-3): """
    ).strip()

    if choice == "1":
        build_optimized_exe()
    elif choice == "2":
        build_with_custom_icon()
    elif choice == "3":
        print("\n🔥 Executando build básico...")
        build_optimized_exe()
        print("\n🎨 Executando build com ícone...")
        build_with_custom_icon()
    else:
        print("❌ Opção inválida!")

    print("\n✅ Processo concluído!")

# =============================================================================
# 📋 **8. COMANDOS MANUAIS ALTERNATIVOS**
# =============================================================================

"""
🔧 **COMANDO MANUAL BÁSICO:**
pyinstaller --onefile --windowed --optimize=2 --strip --name=KeepAliveRDP keep-alive-app.py

🔧 **COMANDO MANUAL SUPER OTIMIZADO:**
pyinstaller --onefile --windowed --optimize=2 --strip --exclude-module=tkinter --exclude-module=matplotlib --exclude-module=numpy --exclude-module=pandas --exclude-module=PyQt6.QtWebEngine --exclude-module=PyQt6.QtMultimedia --name=KeepAliveRDP keep-alive-app.py3

🎯 **TAMANHO ESPERADO:**
- Build básico: ~25-35 MB
- Build otimizado: ~15-25 MB  
- Com UPX (opcional): ~8-15 MB

⚡ **TEMPO DE CARREGAMENTO:**
- SSD: ~2-4 segundos
- HDD: ~3-6 segundos

🎨 **ÍCONE:**
- O programa usa QStyle.StandardPixmap.SP_ComputerIcon
- É o mesmo ícone que aparece na bandeja
- Ícone padrão do Windows (computador)
"""

# 🚀 **EXECUTE:** python build_keepalive.py
