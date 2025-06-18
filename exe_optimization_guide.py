# =============================================================================
# 🔥 GERAÇÃO DE EXE OTIMIZADO - KeepAlive RDP Connection
# Versão Multi-Máquina Compatível
# =============================================================================

import os
import platform
import shutil
import subprocess
import sys
from pathlib import Path


def detect_python_versions():
    """Detecta versões do Python disponíveis no sistema"""
    versions = []
    system = platform.system().lower()

    # Verificar Python atual
    try:
        result = subprocess.run(
            [sys.executable, "--version"], capture_output=True, text=True, timeout=10
        )
        if result.returncode == 0:
            version = result.stdout.strip().replace("Python ", "")
            versions.append((sys.executable, version, "atual"))
    except:
        pass

    # Caminhos específicos por sistema operacional
    if system == "windows":
        # Windows - verificar caminhos comuns
        base_paths = [
            r"C:\Python{}\python.exe",
            r"C:\Users\{}\AppData\Local\Programs\Python\Python{}\python.exe",
            r"C:\Program Files\Python{}\python.exe",
            r"C:\Program Files (x86)\Python{}\python.exe",
        ]

        # Versões para testar
        python_versions = ["39", "310", "311", "312", "313"]

        for version in python_versions:
            for base_path in base_paths:
                if "{}" in base_path:
                    # Para caminhos com usuário
                    if "Users" in base_path:
                        try:
                            username = os.getenv("USERNAME", "")
                            if username:
                                path = base_path.format(username, version)
                            else:
                                continue
                        except:
                            continue
                    else:
                        path = base_path.format(version)
                else:
                    path = base_path

                if os.path.exists(path):
                    try:
                        result = subprocess.run(
                            [path, "--version"],
                            capture_output=True,
                            text=True,
                            timeout=5,
                        )
                        if result.returncode == 0:
                            ver = result.stdout.strip().replace("Python ", "")
                            if not any(v[0] == path for v in versions):
                                versions.append((path, ver, ""))
                    except:
                        continue

    # Comandos genéricos para qualquer sistema
    generic_commands = [
        "python",
        "python3",
        "py",
        "python.exe",
        "python3.9",
        "python3.10",
        "python3.11",
        "python3.12",
        "python3.13",
    ]

    for cmd in generic_commands:
        try:
            result = subprocess.run(
                [cmd, "--version"], capture_output=True, text=True, timeout=5
            )
            if result.returncode == 0:
                version = result.stdout.strip().replace("Python ", "")
                # Verificar se já não foi adicionado
                if not any(v[1] == version and cmd in v[0] for v in versions):
                    # Obter caminho completo
                    try:
                        which_result = subprocess.run(
                            ["where" if system == "windows" else "which", cmd],
                            capture_output=True,
                            text=True,
                            timeout=5,
                        )
                        if which_result.returncode == 0:
                            full_path = which_result.stdout.strip().split("\n")[0]
                            versions.append((full_path, version, ""))
                    except:
                        versions.append((cmd, version, "comando"))
        except:
            continue

    return versions


def select_python_version():
    """Interface para seleção da versão do Python"""
    versions = detect_python_versions()

    if not versions:
        print("❌ Nenhuma versão do Python encontrada!")
        print("💡 Certifique-se de que o Python está instalado e no PATH")
        return None

    print("\n🐍 Versões do Python encontradas:")
    print("-" * 60)

    for i, (path, version, note) in enumerate(versions):
        note_str = f" ({note})" if note else ""
        print(f"{i+1:2d}. Python {version}{note_str}")
        print(f"     📁 {path}")

    print("-" * 60)

    while True:
        try:
            choice = input(f"\n🎯 Escolha a versão (1-{len(versions)}) [1]: ").strip()

            # Default para primeira opção
            if not choice:
                choice = "1"

            choice_idx = int(choice) - 1

            if 0 <= choice_idx < len(versions):
                selected = versions[choice_idx]
                print(f"✅ Selecionado: Python {selected[1]}")
                print(f"📁 Caminho: {selected[0]}")
                return selected[0]
            else:
                print(f"❌ Opção inválida! Digite um número entre 1 e {len(versions)}")
        except ValueError:
            print("❌ Digite um número válido!")
        except KeyboardInterrupt:
            print("\n🔴 Operação cancelada pelo usuário")
            return None


def fix_pyinstaller_conflicts(python_exe):
    """Corrige conflitos conhecidos do PyInstaller"""
    print("\n🔧 Verificando conflitos do PyInstaller...")

    # Lista de pacotes problemáticos básicos
    problematic_packages = [
        "typing",  # Backport obsoleto
        "enum34",  # Enum backport obsoleto
        "pathlib2",  # Pathlib backport obsoleto
        "futures",  # Concurrent.futures backport obsoleto
        "importlib-metadata",  # Pode causar conflitos
    ]

    fixed_any = False

    for package in problematic_packages:
        try:
            # Verificar se está instalado
            check_cmd = [python_exe, "-c", f"import {package.replace('-', '_')}"]
            result = subprocess.run(
                check_cmd, capture_output=True, text=True, timeout=5
            )

            if result.returncode == 0:
                print(f"🗑️ Removendo pacote problemático: {package}")
                uninstall_cmd = [python_exe, "-m", "pip", "uninstall", package, "-y"]
                result = subprocess.run(
                    uninstall_cmd, capture_output=True, text=True, timeout=30
                )

                if result.returncode == 0:
                    print(f"✅ {package} removido com sucesso")
                    fixed_any = True
                else:
                    print(
                        f"⚠️ Falha ao remover {package}: {result.stderr.strip()[:50]}..."
                    )
        except:
            # Pacote não instalado, tudo ok
            continue

    # 🔥 NOVO: Verificar conflito Qt (PyQt5 vs PyQt6)
    qt_conflict = check_qt_conflict(python_exe)
    if qt_conflict:
        fixed_any = True

    if fixed_any:
        print("🔄 Conflitos corrigidos! PyInstaller deve funcionar agora.")
    else:
        print("✅ Nenhum conflito encontrado.")


def check_qt_conflict(python_exe):
    """Verifica e resolve conflito entre PyQt5 e PyQt6"""
    print("\n🎨 Verificando conflito Qt...")

    has_pyqt5 = False
    has_pyqt6 = False
    has_pyside6 = False

    # Verificar PyQt5
    try:
        result = subprocess.run(
            [python_exe, "-c", "import PyQt5; print('OK')"],
            capture_output=True,
            text=True,
            timeout=5,
        )
        if result.returncode == 0:
            has_pyqt5 = True
            print("📦 PyQt5 encontrado")
    except:
        pass

    # Verificar PyQt6
    try:
        result = subprocess.run(
            [python_exe, "-c", "import PyQt6; print('OK')"],
            capture_output=True,
            text=True,
            timeout=5,
        )
        if result.returncode == 0:
            has_pyqt6 = True
            print("📦 PyQt6 encontrado")
    except:
        pass

    # Verificar PySide6
    try:
        result = subprocess.run(
            [python_exe, "-c", "import PySide6; print('OK')"],
            capture_output=True,
            text=True,
            timeout=5,
        )
        if result.returncode == 0:
            has_pyside6 = True
            print("📦 PySide6 encontrado")
    except:
        pass

    # Se há conflito, resolver
    if (
        (has_pyqt5 and has_pyqt6)
        or (has_pyqt5 and has_pyside6)
        or (has_pyqt6 and has_pyside6)
    ):
        print("⚠️ CONFLITO: Múltiplas bibliotecas Qt encontradas!")
        print("🎯 O PyInstaller não suporta múltiplas bibliotecas Qt.")

        # Priorizar PyQt6 (mais moderno)
        if has_pyqt6:
            print("🔧 Mantendo PyQt6 e removendo outras...")
            if has_pyqt5:
                remove_package(python_exe, "PyQt5")
            if has_pyside6:
                remove_package(python_exe, "PySide6")
        elif has_pyside6:
            print("🔧 Mantendo PySide6 e removendo PyQt5...")
            if has_pyqt5:
                remove_package(python_exe, "PyQt5")
        else:
            print("🔧 Mantendo PyQt5...")

        return True
    else:
        print("✅ Nenhum conflito Qt encontrado")
        return False


def remove_package(python_exe, package):
    """Remove um pacote específico"""
    try:
        print(f"🗑️ Removendo {package}...")
        cmd = [python_exe, "-m", "pip", "uninstall", package, "-y"]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)

        if result.returncode == 0:
            print(f"✅ {package} removido com sucesso")
        else:
            print(f"⚠️ Falha ao remover {package}: {result.stderr.strip()[:100]}...")
    except Exception as e:
        print(f"❌ Erro removendo {package}: {e}")


def check_and_install_dependencies(python_exe):
    """Verifica e instala dependências necessárias"""
    print("\n📦 Verificando dependências...")

    # Primeiro, corrigir conflitos conhecidos
    fix_pyinstaller_conflicts(python_exe)

    dependencies = {
        "pyinstaller": ">=6.0.0",
        "PyQt6": ">=6.0.0",
        "pyautogui": ">=0.9.50",
    }

    # Adicionar pywin32 apenas no Windows
    if platform.system().lower() == "windows":
        dependencies["pywin32"] = ">=300"

    print(f"\n🔍 Usando Python: {python_exe}")

    for package, version in dependencies.items():
        print(f"\n📦 Verificando {package}{version}...")

        # Verificar se já está instalado
        try:
            import_name = package.lower().replace("-", "_")
            if package == "PyQt6":
                import_name = "PyQt6.QtCore"

            check_cmd = [python_exe, "-c", f"import {import_name}; print('OK')"]
            result = subprocess.run(
                check_cmd, capture_output=True, text=True, timeout=10
            )

            if result.returncode == 0 and "OK" in result.stdout:
                print(f"✅ {package} já instalado")
                continue
        except:
            pass

        # Instalar se necessário
        print(f"📥 Instalando {package}{version}...")
        try:
            install_cmd = [python_exe, "-m", "pip", "install", f"{package}{version}"]
            result = subprocess.run(
                install_cmd, capture_output=True, text=True, timeout=120
            )

            if result.returncode == 0:
                print(f"✅ {package} instalado com sucesso")
            else:
                print(f"⚠️ Aviso ao instalar {package}:")
                print(f"   {result.stderr.strip()[:100]}...")

        except subprocess.TimeoutExpired:
            print(f"⏱️ Timeout ao instalar {package} - pode ter sido instalado")
        except Exception as e:
            print(f"❌ Erro instalando {package}: {str(e)[:100]}...")

    print("\n✅ Verificação de dependências concluída!")


def find_main_file():
    """Encontra o arquivo principal do KeepAlive"""
    possible_names = [
        "keep-alive-app.py",
        "keepalive-app.py",
        "keepalive.py",
        "keep_alive.py",
        "main.py",
        "app.py",
    ]

    current_dir = Path(".")

    for name in possible_names:
        if (current_dir / name).exists():
            return name

    # Listar arquivos .py disponíveis
    py_files = list(current_dir.glob("*.py"))
    if py_files:
        print("📁 Arquivos .py encontrados:")
        for i, file in enumerate(py_files):
            print(f"   {i+1}. {file.name}")

        while True:
            try:
                choice = input(
                    f"\nEscolha o arquivo principal (1-{len(py_files)}): "
                ).strip()
                if choice:
                    idx = int(choice) - 1
                    if 0 <= idx < len(py_files):
                        return py_files[idx].name
                print("❌ Opção inválida!")
            except (ValueError, KeyboardInterrupt):
                return None

    return None


def test_pyinstaller(python_exe):
    """Testa se PyInstaller está funcionando corretamente"""
    print("\n🧪 Testando PyInstaller...")

    try:
        test_cmd = [python_exe, "-m", "PyInstaller", "--version"]
        result = subprocess.run(test_cmd, capture_output=True, text=True, timeout=10)

        if result.returncode == 0:
            version = result.stdout.strip()
            print(f"✅ PyInstaller funcional - Versão: {version}")
            return True
        else:
            print(f"❌ PyInstaller com problemas: {result.stderr}")
            return False
    except Exception as e:
        print(f"❌ Erro testando PyInstaller: {e}")
        return False


def build_optimized_exe(python_exe=None):
    """Gera EXE otimizado com configurações agressivas de redução de tamanho"""

    if python_exe is None:
        python_exe = sys.executable

    print(f"\n🔥 Iniciando build otimizado")
    print(f"🐍 Python: {python_exe}")

    # Testar PyInstaller primeiro
    if not test_pyinstaller(python_exe):
        print("❌ PyInstaller não está funcionando. Tentando corrigir...")
        fix_pyinstaller_conflicts(python_exe)

        # Testar novamente
        if not test_pyinstaller(python_exe):
            print("❌ Não foi possível corrigir problemas do PyInstaller")
            return False

    # Encontrar arquivo principal
    main_file = find_main_file()
    if not main_file:
        print("❌ Arquivo principal não encontrado!")
        return False

    print(f"📄 Arquivo principal: {main_file}")

    # Preparar diretórios
    output_dir = Path("dist_keepalive")
    build_dir = Path("build_temp")
    spec_dir = Path("specs")

    # Limpar diretórios antigos
    for dir_path in [output_dir, build_dir, spec_dir]:
        if dir_path.exists():
            try:
                shutil.rmtree(dir_path)
                print(f"🧹 Removido: {dir_path}")
            except:
                pass

    # Comando PyInstaller ULTRA OTIMIZADO com exclusões Qt específicas
    cmd = [
        python_exe,
        "-m",
        "PyInstaller",
        "--onefile",  # Arquivo único
        "--windowed",  # Sem console
        "--optimize=2",  # Otimização máxima
        "--strip",  # Remove símbolos debug
        "--noupx",  # Evita problemas com antivírus
        "--clean",  # Limpa cache
        # 🚫 EXCLUSÕES ESPECÍFICAS PARA CONFLITO Qt
        "--exclude-module=PyQt5",  # Força exclusão PyQt5
        "--exclude-module=PySide6",  # Força exclusão PySide6
        "--exclude-module=qtpy",  # Exclui qtpy que causa detecção múltipla
        # 🎯 EXCLUSÕES AGRESSIVAS - Módulos desnecessários
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
        # 🚫 PyQt6 - Módulos pesados desnecessários
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
        "--exclude-module=PyQt6.QtPrintSupport",
        "--exclude-module=PyQt6.QtQml",
        "--exclude-module=PyQt6.QtQuick",
        "--exclude-module=PyQt6.QtSvg",
        "--exclude-module=PyQt6.QtSvgWidgets",
        "--exclude-module=PyQt6.Qt3D",
        "--exclude-module=PyQt6.QtCharts",
        "--exclude-module=PyQt6.QtDataVisualization",
        # 📂 Configurações de saída
        f"--name=KeepAliveRDP",
        f"--distpath={output_dir}",
        f"--workpath={build_dir}",
        f"--specpath={spec_dir}",
        # 📄 Arquivo principal
        main_file,
    ]

    # 🔥 Adicionar variável de ambiente para forçar PyQt6
    env = os.environ.copy()
    env["QT_API"] = "pyqt6"

    print("\n⚡ Executando PyInstaller...")
    print("🎯 Forçando uso exclusivo do PyQt6...")
    print("⏱️ Isso pode levar alguns minutos...")

    try:
        # Executar com variável de ambiente específica
        result = subprocess.run(
            cmd, capture_output=True, text=True, timeout=300, env=env
        )

        # Mostrar output relevante
        if result.stdout:
            lines = result.stdout.split("\n")
            for line in lines:
                if any(
                    keyword in line
                    for keyword in [
                        "INFO:",
                        "WARNING:",
                        "ERROR:",
                        "Building",
                        "Analyzing",
                    ]
                ):
                    print(f"  {line}")

        if result.stderr:
            print("🔍 Erros/Avisos:")
            print(result.stderr)

        if result.returncode == 0:
            print("✅ Build concluído com sucesso!")

            # Verificar arquivo gerado
            exe_path = output_dir / "KeepAliveRDP.exe"
            if exe_path.exists():
                size_mb = exe_path.stat().st_size / (1024 * 1024)
                print(f"\n📦 Arquivo gerado:")
                print(f"   📁 Local: {exe_path.absolute()}")
                print(f"   📏 Tamanho: {size_mb:.1f} MB")

                # Limpar arquivos temporários
                cleanup_build_files(build_dir, spec_dir)

                print(f"\n🎉 EXE pronto para uso!")
                return True
            else:
                print("❌ Arquivo EXE não foi encontrado!")
                return False
        else:
            print(f"❌ Build falhou com código: {result.returncode}")
            print("📋 Detalhes do erro:")
            if result.stderr:
                print(result.stderr)

            # Verificar se é erro Qt específico
            if "multiple Qt bindings" in result.stderr:
                print("\n🔧 SOLUÇÃO: Conflito Qt detectado!")
                print("Execute novamente - o script vai corrigir automaticamente.")

            return False

    except subprocess.TimeoutExpired:
        print("⏰ Build interrompido por timeout (5 minutos)")
        return False
    except Exception as e:
        print(f"❌ Erro durante build: {e}")
        return False


def extract_system_icon():
    """Extrai ícone do sistema Windows para usar no EXE"""
    try:
        print("🎨 Tentando extrair ícone do sistema...")

        # Método 1: Usar PowerShell para extrair ícone
        powershell_cmd = [
            "powershell",
            "-Command",
            """
            Add-Type -AssemblyName System.Drawing
            $icon = [System.Drawing.SystemIcons]::Computer
            $bitmap = $icon.ToBitmap()
            $bitmap.Save('computer_icon.png', [System.Drawing.Imaging.ImageFormat]::Png)
            Write-Host 'Icon extracted successfully'
            """,
        ]

        result = subprocess.run(
            powershell_cmd, capture_output=True, text=True, timeout=30
        )

        if result.returncode == 0 and os.path.exists("computer_icon.png"):
            print("✅ Ícone extraído: computer_icon.png")
            return "computer_icon.png"

        # Método 2: Criar ícone simples baseado em texto
        print("🎨 Criando ícone simples...")
        create_simple_icon()
        if os.path.exists("simple_icon.ico"):
            return "simple_icon.ico"

    except Exception as e:
        print(f"⚠️ Erro extraindo ícone: {e}")

    return None


def create_simple_icon():
    """Cria um ícone simples usando PIL se disponível"""
    try:
        from PIL import Image, ImageDraw, ImageFont

        # Criar imagem 32x32
        img = Image.new("RGBA", (32, 32), (0, 100, 200, 255))
        draw = ImageDraw.Draw(img)

        # Desenhar forma simples
        draw.rectangle([4, 4, 28, 28], fill=(255, 255, 255, 255))
        draw.rectangle([8, 8, 24, 24], fill=(0, 100, 200, 255))

        # Salvar como ICO
        img.save("simple_icon.ico", format="ICO")
        print("✅ Ícone simples criado")

    except ImportError:
        print("⚠️ PIL não disponível para criar ícone")
    except Exception as e:
        print(f"⚠️ Erro criando ícone: {e}")


def build_with_icon(python_exe):
    """Build com tentativa de ícone personalizado"""
    print("\n🎨 Tentando build com ícone personalizado...")

    # Tentar extrair ícone
    icon_file = extract_system_icon()

    # Se não conseguiu, fazer build sem ícone mesmo
    if not icon_file:
        print("🔄 Fazendo build sem ícone personalizado...")
        return build_optimized_exe(python_exe)

    # Build com ícone
    print(f"🎨 Usando ícone: {icon_file}")

    # Encontrar arquivo principal
    main_file = find_main_file()
    if not main_file:
        print("❌ Arquivo principal não encontrado!")
        return False

    # Preparar diretórios
    output_dir = Path("dist_with_icon")
    build_dir = Path("build_icon")
    spec_dir = Path("specs_icon")

    # Limpar diretórios antigos
    for dir_path in [output_dir, build_dir, spec_dir]:
        if dir_path.exists():
            try:
                shutil.rmtree(dir_path)
            except:
                pass

    # Comando básico com ícone
    cmd = [
        python_exe,
        "-m",
        "PyInstaller",
        "--onefile",
        "--windowed",
        "--optimize=2",
        "--strip",
        "--clean",
        f"--icon={icon_file}",
        f"--name=KeepAliveRDP_Icon",
        f"--distpath={output_dir}",
        f"--workpath={build_dir}",
        f"--specpath={spec_dir}",
        main_file,
    ]

    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)

        if result.returncode == 0:
            exe_path = output_dir / "KeepAliveRDP_Icon.exe"
            if exe_path.exists():
                size_mb = exe_path.stat().st_size / (1024 * 1024)
                print(f"✅ Build com ícone concluído!")
                print(f"📁 Local: {exe_path.absolute()}")
                print(f"📏 Tamanho: {size_mb:.1f} MB")

                cleanup_build_files(build_dir, spec_dir)
                return True

        print("❌ Build com ícone falhou")
        if result.stderr:
            print(f"Erro: {result.stderr[:200]}...")

    except Exception as e:
        print(f"❌ Erro no build com ícone: {e}")

    return False


def show_system_info():
    """Mostra informações do sistema"""
    print("💻 Informações do Sistema:")
    print(f"   🖥️ OS: {platform.system()} {platform.release()}")
    print(f"   🏗️ Arquitetura: {platform.machine()}")
    print(f"   🐍 Python atual: {sys.version.split()[0]}")


def cleanup_build_files(build_dir, spec_dir):
    """Remove arquivos temporários de build"""
    try:
        if build_dir.exists():
            shutil.rmtree(build_dir)
            print("🧹 Arquivos temporários removidos")

        if spec_dir.exists():
            for spec_file in spec_dir.glob("*.spec"):
                spec_file.unlink()
            if not list(spec_dir.iterdir()):
                spec_dir.rmdir()
    except Exception as e:
        print(f"⚠️ Erro limpando arquivos temporários: {e}")


def show_build_options():
    """Mostra opções de build disponíveis"""
    print("\n🔥 Opções de Build Disponíveis:")
    print("=" * 40)
    print("1 - Build básico otimizado (recomendado)")
    print("2 - Build com ícone personalizado")
    print("3 - Ambos (básico + com ícone)")
    print("4 - Build de teste rápido")
    print("=" * 40)


def build_quick_test(python_exe):
    """Build rápido para teste (sem otimizações pesadas)"""
    print("\n⚡ Build de teste rápido...")

    main_file = find_main_file()
    if not main_file:
        print("❌ Arquivo principal não encontrado!")
        return False

    # Build simples e rápido
    cmd = [
        python_exe,
        "-m",
        "PyInstaller",
        "--onefile",
        "--windowed",
        "--clean",
        "--name=KeepAliveRDP_Test",
        "--distpath=dist_test",
        main_file,
    ]

    try:
        print("⚡ Executando build de teste...")
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=180)

        if result.returncode == 0:
            exe_path = Path("dist_test/KeepAliveRDP_Test.exe")
            if exe_path.exists():
                size_mb = exe_path.stat().st_size / (1024 * 1024)
                print(f"✅ Build de teste concluído!")
                print(f"📁 Local: {exe_path.absolute()}")
                print(f"📏 Tamanho: {size_mb:.1f} MB")
                return True

        print("❌ Build de teste falhou")
        if result.stderr:
            print(f"Erro: {result.stderr[:200]}...")

    except Exception as e:
        print(f"❌ Erro no build de teste: {e}")

    return False


def main():
    """Função principal com opções de build"""
    print("🔥 KeepAlive RDP - Gerador de EXE Otimizado")
    print("=" * 55)

    show_system_info()

    # Selecionar versão do Python
    python_exe = select_python_version()
    if not python_exe:
        print("❌ Não foi possível selecionar uma versão do Python")
        return

    # Sempre verificar dependências e conflitos
    print("\n🔧 Verificando ambiente Python...")
    check_and_install_dependencies(python_exe)

    # Mostrar opções de build
    show_build_options()

    while True:
        try:
            choice = input("\n🎯 Escolha o tipo de build (1-4) [1]: ").strip()

            if not choice:
                choice = "1"

            if choice == "1":
                print("\n🔥 Executando build básico otimizado...")
                success = build_optimized_exe(python_exe)
                break

            elif choice == "2":
                print("\n🎨 Executando build com ícone...")
                success = build_with_icon(python_exe)
                break

            elif choice == "3":
                print("\n🔥 Executando build básico...")
                success1 = build_optimized_exe(python_exe)

                print("\n🎨 Executando build com ícone...")
                success2 = build_with_icon(python_exe)

                success = success1 or success2

                if success1 and success2:
                    print("\n🎉 Ambos builds concluídos com sucesso!")
                elif success1:
                    print("\n✅ Build básico concluído. Build com ícone falhou.")
                elif success2:
                    print("\n✅ Build com ícone concluído. Build básico falhou.")
                else:
                    print("\n❌ Ambos builds falharam.")
                break

            elif choice == "4":
                print("\n⚡ Executando build de teste...")
                success = build_quick_test(python_exe)
                break

            else:
                print("❌ Opção inválida! Digite 1, 2, 3 ou 4")
                continue

        except KeyboardInterrupt:
            print("\n🔴 Operação cancelada pelo usuário")
            return

    # Resultado final
    if success:
        print("\n🎉 Processo concluído com sucesso!")
        print("\n📂 Verifique as pastas de saída:")
        print("   - dist_keepalive/ (build básico)")
        print("   - dist_with_icon/ (build com ícone)")
        print("   - dist_test/ (build de teste)")
    else:
        print("\n❌ Build falhou.")
        print("\n💡 Dicas para resolver problemas:")
        print("   1. Tente o build de teste (opção 4) primeiro")
        print("   2. Execute novamente - o script corrigiu conflitos")
        print("   3. Verifique se o arquivo Python principal existe")
        print("   4. Use uma versão mais recente do Python (3.9+)")

        # Oferecer tentar novamente
        retry = input("\n🔄 Tentar build de teste? [S/n]: ").strip().lower()
        if retry in ("", "s", "sim", "y", "yes"):
            success = build_quick_test(python_exe)
            if success:
                print(
                    "\n🎉 Build de teste funcionou! Problema pode ser nas otimizações."
                )


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n🔴 Operação cancelada pelo usuário")
    except Exception as e:
        print(f"\n❌ Erro inesperado: {e}")

    input("\n⏸️ Pressione Enter para sair...")

# =============================================================================
# 📋 COMANDOS MANUAIS PARA REFERÊNCIA
# =============================================================================

"""
🔧 COMANDO MANUAL BÁSICO:
pyinstaller --onefile --windowed --optimize=2 --strip --name=KeepAliveRDP keep-alive-app.py

🔧 COMANDO MANUAL COMPLETO:
pyinstaller --onefile --windowed --optimize=2 --strip \
  --exclude-module=tkinter --exclude-module=matplotlib \
  --exclude-module=PyQt6.QtWebEngine --exclude-module=PyQt6.QtMultimedia \
  --name=KeepAliveRDP keep-alive-app.py

🎯 TAMANHO ESPERADO:
- Build básico: ~25-35 MB
- Build otimizado: ~15-25 MB
- Com todas exclusões: ~10-20 MB

⚡ COMPATIBILIDADE:
- Windows 7/8/10/11
- Não requer instalação do Python na máquina destino
- Executável independente (standalone)

🚀 USO:
1. Coloque este arquivo na mesma pasta do seu script Python
2. Execute: python build_keepalive_exe.py
3. Siga as instruções na tela
4. EXE será gerado em dist_keepalive/KeepAliveRDP.exe
"""
