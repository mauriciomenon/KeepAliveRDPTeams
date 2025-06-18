def build_optimized_exe(python_exe=None):
    """Gerador inteligente de EXE com tratamento de erros adaptativo"""
    
    if python_exe is None:
        python_exe = sys.executable
    
    print(f"\n🔥 Iniciando build otimizado")
    print(f"🐍 Python: {python_exe}")
    
    # Testar PyInstaller primeiro
    if not test_pyinstaller(python_exe):
        print("❌ PyInstaller não está funcionando. Tentando correção inteligente...")
        smart_conflict_resolver(python_exe)
        
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
    for dir_path in [output_dir, build_dir, spec_dir]:# =============================================================================
# 🔥 GERAÇÃO DE EXE OTIMIZADO - KeepAlive RDP Connection
# Versão Multi-Máquina Compatível
# =============================================================================

import os
import subprocess
import sys
import shutil
import platform
from pathlib import Path

def detect_python_versions():
    """Detecta versões do Python disponíveis no sistema"""
    versions = []
    system = platform.system().lower()
    
    # Verificar Python atual
    try:
        result = subprocess.run([sys.executable, "--version"], 
                              capture_output=True, text=True, timeout=10)
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
            r"C:\Program Files (x86)\Python{}\python.exe"
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
                        result = subprocess.run([path, "--version"], 
                                              capture_output=True, text=True, timeout=5)
                        if result.returncode == 0:
                            ver = result.stdout.strip().replace("Python ", "")
                            if not any(v[0] == path for v in versions):
                                versions.append((path, ver, ""))
                    except:
                        continue
    
    # Comandos genéricos para qualquer sistema
    generic_commands = [
        "python", "python3", "py", "python.exe",
        "python3.9", "python3.10", "python3.11", "python3.12", "python3.13"
    ]
    
    for cmd in generic_commands:
        try:
            result = subprocess.run([cmd, "--version"], 
                                  capture_output=True, text=True, timeout=5)
            if result.returncode == 0:
                version = result.stdout.strip().replace("Python ", "")
                # Verificar se já não foi adicionado
                if not any(v[1] == version and cmd in v[0] for v in versions):
                    # Obter caminho completo
                    try:
                        which_result = subprocess.run(
                            ["where" if system == "windows" else "which", cmd],
                            capture_output=True, text=True, timeout=5
                        )
                        if which_result.returncode == 0:
                            full_path = which_result.stdout.strip().split('\n')[0]
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

# =============================================================================
# 🔥 GERAÇÃO DE EXE OTIMIZADO - KeepAlive RDP Connection
# Versão Inteligente com Memória de Problemas Resolvidos
# =============================================================================

import os
import subprocess
import sys
import shutil
import platform
import json
from pathlib import Path
from datetime import datetime

# 🧠 Sistema de memória para problemas resolvidos
MEMORY_FILE = ".pyinstaller_memory.json"

def load_memory():
    """Carrega memória de problemas já resolvidos"""
    try:
        if os.path.exists(MEMORY_FILE):
            with open(MEMORY_FILE, 'r') as f:
                return json.load(f)
    except:
        pass
    
    return {
        "resolved_conflicts": [],
        "last_successful_build": None,
        "known_issues": {},
        "user_preferences": {}
    }

def save_memory(memory):
    """Salva memória de problemas resolvidos"""
    try:
        memory["last_update"] = datetime.now().isoformat()
        with open(MEMORY_FILE, 'w') as f:
            json.dump(memory, f, indent=2)
    except:
        pass

def detect_python_versions():
    """Detecta versões do Python disponíveis no sistema"""
    versions = []
    system = platform.system().lower()
    
    # Verificar Python atual
    try:
        result = subprocess.run([sys.executable, "--version"], 
                              capture_output=True, text=True, timeout=10)
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
            r"C:\Program Files (x86)\Python{}\python.exe"
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
                        result = subprocess.run([path, "--version"], 
                                              capture_output=True, text=True, timeout=5)
                        if result.returncode == 0:
                            ver = result.stdout.strip().replace("Python ", "")
                            if not any(v[0] == path for v in versions):
                                versions.append((path, ver, ""))
                    except:
                        continue
    
    # Comandos genéricos para qualquer sistema
    generic_commands = [
        "python", "python3", "py", "python.exe",
        "python3.9", "python3.10", "python3.11", "python3.12", "python3.13"
    ]
    
    for cmd in generic_commands:
        try:
            result = subprocess.run([cmd, "--version"], 
                                  capture_output=True, text=True, timeout=5)
            if result.returncode == 0:
                version = result.stdout.strip().replace("Python ", "")
                # Verificar se já não foi adicionado
                if not any(v[1] == version and cmd in v[0] for v in versions):
                    # Obter caminho completo
                    try:
                        which_result = subprocess.run(
                            ["where" if system == "windows" else "which", cmd],
                            capture_output=True, text=True, timeout=5
                        )
                        if which_result.returncode == 0:
                            full_path = which_result.stdout.strip().split('\n')[0]
                            versions.append((full_path, version, ""))
                    except:
                        versions.append((cmd, version, "comando"))
        except:
            continue
    
    return versions

def select_python_version():
    """Interface inteligente para seleção da versão do Python"""
    memory = load_memory()
    versions = detect_python_versions()
    
    if not versions:
        print("❌ Nenhuma versão do Python encontrada!")
        print("💡 Certifique-se de que o Python está instalado e no PATH")
        return None
    
    # Verificar se há preferência do usuário
    last_successful = memory.get("last_successful_build")
    if last_successful and last_successful.get("python_exe"):
        last_python = last_successful["python_exe"]
        # Verificar se ainda existe
        if any(v[0] == last_python for v in versions):
            use_last = input(f"\n🎯 Usar Python da última execução bem-sucedida?\n   📁 {last_python}\n   [S/n]: ").strip().lower()
            if use_last in ('', 's', 'sim', 'y', 'yes'):
                print(f"✅ Usando Python anterior: {last_python}")
                return last_python
    
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

def smart_conflict_resolver(python_exe):
    """Sistema inteligente de resolução de conflitos"""
    memory = load_memory()
    resolved_conflicts = memory.get("resolved_conflicts", [])
    
    print("\n🧠 Sistema inteligente de resolução de conflitos ativo...")
    
    # Verificar conflitos já resolvidos
    current_conflicts = detect_all_conflicts(python_exe)
    new_conflicts = [c for c in current_conflicts if c not in resolved_conflicts]
    
    if not new_conflicts:
        if resolved_conflicts:
            print(f"✅ Todos os {len(resolved_conflicts)} conflitos conhecidos já foram resolvidos")
        else:
            print("✅ Nenhum conflito detectado")
        return
    
    print(f"🔍 Encontrados {len(new_conflicts)} novos conflitos:")
    for conflict in new_conflicts:
        print(f"   - {conflict}")
    
    # Resolver automaticamente conflitos conhecidos
    auto_resolved = []
    for conflict in new_conflicts:
        if auto_resolve_conflict(python_exe, conflict):
            auto_resolved.append(conflict)
            resolved_conflicts.append(conflict)
    
    if auto_resolved:
        print(f"🤖 Resolvidos automaticamente: {len(auto_resolved)} conflitos")
        memory["resolved_conflicts"] = resolved_conflicts
        save_memory(memory)
    
    # Para conflitos desconhecidos, pedir aprovação UMA VEZ
    unknown_conflicts = [c for c in new_conflicts if c not in auto_resolved]
    if unknown_conflicts:
        print(f"\n🆕 Conflitos novos detectados: {len(unknown_conflicts)}")
        if ask_user_once("Tentar resolver automaticamente?"):
            for conflict in unknown_conflicts:
                if smart_resolve_unknown(python_exe, conflict):
                    resolved_conflicts.append(conflict)
            
            memory["resolved_conflicts"] = resolved_conflicts
            save_memory(memory)

def detect_all_conflicts(python_exe):
    """Detecta todos os tipos de conflitos possíveis"""
    conflicts = []
    
    # Conflitos de pacotes obsoletos
    obsolete_packages = ["typing", "enum34", "pathlib2", "futures", "importlib-metadata"]
    for package in obsolete_packages:
        if is_package_installed(python_exe, package):
            conflicts.append(f"obsolete_package:{package}")
    
    # Conflitos Qt
    qt_libs = check_qt_libraries(python_exe)
    if len(qt_libs) > 1:
        conflicts.append(f"qt_conflict:{','.join(qt_libs)}")
    
    # Conflitos de hooks PyInstaller
    if has_hook_conflicts(python_exe):
        conflicts.append("pyinstaller_hooks")
    
    return conflicts

def is_package_installed(python_exe, package):
    """Verifica se um pacote está instalado"""
    try:
        import_name = package.replace('-', '_')
        check_cmd = [python_exe, "-c", f"import {import_name}"]
        result = subprocess.run(check_cmd, capture_output=True, text=True, timeout=5)
        return result.returncode == 0
    except:
        return False

def check_qt_libraries(python_exe):
    """Verifica quais bibliotecas Qt estão instaladas"""
    qt_libs = []
    
    for qt_lib in ["PyQt5", "PyQt6", "PySide6"]:
        if is_package_installed(python_exe, qt_lib):
            qt_libs.append(qt_lib)
    
    return qt_libs

def has_hook_conflicts(python_exe):
    """Verifica conflitos de hooks do PyInstaller"""
    try:
        # Verificar se PyInstaller funciona sem erro de hooks
        test_cmd = [python_exe, "-m", "PyInstaller", "--version"]
        result = subprocess.run(test_cmd, capture_output=True, text=True, timeout=10)
        
        if result.returncode == 0:
            # Verificar se há warnings de hooks no stderr
            return "hook_dirs" in result.stderr or "hooks_contrib" in result.stderr
        
        return True
    except:
        return True

def auto_resolve_conflict(python_exe, conflict):
    """Resolve automaticamente conflitos conhecidos"""
    if conflict.startswith("obsolete_package:"):
        package = conflict.split(":")[1]
        return remove_package_silent(python_exe, package)
    
    elif conflict.startswith("qt_conflict:"):
        qt_libs = conflict.split(":")[1].split(",")
        # Manter PyQt6, remover outros
        if "PyQt6" in qt_libs:
            for lib in qt_libs:
                if lib != "PyQt6":
                    remove_package_silent(python_exe, lib)
        return True
    
    elif conflict == "pyinstaller_hooks":
        return fix_pyinstaller_hooks(python_exe)
    
    return False

def smart_resolve_unknown(python_exe, conflict):
    """Tenta resolver conflitos desconhecidos de forma inteligente"""
    print(f"🔧 Tentando resolver: {conflict}")
    
    # Lógica básica para novos tipos de conflito
    if "package" in conflict.lower():
        package_name = conflict.split(":")[-1] if ":" in conflict else conflict
        return remove_package_silent(python_exe, package_name)
    
    return False

def remove_package_silent(python_exe, package):
    """Remove pacote silenciosamente"""
    try:
        cmd = [python_exe, "-m", "pip", "uninstall", package, "-y"]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
        return result.returncode == 0
    except:
        return False

def fix_pyinstaller_hooks(python_exe):
    """Corrige problemas de hooks do PyInstaller"""
    try:
        # Reinstalar PyInstaller para corrigir hooks
        cmd = [python_exe, "-m", "pip", "install", "--upgrade", "--force-reinstall", "pyinstaller"]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
        return result.returncode == 0
    except:
        return False

def ask_user_once(question):
    """Pergunta ao usuário apenas uma vez, lembra da resposta"""
    memory = load_memory()
    user_prefs = memory.get("user_preferences", {})
    
    # Gerar chave única para esta pergunta
    question_key = f"auto_resolve_{hash(question) % 10000}"
    
    if question_key in user_prefs:
        choice = user_prefs[question_key]
        print(f"🧠 Lembrado: {question} -> {'Sim' if choice else 'Não'}")
        return choice
    
    # Primeira vez perguntando
    response = input(f"\n❓ {question} [S/n]: ").strip().lower()
    choice = response in ('', 's', 'sim', 'y', 'yes')
    
    # Salvar resposta
    user_prefs[question_key] = choice
    memory["user_preferences"] = user_prefs
    save_memory(memory)
    
    return choice

def check_and_install_dependencies(python_exe):
    """Sistema inteligente de verificação de dependências"""
    memory = load_memory()
    
    print("\n📦 Verificando dependências...")
    
    # Usar sistema inteligente de resolução de conflitos
    smart_conflict_resolver(python_exe)
    
    dependencies = {
        "pyinstaller": ">=6.0.0",
        "PyQt6": ">=6.0.0", 
        "pyautogui": ">=0.9.50"
    }
    
    # Adicionar pywin32 apenas no Windows
    if platform.system().lower() == "windows":
        dependencies["pywin32"] = ">=300"
    
    print(f"\n🔍 Usando Python: {python_exe}")
    
    for package, version in dependencies.items():
        if not check_dependency_smart(python_exe, package, version):
            print(f"❌ Falha ao verificar/instalar {package}")

def check_dependency_smart(python_exe, package, version):
    """Verificação inteligente de dependência individual"""
    print(f"\n📦 Verificando {package}{version}...")
    
    # Verificar se já está instalado
    try:
        import_name = package.lower().replace('-', '_')
        if package == "PyQt6":
            import_name = "PyQt6.QtCore"
        
        check_cmd = [python_exe, "-c", f"import {import_name}; print('OK')"]
        result = subprocess.run(check_cmd, capture_output=True, text=True, timeout=10)
        
        if result.returncode == 0 and "OK" in result.stdout:
            print(f"✅ {package} já instalado")
            return True
    except:
        pass
    
    # Instalar se necessário
    print(f"📥 Instalando {package}{version}...")
    try:
        install_cmd = [python_exe, "-m", "pip", "install", f"{package}{version}"]
        result = subprocess.run(install_cmd, capture_output=True, text=True, timeout=120)
        
        if result.returncode == 0:
            print(f"✅ {package} instalado com sucesso")
            return True
        else:
            print(f"⚠️ Aviso ao instalar {package}:")
            print(f"   {result.stderr.strip()[:100]}...")
            return False
            
    except subprocess.TimeoutExpired:
        print(f"⏱️ Timeout ao instalar {package}")
        return False
    except Exception as e:
        print(f"❌ Erro instalando {package}: {str(e)[:100]}...")
        return False

def intelligent_build_error_handler(python_exe, error_output):
    """Sistema inteligente de tratamento de erros de build"""
    memory = load_memory()
    known_issues = memory.get("known_issues", {})
    
    # Detectar tipo de erro
    error_type = classify_build_error(error_output)
    
    if error_type in known_issues:
        print(f"🧠 Erro conhecido detectado: {error_type}")
        solution = known_issues[error_type]
        
        if ask_user_once(f"Aplicar solução conhecida para '{error_type}'?"):
            return apply_solution(python_exe, solution)
    else:
        # Novo tipo de erro
        print(f"🆕 Novo tipo de erro detectado: {error_type}")
        solution = suggest_solution(error_output)
        
        if solution and ask_user_once(f"Tentar solução sugerida: {solution['description']}?"):
            success = apply_solution(python_exe, solution)
            
            if success:
                # Salvar solução que funcionou
                known_issues[error_type] = solution
                memory["known_issues"] = known_issues
                save_memory(memory)
                print(f"💾 Solução salva para futuros usos")
            
            return success
    
    return False

def classify_build_error(error_output):
    """Classifica o tipo de erro baseado na saída"""
    error_lower = error_output.lower()
    
    if "multiple qt bindings" in error_lower:
        return "qt_conflict"
    elif "modulenotfounderror" in error_lower and "zipfile" in error_lower:
        return "missing_zipfile"
    elif "modulenotfounderror" in error_lower:
        return "missing_module"
    elif "hook_dirs" in error_lower or "hooks_contrib" in error_lower:
        return "hook_error"
    elif "timeout" in error_lower or "timed out" in error_lower:
        return "timeout_error"
    elif "permission" in error_lower or "access" in error_lower:
        return "permission_error"
    else:
        return "unknown_error"

def suggest_solution(error_output):
    """Sugere solução baseada no erro"""
    error_type = classify_build_error(error_output)
    
    solutions = {
        "qt_conflict": {
            "description": "Remover bibliotecas Qt conflitantes",
            "action": "remove_qt_conflicts"
        },
        "missing_zipfile": {
            "description": "Reduzir exclusões de módulos",
            "action": "reduce_exclusions"
        },
        "missing_module": {
            "description": "Instalar módulo faltante",
            "action": "install_missing_module"
        },
        "hook_error": {
            "description": "Reinstalar PyInstaller",
            "action": "reinstall_pyinstaller"
        },
        "timeout_error": {
            "description": "Aumentar timeout e limpar cache",
            "action": "increase_timeout"
        },
        "permission_error": {
            "description": "Executar como administrador",
            "action": "admin_rights"
        }
    }
    
    return solutions.get(error_type)

def apply_solution(python_exe, solution):
    """Aplica uma solução específica"""
    action = solution.get("action")
    
    if action == "remove_qt_conflicts":
        return auto_resolve_conflict(python_exe, "qt_conflict:PyQt5,PyQt6,PySide6")
    elif action == "reduce_exclusions":
        print("🔧 Próximo build usará menos exclusões")
        return True
    elif action == "reinstall_pyinstaller":
        return fix_pyinstaller_hooks(python_exe)
    elif action == "increase_timeout":
        print("🔧 Próximo build usará timeout maior")
        return True
    else:
        print(f"⚠️ Solução '{action}' não implementada")
        return False

def find_main_file():
    """Encontra o arquivo principal do KeepAlive"""
    possible_names = [
        "keep-alive-app.py",
        "keepalive-app.py", 
        "keepalive.py",
        "keep_alive.py",
        "main.py",
        "app.py"
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
                choice = input(f"\nEscolha o arquivo principal (1-{len(py_files)}): ").strip()
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
    
    # Verificar se é aplicação GUI (para evitar console)
    is_gui_app = detect_gui_application(main_file)
    
    # Comando PyInstaller ULTRA OTIMIZADO com configurações GUI corretas
    cmd = [
        python_exe, "-m", "PyInstaller",
        "--onefile",                       # Arquivo único
        "--optimize=2",                    # Otimização máxima
        "--strip",                         # Remove símbolos debug
        "--noupx",                         # Evita problemas com antivírus
        "--clean",                         # Limpa cache
        "--noconfirm",                     # Não pedir confirmação
    ]
    
    # 🔥 CONFIGURAÇÃO CORRETA PARA GUI
    if is_gui_app:
        cmd.extend([
            "--windowed",                  # SEM console para aplicações GUI
            "--noconsole",                 # Força remoção do console
        ])
        print("🖥️ Aplicação GUI detectada - console será ocultado")
    else:
        cmd.append("--console")            # COM console para aplicações CLI
        print("⌨️ Aplicação CLI detectada - console será mantido")
    
    # 🚫 EXCLUSÕES ESPECÍFICAS PARA CONFLITO Qt
    cmd.extend([
        "--exclude-module=PyQt5",          # Força exclusão PyQt5
        "--exclude-module=PySide6",        # Força exclusão PySide6
        "--exclude-module=qtpy",           # Exclui qtpy que causa detecção múltipla
        
        # 🎯 EXCLUSÕES AGRESSIVAS - Módulos desnecessários (CORRIGIDO)
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
        # REMOVIDO: zipfile, json, xml, html, email, http - PyInstaller precisa desses
        # REMOVIDO: pickle, gzip, bz2, lzma - podem ser necessários
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
        main_file
    ])
    
    # 🔥 Adicionar variável de ambiente para forçar PyQt6
    env = os.environ.copy()
    env['QT_API'] = 'pyqt6'
    
    print("\n⚡ Executando PyInstaller...")
    print("🎯 Forçando uso exclusivo do PyQt6...")
    if is_gui_app:
        print("🚫 Console será ocultado (aplicação GUI)")
    print("⏱️ Isso pode levar alguns minutos...")
    
    try:
        # Executar com variável de ambiente específica
        result = subprocess.run(cmd, capture_output=True, text=True, 
                              timeout=300, env=env)
        
        # Mostrar output relevante
        if result.stdout:
            lines = result.stdout.split('\n')
            for line in lines:
                if any(keyword in line for keyword in ["INFO:", "WARNING:", "ERROR:", "Building", "Analyzing"]):
                    if len(line.strip()) > 0:  # Evitar linhas vazias
                        print(f"  {line}")
        
        if result.returncode == 0:
            print("✅ Build concluído com sucesso!")
            
            # Verificar arquivo gerado
            exe_path = output_dir / "KeepAliveRDP.exe"
            if exe_path.exists():
                size_mb = exe_path.stat().st_size / (1024 * 1024)
                print(f"\n📦 Arquivo gerado:")
                print(f"   📁 Local: {exe_path.absolute()}")
                print(f"   📏 Tamanho: {size_mb:.1f} MB")
                
                if is_gui_app:
                    print(f"   🖥️ Tipo: Aplicação GUI (sem console)")
                else:
                    print(f"   ⌨️ Tipo: Aplicação CLI (com console)")
                
                # Salvar build bem-sucedido na memória
                memory = load_memory()
                memory["last_successful_build"] = {
                    "python_exe": python_exe,
                    "timestamp": datetime.now().isoformat(),
                    "exe_path": str(exe_path.absolute()),
                    "size_mb": size_mb,
                    "is_gui": is_gui_app
                }
                save_memory(memory)
                
                # Limpar arquivos temporários
                cleanup_build_files(build_dir, spec_dir)
                
                print(f"\n🎉 EXE pronto para uso!")
                return True
            else:
                print("❌ Arquivo EXE não foi encontrado!")
                return False
        else:
            print(f"❌ Build falhou com código: {result.returncode}")
            
            # Usar sistema inteligente de tratamento de erros
            error_output = result.stderr + result.stdout
            if intelligent_build_error_handler(python_exe, error_output):
                print("🔧 Erro corrigido automaticamente. Tentando build novamente...")
                return build_optimized_exe(python_exe)  # Retry uma vez
            
            print("📋 Detalhes do erro:")
            if result.stderr:
                print(result.stderr[:500] + "..." if len(result.stderr) > 500 else result.stderr)
            
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

def detect_gui_application(main_file):
    """Detecta se é uma aplicação GUI baseada no código"""
    try:
        with open(main_file, 'r', encoding='utf-8') as f:
            content = f.read().lower()
        
        # Indicadores de aplicação GUI
        gui_indicators = [
            'pyqt', 'pyside', 'tkinter', 'kivy', 'wxpython',
            'qapplication', 'qwidget', 'qmainwindow',
            'systemtrayicon', 'qtrayicon', 'qtoolbar',
            'qmenu', 'qdialog', 'qpushbutton',
            'from PyQt', 'from PySide', 'import PyQt', 'import PySide'
        ]
        
        # Verificar se contém indicadores GUI
        for indicator in gui_indicators:
            if indicator in content:
                return True
        
        # Verificar por extensões de arquivo específicas
        if main_file.endswith(('-gui.py', '-app.py', '-window.py')):
            return True
        
        return False
        
    except Exception:
        # Se não conseguir ler o arquivo, assumir que é GUI baseado no nome
        return 'app' in main_file.lower() or 'gui' in main_file.lower()

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

def extract_system_icon():
    """Extrai ícone do sistema Windows para usar no EXE"""
    try:
        print("🎨 Tentando extrair ícone do sistema...")
        
        # Método 1: Usar PowerShell para extrair ícone
        powershell_cmd = [
            "powershell", "-Command",
            """
            Add-Type -AssemblyName System.Drawing
            $icon = [System.Drawing.SystemIcons]::Computer
            $bitmap = $icon.ToBitmap()
            $bitmap.Save('computer_icon.png', [System.Drawing.Imaging.ImageFormat]::Png)
            Write-Host 'Icon extracted successfully'
            """
        ]
        
        result = subprocess.run(powershell_cmd, capture_output=True, text=True, timeout=30)
        
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
        img = Image.new('RGBA', (32, 32), (0, 100, 200, 255))
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
    
    # Comando otimizado MAS com menos exclusões para evitar problemas
    cmd = [
        python_exe, "-m", "PyInstaller",
        "--onefile",
        "--windowed",
        "--optimize=2",
        "--strip",
        "--clean",
        f"--icon={icon_file}",
        
        # 🚫 EXCLUSÕES ESPECÍFICAS PARA CONFLITO Qt
        "--exclude-module=PyQt5",          
        "--exclude-module=PySide6",        
        "--exclude-module=qtpy",           
        
        # 🎯 EXCLUSÕES BÁSICAS (removemos zipfile da lista!)
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
        # NÃO excluir: zipfile, json, xml, html - PyInstaller precisa deles
        "--exclude-module=unittest",
        "--exclude-module=doctest",
        "--exclude-module=pdb",
        "--exclude-module=turtle",
        
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
        
        f"--name=KeepAliveRDP_Icon",
        f"--distpath={output_dir}",
        f"--workpath={build_dir}",
        f"--specpath={spec_dir}",
        main_file
    ]
    
    # Variável de ambiente para forçar PyQt6
    env = os.environ.copy()
    env['QT_API'] = 'pyqt6'
    
    try:
        print("⚡ Executando build com ícone...")
        result = subprocess.run(cmd, capture_output=True, text=True, 
                              timeout=300, env=env)
        
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
            error_msg = result.stderr[:300]
            print(f"Erro: {error_msg}...")
            
            # Se é erro de módulo faltando, tentar build simplificado
            if "ModuleNotFoundError" in result.stderr or "zipfile" in result.stderr:
                print("🔧 Erro de módulo detectado. Tentando build simplificado...")
                return build_simple_with_icon(python_exe, icon_file)
        
    except Exception as e:
        print(f"❌ Erro no build com ícone: {e}")
    
    return False

def build_simple_with_icon(python_exe, icon_file):
    """Build simplificado com ícone (menos exclusões)"""
    print("\n🔧 Tentando build simplificado com ícone...")
    
    main_file = find_main_file()
    if not main_file:
        return False
    
    # Preparar diretórios
    output_dir = Path("dist_simple_icon")
    build_dir = Path("build_simple")
    spec_dir = Path("specs_simple")
    
    # Limpar diretórios antigos
    for dir_path in [output_dir, build_dir, spec_dir]:
        if dir_path.exists():
            try:
                shutil.rmtree(dir_path)
            except:
                pass
    
    # Comando MÍNIMO com ícone
    cmd = [
        python_exe, "-m", "PyInstaller",
        "--onefile",
        "--windowed",
        "--clean",
        f"--icon={icon_file}",
        
        # Apenas exclusões essenciais Qt
        "--exclude-module=PyQt5",          
        "--exclude-module=PySide6",        
        
        f"--name=KeepAliveRDP_SimpleIcon",
        f"--distpath={output_dir}",
        f"--workpath={build_dir}",
        f"--specpath={spec_dir}",
        main_file
    ]
    
    # Variável de ambiente
    env = os.environ.copy()
    env['QT_API'] = 'pyqt6'
    
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, 
                              timeout=300, env=env)
        
        if result.returncode == 0:
            exe_path = output_dir / "KeepAliveRDP_SimpleIcon.exe"
            if exe_path.exists():
                size_mb = exe_path.stat().st_size / (1024 * 1024)
                print(f"✅ Build simplificado com ícone concluído!")
                print(f"📁 Local: {exe_path.absolute()}")
                print(f"📏 Tamanho: {size_mb:.1f} MB")
                
                cleanup_build_files(build_dir, spec_dir)
                return True
        
        print("❌ Build simplificado também falhou")
        if result.stderr:
            print(f"Erro: {result.stderr[:200]}...")
        
    except Exception as e:
        print(f"❌ Erro no build simplificado: {e}")
    
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
        python_exe, "-m", "PyInstaller",
        "--onefile",
        "--windowed",
        "--clean",
        "--name=KeepAliveRDP_Test",
        "--distpath=dist_test",
        main_file
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
        if retry in ('', 's', 'sim', 'y', 'yes'):
            success = build_quick_test(python_exe)
            if success:
                print("\n🎉 Build de teste funcionou! Problema pode ser nas otimizações.")

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