"""
Módulo para verificação do status do Microsoft Teams
"""

import random
import subprocess
import time

import pyautogui
import win32con
import win32gui

# Novo import para UI Automation
try:
    import uiautomation as auto

    UI_AUTOMATION_AVAILABLE = True
    print("DEBUG TEAMS: UI Automation disponível")
except ImportError:
    UI_AUTOMATION_AVAILABLE = False
    print(
        "DEBUG TEAMS: UI Automation não disponível. Instale com: pip install uiautomation"
    )


def get_teams_status():
    """Verifica status real do Teams (disponível/ausente/ocupado)"""
    try:
        print("DEBUG TEAMS: Iniciando verificação de status do Teams...")

        # Método 0: UI Automation - NOVO MÉTODO PRINCIPAL
        if UI_AUTOMATION_AVAILABLE:
            print("DEBUG TEAMS: Método 0 - UI Automation (lendo avatar)...")
            try:
                print("DEBUG TEAMS: Procurando janelas Teams com UI Automation...")

                # Busca todas as janelas e filtra as do Teams
                teams_windows_found = []

                # Método alternativo: busca por todas as janelas que terminam com "Microsoft Teams"
                def find_teams_windows():
                    found_windows = []
                    try:
                        # Enumera todas as janelas top-level
                        all_windows = auto.GetRootControl().GetChildren()
                        print(f"DEBUG TEAMS: Verificando {len(all_windows)} janelas...")

                        for window in all_windows:
                            try:
                                if hasattr(window, "Name") and window.Name:
                                    window_name = window.Name
                                    if "microsoft teams" in window_name.lower():
                                        print(
                                            f"DEBUG TEAMS: Janela Teams encontrada: '{window_name}'"
                                        )
                                        found_windows.append(window)
                            except Exception:
                                continue
                    except Exception as e:
                        print(f"DEBUG TEAMS: Erro ao enumerar janelas: {str(e)}")

                    return found_windows

                teams_windows_found = find_teams_windows()

                if not teams_windows_found:
                    print(
                        "DEBUG TEAMS: Nenhuma janela Teams encontrada via UI Automation"
                    )
                else:
                    print(
                        f"DEBUG TEAMS: {len(teams_windows_found)} janela(s) Teams encontrada(s)"
                    )

                    # Para cada janela Teams, procura o botão do avatar
                    for teams_window in teams_windows_found:
                        try:
                            print(
                                f"DEBUG TEAMS: Analisando janela: '{teams_window.Name}'"
                            )

                            # Procura botões na janela
                            buttons = teams_window.GetChildren()
                            print(
                                f"DEBUG TEAMS: Verificando {len(buttons)} elementos na janela..."
                            )

                            def search_avatar_button(element, depth=0):
                                if depth > 3:  # Limita profundidade para evitar loops
                                    return None

                                try:
                                    if (
                                        hasattr(element, "ControlTypeName")
                                        and element.ControlTypeName == "Button"
                                        and hasattr(element, "Name")
                                        and element.Name
                                    ):

                                        name_lower = element.Name.lower()
                                        print(
                                            f"DEBUG TEAMS: Botão encontrado: '{element.Name}'"
                                        )

                                        # Verifica frases-chave de status
                                        status_phrases = [
                                            "status displayed as",
                                            "status exibido como",
                                            "com status exibido como",
                                            "profile picture",
                                            "foto de perfil",
                                            "your avatar",
                                            "seu avatar",
                                        ]

                                        if any(
                                            phrase in name_lower
                                            for phrase in status_phrases
                                        ):
                                            print(
                                                f"DEBUG TEAMS: AVATAR ENCONTRADO: '{element.Name}'"
                                            )
                                            return element

                                    # Busca recursivamente nos filhos
                                    for child in element.GetChildren():
                                        result = search_avatar_button(child, depth + 1)
                                        if result:
                                            return result

                                except Exception:
                                    pass

                                return None

                            avatar_button = search_avatar_button(teams_window)

                            if avatar_button:
                                # Extrai status do nome do botão
                                button_text = avatar_button.Name
                                print(f"DEBUG TEAMS: Texto do avatar: '{button_text}'")

                                # Procura por palavras de status no texto
                                words = button_text.split()
                                status_raw = None

                                # Procura a palavra após "as" ou "como"
                                for i, word in enumerate(words):
                                    if word.lower() in ["as", "como"] and i + 1 < len(
                                        words
                                    ):
                                        status_raw = words[i + 1].strip(".,!").strip()
                                        break

                                if not status_raw and words:
                                    status_raw = words[-1].strip(".,!").strip()

                                if status_raw:
                                    print(
                                        f"DEBUG TEAMS: Status extraído: '{status_raw}'"
                                    )

                                    # Mapeia status (sem emojis)
                                    status_map = {
                                        "Available": "DISPONÍVEL",
                                        "Online": "DISPONÍVEL",
                                        "Disponível": "DISPONÍVEL",
                                        "Busy": "OCUPADO",
                                        "Ocupado": "OCUPADO",
                                        "InAMeeting": "OCUPADO",
                                        "Away": "AUSENTE",
                                        "Ausente": "AUSENTE",
                                        "BeRightBack": "AUSENTE",
                                        "Volto": "AUSENTE",
                                        "DoNotDisturb": "NÃO_PERTURBE",
                                        "Dnd": "NÃO_PERTURBE",
                                        "Não": "NÃO_PERTURBE",
                                        "incomodar": "NÃO_PERTURBE",
                                        "Offline": "OFFLINE",
                                        "Desconectado": "OFFLINE",
                                        "offline": "OFFLINE",
                                    }

                                    # Busca mapeamento
                                    final_status = None
                                    for key, value in status_map.items():
                                        if key.lower() in status_raw.lower():
                                            final_status = value
                                            break

                                    if not final_status:
                                        final_status = status_raw.upper()

                                    print(
                                        f"DEBUG TEAMS: Status final via UI Automation: {final_status}"
                                    )
                                    return (
                                        final_status,
                                        f"UI Automation: {final_status}",
                                    )

                        except Exception as e:
                            print(
                                f"DEBUG TEAMS: Erro ao analisar janela Teams: {str(e)}"
                            )
                            continue

                    print(
                        "DEBUG TEAMS: Botão do avatar não encontrado em nenhuma janela"
                    )

            except Exception as e:
                print(f"DEBUG TEAMS: Erro no UI Automation: {str(e)}")

        # Método 1: Verificar processo do Teams
        print("DEBUG TEAMS: Método 1 - Verificando processos do Teams...")
        teams_processes = []
        try:
            print("DEBUG TEAMS: Procurando ms-teams.exe...")
            for proc in subprocess.run(
                ["tasklist", "/FI", "IMAGENAME eq ms-teams.exe", "/FO", "CSV"],
                capture_output=True,
                text=True,
                shell=True,
            ).stdout.split("\n"):
                if "ms-teams.exe" in proc.lower():
                    teams_processes.append(proc)
                    print(f"DEBUG TEAMS: Encontrado ms-teams.exe: {proc.strip()}")

            # Se não tem processo do Teams, tenta o clássico
            if not teams_processes:
                print(
                    "DEBUG TEAMS: ms-teams.exe não encontrado, procurando Teams.exe clássico..."
                )
                for proc in subprocess.run(
                    ["tasklist", "/FI", "IMAGENAME eq Teams.exe", "/FO", "CSV"],
                    capture_output=True,
                    text=True,
                    shell=True,
                ).stdout.split("\n"):
                    if "Teams.exe" in proc.lower():
                        teams_processes.append(proc)
                        print(f"DEBUG TEAMS: Encontrado Teams.exe: {proc.strip()}")

            if not teams_processes:
                print("DEBUG TEAMS: Nenhum processo Teams encontrado")
                return "INATIVO", "Teams não está em execução"
            else:
                print(
                    f"DEBUG TEAMS: {len(teams_processes)} processo(s) Teams encontrado(s)"
                )

        except Exception as e:
            print(f"DEBUG TEAMS: Erro no Método 1: {str(e)}")

        # Método 2: Verificar janelas do Teams e título para status
        print("DEBUG TEAMS: Método 2 - Verificando janelas do Teams...")
        teams_windows = []
        try:

            def enum_windows_callback(hwnd, results):
                try:
                    if win32gui.IsWindowVisible(hwnd):
                        window_text = win32gui.GetWindowText(hwnd)
                        if "teams" in window_text.lower() and len(window_text) > 5:
                            if "microsoft teams" in window_text.lower():
                                print(
                                    f"DEBUG TEAMS: Janela TEAMS REAL encontrada: '{window_text}'"
                                )
                                rect = win32gui.GetWindowRect(hwnd)
                                width = rect[2] - rect[0]
                                height = rect[3] - rect[1]
                                print(
                                    f"DEBUG TEAMS: Tamanho da janela: {width}x{height}"
                                )
                                if width > 100 and height > 100:
                                    results.append(window_text)

                                    title_lower = window_text.lower()
                                    print(
                                        f"DEBUG TEAMS: Analisando título: '{title_lower}'"
                                    )

                                    away_words = [
                                        "away",
                                        "ausente",
                                        "busy",
                                        "ocupado",
                                        "offline",
                                        "não disponível",
                                    ]
                                    available_words = [
                                        "available",
                                        "disponível",
                                        "online",
                                        "ativo",
                                    ]

                                    if any(word in title_lower for word in away_words):
                                        print(
                                            f"DEBUG TEAMS: Status AUSENTE detectado no título (palavra encontrada)"
                                        )
                                        results.append("STATUS_AUSENTE")
                                    elif any(
                                        word in title_lower for word in available_words
                                    ):
                                        print(
                                            f"DEBUG TEAMS: Status DISPONÍVEL detectado no título (palavra encontrada)"
                                        )
                                        results.append("STATUS_DISPONIVEL")
                                    else:
                                        print(
                                            "DEBUG TEAMS: Nenhum status específico detectado no título - assumindo ATIVO"
                                        )
                                        results.append("STATUS_ATIVO_SEM_INDICACAO")
                                else:
                                    print(
                                        "DEBUG TEAMS: Janela muito pequena, ignorando"
                                    )
                            else:
                                print(
                                    f"DEBUG TEAMS: Janela ignorada (não é Teams real): '{window_text}'"
                                )
                except Exception as e:
                    print(f"DEBUG TEAMS: Erro ao processar janela: {str(e)}")
                return True

            windows_found = []
            win32gui.EnumWindows(enum_windows_callback, windows_found)

            teams_windows = [w for w in windows_found if not w.startswith("STATUS_")]
            status_flags = [w for w in windows_found if w.startswith("STATUS_")]

            print(f"DEBUG TEAMS: Total de janelas Teams reais: {len(teams_windows)}")
            print(f"DEBUG TEAMS: Status encontrados: {status_flags}")

            if teams_windows:
                if "STATUS_AUSENTE" in status_flags:
                    print("DEBUG TEAMS: Resultado final - AUSENTE")
                    return "AUSENTE", "Ausente"
                elif "STATUS_DISPONIVEL" in status_flags:
                    print("DEBUG TEAMS: Resultado final - DISPONÍVEL")
                    return "DISPONÍVEL", "Disponível"
                elif "STATUS_ATIVO_SEM_INDICACAO" in status_flags:
                    print(
                        "DEBUG TEAMS: Resultado final - ATIVO (sem indicação no título)"
                    )
                    return "ATIVO", f"Ativo - {len(teams_windows)} janela(s) Teams"
                else:
                    print("DEBUG TEAMS: Resultado final - ATIVO (indeterminado)")
                    return "ATIVO", f"Indeterminado - {len(teams_windows)} janela(s)"

        except Exception as e:
            print(f"DEBUG TEAMS: Erro no Método 2: {str(e)}")

        # Método 3: Fallback simples - só verifica se está rodando
        print("DEBUG TEAMS: Método 3 - Fallback com pyautogui...")
        try:
            teams_found = False
            windows_count = 0
            for window in pyautogui.getAllWindows():
                if hasattr(window, "title") and window.title:
                    if (
                        "microsoft teams" in window.title.lower()
                        and window.width > 100
                        and window.height > 100
                    ):
                        print(
                            f"DEBUG TEAMS: Janela pyautogui Teams: '{window.title}' {window.width}x{window.height}"
                        )
                        teams_found = True
                        windows_count += 1

            print(f"DEBUG TEAMS: Método 3 encontrou {windows_count} janela(s) Teams")
            if teams_found:
                print("DEBUG TEAMS: Resultado final - DETECTADO (fallback)")
                return "DETECTADO", f"Detectado (fallback) - {windows_count} janela(s)"

        except Exception as e:
            print(f"DEBUG TEAMS: Erro no Método 3: {str(e)}")

        # Se chegou até aqui mas tem processo, Teams está rodando mas sem janelas visíveis
        if teams_processes:
            print("DEBUG TEAMS: Resultado final - PROCESSO SEM JANELA")
            return "PROCESSO", "Processo ativo mas sem janelas visíveis"

        print("DEBUG TEAMS: Resultado final - INDETERMINADO")
        return "INDETERMINADO", "Não foi possível determinar status do Teams"

    except Exception as e:
        print(f"DEBUG TEAMS: Erro geral: {str(e)}")
        return None, f"Erro na verificação: {str(e)}"


def test_teams_detection():
    """Testa a detecção do Teams uma vez"""
    print("\n" + "=" * 50)
    print("TESTE DE DETECÇÃO DO TEAMS")
    print("=" * 50)

    status, message = get_teams_status()

    print(f"\nRESULTADO:")
    print(f"   Status: {status}")
    print(f"   Mensagem: {message}")
    print("=" * 50 + "\n")

    return status, message


def monitor_teams_continuous():
    """Monitora o Teams continuamente"""
    print("\nMONITORAMENTO CONTÍNUO DO TEAMS")
    print("   Pressione Ctrl+C para parar\n")

    last_status = None
    check_count = 0

    try:
        while True:
            check_count += 1
            print(f"[Verificação #{check_count}] ", end="", flush=True)

            status, message = get_teams_status()

            # Só mostra resultado se mudou ou a cada 10 verificações
            if status != last_status or check_count % 10 == 0:
                timestamp = time.strftime("%H:%M:%S")
                if status != last_status:
                    print(f"[{timestamp}] MUDANÇA: {status} - {message}")
                    last_status = status
                else:
                    print(f"[{timestamp}] Status: {status}")
            else:
                print(".", end="", flush=True)

            time.sleep(5)  # Verifica a cada 5 segundos

    except KeyboardInterrupt:
        print("\n\nMonitoramento interrompido pelo usuário")
    except Exception as e:
        print(f"\nErro no monitoramento: {str(e)}")


if __name__ == "__main__":
    import sys

    print("TEAMS STATUS CHECKER")
    print("=" * 30)

    # Verifica se UI Automation está disponível
    if UI_AUTOMATION_AVAILABLE:
        print("UI Automation: Disponível")
    else:
        print("UI Automation: Não disponível")
        print("   Para detecção avançada, instale: pip install uiautomation")

    print("\nOpções:")
    print("1. Teste único")
    print("2. Monitoramento contínuo")
    print("3. Sair")

    try:
        choice = input("\nEscolha uma opção (1-3): ").strip()

        if choice == "1":
            test_teams_detection()
        elif choice == "2":
            monitor_teams_continuous()
        elif choice == "3":
            print("Até logo!")
        else:
            print("Opção inválida. Executando teste único...")
            test_teams_detection()

    except KeyboardInterrupt:
        print("\nAté logo!")
    except Exception as e:
        print(f"\nErro: {str(e)}")
        sys.exit(1)
