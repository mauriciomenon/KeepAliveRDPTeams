"""
Módulo para verificação do status do Microsoft Teams
"""

import random
import subprocess
import time

import pyautogui
import win32con
import win32gui


def get_teams_status():
    """Verifica status real do Teams (disponível/ausente/ocupado)"""
    try:
        print("DEBUG TEAMS: Iniciando verificação de status do Teams...")

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
