# pylint: disable=E0602,E0102,E1101
import time
import logging
import uiautomation as auto

logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)


class TeamsStatusChanger:
    """Classe para mudar status do Teams usando UI Automation"""

    def __init__(self):
        self.teams_window = None
        self._find_teams_window()

    def _find_teams_window(self) -> None:
        """Encontra a janela principal do Teams usando UIAutomation"""
        try:
            # Procura por janelas do Teams usando o ClassName
            teams_windows = auto.WindowControl(ClassName="TeamsWebView")

            if teams_windows.Exists(maxSearchSeconds=3):
                self.teams_window = teams_windows
                logger.info(f"Found Teams window: {teams_windows.Name}")
            else:
                # Tenta encontrar usando parte do título
                teams_windows = auto.WindowControl(
                    SearchDepth=1, SubName="Microsoft Teams"
                )
                if teams_windows.Exists(maxSearchSeconds=3):
                    self.teams_window = teams_windows
                    logger.info(f"Found Teams window by title: {teams_windows.Name}")
                else:
                    logger.error("Teams window not found")

        except Exception as e:
            logger.error(f"Error finding Teams window: {e}")
            logger.exception("Full traceback:")

    def _print_element_info(self, element, depth=0):
        """Imprime informações detalhadas sobre um elemento e seus filhos"""
        indent = "  " * depth
        logger.info(f"{indent}Element: {element.Name}")
        logger.info(f"{indent}Class: {element.ClassName}")
        logger.info(f"{indent}AutomationId: {element.AutomationId}")
        logger.info(f"{indent}ControlType: {element.ControlTypeName}")
        logger.info(f"{indent}Rect: {element.BoundingRectangle}")

        # Lista todas as propriedades disponíveis em modo debug
        logger.debug(f"{indent}Properties:")
        for prop in dir(element):
            if not prop.startswith("_"):
                try:
                    value = getattr(element, prop)
                    if not callable(value):
                        logger.debug(f"{indent}  {prop}: {value}")
                except:
                    pass

        # Recursivamente imprime informações dos filhos
        for child in element.GetChildren():
            self._print_element_info(child, depth + 1)

    def _find_clickable_element(
        self, element, search_terms, depth=0, max_depth=15, timeout=1.0
    ):
        """Procura recursivamente por um elemento clicável com timeout"""
        if depth > max_depth:
            return None

        try:
            # Define um timeout para evitar congelamentos
            auto.SetGlobalSearchTimeout(timeout)

            # Tenta encontrar por padrões de automation id primeiro - com timeout curto
            for term in search_terms:
                try:
                    button = element.ButtonControl(AutomationId=term, searchDepth=1)
                    if button.Exists(maxSearchSeconds=0.5):
                        logger.info(f"Found button with AutomationId: {term}")
                        return button
                except Exception as e:
                    logger.debug(f"Error finding button by AutomationId {term}: {e}")

            # Tenta encontrar por nome - com timeout curto
            for term in search_terms:
                try:
                    button = element.ButtonControl(Name=term, searchDepth=1)
                    if button.Exists(maxSearchSeconds=0.5):
                        logger.info(f"Found button with Name: {term}")
                        return button
                except Exception as e:
                    logger.debug(f"Error finding button by Name {term}: {e}")

            # Se não encontrou por botão, tenta encontrar qualquer elemento clicável
            try:
                name = element.Name.lower() if element.Name else ""
                class_name = element.ClassName.lower() if element.ClassName else ""
                automation_id = (
                    element.AutomationId.lower() if element.AutomationId else ""
                )

                logger.debug(f"Checking element at depth {depth}:")
                logger.debug(f"  Name: {name}")
                logger.debug(f"  Class: {class_name}")
                logger.debug(f"  AutomationId: {automation_id}")

                # Verifica se este elemento corresponde aos critérios
                for term in search_terms:
                    if (
                        term in name or term in class_name or term in automation_id
                    ) and element.IsEnabled:
                        logger.info(f"Found matching element: {name} ({class_name})")
                        return element

                # Procura nos filhos diretos primeiro (searchDepth=1)
                children = element.GetChildren()
                if children:
                    for child in children:
                        result = self._find_clickable_element(
                            child, search_terms, depth + 1, max_depth, timeout
                        )
                        if result:
                            return result

            except Exception as e:
                logger.debug(f"Error checking element attributes: {e}")

        except Exception as e:
            logger.debug(f"Error in find_clickable_element: {e}")
        finally:
            # Restaura o timeout padrão
            auto.SetGlobalSearchTimeout(10.0)

        return None

    def change_status(self, status: str) -> bool:
        """Muda o status do Teams"""
        if not self.teams_window:
            logger.error("Teams window not found")
            return False

        try:
            # Verifica se a janela do Teams foi encontrada
            if not self.teams_window:
                logger.error("Teams window not found")
                return False

            # Ativa a janela do Teams
            self.teams_window.SetFocus()
            time.sleep(1)  # Aguarda a janela ser ativada

            # Usa a própria janela do Teams como controle principal
            teams_control = self.teams_window

            logger.info("Starting element tree inspection...")
            self._print_element_info(self.teams_window)

            # Procura pelo botão de perfil/status
            # Amplia os termos de busca para incluir IDs conhecidos do Teams
            profile_terms = [
                "personButton",  # ID comum do botão de perfil
                "collaboratorprofilephoto",  # ID da foto de perfil
                "status-button",
                "presence-button",
                "me-button",
            ]

            logger.info("Searching for profile/status button...")
            profile_elem = self._find_clickable_element(teams_control, profile_terms)

            if profile_elem:
                logger.info("Found profile element, attempting to click...")
                profile_elem.Click()
                time.sleep(1)

                # Procura pela opção de status específica
                status_option = self._find_clickable_element(
                    teams_control, [status.lower()]
                )
                if status_option:
                    logger.info(f"Found status option {status}, clicking...")
                    status_option.Click()
                    return True
                else:
                    logger.error(f"Status option {status} not found")
            else:
                logger.error("Profile/status button not found")

            return False

        except Exception as e:
            logger.error(f"Error changing status: {e}")
            logger.exception("Full traceback:")
            return False


def main():
    changer = TeamsStatusChanger()

    statuses = [
        "Disponível",
        "Ocupado",
        "Não incomodar",
        "Volto logo",
        "Aparecer como ausente",
        "Aparecer offline",
    ]

    print("\nStatus disponíveis:")
    for s in statuses:
        print(f"- {s}")

    status = "Disponível"
    print(f"\nTentando mudar status para: {status}")
    if changer.change_status(status):
        print("Status alterado com sucesso")
    else:
        print("Falha ao alterar status")


if __name__ == "__main__":
    main()
