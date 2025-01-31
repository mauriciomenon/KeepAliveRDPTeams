import os
import json
import logging
from enum import Enum
from typing import Optional
import requests
from datetime import datetime, timedelta

# Configuração de logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

class TeamsStatus(Enum):
    """Enumeração dos status possíveis do Teams"""
    AVAILABLE = ("Disponível", "Available")
    BUSY = ("Ocupado", "Busy")
    DO_NOT_DISTURB = ("Não incomodar", "DoNotDisturb")
    BE_RIGHT_BACK = ("Volto logo", "BeRightBack")
    AWAY = ("Ausente", "Away")
    OFFLINE = ("Offline", "Offline")

    def __init__(self, display_name: str, graph_status: str):
        self.display_name = display_name
        self.graph_status = graph_status

class TeamsGraphManager:
    """Gerenciador de status do Teams usando Microsoft Graph API"""
    
    def __init__(self, client_id: str, client_secret: str, tenant_id: str):
        """
        Inicializa o gerenciador
        Args:
            client_id: ID do aplicativo registrado no Azure AD
            client_secret: Segredo do aplicativo
            tenant_id: ID do tenant do Azure AD
        """
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self.access_token = None
        self.token_expires = None
        
    def _get_access_token(self) -> Optional[str]:
        """
        Obtém ou renova o token de acesso
        Returns:
            str: Token de acesso ou None em caso de erro
        """
        try:
            # Verificar se o token atual ainda é válido
            if self.access_token and self.token_expires:
                if datetime.now() < self.token_expires - timedelta(minutes=5):
                    return self.access_token

            # Endpoint para obtenção do token
            token_url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
            
            # Parâmetros para a requisição
            data = {
                'grant_type': 'client_credentials',
                'client_id': self.client_id,
                'client_secret': self.client_secret,
                'scope': 'https://graph.microsoft.com/.default'
            }
            
            # Fazer requisição de token
            response = requests.post(token_url, data=data)
            response.raise_for_status()
            
            # Processar resposta
            token_data = response.json()
            self.access_token = token_data['access_token']
            self.token_expires = datetime.now() + timedelta(seconds=token_data['expires_in'])
            
            logger.info("Token de acesso obtido com sucesso")
            return self.access_token
            
        except Exception as e:
            logger.error(f"Erro ao obter token de acesso: {str(e)}")
            return None
            
    def set_user_status(self, user_id: str, status: TeamsStatus) -> bool:
        """
        Define o status de um usuário no Teams
        Args:
            user_id: ID do usuário no Azure AD
            status: Novo status a ser definido
        Returns:
            bool: True se sucesso, False caso contrário
        """
        try:
            # Obter token de acesso
            access_token = self._get_access_token()
            if not access_token:
                return False
                
            # Endpoint para atualização de status
            url = f"https://graph.microsoft.com/v1.0/users/{user_id}/presence/setStatusMessage"
            
            # Headers da requisição
            headers = {
                'Authorization': f'Bearer {access_token}',
                'Content-Type': 'application/json'
            }
            
            # Payload com novo status
            data = {
                'statusMessage': {
                    'message': '',  # Mensagem opcional
                    'status': status.graph_status
                }
            }
            
            # Fazer requisição
            response = requests.post(url, headers=headers, json=data)
            response.raise_for_status()
            
            logger.info(f"Status alterado com sucesso para: {status.display_name}")
            return True
            
        except Exception as e:
            logger.error(f"Erro ao alterar status: {str(e)}")
            return False
            
    def get_user_status(self, user_id: str) -> Optional[TeamsStatus]:
        """
        Obtém o status atual de um usuário no Teams
        Args:
            user_id: ID do usuário no Azure AD
        Returns:
            TeamsStatus: Status atual ou None em caso de erro
        """
        try:
            # Obter token de acesso
            access_token = self._get_access_token()
            if not access_token:
                return None
                
            # Endpoint para consulta de status
            url = f"https://graph.microsoft.com/v1.0/users/{user_id}/presence"
            
            # Headers da requisição
            headers = {
                'Authorization': f'Bearer {access_token}',
                'Content-Type': 'application/json'
            }
            
            # Fazer requisição
            response = requests.get(url, headers=headers)
            response.raise_for_status()
            
            # Processar resposta
            presence_data = response.json()
            status = presence_data.get('availability')
            
            # Converter para enum
            for teams_status in TeamsStatus:
                if teams_status.graph_status.lower() == status.lower():
                    logger.info(f"Status atual: {teams_status.display_name}")
                    return teams_status
                    
            return None
            
        except Exception as e:
            logger.error(f"Erro ao obter status: {str(e)}")
            return None


# Exemplo de uso
if __name__ == "__main__":
    # Configurações do Azure AD (substitua pelos valores reais)
    CONFIG = {
        'client_id': 'seu_client_id',
        'client_secret': 'seu_client_secret',
        'tenant_id': 'seu_tenant_id',
        'user_id': 'id_do_usuario'  # ID do usuário no Azure AD
    }
    
    # Criar instância do gerenciador
    teams_manager = TeamsGraphManager(
        CONFIG['client_id'],
        CONFIG['client_secret'],
        CONFIG['tenant_id']
    )
    
    # Exemplo: Definir status como Ocupado
    if teams_manager.set_user_status(CONFIG['user_id'], TeamsStatus.BUSY):
        print("Status alterado com sucesso!")
        
        # Consultar status atual
        current_status = teams_manager.get_user_status(CONFIG['user_id'])
        if current_status:
            print(f"Status atual: {current_status.display_name}")