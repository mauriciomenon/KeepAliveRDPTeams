import requests

# Configuração das credenciais corretas
CONFIG = {
    "client_id": "c36ea277-60f5-48f3-b111-52133d9a6859",  # ID do aplicativo (cliente)
  
    "tenant_id": "db31e804-2b01-4548-b3d4-642d3cde201f",  # ID do diretório (locatário)
    "user_id": "b206ca7f-4f0e-4599-ab2b-4740d93ad276",  # ID do Objeto do usuário Maurício Menon
}


# Obter token de acesso
def get_access_token():
    token_url = (
        f"https://login.microsoftonline.com/{CONFIG['tenant_id']}/oauth2/v2.0/token"
    )
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    data = {
        "grant_type": "client_credentials",
        "client_id": CONFIG["client_id"],
        "client_secret": CONFIG["client_secret"],
        "scope": "https://graph.microsoft.com/.default",
    }

    response = requests.post(token_url, headers=headers, data=data)
    if response.status_code == 200:
        return response.json().get("access_token")
    else:
        print(f"❌ Erro ao obter token: {response.status_code} - {response.text}")
        return None


# Consultar status do usuário no Teams
def get_teams_status():
    access_token = get_access_token()
    if not access_token:
        return

    graph_url = f"https://graph.microsoft.com/v1.0/users/{CONFIG['user_id']}/presence"
    headers_auth = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
    }

    response = requests.get(graph_url, headers=headers_auth)
    if response.status_code == 200:
        print("✅ Status do usuário no Teams:")
        print(response.json())
    else:
        print(f"❌ Erro ao obter status: {response.status_code} - {response.text}")


# Executar a verificação de status
get_teams_status()
