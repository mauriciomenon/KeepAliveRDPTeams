# Ambiente RDP/VNC com Podman

Este repositório contém scripts para criar um ambiente remoto com acesso via RDP e VNC (noVNC) usando Podman.

## Pré-requisitos
- Podman instalado
- PowerShell 7+ (recomendado)
- Execução como administrador (para acessar portas baixas)

## Configuração Inicial

1. Clone o repositório:
   ```bash
   git clone https://github.com/seu-usuario/podman-rdp-vnc.git
   cd podman-rdp-vnc


Construa a imagem (apenas primeira vez):

podman build -t danielguerra/ubuntu-xrdp .

# Ver status
.\manage-container.ps1 -Action status

# Iniciar container
.\manage-container.ps1 -Action start

# Parar container
.\manage-container.ps1 -Action stop

# Reiniciar container
.\manage-container.ps1 -Action restart