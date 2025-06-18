<#
.SYNOPSIS
    Inicia um container com servidor RDP e VNC (noVNC) usando Podman
.DESCRIPTION
    Este script:
    - Remove containers antigos com o mesmo nome
    - Constrói a imagem (se não existir) usando o Dockerfile no diretório atual
    - Encontra portas livres para VNC (noVNC) e RDP
    - Inicia o container mapeando para as portas livres
    - Abre o navegador para o VNC e o cliente RDP
.NOTES
    Credenciais padrão:
      Usuário: user
      Senha: password
    Requer Podman instalado e configurado.
    PowerShell 7+ e execução como administrador são recomendados.

.Construa a imagem (apenas primeira vez):
    podman build -t danielguerra/ubuntu-xrdp .
#>

# Configurações
$IMAGE_NAME = "danielguerra/ubuntu-xrdp"
$CONTAINER_NAME = "rdp-vnc-test"

# 1. Remove container antigo se existir
podman stop $CONTAINER_NAME 2>$null
podman rm $CONTAINER_NAME 2>$null

# 2. Constrói a imagem se não existir
if (-not (podman images --format "{{.Repository}}" | Where-Object { $_ -eq $IMAGE_NAME })) {
    Write-Host "Construindo imagem persistente..." -ForegroundColor Cyan
    podman build -t $IMAGE_NAME .
}
else {
    Write-Host "Usando imagem persistente existente" -ForegroundColor Green
}

# 3. Encontra portas livres
function Get-FreePort {
    $port = Get-Random -Minimum 10000 -Maximum 60000
    while ((Test-NetConnection -ComputerName localhost -Port $port -WarningAction SilentlyContinue).TcpTestSucceeded) {
        $port++
    }
    return $port
}

$vncPort = Get-FreePort
$rdpPort = Get-FreePort

# 4. Inicia o container
Write-Host "Iniciando container..." -ForegroundColor Cyan
podman run -d --name $CONTAINER_NAME `
    -p ${vncPort}:6080 `
    -p ${rdpPort}:3389 `
    $IMAGE_NAME

# 5. Aguarda inicialização (20 segundos para garantir que todos os serviços estejam prontos)
Write-Host "Aguardando inicialização (20 segundos)..." -ForegroundColor Cyan
Start-Sleep -Seconds 20

# 6. Abre as conexões
$vncUrl = "http://localhost:$vncPort/vnc.html"
$rdpArgs = "/v:localhost:$rdpPort"

Write-Host "Abrindo cliente RDP e navegador para VNC..." -ForegroundColor Cyan
try {
    Start-Process "mstsc" -ArgumentList $rdpArgs -ErrorAction Stop
}
catch {
    Write-Warning "Não foi possível abrir o mstsc. Conecte manualmente com: mstsc /v:localhost:$rdpPort"
}

try {
    # Tenta abrir no navegador padrão
    Start-Process $vncUrl -ErrorAction Stop
}
catch {
    # Fallback para navegadores específicos
    $browsers = @("chrome.exe", "msedge.exe", "firefox.exe")
    foreach ($browser in $browsers) {
        if (Get-Command $browser -ErrorAction SilentlyContinue) {
            Start-Process $browser -ArgumentList $vncUrl
            break
        }
    }
}

# 7. Mostra informações
Write-Host "`n=====================================================" -ForegroundColor Green
Write-Host "Container iniciado com sucesso!" -ForegroundColor Green
Write-Host "Nome do container: $CONTAINER_NAME" -ForegroundColor Green
Write-Host "VNC (noVNC): $vncUrl" -ForegroundColor Green
Write-Host "RDP: mstsc /v:localhost:$rdpPort" -ForegroundColor Green
Write-Host "`nCredenciais:" -ForegroundColor Yellow
Write-Host "  Usuário: user" -ForegroundColor Yellow
Write-Host "  Senha: password" -ForegroundColor Yellow
Write-Host "`nPara parar o container: podman stop $CONTAINER_NAME" -ForegroundColor Yellow
Write-Host "Para reiniciar: podman start $CONTAINER_NAME" -ForegroundColor Yellow
Write-Host "=====================================================`n" -ForegroundColor Green