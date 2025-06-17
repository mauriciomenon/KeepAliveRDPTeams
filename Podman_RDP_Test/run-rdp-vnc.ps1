# run-rdp-vnc.ps1
# Versão definitiva - PowerShell 7.5.1+

# 1. Força remoção de containers antigos
podman stop rdp-vnc-test 2>$null
podman rm rdp-vnc-test 2>$null

# 2. Construção da imagem (apenas primeira execução)
if (-not (podman images --format "{{.Repository}}" | Select-String "danielguerra/ubuntu-xrdp")) {
    Write-Host "Construindo imagem persistente..." -ForegroundColor Cyan
    @"
FROM ubuntu:22.04

RUN apt-get update && \
    DEBIAN_FRONTEND=noninteractive apt-get install -y --no-install-recommends \
        xrdp \
        xorgxrdp \
        x11vnc \
        novnc \
        websockify \
        lxde-core \
        firefox \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

RUN echo "startlxde" > /etc/xrdp/startwm.sh
RUN adduser xrdp ssl-cert
RUN echo "lxsession" > /etc/skel/.xsession

EXPOSE 3389 6080
CMD ["sh", "-c", "service xrdp start && websockify --web=/usr/share/novnc 6080 localhost:5900"]
"@ | Set-Content -Path Dockerfile -Force

    podman build -t danielguerra/ubuntu-xrdp -f Dockerfile
}
else {
    Write-Host "Usando imagem persistente existente" -ForegroundColor Green
}

# 3. Encontra portas livres dinamicamente
function Get-FreePort {
    $port = Get-Random -Minimum 10000 -Maximum 60000
    while (Get-NetTCPConnection -LocalPort $port -ErrorAction SilentlyContinue) {
        $port++
    }
    return $port
}

$vncPort = Get-FreePort
$rdpPort = Get-FreePort

# 4. Executa o container
podman run -d --name rdp-vnc-test `
    -p ${vncPort}:6080 `
    -p ${rdpPort}:3389 `
    danielguerra/ubuntu-xrdp

# 5. Aguarda inicialização
Start-Sleep -Seconds 5

# 6. Abre as conexões
$vncUrl = "http://localhost:$vncPort/vnc.html"
Start-Process "chrome.exe" -ArgumentList $vncUrl  # Altere para seu navegador
Start-Process "mstsc" -ArgumentList "/v:localhost:$rdpPort"

Write-Host "`nCONEXÕES PRONTAS:" -ForegroundColor Green
Write-Host "RDP: mstsc /v:localhost:$rdpPort"
Write-Host "VNC: $vncUrl"
Write-Host "`nImagem persistente: danielguerra/ubuntu-xrdp" -ForegroundColor Cyan