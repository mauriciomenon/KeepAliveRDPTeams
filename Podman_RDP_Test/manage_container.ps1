<#
.SYNOPSIS
    Gerencia o container RDP/VNC
.DESCRIPTION
    Permite iniciar, parar, reiniciar e ver status do container
.PARAMETER Action
    Ação a ser executada: start, stop, restart, status
#>

param(
    [ValidateSet("start", "stop", "restart", "status")]
    [string]$Action = "status"
)

$CONTAINER_NAME = "rdp-vnc-test"

switch ($Action) {
    "start" {
        podman start $CONTAINER_NAME
        $ip = podman inspect -f '{{range.NetworkSettings.Networks}}{{.IPAddress}}{{end}}' $CONTAINER_NAME
        $ports = podman port $CONTAINER_NAME
        
        Write-Host "`nContainer iniciado!" -ForegroundColor Green
        Write-Host "Endereço: $ip"
        Write-Host "Portas:"
        $ports | ForEach-Object { Write-Host "  $_" }
        
        # Encontra portas mapeadas
        $rdpPort = ($ports | Where-Object { $_ -match ":3389" } -split ":")[0]
        $vncPort = ($ports | Where-Object { $_ -match ":6080" } -split ":")[0]
        
        Write-Host "`nAcesse via RDP: mstsc /v:localhost:$rdpPort" -ForegroundColor Cyan
        Write-Host "Acesse via VNC: http://localhost:$vncPort/vnc.html" -ForegroundColor Cyan
        Write-Host "`nCredenciais:" -ForegroundColor Yellow
        Write-Host "  Usuário: user" -ForegroundColor Yellow
        Write-Host "  Senha: password" -ForegroundColor Yellow
    }
    "stop" {
        podman stop $CONTAINER_NAME
        Write-Host "Container parado" -ForegroundColor Yellow
    }
    "restart" {
        podman restart $CONTAINER_NAME
        Write-Host "Container reiniciado" -ForegroundColor Green
    }
    "status" {
        $status = podman ps -a --filter "name=$CONTAINER_NAME" --format "{{.Status}}"
        if ($status) {
            Write-Host "Status do container $CONTAINER_NAME : $status" -ForegroundColor Cyan
        }
        else {
            Write-Host "Container $CONTAINER_NAME não encontrado" -ForegroundColor Red
        }
    }
}