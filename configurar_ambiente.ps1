$repo = Get-Location

Write-Host "Configurando ambiente do projeto Python..." -ForegroundColor Cyan

Copy-Item "$PSScriptRoot\.pylintrc" -Destination $repo -Force
Copy-Item "$PSScriptRoot\.flake8" -Destination $repo -Force
Copy-Item "$PSScriptRoot\.editorconfig" -Destination $repo -Force

if (-not (Test-Path "$repo\.git")) {
    git init
    Write-Host "Repositório Git inicializado." -ForegroundColor Green
}

git config core.autocrlf input
git config core.eol lf
Write-Host "Git configurado com fim de linha LF." -ForegroundColor Green

Write-Host "Ambiente configurado com sucesso!" -ForegroundColor Green
