# setup.ps1 — Instala skills e ferramentas Claude Code
# Uso: .\setup.ps1
# Requer: Node.js instalado

$ErrorActionPreference = "Stop"

$tools  = "$HOME\claude-tools"
$cmds   = "$HOME\.claude\commands"
$script = $PSScriptRoot

Write-Host ""
Write-Host "=== Claude Dotfiles — Instalando ===" -ForegroundColor Cyan

# Cria diretórios
New-Item -ItemType Directory -Force $tools | Out-Null
New-Item -ItemType Directory -Force $cmds  | Out-Null

# Copia ferramentas (catalog.json, gerar-planilha.js, etc.)
Copy-Item -Force "$script\tools\*" $tools
Write-Host "  Ferramentas copiadas para: $tools" -ForegroundColor Green

# Instala cada skill substituindo o placeholder pelo caminho real
Get-ChildItem "$script\commands\*.md" | ForEach-Object {
    $content = Get-Content $_.FullName -Raw -Encoding UTF8
    $content = $content -replace '\{\{CLAUDE_TOOLS_PATH\}\}', $tools
    $dest = "$cmds\$($_.Name)"
    [System.IO.File]::WriteAllText($dest, $content, [System.Text.Encoding]::UTF8)
    Write-Host "  Skill instalado: $($_.Name)" -ForegroundColor Green
}

Write-Host ""
Write-Host "=== Instalação concluída! ===" -ForegroundColor Cyan
Write-Host "  Skills disponíveis no Claude Code:" -ForegroundColor White
Get-ChildItem "$cmds\*.md" | ForEach-Object { Write-Host "    /$(($_.BaseName))" }
Write-Host ""
Write-Host "  Para atualizar no futuro:" -ForegroundColor DarkGray
Write-Host "    git pull && .\setup.ps1" -ForegroundColor DarkGray
Write-Host ""
