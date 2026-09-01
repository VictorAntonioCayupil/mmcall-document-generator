# MMCall Academy Skill Installer for Windows (PowerShell)
# Installs the mmcall-document-generator skill into Claude Code and Gemini Antigravity

$ErrorActionPreference = "Stop"

Write-Host "🎓 Instalando MMCall Academy Document Generator Skill..." -ForegroundColor Cyan

# 1. Rutas de destino
$geminiSkillDir = "$env:USERPROFILE\.gemini\config\skills\mmcall-document-generator"
$claudeSkillDir = "$env:USERPROFILE\.claude\skills\mmcall-document-generator"

$sourceDir = $PSScriptRoot

# 2. Instalar en Gemini Antigravity
if (-not (Test-Path $geminiSkillDir)) {
    New-Item -ItemType Directory -Path $geminiSkillDir -Force | Out-Null
}
Copy-Item -Path "$sourceDir\*" -Destination $geminiSkillDir -Recurse -Force
Write-Host "✅ Instalado con éxito en Gemini Antigravity: $geminiSkillDir" -ForegroundColor Green

# 3. Instalar en Claude Code (si existe la carpeta .claude)
$claudeConfigDir = "$env:USERPROFILE\.claude"
if (Test-Path $claudeConfigDir) {
    if (-not (Test-Path $claudeSkillDir)) {
        New-Item -ItemType Directory -Path $claudeSkillDir -Force | Out-Null
    }
    Copy-Item -Path "$sourceDir\*" -Destination $claudeSkillDir -Recurse -Force
    Write-Host "✅ Instalado con éxito en Claude Code: $claudeSkillDir" -ForegroundColor Green
}

Write-Host "`n🚀 ¡Skill instalada correctamente! Ahora tu IA conoce automáticamente los formatos de MMCall Academy." -ForegroundColor Yellow
