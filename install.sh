#!/usr/bin/env bash
# MMCall Academy Skill Installer for Linux / MacOS / WSL
set -e

echo "🎓 Instalando MMCall Academy Document Generator Skill..."

SOURCE_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
GEMINI_SKILL_DIR="$HOME/.gemini/config/skills/mmcall-document-generator"
CLAUDE_SKILL_DIR="$HOME/.claude/skills/mmcall-document-generator"

# Install in Gemini
mkdir -p "$GEMINI_SKILL_DIR"
cp -r "$SOURCE_DIR/"* "$GEMINI_SKILL_DIR/"
echo "✅ Instalado con éxito en Gemini Antigravity: $GEMINI_SKILL_DIR"

# Install in Claude (if .claude exists)
if [ -d "$HOME/.claude" ]; then
  mkdir -p "$CLAUDE_SKILL_DIR"
  cp -r "$SOURCE_DIR/"* "$CLAUDE_SKILL_DIR/"
  echo "✅ Instalado con éxito en Claude Code: $CLAUDE_SKILL_DIR"
fi

echo "🚀 ¡Skill instalada correctamente!"
