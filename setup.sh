#!/bin/bash
# setup.sh — Instala skills e ferramentas Claude Code (Mac/Linux)
# Uso: bash setup.sh

set -e

TOOLS="$HOME/claude-tools"
CMDS="$HOME/.claude/commands"
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

echo ""
echo "=== Claude Dotfiles — Instalando ==="

mkdir -p "$TOOLS" "$CMDS"

# Copia ferramentas
cp -f "$SCRIPT_DIR"/tools/* "$TOOLS/"
echo "  Ferramentas copiadas para: $TOOLS"

# Instala cada skill substituindo placeholder pelo caminho real
for f in "$SCRIPT_DIR"/commands/*.md; do
    name=$(basename "$f")
    sed "s|{{CLAUDE_TOOLS_PATH}}|$TOOLS|g" "$f" > "$CMDS/$name"
    echo "  Skill instalado: $name"
done

echo ""
echo "=== Instalação concluída! ==="
echo "  Para atualizar: git pull && bash setup.sh"
echo ""
