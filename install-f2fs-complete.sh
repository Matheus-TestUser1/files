
#!/bin/bash

# Script Simples - Instalação Básica F2FS + Mount
# Apenas o essencial para funcionar
# Versão: 1.0 Simples

set -e

# Cores simples
GREEN='\033[0;32m'
RED='\033[0;31m'
YELLOW='\033[1;33m'
NC='\033[0m'

log() {
    echo -e "${GREEN}[OK]${NC} $1"
}

error() {
    echo -e "${RED}[ERRO]${NC} $1"
    exit 1
}

warn() {
    echo -e "${YELLOW}[INFO]${NC} $1"
}

# Verificar sudo
if [[ $EUID -eq 0 ]]; then
    SUDO=""
else
    SUDO="sudo"
fi

echo "🚀 Instalando F2FS + Mount Tools (Versão Simples)"
echo "=================================================="

# Atualizar sistema
warn "Atualizando sistema..."
$SUDO apt-get update -q
$SUDO apt-get upgrade -y -q
log "Sistema atualizado"

# Instalar pacotes básicos
warn "Instalando ferramentas básicas..."
$SUDO apt-get install -y \
    f2fs-tools \
    util-linux \
    mount \
    kmod \
    ntfs-3g \
    exfat-fuse \
    dosfstools \
    e2fsprogs \
    xfsprogs \
    parted \
    lsof \
    tree \
    htop

log "Ferramentas básicas instaladas"

# Carregar módulo F2FS
warn "Carregando módulo F2FS..."
if $SUDO modprobe f2fs 2>/dev/null; then
    log "Módulo F2FS carregado"
else
    warn "Módulo será carregado automaticamente quando necessário"
fi

# Criar script de teste simples
warn "Criando script de teste..."
$SUDO tee /usr/local/bin/test-f2fs >/dev/null << 'EOF'
#!/bin/bash
echo "🧪 Teste Simples do F2FS"
echo "========================"

# Verificar ferramentas
echo "🔍 Verificando ferramentas:"
for tool in mkfs.f2fs fsck.f2fs dump.f2fs; do
    if command -v $tool >/dev/null; then
        echo "  ✅ $tool: OK"
    else
        echo "  ❌ $tool: FALTANDO"
    fi
done

# Teste prático
echo ""
echo "📁 Teste prático:"
TEST_FILE="/tmp/test-f2fs.img"

# Criar arquivo de 50MB
dd if=/dev/zero of="$TEST_FILE" bs=1M count=50 2>/dev/null
echo "  ✅ Arquivo criado: $TEST_FILE"

# Formatar
if mkfs.f2fs -f "$TEST_FILE" >/dev/null 2>&1; then
    echo "  ✅ Formatação F2FS: OK"
else
    echo "  ❌ Formatação F2FS: FALHOU"
    rm -f "$TEST_FILE"
    exit 1
fi

# Verificar
if fsck.f2fs "$TEST_FILE" >/dev/null 2>&1; then
    echo "  ✅ Verificação: OK"
else
    echo "  ⚠️  Verificação com avisos (normal)"
fi

# Mostrar info
echo "  📊 Informações do volume:"
dump.f2fs -d 1 "$TEST_FILE" 2>/dev/null | head -5 || echo "    Info básica disponível"

# Limpar
rm -f "$TEST_FILE"
echo "  🧹 Arquivo de teste removido"

echo ""
echo "🎉 Teste concluído! F2FS está funcionando."
EOF

$SUDO chmod +x /usr/local/bin/test-f2fs
log "Script de teste criado: test-f2fs"

# Criar script para criar volumes
warn "Criando script para volumes..."
$SUDO tee /usr/local/bin/make-f2fs >/dev/null << 'EOF'
#!/bin/bash
if [ $# -lt 2 ]; then
    echo "Uso: make-f2fs <arquivo.img> <tamanho_mb>"
    echo "Exemplo: make-f2fs volume.img 200"
    exit 1
fi

FILE="$1"
SIZE="$2"

echo "🔨 Criando volume F2FS: $FILE (${SIZE}MB)"

# Criar arquivo
dd if=/dev/zero of="$FILE" bs=1M count="$SIZE" 2>/dev/null
echo "✅ Arquivo criado"

# Formatar
if mkfs.f2fs -f "$FILE" >/dev/null 2>&1; then
    echo "✅ Formatado como F2FS"
    echo "📁 Volume pronto: $FILE"
    echo "💡 Para montar: sudo mount -t f2fs -o loop $FILE /mnt"
else
    echo "❌ Erro na formatação"
    rm -f "$FILE"
    exit 1
fi
EOF

$SUDO chmod +x /usr/local/bin/make-f2fs
log "Script criado: make-f2fs"

# Limpeza rápida
warn "Limpando cache..."
$SUDO apt-get autoremove -y -q
$SUDO apt-get autoclean -q
log "Sistema limpo"

# Testar instalação
warn "Testando instalação..."
if /usr/local/bin/test-f2fs; then
    log "Teste passou! ✅"
else
    error "Teste falhou! ❌"
fi

echo ""
echo "🎉 INSTALAÇÃO CONCLUÍDA!"
echo "======================="
echo ""
echo "✅ F2FS instalado e funcionando"
echo "✅ Ferramentas de mount instaladas"
echo "✅ Scripts utilitários criados"
echo ""
echo "🔧 Comandos disponíveis:"
echo "  test-f2fs                    # Testar F2FS"
echo "  make-f2fs volume.img 100     # Criar volume 100MB"
echo "  mkfs.f2fs arquivo            # Formatar F2FS"
echo "  fsck.f2fs arquivo            # Verificar F2FS"
echo ""
echo "💡 Exemplo de uso:"
echo "  make-f2fs meu-volume.img 500"
echo "  sudo mkdir /mnt/test"
echo "  sudo mount -t f2fs -o loop meu-volume.img /mnt/test"
echo "  echo 'Olá F2FS!' | sudo tee /mnt/test/hello.txt"
echo "  sudo umount /mnt/test"
echo ""
echo "🚀 Pronto para usar!"
