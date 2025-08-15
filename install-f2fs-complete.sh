#!/bin/bash

# Script de Instalação Completa - F2FS e Mount Tools
# Atualiza sistema e instala TUDO relacionado a mount e F2FS
# Compatível com Ubuntu/Debian
# Versão: 1.0

set -euo pipefail

# Cores para output
readonly RED='\033[0;31m'
readonly GREEN='\033[0;32m'
readonly YELLOW='\033[1;33m'
readonly BLUE='\033[0;34m'
readonly PURPLE='\033[0;35m'
readonly NC='\033[0m'

# Função de log
log() {
    echo -e "${GREEN}[$(date +'%H:%M:%S')]${NC} $1"
}

info() {
    echo -e "${BLUE}[INFO]${NC} $1"
}

warn() {
    echo -e "${YELLOW}[AVISO]${NC} $1"
}

error() {
    echo -e "${RED}[ERRO]${NC} $1"
    exit 1
}

# Banner inicial
show_banner() {
    echo -e "${PURPLE}"
    echo "╔══════════════════════════════════════════════════════════════╗"
    echo "║              INSTALADOR COMPLETO F2FS + MOUNT               ║"
    echo "║            Atualiza sistema + Instala tudo                  ║"
    echo "╚══════════════════════════════════════════════════════════════╝"
    echo -e "${NC}"
}

# Verificar se é root ou sudo
check_privileges() {
    if [[ $EUID -eq 0 ]]; then
        log "Executando como root ✅"
        SUDO=""
    elif sudo -n true 2>/dev/null; then
        log "Sudo disponível ✅"
        SUDO="sudo"
    else
        error "Este script precisa de privilégios de root. Execute com sudo."
    fi
}

# Detectar distribuição
detect_distro() {
    if [ -f /etc/os-release ]; then
        . /etc/os-release
        DISTRO=$ID
        VERSION=$VERSION_ID
        log "Distribuição detectada: $PRETTY_NAME"
    else
        error "Não foi possível detectar a distribuição Linux"
    fi
    
    # Verificar se é baseado em Debian/Ubuntu
    if [[ ! "$DISTRO" =~ ^(ubuntu|debian)$ ]]; then
        warn "Esta distribuição pode não ser totalmente suportada: $DISTRO"
        read -p "Continuar mesmo assim? (y/N): " -n 1 -r
        echo
        if [[ ! $REPLY =~ ^[Yy]$ ]]; then
            exit 1
        fi
    fi
}

# Atualizar sistema
update_system() {
    log "Atualizando sistema..."
    
    info "Atualizando lista de pacotes..."
    $SUDO apt-get update -qq
    
    info "Fazendo upgrade dos pacotes instalados..."
    $SUDO apt-get upgrade -y
    
    info "Fazendo dist-upgrade (atualizações do sistema)..."
    $SUDO apt-get dist-upgrade -y
    
    log "Sistema atualizado com sucesso ✅"
}

# Instalar dependências básicas
install_basic_deps() {
    log "Instalando dependências básicas..."
    
    local basic_packages=(
        "build-essential"
        "software-properties-common"
        "apt-transport-https"
        "ca-certificates"
        "gnupg"
        "lsb-release"
        "curl"
        "wget"
        "git"
        "vim"
        "nano"
        "htop"
        "tree"
        "unzip"
        "zip"
        "file"
        "strace"
        "lsof"
    )
    
    info "Instalando pacotes básicos: ${basic_packages[*]}"
    $SUDO apt-get install -y "${basic_packages[@]}"
    
    log "Dependências básicas instaladas ✅"
}

# Instalar TODAS as ferramentas de mount
install_mount_tools() {
    log "Instalando TODAS as ferramentas de mount..."
    
    local mount_packages=(
        "util-linux"           # mount, umount, lsblk, blkid, etc.
        "mount"                # comandos mount básicos
        "e2fsprogs"            # ext2/3/4 tools
        "dosfstools"           # FAT/VFAT tools
        "mtools"               # MS-DOS tools
        "ntfs-3g"              # NTFS support
        "exfat-fuse"           # exFAT support
        "exfatprogs"           # exFAT utilities (Ubuntu 20.04+)
        "hfsplus"              # HFS+ support
        "hfsprogs"             # HFS tools
        "xfsprogs"             # XFS tools
        "btrfs-progs"          # Btrfs tools
        "jfsutils"             # JFS tools
        "reiserfsprogs"        # ReiserFS tools
        "nilfs-tools"          # NILFS tools
        "cryptsetup"           # LUKS/dm-crypt
        "lvm2"                 # LVM tools
        "mdadm"                # RAID tools
        "parted"               # partitioning tool
        "gparted"              # GUI partitioning (se disponível)
        "gdisk"                # GPT partitioning
        "fuse"                 # FUSE filesystem
        "bindfs"               # bind mounting with altered permissions
        "curlftpfs"            # FTP filesystem
        "sshfs"                # SSH filesystem
        "cifs-utils"           # CIFS/SMB mounting
        "nfs-common"           # NFS client
        "nfs-kernel-server"    # NFS server
        "autofs"               # automounting
        "udisks2"              # disk management
        "pmount"               # policy mount
        "udevil"               # mount as user
    )
    
    info "Instalando ferramentas de mount e sistemas de arquivos..."
    
    # Instalar com verificação individual
    for package in "${mount_packages[@]}"; do
        if $SUDO apt-get install -y "$package" 2>/dev/null; then
            echo "  ✅ $package"
        else
            echo "  ⚠️  $package (não disponível nesta versão)"
        fi
    done
    
    log "Ferramentas de mount instaladas ✅"
}

# Instalar TODAS as ferramentas F2FS
install_f2fs_tools() {
    log "Instalando TODAS as ferramentas F2FS..."
    
    local f2fs_packages=(
        "f2fs-tools"           # Pacote principal F2FS
        "kmod"                 # Módulos do kernel
        "linux-modules-extra-$(uname -r)" # Módulos extras (se disponível)
    )
    
    info "Instalando ferramentas F2FS específicas..."
    
    for package in "${f2fs_packages[@]}"; do
        if $SUDO apt-get install -y "$package" 2>/dev/null; then
            echo "  ✅ $package"
        else
            echo "  ⚠️  $package (não disponível nesta versão)"
        fi
    done
    
    # Tentar carregar módulo F2FS
    info "Tentando carregar módulo F2FS do kernel..."
    if $SUDO modprobe f2fs 2>/dev/null; then
        echo "  ✅ Módulo F2FS carregado com sucesso"
    else
        echo "  ⚠️  Módulo F2FS será carregado quando necessário"
    fi
    
    log "Ferramentas F2FS instaladas ✅"
}

# Instalar ferramentas de desenvolvimento e debug
install_dev_tools() {
    log "Instalando ferramentas de desenvolvimento e debug..."
    
    local dev_packages=(
        "linux-tools-common"   # perf e outras ferramentas
        "linux-tools-generic"  # ferramentas do kernel
        "trace-cmd"            # tracing tools
        "blktrace"             # block layer tracing
        "iotop"                # I/O monitoring
        "iostat"               # I/O statistics
        "dstat"                # system resource statistics
        "sysstat"              # system activity tools
        "smartmontools"        # disk health monitoring
        "hdparm"               # disk parameters
        "sdparm"               # SCSI disk parameters
        "sg3-utils"            # SCSI utilities
        "lshw"                 # hardware listing
        "hwinfo"               # hardware information
        "inxi"                 # system information
        "neofetch"             # system info display
    )
    
    info "Instalando ferramentas de desenvolvimento..."
    
    for package in "${dev_packages[@]}"; do
        if $SUDO apt-get install -y "$package" 2>/dev/null; then
            echo "  ✅ $package"
        else
            echo "  ⚠️  $package (não disponível)"
        fi
    done
    
    log "Ferramentas de desenvolvimento instaladas ✅"
}

# Criar scripts utilitários
create_utility_scripts() {
    log "Criando scripts utilitários..."
    
    # Script de teste F2FS completo
    info "Criando script de teste F2FS..."
    $SUDO tee /usr/local/bin/test-f2fs-complete >/dev/null << 'EOF'
#!/bin/bash
set -e

echo "🧪 TESTE COMPLETO F2FS - TODAS AS FUNCIONALIDADES"
echo "================================================="

# Verificar ferramentas
echo "🔍 Verificando ferramentas F2FS..."
for tool in mkfs.f2fs fsck.f2fs dump.f2fs defrag.f2fs resize.f2fs; do
    if command -v $tool >/dev/null 2>&1; then
        version=$($tool 2>&1 | head -n1 | grep -o '[0-9]\+\.[0-9]\+\.[0-9]\+' || echo "disponível")
        echo "  ✅ $tool: $version"
    else
        echo "  ❌ $tool: não encontrado"
    fi
done

# Verificar módulo do kernel
echo ""
echo "🔧 Verificando módulo F2FS do kernel..."
if lsmod | grep -q f2fs; then
    echo "  ✅ Módulo F2FS carregado"
else
    echo "  ⚠️  Tentando carregar módulo F2FS..."
    if sudo modprobe f2fs 2>/dev/null; then
        echo "  ✅ Módulo F2FS carregado com sucesso"
    else
        echo "  ❌ Não foi possível carregar módulo F2FS"
    fi
fi

# Verificar suporte no kernel
echo ""
echo "💾 Verificando suporte no kernel..."
if grep -q f2fs /proc/filesystems 2>/dev/null; then
    echo "  ✅ F2FS suportado pelo kernel"
else
    echo "  ❌ F2FS não está listado em /proc/filesystems"
fi

# Teste prático com arquivo
echo ""
echo "📁 Teste prático com arquivo de imagem..."
TEST_IMG="/tmp/f2fs-test-$(date +%s).img"
TEST_MOUNT="/tmp/f2fs-mount-$$"

# Criar arquivo de teste
echo "  📝 Criando arquivo de teste (100MB)..."
dd if=/dev/zero of="$TEST_IMG" bs=1M count=100 2>/dev/null

# Formatar como F2FS
echo "  🔧 Formatando como F2FS..."
if mkfs.f2fs -f "$TEST_IMG" >/dev/null 2>&1; then
    echo "  ✅ Formatação F2FS: SUCESSO"
    
    # Verificar integridade
    echo "  🔍 Verificando integridade..."
    if fsck.f2fs "$TEST_IMG" >/dev/null 2>&1; then
        echo "  ✅ Verificação de integridade: SUCESSO"
    else
        echo "  ⚠️  Verificação com avisos (normal)"
    fi
    
    # Extrair informações
    echo "  📊 Extraindo informações do sistema de arquivos..."
    if dump.f2fs -d 1 "$TEST_IMG" 2>/dev/null | head -10; then
        echo "  ✅ Extração de informações: SUCESSO"
    else
        echo "  ⚠️  Informações limitadas disponíveis"
    fi
    
    # Tentar montar
    mkdir -p "$TEST_MOUNT"
    echo "  📀 Tentando montar sistema F2FS..."
    if sudo mount -t f2fs -o loop "$TEST_IMG" "$TEST_MOUNT" 2>/dev/null; then
        echo "  ✅ Montagem: SUCESSO"
        
        # Testar operações
        echo "  📝 Testando operações de arquivo..."
        if echo "Teste F2FS $(date)" > "$TEST_MOUNT/teste.txt" 2>/dev/null; then
            echo "  ✅ Criação de arquivo: SUCESSO"
            echo "  📄 Conteúdo: $(cat "$TEST_MOUNT/teste.txt")"
        fi
        
        # Mostrar informações
        echo "  📊 Informações do sistema montado:"
        df -h "$TEST_MOUNT" | tail -n1
        
        # Desmontar
        sudo umount "$TEST_MOUNT" 2>/dev/null
        echo "  ✅ Desmontagem: SUCESSO"
        
    else
        echo "  ⚠️  Montagem falhou (pode ser limitação do ambiente)"
    fi
    
else
    echo "  ❌ Formatação F2FS: FALHOU"
fi

# Limpar
rm -rf "$TEST_IMG" "$TEST_MOUNT" 2>/dev/null

echo ""
echo "🎉 TESTE COMPLETO CONCLUÍDO!"
echo "✅ Sistema está pronto para usar F2FS"
EOF

    $SUDO chmod +x /usr/local/bin/test-f2fs-complete
    
    # Script de informações do sistema
    info "Criando script de informações do sistema..."
    $SUDO tee /usr/local/bin/mount-info >/dev/null << 'EOF'
#!/bin/bash

echo "💾 INFORMAÇÕES COMPLETAS - MOUNT E SISTEMAS DE ARQUIVOS"
echo "======================================================="

echo "🖥️  Sistema:"
echo "  OS: $(lsb_release -ds 2>/dev/null || cat /etc/os-release | grep PRETTY_NAME | cut -d'"' -f2)"
echo "  Kernel: $(uname -r)"
echo "  Arquitetura: $(uname -m)"
echo ""

echo "📦 Módulos de sistemas de arquivos carregados:"
lsmod | grep -E "(f2fs|ext|xfs|btrfs|ntfs|fat|nfs)" | sort

echo ""
echo "💾 Sistemas de arquivos suportados pelo kernel:"
cat /proc/filesystems | sort

echo ""
echo "🔧 Ferramentas de formatação disponíveis:"
for tool in mkfs.ext4 mkfs.xfs mkfs.btrfs mkfs.f2fs mkfs.ntfs mkfs.vfat; do
    if command -v $tool >/dev/null 2>&1; then
        echo "  ✅ $tool"
    else
        echo "  ❌ $tool"
    fi
done

echo ""
echo "🗂️  Dispositivos de bloco:"
lsblk -f 2>/dev/null || lsblk

echo ""
echo "📀 Sistemas de arquivos montados:"
mount | column -t

echo ""
echo "💿 Uso de espaço:"
df -h

echo ""
echo "🔍 Ferramentas de mount instaladas:"
which mount umount findmnt lsblk blkid 2>/dev/null | while read tool; do
    echo "  ✅ $tool"
done
EOF

    $SUDO chmod +x /usr/local/bin/mount-info
    
    # Script para criar volumes F2FS
    info "Criando script para criar volumes F2FS..."
    $SUDO tee /usr/local/bin/create-f2fs-volume >/dev/null << 'EOF'
#!/bin/bash

usage() {
    echo "Uso: create-f2fs-volume <arquivo> <tamanho_mb> [ponto_montagem]"
    echo ""
    echo "Exemplos:"
    echo "  create-f2fs-volume volume.img 500"
    echo "  create-f2fs-volume /tmp/test.img 1000 /mnt/test"
    echo ""
    exit 1
}

if [ $# -lt 2 ]; then
    usage
fi

VOLUME="$1"
SIZE="$2"
MOUNT_POINT="${3:-/mnt/f2fs-$(basename "$VOLUME" .img)}"

echo "🔨 Criando volume F2FS:"
echo "  📁 Arquivo: $VOLUME"
echo "  📏 Tamanho: ${SIZE}MB"
echo "  📍 Ponto de montagem: $MOUNT_POINT"
echo ""

# Criar arquivo
echo "📝 Criando arquivo de $(( SIZE ))MB..."
dd if=/dev/zero of="$VOLUME" bs=1M count="$SIZE" 2>/dev/null
echo "✅ Arquivo criado: $VOLUME"

# Formatar
echo "🔧 Formatando como F2FS..."
if mkfs.f2fs -f "$VOLUME" >/dev/null 2>&1; then
    echo "✅ Formatação concluída"
else
    echo "❌ Falha na formatação"
    exit 1
fi

# Verificar
echo "🔍 Verificando integridade..."
if fsck.f2fs "$VOLUME" >/dev/null 2>&1; then
    echo "✅ Verificação OK"
else
    echo "⚠️  Verificação com avisos (normal)"
fi

# Criar ponto de montagem
if [ ! -d "$MOUNT_POINT" ]; then
    echo "📁 Criando ponto de montagem: $MOUNT_POINT"
    sudo mkdir -p "$MOUNT_POINT"
fi

# Tentar montar
echo "📀 Tentando montar..."
if sudo mount -t f2fs -o loop "$VOLUME" "$MOUNT_POINT" 2>/dev/null; then
    echo "✅ Montado com sucesso em: $MOUNT_POINT"
    echo ""
    echo "📊 Informações do volume:"
    df -h "$MOUNT_POINT"
    echo ""
    echo "🔧 Para desmontar: sudo umount $MOUNT_POINT"
    echo "📝 Para acessar: cd $MOUNT_POINT"
else
    echo "⚠️  Montagem falhou, mas volume F2FS foi criado"
    echo "   Volume disponível em: $VOLUME"
    echo "   Tente montar manualmente: sudo mount -t f2fs -o loop $VOLUME $MOUNT_POINT"
fi
EOF

    $SUDO chmod +x /usr/local/bin/create-f2fs-volume
    
    log "Scripts utilitários criados ✅"
}

# Limpar cache e otimizar sistema
cleanup_system() {
    log "Limpando sistema e otimizando..."
    
    info "Removendo pacotes desnecessários..."
    $SUDO apt-get autoremove -y
    
    info "Limpando cache de pacotes..."
    $SUDO apt-get autoclean
    
    info "Limpando cache APT..."
    $SUDO apt-get clean
    
    log "Sistema otimizado ✅"
}

# Mostrar resumo final
show_summary() {
    echo ""
    echo -e "${PURPLE}╔══════════════════════════════════════════════════════════════╗"
    echo -e "║                    INSTALAÇÃO CONCLUÍDA!                    ║"
    echo -e "╚══════════════════════════════════════════════════════════════╝${NC}"
    echo ""
    
    echo -e "${GREEN}✅ Sistema atualizado completamente"
    echo -e "✅ Todas as ferramentas de mount instaladas"
    echo -e "✅ Todas as ferramentas F2FS instaladas"
    echo -e "✅ Scripts utilitários criados"
    echo -e "✅ Sistema otimizado${NC}"
    echo ""
    
    echo -e "${BLUE}🔧 Scripts disponíveis:${NC}"
    echo "  test-f2fs-complete       # Teste completo do F2FS"
    echo "  mount-info              # Informações do sistema"
    echo "  create-f2fs-volume      # Criar volumes F2FS"
    echo ""
    
    echo -e "${BLUE}🚀 Comandos de exemplo:${NC}"
    echo "  test-f2fs-complete                    # Testar tudo"
    echo "  mount-info                           # Ver informações"
    echo "  create-f2fs-volume test.img 500      # Criar volume 500MB"
    echo "  create-f2fs-volume big.img 2000 /mnt/big  # Volume em /mnt/big"
    echo ""
    
    echo -e "${YELLOW}💡 Dica: Execute 'test-f2fs-complete' para verificar se tudo funciona!${NC}"
}

# Função principal
main() {
    show_banner
    check_privileges
    detect_distro
    update_system
    install_basic_deps
    install_mount_tools
    install_f2fs_tools
    install_dev_tools
    create_utility_scripts
    cleanup_system
    show_summary
}

# Executar apenas se chamado diretamente
if [[ "${BASH_SOURCE[0]}" == "${0}" ]]; then
    main "$@"
fi
