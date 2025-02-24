#!/bin/bash
# Samsung Image Extractor
# Script para extrair arquivos de imagens system.img e vendor.img do Samsung

# Definir cores para melhor visualização
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[0;33m'
NC='\033[0m' # No Color

# Função para exibir mensagens
log() {
  echo -e "${GREEN}[INFO]${NC} $1"
}

error() {
  echo -e "${RED}[ERROR]${NC} $1"
  exit 1
}

warn() {
  echo -e "${YELLOW}[WARNING]${NC} $1"
}

# Verificar se está rodando como root
if [ "$EUID" -ne 0 ]; then
  error "Este script precisa ser executado como root (sudo)."
fi

# Verificar dependências
check_dependencies() {
  log "Verificando dependências..."
  
  deps=("simg2img" "mount" "losetup" "file")
  missing=()
  
  for dep in "${deps[@]}"; do
    if ! command -v $dep &> /dev/null; then
      missing+=($dep)
    fi
  done
  
  if [ ${#missing[@]} -ne 0 ]; then
    warn "Dependências ausentes: ${missing[*]}"
    log "Instalando dependências..."
    apt-get update
    apt-get install -y android-tools-fsutils f2fs-tools util-linux
  else
    log "Todas as dependências estão instaladas."
  fi
}

# Converter imagens de sparse para raw
convert_images() {
  if [ ! -f "$1" ]; then
    error "Arquivo $1 não encontrado."
  fi
  
  local basename=$(basename "$1" .img)
  local output="${2}/${basename}.raw.img"
  
  log "Convertendo $1 para raw..."
  simg2img "$1" "$output"
  
  if [ $? -ne 0 ]; then
    # Se falhar, verificar se já é uma imagem raw
    if file "$1" | grep -q "Linux rev 1.0 ext4"; then
      log "$1 já parece ser uma imagem raw. Copiando diretamente..."
      cp "$1" "$output"
    else
      error "Falha ao converter $1"
    fi
  fi
  
  echo "$output"
}

# Montar imagem
mount_image() {
  local img="$1"
  local mount_point="$2"
  
  # Criar ponto de montagem se não existir
  mkdir -p "$mount_point"
  
  # Detectar sistema de arquivos
  local fstype=$(file "$img" | grep -oE 'ext4|f2fs')
  if [ -z "$fstype" ]; then
    # Tentar detectar de outra forma
    if blkid -s TYPE "$img" | grep -q "ext4"; then
      fstype="ext4"
    elif blkid -s TYPE "$img" | grep -q "f2fs"; then
      fstype="f2fs"
    else
      fstype="ext4" # Assumir ext4 como padrão
      warn "Não foi possível detectar o sistema de arquivos. Assumindo ext4."
    fi
  fi
  
  log "Montando $img ($fstype) em $mount_point"
  
  # Tentar montar diretamente com loop
  if mount -t "$fstype" -o loop,ro "$img" "$mount_point"; then
    log "Imagem montada com sucesso."
    return 0
  fi
  
  # Se falhar, tentar com losetup manualmente
  warn "Montagem direta falhou. Tentando com losetup..."
  
  # Encontrar um dispositivo de loop disponível
  local loop_dev=$(losetup -f)
  
  losetup "$loop_dev" "$img"
  if mount -t "$fstype" -o ro "$loop_dev" "$mount_point"; then
    log "Imagem montada com sucesso usando $loop_dev"
    echo "$loop_dev" # Retornar o dispositivo de loop para desmontar depois
    return 0
  else
    losetup -d "$loop_dev"
    error "Falha ao montar $img"
  fi
}

# Copiar arquivos
copy_files() {
  local src="$1"
  local dest="$2"
  
  log "Copiando arquivos de $src para $dest..."
  mkdir -p "$dest"
  
  # Usar rsync se disponível para cópia mais eficiente
  if command -v rsync &> /dev/null; then
    rsync -a "$src/" "$dest/"
  else
    cp -r "$src/"* "$dest/"
  fi
  
  if [ $? -eq 0 ]; then
    log "Arquivos copiados com sucesso."
  else
    error "Falha ao copiar arquivos."
  fi
}

# Função principal
main() {
  if [ $# -lt 5 ]; then
    echo "Uso: sudo $0 --system /caminho/para/system.img --vendor /caminho/para/vendor.img --output /pasta/destino"
    exit 1
  fi
  
  local system_img=""
  local vendor_img=""
  local output_dir=""
  
  # Processar argumentos
  while [[ "$#" -gt 0 ]]; do
    case $1 in
      --system) system_img="$2"; shift ;;
      --vendor) vendor_img="$2"; shift ;;
      --output) output_dir="$2"; shift ;;
      *) error "Parâmetro desconhecido: $1" ;;
    esac
    shift
  done
  
  # Verificar se todos os parâmetros necessários foram fornecidos
  if [ -z "$system_img" ] || [ -z "$vendor_img" ] || [ -z "$output_dir" ]; then
    error "Todos os parâmetros (--system, --vendor, --output) são obrigatórios."
  fi
  
  # Verificar se os arquivos existem
  if [ ! -f "$system_img" ]; then
    error "Arquivo system.img não encontrado: $system_img"
  fi
  
  if [ ! -f "$vendor_img" ]; then
    error "Arquivo vendor.img não encontrado: $vendor_img"
  fi
  
  # Criar diretório temporário
  local temp_dir=$(mktemp -d)
  log "Diretório temporário: $temp_dir"
  
  # Verificar dependências
  check_dependencies
  
  # Criar diretório de saída
  mkdir -p "$output_dir"
  mkdir -p "$output_dir/system"
  mkdir -p "$output_dir/vendor"
  
  # Converter imagens
  local system_raw=$(convert_images "$system_img" "$temp_dir")
  local vendor_raw=$(convert_images "$vendor_img" "$temp_dir")
  
  # Montar imagens
  local system_mount="$temp_dir/system_mount"
  local vendor_mount="$temp_dir/vendor_mount"
  
  local system_loop=$(mount_image "$system_raw" "$system_mount")
  local vendor_loop=$(mount_image "$vendor_raw" "$vendor_mount")
  
  # Copiar arquivos
  copy_files "$system_mount" "$output_dir/system"
  copy_files "$vendor_mount" "$output_dir/vendor"
  
  # Desmontar e limpar
  log "Limpando recursos..."
  
  umount "$system_mount" || warn "Falha ao desmontar $system_mount"
  umount "$vendor_mount" || warn "Falha ao desmontar $vendor_mount"
  
  if [ -n "$system_loop" ]; then
    losetup -d "$system_loop" || warn "Falha ao liberar $system_loop"
  fi
  
  if [ -n "$vendor_loop" ]; then
    losetup -d "$vendor_loop" || warn "Falha ao liberar $vendor_loop"
  fi
  
  # Remover arquivos temporários
  rm -f "$system_raw" "$vendor_raw"
  rmdir "$system_mount" "$vendor_mount" || true
  rmdir "$temp_dir" || true
  
  log "Processo completo! Arquivos extraídos para: $output_dir"
  log "Pasta system: $output_dir/system"
  log "Pasta vendor: $output_dir/vendor"
}

# Executar função principal com todos os argumentos
main "$@"