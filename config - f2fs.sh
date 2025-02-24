# Instalar dependências necessárias
sudo apt update
sudo apt install -y build-essential flex bison libelf-dev libssl-dev bc



# Copiar a configuração padrão do WSL2
cp Microsoft/config-wsl .config

# Habilitar F2FS no arquivo .config
cat << EOF >> .config
CONFIG_F2FS_FS=y
CONFIG_F2FS_STAT_FS=y
CONFIG_F2FS_FS_XATTR=y
CONFIG_F2FS_FS_POSIX_ACL=y
CONFIG_F2FS_FS_SECURITY=y
CONFIG_F2FS_CHECK_FS=y
CONFIG_F2FS_FS_ENCRYPTION=y
CONFIG_F2FS_IO_TRACE=y
CONFIG_F2FS_FAULT_INJECTION=y
EOF

# Compilar o kernel
make -j$(nproc)

# Copiar o kernel compilado para o Windows
cp arch/x86/boot/bzImage /mnt/c/Users/SEU_USUARIO/kernel

# Criar/atualizar o arquivo .wslconfig no Windows
cat << EOF > /mnt/c/Users/SEU_USUARIO/.wslconfig
[wsl2]
kernel=C:\\Users\\SEU_USUARIO\\kernel
EOF
