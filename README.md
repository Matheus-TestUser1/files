# 🚀 Comandos One-Liner Prontos

## **1. Container Ubuntu F2FS (Docker):**

```bash
curl -fsSL https://raw.githubusercontent.com/Matheus-TestUser1/files/main/ubuntu-f2fs.sh | bash
```

## **2. Instalação Completa F2FS + Mount (Sistema Local):**

```bash
curl -fsSL https://raw.githubusercontent.com/Matheus-TestUser1/files/main/install-f2fs-complete.sh | sudo bash
```

## **Comando Alternativo (Mais Seguro):**

```bash
# Baixar primeiro, verificar depois executar
curl -o ubuntu-f2fs.sh https://raw.githubusercontent.com/Matheus-TestUser1/files/main/ubuntu-f2fs.sh
chmod +x ubuntu-f2fs.sh
./ubuntu-f2fs.sh
```

## **Uso Após Instalação:**

```bash
# Acessar o container
docker exec -it ubuntu-f2fs bash

# Testar F2FS
docker exec ubuntu-f2fs test-f2fs

# Ver informações do sistema
docker exec ubuntu-f2fs f2fs-info

# Criar volume F2FS de 200MB
docker exec ubuntu-f2fs create-f2fs volume.img 200
```

## **Comandos de Gerenciamento:**

```bash
# Verificar status
docker ps | grep ubuntu-f2fs

# Parar container
docker stop ubuntu-f2fs

# Remover container
docker rm ubuntu-f2fs

# Remover imagem customizada
docker rmi ubuntu-f2fs:custom
```

## **Troubleshooting:**

Se o comando falhar, execute passo a passo:

```bash
# 1. Verificar Docker
docker --version
docker info

# 2. Baixar script
curl -O https://raw.githubusercontent.com/Matheus-TestUser1/files/main/ubuntu-f2fs.sh

# 3. Dar permissão
chmod +x ubuntu-f2fs.sh

# 4. Executar
./ubuntu-f2fs.sh
```

## **Verificação de Funcionamento:**

```bash
# Após execução, testar:
docker exec ubuntu-f2fs bash -c "
echo '=== TESTE COMPLETO F2FS ==='
test-f2fs
echo '=== INFORMAÇÕES SISTEMA ==='
f2fs-info
echo '=== TESTE CONCLUÍDO ==='
"
```
