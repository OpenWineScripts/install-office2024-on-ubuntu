#!/bin/bash
# Instalação do Office 2024 via Wine em Linux e macOS com download dos arquivos offline
# Este script realiza:
# 1. Instalação do Wine e Winetricks (detectando a distribuição Linux: Arch, openSUSE, Debian/Ubuntu, Fedora)
# 2. Criação do prefixo Wine para o Office
# 3. Instalação das dependências necessárias para o Office via Winetricks
# 4. Geração do arquivo configuration.xml para o Office Deployment Tool (ODT)
# 5. Download dos arquivos offline do Office usando setup.exe /download

set -e

#########################################
# Funções para instalação em distribuições Linux
#########################################

install_arch_dependencies() {
    echo "Distribuição Arch Linux detectada. Atualizando e instalando Wine e Winetricks..."
    sudo pacman -Syu --noconfirm wine winetricks
}

install_debian_dependencies() {
    echo "Distribuição Debian/Ubuntu detectada. Atualizando e instalando Wine e Winetricks..."
    sudo apt update
    sudo apt install -y wine64 wine winetricks
}

install_fedora_dependencies() {
    echo "Distribuição Fedora detectada. Atualizando e instalando Wine e Winetricks..."
    sudo dnf install -y wine winetricks
}

install_opensuse_dependencies() {
    echo "Distribuição openSUSE detectada. Atualizando e instalando Wine e Winetricks..."
    sudo zypper refresh
    sudo zypper install -y wine wine-winetricks
}

install_linux_dependencies() {
    if [ -f /etc/os-release ]; then
        . /etc/os-release
        case "$ID" in
            arch|manjaro)
                install_arch_dependencies
                ;;
            debian|ubuntu)
                install_debian_dependencies
                ;;
            fedora)
                install_fedora_dependencies
                ;;
            opensuse*|suse)
                install_opensuse_dependencies
                ;;
            *)
                echo "Distribuição Linux não reconhecida. Tente instalar manualmente o Wine e o Winetricks."
                exit 1
                ;;
        esac
    else
        echo "Arquivo /etc/os-release não encontrado. Não foi possível identificar a distribuição Linux."
        exit 1
    fi
}

#########################################
# Função para instalação no macOS via Homebrew
#########################################

install_macos_dependencies() {
    echo "Detectado macOS. Verificando Homebrew..."
    if ! command -v brew &> /dev/null; then
        echo "Homebrew não está instalado. Por favor, instale o Homebrew primeiro: https://brew.sh"
        exit 1
    fi
    echo "Atualizando Homebrew e instalando Wine e Winetricks..."
    brew update
    brew install --cask wine-stable
    brew install winetricks
}

#########################################
# Detecção do sistema operacional
#########################################

OS_TYPE=$(uname)
if [ "$OS_TYPE" = "Linux" ]; then
    install_linux_dependencies
elif [ "$OS_TYPE" = "Darwin" ]; then
    install_macos_dependencies
else
    echo "Sistema operacional não suportado para instalação automática do Wine pelo script."
    exit 1
fi

#########################################
# Configuração do prefixo Wine e instalação das dependências do Office
#########################################

# Define o prefixo Wine para manter o Office isolado
WINEPREFIX="$HOME/.wine_office"
WINEARCH="win64"    # Use "win32" se preferir a versão 32-bit

# Cria o prefixo se não existir
if [ ! -d "$WINEPREFIX" ]; then
    echo "Criando Wine prefix em: $WINEPREFIX"
    env WINEARCH=$WINEARCH WINEPREFIX="$WINEPREFIX" wineboot -i
fi

# Lista de dependências recomendadas via Winetricks para o Office
declare -a DEPENDENCIAS=("corefonts" "msxml6" "gdiplus" "dotnet472" "vcrun2017")
echo "Instalando dependências com Winetricks..."
for dep in "${DEPENDENCIAS[@]}"; do
    echo "Instalando $dep ..."
    env WINEPREFIX="$WINEPREFIX" winetricks -q "$dep"
done

#########################################
# Geração do arquivo configuration.xml
#########################################

echo "Gerando arquivo configuration.xml..."
cat <<'EOF' > configuration.xml
<?xml version="1.0" encoding="UTF-8"?>
<Configuration>
  <Add OfficeClientEdition="64" SourcePath="Office2024Offline" Channel="Broad">
    <Product ID="ProPlus2024Retail">
      <Language ID="pt-br" />
    </Product>
  </Add>
  <Display Level="None" AcceptEULA="TRUE" />
  <Property Name="AUTOACTIVATE" Value="1" />
</Configuration>
EOF

echo "Arquivo configuration.xml gerado com sucesso."

#########################################
# Validação da presença do setup.exe (Office Deployment Tool)
#########################################

if [ ! -f "setup.exe" ]; then
    echo "Erro: O arquivo setup.exe (Office Deployment Tool) não foi encontrado."
    echo "Certifique-se de ter baixado e extraído o ODT (setup.exe) para o mesmo diretório deste script."
    exit 1
fi

#########################################
# Download dos arquivos offline do Office
#########################################

echo "Iniciando o download dos arquivos de instalação offline do Office 2024..."
env WINEPREFIX="$WINEPREFIX" wine setup.exe /download configuration.xml

echo "Download concluído. Os arquivos foram salvos na pasta definida em SourcePath (Office2024Offline)."

echo ""
echo "--------------------------------------------"
echo "Próximos Passos:"
echo "1. Verifique se a pasta 'Office2024Offline' contém os arquivos baixados."
echo "2. Para instalar o Office, execute:"
echo "      env WINEPREFIX=\"$WINEPREFIX\" wine setup.exe /configure configuration.xml"
echo "--------------------------------------------"
