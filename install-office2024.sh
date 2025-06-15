#!/usr/bin/env bash
set -e

# Caminho do prefixo Wine
PREFIX="$HOME/.wine-office2024"

# Verificar se o Wine está instalado
if ! command -v wine &>/dev/null; then
  echo "Wine não encontrado. Instalando WineHQ Stable 10.0..."

  # Adicionar arquitetura i386
  sudo dpkg --add-architecture i386
  sudo apt update 

  # Adicionar repositório WineHQ 10.0
  sudo mkdir -pm755 /etc/apt/keyrings
  sudo wget -O /etc/apt/keyrings/winehq-archive.key \
    https://dl.winehq.org/wine-builds/winehq.key
  sudo wget -NP /etc/apt/sources.list.d/ \
    https://dl.winehq.org/wine-builds/ubuntu/dists/noble/winehq-noble.sources

  echo "Atualizando pacotes..."
  sudo apt update

  echo "Instalando WineHQ Stable 10.0..."
  sudo apt install --install-recommends winehq-stable

  echo "Wine instalado com sucesso!"
else
  echo "Wine já está instalado. Versão atual:"
  wine --version
fi

# Instalar pacotes essenciais
sudo apt install -y \
  gcc make perl software-properties-common \
  winetricks \
  winbind smbclient \
  libgl1-mesa-glx:i386 libglu1-mesa:i386

# Preparar prefixo Wine 32-bit
export WINEARCH=win32
export WINEPREFIX="$PREFIX"
wineboot -i

# Instalar bibliotecas do Windows via Winetricks
winetricks -q cmd corefonts msxml6 riched20 gdiplus

# Aplicar substituições de DLL
winetricks dlls=riched20,msxml6,gdiplus

# Criar e aplicar configurações de DLL
REGFILE="$PREFIX/officedlloverrides.reg"
cat <<EOF > "$REGFILE"
REGEDIT4

[HKEY_CURRENT_USER\Software\Wine\DllOverrides]
"riched20"="native,builtin"
"msxml6"="native,builtin"
"gdiplus"="native,builtin"
EOF

WINEPREFIX="$PREFIX" wine regedit "$REGFILE"
rm "$REGFILE"

echo
echo "✅ Wine 10.0 e dependências do Office instalados!"
echo
echo "▶ Para instalar o Office, execute:"
echo "   WINEPREFIX=\"$PREFIX\" wine /caminho/para/OfficeSetup.exe"
echo
echo "▶ Para iniciar o Word após a instalação:"
echo "   WINEPREFIX=\"$PREFIX\" wine \"$PREFIX/drive_c/Program Files/Microsoft Office/root/Office16/WINWORD.EXE\""
