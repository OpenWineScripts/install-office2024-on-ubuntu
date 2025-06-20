#!/bin/bash
# Office 2024 installation via Wine for Linux and macOS with offline download support
# This script performs the following:
# 1. Detects and installs Wine and Winetricks (supports Arch, Debian/Ubuntu, Fedora, openSUSE, and macOS)
# 2. Creates a Wine prefix specifically for Office
# 3. Installs necessary Office dependencies using Winetricks
# 4. Generates configuration.xml for Office Deployment Tool (ODT)
# 5. Downloads offline Office setup files using setup.exe /download

set -e

#########################################
# Functions for Linux-based distributions
#########################################

install_arch_dependencies() {
    echo "Detected Arch Linux. Updating and installing Wine and Winetricks..."
    sudo pacman -Syu --noconfirm wine winetricks
}

install_debian_dependencies() {
    echo "Detected Debian/Ubuntu. Updating and installing Wine and Winetricks..."
    sudo apt update
    sudo apt install -y wine64 wine32 winetricks
}

install_fedora_dependencies() {
    echo "Detected Fedora. Installing Wine and Winetricks..."
    sudo dnf install -y wine winetricks
}

install_opensuse_dependencies() {
    echo "Detected openSUSE. Installing Wine and Winetricks..."
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
                echo "Unsupported Linux distribution. Please install Wine and Winetricks manually."
                exit 1
                ;;
        esac
    else
        echo "/etc/os-release not found. Unable to detect Linux distribution."
        exit 1
    fi
}

#########################################
# Function for macOS using Homebrew
#########################################

install_macos_dependencies() {
    echo "Detected macOS. Checking for Homebrew..."
    if ! command -v brew &> /dev/null; then
        echo "Homebrew is not installed. Please install it from https://brew.sh"
        exit 1
    fi
    echo "Installing Wine and Winetricks via Homebrew..."
    brew update
    brew install --cask wine-stable
    brew install winetricks
}

#########################################
# OS detection and setup
#########################################

OS_TYPE=$(uname)
if [ "$OS_TYPE" = "Linux" ]; then
    install_linux_dependencies
elif [ "$OS_TYPE" = "Darwin" ]; then
    install_macos_dependencies
else
    echo "Unsupported OS for automated Wine installation."
    exit 1
fi

##########################################################
# Path creation, Wine prefix setup and Office dependencies
##########################################################

WINEPREFIX="$HOME/Documents/wine_office"
WINEARCH="win64"

if [ ! -d "$WINEPREFIX" ]; then
    echo "Creating the directory path of variable $WINEPREFIX, if doesn't exists..."
    mkdir "$WINEPREFIX"
fi

if [ ! -d "$WINEPREFIX" ]; then
    echo "Creating Wine prefix at: $WINEPREFIX"
    env WINEARCH=$WINEARCH WINEPREFIX="$WINEPREFIX" wineboot -i
fi

DEPENDENCIES=("corefonts" "msxml6" "gdiplus" "dotnet472" "vcrun2017")
echo "Installing Office dependencies via Winetricks..."
for dep in "${DEPENDENCIES[@]}"; do
    echo "Installing $dep ..."
    env WINEPREFIX="$WINEPREFIX" winetricks -q "$dep"
done

############################################################################
# Create the folder Office2024Offline and generate configuration.xml for ODT
############################################################################

if [ ! -d "$WINEPREFIX/drive_c/Office2024Offline" ]; then
    echo "Creating the folder Office2024Offline, it doesn't exists..."
    mkdir "$WINEPREFIX/drive_c/Office2024Offline"
fi

echo "Generating configuration.xml..."
cat <<'EOF' > $WINEPREFIX/drive_c/Office2024Offline/configuration.xml
<?xml version="1.0" encoding="UTF-8"?>
<Configuration>
  <Add OfficeClientEdition="64" SourcePath="C:\Office2024Offline" Channel="Broad">
    <Product ID="ProPlus2024Retail">
      <Language ID="pt-br" />
    </Product>
  </Add>
  <Display Level="Full" AcceptEULA="TRUE" />
  <Property Name="AUTOACTIVATE" Value="1" />
</Configuration>
EOF

echo "configuration.xml successfully generated."

###################
# Extract setup.exe
###################

cp $HOME/Downloads/office_deployment_tool.exe $WINEPREFIX/drive_c/Office2024Offline/office_deployment_tool.exe
env WINEPREFIX="$WINEPREFIX" wine $WINEPREFIX/drive_c/Office2024Offline/office_deployment_tool.exe /extract:$WINEPREFIX/drive_c/Office2024Offline

#########################################
# Check for setup.exe and download Office
#########################################

if [ ! -f "$HOME/Documents/wine_office/drive_c/Office2024Offline/setup.exe" ]; then
    echo "Error: setup.exe (Office Deployment Tool) not found in the current directory."
    echo "Please download and extract it from Microsoft's website before proceeding."
    exit 1
fi

echo "Downloading Office offline setup files..."
env WINEPREFIX="$WINEPREFIX" wine setup.exe /download configuration.xml

echo "Download completed. Files saved in the folder specified in configuration.xml (Office2024Offline)."

echo ""
echo "--------------------------------------------"
echo "Next Steps:"
echo "1. Verify the 'Office2024Offline' folder contains the downloaded files."
echo "2. To install Office, run:"
echo "      env WINEPREFIX=\"$WINEPREFIX\" wine setup.exe /configure configuration.xml"
echo "--------------------------------------------"
