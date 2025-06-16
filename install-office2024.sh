#!/bin/bash
set -e

# Script to install Microsoft Office 365 using Wine and its dependencies

# Variables - customize as needed
WINEPREFIX="$HOME/Documentos/office-2024"
WINEARCH="win64"
OFFICE_SETUP_URL=""  # User must provide Office 365 offline installer path or URL
OFFICE_INSTALLER_PATH="$WINEPREFIX/office-2024/office2024-installer/setup.exe" # default download path

# Colors for messages
GREEN="\033[0;32m"
YELLOW="\033[1;33m"
RED="\033[0;31m"
NC="\033[0m" # No Color

echo -e "${GREEN}=== Microsoft Office 365 Installation with Wine ===${NC}"

# Helper function for error exit
error_exit() {
    echo -e "${RED}Error: $1${NC}"
    exit 1
}

# Check for sudo privileges
if ! sudo -v >/dev/null 2>&1; then
    error_exit "This script requires sudo privileges. Please run as a user with sudo access."
fi

# Check platform and package manager
if [[ "$(uname)" == "Linux" ]]; then
    if command -v apt >/dev/null 2>&1; then
        PM="apt"
    elif command -v dnf >/dev/null 2>&1; then
        PM="dnf"
    elif command -v pacman >/dev/null 2>&1; then
        PM="pacman"
    else
        error_exit "Unsupported Linux package manager. Please install Wine manually."
    fi
elif [[ "$(uname)" == "Darwin" ]]; then
    if command -v brew >/dev/null 2>&1; then
        PM="brew"
    else
        error_exit "Homebrew not found. Please install Homebrew first: https://brew.sh/"
    fi
else
    error_exit "Unsupported OS: $(uname). This script supports Linux and macOS only."
fi

echo -e "${YELLOW}Using package manager: $PM${NC}"

# Install Wine and dependencies
echo -e "${GREEN}Installing Wine and required dependencies...${NC}"

if [[ "$PM" == "apt" ]]; then
    sudo dpkg --add-architecture i386
    sudo apt update
    sudo apt install -y --install-recommends wine32 wine64 winetricks cabextract p7zip-full fontconfig
elif [[ "$PM" == "dnf" ]]; then
    sudo dnf install -y wine winetricks cabextract p7zip fontconfig
elif [[ "$PM" == "pacman" ]]; then
    sudo pacman -Sy --noconfirm wine winetricks cabextract p7zip fontconfig
elif [[ "$PM" == "brew" ]]; then
    brew update
    # Install wine-stable or wine depending on availability
    if brew info wine-stable >/dev/null 2>&1; then
        brew install wine-stable winetricks cabextract p7zip fontconfig
    else
        brew install wine winetricks cabextract p7zip fontconfig
    fi
fi

# Verify Wine installed
if ! command -v wine &>/dev/null; then
    error_exit "Wine installation failed or Wine command not found."
fi
if ! command -v winetricks &>/dev/null; then
    error_exit "Winetricks installation failed or winetricks command not found."
fi

# Setup Wine prefix and architecture for Office
echo -e "${GREEN}Setting up Wine prefix at ${WINEPREFIX} with architecture ${WINEARCH}${NC}"

export WINEPREFIX=$WINEPREFIX
export WINEARCH=$WINEARCH

if [ ! -d "$WINEPREFIX" ]; then
    echo -e "${YELLOW}Creating new Wine prefix...${NC}"
    wineboot --init
fi

# Install Wine dependencies via winetricks required by Office 365
echo -e "${GREEN}Installing required Wine components for Office 365...${NC}"

WINEPREFIX="$WINEPREFIX" winetricks -q corefonts fontsmooth=rgb msxml6 riched20 riched30 msxml3 atmlib gdiplus vcrun2013 vcrun2017
WINEPREFIX="$WINEPREFIX" winetricks -q -q win10

# Install Mono and Gecko if missing (Wine usually asks on first run)
echo -e "${GREEN}Checking Wine Mono and Gecko installation...${NC}"

MONO_INSTALLED=$(wine --version && WINEPREFIX="$WINEPREFIX" wine reg query "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\App Paths\\mscoree.exe" > /dev/null 2>&1 && echo "yes" || echo "no")
GECKO_INSTALLED=$(wine --version && WINEPREFIX="$WINEPREFIX" wine reg query "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\App Paths\\iexplore.exe" > /dev/null 2>&1 && echo "yes" || echo "no")

if [ "$MONO_INSTALLED" != "yes" ]; then
    echo -e "${YELLOW}Installing Wine Mono...${NC}"
    WINEPREFIX="$WINEPREFIX" winetricks -q mono210
fi

if [ "$GECKO_INSTALLED" != "yes" ]; then
    echo -e "${YELLOW}Installing Wine Gecko...${NC}"
    WINEPREFIX="$WINEPREFIX"winetricks -q gecko37
fi

# Run Office 365 installer via Wine
echo -e "${GREEN}Launching Microsoft Office 365 installer. Please follow the installation wizard.${NC}"
WINEPREFIX="$WINEPREFIX" wine $OFFICE_INSTALLER_PATH /configure configuration.xml

echo -e "${GREEN}Microsoft Office 365 installation completed (or in progress).${NC}"
echo -e "${YELLOW}To run Office applications, use the following command:${NC}"
echo -e "${YELLOW}WINEPREFIX=$WINEPREFIX wine start 'C:\\Program Files\\Microsoft Office\\<program.exe>"
echo -e "${YELLOW}Replace <program.exe> for other Office apps.${NC}"

echo -e "${GREEN}Installation script completed successfully.${NC}"
