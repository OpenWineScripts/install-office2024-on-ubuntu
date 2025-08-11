#!/bin/bash

# --- Colors for better visualization ---
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[0;33m'
NC='\033[0m' # No Color

echo -e "${GREEN}Starting Microsoft Office installation script with Wine...${NC}"

# --- Function to check and install Wine and Winetricks ---
install_wine_winetricks() {
    echo -e "${YELLOW}Checking Wine and Winetricks installation...${NC}"
    if ! command -v wine &> /dev/null || ! command -v winetricks &> /dev/null; then
        echo -e "${YELLOW}Wine or Winetricks not found. Attempting to install...${NC}"

        # Detect the distribution
        if [ -f /etc/os-release ]; then
            # shellcheck disable=SC1091
            . /etc/os-release
            OS=$ID
        elif [[ "$(uname -s)" == "Darwin" ]]; then
            OS="macos"
        else
            echo -e "${RED}Could not detect your Linux distribution. Please install Wine and Winetricks manually.${NC}"
            exit 1
        fi

        case "$OS" in
            "arch" | "manjaro" | "endeavouros")
                echo -e "${YELLOW}Detected: Arch Linux or derivative. Installing via Pacman...${NC}"
                sudo pacman -Sy wine winetricks --noconfirm
                ;;
            "fedora" | "centos" | "rhel")
                echo -e "${YELLOW}Detected: Fedora or derivative. Installing via DNF...${NC}"
                sudo dnf install wine winetricks -y
                ;;
            "debian" | "ubuntu" | "linuxmint")
                echo -e "${YELLOW}Detected: Debian/Ubuntu or derivative. Installing via APT...${NC}"
                sudo dpkg --add-architecture i386 # Enable 32-bit architecture
                sudo apt update
                sudo apt install wine-stable winetricks -y
                ;;
            "opensuse" | "suse")
                echo -e "${YELLOW}Detected: openSUSE. Installing via Zypper...${NC}"
                sudo zypper install wine winetricks -y
                ;;
            "macos")
                echo -e "${YELLOW}Detected: macOS. Installing via Homebrew...${NC}"
                if ! command -v brew &> /dev/null; then
                    echo -e "${RED}Homebrew not found. Please install it first: /bin/bash -c \"$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)\"${NC}"
                    exit 1
                fi
                brew install wine winetricks
                ;;
            *)
                echo -e "${RED}Distribution $OS not supported for automatic Wine/Winetricks installation. Please install them manually.${NC}"
                exit 1
                ;;
        esac
    else
        echo -e "${GREEN}Wine and Winetricks are already installed.${NC}"
    fi
}

# --- Call the installation function ---
install_wine_winetricks

# --- 1. Define the Wine prefix (virtual environment) ---
export WINEPREFIX="$HOME/Documents/office2024"
echo -e "${GREEN}Creating or using 64-bit Wine prefix in: $WINEPREFIX${NC}"

# Create the 64-bit prefix if it doesn't exist
if [ ! -d "$WINEPREFIX" ]; then
    mkdir -p "$WINEPREFIX"
    WINEARCH=win64 winecfg # Creates the prefix as 64-bit
else
    echo -e "${YELLOW}Wine prefix '$WINEPREFIX' already exists.${NC}"
fi

# --- 2. Install necessary dependencies with Winetricks ---
echo -e "${GREEN}Installing Office dependencies with Winetricks...${NC}"
# Dependencies may vary. .NET and Visual C++ are common.
# For Office, you might need: dotnet48 (or another version), vcrun2019, etc.
# Consult WineHQ documentation for specific Office compatibility.
# LATEST UPDATE ON LINE 85: Added msxml6, gdiplus, riched20, wininet
if ! winetricks -q corefonts dotnet48 vcrun2019 msxml6 gdiplus riched20 wininet; then # Corrected SC2181
    echo -e "${RED}Error installing dependencies with Winetricks. Check the output above.${NC}"
    exit 1
else
    echo -e "${GREEN}Dependencies installed successfully.${NC}"
fi

# --- 3. Download and configure Office Deployment Tool (ODT) ---
ODT_DIR="$HOME/Office_ODT"
mkdir -p "$ODT_DIR"
cd "$ODT_DIR" || exit 1 # Corrected SC2164 - exit if cd fails

ODT_URL="https://download.microsoft.com/download/6c1eeb25-cf8b-41d9-8d0d-cc1dbc032140/officedeploymenttool_18827-20140.exe" # ODT URL (may change)
ODT_EXE="office_deployment_tool.exe"

if [ ! -f "$ODT_EXE" ]; then
    echo -e "${YELLOW}Downloading Office Deployment Tool...${NC}"
    if ! wget -O "$ODT_EXE" "$ODT_URL"; then # Corrected SC2181
        echo -e "${RED}Error downloading Office Deployment Tool. Check the URL and your connection.${NC}"
        exit 1
    fi
else
    echo -e "${GREEN}Office Deployment Tool already downloaded.${NC}"
fi

echo -e "${GREEN}Extracting Office Deployment Tool...${NC}"
if ! wine "$ODT_EXE" /extract:"$ODT_DIR" /quiet /noreboot; then # Corrected SC2181
    echo -e "${RED}Error extracting Office Deployment Tool. Check if Wine is working correctly.${NC}"
    exit 1
fi

# --- 4. Generate configuration.xml for download ---
# This XML file defines which Office version and components will be downloaded.
# Example for Office Professional Plus 2021. Adjust as needed.
# For Office 2024, if and when released, the Product ID might be different.
echo -e "${GREEN}Generating configuration.xml for Office download...${NC}"

cat <<EOF > configuration_download.xml
<Configuration>
  <Add SourcePath="." OfficeClientEdition="64" Channel="PerpetualVL2021">
    <Product ID="ProPlus2021Volume">
      <Language ID="pt-br" />
    </Product>
  </Add>
  <Updates Enabled="TRUE" />
  <Display Level="None" AcceptEULA="TRUE" />
  <Logging Level="Standard" Path="%temp%" />
</Configuration>
EOF

# --- 5. Realizar o download offline do Office ---
echo -e "${GREEN}Iniciando download offline do Office... Isso pode levar um tempo considerável.${NC}"
if ! wine "$ODT_DIR/setup.exe" /download "$ODT_DIR/configuration_download.xml"; then # Corrected SC2181
    echo -e "${RED}Error downloading Office files. Check your connection and the configuration.xml.${NC}"
    exit 1
else
    echo -e "${GREEN}Download do Office concluído com sucesso.${NC}"
fi

# --- 6. Gerar configuration.xml para instalação ---
echo -e "${GREEN}Gerando configuration.xml para instalação do Office...${NC}"

cat <<EOF > configuration_install.xml
<Configuration>
  <Add SourcePath="." OfficeClientEdition="64" Channel="PerpetualVL2021">
    <Product ID="ProPlus2021Volume">
      <Language ID="pt-br" />
    </Product>
  </Add>
  <Updates Enabled="TRUE" />
  <Display Level="Full" AcceptEULA="TRUE" />
  <Logging Level="Standard" Path="%temp%" />
</Configuration>
EOF

# --- 7. Executar a instalação offline do Office ---
echo -e "${GREEN}Iniciando instalação offline do Office...${NC}"
if ! wine "$ODT_DIR/setup.exe" /configure "$ODT_DIR/configuration_install.xml"; then # Corrected SC2181
    echo -e "${RED}Error installing Microsoft Office. Check the output above for more details.${NC}"
    exit 1
else
    echo -e "${GREEN}Microsoft Office installed successfully!${NC}"
    echo -e "${GREEN}Você pode encontrar os atalhos no menu de aplicativos do seu ambiente de trabalho ou executar os programas diretamente via Wine.${NC}"
fi

echo -e "${GREEN}Script concluído.${NC}"
