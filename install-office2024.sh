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
winetricks -q corefonts dotnet48 vcrun2019 msxml6 gdiplus riched20 wininet # Example dependencies

if [ $? -ne 0 ]; then
    echo -e "${RED}Error installing dependencies with Winetricks. Check the output above.${NC}"
    exit 1
else
    echo -e "${GREEN}Dependencies installed successfully.${NC}"
fi

# --- 3. Download and configure Office Deployment Tool (ODT) ---
ODT_DIR="$HOME/Office_ODT"
mkdir -p "$ODT_DIR"
cd "$ODT_DIR"

ODT_URL="https://download.microsoft.com/download/6/4/3/643A248A-024F-4C87-8FCE-8F8D98BB2117/officedeploymenttool_16709-20000.exe" # ODT URL (may change)
ODT_EXE="officedeploymenttool.exe"

if [ ! -f "$ODT_EXE" ]; then
    echo -e "${YELLOW}Downloading Office Deployment Tool...${NC}"
    wget -O "$ODT_EXE" "$ODT_URL"
    if [ $? -ne 0 ]; then
        echo -e "${RED}Error downloading Office Deployment Tool. Check the URL and your connection.${NC}"
        exit 1
    fi
else
    echo -e "${GREEN}Office Deployment Tool already downloaded.${NC}"
fi

echo -e "${GREEN}Extracting Office Deployment Tool...${NC}"
wine "$ODT_EXE" /extract:"$ODT_DIR" /quiet /noreboot

if [ $? -ne 0 ]; then
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

# --- 5. Perform offline Office download ---
echo -e "${GREEN}Starting offline Office download... This may take a considerable amount of time.${NC}"
wine "$ODT_DIR/setup.exe" /download "$ODT_DIR/configuration_download.xml"

if [ $? -ne 0 ]; then
    echo -e "${RED}Error downloading Office files. Check your connection and the configuration.xml.${NC}"
    exit 1
else
    echo -e "${GREEN}Office download completed successfully.${NC}"
fi

# --- 6. Generate configuration.xml for installation ---
echo -e "${GREEN}Generating configuration.xml for Office installation...${NC}"

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

# --- 7. Execute offline Office installation ---
echo -e "${GREEN}Starting offline Office installation...${NC}"
wine "$ODT_DIR/setup.exe" /configure "$ODT_DIR/configuration_install.xml"

if [ $? -ne 0 ]; then
    echo -e "${RED}Error installing Microsoft Office. Check the output above for more details.${NC}"
    exit 1
else
    echo -e "${GREEN}Microsoft Office installed successfully!${NC}"
    echo -e "${GREEN}You can find shortcuts in your desktop environment's application menu or run programs directly via Wine.${NC}"
fi

echo -e "${GREEN}Script completed.${NC}"