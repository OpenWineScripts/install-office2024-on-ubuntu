#!/bin/bash

# Office 365 installation script using Wine
# For Ubuntu/Debian-based distributions

# Function to print status
print_status() {
    echo -e "\033[1;32m[*]\033[0m $1"
}

# Update system and install dependencies
print_status "Updating system and installing Wine, Winetricks, and required tools..."
sudo dpkg --add-architecture i386
sudo apt update && sudo apt upgrade -y
sudo apt install -y wine64 wine32 winetricks cabextract wget unzip p7zip-full

# Create Wine prefix
WINEPREFIX="$HOME/Documentos/office-2024"
ARCHITECTURE="win64"

print_status "Creating Wine prefix at $WINEPREFIX..."
WINEARCH=$ARCHITECTURE wineboot -u

# Configure required Winetricks dependencies
print_status "Installing required Windows components with Winetricks..."
WINEPREFIX=$WINEPREFIX winetricks -q dotnet48 corefonts fontsmooth=rgb wininet winhttp msxml6 riched20 riched30 urlmon

# Download Office Deployment Tool
print_status "Downloading Office Deployment Tool..."
mkdir -p ~/Documentos/office-2024/office365-installer && cd ~/Documentos/office-2024/office365-installer
wget -O setup.exe https://officecdn.microsoft.com/pr/wsus/setup.exe

# Create configuration XML file
print_status "Creating configuration file..."
cat > configuration.xml <<EOF
<Configuration>
  <Add OfficeClientEdition="64" Channel="Current">
    <Product ID="O365ProPlusRetail">
      <Language ID="en-us" />
    </Product>
  </Add>
  <Display Level="None" AcceptEULA="TRUE" />
  <Property Name="AUTOACTIVATE" Value="1" />
</Configuration>
EOF

# Start Office installation
print_status "Starting Office installation... (this might take a while)"
WINEPREFIX=$WINEPREFIX wine setup.exe /configure configuration.xml

print_status "Installation script completed. You may now run Office apps using Wine."
