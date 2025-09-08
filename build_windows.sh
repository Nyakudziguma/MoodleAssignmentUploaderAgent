#!/bin/bash

# Build Windows executable from Linux

echo "=== Setting up Wine environment ==="
export WINEPREFIX=~/.wine_python
export WINEARCH=win32

# Check if wine32 is installed
if ! dpkg -l | grep -q "wine32"; then
    echo "Installing wine32..."
    sudo dpkg --add-architecture i386
    sudo apt update
    sudo apt install wine32:i386 wine winetricks
fi

# Initialize Wine if not already done
if [ ! -d "$WINEPREFIX" ]; then
    echo "Initializing Wine..."
    wineboot --init
    echo "Installing Windows components..."
    winetricks corefonts vcrun2019 vcrun2017
fi

echo "=== Installing Python ==="
if [ ! -f "python-3.9.13.exe" ]; then
    wget https://www.python.org/ftp/python/3.9.13/python-3.9.13.exe
fi

# Check if Python is already installed in Wine
if [ ! -d "$WINEPREFIX/drive_c/Python39" ]; then
    wine python-3.9.13.exe /quiet InstallAllUsers=1 PrependPath=1
    echo "Python installed in Wine"
else
    echo "Python already installed in Wine"
fi

echo "=== Installing Python packages ==="
wine pip install pyinstaller selenium webdriver-manager

echo "=== Building Windows executable ==="
cd ~/Documents/myaoc_automation

wine ~/.wine_python/drive_c/Python39/python.exe -m PyInstaller \
    --onefile \
    --windowed \
    --name "MoodleUploader" \
    --add-data "config.json;." \
    moodle_uploader.py

if [ $? -eq 0 ]; then
    echo "=== Build successful! ==="
    echo "Windows executable created at: dist/MoodleUploader.exe"
else
    echo "=== Build failed ==="
    exit 1
fi