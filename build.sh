#!/bin/zsh

set -e

SCRIPT="main.py"
APP_NAME=$(basename "$SCRIPT" .py)

pyinstaller --clean --noconfirm --onefile --windowed "$SCRIPT"

# Must sign with Developer ID if app is to be distributed to other ARM64 MacOS users
codesign --deep --force --verify --sign 11CB69B5F69442FE499B83CABAC1ABE5156341AF dist/main.app

if [ -f "dist/$APP_NAME" ]; then
    cd dist
    zip -r "${APP_NAME}.zip" "main.app"
    mv "${APP_NAME}.zip" ..
    echo "Build successful"
else
    echo "Build failed"
fi
