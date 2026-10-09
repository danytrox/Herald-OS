#!/usr/bin/env bash
set -euo pipefail
cd "$HOME/Herald-OS/node_modules/electron"

rm -rf dist path.txt electron-v44.5.1-win32-x64.zip
URL="https://github.com/electron/electron/releases/download/v44.5.1/electron-v44.5.1-win32-x64.zip"
SHA="9b382492dcfee91f8f9e92c91f7972550a1b95d2299cac72279dab33a600d7db"

echo "==> downloading electron v44.5.1 win32-x64 ($(date))"
curl -fL --retry 4 --retry-delay 3 --connect-timeout 25 -o electron-v44.5.1-win32-x64.zip "$URL"
echo "==> download done ($(date)), size:"; ls -la electron-v44.5.1-win32-x64.zip

echo "==> verifying sha256"
echo "$SHA  electron-v44.5.1-win32-x64.zip" | sha256sum -c -

echo "==> extracting (Expand-Archive)"
powershell.exe -NoProfile -Command "Expand-Archive -LiteralPath 'electron-v44.5.1-win32-x64.zip' -DestinationPath 'dist' -Force" 2>&1 | head -20

echo "==> writing path.txt"
printf 'electron.exe' > path.txt

echo "==> moving electron.d.ts up if present"
[ -f dist/electron.d.ts ] && mv -f dist/electron.d.ts electron.d.ts

rm -f electron-v44.5.1-win32-x64.zip

echo "==> verify"
ls -la dist/electron.exe dist/version path.txt
echo "DONE"
