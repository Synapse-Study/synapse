#!/usr/bin/env bash
# =============================================================================
# build_app.sh - Synapse macOS Application & DMG Packager
# =============================================================================
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$ROOT_DIR"

# Direct PyInstaller cache to workspace directory
export PYINSTALLER_CONFIG_DIR="$ROOT_DIR/build/pyinstaller_config"
mkdir -p "$PYINSTALLER_CONFIG_DIR"

echo "=================================================="
echo " Starting Synapse Build (macOS .app + .dmg)"
echo " Project Directory: $ROOT_DIR"
echo "=================================================="

# 1. Resolve Python 3 binary (prefer local .venv if available)
if [ -x "$ROOT_DIR/.venv/bin/python3" ]; then
    PYTHON="$ROOT_DIR/.venv/bin/python3"
elif command -v python3 >/dev/null 2>&1; then
    PYTHON="$(command -v python3)"
else
    echo "Error: Python 3 executable could not be found." >&2
    exit 1
fi
echo "Using Python: $PYTHON ($("$PYTHON" --version))"

# 2. Ensure PyInstaller is installed in the target Python environment
if ! "$PYTHON" -m PyInstaller --version >/dev/null 2>&1 && ! "$PYTHON" -m pyinstaller --version >/dev/null 2>&1; then
    echo "PyInstaller not found. Installing pyinstaller via pip..."
    "$PYTHON" -m pip install pyinstaller
fi
PYINSTALLER_MOD="PyInstaller"
if ! "$PYTHON" -m PyInstaller --version >/dev/null 2>&1; then
    PYINSTALLER_MOD="pyinstaller"
fi
PYINSTALLER_VER="$("$PYTHON" -m "$PYINSTALLER_MOD" --version)"
echo "PyInstaller version: $PYINSTALLER_VER (via -m $PYINSTALLER_MOD)"

# 3. Detect App Icon if available
ICON_ARGS=()
if [ -f "$ROOT_DIR/Synapse.app/Contents/Resources/AppIcon.icns" ]; then
    echo "Found macOS AppIcon: $ROOT_DIR/Synapse.app/Contents/Resources/AppIcon.icns"
    ICON_ARGS=(--icon "$ROOT_DIR/Synapse.app/Contents/Resources/AppIcon.icns")
fi

# 4. Resolve Torchvision C extensions & dylibs (fixes operator torchvision::nms missing in EasyOCR)
TV_ARGS=()
TV_DIR="$("$PYTHON" -c 'import torchvision, os; print(os.path.dirname(torchvision.__file__))' 2>/dev/null || true)"
if [ -n "$TV_DIR" ] && [ -d "$TV_DIR" ]; then
    echo "Torchvision package found: $TV_DIR"
    if [ -f "$TV_DIR/_C_stable.so" ]; then
        TV_ARGS+=(--add-binary "$TV_DIR/_C_stable.so:torchvision")
    fi
    if [ -f "$TV_DIR/image_stable.so" ]; then
        TV_ARGS+=(--add-binary "$TV_DIR/image_stable.so:torchvision")
    fi
    if [ -d "$TV_DIR/.dylibs" ]; then
        TV_ARGS+=(--add-data "$TV_DIR/.dylibs:torchvision/.dylibs")
    fi
fi

# 5. Clean prior build artifacts
echo "Cleaning existing dist/ and build/ folders..."
rm -rf dist/Synapse.app dist/Synapse.dmg build/Synapse

# 6. Build dist/Synapse.app with PyInstaller
echo "Building dist/Synapse.app with PyInstaller..."
"$PYTHON" -m "$PYINSTALLER_MOD" \
    --noconfirm \
    --clean \
    --windowed \
    --name "Synapse" \
    --osx-bundle-identifier "com.synapse.app" \
    "${ICON_ARGS[@]}" \
    "${TV_ARGS[@]}" \
    --add-data "gui:gui" \
    --add-data "references:references" \
    --add-data "my-lecture-notes.md:." \
    --collect-all webview \
    --collect-all easyocr \
    --collect-all torchvision \
    --collect-all cv2 \
    --copy-metadata torchvision \
    --copy-metadata torch \
    --copy-metadata easyocr \
    --hidden-import torchvision._C_stable \
    --hidden-import torchvision.image_stable \
    --hidden-import webview \
    --hidden-import supabase \
    --hidden-import easyocr \
    --hidden-import synapse_auth \
    --hidden-import supabase_auth \
    --hidden-import genanki \
    --hidden-import pypdfium2 \
    --hidden-import pptx \
    --hidden-import PIL \
    --hidden-import keyring \
    --hidden-import keyring.backends.macOS \
    anki_study_suite.py

# 7. Post-build verification and fallback symlinks for torchvision dylibs
if [ -d "dist/Synapse.app/Contents/Resources/torchvision" ]; then
    if [ -n "$TV_DIR" ]; then
        cp -f "$TV_DIR"/*.so "dist/Synapse.app/Contents/Resources/torchvision/" 2>/dev/null || true
        if [ -d "$TV_DIR/.dylibs" ]; then
            cp -Rf "$TV_DIR/.dylibs" "dist/Synapse.app/Contents/Resources/torchvision/" 2>/dev/null || true
        fi
    fi
fi

# 8. Verify dist/Synapse.app was generated
if [ ! -d "dist/Synapse.app" ]; then
    echo "Error: dist/Synapse.app build failed!" >&2
    exit 1
fi
echo "Successfully created dist/Synapse.app"

# 8.1 Re-sign app bundle ad-hoc to fix code signatures invalidated by post-build copies
echo "Signing dist/Synapse.app bundle ad-hoc..."
codesign --force --deep --sign - "dist/Synapse.app"
codesign -vvv --deep --strict "dist/Synapse.app"

# 9. Package dist/Synapse.dmg with proper Drag-to-Applications layout
echo "Preparing DMG staging layout..."
DMG_STAGING="dist/dmg_staging"
rm -rf "$DMG_STAGING"
mkdir -p "$DMG_STAGING"

# Copy the signed application
cp -R "dist/Synapse.app" "$DMG_STAGING/"

# Add Applications symlink for standard drag-and-drop install
ln -s /Applications "$DMG_STAGING/Applications"

# Add a one-click helper script that installs to /Applications and clears Gatekeeper quarantine
cat << 'EOF' > "$DMG_STAGING/Install & Launch Synapse.command"
#!/usr/bin/env bash
set -e
DIR="$(cd "$(dirname "$0")" && pwd)"
echo "=================================================="
echo " Installing Synapse to /Applications..."
echo "=================================================="
if [ -d "/Applications/Synapse.app" ]; then
    echo "Removing previous version from /Applications..."
    rm -rf "/Applications/Synapse.app"
fi
cp -R "$DIR/Synapse.app" /Applications/
echo "Clearing macOS download quarantine flags..."
xattr -cr /Applications/Synapse.app 2>/dev/null || true
echo "=================================================="
echo " Installation complete! Launching Synapse..."
echo "=================================================="
open /Applications/Synapse.app
exit 0
EOF
chmod +x "$DMG_STAGING/Install & Launch Synapse.command"

echo "Packaging dist/Synapse.dmg with hdiutil..."
rm -f "dist/Synapse.dmg"
hdiutil create \
    -volname "Synapse" \
    -srcfolder "$DMG_STAGING" \
    -ov \
    -format UDZO \
    "dist/Synapse.dmg"

rm -rf "$DMG_STAGING"

# Sign the DMG container ad-hoc
echo "Signing dist/Synapse.dmg..."
codesign --force --sign - "dist/Synapse.dmg" 2>/dev/null || true

echo "=================================================="
echo " Synapse Build Complete!"
echo " App Bundle: $ROOT_DIR/dist/Synapse.app"
echo " DMG Image:  $ROOT_DIR/dist/Synapse.dmg"
echo "=================================================="
