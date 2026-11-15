#!/usr/bin/env bash
set -e

ARCH_LABEL="${1:-Universal}"
export MACOSX_DEPLOYMENT_TARGET="10.15"

echo "============================================================================"
echo "🍎 ILBERTech Platform v2.0 — macOS ($ARCH_LABEL) .app + .dmg Derleniyor..."
echo "============================================================================"

cd "$(dirname "$0")"

# 1. Generate macOS .icns icon from ilber_logo.jpg using macOS built-in sips + iconutil
if [ -f "ilber_logo.jpg" ] && command -v sips >/dev/null 2>&1 && command -v iconutil >/dev/null 2>&1; then
    echo "🎨 macOS .icns ikonu oluşturuluyor..."
    rm -rf ilber_logo.iconset
    mkdir -p ilber_logo.iconset
    sips -s format png -z 16 16     ilber_logo.jpg --out ilber_logo.iconset/icon_16x16.png >/dev/null
    sips -s format png -z 32 32     ilber_logo.jpg --out ilber_logo.iconset/icon_16x16@2x.png >/dev/null
    sips -s format png -z 32 32     ilber_logo.jpg --out ilber_logo.iconset/icon_32x32.png >/dev/null
    sips -s format png -z 64 64     ilber_logo.jpg --out ilber_logo.iconset/icon_32x32@2x.png >/dev/null
    sips -s format png -z 128 128   ilber_logo.jpg --out ilber_logo.iconset/icon_128x128.png >/dev/null
    sips -s format png -z 256 256   ilber_logo.jpg --out ilber_logo.iconset/icon_128x128@2x.png >/dev/null
    sips -s format png -z 256 256   ilber_logo.jpg --out ilber_logo.iconset/icon_256x256.png >/dev/null
    sips -s format png -z 512 512   ilber_logo.jpg --out ilber_logo.iconset/icon_256x256@2x.png >/dev/null
    sips -s format png -z 512 512   ilber_logo.jpg --out ilber_logo.iconset/icon_512x512.png >/dev/null
    sips -s format png -z 1024 1024 ilber_logo.jpg --out ilber_logo.iconset/icon_512x512@2x.png >/dev/null
    iconutil -c icns ilber_logo.iconset -o ilber_logo.icns
    rm -rf ilber_logo.iconset
fi

ICON_ARG=""
if [ -f "ilber_logo.icns" ]; then
    ICON_ARG="--icon=ilber_logo.icns"
fi

# 2. Bundle with PyInstaller into macOS .app bundle
echo "⚙️ PyInstaller ile ILBERTech_Platform.app derleniyor..."
rm -rf build dist
python -m PyInstaller --noconfirm --onedir --windowed --name "ILBERTech_Platform" \
    $ICON_ARG \
    --osx-bundle-identifier "com.ilbertech.platform.v2" \
    --add-data "icons:icons" \
    --add-data "ilber_logo.jpg:." \
    --add-data "ilber_logo.ico:." \
    --collect-all "customtkinter" \
    --collect-data "whisper" \
    --collect-data "tiktoken_ext" \
    --collect-data "certifi" \
    --collect-data "pdfplumber" \
    --collect-data "rapidocr_onnxruntime" \
    --hidden-import "ilber_core_utils" \
    --hidden-import "ilber_ui_dialogs" \
    --hidden-import "ilber_license" \
    --hidden-import "torch" \
    --hidden-import "torchaudio" \
    --hidden-import "whisper" \
    --hidden-import "tiktoken" \
    --hidden-import "tiktoken_ext" \
    --hidden-import "tiktoken_ext.openai_public" \
    --hidden-import "transformers" \
    --hidden-import "sentencepiece" \
    --hidden-import "docx" \
    --hidden-import "openpyxl" \
    --hidden-import "fitz" \
    --hidden-import "pymupdf" \
    --hidden-import "srt" \
    --hidden-import "pdfplumber" \
    --hidden-import "sounddevice" \
    --hidden-import "scipy.io.wavfile" \
    --hidden-import "rapidocr_onnxruntime" \
    --hidden-import "onnxruntime" \
    ilber_v2.py

# 3. Ensure Info.plist supports macOS 10.15+ (Catalina, Big Sur, Monterey, Ventura, Sonoma, Sequoia)
PLIST_PATH="dist/ILBERTech_Platform.app/Contents/Info.plist"
if [ -f "$PLIST_PATH" ]; then
    /usr/libexec/PlistBuddy -c "Set :LSMinimumSystemVersion 10.15.0" "$PLIST_PATH" 2>/dev/null || \
    /usr/libexec/PlistBuddy -c "Add :LSMinimumSystemVersion string 10.15.0" "$PLIST_PATH" 2>/dev/null || true
    /usr/libexec/PlistBuddy -c "Set :NSHighResolutionCapable true" "$PLIST_PATH" 2>/dev/null || \
    /usr/libexec/PlistBuddy -c "Add :NSHighResolutionCapable bool true" "$PLIST_PATH" 2>/dev/null || true
fi

# 4. Bundle ffmpeg & ffprobe binaries inside .app/Contents/MacOS
if command -v ffmpeg >/dev/null 2>&1; then
    cp -f "$(command -v ffmpeg)" "dist/ILBERTech_Platform.app/Contents/MacOS/" || true
    chmod +x "dist/ILBERTech_Platform.app/Contents/MacOS/ffmpeg" || true
fi
if command -v ffprobe >/dev/null 2>&1; then
    cp -f "$(command -v ffprobe)" "dist/ILBERTech_Platform.app/Contents/MacOS/" || true
    chmod +x "dist/ILBERTech_Platform.app/Contents/MacOS/ffprobe" || true
fi

chmod +x "dist/ILBERTech_Platform.app/Contents/MacOS/ILBERTech_Platform"

# 5. Stage with `ditto` (CRITICAL: `cp -R` breaks PyInstaller 6 symlinks in .app/Contents/Frameworks and causes the 🚫 icon!)
OUT_DMG="ILBERTech_Platform_Mac_${ARCH_LABEL}_Setup_v2.dmg"
echo "💿 $OUT_DMG oluşturuluyor (ditto ile sembolik bağlar korunuyor)..."
rm -rf dmg_staging "$OUT_DMG"
mkdir -p dmg_staging
ditto "dist/ILBERTech_Platform.app" "dmg_staging/ILBERTech_Platform.app"
ln -s /Applications dmg_staging/Applications

# Clear quarantine attributes and ad-hoc sign inside staging right before creating DMG
xattr -cr "dmg_staging/ILBERTech_Platform.app" || true
codesign --force --deep --sign - "dmg_staging/ILBERTech_Platform.app" || true

hdiutil create -volname "ILBERTech Platform v2.0" \
    -srcfolder dmg_staging \
    -ov -format UDZO \
    "$OUT_DMG"

rm -rf dmg_staging

echo "============================================================================"
echo "✅ TAMAMLANDI: $(pwd)/$OUT_DMG"
echo "============================================================================"
