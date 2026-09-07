#!/bin/bash
# ============================================================
#  PDF 处理工具 - macOS 打包脚本（产出 .dmg）
#  用法: bash build_dmg.sh
#  产物: dist_mac/PDF处理工具.dmg
# ============================================================
set -euo pipefail
cd "$(dirname "$0")"

APP_NAME="PDF处理工具"

# 优先使用项目内 venv（避免污染系统 Python）
if [ -x ".venv/bin/python" ]; then
    PY=".venv/bin/python"
else
    PY="python3"
fi

echo "[1/4] 安装/检查依赖..."
"$PY" -m pip install -q --no-cache-dir -r requirements.txt pyinstaller \
    || echo "[警告] pip 安装失败，将尝试使用现有环境继续构建"

echo "[2/4] PyInstaller 构建 .app ..."
rm -rf build_mac dist_mac
"$PY" -m PyInstaller --noconfirm --clean --distpath dist_mac \
    --workpath build_mac pdf_tool_mac.spec

APP="dist_mac/${APP_NAME}.app"
if [ ! -d "$APP" ]; then
    echo "构建失败：找不到 $APP" >&2
    exit 1
fi

echo "[3/4] 生成 .dmg ..."
DMG_STAGE="build_mac/dmg_stage"
rm -rf "$DMG_STAGE"
mkdir -p "$DMG_STAGE"
cp -R "$APP" "$DMG_STAGE/"
ln -s /Applications "$DMG_STAGE/Applications"
rm -f "dist_mac/${APP_NAME}.dmg"
hdiutil create -volname "$APP_NAME" \
    -srcfolder "$DMG_STAGE" \
    -ov -format UDZO \
    "dist_mac/${APP_NAME}.dmg"

echo "[4/4] 完成！"
echo "产物: $(pwd)/dist_mac/${APP_NAME}.dmg"
echo
echo "提示：首次打开会提示“无法验证开发者”，请在 系统设置 -> 隐私与安全性 中点击“仍要打开”。"
