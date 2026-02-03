#!/bin/bash
# 将 DocuPilot/manifest.xml 拷贝到本地 Office wef 目录（仅 macOS）
# 用法: ./copy-manifest-to-wef.sh  或  bash copy-manifest-to-wef.sh

set -e

RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m'

log_ok()   { echo -e "${GREEN}$*${NC}"; }
log_warn() { echo -e "${YELLOW}$*${NC}"; }
log_err()  { echo -e "${RED}$*${NC}"; }

# 脚本所在目录即项目根目录（DocuPilot）
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
MANIFEST_SRC="$SCRIPT_DIR/manifest.xml"

if [[ "$(uname)" != "Darwin" ]]; then
    log_err "此脚本仅支持 macOS。"
    exit 1
fi

if [[ ! -f "$MANIFEST_SRC" ]]; then
    log_err "未找到 manifest.xml: $MANIFEST_SRC"
    exit 1
fi

WEF_BASE="$HOME/Library/Containers"
APPS="com.microsoft.Excel com.microsoft.Word com.microsoft.Powerpoint"

log_ok "正在将 manifest.xml 拷贝到 Office wef 目录..."
echo "  源文件: $MANIFEST_SRC"
echo ""

for app in $APPS; do
    wef_dir="$WEF_BASE/$app/Data/Documents/wef"
    mkdir -p "$wef_dir"
    if cp "$MANIFEST_SRC" "$wef_dir/manifest.xml" 2>/dev/null; then
        log_ok "  ✓ $app: $wef_dir"
    else
        log_warn "  - $app: 拷贝失败（可能未安装或无权限）"
    fi
done

echo ""
log_ok "完成。请在 Excel/Word/PowerPoint 中插入 → 我的加载项 → 共享文件夹 中加载 DocuPilot。"
