#!/bin/bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")/.." && pwd)"
VERSION="${1:-}"
ADDIN_BASE_URL="${2:-https://try2love.github.io/SlideSCI_for_MAC}"
RELEASE_ROOT="${ROOT_DIR}/dist/release"

if [ -z "$VERSION" ]; then
  echo "用法: bash scripts/package-release.sh <version> [addin_base_url]"
  exit 1
fi

VERSION="${VERSION#v}"
ARTIFACT_NAME="SlideSCI-for-Mac-v${VERSION}"
STAGE_DIR="${RELEASE_ROOT}/${ARTIFACT_NAME}"
HELPER_STAGE_DIR="${STAGE_DIR}/helper"
BIN_STAGE_DIR="${STAGE_DIR}/bin"
RELEASE_NOTES_PATH="${RELEASE_ROOT}/release-notes-v${VERSION}.md"

rm -rf "$STAGE_DIR"
rm -f "$RELEASE_NOTES_PATH"

npm run build >/dev/null
mkdir -p "$HELPER_STAGE_DIR" "$BIN_STAGE_DIR"
bash "$ROOT_DIR/scripts/build-companion.sh" "$BIN_STAGE_DIR"
node "$ROOT_DIR/scripts/render-manifest.mjs" "$ADDIN_BASE_URL" "$STAGE_DIR/manifest.xml"

cp "$ROOT_DIR/scripts/native-equation-helper.mjs" "$HELPER_STAGE_DIR/native-equation-helper.mjs"
cp "$ROOT_DIR/scripts/install-slidesci-mac.sh" "$STAGE_DIR/install-slidesci-mac.sh"
cp "$ROOT_DIR/scripts/uninstall-slidesci-mac.sh" "$STAGE_DIR/uninstall-slidesci-mac.sh"
cp "$ROOT_DIR/scripts/install-slidesci-mac.sh" "$STAGE_DIR/install-slidesci-mac.command"
cp "$ROOT_DIR/scripts/uninstall-slidesci-mac.sh" "$STAGE_DIR/uninstall-slidesci-mac.command"
chmod +x \
  "$STAGE_DIR/install-slidesci-mac.sh" \
  "$STAGE_DIR/uninstall-slidesci-mac.sh" \
  "$STAGE_DIR/install-slidesci-mac.command" \
  "$STAGE_DIR/uninstall-slidesci-mac.command" \
  "$BIN_STAGE_DIR/SlideSCICompanion"

cat > "$STAGE_DIR/README.txt" <<EOF
SlideSCI for Mac v${VERSION}

安装步骤：
1. 请优先双击 install-slidesci-mac.command。
2. 如果你的 Mac 还没有安装 Node.js，请先安装 Node.js LTS：
   https://nodejs.org/
3. 确保 Microsoft PowerPoint 已完全退出。
4. 双击运行 install-slidesci-mac.command。
5. 如果 macOS 阻止运行，请右键该文件，选择“打开”，再确认一次。
6. 按系统提示完成安装。
7. 如果系统弹出权限提示，请允许终端或 Node 控制 Microsoft PowerPoint，并允许辅助功能权限。
8. 安装完成后，重新打开 PowerPoint。
9. 在“视图”选项卡右侧找到 SlideSCI。

说明：
- 本安装包会安装 companion watcher，使 helper 跟随 PowerPoint 启停。
- 当前版本仍需要本机安装 Node.js，以运行 native-equation-helper.mjs。
- 加载项前端地址：${ADDIN_BASE_URL}

卸载：
1. 完全退出 PowerPoint。
2. 双击 uninstall-slidesci-mac.command。
3. 卸载完成后重新打开 PowerPoint。
EOF

node "$ROOT_DIR/scripts/generate-release-notes.mjs" "v${VERSION}" "$ADDIN_BASE_URL" "$RELEASE_NOTES_PATH"

(cd "$RELEASE_ROOT" && /usr/bin/zip -qry "${ARTIFACT_NAME}.zip" "$ARTIFACT_NAME")
(cd "$RELEASE_ROOT" && /usr/bin/shasum -a 256 "${ARTIFACT_NAME}.zip" > "${ARTIFACT_NAME}.sha256")

echo "Release artifacts:"
echo "  ${RELEASE_ROOT}/${ARTIFACT_NAME}.zip"
echo "  ${RELEASE_ROOT}/${ARTIFACT_NAME}.sha256"
echo "  ${RELEASE_NOTES_PATH}"
