#!/bin/bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")/.." && pwd)"
VERSION="${1:-}"
ADDIN_BASE_URL="${2:-https://127.0.0.1:18443}"
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
SERVER_STAGE_DIR="${STAGE_DIR}/server"
WEB_STAGE_DIR="${STAGE_DIR}/web"
RELEASE_NOTES_PATH="${RELEASE_ROOT}/release-notes-v${VERSION}.md"

rm -rf "$STAGE_DIR"
rm -f "$RELEASE_NOTES_PATH"

npm run build >/dev/null
mkdir -p "$HELPER_STAGE_DIR" "$BIN_STAGE_DIR" "$SERVER_STAGE_DIR" "$WEB_STAGE_DIR"

bash "$ROOT_DIR/scripts/build-companion.sh" "$BIN_STAGE_DIR"
node "$ROOT_DIR/scripts/render-manifest.mjs" "$ADDIN_BASE_URL" "$STAGE_DIR/manifest.xml"

cp "$ROOT_DIR/scripts/native-equation-helper.mjs" "$HELPER_STAGE_DIR/native-equation-helper.mjs"
cp "$ROOT_DIR/scripts/local-addin-server.mjs" "$SERVER_STAGE_DIR/local-addin-server.mjs"
cp "$ROOT_DIR/dist/index.html" "$WEB_STAGE_DIR/index.html"
cp -R "$ROOT_DIR/dist/assets" "$WEB_STAGE_DIR/assets"
cp "$ROOT_DIR/scripts/install-slidesci-mac.sh" "$STAGE_DIR/install-slidesci-mac.sh"
cp "$ROOT_DIR/scripts/uninstall-slidesci-mac.sh" "$STAGE_DIR/uninstall-slidesci-mac.sh"
cp "$ROOT_DIR/scripts/start-slidesci-mac.sh" "$STAGE_DIR/start-slidesci-mac.sh"
cp "$ROOT_DIR/scripts/stop-slidesci-mac.sh" "$STAGE_DIR/stop-slidesci-mac.sh"
cp "$ROOT_DIR/scripts/install-slidesci-mac.sh" "$STAGE_DIR/install-slidesci-mac.command"
cp "$ROOT_DIR/scripts/uninstall-slidesci-mac.sh" "$STAGE_DIR/uninstall-slidesci-mac.command"
cp "$ROOT_DIR/scripts/start-slidesci-mac.sh" "$STAGE_DIR/start-slidesci-mac.command"
cp "$ROOT_DIR/scripts/stop-slidesci-mac.sh" "$STAGE_DIR/stop-slidesci-mac.command"

chmod +x \
  "$STAGE_DIR/install-slidesci-mac.sh" \
  "$STAGE_DIR/uninstall-slidesci-mac.sh" \
  "$STAGE_DIR/start-slidesci-mac.sh" \
  "$STAGE_DIR/stop-slidesci-mac.sh" \
  "$STAGE_DIR/install-slidesci-mac.command" \
  "$STAGE_DIR/uninstall-slidesci-mac.command" \
  "$STAGE_DIR/start-slidesci-mac.command" \
  "$STAGE_DIR/stop-slidesci-mac.command" \
  "$BIN_STAGE_DIR/SlideSCICompanion"

cat > "$STAGE_DIR/README.txt" <<EOF
SlideSCI for Mac v${VERSION}

安装步骤：
1. 请优先双击 install-slidesci-mac.command。
2. 如果你的 Mac 还没有安装 Node.js，请先安装 Node.js LTS：
   https://nodejs.org/
3. 确保 Microsoft PowerPoint 已完全退出。
4. 双击运行 install-slidesci-mac.command。
5. 如果 macOS 阻止运行，请优先执行：
   xattr -dr com.apple.quarantine "/你的解压目录/${ARTIFACT_NAME}"
   然后再双击 install-slidesci-mac.command。
   如果仍被阻止，再右键该文件，选择“打开”，再确认一次。
6. 安装脚本会自动：
   - 安装 SlideSCICompanion
   - 安装本地公式 helper
   - 安装本地任务窗格 HTTPS 服务
   - 生成并信任本地 HTTPS 证书
   - 注册 launchd LaunchAgent
   - 复制 manifest.xml 到 PowerPoint 侧载目录
7. 如果系统弹出权限提示，请允许终端或 Node 控制 Microsoft PowerPoint，并允许辅助功能权限。
8. 安装完成后，重新打开 PowerPoint。
9. 在 PowerPoint 顶部找到独立的 SlideSCI 选项卡。

说明：
- 本安装包运行时不依赖 GitHub Pages；任务窗格页面、helper 和 companion 都使用本机资源。
- 当前版本仍需要本机安装 Node.js，以运行 native-equation-helper.mjs 和本地任务窗格服务。
- 默认本地加载地址：${ADDIN_BASE_URL}
- 如需临时完全停用 SlideSCI，而不是卸载，请双击 stop-slidesci-mac.command。
- 如需重新启用，请双击 start-slidesci-mac.command。

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
