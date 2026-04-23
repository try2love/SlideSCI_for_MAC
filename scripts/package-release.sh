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

rm -rf "$STAGE_DIR"

npm run build >/dev/null
mkdir -p "$HELPER_STAGE_DIR" "$BIN_STAGE_DIR"
bash "$ROOT_DIR/scripts/build-companion.sh" "$BIN_STAGE_DIR"
node "$ROOT_DIR/scripts/render-manifest.mjs" "$ADDIN_BASE_URL" "$STAGE_DIR/manifest.xml"

cp "$ROOT_DIR/scripts/native-equation-helper.mjs" "$HELPER_STAGE_DIR/native-equation-helper.mjs"
cp "$ROOT_DIR/scripts/install-slidesci-mac.sh" "$STAGE_DIR/install-slidesci-mac.sh"
cp "$ROOT_DIR/scripts/uninstall-slidesci-mac.sh" "$STAGE_DIR/uninstall-slidesci-mac.sh"
chmod +x "$STAGE_DIR/install-slidesci-mac.sh" "$STAGE_DIR/uninstall-slidesci-mac.sh" "$BIN_STAGE_DIR/SlideSCICompanion"

cat > "$STAGE_DIR/README.txt" <<EOF
SlideSCI for Mac v${VERSION}

安装步骤：
1. 确保 PowerPoint 已完全退出。
2. 双击或在终端运行 install-slidesci-mac.sh。
3. 完全重启 PowerPoint。
4. 在“视图”选项卡右侧找到 SlideSCI。

说明：
- 本安装包会安装 companion watcher，使 helper 跟随 PowerPoint 启停。
- 当前版本仍需要本机安装 Node.js，以运行 native-equation-helper.mjs。
- 加载项前端地址：${ADDIN_BASE_URL}
EOF

(cd "$RELEASE_ROOT" && /usr/bin/zip -qry "${ARTIFACT_NAME}.zip" "$ARTIFACT_NAME")
(cd "$RELEASE_ROOT" && /usr/bin/shasum -a 256 "${ARTIFACT_NAME}.zip" > "${ARTIFACT_NAME}.sha256")

echo "Release artifacts:"
echo "  ${RELEASE_ROOT}/${ARTIFACT_NAME}.zip"
echo "  ${RELEASE_ROOT}/${ARTIFACT_NAME}.sha256"
