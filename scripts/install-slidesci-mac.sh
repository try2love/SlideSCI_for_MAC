#!/bin/bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
SOURCE_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
INSTALL_ROOT="${HOME}/Library/Application Support/SlideSCI"
BIN_DIR="$INSTALL_ROOT/bin"
HELPER_DIR="$INSTALL_ROOT/helper"
MANIFEST_DIR="${HOME}/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"
LAUNCH_AGENT_DIR="${HOME}/Library/LaunchAgents"
LAUNCH_AGENT_ID="com.slidesci.helper-watcher"
LAUNCH_AGENT_PATH="$LAUNCH_AGENT_DIR/$LAUNCH_AGENT_ID.plist"
MANIFEST_BASE_URL="${1:-https://localhost:3000}"
TMP_MANIFEST="$(mktemp /tmp/slidesci-manifest.XXXXXX.xml)"
PACKAGE_HELPER_SCRIPT="$SCRIPT_DIR/helper/native-equation-helper.mjs"
SOURCE_HELPER_SCRIPT="$SOURCE_ROOT/scripts/native-equation-helper.mjs"
PACKAGE_COMPANION_BIN="$SCRIPT_DIR/bin/SlideSCICompanion"
SOURCE_COMPANION_BIN="$SOURCE_ROOT/bin/SlideSCICompanion"
PACKAGE_MANIFEST="$SCRIPT_DIR/manifest.xml"
SOURCE_MANIFEST="$SOURCE_ROOT/manifest.xml"
BUILD_COMPANION_SCRIPT="$SOURCE_ROOT/scripts/build-companion.sh"
RENDER_MANIFEST_SCRIPT="$SOURCE_ROOT/scripts/render-manifest.mjs"

NODE_BIN="$(command -v node || true)"
if [ -z "$NODE_BIN" ]; then
  echo "未检测到 node。当前安装脚本仍需要 node 运行 helper。"
  echo "请先安装 Node.js，或后续将 helper 打包为独立可执行文件。"
  exit 1
fi

mkdir -p "$BIN_DIR" "$HELPER_DIR" "$MANIFEST_DIR" "$LAUNCH_AGENT_DIR"

if [ -x "$PACKAGE_COMPANION_BIN" ]; then
  cp "$PACKAGE_COMPANION_BIN" "$BIN_DIR/SlideSCICompanion"
elif [ -x "$SOURCE_COMPANION_BIN" ]; then
  cp "$SOURCE_COMPANION_BIN" "$BIN_DIR/SlideSCICompanion"
elif [ -f "$BUILD_COMPANION_SCRIPT" ]; then
  bash "$BUILD_COMPANION_SCRIPT" "$BIN_DIR"
else
  echo "未找到 SlideSCICompanion。"
  echo "如果你是从 GitHub Release 安装，请确认 zip 已完整解压，并保留 bin/ 文件夹。"
  echo "如果你是从源码安装，请确认仓库结构完整。"
  exit 1
fi

if [ -f "$PACKAGE_HELPER_SCRIPT" ]; then
  cp "$PACKAGE_HELPER_SCRIPT" "$HELPER_DIR/native-equation-helper.mjs"
elif [ -f "$SOURCE_HELPER_SCRIPT" ]; then
  cp "$SOURCE_HELPER_SCRIPT" "$HELPER_DIR/native-equation-helper.mjs"
else
  echo "未找到 native-equation-helper.mjs。"
  echo "如果你是从 GitHub Release 安装，请确认 zip 已完整解压，并保留 helper/ 文件夹。"
  echo "如果你是从源码安装，请确认仓库结构完整。"
  exit 1
fi

if [ -f "$PACKAGE_MANIFEST" ]; then
  cp "$PACKAGE_MANIFEST" "$MANIFEST_DIR/manifest.xml"
elif [ -f "$SOURCE_MANIFEST" ]; then
  cp "$SOURCE_MANIFEST" "$MANIFEST_DIR/manifest.xml"
elif [ -f "$RENDER_MANIFEST_SCRIPT" ]; then
  node "$RENDER_MANIFEST_SCRIPT" "$MANIFEST_BASE_URL" "$TMP_MANIFEST"
  cp "$TMP_MANIFEST" "$MANIFEST_DIR/manifest.xml"
  rm -f "$TMP_MANIFEST"
else
  echo "未找到 manifest.xml，也无法动态生成 manifest。"
  exit 1
fi

cat > "$LAUNCH_AGENT_PATH" <<PLIST
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
  <dict>
    <key>Label</key>
    <string>$LAUNCH_AGENT_ID</string>
    <key>ProgramArguments</key>
    <array>
      <string>$BIN_DIR/SlideSCICompanion</string>
      <string>--helper-command</string>
      <string>$NODE_BIN</string>
      <string>--helper-script</string>
      <string>$HELPER_DIR/native-equation-helper.mjs</string>
      <string>--helper-port</string>
      <string>17926</string>
      <string>--poll-interval</string>
      <string>3</string>
      <string>--shutdown-grace-period</string>
      <string>5</string>
      <string>--powerpoint-bundle-id</string>
      <string>com.microsoft.Powerpoint</string>
    </array>
    <key>RunAtLoad</key>
    <true/>
    <key>KeepAlive</key>
    <true/>
    <key>StandardOutPath</key>
    <string>$INSTALL_ROOT/companion.log</string>
    <key>StandardErrorPath</key>
    <string>$INSTALL_ROOT/companion.log</string>
  </dict>
</plist>
PLIST

launchctl unload "$LAUNCH_AGENT_PATH" >/dev/null 2>&1 || true
launchctl load "$LAUNCH_AGENT_PATH"

echo "SlideSCI 已安装。"
echo "1. manifest 已复制到 $MANIFEST_DIR/manifest.xml"
echo "2. companion LaunchAgent 已注册：$LAUNCH_AGENT_PATH"
echo "3. 请完全退出并重启 PowerPoint。插件入口现在位于“视图”选项卡。"
