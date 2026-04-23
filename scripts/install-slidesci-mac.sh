#!/bin/bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")/.." && pwd)"
INSTALL_ROOT="${HOME}/Library/Application Support/SlideSCI"
BIN_DIR="$INSTALL_ROOT/bin"
HELPER_DIR="$INSTALL_ROOT/helper"
MANIFEST_DIR="${HOME}/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"
LAUNCH_AGENT_DIR="${HOME}/Library/LaunchAgents"
LAUNCH_AGENT_ID="com.slidesci.helper-watcher"
LAUNCH_AGENT_PATH="$LAUNCH_AGENT_DIR/$LAUNCH_AGENT_ID.plist"
MANIFEST_BASE_URL="${1:-https://localhost:3000}"
TMP_MANIFEST="$(mktemp /tmp/slidesci-manifest.XXXXXX.xml)"
SOURCE_HELPER_SCRIPT="$ROOT_DIR/scripts/native-equation-helper.mjs"
BUNDLED_HELPER_SCRIPT="$ROOT_DIR/helper/native-equation-helper.mjs"
SOURCE_COMPANION_BIN="$ROOT_DIR/bin/SlideSCICompanion"
BUNDLED_MANIFEST="$ROOT_DIR/manifest.xml"

NODE_BIN="$(command -v node || true)"
if [ -z "$NODE_BIN" ]; then
  echo "未检测到 node。当前安装脚本仍需要 node 运行 helper。"
  echo "请先安装 Node.js，或后续将 helper 打包为独立可执行文件。"
  exit 1
fi

mkdir -p "$BIN_DIR" "$HELPER_DIR" "$MANIFEST_DIR" "$LAUNCH_AGENT_DIR"

if [ -x "$SOURCE_COMPANION_BIN" ]; then
  cp "$SOURCE_COMPANION_BIN" "$BIN_DIR/SlideSCICompanion"
else
  bash "$ROOT_DIR/scripts/build-companion.sh" "$BIN_DIR"
fi

if [ -f "$BUNDLED_HELPER_SCRIPT" ]; then
  cp "$BUNDLED_HELPER_SCRIPT" "$HELPER_DIR/native-equation-helper.mjs"
else
  cp "$SOURCE_HELPER_SCRIPT" "$HELPER_DIR/native-equation-helper.mjs"
fi

if [ -f "$BUNDLED_MANIFEST" ]; then
  cp "$BUNDLED_MANIFEST" "$MANIFEST_DIR/manifest.xml"
else
  node "$ROOT_DIR/scripts/render-manifest.mjs" "$MANIFEST_BASE_URL" "$TMP_MANIFEST"
  cp "$TMP_MANIFEST" "$MANIFEST_DIR/manifest.xml"
  rm -f "$TMP_MANIFEST"
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
