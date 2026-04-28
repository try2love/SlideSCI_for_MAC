#!/bin/bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
LAUNCH_AGENT_PATH="${HOME}/Library/LaunchAgents/com.slidesci.helper-watcher.plist"

if command -v xattr >/dev/null 2>&1; then
  xattr -dr com.apple.quarantine "$SCRIPT_DIR" >/dev/null 2>&1 || true
fi

launchctl unload "$LAUNCH_AGENT_PATH" >/dev/null 2>&1 || true
pkill -f "SlideSCICompanion" >/dev/null 2>&1 || true
pkill -f "native-equation-helper.mjs" >/dev/null 2>&1 || true
pkill -f "local-addin-server.mjs" >/dev/null 2>&1 || true

echo "SlideSCI 已停止。"
echo "PowerPoint 下次需要重新使用时，请双击 start-slidesci-mac.command。"
