#!/bin/bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
INSTALL_ROOT="${HOME}/Library/Application Support/SlideSCI"
LAUNCH_AGENT_PATH="${HOME}/Library/LaunchAgents/com.slidesci.helper-watcher.plist"
MANIFEST_PATH="${HOME}/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/manifest.xml"
TASKPANE_HEALTH_URL="https://127.0.0.1:18443/health"

if command -v xattr >/dev/null 2>&1; then
  xattr -dr com.apple.quarantine "$SCRIPT_DIR" >/dev/null 2>&1 || true
fi

if [ ! -f "$LAUNCH_AGENT_PATH" ]; then
  echo "未找到 LaunchAgent：$LAUNCH_AGENT_PATH"
  echo "请先运行 install-slidesci-mac.command 安装 SlideSCI。"
  exit 1
fi

if [ ! -f "$MANIFEST_PATH" ]; then
  echo "未找到 PowerPoint manifest：$MANIFEST_PATH"
  echo "请先重新运行 install-slidesci-mac.command。"
  exit 1
fi

launchctl unload "$LAUNCH_AGENT_PATH" >/dev/null 2>&1 || true
launchctl load "$LAUNCH_AGENT_PATH"

for _ in $(seq 1 20); do
  if curl --silent --show-error --insecure --fail --max-time 3 "$TASKPANE_HEALTH_URL" >/dev/null 2>&1; then
    echo "SlideSCI 已启动。"
    echo "本地 taskpane 服务：$TASKPANE_HEALTH_URL"
    echo "如需彻底关停，请双击 stop-slidesci-mac.command。"
    exit 0
  fi
  sleep 0.5
done

echo "SlideSCI 启动失败，请查看日志：$INSTALL_ROOT/companion.log"
exit 1
