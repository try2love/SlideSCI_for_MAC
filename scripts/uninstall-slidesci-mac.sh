#!/bin/bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
INSTALL_ROOT="${HOME}/Library/Application Support/SlideSCI"
COMPANION_APP_PATH="${HOME}/Applications/SlideSCI Companion.app"
MANIFEST_PATH="${HOME}/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/manifest.xml"
LAUNCH_AGENT_PATH="${HOME}/Library/LaunchAgents/com.slidesci.helper-watcher.plist"
LOGIN_KEYCHAIN="${HOME}/Library/Keychains/login.keychain-db"
CERT_COMMON_NAME="SlideSCI Local Add-in"

if command -v xattr >/dev/null 2>&1; then
  xattr -dr com.apple.quarantine "$SCRIPT_DIR" >/dev/null 2>&1 || true
fi

launchctl unload "$LAUNCH_AGENT_PATH" >/dev/null 2>&1 || true
rm -f "$LAUNCH_AGENT_PATH"
rm -f "$MANIFEST_PATH"
rm -rf "$INSTALL_ROOT"
rm -rf "$COMPANION_APP_PATH"
security delete-certificate -c "$CERT_COMMON_NAME" "$LOGIN_KEYCHAIN" >/dev/null 2>&1 || true

echo "SlideSCI 已卸载。请完全退出并重启 PowerPoint。"
