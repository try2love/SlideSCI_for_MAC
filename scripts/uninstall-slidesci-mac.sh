#!/bin/bash
set -euo pipefail

INSTALL_ROOT="${HOME}/Library/Application Support/SlideSCI"
MANIFEST_PATH="${HOME}/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/manifest.xml"
LAUNCH_AGENT_PATH="${HOME}/Library/LaunchAgents/com.slidesci.helper-watcher.plist"
LOGIN_KEYCHAIN="${HOME}/Library/Keychains/login.keychain-db"
CERT_COMMON_NAME="SlideSCI Local Add-in"

launchctl unload "$LAUNCH_AGENT_PATH" >/dev/null 2>&1 || true
rm -f "$LAUNCH_AGENT_PATH"
rm -f "$MANIFEST_PATH"
rm -rf "$INSTALL_ROOT"
security delete-certificate -c "$CERT_COMMON_NAME" "$LOGIN_KEYCHAIN" >/dev/null 2>&1 || true

echo "SlideSCI 已卸载。请完全退出并重启 PowerPoint。"
