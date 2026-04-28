#!/bin/bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
SOURCE_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
INSTALL_ROOT="${HOME}/Library/Application Support/SlideSCI"
BIN_DIR="$INSTALL_ROOT/bin"
HELPER_DIR="$INSTALL_ROOT/helper"
SERVER_DIR="$INSTALL_ROOT/server"
WEB_DIR="$INSTALL_ROOT/web"
CERT_DIR="$INSTALL_ROOT/certs"
MANIFEST_DIR="${HOME}/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"
LAUNCH_AGENT_DIR="${HOME}/Library/LaunchAgents"
LAUNCH_AGENT_ID="com.slidesci.helper-watcher"
LAUNCH_AGENT_PATH="$LAUNCH_AGENT_DIR/$LAUNCH_AGENT_ID.plist"
LOGIN_KEYCHAIN="${HOME}/Library/Keychains/login.keychain-db"
DEFAULT_LOCAL_ADDIN_HOST="${SLIDESCI_LOCAL_HOST:-127.0.0.1}"
DEFAULT_LOCAL_ADDIN_PORT="${SLIDESCI_LOCAL_PORT:-18443}"
MANIFEST_BASE_URL="${1:-https://${DEFAULT_LOCAL_ADDIN_HOST}:${DEFAULT_LOCAL_ADDIN_PORT}}"
CERT_COMMON_NAME="SlideSCI Local Add-in"
CERT_PATH="$CERT_DIR/slidesci-local-cert.pem"
KEY_PATH="$CERT_DIR/slidesci-local-key.pem"
TMP_MANIFEST="$(mktemp /tmp/slidesci-manifest.XXXXXX.xml)"
TMP_OPENSSL_CONFIG="$(mktemp /tmp/slidesci-openssl.XXXXXX.cnf)"

PACKAGE_HELPER_SCRIPT="$SCRIPT_DIR/helper/native-equation-helper.mjs"
SOURCE_HELPER_SCRIPT="$SOURCE_ROOT/scripts/native-equation-helper.mjs"
PACKAGE_SERVER_SCRIPT="$SCRIPT_DIR/server/local-addin-server.mjs"
SOURCE_SERVER_SCRIPT="$SOURCE_ROOT/scripts/local-addin-server.mjs"
PACKAGE_COMPANION_BIN="$SCRIPT_DIR/bin/SlideSCICompanion"
SOURCE_COMPANION_BIN="$SOURCE_ROOT/bin/SlideSCICompanion"
PACKAGE_MANIFEST="$SCRIPT_DIR/manifest.xml"
SOURCE_MANIFEST="$SOURCE_ROOT/manifest.xml"
PACKAGE_WEB_ROOT="$SCRIPT_DIR/web"
SOURCE_WEB_ROOT="$SOURCE_ROOT/dist"
BUILD_COMPANION_SCRIPT="$SOURCE_ROOT/scripts/build-companion.sh"
RENDER_MANIFEST_SCRIPT="$SOURCE_ROOT/scripts/render-manifest.mjs"

cleanup_tmp_files() {
  rm -f "$TMP_MANIFEST" "$TMP_OPENSSL_CONFIG"
}

trap cleanup_tmp_files EXIT

if command -v xattr >/dev/null 2>&1; then
  xattr -dr com.apple.quarantine "$SCRIPT_DIR" >/dev/null 2>&1 || true
fi

read -r LOCAL_ADDIN_HOST LOCAL_ADDIN_PORT <<EOF
$(python3 - "$MANIFEST_BASE_URL" <<'PY'
import sys
from urllib.parse import urlparse

url = urlparse(sys.argv[1])
host = url.hostname or "127.0.0.1"
port = url.port or 18443
print(host, port)
PY
)
EOF

require_command() {
  local command_name="$1"
  local help_message="$2"
  if ! command -v "$command_name" >/dev/null 2>&1; then
    echo "$help_message"
    exit 1
  fi
}

render_manifest() {
  if [ -f "$RENDER_MANIFEST_SCRIPT" ]; then
    node "$RENDER_MANIFEST_SCRIPT" "$MANIFEST_BASE_URL" "$TMP_MANIFEST" >/dev/null
    return
  fi

  if [ -f "$PACKAGE_MANIFEST" ]; then
    cp "$PACKAGE_MANIFEST" "$TMP_MANIFEST"
    return
  fi

  if [ -f "$SOURCE_MANIFEST" ]; then
    cp "$SOURCE_MANIFEST" "$TMP_MANIFEST"
    return
  fi

  echo "未找到 manifest.xml，也无法动态生成 manifest。"
  exit 1
}

extract_taskpane_url() {
  local manifest_path="$1"
  python3 - "$manifest_path" <<'PY'
import re
import sys
from pathlib import Path

text = Path(sys.argv[1]).read_text(encoding="utf-8")
patterns = [
    r'<SourceLocation[^>]*DefaultValue="([^"]+)"',
    r'<bt:Url[^>]*id="Taskpane\.Url"[^>]*DefaultValue="([^"]+)"',
]
for pattern in patterns:
    match = re.search(pattern, text)
    if match:
        print(match.group(1))
        raise SystemExit(0)
raise SystemExit(1)
PY
}

copy_companion() {
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
}

copy_helper() {
  if [ -f "$PACKAGE_HELPER_SCRIPT" ]; then
    cp "$PACKAGE_HELPER_SCRIPT" "$HELPER_DIR/native-equation-helper.mjs"
  elif [ -f "$SOURCE_HELPER_SCRIPT" ]; then
    cp "$SOURCE_HELPER_SCRIPT" "$HELPER_DIR/native-equation-helper.mjs"
  else
    echo "未找到 native-equation-helper.mjs。"
    exit 1
  fi
}

copy_local_server() {
  if [ -f "$PACKAGE_SERVER_SCRIPT" ]; then
    cp "$PACKAGE_SERVER_SCRIPT" "$SERVER_DIR/local-addin-server.mjs"
  elif [ -f "$SOURCE_SERVER_SCRIPT" ]; then
    cp "$SOURCE_SERVER_SCRIPT" "$SERVER_DIR/local-addin-server.mjs"
  else
    echo "未找到 local-addin-server.mjs。"
    exit 1
  fi
}

ensure_source_web_root() {
  if [ -f "$SOURCE_WEB_ROOT/index.html" ] && [ -d "$SOURCE_WEB_ROOT/assets" ]; then
    return
  fi

  if ! command -v npm >/dev/null 2>&1; then
    echo "源码安装缺少前端构建产物 dist/，且未检测到 npm。"
    echo "请先执行 npm run build。"
    exit 1
  fi

  echo "未检测到 dist/ 前端构建产物，正在执行 npm run build ..."
  (cd "$SOURCE_ROOT" && npm run build >/dev/null)
}

copy_web_assets() {
  rm -rf "$WEB_DIR"
  mkdir -p "$WEB_DIR"

  if [ -f "$PACKAGE_WEB_ROOT/index.html" ] && [ -d "$PACKAGE_WEB_ROOT/assets" ]; then
    cp "$PACKAGE_WEB_ROOT/index.html" "$WEB_DIR/index.html"
    cp -R "$PACKAGE_WEB_ROOT/assets" "$WEB_DIR/assets"
    return
  fi

  ensure_source_web_root
  cp "$SOURCE_WEB_ROOT/index.html" "$WEB_DIR/index.html"
  cp -R "$SOURCE_WEB_ROOT/assets" "$WEB_DIR/assets"
}

generate_local_certificate() {
  if [ -f "$CERT_PATH" ] && [ -f "$KEY_PATH" ]; then
    return
  fi

  cat > "$TMP_OPENSSL_CONFIG" <<EOF
[req]
default_bits = 2048
prompt = no
default_md = sha256
distinguished_name = dn
x509_extensions = v3_req

[dn]
CN = ${CERT_COMMON_NAME}

[v3_req]
subjectAltName = @alt_names
keyUsage = critical, digitalSignature, keyEncipherment
extendedKeyUsage = serverAuth

[alt_names]
DNS.1 = localhost
IP.1 = 127.0.0.1
EOF

  /usr/bin/openssl req \
    -x509 \
    -nodes \
    -days 3650 \
    -newkey rsa:2048 \
    -keyout "$KEY_PATH" \
    -out "$CERT_PATH" \
    -config "$TMP_OPENSSL_CONFIG" >/dev/null 2>&1
}

trust_local_certificate() {
  local cert_fingerprint
  cert_fingerprint="$(/usr/bin/openssl x509 -noout -fingerprint -sha1 -in "$CERT_PATH" | awk -F= '{print $2}' | tr -d ':')"

  if security find-certificate -Z -c "$CERT_COMMON_NAME" "$LOGIN_KEYCHAIN" 2>/dev/null | grep -q "$cert_fingerprint"; then
    return
  fi

  security delete-certificate -c "$CERT_COMMON_NAME" "$LOGIN_KEYCHAIN" >/dev/null 2>&1 || true
  if ! security add-trusted-cert -d -r trustRoot -k "$LOGIN_KEYCHAIN" "$CERT_PATH" >/dev/null; then
    echo "无法将 SlideSCI 本地 HTTPS 证书加入登录钥匙串。"
    echo "请在系统提示中允许证书安装后，再重新运行安装脚本。"
    exit 1
  fi
}

wait_for_local_server() {
  local health_url="$1"
  local attempt
  for attempt in $(seq 1 20); do
    if curl --silent --show-error --insecure --fail --max-time 3 "$health_url" >/dev/null 2>&1; then
      return
    fi
    sleep 0.5
  done

  echo "SlideSCI 本地任务窗格服务未能成功启动：$health_url"
  echo "请查看日志：$INSTALL_ROOT/companion.log"
  exit 1
}

require_command "python3" "未检测到 python3，无法安装 SlideSCI。"
require_command "node" "未检测到 node。当前安装脚本仍需要 node 运行 helper 和本地任务窗格服务。请先安装 Node.js LTS。"
require_command "openssl" "未检测到 openssl，无法生成本地 HTTPS 证书。"
require_command "curl" "未检测到 curl，无法验证本地任务窗格服务。"

NODE_BIN="$(command -v node)"

mkdir -p "$BIN_DIR" "$HELPER_DIR" "$SERVER_DIR" "$CERT_DIR" "$MANIFEST_DIR" "$LAUNCH_AGENT_DIR"

copy_companion
copy_helper
copy_local_server
copy_web_assets
generate_local_certificate
trust_local_certificate
render_manifest
cp "$TMP_MANIFEST" "$MANIFEST_DIR/manifest.xml"

chmod +x "$BIN_DIR/SlideSCICompanion"

if command -v xattr >/dev/null 2>&1; then
  xattr -dr com.apple.quarantine "$INSTALL_ROOT" >/dev/null 2>&1 || true
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
      <string>--web-command</string>
      <string>$NODE_BIN</string>
      <string>--web-script</string>
      <string>$SERVER_DIR/local-addin-server.mjs</string>
      <string>--web-root</string>
      <string>$WEB_DIR</string>
      <string>--host</string>
      <string>$LOCAL_ADDIN_HOST</string>
      <string>--web-port</string>
      <string>$LOCAL_ADDIN_PORT</string>
      <string>--web-cert</string>
      <string>$CERT_PATH</string>
      <string>--web-key</string>
      <string>$KEY_PATH</string>
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

wait_for_local_server "${MANIFEST_BASE_URL}/health"

TASKPANE_URL="$(extract_taskpane_url "$TMP_MANIFEST" || true)"

echo "SlideSCI 已安装。"
echo "1. manifest 已复制到 $MANIFEST_DIR/manifest.xml"
echo "2. 本地任务窗格服务已就绪：${MANIFEST_BASE_URL}/health"
echo "3. companion LaunchAgent 已注册：$LAUNCH_AGENT_PATH"
echo "4. 公式 helper 现在由 companion 常驻托管，本地地址为 http://127.0.0.1:17926"
echo "5. 如果你想临时完全关停 SlideSCI，请双击 stop-slidesci-mac.command；恢复时双击 start-slidesci-mac.command。"
echo "6. 请完全退出并重启 PowerPoint。插件入口现在位于顶部 SlideSCI 选项卡。"
if [ -n "$TASKPANE_URL" ]; then
  echo "7. Taskpane 地址：$TASKPANE_URL"
fi
