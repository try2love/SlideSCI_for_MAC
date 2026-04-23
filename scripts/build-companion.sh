#!/bin/bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")/.." && pwd)"
OUTPUT_DIR="${1:-$ROOT_DIR/dist/companion}"
OUTPUT_BIN="$OUTPUT_DIR/SlideSCICompanion"
SWIFT_CACHE_DIR="${SWIFT_MODULECACHE_PATH:-/tmp/swift-module-cache}"
CLANG_CACHE_DIR="${CLANG_MODULE_CACHE_PATH:-/tmp/clang-module-cache}"

mkdir -p "$OUTPUT_DIR"
mkdir -p "$SWIFT_CACHE_DIR" "$CLANG_CACHE_DIR"
env SWIFT_MODULECACHE_PATH="$SWIFT_CACHE_DIR" CLANG_MODULE_CACHE_PATH="$CLANG_CACHE_DIR" \
  swiftc "$ROOT_DIR/companion/SlideSCICompanion.swift" -o "$OUTPUT_BIN"
echo "Built companion: $OUTPUT_BIN"
