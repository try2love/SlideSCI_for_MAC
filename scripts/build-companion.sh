#!/bin/bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")/.." && pwd)"
OUTPUT_DIR="${1:-$ROOT_DIR/dist/companion}"
OUTPUT_BIN="$OUTPUT_DIR/SlideSCICompanion"

mkdir -p "$OUTPUT_DIR"
swiftc "$ROOT_DIR/companion/SlideSCICompanion.swift" -o "$OUTPUT_BIN"
echo "Built companion: $OUTPUT_BIN"
