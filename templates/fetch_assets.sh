#!/usr/bin/env bash
# ============================================================================
# fetch_assets.sh
#   index.html を「フォント/ライブラリ自己ホスト」構成で動かすための初回セットアップ。
#   インターネットに接続できる環境で一度だけ実行してください。
#   static/vendor/ と static/fonts/ に必要なファイルを配置します。
#
#   使い方:  bash fetch_assets.sh         (main.py と同じ階層で実行)
# ============================================================================
set -euo pipefail

ROOT="$(cd "$(dirname "$0")" && pwd)"
VENDOR="$ROOT/static/vendor"
FONTS="$ROOT/static/fonts"
mkdir -p "$VENDOR" "$FONTS"

# --- バージョン（必要に応じて固定/更新） ---
ALPINE_VER="3.14.8"
XLSX_VER="0.20.3"

# --- ホスト（ベースURLは変数に分離） ---
JD="https://cdn.jsdelivr.net"
XLSX_HOST="https://cdn.sheetjs.com"
FS="$JD/fontsource/fonts"

dl () {  # dl <url> <dest>
  local url="$1" dest="$2"
  echo "  -> $(basename "$dest")"
  if command -v curl >/dev/null 2>&1; then
    curl -fsSL --retry 3 "$url" -o "$dest"
  elif command -v wget >/dev/null 2>&1; then
    wget -q "$url" -O "$dest"
  else
    echo "ERROR: curl も wget も見つかりません" >&2; exit 1
  fi
  if [ ! -s "$dest" ]; then echo "ERROR: 空ファイル $dest" >&2; exit 1; fi
}

echo "[1/2] JS ライブラリを取得 -> static/vendor/"
dl "$JD/npm/alpinejs@${ALPINE_VER}/dist/cdn.min.js"             "$VENDOR/alpine.min.js"
dl "$XLSX_HOST/xlsx-${XLSX_VER}/package/dist/xlsx.full.min.js"  "$VENDOR/xlsx.full.min.js"

echo "[2/2] Web フォント (woff2) を取得 -> static/fonts/"
# Inter (latin)
dl "$FS/inter@latest/latin-400-normal.woff2"               "$FONTS/inter-400.woff2"
dl "$FS/inter@latest/latin-500-normal.woff2"               "$FONTS/inter-500.woff2"
dl "$FS/inter@latest/latin-600-normal.woff2"               "$FONTS/inter-600.woff2"
dl "$FS/inter@latest/latin-700-normal.woff2"               "$FONTS/inter-700.woff2"
# Cormorant Garamond (latin)
dl "$FS/cormorant-garamond@latest/latin-500-normal.woff2"  "$FONTS/cormorant-garamond-500.woff2"
dl "$FS/cormorant-garamond@latest/latin-600-normal.woff2"  "$FONTS/cormorant-garamond-600.woff2"
dl "$FS/cormorant-garamond@latest/latin-700-normal.woff2"  "$FONTS/cormorant-garamond-700.woff2"
# Noto Sans JP (japanese subset)
dl "$FS/noto-sans-jp@latest/japanese-400-normal.woff2"     "$FONTS/noto-sans-jp-400.woff2"
dl "$FS/noto-sans-jp@latest/japanese-500-normal.woff2"     "$FONTS/noto-sans-jp-500.woff2"
dl "$FS/noto-sans-jp@latest/japanese-700-normal.woff2"     "$FONTS/noto-sans-jp-700.woff2"

echo ""
echo "完了: static/vendor/(2ファイル) と static/fonts/(10ファイル) を配置しました。"
echo "index.html は /static/vendor/*.js, /static/fonts/*.woff2 を参照します。"
