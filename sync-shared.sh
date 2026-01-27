#!/bin/bash
#
# 共通コード同期スクリプト
# メインシステムの sharedLib.js を AllUser にコピー
#
# 使用方法:
#   ./sync-shared.sh
#
# AllUserをデプロイする前に必ず実行してください

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SRC_FILE="$SCRIPT_DIR/src/sharedLib.js"
DEST_DIR="$SCRIPT_DIR/AllUser/src"
DEST_FILE="$DEST_DIR/sharedLib.js"

# 色付き出力
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

echo "========================================"
echo "共通コード同期スクリプト"
echo "========================================"

# ソースファイルの存在確認
if [ ! -f "$SRC_FILE" ]; then
    echo -e "${RED}エラー: ソースファイルが見つかりません${NC}"
    echo "  $SRC_FILE"
    exit 1
fi

# 宛先ディレクトリの存在確認
if [ ! -d "$DEST_DIR" ]; then
    echo -e "${RED}エラー: 宛先ディレクトリが見つかりません${NC}"
    echo "  $DEST_DIR"
    exit 1
fi

# 既存ファイルがあればバックアップ
if [ -f "$DEST_FILE" ]; then
    BACKUP_FILE="$DEST_FILE.bak.$(date +%Y%m%d_%H%M%S)"
    cp "$DEST_FILE" "$BACKUP_FILE"
    echo -e "${YELLOW}既存ファイルをバックアップ:${NC}"
    echo "  $BACKUP_FILE"
fi

# コピー実行
cp "$SRC_FILE" "$DEST_FILE"

echo -e "${GREEN}同期完了!${NC}"
echo "  コピー元: $SRC_FILE"
echo "  コピー先: $DEST_FILE"
echo ""
echo "========================================"
echo "次のステップ:"
echo "  1. AllUser プロジェクトを GAS エディタで開く"
echo "  2. sharedLib.js が正しく追加されていることを確認"
echo "  3. デプロイを実行"
echo "========================================"
