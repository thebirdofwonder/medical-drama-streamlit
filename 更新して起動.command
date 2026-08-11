#!/bin/bash
# 最新版を取り込んでから、医学ドラマ動画メーカーを起動します。
# Finder でこのファイルをダブルクリックするか、
# ターミナルで: bash 更新して起動.command

set -e
cd "$(dirname "$0")"

BRANCH="cursor/cloud-agent-1786051249580-ewceu"
NEED_BUILD="ui-slim-20260811"

echo "========================================"
echo " 最新版を取得して起動します"
echo "========================================"
echo ""
echo "いまのフォルダ:"
pwd
echo ""

if [ ! -f "app.py" ]; then
  echo "エラー: このフォルダに app.py がありません。"
  echo "正しいプロジェクトフォルダで実行してください。"
  read -r -p "Enter で閉じる…"
  exit 1
fi

if [ -d ".git" ]; then
  echo "GitHub から最新版を取得中…"
  echo "  枝（ブランチ）: $BRANCH"
  git fetch origin "$BRANCH" || git fetch origin
  git checkout "$BRANCH"
  git pull origin "$BRANCH"
  echo ""
  echo "いまの版:"
  git log -1 --oneline || true
  echo "いまの枝:"
  git branch --show-current || true
  echo ""
else
  echo "注意: このフォルダは Git 管理ではありません。"
  echo "最新版が入っていない可能性があります。"
  echo ""
fi

# app.py に修正版番号があるか確認（古い main のままだと取り込みが失敗する）
BUILD_LINE="$(grep -E '^APP_BUILD\s*=' app.py | head -1 || true)"
echo "app.py の修正版: $BUILD_LINE"
if echo "$BUILD_LINE" | grep -q "$NEED_BUILD"; then
  echo "OK: 取り込みエラー修正版が入っています。"
else
  echo "========================================"
  echo " 警告: まだ古い app.py のようです。"
  echo " 左の「設定」に 修正版: ui-slim-20260810… が"
  echo " 出ない場合は、別フォルダを開いている可能性があります。"
  echo "========================================"
  echo ""
  read -r -p "このまま起動しますか？ (y/N): " ans
  case "$ans" in
    y|Y|yes|YES) ;;
    *) echo "中止しました。"; read -r -p "Enter で閉じる…"; exit 1 ;;
  esac
fi
echo ""

exec bash "./起動する.command"
