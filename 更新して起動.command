#!/bin/bash
# 最新版を取り込んでから、医学ドラマ動画メーカーを起動します。
# Finder でこのファイルをダブルクリックするか、
# ターミナルで: bash 更新して起動.command

set -e
cd "$(dirname "$0")"

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
  git fetch origin cursor/cloud-agent-1786051249580-ewceu || git fetch origin
  git checkout cursor/cloud-agent-1786051249580-ewceu 2>/dev/null || true
  git pull origin cursor/cloud-agent-1786051249580-ewceu || git pull
  echo ""
  echo "いまの版:"
  git log -1 --oneline || true
  echo ""
else
  echo "注意: このフォルダは Git 管理ではありません。"
  echo "最新版が入っていない可能性があります。"
  echo ""
fi

# app.py に修正版番号があるか確認
if grep -q 'fix-widget-20260809c' app.py 2>/dev/null; then
  echo "OK: 修正版 fix-widget-20260809c が入っています。"
else
  echo "警告: 新しい修正版がまだ入っていないようです。"
  echo "フォルダを間違えているか、pull に失敗している可能性があります。"
fi
echo ""

exec bash "./起動する.command"
