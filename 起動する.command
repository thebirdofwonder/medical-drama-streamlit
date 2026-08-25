#!/bin/bash
# 医学ドラマ動画メーカーを、ネット制限のない普通の環境で起動します。

cd "$(dirname "$0")" || exit 1

# プロキシ設定を外す（ローカル起動を安定させる）
unset HTTP_PROXY HTTPS_PROXY ALL_PROXY http_proxy https_proxy all_proxy
unset SOCKS_PROXY SOCKS5_PROXY socks_proxy socks5_proxy
unset GIT_HTTP_PROXY GIT_HTTPS_PROXY

APP_URL="http://localhost:8501/?reset=1"

echo "========================================"
echo " 医学ドラマ動画メーカー を起動します"
echo "========================================"
echo ""
echo "フォルダ: $(pwd)"
echo ""

# すでに動いている古いアプリを止める（反応しないとき用）
echo "古いアプリが残っていれば止めます…"
if command -v lsof >/dev/null 2>&1; then
  lsof -ti:8501 2>/dev/null | while read -r pid; do
    kill -9 "$pid" 2>/dev/null || true
  done
fi
pkill -f "streamlit run app.py" 2>/dev/null || true
pkill -f "python3 -m streamlit run app.py" 2>/dev/null || true
sleep 1

# 必須ファイル
if [ ! -f "app.py" ]; then
  echo "エラー: app.py がありません。"
  echo "先に「更新して起動.command」で最新版を取得してください。"
  read -r -p "Enter で閉じる…"
  exit 1
fi

# 背景フォルダ
mkdir -p "./custom_backgrounds"
if [ ! -f "./custom_backgrounds/使い方.txt" ]; then
  cat > "./custom_backgrounds/使い方.txt" <<'EOF'
【背景静止画の使い方】
1. 背景画は別途事前に作成する
2. このフォルダ（custom_backgrounds）に .jpg または .png で保存
3. ファイル名は台本の 〈〉 内の文字と同じにする
4. 「最終版（背景あり）」で MP4 を作る
EOF
fi

echo "起動後、ブラウザで次を開きます:"
echo "  $APP_URL"
echo ""
echo "Cursor 内部ブラウザで見る場合（手動）:"
echo "  1. Cursor で Cmd+Shift+P"
echo "  2. Simple Browser: Show を選ぶ"
echo "  3. 上のアドレスを貼り付けて Enter"
echo ""
echo "止めるときは、この窓で Ctrl+C を押してください。"
echo ""

# 少し待ってから外部ブラウザを開く（macOS）
# ※ cursor:// の自動オープンは Cursor が「ファイルが見つからない」エラーを出すため使わない
if command -v open >/dev/null 2>&1; then
  (sleep 4 && open "$APP_URL") &
fi

# streamlit が無い場合の案内
if ! python3 -c "import streamlit" 2>/dev/null; then
  echo "streamlit が入っていません。インストールします…"
  python3 -m pip install -r requirements.txt --user
fi

set +e
# 手動で Simple Browser を使う場合にも通るよう CORS / XSRF をオフ（ローカル専用）
python3 -m streamlit run app.py \
  --server.port 8501 \
  --server.address 127.0.0.1 \
  --server.headless true \
  --server.enableCORS false \
  --server.enableXsrfProtection false \
  --browser.serverAddress localhost \
  --browser.gatherUsageStats false
STATUS=$?
set -e

if [ "$STATUS" -ne 0 ]; then
  echo ""
  echo "========================================"
  echo " 起動に失敗しました（終了コード: $STATUS）"
  echo " 上の赤いエラーメッセージを確認してください。"
  echo "========================================"
  read -r -p "Enter で閉じる…"
  exit "$STATUS"
fi
