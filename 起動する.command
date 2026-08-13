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

echo "起動後、次のアドレスで開きます:"
echo "  $APP_URL"
echo "  （Cursor 内部ブラウザを優先。だめなら外部ブラウザ）"
echo ""
echo "止めるときは、この窓で Ctrl+C を押してください。"
echo ""

open_in_cursor_internal_browser() {
  local url="$1"
  local enc=""
  enc="$(python3 -c "import urllib.parse,sys; print(urllib.parse.quote(sys.argv[1], safe=''))" "$url" 2>/dev/null)" || return 1
  # Cursor / VS Code の Simple Browser（内部ブラウザ）を開く
  if open "cursor://vscode.simple-browser/show?url=${enc}" 2>/dev/null; then
    return 0
  fi
  if open "vscode://vscode.simple-browser/show?url=${enc}" 2>/dev/null; then
    return 0
  fi
  return 1
}

# 少し待ってからブラウザを開く（macOS）
(
  sleep 4
  if open_in_cursor_internal_browser "$APP_URL"; then
    echo "Cursor 内部ブラウザで開く操作を送りました: $APP_URL"
  elif command -v open >/dev/null 2>&1; then
    echo "内部ブラウザを開けなかったので、外部ブラウザで開きます: $APP_URL"
    open "$APP_URL"
  fi
) &

# streamlit が無い場合の案内
if ! python3 -c "import streamlit" 2>/dev/null; then
  echo "streamlit が入っていません。インストールします…"
  python3 -m pip install -r requirements.txt --user
fi

set +e
# Cursor 内部ブラウザ（iframe）でも動くよう CORS / XSRF を明示的にオフ
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
