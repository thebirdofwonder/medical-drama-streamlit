#!/bin/bash
# 医学ドラマ動画メーカーを、ネット制限のない普通の環境で起動します。

cd "$(dirname "$0")"

# プロキシ設定を外す（Claude API 接続のため）
unset HTTP_PROXY HTTPS_PROXY ALL_PROXY http_proxy https_proxy all_proxy
unset SOCKS_PROXY SOCKS5_PROXY socks_proxy socks5_proxy
unset GIT_HTTP_PROXY GIT_HTTPS_PROXY

echo "========================================"
echo " 医学ドラマ動画メーカー を起動します"
echo "========================================"
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

echo ""
echo "起動後、ブラウザで次を開きます:"
echo "  http://localhost:8501/?reset=1"
echo ""

# 自由文の背景画像フォルダを必ず用意する
mkdir -p "./custom_backgrounds"
if [ ! -f "./custom_backgrounds/使い方.txt" ]; then
  cat > "./custom_backgrounds/使い方.txt" <<'EOF'
【自由文の背景の使い方】
1. 台本に 〈夜の暗い手術室〉 のように書く
2. このフォルダに、同じ名前の画像を置く
   例: 夜の暗い手術室.jpg  /  夜の暗い手術室.png
3. アプリで「最終版（背景あり）」の MP4 を作る
EOF
fi
echo "背景画像フォルダ:"
echo "  $(pwd)/custom_backgrounds"
echo ""
echo "止めるときは、この窓で Ctrl+C を押してください。"
echo ""

# 少し待ってからブラウザを開く（macOS）
if command -v open >/dev/null 2>&1; then
  (sleep 3 && open "http://localhost:8501/?reset=1") &
fi

python3 -m streamlit run app.py \
  --server.port 8501 \
  --server.address localhost \
  --browser.gatherUsageStats false
