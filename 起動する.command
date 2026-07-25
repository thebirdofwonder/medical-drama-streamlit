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
echo "すでに同じアプリが動いている場合は、"
echo "先にそのターミナルで Ctrl+C を押して止めてください。"
echo ""
echo "起動後、ブラウザで次を開いてください:"
echo "  http://localhost:8501"
echo ""

python3 -m streamlit run app.py
