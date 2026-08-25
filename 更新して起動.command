#!/bin/bash
# 最新版を取り込んでから、医学ドラマ動画メーカーを起動します。
# Finder でこのファイルをダブルクリックするか、
# ターミナルで: bash "更新して起動.command"

cd "$(dirname "$0")" || exit 1

BRANCH="cursor/cloud-agent-1786051249580-ewceu"
NEED_BUILD="ui-slim-202608"

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
  echo "例: cd ~/Documents/Cursor/medical-drama-app"
  read -r -p "Enter で閉じる…"
  exit 1
fi

# Git の失敗で起動まで進まない、という状態を防ぐ
if [ -d ".git" ]; then
  echo "GitHub から最新版を取得中…"
  echo "  枝（ブランチ）: $BRANCH"
  echo ""

  # 使い方.txt など、自動生成ファイルのローカル差分で pull が止まらないようにする
  git checkout -- "custom_backgrounds/使い方.txt" 2>/dev/null || true
  git restore -- "custom_backgrounds/使い方.txt" 2>/dev/null || true

  # ローカルだけの変更が pull を邪魔することがあるので退避
  if ! git diff --quiet 2>/dev/null || ! git diff --cached --quiet 2>/dev/null; then
    echo "注意: このフォルダに未保存の変更があります。"
    echo "  いったん退避（stash）してから最新版を取ります。"
    git stash push -u -m "auto-stash-before-update-$(date +%Y%m%d-%H%M%S)" 2>/dev/null || true
  fi

  git fetch origin "$BRANCH" 2>&1 || git fetch origin 2>&1 || echo "警告: fetch に失敗しました（ネット確認）"
  git checkout "$BRANCH" 2>&1 || echo "警告: checkout に失敗しました"
  if ! git pull --ff-only origin "$BRANCH" 2>&1; then
    echo "警告: 早送り pull に失敗しました。使い方.txt を捨てて再試行します…"
    git checkout -- "custom_backgrounds/使い方.txt" 2>/dev/null || true
    if ! git pull --ff-only origin "$BRANCH" 2>&1; then
      echo "警告: まだ失敗したので、remote に強制的に合わせます…"
      git fetch origin "$BRANCH" 2>&1 || true
      git reset --hard "origin/$BRANCH" 2>&1 || echo "警告: reset にも失敗しました"
    fi
  fi

  echo ""
  echo "いまの版:"
  git log -1 --oneline 2>/dev/null || true
  echo "いまの枝:"
  git branch --show-current 2>/dev/null || true
  echo ""

  # pull でこのスクリプト自身が新しくなっても、いま動いているのは古い版のまま。
  # 一度だけやり直して、新しい起動スクリプトで続きを実行する。
  if [ "${MDS_REEXEC_AFTER_PULL:-}" != "1" ]; then
    export MDS_REEXEC_AFTER_PULL=1
    echo "最新の起動スクリプトでやり直します…"
    echo ""
    exec bash "$0" "$@"
  fi
else
  echo "注意: このフォルダは Git 管理ではありません。"
  echo "最新版が入っていない可能性があります。"
  echo ""
fi

# 必須ファイル確認
if [ ! -f "app.py" ]; then
  echo "========================================"
  echo " エラー: app.py がありません。"
  echo " 最新版の取得に失敗している可能性があります。"
  echo "========================================"
  read -r -p "Enter で閉じる…"
  exit 1
fi

BUILD_LINE="$(grep -E '^APP_BUILD\s*=' app.py | head -1 || true)"
echo "app.py の修正版: $BUILD_LINE"
if echo "$BUILD_LINE" | grep -q "$NEED_BUILD"; then
  echo "OK: 新しい修正版が入っています。"
else
  echo "========================================"
  echo " 警告: まだ古い app.py のようです。"
  echo "========================================"
  echo ""
  read -r -p "このまま起動しますか？ (y/N): " ans
  case "$ans" in
    y|Y|yes|YES) ;;
    *) echo "中止しました。"; read -r -p "Enter で閉じる…"; exit 1 ;;
  esac
fi
echo ""

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
echo "背景画像フォルダ:"
echo "  $(pwd)/custom_backgrounds"
echo ""

# 依存パッケージの最低限チェック
echo "起動前チェック…"
if ! python3 -c "import streamlit, requests, docx, PIL, pypdf" 2>/dev/null; then
  echo "必要な Python パッケージが不足しています。インストールを試します…"
  python3 -m pip install -r requirements.txt --user
fi
if ! python3 -c "import app" 2>/dev/null; then
  echo "========================================"
  echo " エラー: app.py を読み込めません。"
  echo " 下の詳細をコピーして共有してください。"
  echo "========================================"
  python3 -c "import app" 2>&1 || true
  read -r -p "Enter で閉じる…"
  exit 1
fi
echo "OK: アプリの読み込み確認できました。"
echo ""

exec bash "./起動する.command"
