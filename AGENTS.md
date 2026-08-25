# AGENTS.md

## Cursor Cloud specific instructions

### What this app is
A single-file Streamlit app (`app.py`) — 「医学ドラマ動画メーカー」 — that turns a
Japanese text script into a 1920x1080 MP4. Pipeline: script import → confirm →
VOICEVOX narration → background frames + subtitles → MP4 (via moviepy/ffmpeg).

### Services / how to run
- Streamlit app: started by the `start` command in `.cursor/environment.json`
  (serves on port `8501`). It runs with system `python3` (deps install to the
  `--user` site via the update script), so no virtualenv activation is needed.
- VOICEVOX engine (required for ANY audio/video generation): the app connects to
  `http://127.0.0.1:50021` (`VOICEVOX_URL` in `app.py`). Without it the sidebar
  shows "VOICEVOX 未接続" and video creation is blocked.

### Non-obvious startup caveats (do this before generating video)
- Docker is installed in the snapshot but the daemon does NOT auto-start. Start it
  with `sudo dockerd > /tmp/dockerd.log 2>&1 &` (uses `fuse-overlayfs` storage
  driver + iptables-legacy, already configured in `/etc/docker/daemon.json`).
- The VOICEVOX engine image (`voicevox/voicevox_engine:cpu-latest`) is pulled in
  the snapshot but the container is run with `--rm`, so it does NOT persist across
  reboots. Start it fresh with:
  `sudo docker run -d --rm --name voicevox -p 50021:50021 voicevox/voicevox_engine:cpu-latest`
  then wait ~10s and confirm with `curl -s http://127.0.0.1:50021/version`.
- Japanese subtitles need CJK fonts. `fonts-noto-cjk` is installed in the snapshot;
  `app.load_jp_font()` falls back to `/usr/share/fonts/opentype/noto/NotoSansCJK-*.ttc`.
  If Japanese renders as boxes, reinstall with `sudo apt-get install -y fonts-noto-cjk`.
- `ffmpeg` (system) is used by moviepy; it is pre-installed on the base image.

### How to smoke-test end to end
Upload a `.txt` script under "1. 台本" → "台本を取り込む" → "2. 確定して動画作成へ"
→ "ドラフトMP4を作成する（背景なし）". Draft mode needs no background images and
proves the full pipeline. "最終版（背景あり）" additionally requires images placed in
`custom_backgrounds/` whose filenames match the `〈〉` scene markers in the script.
Output MP4 is written to the desktop dir (e.g. `~/Desktop/medical_drama.mp4`).

### Notes
- No automated test suite and no lint config in this repo; validation is manual
  (run the app and generate a video).
- The `ANTHROPIC_API_KEY` in `.env.example` is not referenced by the current
  `app.py`; it is not required to run or generate videos.
- Update script (VM startup) only refreshes Python deps
  (`pip3 install -r requirements.txt`). Docker/VOICEVOX/fonts live in the snapshot,
  not the update script.
