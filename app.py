"""
医学ドラマ台本 → AIレビュー → VOICEVOX音声 → 背景＋BGM → MP4 生成
Streamlit アプリ（macOS / Apple Silicon 向け）
"""

from __future__ import annotations

import io
import json
import os
import re
import struct
import tempfile
import wave
from pathlib import Path
from typing import Any

import requests
import streamlit as st
from docx import Document
from PIL import Image, ImageDraw, ImageFont

# ---------------------------------------------------------------------------
# 定数
# ---------------------------------------------------------------------------
VOICEVOX_URL = "http://127.0.0.1:50021"
SPEAKER_ID = 84  # 青山龍星
VIDEO_SIZE = (1920, 1080)
CREDIT_TEXT = "音声: VOICEVOX 青山龍星"
BGM_FILENAME = "bgm.mp3"
# Pixabay のフリー音源（暗め・シリアス寄り）。取得失敗時は簡易BGMを自動生成します。
BGM_CANDIDATE_URLS = [
    "https://cdn.pixabay.com/download/audio/2022/01/18/audio_d0c6ff1bab.mp3?filename=dark-ambient-11495.mp3",
    "https://cdn.pixabay.com/download/audio/2021/08/09/audio_dc39bde808.mp3?filename=cinematic-documentary-11521.mp3",
]
WORK_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = WORK_DIR / "outputs"
MAX_VOICEVOX_CHARS = 180  # 長文を分割して送る目安（30分級向けにやや短め）
DEFAULT_FOOTNOTE = ""
# Anthropic の現行モデル（旧 claude-sonnet-4-20250514 は引退済み）
CLAUDE_MODEL_CANDIDATES = [
    "claude-sonnet-4-6",
    "claude-sonnet-5",
    "claude-sonnet-4-5",
    "claude-haiku-4-5",
]
# レビュー用に送る台本の上限（長すぎると API が失敗しやすい）
REVIEW_SCRIPT_MAX_CHARS = 12000


# ---------------------------------------------------------------------------
# ユーティリティ: 台本読み込み
# ---------------------------------------------------------------------------
def extract_text_from_upload(uploaded_file) -> str:
    """アップロードされた .txt / .docx からテキストを取り出す。"""
    name = (uploaded_file.name or "").lower()
    raw = uploaded_file.read()

    if name.endswith(".txt"):
        for encoding in ("utf-8", "utf-8-sig", "cp932", "shift_jis"):
            try:
                return raw.decode(encoding)
            except UnicodeDecodeError:
                continue
        return raw.decode("utf-8", errors="replace")

    if name.endswith(".docx"):
        doc = Document(io.BytesIO(raw))
        parts = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        # 表の中の文章も拾う
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    t = cell.text.strip()
                    if t:
                        parts.append(t)
        return "\n".join(parts)

    raise ValueError("対応形式は .txt または .docx のみです。")


# ---------------------------------------------------------------------------
# AI レビュー（Anthropic Claude API / フォールバック簡易レビュー）
# ---------------------------------------------------------------------------
def get_api_key() -> str:
    """環境変数または Streamlit secrets から API キーを取得（コードに直書きしない）。"""
    key = os.environ.get("ANTHROPIC_API_KEY", "").strip()
    if key:
        return key
    try:
        return st.secrets.get("ANTHROPIC_API_KEY", "").strip()
    except Exception:
        return ""


def build_review_prompt(script: str) -> str:
    return f"""あなたは現役の救急・集中治療に詳しい医療監修医です。
以下の医学ドラマ台本を検証し、JSONオブジェクトだけを返してください。
前後に説明文・Markdown・コードフェンスは付けないでください。

厳守ルール:
- 有効なJSONのみ（末尾カンマ禁止）
- 文字列内に半角ダブルクォート " を書かない（必要なら『』や「」を使う）
- 各配列は最大5件
- original は40文字以内、issue/suggestion は120文字以内

形式:
{{
  "medical_contradictions": [
    {{"original": "引用", "issue": "問題", "suggestion": "修正案"}}
  ],
  "awkward_for_doctors": [
    {{"original": "引用", "issue": "問題", "suggestion": "修正案"}}
  ],
  "immersion_improvements": [
    {{"original": "引用", "issue": "問題", "suggestion": "修正案"}}
  ]
}}

該当が無い観点は空配列 [] にしてください。

台本:
---
{script}
---
"""


def strip_code_fence(text: str) -> str:
    text = text.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json|JSON)?\s*", "", text)
        text = re.sub(r"\s*```$", "", text)
    return text.strip()


def repair_json_text(text: str) -> str:
    """よくある壊れ方を直してから json.loads する。"""
    text = strip_code_fence(text)
    # 最初の { から最後の } まで
    start = text.find("{")
    end = text.rfind("}")
    if start >= 0 and end > start:
        text = text[start : end + 1]
    # 末尾カンマ: ,] や ,} を除去
    text = re.sub(r",\s*([\]}])", r"\1", text)
    # スマートクォートを半角に
    text = (
        text.replace("“", '"')
        .replace("”", '"')
        .replace("‘", "'")
        .replace("’", "'")
    )
    return text


def parse_json_loose(text: str) -> dict[str, Any]:
    """モデル出力から JSON をできるだけ取り出す。"""
    candidates = [text, repair_json_text(text)]
    errors: list[str] = []
    for cand in candidates:
        try:
            data = json.loads(cand)
            if isinstance(data, dict):
                return data
        except json.JSONDecodeError as e:
            errors.append(str(e))

    # さらに: 制御文字除去して再試行
    cleaned = repair_json_text(re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", text))
    try:
        data = json.loads(cleaned)
        if isinstance(data, dict):
            return data
    except json.JSONDecodeError as e:
        errors.append(str(e))

    raise ValueError(
        "AIの返答をJSONとして読めませんでした。\n"
        + (errors[-1] if errors else "")
    )


def ask_claude_fix_json(api_key: str, model: str, broken: str) -> str:
    """壊れたJSONを、同じモデルに直してもらう。"""
    url = "https://api.anthropic.com/v1/messages"
    headers = {
        "x-api-key": api_key,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json",
    }
    body = {
        "model": model,
        "max_tokens": 4096,
        "messages": [
            {
                "role": "user",
                "content": (
                    "次のテキストを、有効なJSONオブジェクトだけに直してください。"
                    "説明文やコードフェンスは不要です。"
                    "キーは medical_contradictions / awkward_for_doctors / "
                    "immersion_improvements を維持してください。\n\n"
                    f"{broken[:12000]}"
                ),
            }
        ],
    }
    resp = http_session_direct().post(url, headers=headers, json=body, timeout=120)
    if resp.status_code != 200:
        raise RuntimeError(
            f"JSON修正リクエスト失敗 (HTTP {resp.status_code}): {resp.text[:300]}"
        )
    data = resp.json()
    parts = data.get("content", [])
    return "".join(p.get("text", "") for p in parts if p.get("type") == "text")


def http_session_direct() -> requests.Session:
    """
    プロキシ（通信の仲介役）を使わず、インターネットへ直接つなぐ。
    Cursor や社内ネットのプロキシ設定があると Claude API が
    403 Forbidden で失敗することがあるため。
    """
    session = requests.Session()
    session.trust_env = False  # HTTP_PROXY 等の環境変数を無視
    session.proxies = {"http": None, "https": None}
    return session


def review_with_claude(script: str, api_key: str) -> dict[str, Any]:
    """Anthropic Messages API でレビュー（APIキーは引数経由、直書きしない）。"""
    url = "https://api.anthropic.com/v1/messages"
    headers = {
        "x-api-key": api_key,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json",
    }
    prompt = build_review_prompt(script)
    last_error = ""

    for model in CLAUDE_MODEL_CANDIDATES:
        body = {
            "model": model,
            "max_tokens": 8192,
            "messages": [{"role": "user", "content": prompt}],
        }
        try:
            resp = http_session_direct().post(
                url, headers=headers, json=body, timeout=120
            )
        except requests.exceptions.ProxyError as e:
            raise RuntimeError(
                "プロキシ（通信の仲介）のせいで Claude に接続できませんでした。\n"
                "ターミナルで次を実行してから、アプリを再起動してください:\n"
                "  unset HTTP_PROXY HTTPS_PROXY ALL_PROXY http_proxy https_proxy all_proxy\n"
                f"詳細: {e}"
            ) from e
        except requests.exceptions.RequestException as e:
            raise RuntimeError(
                "Claude API への通信に失敗しました。"
                "ネット接続と APIキーを確認してください。\n"
                f"詳細: {e}"
            ) from e

        if resp.status_code == 404 and "model" in resp.text.lower():
            last_error = f"{model}: {resp.text[:200]}"
            continue  # 次のモデル候補を試す

        if resp.status_code != 200:
            raise RuntimeError(
                f"Claude API エラー (HTTP {resp.status_code}): {resp.text[:500]}"
            )

        data = resp.json()
        parts = data.get("content", [])
        text = "".join(p.get("text", "") for p in parts if p.get("type") == "text")
        if not text:
            raise RuntimeError("Claude API から空の返答が返りました。")

        try:
            parsed = parse_json_loose(text)
        except ValueError:
            # 壊れたJSONを一度だけ修正依頼
            try:
                fixed = ask_claude_fix_json(api_key, model, text)
                parsed = parse_json_loose(fixed)
            except Exception as e:  # noqa: BLE001
                raise RuntimeError(
                    "Claudeのレビュー結果を読み取れませんでした。"
                    "もう一度『1. 台本をレビューする』を押してください。\n"
                    f"詳細: {e}"
                ) from e

        result = normalize_review(parsed)
        result["mode"] = "claude"
        result["model_used"] = model
        return result

    raise RuntimeError(
        "利用できる Claude モデルが見つかりませんでした。\n"
        f"試したモデル: {', '.join(CLAUDE_MODEL_CANDIDATES)}\n"
        f"最後のエラー: {last_error}"
    )

def heuristic_review(script: str) -> dict[str, Any]:
    """
    APIキーが無いとき用の簡易レビュー。
    本格的な医学監修の代わりではなく、プロセス確認用です。
    """
    medical: list[dict[str, str]] = []
    awkward: list[dict[str, str]] = []
    immersion: list[dict[str, str]] = []

    patterns_medical = [
        (
            r"血圧が\s*200[/\／]20",
            "血圧の下の値（拡張期）が極端に低すぎる表記は非現実的です。",
            "例: 「血圧 200/110」など臨床で起こりうる数値に修正してください。",
        ),
        (
            r"心電図(が|で)?(止まっ|フラット|直線)",
            "心停止の表現が曖昧です。波形の種類を明示した方が正確です。",
            "「心電図は心静止（アスystole）」「心室細動（VF）」など具体名に。",
        ),
        (
            r"酸素飽和度\s*(が)?\s*0%|SpO2\s*(が)?\s*0",
            "SpO2 0% は測定不能やプローブ外れの可能性が高く、物語上も説明が必要です。",
            "「測定不能」「プローブ外れの可能性」など状況説明を添えてください。",
        ),
    ]
    for pat, issue, suggestion in patterns_medical:
        m = re.search(pat, script)
        if m:
            medical.append(
                {
                    "original": m.group(0),
                    "issue": issue,
                    "suggestion": suggestion,
                }
            )

    patterns_awkward = [
        (
            r"オペ(を)?しよう|オペる",
            "現場では「オペ」単体より手技名・適応を言うことが多いです。",
            "「緊急開腹術を開始する」「気管内挿管する」など具体的に。",
        ),
        (
            r"点滴(を)?打(つ|って)",
            "医療者は「点滴を入れる／開始する」と言うことが多いです。",
            "「末梢ルートを確保して補液を開始」などに。",
        ),
        (
            r"心臓マッサージ",
            "現在は「胸骨圧迫」が標準的な言い方です。",
            "「胸骨圧迫を開始」に言い換えると医師の耳に自然です。",
        ),
    ]
    for pat, issue, suggestion in patterns_awkward:
        m = re.search(pat, script)
        if m:
            awkward.append(
                {
                    "original": m.group(0),
                    "issue": issue,
                    "suggestion": suggestion,
                }
            )

    if len(script) < 200:
        immersion.append(
            {
                "original": script[:80] + ("…" if len(script) > 80 else ""),
                "issue": "短いため、現場の音・時間経過・バイタルの変化が弱い可能性があります。",
                "suggestion": "モニター音、時刻、SpO2/血圧の推移、スタッフの短い掛け声を1〜2文足す。",
            }
        )
    if "…" not in script and "……" not in script and "——" not in script:
        immersion.append(
            {
                "original": "（全体）",
                "issue": "間（ま）や沈黙の描写が少なく、緊張感が平坦になりがちです。",
                "suggestion": "重要な決断の直前に短い沈黙やモニター音だけの一瞬を入れてください。",
            }
        )
    if not re.search(r"(血圧|SpO2|心拍数|脈拍|呼吸数)", script):
        immersion.append(
            {
                "original": "（バイタル表記なし）",
                "issue": "数値がないと救急シーンの臨場感が落ちます。",
                "suggestion": "「血圧 82/40、脈 130、SpO2 88%」など具体値を1か所入れてください。",
            }
        )

    if not medical and not awkward and not immersion:
        immersion.append(
            {
                "original": "（全体）",
                "issue": "自動チェックでは大きな問題は見つかりませんでした（簡易モード）。",
                "suggestion": "APIキーを設定すると、Claudeによる本格レビューに切り替えられます。",
            }
        )

    return normalize_review(
        {
            "medical_contradictions": medical,
            "awkward_for_doctors": awkward,
            "immersion_improvements": immersion,
            "mode": "heuristic",
        }
    )


def normalize_review(data: dict[str, Any]) -> dict[str, Any]:
    def _items(key: str) -> list[dict[str, str]]:
        raw = data.get(key, []) or []
        out: list[dict[str, str]] = []
        for item in raw:
            if not isinstance(item, dict):
                continue
            out.append(
                {
                    "original": str(item.get("original", "")).strip(),
                    "issue": str(item.get("issue", "")).strip(),
                    "suggestion": str(item.get("suggestion", "")).strip(),
                }
            )
        return out

    return {
        "medical_contradictions": _items("medical_contradictions"),
        "awkward_for_doctors": _items("awkward_for_doctors"),
        "immersion_improvements": _items("immersion_improvements"),
        "mode": data.get("mode", "claude"),
    }


def run_script_review(script: str) -> dict[str, Any]:
    api_key = get_api_key()
    review_text = script
    truncated = False
    if len(script) > REVIEW_SCRIPT_MAX_CHARS:
        review_text = (
            script[:REVIEW_SCRIPT_MAX_CHARS]
            + "\n\n…（以下省略。レビューは先頭部分のみ）"
        )
        truncated = True
    if api_key:
        result = review_with_claude(review_text, api_key)
    else:
        result = heuristic_review(review_text)
    result["review_truncated"] = truncated
    return result


# ---------------------------------------------------------------------------
# VOICEVOX
# ---------------------------------------------------------------------------
def check_voicevox() -> tuple[bool, str]:
    try:
        r = requests.get(f"{VOICEVOX_URL}/version", timeout=3)
        if r.status_code == 200:
            return True, r.text.strip().strip('"')
        return False, f"HTTP {r.status_code}"
    except requests.RequestException as e:
        return False, str(e)


def split_text_for_voicevox(text: str, max_chars: int = MAX_VOICEVOX_CHARS) -> list[str]:
    """句点などで区切り、長すぎる場合はさらに分割。"""
    text = text.replace("\r\n", "\n").strip()
    if not text:
        return []
    # 空行は短い間として扱う
    blocks = re.split(r"\n+", text)
    chunks: list[str] = []
    for block in blocks:
        block = block.strip()
        if not block:
            continue
        parts = re.split(r"(?<=[。！？!?])", block)
        buf = ""
        for part in parts:
            part = part.strip()
            if not part:
                continue
            if len(buf) + len(part) <= max_chars:
                buf += part
            else:
                if buf:
                    chunks.append(buf)
                if len(part) <= max_chars:
                    buf = part
                else:
                    # さらに強制分割
                    for i in range(0, len(part), max_chars):
                        chunks.append(part[i : i + max_chars])
                    buf = ""
        if buf:
            chunks.append(buf)
    return chunks


def synthesize_wav_bytes(text: str, speaker: int = SPEAKER_ID) -> bytes:
    q = requests.post(
        f"{VOICEVOX_URL}/audio_query",
        params={"text": text, "speaker": speaker},
        timeout=60,
    )
    if q.status_code != 200:
        raise RuntimeError(f"audio_query 失敗: HTTP {q.status_code} / {q.text[:300]}")
    query = q.json()
    # 少しゆっくり・はっきり（医学ドラマ向け）
    query["speedScale"] = float(query.get("speedScale", 1.0)) * 0.95
    query["intonationScale"] = float(query.get("intonationScale", 1.0))

    s = requests.post(
        f"{VOICEVOX_URL}/synthesis",
        params={"speaker": speaker},
        headers={"Content-Type": "application/json"},
        data=json.dumps(query),
        timeout=120,
    )
    if s.status_code != 200:
        raise RuntimeError(f"synthesis 失敗: HTTP {s.status_code} / {s.text[:300]}")
    return s.content


def concat_wav_files(wav_paths: list[Path], out_path: Path, pause_ms: int = 350) -> Path:
    """複数WAVを無音つきでつなぎ、ファイルへ書き出す（長尺向け・メモリ節約）。"""
    if not wav_paths:
        raise ValueError("音声データが空です。")

    params = None
    with wave.open(str(out_path), "wb") as out_w:
        for i, p in enumerate(wav_paths):
            with wave.open(str(p), "rb") as w:
                if params is None:
                    params = w.getparams()
                    out_w.setparams(params)
                else:
                    if (
                        w.getnchannels() != params.nchannels
                        or w.getsampwidth() != params.sampwidth
                        or w.getframerate() != params.framerate
                    ):
                        raise RuntimeError("WAVの形式が一致しません。")
                out_w.writeframes(w.readframes(w.getnframes()))
                if i < len(wav_paths) - 1 and pause_ms > 0 and params is not None:
                    n = int(params.framerate * pause_ms / 1000)
                    silence = b"\x00" * n * params.nchannels * params.sampwidth
                    out_w.writeframes(silence)
    return out_path


def generate_narration_wav_to_file(
    script: str,
    out_wav: Path,
    progress_callback=None,
) -> Path:
    """
    長い台本（30分前後）向け: 分割してVOICEVOXへ送り、1本のWAVに結合する。
    progress_callback(done, total) があれば進捗を伝える。
    """
    chunks = split_text_for_voicevox(script)
    if not chunks:
        raise ValueError("読み上げる文章が空です。")

    out_wav.parent.mkdir(parents=True, exist_ok=True)
    part_dir = out_wav.parent / "_voice_parts"
    part_dir.mkdir(parents=True, exist_ok=True)
    part_paths: list[Path] = []

    try:
        for i, chunk in enumerate(chunks):
            if progress_callback:
                progress_callback(i, len(chunks))
            part = part_dir / f"part_{i:05d}.wav"
            part.write_bytes(synthesize_wav_bytes(chunk))
            part_paths.append(part)
        if progress_callback:
            progress_callback(len(chunks), len(chunks))
        concat_wav_files(part_paths, out_wav)
    finally:
        for p in part_paths:
            try:
                p.unlink(missing_ok=True)
            except OSError:
                pass
        try:
            part_dir.rmdir()
        except OSError:
            pass
    return out_wav


def generate_narration_wav(script: str) -> bytes:
    """短い用途向け互換。長尺は generate_narration_wav_to_file を使う。"""
    with tempfile.TemporaryDirectory(prefix="vvox_") as tmp:
        path = Path(tmp) / "n.wav"
        generate_narration_wav_to_file(script, path)
        return path.read_bytes()


# ---------------------------------------------------------------------------
# 背景画像（Pillow）— YouTubeサムネ風タイトル＋VOICEVOXクレジット
# ---------------------------------------------------------------------------
def load_jp_font(size: int, bold: bool = True) -> ImageFont.ImageFont:
    """macOS で使える日本語フォントを探す。"""
    candidates = []
    if bold:
        candidates.extend(
            [
                "/System/Library/Fonts/ヒラギノ角ゴシック W8.ttc",
                "/System/Library/Fonts/ヒラギノ角ゴシック W6.ttc",
                "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
            ]
        )
    candidates.extend(
        [
            "/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc",
            "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
            "/Library/Fonts/Arial Unicode.ttf",
        ]
    )
    for path in candidates:
        try:
            return ImageFont.truetype(path, size)
        except OSError:
            continue
    return ImageFont.load_default()


def draw_outlined_text(
    draw: ImageDraw.ImageDraw,
    xy: tuple[int, int],
    text: str,
    font: ImageFont.ImageFont,
    fill: tuple[int, int, int],
    outline: tuple[int, int, int] = (0, 0, 0),
    outline_width: int = 6,
) -> None:
    """縁取り＋影つきの太字テキスト（YouTubeサムネ風）。"""
    x, y = xy
    # 影
    draw.text((x + 4, y + 4), text, font=font, fill=(0, 0, 0))
    # 縁取り（周囲に同じ文字を重ねる）
    for dx in range(-outline_width, outline_width + 1):
        for dy in range(-outline_width, outline_width + 1):
            if dx == 0 and dy == 0:
                continue
            if dx * dx + dy * dy > outline_width * outline_width:
                continue
            draw.text((x + dx, y + dy), text, font=font, fill=outline)
    draw.text((x, y), text, font=font, fill=fill)


def wrap_text_to_width(
    text: str, font: ImageFont.ImageFont, max_width: int, draw: ImageDraw.ImageDraw
) -> list[str]:
    """画面幅に収まるよう、日本語を1文字ずつ折り返す。"""
    text = (text or "").replace("\r\n", "\n").strip()
    if not text:
        return []
    lines: list[str] = []
    for paragraph in text.split("\n"):
        if not paragraph:
            lines.append("")
            continue
        buf = ""
        for ch in paragraph:
            trial = buf + ch
            bbox = draw.textbbox((0, 0), trial, font=font)
            if bbox[2] - bbox[0] <= max_width:
                buf = trial
            else:
                if buf:
                    lines.append(buf)
                buf = ch
        if buf:
            lines.append(buf)
    return lines


def fit_image_cover(img: Image.Image, size: tuple[int, int]) -> Image.Image:
    """画像を 1920x1080 に合わせて中央切り抜き（余白なし）。"""
    tw, th = size
    iw, ih = img.size
    scale = max(tw / iw, th / ih)
    nw, nh = int(iw * scale), int(ih * scale)
    img = img.resize((nw, nh), Image.Resampling.LANCZOS)
    left = (nw - tw) // 2
    top = (nh - th) // 2
    return img.crop((left, top, left + tw, top + th))


def load_image_from_upload(uploaded_file) -> Image.Image | None:
    if uploaded_file is None:
        return None
    data = uploaded_file.getvalue() if hasattr(uploaded_file, "getvalue") else uploaded_file.read()
    return Image.open(io.BytesIO(data)).convert("RGBA")


def load_text_from_upload(uploaded_file) -> str:
    if uploaded_file is None:
        return ""
    raw = uploaded_file.getvalue() if hasattr(uploaded_file, "getvalue") else uploaded_file.read()
    if isinstance(raw, str):
        return raw
    for enc in ("utf-8", "utf-8-sig", "cp932", "shift_jis"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def paste_centered(
    base: Image.Image,
    overlay: Image.Image,
    max_width_ratio: float = 0.9,
    max_height_ratio: float = 0.35,
    y_ratio: float = 0.08,
) -> None:
    """タイトル画像などを、幅制限つきで中央に重ねる。"""
    w, h = base.size
    max_w = int(w * max_width_ratio)
    max_h = int(h * max_height_ratio)
    ow, oh = overlay.size
    scale = min(max_w / ow, max_h / oh, 1.0)
    nw, nh = max(1, int(ow * scale)), max(1, int(oh * scale))
    overlay = overlay.resize((nw, nh), Image.Resampling.LANCZOS)
    x = (w - nw) // 2
    y = int(h * y_ratio)
    if overlay.mode != "RGBA":
        overlay = overlay.convert("RGBA")
    base.paste(overlay, (x, y), overlay)


def create_background_image(
    path: Path,
    title: str = "",
    subtitle: str = "",
    background_img: Image.Image | None = None,
    title_img: Image.Image | None = None,
    footnote_text: str = "",
    voicevox_credit: str = CREDIT_TEXT,
) -> Path:
    """
    YouTube向け静止画フレームを作る。
    - 背景: アップロード画像 or 自動生成の暗い医療ドラマ背景
    - タイトル: アップロード画像優先。無ければ文字タイトル
    - 脚注: 出典テキスト（著作権表示）
    - 右下: VOICEVOX クレジット（消さない）
    """
    w, h = VIDEO_SIZE

    if background_img is not None:
        base = fit_image_cover(background_img.convert("RGB"), VIDEO_SIZE).convert("RGBA")
        # 文字を読みやすくするため薄い暗幕
        shade = Image.new("RGBA", (w, h), (0, 0, 0, 90))
        base = Image.alpha_composite(base, shade)
    else:
        img = Image.new("RGB", (w, h), (8, 12, 22))
        draw_tmp = ImageDraw.Draw(img)
        for y in range(h):
            t = y / max(h - 1, 1)
            r = int(6 + 18 * t)
            g = int(10 + 28 * t)
            b = int(20 + 40 * t)
            draw_tmp.line([(0, y), (w, y)], fill=(r, g, b))
        overlay = Image.new("RGBA", (w, h), (0, 0, 0, 0))
        od = ImageDraw.Draw(overlay)
        cx, cy = w // 2, int(h * 0.42)
        for radius, alpha in ((520, 28), (360, 40), (200, 55)):
            od.ellipse(
                [cx - radius, cy - radius, cx + radius, cy + radius],
                fill=(30, 80, 120, alpha),
            )
        base = Image.alpha_composite(img.convert("RGBA"), overlay)
        draw_tmp = ImageDraw.Draw(base.convert("RGB"))
        # 水平線は RGB 版に描いて戻す
        rgb = base.convert("RGB")
        d2 = ImageDraw.Draw(rgb)
        for i in range(8):
            yy = int(h * 0.25) + i * 70
            d2.line([(80, yy), (w - 80, yy)], fill=(40, 70, 90), width=1)
        base = rgb.convert("RGBA")

    draw = ImageDraw.Draw(base)

    # タイトル画像があれば優先
    if title_img is not None:
        paste_centered(
            base,
            title_img,
            max_width_ratio=0.92,
            max_height_ratio=0.38,
            y_ratio=0.06,
        )
        draw = ImageDraw.Draw(base)
    else:
        max_text_w = int(w * 0.88)
        title = (title or "").strip() or "医学ドラマ"
        subtitle = (subtitle or "").strip()
        title_font = load_jp_font(92, bold=True)
        sub_font = load_jp_font(64, bold=True)
        pink = (255, 80, 160)
        yellow = (255, 230, 60)
        cyan = (80, 210, 255)

        title_lines = wrap_text_to_width(title, title_font, max_text_w, draw)
        y_cursor = int(h * 0.10)
        for line in title_lines[:4]:
            bbox = draw.textbbox((0, 0), line, font=title_font)
            tw = bbox[2] - bbox[0]
            th = bbox[3] - bbox[1]
            x = (w - tw) // 2
            draw_outlined_text(
                draw,
                (x, y_cursor),
                line,
                title_font,
                fill=pink,
                outline=(0, 0, 0),
                outline_width=8,
            )
            y_cursor += th + 18

        if subtitle:
            sub_lines = wrap_text_to_width(subtitle, sub_font, max_text_w, draw)
            y_sub = int(h * 0.58)
            for i, line in enumerate(sub_lines[:4]):
                bbox = draw.textbbox((0, 0), line, font=sub_font)
                tw = bbox[2] - bbox[0]
                th = bbox[3] - bbox[1]
                x = (w - tw) // 2
                color = yellow if i % 2 == 0 else cyan
                draw_outlined_text(
                    draw,
                    (x, y_sub),
                    line,
                    sub_font,
                    fill=color,
                    outline=(0, 0, 0),
                    outline_width=7,
                )
                y_sub += th + 14

    # --- 脚注（出典・著作権表示）左下 ---
    footnote = (footnote_text or "").strip()
    footnote_font = load_jp_font(28, bold=False)
    margin = 40
    y_foot = h - margin - 80
    if footnote:
        foot_lines = wrap_text_to_width(footnote, footnote_font, int(w * 0.62), draw)
        # 下から積み上げ
        line_heights = []
        for line in foot_lines[-6:]:
            bbox = draw.textbbox((0, 0), line, font=footnote_font)
            line_heights.append(bbox[3] - bbox[1] + 6)
        total_h = sum(line_heights) if line_heights else 0
        y = h - margin - 50 - total_h
        for line in foot_lines[-6:]:
            draw.text((margin + 2, y + 2), line, font=footnote_font, fill=(0, 0, 0))
            draw.text((margin, y), line, font=footnote_font, fill=(200, 205, 210))
            bbox = draw.textbbox((0, 0), line, font=footnote_font)
            y += bbox[3] - bbox[1] + 6

    # --- VOICEVOX クレジット（消さない）右下 ---
    credit_font = load_jp_font(36, bold=False)
    credit = (voicevox_credit or CREDIT_TEXT).strip() or CREDIT_TEXT
    bbox = draw.textbbox((0, 0), credit, font=credit_font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    x = w - tw - margin
    y = h - th - margin
    draw.text((x + 2, y + 2), credit, font=credit_font, fill=(0, 0, 0))
    draw.text((x, y), credit, font=credit_font, fill=(210, 220, 230))

    path.parent.mkdir(parents=True, exist_ok=True)
    base.convert("RGB").save(path, format="PNG")
    return path


# ---------------------------------------------------------------------------
# BGM
# ---------------------------------------------------------------------------
def make_fallback_bgm_wav(path: Path, seconds: float = 60.0, rate: int = 44100) -> Path:
    """ダウンロード失敗時: 低いドローン音のWAVを自作（追加ライブラリ不要）。"""
    import math

    path = path.with_suffix(".wav")
    n = int(seconds * rate)
    frames = bytearray()
    for i in range(n):
        t = i / rate
        # 低いサイン波を重ねてシリアスな雰囲気に
        v = 0.15 * math.sin(2 * math.pi * 55 * t)
        v += 0.08 * math.sin(2 * math.pi * 82.5 * t)
        v += 0.05 * math.sin(2 * math.pi * 110 * t)
        # フェードイン/アウト
        fade = min(t / 2.0, 1.0, max(0.0, (seconds - t) / 2.0))
        sample = int(max(-1.0, min(1.0, v * fade)) * 32767)
        frames += struct.pack("<h", sample)

    with wave.open(str(path), "wb") as w:
        w.setnchannels(1)
        w.setsampwidth(2)
        w.setframerate(rate)
        w.writeframes(frames)
    return path


def ensure_bgm(work_dir: Path = WORK_DIR) -> Path:
    """bgm.mp3 が無ければ取得。失敗時はWAVを生成。"""
    mp3_path = work_dir / BGM_FILENAME
    if mp3_path.exists() and mp3_path.stat().st_size > 1000:
        return mp3_path

    for url in BGM_CANDIDATE_URLS:
        try:
            r = http_session_direct().get(url, timeout=60, allow_redirects=True)
            if r.status_code == 200 and len(r.content) > 1000:
                ctype = r.headers.get("Content-Type", "")
                if "html" in ctype.lower():
                    continue
                mp3_path.write_bytes(r.content)
                return mp3_path
        except requests.RequestException:
            continue

    return make_fallback_bgm_wav(work_dir / "bgm_fallback")


# ---------------------------------------------------------------------------
# 動画合成（moviepy）
# ---------------------------------------------------------------------------
def build_mp4(
    narration_wav: Path,
    background_png: Path,
    bgm_path: Path,
    output_mp4: Path,
) -> Path:
    """静止画＋長尺音声向け。fpsは1（YouTube静止画でOK・高速）。"""
    # moviepy 1.x
    from moviepy.editor import (
        AudioFileClip,
        CompositeAudioClip,
        ImageClip,
        afx,
    )

    voice = AudioFileClip(str(narration_wav))
    duration = float(voice.duration)
    if duration <= 0:
        voice.close()
        raise RuntimeError("音声の長さが 0 です。")

    bgm = AudioFileClip(str(bgm_path))
    # BGMを本編長さに合わせ、音量を下げる
    if bgm.duration < duration:
        bgm = afx.audio_loop(bgm, duration=duration)
    else:
        bgm = bgm.subclip(0, duration)
    bgm = bgm.volumex(0.18)

    mixed = CompositeAudioClip([voice, bgm])

    # 静止画なので fps=1 で十分（30分でも処理が軽い）
    still_fps = 1
    video = (
        ImageClip(str(background_png))
        .set_duration(duration)
        .set_fps(still_fps)
        .set_audio(mixed)
    )

    output_mp4.parent.mkdir(parents=True, exist_ok=True)
    tried_errors: list[str] = []

    encode_attempts = [
        {
            "codec": "h264_videotoolbox",
            "audio_codec": "aac",
            "ffmpeg_params": ["-b:v", "3M", "-pix_fmt", "yuv420p"],
        },
        {
            "codec": "libx264",
            "audio_codec": "aac",
            "preset": "ultrafast",
            "ffmpeg_params": ["-pix_fmt", "yuv420p", "-tune", "stillimage"],
        },
    ]

    last_err: Exception | None = None
    for opts in encode_attempts:
        try:
            kwargs = {
                "fps": still_fps,
                "codec": opts["codec"],
                "audio_codec": opts["audio_codec"],
                "threads": 0,
                "logger": None,
            }
            if "preset" in opts:
                kwargs["preset"] = opts["preset"]
            if "ffmpeg_params" in opts:
                kwargs["ffmpeg_params"] = opts["ffmpeg_params"]

            video.write_videofile(str(output_mp4), **kwargs)
            last_err = None
            break
        except Exception as e:  # noqa: BLE001 — エンコード失敗時に次の設定へ
            tried_errors.append(f"{opts['codec']}: {e}")
            last_err = e
            if output_mp4.exists():
                try:
                    output_mp4.unlink()
                except OSError:
                    pass

    voice.close()
    bgm.close()
    mixed.close()
    video.close()

    if last_err is not None:
        raise RuntimeError(
            "動画エンコードに失敗しました。\n" + "\n".join(tried_errors)
        ) from last_err

    return output_mp4


# ---------------------------------------------------------------------------
# UI ヘルパー
# ---------------------------------------------------------------------------
def render_review_section(title: str, items: list[dict[str, str]]) -> None:
    st.subheader(title)
    if not items:
        st.caption("該当なし")
        return
    for i, item in enumerate(items, start=1):
        with st.expander(f"{i}. {item.get('original') or '（箇所）'}", expanded=(i == 1)):
            st.markdown("**問題点**")
            st.write(item.get("issue") or "（なし）")
            st.markdown("**修正案**")
            st.write(item.get("suggestion") or "（なし）")


def init_state() -> None:
    defaults = {
        "raw_script": "",
        "review": None,
        "final_script": "",
        "review_done": False,
        "script_confirmed": False,
        "mp4_bytes": None,
        "mp4_path": "",
        "mp4_name": "medical_drama.mp4",
        "last_error": "",
        "video_title": "命を賭けた決断",
        "video_subtitle": "その一言が、すべてを変えた",
        "footnote_text": "",
        "last_script_path": "",
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


# ---------------------------------------------------------------------------
# Streamlit メイン
# ---------------------------------------------------------------------------
def main() -> None:
    st.set_page_config(
        page_title="医学ドラマ動画メーカー",
        page_icon="🎬",
        layout="wide",
    )
    init_state()

    st.title("医学ドラマ動画メーカー")
    st.write(
        "台本（テキストまたはWord）を上げると、"
        "AIレビュー → VOICEVOX音声 → 静止画背景・BGM合成 → MP4ダウンロード、まで進めます。"
        "台本は30分前後の長さにも対応しています。"
    )

    with st.sidebar:
        st.header("設定")
        st.caption("VOICEVOX はあらかじめ起動しておいてください。")
        ok, ver = check_voicevox()
        if ok:
            st.success(f"VOICEVOX 接続OK（version: {ver}）")
        else:
            st.error(f"VOICEVOX に接続できません: {ver}")
            st.info("VOICEVOXアプリを起動し、エンジンが 50021 番で待っているか確認してください。")

        st.divider()
        st.markdown("**AIレビュー用 APIキー**")
        st.caption(
            "Anthropic（クロード）の APIキーを使う場合は、"
            "ターミナルで環境変数 `ANTHROPIC_API_KEY` を設定するか、"
            "下の欄に一時入力してください（コードには保存しません）。"
        )
        typed = st.text_input("ANTHROPIC_API_KEY（任意）", type="password")
        if typed.strip():
            os.environ["ANTHROPIC_API_KEY"] = typed.strip()

        if get_api_key():
            st.success("Claude API キーを検出 → 本格レビューモード")
        else:
            st.warning("APIキーなし → 簡易レビューモード（動作確認用）")

    # ----- Step 1: アップロード -----
    st.header("ステップ1: 台本をアップロード")
    uploaded = st.file_uploader(
        "台本ファイル（.txt または .docx）",
        type=["txt", "docx"],
    )

    if uploaded is not None:
        try:
            text = extract_text_from_upload(uploaded)
            if text.strip() != st.session_state.raw_script.strip():
                st.session_state.raw_script = text
                st.session_state.final_script = text
                st.session_state.final_script_editor = text
                st.session_state.review = None
                st.session_state.review_done = False
                st.session_state.script_confirmed = False
                st.session_state.mp4_bytes = None
        except Exception as e:  # noqa: BLE001
            st.error(f"ファイルの読み込みに失敗しました: {e}")

    if st.session_state.raw_script:
        n_chars = len(st.session_state.raw_script)
        # 目安: 日本語読み上げおおよそ 300〜350字/分 → ここでは 320字/分
        est_min = max(1, round(n_chars / 320))
        st.info(
            f"読み込んだ文字数: 約 {n_chars:,} 字　／　読み上げ目安: 約 {est_min} 分"
            "（実際の長さはVOICEVOXの速さで変わります）"
        )
        with st.expander("読み込んだ台本（原文）", expanded=False):
            st.text(st.session_state.raw_script)

    col1, _ = st.columns([1, 2])
    with col1:
        review_clicked = st.button(
            "1. 台本をレビューする",
            type="primary",
            disabled=not bool(st.session_state.raw_script.strip()),
            use_container_width=True,
        )

    if review_clicked:
        with st.spinner("台本をレビューしています…"):
            try:
                st.session_state.review = run_script_review(st.session_state.raw_script)
                st.session_state.review_done = True
                st.session_state.script_confirmed = False
                st.session_state.mp4_bytes = None
                st.session_state.last_error = ""
            except Exception as e:  # noqa: BLE001
                st.session_state.last_error = str(e)
                st.error(f"レビューに失敗しました: {e}")

    # ----- Step 2: レビュー結果と採否 -----
    if st.session_state.review_done and st.session_state.review:
        st.header("ステップ2: レビュー結果を確認し、最終台本を確定")
        review = st.session_state.review
        mode = review.get("mode", "claude")
        if mode == "heuristic":
            st.info(
                "いまは簡易レビューです。"
                "より本格的な医学チェックには Anthropic API キーを設定してください。"
            )
        else:
            st.success("Claude によるレビュー結果です。採否を判断して台本を直してください。")
        if review.get("review_truncated"):
            st.warning(
                "台本がとても長いため、レビューは先頭部分のみです。"
                "最終台本は全文を確認・編集してください。"
            )

        render_review_section(
            "(1) 医学的に矛盾している箇所",
            review.get("medical_contradictions", []),
        )
        render_review_section(
            "(2) 現役の医師が聞くと違和感がある表現",
            review.get("awkward_for_doctors", []),
        )
        render_review_section(
            "(3) 修正すると臨場感が増す箇所",
            review.get("immersion_improvements", []),
        )

        st.subheader("最終台本（ここで直してから確定）")
        if "final_script_editor" not in st.session_state:
            st.session_state.final_script_editor = (
                st.session_state.final_script or st.session_state.raw_script
            )
        st.text_area(
            "レビューを反映した文章を編集できます",
            height=320,
            key="final_script_editor",
        )

        if st.button("2. この台本で動画制作に進む", type="primary"):
            edited = (st.session_state.get("final_script_editor") or "").strip()
            if not edited:
                st.error("最終台本が空です。文章を入れてください。")
            else:
                st.session_state.final_script = edited
                st.session_state.script_confirmed = True
                st.success("台本を確定しました。次のステップで動画を生成できます。")

    # ----- Step 3: 動画生成 -----
    if st.session_state.script_confirmed:
        st.header("ステップ3: 素材を用意してMP4を作る")
        st.caption(
            "VOICEVOX（青山龍星）＋ 静止画背景 ＋ BGM → MP4"
            "（YouTube向けは静止画でOK。30分級も対応）"
        )

        st.subheader("① 背景画像（任意）")
        bg_upload = st.file_uploader(
            "背景画像をアップロード（.png / .jpg / .webp）",
            type=["png", "jpg", "jpeg", "webp"],
            key="bg_image_upload",
            help="未設定なら自動の暗い医療ドラマ背景を使います",
        )
        if bg_upload is not None:
            st.image(bg_upload, caption="背景プレビュー", use_container_width=True)

        st.subheader("② タイトル（画像 or 文字）")
        title_upload = st.file_uploader(
            "タイトル画像をアップロード（.png 推奨・透明OK）",
            type=["png", "jpg", "jpeg", "webp"],
            key="title_image_upload",
            help="画像があれば文字タイトルより優先されます",
        )
        if title_upload is not None:
            st.image(title_upload, caption="タイトル画像プレビュー", use_container_width=True)

        st.text_input(
            "タイトル文字（画像が無いとき用・ピンク）",
            key="video_title",
        )
        st.text_input(
            "サブタイトル文字（黄／水色）",
            key="video_subtitle",
        )

        st.subheader("③ 脚注・出典（著作権表示）")
        footnote_upload = st.file_uploader(
            "出典テキストをアップロード（.txt）",
            type=["txt"],
            key="footnote_upload",
            help="例: 画像提供 ○○ / BGM: Pixabay など",
        )
        if footnote_upload is not None:
            file_id = f"{footnote_upload.name}-{footnote_upload.size}"
            if st.session_state.get("_footnote_file_id") != file_id:
                st.session_state.footnote_text = load_text_from_upload(
                    footnote_upload
                ).strip()
                st.session_state["_footnote_file_id"] = file_id
        st.text_area(
            "脚注に入れる出典の文字（左下に表示）",
            key="footnote_text",
            height=100,
            placeholder="例）背景写真: ○○（利用許諾済） / 参考資料: △△",
        )
        st.caption(f"右下のVOICEVOXクレジットは残します: {CREDIT_TEXT}")

        if st.button("3. 動画を生成する", type="primary"):
            progress = st.progress(0, text="準備中…")
            status = st.empty()
            try:
                ok, ver = check_voicevox()
                if not ok:
                    raise RuntimeError(
                        "VOICEVOX に接続できません。アプリを起動してから再実行してください。"
                        f"（詳細: {ver}）"
                    )

                OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
                # 台本を残す（次回「最後の台本を見せて」に備える）
                script_path = OUTPUT_DIR / "last_script.txt"
                script_path.write_text(
                    st.session_state.final_script, encoding="utf-8"
                )
                st.session_state.last_script_path = str(script_path)

                with tempfile.TemporaryDirectory(prefix="meddrama_") as tmp:
                    tmp_path = Path(tmp)

                    status.info("① 音声を生成しています（長い台本は時間がかかります）…")
                    wav_path = tmp_path / "narration.wav"
                    n_chunks = len(
                        split_text_for_voicevox(st.session_state.final_script)
                    )

                    def _voice_prog(done: int, total: int) -> None:
                        pct = 5 + int(50 * (done / max(total, 1)))
                        progress.progress(
                            min(pct, 55),
                            text=f"VOICEVOX 音声生成中… {done}/{total} 区間",
                        )
                        status.info(
                            f"① VOICEVOXで読み上げ中（{done}/{total}）…"
                            f" 予定区間数 {n_chunks}"
                        )

                    generate_narration_wav_to_file(
                        st.session_state.final_script,
                        wav_path,
                        progress_callback=_voice_prog,
                    )

                    status.info("② 静止画フレーム（背景・タイトル・出典・クレジット）を作っています…")
                    progress.progress(60, text="静止画を合成中…")
                    bg_path = tmp_path / "background.png"
                    bg_img = load_image_from_upload(bg_upload)
                    title_img = load_image_from_upload(title_upload)
                    create_background_image(
                        bg_path,
                        title=st.session_state.get("video_title", ""),
                        subtitle=st.session_state.get("video_subtitle", ""),
                        background_img=bg_img,
                        title_img=title_img,
                        footnote_text=st.session_state.get("footnote_text", ""),
                    )
                    # プレビュー用にも保存
                    preview_path = OUTPUT_DIR / "last_frame.png"
                    preview_path.write_bytes(bg_path.read_bytes())

                    status.info("③ BGMを用意しています…")
                    progress.progress(70, text="BGMを取得中…")
                    bgm_path = ensure_bgm(WORK_DIR)

                    status.info("④ 静止画＋音声をMP4にしています（30分級は数分かかることがあります）…")
                    progress.progress(80, text="MP4へエンコード中…")
                    out_path = OUTPUT_DIR / "medical_drama.mp4"
                    build_mp4(wav_path, bg_path, bgm_path, out_path)

                    st.session_state.mp4_path = str(out_path)
                    st.session_state.mp4_name = "medical_drama.mp4"
                    # 巨大ファイルはメモリに載せない（ダウンロードはファイルから）
                    st.session_state.mp4_bytes = None
                    progress.progress(100, text="完了")
                    status.success(
                        "動画の生成が完了しました。"
                        f"保存先: {out_path}"
                    )

            except Exception as e:  # noqa: BLE001
                st.session_state.mp4_path = ""
                st.session_state.mp4_bytes = None
                st.error(f"動画生成に失敗しました: {e}")
                st.exception(e)

        mp4_path = st.session_state.get("mp4_path") or ""
        if mp4_path and Path(mp4_path).exists():
            size_mb = Path(mp4_path).stat().st_size / (1024 * 1024)
            st.success(f"完成ファイル: `{mp4_path}` （約 {size_mb:.1f} MB）")
            # 大きいMP4を毎回メモリに載せると落ちやすいので上限を設ける
            if size_mb < 180:
                st.download_button(
                    label="MP4をダウンロード",
                    data=Path(mp4_path).read_bytes(),
                    file_name=st.session_state.mp4_name,
                    mime="video/mp4",
                    type="primary",
                )
            else:
                st.info(
                    "ファイルが大きいため、ブラウザからのダウンロードは省略しました。"
                    "Finder で次の場所を開いてください。"
                )
                st.code(mp4_path)
            st.caption(
                "30分前後の動画はブラウザ再生が重いことがあります。"
                "QuickTime などで確認してください。"
            )
            frame = OUTPUT_DIR / "last_frame.png"
            if frame.exists():
                st.image(
                    str(frame),
                    caption="動画に使う静止画フレーム",
                    use_container_width=True,
                )
            script_saved = st.session_state.get("last_script_path") or ""
            if script_saved and Path(script_saved).exists():
                st.download_button(
                    label="今回の台本テキストをダウンロード",
                    data=Path(script_saved).read_text(encoding="utf-8"),
                    file_name="last_script.txt",
                    mime="text/plain",
                )
        elif st.session_state.mp4_bytes:
            st.download_button(
                label="MP4をダウンロード",
                data=st.session_state.mp4_bytes,
                file_name=st.session_state.mp4_name,
                mime="video/mp4",
                type="primary",
            )
            st.video(st.session_state.mp4_bytes)


if __name__ == "__main__":
    main()
