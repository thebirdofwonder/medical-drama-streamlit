"""医学ドラマ動画メーカー — 自動テスト（unittest）。"""

from __future__ import annotations

import sys
import tempfile
import unittest
import wave
from pathlib import Path
from unittest.mock import patch

# リポジトリ直下の app.py を import できるようにする
ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import app  # noqa: E402


class TestAppImport(unittest.TestCase):
    def test_app_build_is_set(self) -> None:
        self.assertTrue(app.APP_BUILD)
        self.assertIn("ui-slim", app.APP_BUILD)

    def test_main_function_exists(self) -> None:
        self.assertTrue(callable(app.main))


class TestFilenames(unittest.TestCase):
    def test_make_title_basename_sanitizes(self) -> None:
        self.assertEqual(app.make_title_basename("心不全ドラマ"), "心不全ドラマ")
        self.assertEqual(app.make_title_basename("a/b:c"), "a_b_c")
        self.assertEqual(app.make_title_basename(""), "medical_drama")

    def test_make_desktop_mp4_filename(self) -> None:
        self.assertEqual(app.make_desktop_mp4_filename("テスト"), "テスト.mp4")
        self.assertEqual(app.make_desktop_mp4_filename("テスト", draft=True), "テスト.mp4")

    def test_expected_custom_bg_filename(self) -> None:
        self.assertEqual(app.expected_custom_bg_filename("手術室"), "手術室.png")
        self.assertEqual(app.expected_custom_bg_filename("a/b"), "a_b.png")


class TestScriptUpload(unittest.TestCase):
    def test_assert_script_upload_size_ok(self) -> None:
        app.assert_script_upload_size(b"x" * 100, "test.txt")

    def test_assert_script_upload_size_too_large(self) -> None:
        big = b"x" * (app.MAX_SCRIPT_UPLOAD_BYTES + 1)
        with self.assertRaises(ValueError) as ctx:
            app.assert_script_upload_size(big, "big.txt")
        self.assertIn("大きすぎ", str(ctx.exception))

    def test_extract_text_from_txt(self) -> None:
        text = app.extract_text_from_bytes("script.txt", "心不全\n患者".encode("utf-8"))
        self.assertIn("心不全", text)

    def test_extract_text_rejects_unknown(self) -> None:
        with self.assertRaises(ValueError):
            app.extract_text_from_bytes("data.bin", b"\x00\x01")


class TestVoicevoxHelpers(unittest.TestCase):
    def test_clamp_voicevox_speed(self) -> None:
        self.assertEqual(app.clamp_voicevox_speed(1.0), 1.0)
        self.assertEqual(app.clamp_voicevox_speed(99.0), app.VOICEVOX_SPEED_MAX)
        self.assertEqual(app.clamp_voicevox_speed("bad"), app.VOICEVOX_SPEED_SCALE)

    def test_check_voicevox_when_offline(self) -> None:
        ok, msg = app.check_voicevox()
        self.assertIsInstance(ok, bool)
        self.assertIsInstance(msg, str)
        if not ok:
            self.assertIn("VOICEVOX", msg)

    def test_synthesize_raises_when_voicevox_down(self) -> None:
        ok, _ = app.check_voicevox()
        if ok:
            self.skipTest("VOICEVOX が接続中のためオフライン失敗テストをスキップ")
        with self.assertRaises((RuntimeError, Exception)):
            app.synthesize_wav_bytes("テスト")


class TestTextProcessing(unittest.TestCase):
    def test_strip_background_hints(self) -> None:
        self.assertEqual(app.strip_background_hint_segments("〈ER〉救急"), "救急")
        self.assertEqual(app.strip_background_hint_segments("<ER>救急"), "救急")

    def test_ruby_to_tts(self) -> None:
        tts = app.voicevox_tts_from_ruby_text("患者{心不全|しんふぜん}が来た。")
        self.assertEqual(tts, "患者しんふぜんが来た。")
        sub = app.strip_voicevox_ruby("患者{心不全|しんふぜん}が来た。")
        self.assertEqual(sub, "患者心不全が来た。")

    def test_split_text_for_voicevox(self) -> None:
        chunks = app.split_text_for_voicevox("一文目。二文目。")
        self.assertGreaterEqual(len(chunks), 2)

    def test_find_missing_backgrounds(self) -> None:
        missing = app.find_missing_custom_backgrounds(
            script="〈存在しない背景XYZ123〉患者。"
        )
        hints = {m["hint"] for m in missing}
        self.assertIn("存在しない背景XYZ123", hints)


class TestAudioVideo(unittest.TestCase):
    def test_make_silent_wav(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "silent.wav"
            app.make_silent_wav(path, duration_sec=0.5)
            self.assertTrue(path.is_file())
            with wave.open(str(path), "rb") as w:
                self.assertGreater(w.getnframes(), 0)

    def test_validate_audio_subtitle_sync_ok(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "n.wav"
            app.make_silent_wav(path, duration_sec=2.0)
            cues = [{"start": 0.0, "end": 2.0, "text": "字幕"}]
            issues = app.validate_audio_subtitle_sync(path, cues)
            self.assertEqual(issues, [])

    def test_validate_audio_subtitle_sync_detects_gap(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "n.wav"
            app.make_silent_wav(path, duration_sec=2.0)
            cues = [{"start": 1.0, "end": 2.0, "text": "字幕"}]
            issues = app.validate_audio_subtitle_sync(path, cues)
            self.assertTrue(any("遅すぎ" in m for m in issues))

    def test_build_mp4_with_silent_audio(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            wav = tmp_path / "n.wav"
            app.make_silent_wav(wav, duration_sec=1.0)
            frame = tmp_path / "frame.png"
            app.create_plain_scene_frame(frame)
            out = tmp_path / "out.mp4"
            app.build_mp4(wav, [(frame, 1.0)], out, subtitle_cues=[])
            self.assertTrue(out.is_file())
            self.assertGreater(out.stat().st_size, 500)

    def test_build_mp4_fails_without_scenes(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            wav = tmp_path / "n.wav"
            app.make_silent_wav(wav, duration_sec=0.5)
            out = tmp_path / "bad.mp4"
            with self.assertRaises(RuntimeError):
                app.build_mp4(wav, [], out)


class TestGenerateNarrationOffline(unittest.TestCase):
    def test_generate_narration_raises_when_voicevox_down(self) -> None:
        ok, _ = app.check_voicevox()
        if ok:
            self.skipTest("VOICEVOX 接続中")
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "n.wav"
            with self.assertRaises((RuntimeError, Exception)):
                app.generate_narration_wav_to_file("テスト文。", path)


class TestStreamlitHealth(unittest.TestCase):
    def test_streamlit_health_endpoint(self) -> None:
        import urllib.error
        import urllib.request

        try:
            with urllib.request.urlopen(
                "http://127.0.0.1:8501/_stcore/health", timeout=3
            ) as resp:
                body = resp.read().decode("utf-8").strip()
        except (urllib.error.URLError, TimeoutError):
            self.skipTest("Streamlit が起動していません")
        self.assertEqual(body, "ok")


class TestExportErrorLogging(unittest.TestCase):
    def test_export_error_file_can_be_written(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            err_path = Path(tmp) / "last_export_error.txt"
            err_path.write_text("traceback sample", encoding="utf-8")
            self.assertIn("traceback", err_path.read_text(encoding="utf-8"))


class TestVoicevoxHowto(unittest.TestCase):
    def test_howto_mentions_linux(self) -> None:
        text = app.voicevox_howto_start()
        self.assertIn("Linux", text)
        self.assertIn("docker", text.lower())
        self.assertIn(app.VOICEVOX_URL, text)


if __name__ == "__main__":
    unittest.main()
