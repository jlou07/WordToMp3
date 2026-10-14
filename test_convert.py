import contextlib
import io
import os
from pathlib import Path
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import patch
import wave

from convert import assemble_mp3, main, normalize_text, parser, split_text, synthesize_chunk, wav_duration


class ConverterTests(unittest.TestCase):
    def test_document_path_is_required(self):
        with contextlib.redirect_stderr(io.StringIO()):
            with self.assertRaises(SystemExit) as error:
                parser().parse_args([])
        self.assertEqual(error.exception.code, 2)

    def test_endpoint_from_environment_can_be_overridden(self):
        with patch.dict(os.environ, {"AZURE_SPEECH_ENDPOINT": "https://environment.example/"}):
            self.assertEqual(
                parser().parse_args(["document.txt"]).endpoint, "https://environment.example/",
            )
            self.assertEqual(
                parser().parse_args(["document.txt", "--endpoint", "https://override.example/"]).endpoint,
                "https://override.example/",
            )

    def test_dry_run_needs_no_azure_configuration(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "document.txt"
            source.write_text("Hello.", encoding="utf-8")
            with patch.dict(os.environ, {}, clear=True), patch(
                "sys.argv", ["convert.py", str(source), "--dry-run"],
            ), patch("convert.getpass.getpass") as prompt, contextlib.redirect_stdout(io.StringIO()):
                self.assertEqual(main(), 0)
                prompt.assert_not_called()
            self.assertEqual(list(Path(directory).iterdir()), [source])

    def test_real_conversion_requires_explicit_endpoint(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "document.txt"
            source.write_text("Hello.", encoding="utf-8")
            with patch.dict(os.environ, {}, clear=True), patch(
                "sys.argv", ["convert.py", str(source)],
            ), patch("convert.getpass.getpass") as prompt:
                with self.assertRaisesRegex(ValueError, "Missing Azure endpoint"):
                    main()
                prompt.assert_not_called()

    def test_normalizes_word_controls(self):
        self.assertEqual(normalize_text(" Hello\rworld\x07Cell\x02 "), "Hello\nworld\nCell")

    def test_preserves_all_words_with_limits(self):
        text = ("An original story. With dates and names!\n\n" * 180) + "The end."
        for limit in (100, 2000, 3000):
            chunks = split_text(text, limit)
            self.assertGreater(len(chunks), 1)
            self.assertTrue(all(0 < len(chunk) <= limit for chunk in chunks))
            self.assertEqual(" ".join(" ".join(chunks).split()), " ".join(text.split()))
            self.assertTrue(all(len(chunk.encode("utf-8")) <= 12000 for chunk in chunks))

    def test_long_unbroken_token(self):
        text = "x" * 6100
        chunks = split_text(text)
        self.assertEqual("".join(chunks), text)
        self.assertTrue(all(len(chunk) <= 2000 for chunk in chunks))

    def test_unicode_and_exact_boundary(self):
        text = "\U0001f600" * 2000
        self.assertEqual(split_text(text), [text])
        self.assertEqual(split_text(""), [])
        with self.assertRaises(ValueError):
            split_text("text", 5000)

    @staticmethod
    def write_wav(path: Path, seconds: int = 1):
        with wave.open(str(path), "wb") as audio:
            audio.setnchannels(1)
            audio.setsampwidth(2)
            audio.setframerate(24000)
            audio.writeframes(b"\0\0" * 24000 * seconds)

    def test_detects_truncated_wav(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "audio.wav"
            self.write_wav(path)
            self.assertEqual(wav_duration(path), 1)
            path.write_bytes(path.read_bytes()[:-10])
            with self.assertRaises(ValueError):
                wav_duration(path)

    def test_encodes_all_segments_as_one_mp3(self):
        import imageio_ffmpeg
        with tempfile.TemporaryDirectory(prefix="test audio ") as directory:
            cache = Path(directory) / "cache"
            cache.mkdir()
            self.write_wav(cache / "00001.wav", 1)
            self.write_wav(cache / "00002.wav", 2)
            output = Path(directory) / "audio.mp3"
            ffmpeg = imageio_ffmpeg.get_ffmpeg_exe()
            assemble_mp3(cache, 2, output, ffmpeg)
            self.assertGreater(output.stat().st_size, 1000)
            import subprocess
            decoded = Path(directory) / "decoded.wav"
            subprocess.run(
                [ffmpeg, "-v", "error", "-i", str(output), str(decoded)], check=True,
            )
            self.assertAlmostEqual(wav_duration(decoded), 3.0, delta=0.1)

    def test_failed_encoding_preserves_existing_output(self):
        import subprocess
        with tempfile.TemporaryDirectory() as directory:
            cache = Path(directory) / "cache"
            cache.mkdir()
            output = Path(directory) / "audio.mp3"
            output.write_bytes(b"previous recording")
            with patch("convert.subprocess.run", return_value=subprocess.CompletedProcess(
                [], 1, stdout="", stderr="encoding failure",
            )):
                with self.assertRaisesRegex(RuntimeError, "encoding failure"):
                    assemble_mp3(cache, 1, output, "ffmpeg")
            self.assertEqual(output.read_bytes(), b"previous recording")

    def test_retries_transient_errors_then_saves_complete_audio(self):
        import azure.cognitiveservices.speech as speechsdk
        with tempfile.TemporaryDirectory() as directory:
            destination = Path(directory) / "audio.wav"
            self.write_wav(destination)
            audio_data = destination.read_bytes()
            destination.unlink()
            transient = SimpleNamespace(
                reason=speechsdk.ResultReason.Canceled,
                cancellation_details=SimpleNamespace(
                    reason=speechsdk.CancellationReason.Error,
                    error_code=speechsdk.CancellationErrorCode.TooManyRequests,
                    error_details="Temporary error",
                ),
            )
            success = SimpleNamespace(
                reason=speechsdk.ResultReason.SynthesizingAudioCompleted,
                audio_data=audio_data,
            )
            from unittest.mock import Mock
            synthesizer = Mock()
            synthesizer.speak_text_async.return_value.get.side_effect = [transient, success]
            with patch("convert.time.sleep") as sleep:
                synthesize_chunk(synthesizer, "Hello", destination, "secret", 4)
            self.assertEqual(synthesizer.speak_text_async.call_count, 2)
            sleep.assert_called_once_with(2)
            self.assertEqual(wav_duration(destination), 1)

    def test_authentication_failure_stops_without_retry_or_key_leak(self):
        import azure.cognitiveservices.speech as speechsdk
        from unittest.mock import Mock
        synthesizer = Mock()
        synthesizer.speak_text_async.return_value.get.return_value = SimpleNamespace(
            reason=speechsdk.ResultReason.Canceled,
            cancellation_details=SimpleNamespace(
                reason=speechsdk.CancellationReason.Error,
                error_code=speechsdk.CancellationErrorCode.AuthenticationFailure,
                error_details="Rejected secret",
            ),
        )
        with tempfile.TemporaryDirectory() as directory:
            destination = Path(directory) / "audio.wav"
            with self.assertRaisesRegex(RuntimeError, "Rejected \\[key redacted\\]"):
                synthesize_chunk(synthesizer, "Hello", destination, "secret", 4)
            self.assertFalse(destination.exists())
        self.assertEqual(synthesizer.speak_text_async.call_count, 1)


if __name__ == "__main__":
    unittest.main()
