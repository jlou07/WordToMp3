"""Convert a long Word document into one MP3 using Azure Speech."""

from __future__ import annotations

import argparse
import getpass
import hashlib
import json
import os
from pathlib import Path
import re
import subprocess
import sys
import tempfile
import time
from urllib.parse import urlparse
import wave


DEFAULT_VOICE = "fr-FR-DeniseNeural"
DEFAULT_CHUNK_SIZE = 2000


def normalize_text(text: str) -> str:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[\x07\x0b\x0c]", "\n", text)
    text = "".join(
        char for char in text
        if char in "\n\t" or 0x20 <= ord(char) <= 0xD7FF
        or 0xE000 <= ord(char) <= 0xFFFD or 0x10000 <= ord(char) <= 0x10FFFF
    )
    text = re.sub(r"[^\S\n]+", " ", text)
    return re.sub(r"\n{3,}", "\n\n", text).strip()


def split_text(text: str, limit: int = DEFAULT_CHUNK_SIZE) -> list[str]:
    if not 100 <= limit <= 3000:
        raise ValueError("Chunk size must be between 100 and 3000.")
    remaining = normalize_text(text)
    chunks: list[str] = []
    while remaining:
        if len(remaining) <= limit:
            chunks.append(remaining)
            break
        window = remaining[:limit + 1]
        cut = window.rfind("\n\n", limit // 2, limit + 1)
        if cut < 0:
            boundaries = list(re.finditer(r'[.!?]["\u00bb\u201d]?\s+', window))
            cut = boundaries[-1].end() if boundaries else -1
        if cut < limit // 2:
            spaces = list(re.finditer(r"\s+", window))
            cut = spaces[-1].start() if spaces else limit
        cut = min(cut, limit)
        chunks.append(remaining[:cut].strip())
        remaining = remaining[cut:].lstrip()
    return chunks


def extract_text(source: Path) -> str:
    if source.suffix.lower() == ".txt":
        return normalize_text(source.read_text(encoding="utf-8-sig"))
    if source.suffix.lower() not in {".doc", ".docx"}:
        raise ValueError("Supported formats: .doc, .docx, and UTF-8 .txt.")
    if sys.platform != "win32":
        raise RuntimeError("Word extraction requires Windows and Microsoft Word.")
    import pythoncom
    import pywintypes
    import win32com.client

    pythoncom.CoInitialize()
    word = None
    document = None
    old_update_links = None
    try:
        # DispatchEx isolates this instance from the user's open Word session.
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        word.AutomationSecurity = 3
        old_update_links = word.Options.UpdateLinksAtOpen
        word.Options.UpdateLinksAtOpen = False
        document = word.Documents.Open(
            FileName=str(source), ConfirmConversions=False, ReadOnly=True,
            AddToRecentFiles=False, PasswordDocument="", PasswordTemplate="",
            WritePasswordDocument="", WritePasswordTemplate="",
            Visible=False, OpenAndRepair=False, NoEncodingDialog=True,
        )
        parts = [document.Content.Text]
        for collection in (document.Footnotes, document.Endnotes):
            for index in range(1, collection.Count + 1):
                parts.append(collection.Item(index).Range.Text)
        return normalize_text("\n\n".join(parts))
    except pywintypes.com_error as error:
        raise RuntimeError(f"Word extraction failed for {source.name}: {error}") from error
    finally:
        try:
            if document is not None:
                document.Close(SaveChanges=0)
        finally:
            try:
                if word is not None:
                    try:
                        if old_update_links is not None:
                            word.Options.UpdateLinksAtOpen = old_update_links
                    finally:
                        word.Quit(SaveChanges=0)
            finally:
                pythoncom.CoUninitialize()


def wav_duration(path: Path) -> float:
    with wave.open(str(path), "rb") as audio:
        if (audio.getnchannels(), audio.getsampwidth(), audio.getframerate()) != (1, 2, 24000):
            raise ValueError(f"Unexpected WAV format: {path.name}")
        frames = audio.getnframes()
        if frames <= 0:
            raise ValueError(f"Empty audio segment: {path.name}")
        read = 0
        while data := audio.readframes(24000):
            read += len(data)
        if read != frames * 2:
            raise ValueError(f"Incomplete audio segment: {path.name}")
        duration = frames / 24000
        if duration >= 540:
            raise ValueError("Segment too long: reduce --chunk-size to avoid truncation.")
        return duration


def synthesize_chunk(synthesizer, text: str, destination: Path, key: str, attempts: int) -> None:
    import azure.cognitiveservices.speech as speechsdk

    transient_errors = {
        speechsdk.CancellationErrorCode.ConnectionFailure,
        speechsdk.CancellationErrorCode.ServiceTimeout,
        speechsdk.CancellationErrorCode.ServiceUnavailable,
        speechsdk.CancellationErrorCode.TooManyRequests,
    }
    for attempt in range(1, attempts + 1):
        result = synthesizer.speak_text_async(text).get()
        if result.reason == speechsdk.ResultReason.SynthesizingAudioCompleted:
            temporary = destination.with_suffix(".partial.wav")
            try:
                temporary.write_bytes(result.audio_data)
                wav_duration(temporary)
                temporary.replace(destination)
            finally:
                temporary.unlink(missing_ok=True)
            return
        details = result.cancellation_details
        if details is None:
            raise RuntimeError(f"Unexpected Azure response: {result.reason}")
        message = (
            f"Azure Speech: {details.reason}, {details.error_code}. "
            f"{details.error_details}"
        ).replace(key, "[key redacted]")
        if details.error_code not in transient_errors or attempt == attempts:
            raise RuntimeError(message)
        delay = min(2 ** attempt, 30)
        print(f"{message}\nRetrying in {delay} seconds.", file=sys.stderr, flush=True)
        time.sleep(delay)


def assemble_mp3(cache: Path, count: int, output: Path, ffmpeg: str) -> None:
    concat = cache / "concat.txt"
    concat.write_text(
        "".join(f"file '{index:05d}.wav'\n" for index in range(1, count + 1)),
        encoding="ascii",
    )
    with tempfile.TemporaryDirectory(prefix="speech-encode-", dir=output.parent) as directory:
        temporary = Path(directory) / "audio.mp3"
        command = [
            ffmpeg, "-hide_banner", "-loglevel", "error", "-nostdin", "-y",
            "-f", "concat", "-safe", "1", "-i", str(concat),
            "-map", "0:a:0", "-c:a", "libmp3lame", "-b:a", "128k",
            "-ar", "24000", "-ac", "1", str(temporary),
        ]
        result = subprocess.run(command, capture_output=True, text=True, check=False)
        if result.returncode != 0:
            raise RuntimeError(f"MP3 assembly failed: {result.stderr.strip()}")
        if not temporary.is_file() or temporary.stat().st_size == 0:
            raise RuntimeError("FFmpeg did not produce an MP3.")
        validation = subprocess.run(
            [ffmpeg, "-hide_banner", "-loglevel", "error", "-xerror", "-nostdin",
             "-i", str(temporary), "-f", "null", "-"],
            capture_output=True, text=True, check=False,
        )
        if validation.returncode != 0:
            raise RuntimeError(f"Invalid MP3: {validation.stderr.strip()}")
        temporary.replace(output)


def parser() -> argparse.ArgumentParser:
    result = argparse.ArgumentParser(description=__doc__)
    result.add_argument("document", type=Path, help="Word document or UTF-8 text file")
    result.add_argument("-o", "--output", type=Path, help="Final MP3 (defaults to next to the document)")
    result.add_argument("--endpoint", default=os.environ.get("AZURE_SPEECH_ENDPOINT"),
                        help="Azure endpoint (or AZURE_SPEECH_ENDPOINT environment variable)")
    result.add_argument("--voice", default=DEFAULT_VOICE)
    result.add_argument("--chunk-size", type=int, default=DEFAULT_CHUNK_SIZE)
    result.add_argument("--attempts", type=int, default=4)
    result.add_argument("--overwrite", action="store_true", help="Replace an existing MP3")
    result.add_argument("--dry-run", action="store_true", help="Extract and count without calling Azure")
    return result


def main() -> int:
    args = parser().parse_args()
    source = args.document.expanduser().resolve()
    if not source.is_file():
        raise ValueError(f"Document not found: {source}")
    output = (args.output or source.with_suffix(".mp3")).expanduser().resolve()
    if output.suffix.lower() != ".mp3":
        raise ValueError("The output file must have the .mp3 extension.")
    if not args.dry_run and output.exists() and not args.overwrite:
        raise ValueError(f"The MP3 already exists: {output}. Use --overwrite.")
    if args.attempts < 1 or args.attempts > 10:
        raise ValueError("--attempts must be between 1 and 10.")
    if not args.endpoint and not args.dry_run:
        raise ValueError("Missing Azure endpoint: use --endpoint or AZURE_SPEECH_ENDPOINT.")
    if args.endpoint:
        parsed = urlparse(args.endpoint)
        if (parsed.scheme != "https" or not parsed.hostname or parsed.username
                or parsed.password or parsed.query or parsed.fragment
                or parsed.path not in {"", "/"}):
            raise ValueError("Expected endpoint: https://your-resource.cognitiveservices.azure.com/")
    print(f"Extracting: {source}", flush=True)
    text = extract_text(source)
    chunks = split_text(text, args.chunk_size)
    if not chunks:
        raise ValueError("The document does not contain readable text.")
    print(
        f"{len(text):,} characters, {len(chunks)} segment(s), voice {args.voice}.\n"
        "Document body and tables included; notes appended at the end. Images excluded.",
        flush=True,
    )
    if args.dry_run:
        print("Dry run complete: no Azure calls, no charges.")
        return 0

    import azure.cognitiveservices.speech as speechsdk
    import imageio_ffmpeg

    ffmpeg = imageio_ffmpeg.get_ffmpeg_exe()
    key = os.environ.get("AZURE_SPEECH_KEY") or getpass.getpass("Azure Speech key (hidden): ")
    if not key.strip():
        raise ValueError("Missing Azure Speech key.")
    config = speechsdk.SpeechConfig(subscription=key, endpoint=args.endpoint)
    config.speech_synthesis_voice_name = args.voice
    config.set_speech_synthesis_output_format(speechsdk.SpeechSynthesisOutputFormat.Riff24Khz16BitMonoPcm)
    synthesizer = speechsdk.SpeechSynthesizer(speech_config=config, audio_config=None)
    identity = json.dumps(
        {"version": 1, "chunks": chunks, "voice": args.voice,
         "endpoint": args.endpoint, "format": "pcm-24000-16-mono"},
        ensure_ascii=False, sort_keys=True,
    )
    fingerprint = hashlib.sha256(identity.encode("utf-8")).hexdigest()
    output.parent.mkdir(parents=True, exist_ok=True)
    cache = output.parent / (output.stem + ".speech-cache") / fingerprint
    cache.mkdir(parents=True, exist_ok=True)
    lock = cache / ".lock"
    try:
        descriptor = os.open(lock, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
    except FileExistsError as error:
        raise RuntimeError(
            f"Conversion already running, or lock left by a stopped process: {lock}. "
            "Only delete this lock after verifying that the other conversion has stopped."
        ) from error
    try:
        with os.fdopen(descriptor, "w") as handle:
            handle.write(str(os.getpid()))
        total_seconds = 0.0
        for index, chunk in enumerate(chunks, 1):
            destination = cache / f"{index:05d}.wav"
            if destination.exists():
                try:
                    duration = wav_duration(destination)
                except (wave.Error, EOFError, ValueError) as error:
                    print(f"Invalid cache ({destination.name}): {error}. Regenerating.", file=sys.stderr)
                    destination.unlink()
                else:
                    total_seconds += duration
                    print(f"[{index}/{len(chunks)}] Reused ({duration:.1f} s)", flush=True)
                    continue
            print(f"[{index}/{len(chunks)}] Synthesizing ({len(chunk)} characters)...", flush=True)
            synthesize_chunk(synthesizer, chunk, destination, key, args.attempts)
            total_seconds += wav_duration(destination)
        print("Assembling and verifying the MP3...", flush=True)
        assemble_mp3(cache, len(chunks), output, ffmpeg)
        print(f"Done: {output}\nDuration: {total_seconds / 60:.1f} minutes", flush=True)
    finally:
        lock.unlink(missing_ok=True)
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\nInterrupted. Run the same command again to resume.", file=sys.stderr)
        sys.exit(130)
    except (ValueError, RuntimeError, OSError, ImportError) as error:
        print(f"Error: {error}", file=sys.stderr)
        sys.exit(1)
