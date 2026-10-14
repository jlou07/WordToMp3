# WordToMp3

**Turn a long Word document into one MP3 using Python and Azure Speech.**

Word's **Read Aloud** plays a document inside Word. This command-line application
creates a standalone audio file instead; it does not record Word's playback.
Word's **Dictate** feature does the reverse: speech to text.

Repository: [jlou07/WordToMp3](https://github.com/jlou07/WordToMp3).

## Requirements

- Windows and Microsoft Word installed for `.doc` and `.docx` extraction.
- Python; local checks were run with Python 3.11.9.
- An Azure Speech resource, its HTTPS endpoint, and a valid key for synthesis.
- Network access and sufficient space for intermediate WAV files.

UTF-8 `.txt` input does not require Word. Other operating systems have not been
validated end to end. The original document is opened read-only and is not saved.

## Installation

In PowerShell, from the project directory:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade -r requirements-tools.txt
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\convert.ps1 --help
```

FFmpeg is supplied through `imageio-ffmpeg`; no separate system installation
is required. The launcher uses the Python executable inside `.venv`.

The installer-tooling baseline in [requirements-tools.txt](requirements-tools.txt)
requires Python 3.10 or newer. Upgrade the tooling before installing application
dependencies, including when reusing an older virtual environment. These tools
are not application runtime dependencies; review security updates regularly.

## Try it without Azure or charges

```powershell
.\convert.ps1 .\docs\article\sample-narration.txt --dry-run
```

This extracts and counts text and chunks. It does not request a key, call Azure,
write audio, validate credentials, or check voice availability.

## Convert a document

```powershell
$env:AZURE_SPEECH_ENDPOINT = "https://your-resource.cognitiveservices.azure.com/"
.\convert.ps1 "C:\Documents\book.docx" -o "C:\Audio\book.mp3"
```

The document path is required. There is no search through personal folders.
Without `-o`, the MP3 is created next to the input document.

If `AZURE_SPEECH_KEY` is unset, the application requests the key using a hidden
prompt. Do not put a real key in source code, a published command, shell history,
or a screenshot. The repository contains no personal resource endpoint.

**Synthesis can incur Azure charges.** Text chunks are sent to your configured
Azure Speech resource; the Word document itself is not uploaded as a file.

```powershell
.\convert.ps1 "C:\Documents\book.doc" --dry-run
.\convert.ps1 "C:\Documents\book.docx" --voice "en-US-JennyNeural"
.\convert.ps1 "C:\Documents\book.docx" --chunk-size 1500 --attempts 6
.\convert.ps1 "C:\Documents\book.docx" --overwrite
```

The default narration voice remains `fr-FR-DeniseNeural` to preserve the original
application behavior. Documentation and CLI messages are English; the voice is
an independent choice. Use `--voice` for the language of your document and verify
availability on your resource.

| Option | Default | Purpose |
|---|---|---|
| `document` | Required | DOC, DOCX, or UTF-8 TXT input. |
| `-o`, `--output` | Next to input, with `.mp3` extension | Output path. |
| `--endpoint` | `AZURE_SPEECH_ENDPOINT` | Required for synthesis, optional for dry runs. |
| `--voice` | `fr-FR-DeniseNeural` | Narration voice. |
| `--chunk-size` | `2000` | Characters per chunk; range 100-3000. |
| `--attempts` | `4` | Total attempts per chunk; range 1-10. |
| `--overwrite` | Off | Replace an existing MP3 only after successful encoding and checks. |
| `--dry-run` | Off | Extract and count without synthesis. |

## Long documents and recovery

- Extracts document body and table text; appends footnotes and endnotes.
  Images, scanned text, floating text boxes, headers, and footers are not extracted.
- Normalizes whitespace and splits preferentially at paragraphs and sentences.
  Very long unbroken tokens can be split at the character limit.
- Produces 24 kHz, 16-bit mono WAV chunks, then encodes one 128 kbit/s MP3.
- Rejects empty, incomplete, unexpected-format, or nine-minute-and-longer WAVs.
  Check current Azure limits for your API and voice; this is not a guarantee
  against all possible truncation.
- Decodes the finished MP3 as a technical check before publishing the output.
- Resumes using the adjacent `*.speech-cache` folder. Run the same command again;
  valid cached chunks are reused without new synthesis calls for those chunks.
- Text, voice, endpoint, and resulting chunk boundaries determine cache identity.
  Moving the output to another location changes where the cache is searched.
- Transient Azure errors are retried, with four **total attempts** by default.
  Authentication and configuration errors stop the conversion.
- An interruption preserves completed chunks. An existing MP3 is not replaced
  unless `--overwrite` is specified and finalization succeeds.
- A `.lock` file contains the conversion PID. After a crash, only remove it
  after verifying that the associated conversion is no longer running.

The cache contains unencrypted audio of your document. Keep it private.
Deleting it means a future conversion may require synthesizing the text again.
Technical validation does not check pronunciation or guarantee semantic completeness:
listen to the result before sharing it.

## Private configuration and publishing

[.env.example](.env.example) documents variable names and placeholder values only.
The application does **not** automatically load `.env` files. Configure environment
variables in PowerShell or use `--endpoint`, and use the hidden prompt for the key.

[.gitignore](.gitignore) excludes real `.env` files, Word/PDF documents, MP3/WAV
files, speech caches, virtual environments, and common key/certificate files.
Keep private inputs in the ignored `local` directory. Plain-text inputs outside
that directory are not automatically excluded: review every file before staging.

The application does not save your key to its cache. Logs may contain document
paths, and audio contains the spoken content. Git exclusions are not a substitute
for reviewing staged files. Rotate any credential that has been disclosed.

## Local tests without Azure charges

```powershell
.\.venv\Scripts\python.exe -m unittest -v test_convert.py
```

The suite covers text splitting, WAV validation, actual local FFmpeg encoding,
output preservation on encoding failure, mocked Azure errors, and configuration
requirements. It does not perform a live Azure conversion or automate a real
Word document end to end.
