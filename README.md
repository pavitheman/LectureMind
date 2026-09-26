# LectureMind

LectureMind is a locally run study assistant that turns a lecture recording, a PowerPoint deck, or both into revision materials: a transcript, slide notes, key concepts, a summary, flashcards, a multiple-choice quiz, and a visual mind map.

## Features

- Transcribes lecture audio with Whisper.
- Extracts PowerPoint text and can caption slide images.
- Creates a formatted overview, clickable flashcards, and selectable quiz answers with feedback.
- Builds a mind map from the lecture concepts and provides an SVG download.
- Processes uploaded files in a request-specific temporary folder and removes them when processing finishes.
- Runs AI inference locally. Large inputs are summarized in chunks before study materials are generated.

## Model scope

The current implementation uses three pretrained AI models:

| Model | Use |
| --- | --- |
| OpenAI Whisper `base` | Audio transcription |
| Salesforce BLIP `Salesforce/blip-image-captioning-base` | Optional descriptions of slide images |
| Llama 3.2 through Ollama | Summaries, concepts, flashcards, quizzes, and structured mind-map content |

The mind map is generated from Llama's structured output and drawn as SVG. Graphviz is used as an optional renderer when installed; LectureMind has a built-in SVG fallback. Graphviz is not an additional AI model. Earlier project correspondence mentioned a possible additional AI model, so confirm with the client or supervisor whether a fourth model is still a requirement before treating this model scope as final.

## Requirements

- Windows, macOS, or Linux
- Python 3.10 or newer
- FFmpeg available on `PATH` for Whisper audio decoding
- Ollama installed and running, with `llama3.2` downloaded
- Internet access the first time Whisper and BLIP model weights are needed
- Graphviz is optional

## Setup on Windows

Open PowerShell in the `LectureMind` project directory:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

Install FFmpeg and verify `ffmpeg` runs in PowerShell. Install and start Ollama, then download the language model:

```powershell
ollama pull llama3.2
```

Start the web app:

```powershell
python webapp.py
```

Open <http://localhost:5000>. If PowerShell blocks environment activation, run the project interpreter directly:

```powershell
.\.venv\Scripts\python.exe webapp.py
```

The first inference can take several minutes while model weights load or download. Keep Ollama running while using the app. Set `LECTUREMIND_LLM_MODEL` if you have pulled a different Ollama model name.

## Run the tests

With the project environment activated:

```powershell
python -m unittest discover -s tests -v
```

The automated tests use Python's built-in `unittest` framework and mock AI inference, so they do not download model weights or require Ollama/FFmpeg. They cover parsers, long-input chunking, mind-map fallback output, upload validation, the audio/slides/combined routes, oversized and damaged uploads, and partial model failures. They do not replace a manual end-to-end run with the local models.

## Inputs and limits

- Audio: `.mp3`, `.wav`, `.m4a`, or `.mp4`
- Slides: `.pptx`
- Audio and slides can be submitted together.
- Maximum combined upload size: 500 MB.
- The app currently generates materials for the active request; it does not provide a persistent lecture library or user accounts.

## Project files

- `webapp.py` — Flask routes, upload validation, and results assembly.
- `app.py` — model-backed processing pipeline, parsers, and mind-map rendering.
- `templates/upload.html` — upload page.
- `templates/results.html` — overview, mind map, flashcards, quiz, transcript, and slide notes.
- `tests/test_pipeline.py` — unit and mocked Flask route tests.
- `prd.txt` — project requirements document.

## Project progress

This timeline records dated checkpoints supported by the available Git history and project correspondence. The Git repository contains no commit records for each week between August and September, so this is a milestone history, not a claim of weekly commits or a complete week-by-week work log.

| Date | Evidence-backed milestone |
| --- | --- |
| May 9, 2026 | Git history records the initial repository and the PRD. |
| May 10, 2026 | Git history records the Whisper transcription feature and its merge. |
| May 31, 2026 | Git history records Llama/Ollama summary and flashcard generation. |
| August 20, 2026 | Client correspondence records the assignment and final-submission expectations. |
| August 21, 2026 | Client correspondence discusses the AI model count and mind-map direction. |
| September 9, 2026 | Preliminary-report feedback identifies areas to strengthen, including technical design, work plan, evaluation, and prototype visuals. |
| September 15, 2026 | Project correspondence describes the three-model scope and the mind map as Llama-generated content rendered with Graphviz. |
| September 26, 2026 | Current stabilization work adds the Flask interface, mind-map and study-material flows, dependency setup, ignore rules, and automated tests. These changes are local until committed. |

## Current limitations

- Output quality depends on the recording, slide readability, and locally available model performance.
- BLIP captions are optional; slide text extraction can continue without them.
- Automated tests mock model calls, so verify the complete upload-to-results flow locally before a demonstration or release.
- Confirm whether the client still expects a fourth pretrained AI model; the current implementation and preliminary-report scope describe three.
