# LectureMind

I built LectureMind as a local study assistant that turns a lecture recording, a PowerPoint presentation, or both into revision materials. It produces a transcript, slide notes, key concepts, a summary, flashcards, a multiple-choice quiz, and a visual mind map.

## What I built

- I use Whisper to transcribe lecture audio.
- I extract text from PowerPoint slides and use BLIP to caption embedded images.
- I generate an overview, flashcards, and a multiple-choice quiz from the lecture content.
- I create a mind map from the lecture concepts and let users download it as an SVG.
- I process each upload in a temporary request folder and remove the files when processing finishes.
- I run the AI inference locally. For long lectures, I summarize the content in chunks before generating study materials.

## AI models

I use three pretrained models in the current implementation:

| Model | What I use it for |
| --- | --- |
| OpenAI Whisper `base` | Audio transcription |
| Salesforce BLIP `Salesforce/blip-image-captioning-base` | Optional descriptions of slide images |
| Llama 3.2 through Ollama | Summaries, concepts, flashcards, quizzes, and structured mind-map content |

My preliminary report specifies these three models. I reuse Llama's structured output for the mind map, then render it as SVG. Graphviz is an optional renderer; if it is unavailable, I use the built-in SVG renderer. Graphviz is not an AI model, so the mind map does not add a fourth model to the pipeline.

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

Install FFmpeg and verify that `ffmpeg` runs in PowerShell. Install and start Ollama, then download the language model:

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

I wrote the automated tests with Python's built-in `unittest` framework. They mock AI inference, so they do not download model weights or require Ollama or FFmpeg. They cover parsers, long-input chunking, mind-map fallback output, upload validation, audio and slide routes, oversized or damaged uploads, and partial model failures. I also ran a real combined audio-and-slide upload through all three models; the automated tests do not replace that end-to-end check on another machine.

## Inputs and limits

- Audio: `.mp3`, `.wav`, `.m4a`, or `.mp4`
- Slides: `.pptx`
- I allow audio and slides to be submitted together.
- Maximum combined upload size: 500 MB.
- I generate materials for the active request; I have not added user accounts or a persistent lecture library.

## Project files

- `webapp.py` — I use Flask for upload routes, validation, and results assembly.
- `app.py` — I implement the local AI pipeline, parsers, and mind-map renderer here.
- `templates/upload.html` — upload page.
- `templates/results.html` — overview, mind map, flashcards, quiz, transcript, and slide notes.
- `tests/test_pipeline.py` — unit and mocked Flask route tests.
- `prd.txt` — project requirements document.

## Project progress

I list dated milestones that are supported by the repository history. The repository does not record a code change for every week, so I have not filled gaps with unverified weekly entries.

| Date | Progress |
| --- | --- |
| May 9, 2026 | I created the repository and added the initial project requirements. |
| May 10, 2026 | I added Whisper audio transcription. |
| May 31, 2026 | I connected Ollama and Llama 3.2 for summaries and flashcards. |
| September 26, 2026 | I merged the Flask app, lecture study tools, mind map, setup documentation, and automated tests into `main` in PR #2. |
| September 26, 2026 | I fixed BLIP slide-image captioning and updated the model-scope documentation in PR #3, which is awaiting review. |

## Current limitations

- Output quality depends on the recording, slide readability, and local model performance.
- BLIP captioning is optional; I continue extracting slide text if image captioning is unavailable.
- Automated tests mock model calls. I recommend running the full upload flow with the local models before a demonstration or release.
- The mind map reuses Llama and does not use a separate pretrained model.
