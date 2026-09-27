"""Flask web interface for the local LectureMind pipeline."""

import logging
import os
import tempfile
import zipfile
from pathlib import Path

from flask import Flask, render_template, request
from werkzeug.exceptions import RequestEntityTooLarge
from werkzeug.utils import secure_filename

from app import (
    PipelineError,
    build_context,
    generate_flashcards,
    generate_mindmap,
    generate_quiz,
    generate_summary,
    parse_flashcards,
    parse_overview,
    parse_quiz,
    render_mindmap_svg,
    transcribe,
    process_slides,
    user_safe_error,
)


logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("lecturemind")

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 500 * 1024 * 1024
app.config["UPLOAD_FOLDER"] = str(Path(__file__).resolve().parent / "uploads")
Path(app.config["UPLOAD_FOLDER"]).mkdir(parents=True, exist_ok=True)

AUDIO_EXTENSIONS = {"mp3", "mp4", "wav", "m4a"}
SLIDE_EXTENSIONS = {"pptx"}


def _file_extension(upload):
    filename = secure_filename(upload.filename or "")
    if not filename or "." not in filename:
        return filename, ""
    return filename, filename.rsplit(".", 1)[1].lower()


def _validate_upload(upload, allowed, label):
    filename, extension = _file_extension(upload)
    if not filename or extension not in allowed:
        allowed_text = ", ".join(f".{item}" for item in sorted(allowed))
        raise PipelineError(f"Choose a {label} file in one of these formats: {allowed_text}.")
    return filename, extension


def _valid_powerpoint(path):
    try:
        with zipfile.ZipFile(path) as archive:
            names = set(archive.namelist())
            return "[Content_Types].xml" in names and "ppt/presentation.xml" in names
    except (OSError, zipfile.BadZipFile):
        return False


def _render_upload(error=None):
    return render_template("upload.html", error=error), 400 if error else 200


@app.route("/")
def index():
    return render_template("upload.html")


@app.errorhandler(RequestEntityTooLarge)
def too_large(_error):
    return _render_upload("The selected files exceed the 500 MB upload limit. Choose smaller files and try again.")[0], 413


@app.route("/process", methods=["POST"])
def process():
    audio = request.files.get("audio")
    slides = request.files.get("slides")
    has_audio = bool(audio and audio.filename)
    has_slides = bool(slides and slides.filename)
    if not has_audio and not has_slides:
        return _render_upload("Choose an audio recording, a PowerPoint file, or both.")

    try:
        audio_name = _validate_upload(audio, AUDIO_EXTENSIONS, "audio/video") if has_audio else None
        slides_name = _validate_upload(slides, SLIDE_EXTENSIONS, "PowerPoint") if has_slides else None
    except PipelineError as exc:
        return _render_upload(str(exc))

    transcript = None
    slide_content = None
    processing_errors = []

    # Each request gets an isolated directory; uploads and extracted slide images
    # are removed automatically after all processing and rendering is complete.
    with tempfile.TemporaryDirectory(prefix="lecturemind-", dir=app.config["UPLOAD_FOLDER"]) as request_dir:
        request_path = Path(request_dir)
        if has_audio:
            audio_path = request_path / f"lecture.{audio_name[1]}"
            audio.save(audio_path)
            if audio_path.stat().st_size == 0:
                return _render_upload("The audio file is empty. Choose a valid recording and try again.")
            try:
                transcript = transcribe(audio_path)
            except Exception as exc:
                logger.exception("Audio transcription failed")
                processing_errors.append(user_safe_error(exc))

        if has_slides:
            slides_path = request_path / "slides.pptx"
            slides.save(slides_path)
            if slides_path.stat().st_size == 0 or not _valid_powerpoint(slides_path):
                processing_errors.append("The PowerPoint file is empty or damaged. Please export a valid .pptx file and try again.")
            else:
                try:
                    slide_content = process_slides(slides_path, work_dir=request_path)
                except Exception as exc:
                    logger.exception("PowerPoint processing failed")
                    processing_errors.append(user_safe_error(exc))

        if not transcript and not slide_content:
            return _render_upload(processing_errors[0] if processing_errors else "No usable content was found in the uploaded files.")

        try:
            context = build_context(transcript, slide_content)
        except Exception as exc:
            logger.exception("Lecture context preparation failed")
            context = None
            processing_errors.append(user_safe_error(exc))

        results = {"summary": "", "flashcards": "", "quiz": "", "mindmap": {"title": "Lecture concepts", "concepts": []}}
        if context:
            generators = (
                ("summary", lambda: generate_summary(context=context)),
                ("flashcards", lambda: generate_flashcards(context=context)),
                ("quiz", lambda: generate_quiz(context=context, summary=results["summary"])),
            )
            for name, generate in generators:
                try:
                    results[name] = generate()
                except Exception as exc:
                    logger.exception("Study material generation failed: %s", name)
                    processing_errors.append(user_safe_error(exc))

            try:
                results["mindmap"] = generate_mindmap(context=context, summary=results["summary"])
            except Exception as exc:
                logger.exception("Mind map generation failed")
                processing_errors.append("The mind map could not be generated. Other study materials are still available.")

        try:
            mindmap_svg = render_mindmap_svg(results["mindmap"])
        except Exception:
            logger.exception("Mind map rendering failed")
            mindmap_svg = ""
            processing_errors.append("The mind map could not be rendered. Other study materials are still available.")

        return render_template(
            "results.html",
            transcript=transcript,
            slide_content=slide_content,
            summary=results["summary"],
            overview=parse_overview(results["summary"]),
            flashcards=results["flashcards"],
            flashcard_items=parse_flashcards(results["flashcards"]),
            quiz=results["quiz"],
            quiz_items=parse_quiz(results["quiz"]),
            mindmap=results["mindmap"],
            mindmap_svg=mindmap_svg,
            processing_errors=list(dict.fromkeys(processing_errors)),
        )


if __name__ == "__main__":
    print("LectureMind is starting at http://localhost:5000")
    print("Models load when needed. Keep Ollama running with the configured model pulled.")
    app.run(debug=False, threaded=True, port=5000)
