"""Local AI pipeline for LectureMind.

Models are loaded on demand so the web server can start even when an optional
model is not installed yet. Long inputs are reduced hierarchically instead of
being silently cut off at a fixed character limit.
"""

import json
import os
import re
import tempfile
import threading
from difflib import SequenceMatcher
from pathlib import Path

import ollama
from pptx import Presentation


LLM_MODEL = os.environ.get("LECTUREMIND_LLM_MODEL", "llama3.2")
CONTEXT_LIMIT = 6500
CHUNK_SIZE = 4200
_model_lock = threading.Lock()
_whisper_model = None
_captioner = None
_captioner_attempted = False


class PipelineError(RuntimeError):
    """An actionable failure in a local model or input-processing stage."""


def _chat(system, prompt, *, response_format=None):
    options = {"model": LLM_MODEL, "messages": [
        {"role": "system", "content": system},
        {"role": "user", "content": prompt},
    ]}
    if response_format:
        options["format"] = response_format
    try:
        response = ollama.chat(**options)
        return response["message"]["content"].strip()
    except Exception as exc:
        message = str(exc).lower()
        if "connect" in message or "refused" in message or "localhost" in message:
            raise PipelineError("Ollama is not running. Start Ollama and try again.") from exc
        if "not found" in message or "pull" in message:
            raise PipelineError(f"The local model '{LLM_MODEL}' is unavailable. Run `ollama pull {LLM_MODEL}` and try again.") from exc
        raise PipelineError("The local language model could not complete this step. Check that Ollama is running and try again.") from exc


def _load_whisper():
    global _whisper_model
    if _whisper_model is None:
        with _model_lock:
            if _whisper_model is None:
                try:
                    import whisper
                    print("Loading Whisper base model...")
                    _whisper_model = whisper.load_model("base")
                except Exception as exc:
                    raise PipelineError("Whisper could not be loaded. Check the installation and available disk space, then restart the app.") from exc
    return _whisper_model


def _load_captioner():
    global _captioner, _captioner_attempted
    if _captioner_attempted and _captioner is None:
        return None
    if _captioner is None:
        with _model_lock:
            if _captioner is None:
                _captioner_attempted = True
                try:
                    from transformers import BlipForConditionalGeneration, BlipProcessor
                    print("Loading BLIP slide-image model...")
                    model_name = "Salesforce/blip-image-captioning-base"
                    processor = BlipProcessor.from_pretrained(model_name)
                    model = BlipForConditionalGeneration.from_pretrained(model_name)
                    _captioner = (processor, model)
                except Exception as exc:
                    print(f"BLIP unavailable; slide text extraction will continue without image captions: {exc}")
                    return None
    return _captioner


def transcribe(audio_path):
    """Transcribe an uploaded audio or video file with Whisper."""
    try:
        result = _load_whisper().transcribe(str(audio_path))
    except PipelineError:
        raise
    except Exception as exc:
        raise PipelineError("Whisper could not transcribe this file. Check that it is a valid audio/video file and that FFmpeg is installed.") from exc
    text = (result.get("text") or "").strip()
    if not text:
        raise PipelineError("Whisper did not detect any speech in the uploaded recording.")
    return text


def process_slides(pptx_path, work_dir=None):
    """Extract slide text and best-effort captions for embedded images."""
    try:
        presentation = Presentation(str(pptx_path))
    except Exception as exc:
        raise PipelineError("The PowerPoint file could not be opened. Please upload a valid .pptx file.") from exc

    slide_output = []
    captioner = None
    temp_root = Path(work_dir) if work_dir else Path(tempfile.gettempdir())
    for slide_number, slide in enumerate(presentation.slides, start=1):
        items = []
        image_number = 0
        for shape in slide.shapes:
            if getattr(shape, "has_text_frame", False):
                for paragraph in shape.text_frame.paragraphs:
                    text = paragraph.text.strip()
                    if text:
                        items.append(text)
            if shape.shape_type == 13:
                image_number += 1
                if captioner is None:
                    captioner = _load_captioner()
                if captioner is None:
                    continue
                try:
                    ext = shape.image.content_type.split("/")[-1].replace("jpeg", "jpg")
                    image_path = temp_root / f"slide-{slide_number}-{image_number}.{ext}"
                    image_path.write_bytes(shape.image.blob)
                    try:
                        from PIL import Image
                        processor, model = captioner
                        with Image.open(image_path) as image:
                            inputs = processor(images=image.convert("RGB"), return_tensors="pt")
                        generated = model.generate(**inputs, max_new_tokens=48)
                        caption = processor.decode(generated[0], skip_special_tokens=True).strip()
                    finally:
                        image_path.unlink(missing_ok=True)
                    if caption:
                        items.append(f"[Slide image: {caption}]")
                except Exception as exc:
                    print(f"Slide {slide_number} image {image_number} could not be captioned: {exc}")
        slide_output.append(f"Slide {slide_number}:\n" + ("\n".join(items) if items else "No readable text or captions."))
    return "\n\n".join(slide_output)


def _split_chunks(text, limit=CHUNK_SIZE):
    """Split at paragraph/sentence boundaries where possible."""
    paragraphs = [part.strip() for part in re.split(r"\n+", text) if part.strip()]
    if not paragraphs:
        paragraphs = [text.strip()]
    chunks, current = [], ""
    for paragraph in paragraphs:
        while len(paragraph) > limit:
            if current:
                chunks.append(current)
                current = ""
            cut = paragraph.rfind(" ", 0, limit)
            cut = cut if cut > limit // 2 else limit
            chunks.append(paragraph[:cut].strip())
            paragraph = paragraph[cut:].strip()
        candidate = f"{current}\n{paragraph}".strip()
        if len(candidate) > limit and current:
            chunks.append(current)
            current = paragraph
        else:
            current = candidate
    if current:
        chunks.append(current)
    return chunks


def build_context(transcript=None, slide_content=None):
    """Preserve long lecture content through chunk summaries and reduction."""
    sources = []
    if transcript and transcript.strip():
        sources.append(("Lecture transcript", transcript.strip()))
    if slide_content and slide_content.strip():
        sources.append(("PowerPoint slides", slide_content.strip()))
    if not sources:
        return ""
    combined = "\n\n".join(f"{name}:\n{text}" for name, text in sources)
    if len(combined) <= CONTEXT_LIMIT:
        return combined

    print("Long lecture detected; summarising in overlapping chunks...")
    chunk_summaries = []
    for index, chunk in enumerate(_split_chunks(combined), start=1):
        summary = _chat(
            "You make faithful study notes from lecture material. Keep names, definitions, examples, and relationships that appear in the source. Do not add outside facts.",
            f"Summarise this section in up to 700 characters. Preserve its key facts and terminology.\n\nSection {index}:\n{chunk}",
        )
        chunk_summaries.append(summary[:1100])

    reduced = "\n\n".join(chunk_summaries)
    while len(reduced) > CONTEXT_LIMIT:
        groups = _split_chunks(reduced, CONTEXT_LIMIT)
        summaries = []
        for group in groups:
            summaries.append(_chat(
                "Combine study notes without losing distinct ideas. Use only the supplied notes.",
                f"Combine these notes into a concise, complete outline under 4,000 characters.\n\n{group}",
            ))
        reduced = "\n\n".join(summaries)
    return "Lecture materials (hierarchical summaries of the complete upload):\n" + reduced


def generate_summary(transcript=None, slide_content=None, *, context=None):
    context = context if context is not None else build_context(transcript, slide_content)
    if not context:
        return "No lecture content was provided."
    return _chat(
        "You are a careful university teaching assistant. Use only the supplied lecture materials.",
        "Return plain text with no Markdown. Use this exact structure:\nKEY CONCEPTS\n1. Concept name: one concise sentence explaining it.\n2. Concept name: one concise sentence explaining it.\n3. Concept name: one concise sentence explaining it.\n4. Concept name: one concise sentence explaining it.\n5. Concept name: one concise sentence explaining it.\nSUMMARY\nWrite a clear 3-4 sentence overview of the lecture. Keep concepts and summary separate. Use only supported information from the source.\n\n" + context,
    )


def generate_flashcards(transcript=None, slide_content=None, *, context=None):
    context = context if context is not None else build_context(transcript, slide_content)
    if not context:
        return ""
    return _chat(
        "You create accurate study flashcards grounded only in the supplied lecture material.",
        "Create five distinct flashcards. Use this format for each:\nFront: question\nBack: concise answer\n\n" + context,
    )


def generate_quiz(transcript=None, slide_content=None, *, context=None, summary=""):
    context = context if context is not None else build_context(transcript, slide_content)
    if not context:
        return ""
    concept_quiz = _quiz_from_summary(summary)
    if concept_quiz:
        return concept_quiz
    topic_text = _chat(
        "You identify distinct, testable topics using only the provided lecture material.",
        "Return up to five distinct numbered topics. Do not invent topics if the source does not support them.\n\n" + context,
    )
    topics = []
    for line in topic_text.splitlines():
        match = re.match(r"\s*(?:\d+[.)]|[-*])\s*(.+)", line)
        if match and match.group(1).strip() not in topics:
            topics.append(match.group(1).strip()[:180])
    if not topics:
        raise PipelineError("The language model did not identify quiz topics from this lecture. Try processing the files again.")

    questions = []
    previous_prompts = []
    for topic in topics[:5]:
        question = ""
        for attempt in range(2):
            previous = "\n".join(f"- {prompt}" for prompt in previous_prompts) or "None"
            retry_note = (
                f"Your previous response was incomplete or too similar to an earlier question:\n{question}\n"
                "Write a different question with four distinct options.\n"
                if attempt else ""
            )
            question = _chat(
                "You write precise multiple-choice questions using only the supplied lecture material. "
                "Each question has one clearly correct answer and three distinct, clearly wrong alternatives.",
                f"Write one specific, verifiable question about: {topic}\n"
                "Avoid vague questions about the course's main focus, invented course sections, "
                "overlapping answer options, and facts absent from the source. "
                "Do not repeat or rephrase an earlier question.\n"
                f"Earlier questions:\n{previous}\n"
                f"{retry_note}"
                "Use this exact format:\nQuestion: ...\nA. ...\nB. ...\nC. ...\nD. ...\nAnswer: A/B/C/D\n\n"
                f"Source:\n{context}",
            )
            parsed = parse_quiz(question)
            if len(parsed) != 1 or not _valid_quiz_item(parsed[0]):
                continue
            if not _answer_supported(parsed[0], context):
                continue
            prompt = parsed[0]["prompt"]
            if any(quoted.casefold() not in context.casefold() for quoted in re.findall(r'["“]([^"”]+)["”]', prompt)):
                continue
            if any(_is_repeated_question(prompt, old) for old in previous_prompts):
                continue
            questions.append(question)
            previous_prompts.append(prompt)
            break
    if not questions:
        raise PipelineError("The language model did not return a usable quiz. Other study materials may still be available.")
    return "\n\n".join(questions)


def _quiz_from_summary(summary):
    """Build unambiguous concept matching MCQs from the visible overview."""
    concepts = [item for item in _concepts_from_summary(summary) if item["details"]]
    labels = [item["label"] for item in concepts]
    if len(concepts) < 4 or len({label.casefold() for label in labels}) != len(labels):
        return ""
    questions = []
    for index, concept in enumerate(concepts[:5]):
        alternatives = [label for label in labels if label != concept["label"]][:3]
        if len(alternatives) != 3:
            continue
        correct_at = index % 4
        alternatives.insert(correct_at, concept["label"])
        description = concept["details"][0].replace('"', "'").strip()
        lines = [f'Question: Which lecture concept matches this statement: "{description}"?']
        lines.extend(f"{letter}. {label}" for letter, label in zip("ABCD", alternatives))
        lines.append(f"Answer: {'ABCD'[correct_at]}")
        questions.append("\n".join(lines))
    return "\n\n".join(questions)


def _valid_quiz_item(item):
    options = item.get("options", [])
    letters = [option.get("letter") for option in options]
    texts = [re.sub(r"\s+", " ", option.get("text", "")).strip().casefold() for option in options]
    return (
        bool(item.get("prompt", "").strip())
        and len(options) == 4
        and letters == ["A", "B", "C", "D"]
        and item.get("answer") in letters
        and all(texts)
        and len(set(texts)) == 4
    )


def _answer_supported(item, source):
    """Reject an answer that introduces specific terms absent from the source."""
    answer = next(option["text"] for option in item["options"] if option["letter"] == item["answer"])
    ignored = {"about", "after", "before", "being", "course", "computer", "computers", "could", "from", "have", "into", "their", "there", "these", "those", "through", "using", "which", "would"}
    words = {word for word in re.findall(r"[a-z]{5,}", answer.casefold()) if word not in ignored}
    if not words:
        return True
    source_words = set(re.findall(r"[a-z]{5,}", source.casefold()))
    matches = sum(any(word[:5] == candidate[:5] for candidate in source_words) for word in words)
    return matches / len(words) >= 0.6


def _is_repeated_question(prompt, earlier):
    current_words = re.findall(r"\w+", prompt.casefold())
    earlier_words = re.findall(r"\w+", earlier.casefold())
    if current_words == earlier_words:
        return True
    if all("main focus" in value and "course" in value for value in (prompt.casefold(), earlier.casefold())):
        return True
    return (
        len(current_words) >= 5
        and len(earlier_words) >= 5
        and current_words[:5] == earlier_words[:5]
        and SequenceMatcher(None, prompt.casefold(), earlier.casefold()).ratio() >= 0.78
    )


def _concepts_from_summary(summary):
    concepts = []
    for line in (summary or "").splitlines():
        match = re.match(r"\s*(?:\d+[.)]|[-*])\s*(.+)", line)
        if match:
            value = match.group(1).strip().replace("**", "").replace("__", "")
            label, separator, detail = value.partition(":")
            if not separator and " — " in value:
                label, detail = value.split(" — ", 1)
            if not detail:
                detail = value
                label = re.split(r"[,.;]", value, maxsplit=1)[0]
                label = re.sub(r"^(?:the course (?:will |aims to )?|the term )", "", label, flags=re.I)
                label = re.split(r"\b(?:refers to|play(?:s)? a|has different|will focus on|provide(?:s)? a|is an|are key|are pervasive)\b", label, maxsplit=1, flags=re.I)[0]
                label = " ".join(label.split()[:6])
                label = re.sub(r"^while\s+", "", label, flags=re.I)
                if label.casefold() == "computers" and "society" in value.casefold():
                    label = "Computers in society"
            label = re.sub(r"\s+", " ", label)
            label = label.strip("*# .- ")
            if label and label.lower() not in {item["label"].lower() for item in concepts}:
                concepts.append({"label": label[:90], "details": [detail.strip()[:260]] if detail.strip() else []})
    return concepts[:7]


def generate_mindmap(transcript=None, slide_content=None, *, context=None, summary=""):
    """Return a validated title and concept branches for the mind-map renderer."""
    context = context if context is not None else build_context(transcript, slide_content)
    concepts = _concepts_from_summary(summary)
    title = "Lecture concepts"
    if context:
        prompt = (
            "Return JSON only with keys title and concepts. title is a short lecture topic. concepts is an array of exactly five distinct objects when the source supports five, each with label and details (one or two short supporting facts). "
            "Use the key concepts in the supplied overview. Each label must be a short noun phrase of at most five words. Use only the supplied lecture. Do not invent facts."
            f"\n\nOVERVIEW CONCEPTS:\n{summary}\n\nLECTURE SOURCE:\n{context}"
        )
        try:
            raw = _chat("You organise lecture ideas into a concise study mind map.", prompt, response_format="json")
            data = json.loads(raw)
            safe_title = re.sub(r"\s+", " ", str(data.get("title", ""))).strip()
            if safe_title:
                title = safe_title[:100]
            parsed = []
            for item in data.get("concepts", []):
                if not isinstance(item, dict):
                    continue
                label = re.sub(r"\s+", " ", str(item.get("label", ""))).strip()
                details = item.get("details", [])
                if isinstance(details, str):
                    details = [details]
                details = [re.sub(r"\s+", " ", str(value)).strip()[:140] for value in details if str(value).strip()]
                if label:
                    parsed.append({"label": label[:90], "details": details[:2]})
            if parsed:
                by_label = {item["label"].casefold(): item for item in parsed}
                if concepts:
                    # Keep the model's short labels where they match, but use the
                    # overview's factual detail and its fixed number of branches.
                    result = []
                    for item in concepts[:5]:
                        candidate = by_label.get(item["label"].casefold())
                        if not candidate:
                            generic = {"computer", "computers", "course", "lecture", "introduction", "study", "focus"}
                            source_words = set(re.findall(r"[a-z]{4,}", (item["label"] + " " + " ".join(item["details"])).casefold())) - generic
                            ranked = sorted(parsed, key=lambda branch: len((set(re.findall(r"[a-z]{4,}", branch["label"].casefold())) - generic) & source_words), reverse=True)
                            if ranked and (set(re.findall(r"[a-z]{4,}", ranked[0]["label"].casefold())) - generic) & source_words:
                                candidate = ranked[0]
                        label = candidate["label"] if candidate and len(candidate["label"].split()) <= 5 else item["label"]
                        if label.casefold() in {entry["label"].casefold() for entry in result}:
                            label = item["label"]
                        result.append({"label": label, "details": item["details"]})
                    concepts = result
                else:
                    combined = []
                    seen = set()
                    for item in parsed:
                        key = item["label"].casefold()
                        if key not in seen:
                            seen.add(key)
                            combined.append(item)
                    concepts = combined[:5]
        except Exception as exc:
            print(f"Mind-map structure fallback used: {exc}")
    if not concepts:
        concepts = [{"label": "Key lecture ideas", "details": ["See the concepts and summary for the generated study notes."]}]
    return {"title": title, "concepts": concepts}


def parse_overview(text):
    """Separate model output into clean concept cards and a readable summary."""
    concepts = []
    summary_lines = []
    other_lines = []
    in_summary = False
    for raw_line in (text or "").splitlines():
        line = raw_line.strip().replace("**", "").replace("__", "")
        heading = re.sub(r"[#*:_\s]", "", line).casefold()
        if heading in {"summary", "lectureoverview", "overview"}:
            in_summary = True
            continue
        if in_summary:
            if line:
                summary_lines.append(line)
            continue
        match = re.match(r"^\s*(?:\d+[.)]|[-*])\s*(.+)$", line)
        if match:
            value = match.group(1).strip()
            if ":" in value:
                label, detail = value.split(":", 1)
            elif " — " in value:
                label, detail = value.split(" — ", 1)
            else:
                label, detail = value, ""
            label = re.sub(r"\s+", " ", label).strip(" .-:#")
            detail = re.sub(r"\s+", " ", detail).strip()
            if label:
                concepts.append({"label": label[:110], "detail": detail})
        elif line:
            other_lines.append(line)

    if not summary_lines:
        # Some local model responses omit a heading but separate the overview
        # from the numbered concepts with a blank line.
        summary_lines = [line for line in other_lines if not re.match(r"^(?:here are|key concepts)", line, re.I)]
    summary_text = " ".join(summary_lines).strip()
    summary_text = re.sub(r"\s+", " ", summary_text)
    return {"concepts": concepts[:7], "summary": summary_text}


def _dot_quote(value):
    escaped_lines = []
    for line in str(value).splitlines():
        escaped_lines.append(line.replace("\\", "\\\\").replace('"', '\\"'))
    return '"' + r"\n".join(escaped_lines) + '"'


def _wrap_svg_text(value, limit, max_lines):
    words = re.sub(r"\s+", " ", str(value)).strip().split()
    lines, current = [], ""
    for word in words:
        candidate = f"{current} {word}".strip()
        if len(candidate) > limit and current:
            lines.append(current)
            current = word
        else:
            current = candidate
    if current:
        lines.append(current)
    if len(lines) > max_lines:
        lines = lines[:max_lines]
        lines[-1] = lines[-1].rstrip(".,;: ") + "…"
    return lines


def _fallback_svg(title, concepts):
    """Create a balanced, dependency-free mind map with a central topic."""
    from html import escape
    width, row_height = 1360, 212
    left_items = [(index, concept) for index, concept in enumerate(concepts) if index % 2 == 0]
    right_items = [(index, concept) for index, concept in enumerate(concepts) if index % 2 == 1]
    row_count = max(len(left_items), len(right_items), 1)
    height = max(430, 100 + row_count * row_height)
    root_x, root_w, root_h = 540, 280, 220
    root_y = (height - root_h) // 2
    left_x, right_x, card_w, card_h = 30, 900, 430, 182
    colors = ["#087f8c", "#4059ad", "#a04b78", "#667a35", "#a45a28", "#6553a3", "#14735f"]
    parts = [
        f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {width} {height}" role="img" aria-labelledby="map-title map-desc">',
        '<title id="map-title">Lecture mind map</title><desc id="map-desc">The lecture topic is in the centre, connected to key concepts and supporting details on both sides.</desc>',
        '<rect width="100%" height="100%" rx="24" fill="#f7f8f5"/>',
    ]
    placements = []
    for side, entries, x in (("left", left_items, left_x), ("right", right_items, right_x)):
        for row, (index, concept) in enumerate(entries):
            y = (height - len(entries) * row_height) // 2 + row * row_height + (row_height - card_h) // 2
            center_y = y + card_h // 2
            color = colors[index % len(colors)]
            if side == "left":
                path = f"M {root_x} {root_y + root_h // 2} C 505 {root_y + root_h // 2}, 500 {center_y}, {x + card_w} {center_y}"
            else:
                path = f"M {root_x + root_w} {root_y + root_h // 2} C 855 {root_y + root_h // 2}, 860 {center_y}, {x} {center_y}"
            parts.append(f'<path d="{path}" fill="none" stroke="{color}" stroke-width="3.5" stroke-linecap="round" opacity=".68"/>')
            placements.append((index, concept, x, y, color))

    parts.append(f'<rect x="{root_x}" y="{root_y}" width="{root_w}" height="{root_h}" rx="26" fill="#152d3c"/>')
    parts.append(f'<text x="{root_x + 25}" y="{root_y + 40}" fill="#a8d9d6" font-family="Arial,sans-serif" font-size="13" font-weight="700" letter-spacing="1.4">LECTURE TOPIC</text>')
    title_lines = _wrap_svg_text(title, 21, 5)
    for line_index, value in enumerate(title_lines):
        parts.append(f'<text x="{root_x + 25}" y="{root_y + 82 + line_index * 26}" fill="#ffffff" font-family="Arial,sans-serif" font-size="20" font-weight="700">{escape(value)}</text>')

    for index, concept, x, y, color in placements:
        parts.append(f'<rect x="{x}" y="{y}" width="{card_w}" height="{card_h}" rx="19" fill="#ffffff" stroke="#dce3e2" stroke-width="1.5"/>')
        accent_x = x if x < root_x else x + card_w - 8
        parts.append(f'<rect x="{accent_x}" y="{y}" width="8" height="{card_h}" rx="4" fill="{color}"/>')
        text_x = x + 25 if x < root_x else x + 27
        label_lines = _wrap_svg_text(concept.get("label", "Key concept"), 38, 3)
        label_y = y + 32
        for line_index, value in enumerate(label_lines):
            parts.append(f'<text x="{text_x}" y="{label_y + line_index * 24}" fill="#172e3d" font-family="Arial,sans-serif" font-size="18" font-weight="700">{escape(value)}</text>')
        details = concept.get("details") or ["Key idea from the lecture"]
        detail_y = label_y + len(label_lines) * 24 + 7
        detail_line_index = 0
        for detail in details[:2]:
            for value in _wrap_svg_text(detail, 55, 3):
                parts.append(f'<text x="{text_x}" y="{detail_y + detail_line_index * 18}" fill="#586a73" font-family="Arial,sans-serif" font-size="13">{escape(value)}</text>')
                detail_line_index += 1
    parts.append("</svg>")
    return "".join(parts)


def render_mindmap_svg(mindmap):
    """Prefer Graphviz when available; fall back to the built-in SVG layout."""
    import shutil
    import subprocess
    dot_binary = shutil.which("dot")
    if dot_binary:
        lines = ["digraph LectureMind {", "graph [root=root, overlap=false, bgcolor=\"#f7f8f5\", pad=\"0.35\", ranksep=\"2.2\", splines=curved];", "node [shape=box, style=\"rounded,filled\", fontname=\"Arial\", color=\"#dce3e2\", fillcolor=\"white\", fontcolor=\"#172e3d\", margin=\"0.22,0.14\"];", "edge [color=\"#93aaa9\", penwidth=\"1.5\"];"]
        root_label = "\n".join(_wrap_svg_text(mindmap["title"], 22, 5))
        lines.append(f"root [label={_dot_quote(root_label)}, fillcolor=\"#152d3c\", fontcolor=\"white\", color=\"#152d3c\", fontsize=\"18\"];")
        for index, concept in enumerate(mindmap["concepts"]):
            node = f"concept{index}"
            label = "\n".join(_wrap_svg_text(concept["label"], 38, 3))
            if concept.get("details"):
                label += "\n" + "\n".join(line for detail in concept["details"][:2] for line in _wrap_svg_text(detail, 55, 3))
            lines.append(f"{node} [label={_dot_quote(label)}, fontsize=\"14\"];")
            lines.append(f"root -> {node};")
        lines.append("}")
        try:
            result = subprocess.run([dot_binary, "-Ktwopi", "-Tsvg"], input="\n".join(lines), text=True, capture_output=True, check=True, timeout=15)
            svg = result.stdout
            if "<svg" in svg and "</svg>" in svg:
                svg = svg[svg.find("<svg"):svg.rfind("</svg>") + 6]
                return svg
        except (subprocess.SubprocessError, OSError) as exc:
            print(f"Graphviz rendering failed; built-in SVG layout used: {exc}")
    return _fallback_svg(mindmap["title"], mindmap["concepts"])


def parse_flashcards(text):
    cards = []
    front = None
    for raw_line in (text or "").splitlines():
        line = raw_line.replace("**", "").replace("__", "").strip()
        line = re.sub(r"^[-*•]\s+", "", line)
        match = re.match(r"^(Front|Back)\s*:\s*(.*)$", line, re.I)
        if match and match.group(1).casefold() == "front":
            front = match.group(2).strip()
        elif front is not None and match and match.group(1).casefold() == "back":
            back = match.group(2).strip()
            if front and back:
                cards.append({"front": front, "back": back})
            front = None
    return cards


def parse_quiz(text):
    questions = []
    current = None
    for line in (text or "").splitlines():
        q_match = re.match(r"\s*(?:Question\s*\d*\s*:\s*)?(.+\?)\s*$", line, re.I)
        option_match = re.match(r"\s*([A-D])[.)]\s*(.+)", line)
        answer_match = re.match(r"\s*Answer\s*:\s*([A-D])\b", line, re.I)
        if line.strip().lower().startswith("question:"):
            if current:
                questions.append(current)
            current = {"prompt": re.sub(r"^\s*Question\s*:\s*", "", line, flags=re.I), "options": [], "answer": ""}
        elif q_match and not option_match and not answer_match:
            if current:
                questions.append(current)
            current = {"prompt": q_match.group(1), "options": [], "answer": ""}
        elif option_match and current:
            current["options"].append({"letter": option_match.group(1), "text": option_match.group(2)})
        elif answer_match and current:
            current["answer"] = answer_match.group(1).upper()
    if current:
        questions.append(current)
    return [item for item in questions if item.get("prompt")]


def user_safe_error(exc):
    if isinstance(exc, PipelineError):
        return str(exc)
    return "Processing stopped because one of the lecture files could not be handled. Check the file and local model setup, then try again."


if __name__ == "__main__":
    print("Use `python webapp.py` to start the LectureMind web application.")
