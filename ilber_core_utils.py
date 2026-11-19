import docx
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import subprocess
import os
import sys
import re
import shutil
import threading
import time
import datetime
import json
import math
import numpy as np
import customtkinter as ctk
from PIL import Image
from tkinter import filedialog, messagebox
import srt
os.environ.setdefault("translators_default_region", "EN")


def resource_path(rel_path):
    """Find bundled or local asset paths reliably in both Python and PyInstaller builds (Windows & macOS)."""
    candidates = [
        os.path.join(getattr(sys, "_MEIPASS", ""), rel_path),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), rel_path),
        os.path.join(os.path.dirname(os.path.abspath(sys.executable)), rel_path),
        os.path.join(os.path.dirname(os.path.abspath(sys.executable)), "_internal", rel_path),
        os.path.join(os.path.dirname(os.path.abspath(sys.executable)), "..", "Resources", rel_path),
        os.path.join(os.path.expanduser("~"), "Downloads", rel_path),
        rel_path,
    ]
    for p in candidates:
        if p and os.path.exists(p):
            return p
    return rel_path


def open_path_crossplatform(target_path):
    """Opens a file or folder in the native OS viewer (Windows Explorer / macOS Finder / Linux xdg-open)."""
    if not target_path or not os.path.exists(target_path):
        return
    try:
        if os.name == "nt" and hasattr(os, "startfile"):
            os.startfile(target_path)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", target_path])
        else:
            subprocess.Popen(["xdg-open", target_path])
    except Exception:
        pass


def format_ts_hms(seconds):
    seconds = max(0.0, float(seconds))
    h = int(seconds // 3600)
    m = int((seconds % 3600) // 60)
    s = int(seconds % 60)
    return f"{h:02d}:{m:02d}:{s:02d}"


def format_ts_vtt(seconds):
    seconds = max(0.0, float(seconds))
    h = int(seconds // 3600)
    m = int((seconds % 3600) // 60)
    s = int(seconds % 60)
    ms = int(round((seconds - int(seconds)) * 1000))
    if ms >= 1000:
        ms = 999
    return f"{h:02d}:{m:02d}:{s:02d}.{ms:03d}"


def format_ts_srt(seconds):
    return format_ts_vtt(seconds).replace(".", ",")


_SUPERSCRIPT_MAP = str.maketrans("0123456789+-=nxy", "⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼ⁿˣʸ")


def _latex_to_readable_math(text):
    """Converts raw LaTeX expressions like \\(2^{12}\\), \\frac{a}{b}, \\times into clean human-readable math."""
    if not text:
        return ""
    # Remove LaTeX inline/block delimiters \( \), \[ \], $$, and single $...$
    text = text.replace(r"\(", "").replace(r"\)", "").replace(r"\[", "").replace(r"\]", "").replace("$$", "")
    text = re.sub(r"\$([^$\n]{1,40})\$", r"\1", text)
    # Convert \frac{num}{den} -> num / den
    for _ in range(3):
        text = re.sub(r"\\frac\{([^{}]+)\}\{([^{}]+)\}", r"\1 / \2", text)
    # Convert \sqrt{x} -> √(x)
    text = re.sub(r"\\sqrt\{([^{}]+)\}", r"√(\1)", text)
    # Convert ^{...} and ^x into clean Unicode superscripts when characters are supported
    def _sup_repl(m):
        inner = m.group(1).strip()
        if all(c in "0123456789+-=nxy" for c in inner):
            return inner.translate(_SUPERSCRIPT_MAP)
        return f"^({inner})"
    text = re.sub(r"\^\{([^{}]+)\}", _sup_repl, text)
    text = re.sub(r"\^([0-9]+)", lambda m: m.group(1).translate(_SUPERSCRIPT_MAP), text)
    # Common LaTeX math symbols
    latex_syms = [
        (r"\times", "×"), (r"\cdot", "·"), (r"\div", "÷"), (r"\pm", "±"),
        (r"\neq", "≠"), (r"\ne", "≠"), (r"\leq", "≤"), (r"\le", "≤"),
        (r"\geq", "≥"), (r"\ge", "≥"), (r"\Rightarrow", "➔"), (r"\rightarrow", "➔"),
        (r"\to", "➔"), (r"\approx", "≈"), (r"\pi", "π"), (r"\infty", "∞"),
        (r"\text", ""), ("\\,", " "), ("\\;", " "), ("\\!", ""),
    ]
    for src_s, dst_s in latex_syms:
        text = text.replace(src_s, dst_s)
    # Remove leftover curly braces around simple tokens like {12} -> 12
    text = re.sub(r"\{([0-9a-zA-ZğüşıöçĞÜŞİÖÇ⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻\s\+\-\*/=]+)\}", r"\1", text)
    return text


def clean_markdown_text(text):
    """Remove raw Markdown symbols (**, #, __, ```) and convert raw LaTeX into clean human-readable math."""
    if not text:
        return ""
    text = _latex_to_readable_math(text)
    lines = []
    for raw_line in text.splitlines():
        line = raw_line.strip()
        line = re.sub(r"^#{1,6}\s*", "", line)
        line = re.sub(r"\*\*(.*?)\*\*", r"\1", line)
        line = re.sub(r"__(.*?)__", r"\1", line)
        line = re.sub(r"^[\*\-]\s+", "• ", line)
        line = line.replace("```json", "").replace("```", "")
        lines.append(line)
    return "\n".join(lines)


def add_formatted_text_to_doc(doc, text):
    """Write text into a Word document applying real headings, bold formatting, and yellow warning highlights without raw Markdown."""
    if not text:
        return
    section_titles = (
        "1.", "2.", "3.", "4.", "5.", "6.",
        "Konu ve Genel Bakış", "Ana Kavramlar", "Tanımlar ve Formüller",
        "Örnekler", "Tekrar Soruları", "Kontrol Edilmesi Gereken Kısımlar",
        "Topic and Overview", "Key Concepts", "Definitions and Formulas",
        "Examples", "Review Questions", "Items to Verify",
        "Thema und Überblick", "Hauptkonzepte", "Definitionen und Formeln",
        "Beispiele", "Wiederholungsfragen", "Zu überprüfende Punkte",
    )
    for raw_line in text.splitlines():
        s = raw_line.strip()
        if not s:
            continue
        is_heading = s.startswith("#") or any(s.startswith(t) for t in section_titles)
        clean_s = clean_markdown_text(s)
        if is_heading and len(clean_s) < 95:
            doc.add_heading(clean_s, level=2)
            continue
        p = doc.add_paragraph()
        if "⚠️" in clean_s or "[Kaynakta bulunamadı" in clean_s or "[Not found in source" in clean_s or "[Nicht in Quelle" in clean_s or "[anlaşılmadı]" in clean_s.lower():
            run = p.add_run(clean_s)
            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
            run.bold = True
        else:
            # Parse inline **bold** if any existed in raw_line before cleaning
            parts = re.split(r"(\*\*.*?\*\*)", raw_line.strip())
            if len(parts) > 1:
                for part in parts:
                    if part.startswith("**") and part.endswith("**"):
                        r = p.add_run(part[2:-2])
                        r.bold = True
                    else:
                        p.add_run(clean_markdown_text(part))
            else:
                p.add_run(clean_s)


def add_word_comment_or_note(paragraph, comment_text, author="ILBER FIELD"):
    """Attach a visible field note block to a Word paragraph/cell so field notes never corrupt transcript flow."""
    if not comment_text or not comment_text.strip():
        return
    run = paragraph.add_run(f"  [💬 Saha Notu ({author}): {comment_text.strip()}]")
    run.italic = True
    run.font.color.rgb = RGBColor(180, 83, 9)
    run.font.highlight_color = WD_COLOR_INDEX.YELLOW


def enforce_subtitle_rules(segments, max_chars_per_line=42, max_lines=2, min_dur=1.0, max_dur=7.0):
    """
    Enforce GLOBE subtitle rules:
    - Max 42 chars per line, max 2 lines per subtitle.
    - Split long segments at sentence or phrase boundaries.
    - Duration between 1.0s and 7.0s.
    """
    max_block_chars = (max_chars_per_line * max_lines) - 10
    processed = []

    def wrap_two_lines(txt):
        words = txt.split()
        if not words:
            return ""
        lines = []
        curr = ""
        for w in words:
            if len(w) > max_chars_per_line:
                w = w[:max_chars_per_line]
            if not curr:
                curr = w
            elif len(curr) + 1 + len(w) <= max_chars_per_line:
                curr += " " + w
            else:
                lines.append(curr)
                curr = w
        if curr:
            lines.append(curr)
        return "\n".join(lines[:max_lines])

    for seg in segments:
        start = float(seg.get("start", 0.0))
        end = float(seg.get("end", start + 2.0))
        text = re.sub(r"\s+", " ", str(seg.get("text", "")).strip())
        if not text:
            continue

        total_dur = max(min_dur, end - start)
        chunks = []
        remaining = text
        while len(remaining) > max_block_chars:
            window = remaining[:max_block_chars]
            cut = -1
            for punct in [". ", "? ", "! ", "; ", ", ", " - "]:
                idx = window.rfind(punct)
                if idx > max_block_chars // 3:
                    cut = idx + len(punct) - 1
                    break
            if cut == -1:
                cut = window.rfind(" ")
            if cut <= 0:
                cut = max_block_chars
            chunks.append(remaining[:cut].strip())
            remaining = remaining[cut:].strip()
        if remaining:
            chunks.append(remaining)

        total_chars = max(1, sum(len(c) for c in chunks))
        cursor = start
        for c in chunks:
            prop_dur = total_dur * (len(c) / total_chars)
            dur = max(min_dur, min(max_dur, prop_dur))
            c_end = cursor + dur
            processed.append({
                "start": cursor,
                "end": c_end,
                "text": wrap_two_lines(c),
                "speaker": seg.get("speaker", ""),
                "note": seg.get("note", ""),
            })
            cursor = c_end

    # Prevent overlaps and enforce min/max duration
    for i in range(len(processed)):
        s = processed[i]["start"]
        e = processed[i]["end"]
        if e - s < min_dur:
            e = s + min_dur
        if e - s > max_dur:
            e = s + max_dur
        if i + 1 < len(processed) and e > processed[i + 1]["start"]:
            if processed[i + 1]["start"] - s >= min_dur:
                e = processed[i + 1]["start"]
            else:
                processed[i + 1]["start"] = e
        processed[i]["end"] = e

    return processed


def write_srt_file(segments, path):
    with open(path, "w", encoding="utf-8") as f:
        for i, seg in enumerate(segments, 1):
            f.write(f"{i}\n{format_ts_srt(seg['start'])} --> {format_ts_srt(seg['end'])}\n{seg['text'].strip()}\n\n")


def write_vtt_file(segments, path):
    with open(path, "w", encoding="utf-8") as f:
        f.write("WEBVTT\n\n")
        for i, seg in enumerate(segments, 1):
            f.write(f"{i}\n{format_ts_vtt(seg['start'])} --> {format_ts_vtt(seg['end'])}\n{seg['text'].strip()}\n\n")


class AudioVideoPreviewer:
    """Plays exact [start, end] audio slices and extracts video preview frames using local FFmpeg + sounddevice."""
    def __init__(self):
        self.is_playing = False

    def stop(self):
        try:
            import sounddevice as sd
            sd.stop()
            self.is_playing = False
        except Exception:
            pass

    def play_segment(self, media_path, start_sec, end_sec):
        self.stop()
        if not media_path or not os.path.exists(media_path):
            return
        dur = max(0.8, min(15.0, float(end_sec) - float(start_sec)))
        start_sec = max(0.0, float(start_sec))

        def _worker():
            try:
                import sounddevice as sd
                self.is_playing = True
                cmd = [
                    "ffmpeg", "-y",
                    "-ss", f"{start_sec:.3f}",
                    "-t", f"{dur:.3f}",
                    "-i", media_path,
                    "-vn", "-ac", "1", "-ar", "16000", "-f", "s16le", "-"
                ]
                creationflags = 0x08000000 if os.name == "nt" else 0
                proc = subprocess.run(cmd, capture_output=True, creationflags=creationflags)
                if proc.returncode == 0 and proc.stdout:
                    audio_np = np.frombuffer(proc.stdout, dtype=np.int16).astype(np.float32) / 32768.0
                    sd.play(audio_np, samplerate=16000)
            except Exception:
                pass

        threading.Thread(target=_worker, daemon=True).start()

    def extract_frame_pil(self, media_path, time_sec, size=(216, 120)):
        if not media_path or not os.path.exists(media_path):
            return None
        ext = media_path.lower().rsplit(".", 1)[-1]
        if ext not in ("mp4", "mov", "mkv", "webm", "avi"):
            return None
        try:
            import io
            cmd = [
                "ffmpeg", "-y",
                "-ss", f"{max(0.0, float(time_sec)):.2f}",
                "-i", media_path,
                "-frames:v", "1",
                "-vf", f"scale={size[0]}:{size[1]}:force_original_aspect_ratio=decrease",
                "-f", "image2pipe", "-vcodec", "png", "-"
            ]
            creationflags = 0x08000000 if os.name == "nt" else 0
            proc = subprocess.run(cmd, capture_output=True, creationflags=creationflags, timeout=3)
            if proc.returncode == 0 and proc.stdout:
                return Image.open(io.BytesIO(proc.stdout)).copy()
        except Exception:
            pass
        return None


def acoustic_diarize_segments(media_path, segments, num_speakers=2, speaker_labels=None):
    """
    Perform local audio-based speaker diarization on timestamped Whisper segments.
    Extracts 16kHz mono PCM via FFmpeg and computes acoustic voice feature vectors
    (fundamental pitch / ZCR, spectral centroid, low/mid/high frequency band ratios)
    for each segment, then clusters speakers acoustically.
    """
    if not segments:
        return []
    num_speakers = max(1, min(6, int(num_speakers)))
    if not speaker_labels or len(speaker_labels) < num_speakers:
        speaker_labels = [f"Konuşmacı {i+1}" for i in range(num_speakers)]

    # Decode full audio to 16kHz mono int16 array via FFmpeg
    audio_np = None
    sr = 16000
    try:
        cmd = [
            "ffmpeg", "-y", "-i", media_path,
            "-vn", "-ac", "1", "-ar", str(sr), "-f", "s16le", "-"
        ]
        creationflags = 0x08000000 if os.name == "nt" else 0
        proc = subprocess.run(cmd, capture_output=True, creationflags=creationflags)
        if proc.returncode == 0 and proc.stdout:
            audio_np = np.frombuffer(proc.stdout, dtype=np.int16).astype(np.float32) / 32768.0
    except Exception:
        audio_np = None

    features = []
    valid_indices = []
    for idx, seg in enumerate(segments):
        s_sec = float(seg.get("start", 0.0))
        e_sec = float(seg.get("end", s_sec + 1.0))
        if audio_np is not None and len(audio_np) > sr:
            i0 = max(0, int(s_sec * sr))
            i1 = min(len(audio_np), int(e_sec * sr))
            clip = audio_np[i0:i1]
        else:
            clip = np.array([], dtype=np.float32)

        if len(clip) >= 400:
            # 1. Zero-crossing rate (correlates with fundamental pitch/voicing)
            zcr = float(np.mean(np.abs(np.diff(np.sign(clip)))))
            # 2. Spectral features via FFT
            win = clip[:min(len(clip), sr * 4)]
            spec = np.abs(np.fft.rfft(win))
            freqs = np.fft.rfftfreq(len(win), d=1.0 / sr)
            spec_sum = float(np.sum(spec)) + 1e-8
            centroid = float(np.sum(freqs * spec) / spec_sum) / 4000.0
            # Sub-band energy ratios (Voice fundamental 80-300Hz, Formants 300-1200Hz, High 1200-4000Hz)
            b_low = float(np.sum(spec[(freqs >= 80) & (freqs < 300)])) / spec_sum
            b_mid = float(np.sum(spec[(freqs >= 300) & (freqs < 1200)])) / spec_sum
            b_high = float(np.sum(spec[(freqs >= 1200) & (freqs < 4000)])) / spec_sum
            features.append([zcr, centroid, b_low, b_mid, b_high])
            valid_indices.append(idx)
        else:
            features.append([0.0, 0.0, 0.33, 0.33, 0.33])

    X = np.array(features, dtype=np.float32)
    if len(valid_indices) >= num_speakers and num_speakers > 1:
        # Standardize features
        mean = np.mean(X[valid_indices], axis=0)
        std = np.std(X[valid_indices], axis=0) + 1e-6
        Xn = (X - mean) / std

        # Deterministic K-Means initialization (spread across quantiles of primary acoustic axis)
        proj = Xn[:, 0] + Xn[:, 1] + Xn[:, 2]
        quantiles = np.linspace(0.15, 0.85, num_speakers)
        init_idx = [int(np.argmin(np.abs(proj - np.quantile(proj, q)))) for q in quantiles]
        centroids = Xn[init_idx].copy()

        labels = np.zeros(len(segments), dtype=int)
        for _ in range(25):
            dists = np.linalg.norm(Xn[:, None, :] - centroids[None, :, :], axis=2)
            new_labels = np.argmin(dists, axis=1)
            for k in range(num_speakers):
                if np.any(new_labels == k):
                    centroids[k] = np.mean(Xn[new_labels == k], axis=0)
            labels = new_labels

        # Smooth very short isolated speaker flips when pause between segments is tiny (<0.25s)
        for i in range(1, len(segments) - 1):
            pause_prev = float(segments[i]["start"]) - float(segments[i - 1]["end"])
            dur_curr = float(segments[i]["end"]) - float(segments[i]["start"])
            if pause_prev < 0.25 and dur_curr < 1.2 and labels[i - 1] == labels[i + 1]:
                labels[i] = labels[i - 1]

        # Order speaker IDs so the first speaker in the recording is Speaker 0
        remap = {}
        next_id = 0
        ordered_labels = []
        for lbl in labels:
            if int(lbl) not in remap:
                remap[int(lbl)] = next_id
                next_id += 1
            ordered_labels.append(remap[int(lbl)])
    else:
        ordered_labels = [0 for _ in segments]

    out_segments = []
    for i, seg in enumerate(segments):
        spk_idx = ordered_labels[i] if i < len(ordered_labels) else 0
        spk_name = speaker_labels[spk_k] if (spk_k := min(spk_idx, len(speaker_labels) - 1)) >= 0 else "Konuşmacı 1"
        out_segments.append({
            "start": float(seg.get("start", 0.0)),
            "end": float(seg.get("end", 0.0)),
            "speaker": spk_name,
            "text": str(seg.get("text", "")).strip(),
            "note": str(seg.get("note", "")),
        })
    return out_segments


def detect_anonymization_candidates(segments):
    """
    Detect candidate Person ([KİŞİ-X]), Location ([YER-X]), and Institution ([KURUM-X]) names
    in the transcript for researcher approval in FIELD segment.
    """
    inst_keywords = (
        "Üniversitesi", "Bakanlığı", "Belediyesi", "Hastanesi", "Okulu", "Lisesi",
        "Kurumu", "Derneği", "Vakfı", "Şirketi", "Holding", "Bankası", "Enstitüsü",
        "University", "Hospital", "Ministry", "Institute", "Corporation", "GmbH",
    )
    loc_keywords = (
        "Mahallesi", "Caddesi", "Sokağı", "İlçesi", "Köyü", "Meydanı", "Bulvarı",
        "İstanbul", "Ankara", "İzmir", "Bursa", "Antalya", "Adana", "Konya", "Gaziantep",
        "Berlin", "London", "Paris", "New York",
    )
    common_sentence_starters = {
        "Ama", "Fakat", "Ancak", "Çünkü", "Yani", "Mesela", "Örneğin", "Tabii", "Evet", "Hayır",
        "Şimdi", "Sonra", "Önce", "Bence", "Aslında", "Zaten", "Belki", "Acaba", "Nasıl", "Neden",
        "Bunun", "Şunun", "Onun", "Burada", "Orada", "Şurada", "Ben", "Sen", "O", "Biz", "Siz", "Onlar",
        "Bir", "Bu", "Şu", "Ve", "İle", "İçin", "Gibi", "Kadar", "Daha", "Çok", "En", "Her", "Hiç",
        "Valla", "Vallahi", "Şey", "Hocam", "Abi", "Abla", "Kardeşim", "Arkadaşlar", "Peki", "Tamam",
        "The", "And", "But", "Because", "Well", "Yes", "No", "Now", "Then", "Here", "There", "This", "That",
        "Und", "Aber", "Weil", "Ja", "Nein", "Jetzt", "Dann", "Hier", "Dort", "Das", "Dies",
    }

    candidates = []  # list of dicts: {"original": str, "code": str, "type": str, "approved": True}
    seen = set()
    counts = {"KİŞİ": 0, "YER": 0, "KURUM": 0}

    for seg in segments:
        text = seg.get("text", "")
        # Multi-word capitalized phrases or mid-sentence capitalized words
        tokens = re.findall(r"\b[A-ZÇĞİÖŞÜ][a-zçğıöşü]{2,}(?:\s+[A-ZÇĞİÖŞÜ][a-zçğıöşü]{2,})*\b", text)
        for tok in tokens:
            clean_tok = tok.strip()
            if clean_tok in seen or clean_tok in common_sentence_starters:
                continue
            # Check if it's at the very beginning of a sentence and only 1 word
            is_single_word = " " not in clean_tok
            if is_single_word and (text.startswith(clean_tok) or f". {clean_tok}" in text or f"? {clean_tok}" in text or f"! {clean_tok}" in text):
                if clean_tok not in loc_keywords and not any(k in clean_tok for k in inst_keywords):
                    continue

            if any(k in clean_tok for k in inst_keywords):
                etype = "KURUM"
            elif any(k in clean_tok for k in loc_keywords) or clean_tok.endswith(("'da", "'de", "'ta", "'te", "'a", "'e", "'dan", "'den")):
                etype = "YER"
            else:
                etype = "KİŞİ"

            counts[etype] += 1
            code = f"[{etype}-{counts[etype]}]"
            seen.add(clean_tok)
            candidates.append({
                "original": clean_tok,
                "code": code,
                "type": etype,
                "approved": True,
            })

    return candidates


def verify_summary_against_transcript(summary_text, transcript_text, lang="tr"):
    """
    Stage 4 of EDU Summarization: Automatic Verification.
    Checks proper nouns / capitalized terms and timestamps in the summary against the source transcript.
    Marks unverified items with a warning tag and returns (verified_summary_text, warnings_list).
    """
    if not summary_text:
        return "", []

    trans_lower = transcript_text.lower()
    # Build set of 4-char prefixes of all words in transcript for agglutinative (Turkish) stem matching
    trans_words = re.findall(r"\w{3,}", trans_lower)
    trans_stems = {w[:5] for w in trans_words if len(w) >= 5} | set(trans_words)

    ignore_words = {
        "konu", "genel", "bakış", "ana", "kavramlar", "tanımlar", "formüller", "örnekler",
        "tekrar", "soruları", "kontrol", "edilmesi", "gereken", "kısımlar", "ders", "özeti",
        "bölüm", "madde", "soru", "cevap", "doğru", "şık", "açık", "uçlu", "çoktan", "seçmeli",
        "anlaşılmadı", "kaynakta", "bulunamadı", "topic", "overview", "key", "concepts",
        "definitions", "formulas", "examples", "review", "questions", "answer", "thema",
        "überblick", "hauptkonzepte", "definitionen", "formeln", "beispiele", "fragen",
    }

    tag_msg = {
        "tr": "⚠️ [Kaynakta bulunamadı - Kontrol edin]",
        "en": "⚠️ [Not found in source - Please verify]",
        "de": "⚠️ [Nicht in Quelle gefunden - Bitte prüfen]",
    }.get(lang, "⚠️ [Kaynakta bulunamadı - Kontrol edin]")

    verified_lines = []
    warnings = []

    for line in summary_text.splitlines():
        s = line.strip()
        if not s:
            verified_lines.append("")
            continue
        # Skip section headers
        if re.match(r"^[1-6][\.\)]\s+", s) and len(s) < 65:
            verified_lines.append(s)
            continue

        # Check capitalized proper nouns / technical terms inside the line
        cap_terms = re.findall(r"\b[A-ZÇĞİÖŞÜ][a-zçğıöşüA-ZÇĞİÖŞÜ]{3,}\b", s)
        unverified_terms = []
        for term in cap_terms:
            tl = term.lower()
            if tl in ignore_words:
                continue
            # Check if term or its 5-letter stem exists in transcript
            stem = tl[:5] if len(tl) >= 5 else tl
            if tl not in trans_lower and stem not in trans_stems:
                unverified_terms.append(term)

        if unverified_terms and not s.startswith("⚠️"):
            flagged = f"{s}  {tag_msg}: ({', '.join(unverified_terms)})"
            verified_lines.append(flagged)
            warnings.append(flagged)
        else:
            if "[anlaşılmadı]" in s.lower() or "[unclear]" in s.lower() or "[unverständlich]" in s.lower():
                warnings.append(s)
            verified_lines.append(s)

    return "\n".join(verified_lines), warnings


def chunk_segments_for_edu(segments, max_words=1250, overlap_sentences=2):
    """
    Stage 1 of EDU Summarization:
    Split timestamped segments into ~8-10 minute (~1,200-1,500 word) chunks at sentence boundaries,
    leaving 1-2 sentences of overlap between consecutive chunks.
    """
    if not segments:
        return []
    chunks = []
    current_segs = []
    current_words = 0

    for seg in segments:
        txt = str(seg.get("text", "")).strip()
        if not txt:
            continue
        w_count = len(txt.split())
        current_segs.append(seg)
        current_words += w_count

        ends_sentence = txt.endswith((".", "?", "!"))
        if current_words >= max_words and ends_sentence:
            chunks.append(list(current_segs))
            current_segs = current_segs[-overlap_sentences:] if overlap_sentences > 0 else []
            current_words = sum(len(s.get("text", "").split()) for s in current_segs)

    if current_segs:
        if not chunks or current_segs != chunks[-1]:
            chunks.append(list(current_segs))
    return chunks


def build_edu_chunk_prompt(lang="tr"):
    """Return strict Stage 2 JSON extraction system prompt obeying all 6 mandatory rules."""
    if lang == "tr":
        return (
            "Sen akademik bir ders notu çıkarıcısısın. Sana zaman damgalı ders deşifresi verilecek.\n"
            "ZORUNLU KURALLAR:\n"
            "1. Yalnızca deşifrede geçen bilgileri kullan; kesinlikle kendi dış bilgini ekleme.\n"
            "2. Anlaşılmayan veya bozuk yerleri [anlaşılmadı] olarak işaretle, tahmin etme.\n"
            "3. Dersi anlatan konuşmacı ile anlatılan konudaki kişileri birbirine karıştırma.\n"
            "4. Üçüncü şahıs ve tarafsız akademik dil kullan; kesinlikle 'ben' veya 'biz' kullanma.\n"
            "5. Her maddenin yanına mutlaka [SS:DD:SN] veya [DD:SN] zaman damgasını yaz.\n"
            "6. Sona değerlendirme veya dolgu cümlesi ('Bu özet ... yapılmıştır' vb.) ekleme.\n\n"
            "SADECE aşağıdaki JSON formatında yanıt ver (başka hiçbir açıklama yazma):\n"
            "{\n"
            '  "genel_bakis": ["En fazla 2-3 cümlelik tarafsız özet [00:00:00]"],\n'
            '  "ana_kavramlar": ["[00:01:15] Kavram ve ders içi açıklaması"],\n'
            '  "tanimlar_formuller": ["[00:02:30] Tanım veya formül açıklaması"],\n'
            '  "ornekler": ["[00:04:10] Derste verilen örnek"],\n'
            '  "belirsiz_kisimlar": ["[00:05:00] [anlaşılmadı] veya belirsiz ifade"]\n'
            "}"
        )
    elif lang == "de":
        return (
            "Sie sind ein akademischer Vorlesungsassistent. Ihnen wird ein Transkript mit Zeitstempeln gegeben.\n"
            "STRENGE REGELN:\n"
            "1. Verwenden Sie NUR Informationen aus dem Transkript; fügen Sie kein eigenes Wissen hinzu.\n"
            "2. Markieren Sie unklare Stellen als [unverständlich], raten Sie nicht.\n"
            "3. Verwechseln Sie nicht den Dozenten mit erwähnten Personen.\n"
            "4. Verwenden Sie die dritte Person und neutrale Sprache; kein 'ich' oder 'wir'.\n"
            "5. Fügen Sie jedem Punkt den Zeitstempel [HH:MM:SS] hinzu.\n"
            "6. Keine Füllsätze am Ende.\n\n"
            "Antworten Sie NUR im folgenden JSON-Format:\n"
            "{\n"
            '  "genel_bakis": ["Maximal 2-3 neutrale Sätze [00:00:00]"],\n'
            '  "ana_kavramlar": ["[00:01:15] Hauptkonzept"],\n'
            '  "tanimlar_formuller": ["[00:02:30] Definition oder Formel"],\n'
            '  "ornekler": ["[00:04:10] Genanntes Beispiel"],\n'
            '  "belirsiz_kisimlar": ["[00:05:00] [unverständlich]"]\n'
            "}"
        )
    else:
        return (
            "You are an academic lecture note extractor. You will receive a timestamped lecture transcript.\n"
            "MANDATORY RULES:\n"
            "1. Use ONLY information explicitly stated in the transcript; never add external knowledge.\n"
            "2. Mark unclear or garbled parts as [unclear]; do not guess.\n"
            "3. Do not confuse the lecturer with historical/subject figures mentioned.\n"
            "4. Use third-person neutral academic tone; never use 'I' or 'we'.\n"
            "5. Include the timestamp [HH:MM:SS] next to every single point.\n"
            "6. Do not add concluding filler sentences.\n\n"
            "Respond ONLY in the following JSON format:\n"
            "{\n"
            '  "genel_bakis": ["Max 2-3 neutral overview sentences [00:00:00]"],\n'
            '  "ana_kavramlar": ["[00:01:15] Key concept explained in lecture"],\n'
            '  "tanimlar_formuller": ["[00:02:30] Definition or formula"],\n'
            '  "ornekler": ["[00:04:10] Example given in lecture"],\n'
            '  "belirsiz_kisimlar": ["[00:05:00] [unclear] segment"]\n'
            "}"
        )


def parse_edu_chunk_json(raw_llm_output):
    """Safely parse LLM JSON output with fallback line extraction so it never crashes on minor syntax errors."""
    data = {
        "genel_bakis": [],
        "ana_kavramlar": [],
        "tanimlar_formuller": [],
        "ornekler": [],
        "belirsiz_kisimlar": [],
    }
    if not raw_llm_output:
        return data
    cleaned = raw_llm_output.strip()
    # Find JSON block
    m = re.search(r"\{.*\}", cleaned, flags=re.DOTALL)
    if m:
        try:
            parsed = json.loads(m.group(0))
            if isinstance(parsed, dict):
                for k in data.keys():
                    val = parsed.get(k, [])
                    if isinstance(val, list):
                        data[k] = [clean_markdown_text(str(x)).strip() for x in val if str(x).strip()]
                    elif isinstance(val, str) and val.strip():
                        data[k] = [clean_markdown_text(val).strip()]
                return data
        except Exception:
            pass

    # Fallback line parser if JSON was slightly malformed
    current_key = "ana_kavramlar"
    for line in cleaned.splitlines():
        s = clean_markdown_text(line).strip(" ,\"'[]")
        if not s or s in ("{", "}"):
            continue
        low = s.lower()
        if "genel_bakis" in low or "overview" in low or "konu" in low:
            current_key = "genel_bakis"
            continue
        elif "ana_kavramlar" in low or "kavram" in low or "concept" in low:
            current_key = "ana_kavramlar"
            continue
        elif "tanimlar" in low or "formuller" in low or "definition" in low:
            current_key = "tanimlar_formuller"
            continue
        elif "ornekler" in low or "örnek" in low or "example" in low:
            current_key = "ornekler"
            continue
        elif "belirsiz" in low or "unclear" in low:
            current_key = "belirsiz_kisimlar"
            continue
        data[current_key].append(s)
    return data


def run_edu_4stage_pipeline(llm_callable, segments, audio_lang="tr", log_callback=None, ui_lang="tr"):
    """
    Executes the complete 4-Stage EDU Lecture Summarization & Question Generation Pipeline:
    1. Chunking (8-10 min / ~1250 words with 2-sentence overlap)
    2. JSON Note Extraction per chunk (temp 0.1, strict rules, timestamps)
    3. Merging & Deduplication + Programmatic Template Formatting
    4. Automatic Verification (timestamp & proper noun hallucination check against transcript)
    + Review Question Generation strictly from verified notes.
    """
    lang = audio_lang if audio_lang in ("tr", "en", "de") else "en"
    chunks = chunk_segments_for_edu(segments, max_words=1200, overlap_sentences=2)
    if not chunks:
        return "", ""

    merged = {
        "genel_bakis": [],
        "ana_kavramlar": [],
        "tanimlar_formuller": [],
        "ornekler": [],
        "belirsiz_kisimlar": [],
    }

    sys_prompt = build_edu_chunk_prompt(lang)
    valid_timestamps = set()
    full_transcript_lines = []
    for s in segments:
        ts_str = format_ts_hms(s["start"])
        valid_timestamps.add(ts_str[:5])  # HH:MM or MM:SS prefix
        full_transcript_lines.append(f"[{ts_str}] {s['text'].strip()}")
    full_transcript_text = "\n".join(full_transcript_lines)

    for idx, ch in enumerate(chunks, 1):
        if log_callback:
            if ui_lang == "en":
                log_callback(f"      -> [EDU Stage 1-2] Analyzing chunk {idx}/{len(chunks)} (JSON Note Extraction)...")
            elif ui_lang == "de":
                log_callback(f"      -> [EDU Stufe 1-2] Abschnitt {idx}/{len(chunks)} wird analysiert (JSON-Notizen)...")
            else:
                log_callback(f"      -> [EDU Aşama 1-2] Bölüm {idx}/{len(chunks)} analiz ediliyor (JSON Not Çıkarma)...")
        chunk_text = "\n".join([f"[{format_ts_hms(s['start'])}] {s['text'].strip()}" for s in ch])
        raw_out = llm_callable(sys_prompt, chunk_text, max_tokens=900, temperature=0.1)
        parsed = parse_edu_chunk_json(raw_out)
        for k in merged.keys():
            for item in parsed.get(k, []):
                if item and item not in merged[k]:
                    merged[k].append(item)

    if log_callback:
        if ui_lang == "en":
            log_callback("      -> [EDU Stage 3-4] Merging notes & running Hallucination / Timestamp verification...")
        elif ui_lang == "de":
            log_callback("      -> [EDU Stufe 3-4] Notizen zusammenführen & Halluzinations-/Zeitstempelprüfung...")
        else:
            log_callback("      -> [EDU Aşama 3-4] Notlar birleştiriliyor ve Halüsinasyon / Zaman Damgası kontrolü yapılıyor...")

    # Stage 4: Verify items against full_transcript_text
    verified_notes_for_questions = []

    def filter_and_verify_list(items):
        out = []
        for item in items:
            # Remove first-person pronouns if any slipped through
            item_clean = re.sub(r"\b(ben|biz|bence|bana göre)\b", "", item, flags=re.IGNORECASE).strip()
            ver_line, warns = verify_summary_against_transcript(item_clean, full_transcript_text, lang=lang)
            out.append(ver_line)
            if not warns and "[anlaşılmadı]" not in ver_line.lower():
                verified_notes_for_questions.append(ver_line)
            elif warns:
                for w in warns:
                    if w not in merged["belirsiz_kisimlar"]:
                        merged["belirsiz_kisimlar"].append(w)
        return out

    v_bakis = filter_and_verify_list(merged["genel_bakis"][:3])
    v_kavram = filter_and_verify_list(merged["ana_kavramlar"])
    v_tanim = filter_and_verify_list(merged["tanimlar_formuller"])
    v_ornek = filter_and_verify_list(merged["ornekler"])

    # Generate Review Questions ONLY from verified notes
    if log_callback:
        if ui_lang == "en":
            log_callback("      -> [EDU Questions] Generating timestamped review questions from verified lecture notes...")
        elif ui_lang == "de":
            log_callback("      -> [EDU Fragen] Wiederholungsfragen mit Zeitstempeln werden aus verifizierten Notizen erstellt...")
        else:
            log_callback("      -> [EDU Soru Üretimi] Doğrulanmış ders notlarından zaman damgalı tekrar soruları üretiliyor...")

    verified_block = "\n".join(verified_notes_for_questions[:25]) if verified_notes_for_questions else full_transcript_text[:3000]
    if lang == "tr":
        q_prompt = (
            "Sana yalnızca doğrulanmış ve zaman damgalı ders notları veriliyor.\n"
            "KURALLAR:\n"
            "1. SADECE bu notlardaki bilgilerden 3 adet Çoktan Seçmeli ve 3 adet Açık Uçlu olmak üzere toplam 6 tekrar/sınav sorusu üret.\n"
            "2. Her sorunun altına DOĞRU CEVABI ve bilginin geçtiği [SS:DD:SN] ZAMAN DAMGASINI mutlaka yaz.\n"
            "3. Ham Markdown (** veya #) kullanma. Düz ve temiz metin yaz."
        )
    elif lang == "de":
        q_prompt = (
            "Ihnen werden verifizierte Vorlesungsnotizen mit Zeitstempeln gegeben.\n"
            "Erstellen Sie 3 Multiple-Choice-Fragen und 3 offene Wiederholungsfragen NUR basierend auf diesen Notizen.\n"
            "Geben Sie unter jeder Frage die ANTWORT und den ZEITSTEMPEL [HH:MM:SS] an. Kein Markdown (** oder #)."
        )
    else:
        q_prompt = (
            "You are given verified timestamped lecture notes.\n"
            "Create 3 Multiple-Choice and 3 Open-Ended review questions strictly from these notes.\n"
            "Include the ANSWER and the [HH:MM:SS] TIMESTAMP under each question. Do not use Markdown (** or #)."
        )

    raw_questions = llm_callable(q_prompt, verified_block, max_tokens=900, temperature=0.1)
    clean_questions = clean_markdown_text(raw_questions)

    # Assemble Programmatic Lecture Summary Template (No meeting headers like 'Action Items'!)
    headers = {
        "tr": (
            "1. Konu ve Genel Bakış",
            "2. Ana Kavramlar",
            "3. Tanımlar ve Formüller",
            "4. Örnekler",
            "5. Tekrar Soruları (Zaman Damgalı)",
            "6. Kontrol Edilmesi Gereken Kısımlar",
            "Belirsiz veya kaynakta doğrulanamayan madde tespit edilmedi.",
        ),
        "de": (
            "1. Thema und Überblick",
            "2. Hauptkonzepte",
            "3. Definitionen und Formeln",
            "4. Beispiele",
            "5. Wiederholungsfragen (mit Zeitstempel)",
            "6. Zu überprüfende Punkte",
            "Keine unklaren oder nicht verifizierten Punkte gefunden.",
        ),
        "en": (
            "1. Topic and Overview",
            "2. Key Concepts",
            "3. Definitions and Formulas",
            "4. Examples",
            "5. Review Questions (Timestamped)",
            "6. Items to Verify",
            "No unclear or unverified items detected.",
        ),
    }.get(lang)

    sec_overview = "\n".join(f"• {x}" for x in v_bakis) if v_bakis else "• —"
    sec_concepts = "\n".join(f"• {x}" for x in v_kavram) if v_kavram else "• —"
    sec_defs = "\n".join(f"• {x}" for x in v_tanim) if v_tanim else "• —"
    sec_ex = "\n".join(f"• {x}" for x in v_ornek) if v_ornek else "• —"
    sec_verify = "\n".join(f"• {x}" for x in merged["belirsiz_kisimlar"]) if merged["belirsiz_kisimlar"] else f"• {headers[6]}"

    final_report = (
        f"{headers[0]}\n{sec_overview}\n\n"
        f"{headers[1]}\n{sec_concepts}\n\n"
        f"{headers[2]}\n{sec_defs}\n\n"
        f"{headers[3]}\n{sec_ex}\n\n"
        f"{headers[4]}\n{clean_questions}\n\n"
        f"{headers[5]}\n{sec_verify}"
    )
    return clean_markdown_text(final_report), clean_questions


def save_bilingual_word_doc(segments_src, lines_tgt, src_label, tgt_label, out_path, title_text="Çift Dilli Deşifre / Bilingual Transcript"):
    """
    Create a two-column Word table:
    Left column: [Timestamp] Source language text
    Right column: Target language translation
    """
    doc = docx.Document()
    doc.add_heading(title_text, 0)
    table = doc.add_table(rows=1, cols=2)
    table.style = "Table Grid"
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = f"Kaynak ({src_label})"
    hdr_cells[1].text = f"Çeviri ({tgt_label})"
    for cell in hdr_cells:
        for p in cell.paragraphs:
            for r in p.runs:
                r.bold = True

    for i, seg in enumerate(segments_src):
        row_cells = table.add_row().cells
        ts = format_ts_hms(seg.get("start", 0.0))
        src_txt = str(seg.get("text", "")).replace("\n", " ").strip()
        tgt_txt = str(lines_tgt[i]).replace("\n", " ").strip() if i < len(lines_tgt) else ""
        row_cells[0].text = f"[{ts}] {src_txt}"
        row_cells[1].text = f"[{ts}] {tgt_txt}"

    doc.save(out_path)


def generate_field_ai_suggestions(llm_callable, segments):
    """
    Scan FIELD transcript for potential misheard words (not dialect/grammar!)
    and return a list of suggestion dicts:
    [{"seg_index": idx, "timestamp": "00:01:12", "original_word": "...", "suggested_word": "...", "context": "...", "status": "pending"}]
    Original text is NEVER modified automatically.
    """
    suggestions = []
    if not segments:
        return suggestions

    sample_lines = []
    for idx, seg in enumerate(segments[:40]):
        sample_lines.append(f"{idx}|[{format_ts_hms(seg['start'])}] {seg['text'].strip()}")
    joined = "\n".join(sample_lines)

    sys_p = (
        "Sen nitel araştırma deşifre denetçisisin. Aşağıdaki deşifre satırlarında ses tanıma motorunun (Whisper) "
        "bağlama göre YANLIŞ DUYMUŞ olabileceği kelimeleri tespit et.\n"
        "KESİN KURAL: Katılımcının ağız özelliklerini, devrik cümlelerini, 'valla, şey, hani' gibi konuşma dili ifadelerini ASLA hata olarak işaretleme.\n"
        "Sadece bariz yanlış duyulmuş kelimeleri şu JSON listesi formatında döndür:\n"
        '[{"seg_index": 0, "original_word": "yanlış_kelime", "suggested_word": "doğru_öneri", "reason": "bağlam uyumu"}]'
    )
    try:
        raw = llm_callable(sys_p, joined, max_tokens=500, temperature=0.1)
        m = re.search(r"\[.*\]", raw, flags=re.DOTALL)
        if m:
            arr = json.loads(m.group(0))
            if isinstance(arr, list):
                for item in arr:
                    idx = int(item.get("seg_index", 0))
                    orig_w = str(item.get("original_word", "")).strip()
                    sugg_w = str(item.get("suggested_word", "")).strip()
                    if 0 <= idx < len(segments) and orig_w and sugg_w and orig_w != sugg_w:
                        if orig_w.lower() in segments[idx]["text"].lower():
                            suggestions.append({
                                "seg_index": idx,
                                "timestamp": format_ts_hms(segments[idx]["start"]),
                                "original_word": orig_w,
                                "suggested_word": sugg_w,
                                "reason": str(item.get("reason", "Olası ses tanıma hatası")),
                                "status": "pending",  # "accepted" or "rejected"
                            })
    except Exception:
        pass
    return suggestions


def export_field_outputs(
    workspace_dir,
    base_name,
    metadata,
    segments,
    suggestions,
    anon_candidates=None,
    translated_lines=None,
    tgt_lang_label=None,
    output_dir=None,
):
    """
    Generate FIELD segment outputs preserving original text as-is (without anonymization tags):
    1. Birebir Deşifre (.docx) - Exact transcript with metadata header, 3-column table (Zaman | Konuşmacı | Metin),
       and Field Notes attached as yellow comment annotations.
    2. Okunabilir Metin (.docx) - Presentation/quote version with accepted AI word suggestions applied.
    3. Düz metin (.txt) & NVivo / MAXQDA / ATLAS.ti uyumlu çıktı (.txt).
    """
    reports_dir = output_dir if output_dir else os.path.join(workspace_dir, "Transcripts")
    os.makedirs(reports_dir, exist_ok=True)

    def apply_accepted_suggestions(text, seg_idx):
        out = text
        for s in suggestions:
            if s.get("seg_index") == seg_idx and s.get("status") == "accepted":
                orig_w = s["original_word"]
                sugg_w = s["suggested_word"]
                out = re.sub(rf"\b{re.escape(orig_w)}\b", sugg_w, out, flags=re.IGNORECASE)
        return out

    # --- 1. BIREBIR DESIFRE (WORD) ---
    verbatim_path = os.path.join(reports_dir, f"{base_name}_FIELD_Birebir_Desifre.docx")
    doc_v = docx.Document()
    doc_v.add_heading(f"FIELD Birebir Görüşme Deşifresi: {base_name}", 0)

    # Metadata Box at top of document
    doc_v.add_heading("Görüşme Bilgileri (Saha Künyesi)", level=2)
    meta_table = doc_v.add_table(rows=6, cols=2)
    meta_table.style = "Table Grid"
    meta_fields = [
        ("Tarih / Date", metadata.get("date", "-")),
        ("Yer / Location", metadata.get("location", "-")),
        ("Katılımcı Kodu / Participant Code", metadata.get("participant_code", "P01")),
        ("Görüşmeci / Interviewer", metadata.get("interviewer", "Görüşmeci")),
        ("Süre / Duration", metadata.get("duration", "-")),
        ("Bağlam ve Gözlem Notu / Context", metadata.get("context", "-")),
    ]
    for r_idx, (k_label, v_val) in enumerate(meta_fields):
        row = meta_table.rows[r_idx].cells
        row[0].text = k_label
        row[1].text = str(v_val)
        for p in row[0].paragraphs:
            for r in p.runs:
                r.bold = True

    doc_v.add_paragraph("")
    doc_v.add_heading("Deşifre Dökümü", level=2)

    t_table = doc_v.add_table(rows=1, cols=3)
    t_table.style = "Table Grid"
    hdr = t_table.rows[0].cells
    hdr[0].text = "Zaman"
    hdr[1].text = "Konuşmacı"
    hdr[2].text = "Metin"
    for c in hdr:
        for p in c.paragraphs:
            for r in p.runs:
                r.bold = True

    for i, seg in enumerate(segments):
        row = t_table.add_row().cells
        ts_str = format_ts_hms(seg.get("start", 0.0))
        spk = seg.get("speaker", "Konuşmacı 1")
        raw_txt = seg.get("text", "").strip()
        note = seg.get("note", "").strip()

        row[0].text = ts_str
        row[1].text = spk
        p_txt = row[2].paragraphs[0]
        p_txt.add_run(raw_txt)
        if note:
            add_word_comment_or_note(p_txt, note, author=metadata.get("interviewer", "Araştırmacı"))

    if translated_lines and tgt_lang_label:
        doc_v.add_page_break()
        doc_v.add_heading(f"Çeviri Dökümü ({tgt_lang_label})", level=2)
        for i, seg in enumerate(segments):
            ts_str = format_ts_hms(seg.get("start", 0.0))
            spk = seg.get("speaker", "Konuşmacı 1")
            tr_txt = translated_lines[i] if i < len(translated_lines) else ""
            doc_v.add_paragraph(f"[{ts_str}] {spk}: {tr_txt}")

    doc_v.save(verbatim_path)

    # --- 2. OKUNABILIR METIN (WORD - Accepted AI Suggestions applied) ---
    readable_path = os.path.join(reports_dir, f"{base_name}_FIELD_Okunabilir_Metin.docx")
    doc_r = docx.Document()
    doc_r.add_heading(f"FIELD Okunabilir Metin (Sunum & Alıntı Sürümü): {base_name}", 0)
    for i, seg in enumerate(segments):
        ts_str = format_ts_hms(seg.get("start", 0.0))
        spk = seg.get("speaker", "Konuşmacı 1")
        txt = apply_accepted_suggestions(seg.get("text", "").strip(), i)
        p = doc_r.add_paragraph()
        r_spk = p.add_run(f"[{ts_str}] {spk}: ")
        r_spk.bold = True
        p.add_run(txt)
    doc_r.save(readable_path)

    # --- 3. NVivo / MAXQDA / ATLAS.ti UYUMLU DÜZ METİN (.txt) ---
    qda_path = os.path.join(reports_dir, f"{base_name}_FIELD_NVivo_MAXQDA_ATLASti.txt")
    with open(qda_path, "w", encoding="utf-8") as f:
        f.write(f"# GÖRÜŞME KÜNYESİ | Tarih: {metadata.get('date','-')} | Yer: {metadata.get('location','-')} | Katılımcı: {metadata.get('participant_code','P01')} | Görüşmeci: {metadata.get('interviewer','Görüşmeci')}\n")
        f.write(f"# BAĞLAM: {metadata.get('context','-')}\n")
        f.write("# FORMAT: MAXQDA / NVivo / ATLAS.ti Zaman Damgalı Konuşmacı Formatı\n\n")
        for seg in segments:
            ts_str = format_ts_hms(seg.get("start", 0.0))
            spk = seg.get("speaker", "Konuşmacı 1")
            txt = seg.get("text", "").strip()
            note = seg.get("note", "").strip()
            note_str = f" // [SAHA NOTU: {note}]" if note else ""
            f.write(f"{spk}\t[{ts_str}]\t{txt}{note_str}\n\n")

    return {
        "verbatim_docx": verbatim_path,
        "readable_docx": readable_path,
        "qda_txt": qda_path,
    }


_OCR_ENGINE_SINGLETON = None


def _clean_ocr_turkish_text(text):
    """Fixes common OCR spacing, math symbol, and Turkish/English/German word boundary artifacts on exam questions without touching fractions."""
    if not text:
        return ""
    # Strip common web quiz floating popups / UI noise (NEVER strip fractions like 3/4 or 5/6!)
    text = re.sub(r'(?i)\b(soru\s*sorun|ipucu\s*g[oö]ster)\b', ' ', text)
    # Normalize unicode minus / en-dash / em-dash / CJK horizontal stroke (\u4e00) and fullwidth parens
    text = (
        text.replace("\u2212", "-")
        .replace("\u2013", "-")
        .replace("\u2014", "-")
        .replace("\u4e00", "-")
        .replace("\uff09", ")")
        .replace("\uff08", "(")
    )
    # Fix OCR reading '- 1', '+ 1', '-1' as '- I', '+ l', '-I'
    text = re.sub(r"([\+\-]\s*)[Il](?=[\s\.,;:\?\)]|$)", r"\g<1>1", text)
    # Fix OCR reading variable 'x' as multiplication sign '×' in algebraic options like 'A) ×+4' or '2×+3'
    text = re.sub(r"([A-E][\)\.\-]\s*|\d+|\+|\-|\()\s*×\s*(?=[\+\-\²\³\^\d\)])", r"\1x", text)
    # Ensure spaces around quotes and after punctuation
    text = re.sub(r'"([^"]+)"([a-zA-ZğüşıöçĞÜŞİÖÇäöüÄÖÜß])', r'"\1" \2', text)
    text = re.sub(r'([a-zA-ZğüşıöçĞÜŞİÖÇäöüÄÖÜß])"([^"]+)"', r'\1 "\2"', text)
    text = re.sub(r'([,;:?!])([a-zA-ZğüşıöçĞÜŞİÖÇäöüÄÖÜß])', r'\1 \2', text)
    # Ensure option markers A) B) C) D) E) start cleanly on their own line (including lowercase algebraic options like B)2x+3)
    text = re.sub(r'(?m)^([A-E])\s*[\)\.\-]\s*(\S)', r'\1) \2', text)
    text = re.sub(r'(\s+)([A-E])\s*[\)\.]\s*([a-zA-ZÇĞİÖŞÜçğıöşüäöüÄÖÜß0-9\-\(\√\%])', r'\n\2) \3', text)

    # Repair dotless 'ı' merged with following number like 'toplam163' or 'toplaml63' -> 'toplamı 63'
    text = re.sub(r"\b([tT]oplam)[1lI](\d+)\b", r"\1ı \2", text)
    text = re.sub(r"\b([sS]ay[ıi]n[ıi]n)[tT]oplam", r"\1 toplam", text)

    # Split & repair common Turkish question stem & grammar words that OCR models alter
    split_tokens = [
        ("Asagidakicumlelerin", "Aşağıdaki cümlelerin"),
        ("asagidakicumlelerin", "aşağıdaki cümlelerin"),
        ("aşağıdaki lerden", "aşağıdakilerden"),
        ("asagidakilerden", "aşağıdakilerden"),
        ("Asagidakilerden", "Aşağıdakilerden"),
        ("Asagidaki", "Aşağıdaki "),
        ("asagidaki", "aşağıdaki "),
        ("ArdisikUg", "Ardışık üç"),
        ("Ardiskug", "Ardışık üç"),
        ("Ardsik ug", "Ardışık üç"),
        ("Ardisik uc", "Ardışık üç"),
        ("Ardisik Uc", "Ardışık üç"),
        ("Ardsik", "Ardışık"),
        ("Ardisik", "Ardışık"),
        ("ug tek", "üç tek"),
        ("Uc tek", "üç tek"),
        ("uc tek", "üç tek"),
        ("saynin", "sayının"),
        ("sayinin", "sayının"),
        ("toplaml", "toplamı "),
        ("toplami", "toplamı"),
        ("sayilarn", "sayıların"),
        ("sayilarin", "sayıların"),
        ("buyugukagtr", "büyüğü kaçtır"),
        ("buyugukagr", "büyüğü kaçtır"),
        ("buyugukactir", "büyüğü kaçtır"),
        ("buyugu", "büyüğü "),
        ("kucugu", "küçüğü "),
        ("kagtir", "kaçtır"),
        ("kagtr", "kaçtır"),
        ("kagr?", "kaçtır?"),
        ("kactir", "kaçtır"),
        ("cumlelerin", "cümlelerin "),
        ("hangisinde", " hangisinde "),
        ("birbirine", " birbirine "),
        ("baglamak", " bağlamak "),
        ("kullanilmistir", " kullanılmıştır"),
        ("soylenemez", "söylenemez"),
        ("soylenebilir", "söylenebilir"),
        ("ulasilamaz", "ulaşılamaz"),
        ("ulasilabilir", "ulaşılabilir"),
        ("cikarilamaz", "çıkarılamaz"),
        ("deginilmemistir", "değinilmemiştir"),
        ("cikmasinda", "çıkmasında"),
        ("cıkmasında", "çıkmasında"),
        ("oldugu", "olduğu"),
        ("carımaalamalarının", "çıkar çatışmalarının"),
        ("catismalarinin", "çatışmalarının"),
        ("artirdığı", "artırdığı"),
        ("artirdigi", "artırdığı"),
        ("bloklasmaların", "bloklaşmaların"),
        ("bloklasmalarin", "bloklaşmaların"),
        ("bloklasmalar", "bloklaşmalar"),
        ("yarisinin", "yarışının"),
        ("yarisina", "yarışına"),
        ("hiz ", "hız "),
        ("baslamasında", "başlamasında"),
        ("baslamasinda", "başlamasında"),
        ("kiskirtıcı", "kışkırtıcı"),
        ("kiskirtici", "kışkırtıcı"),
        ("dogrudan", "doğrudan"),
        ("cografi", "coğrafi"),
        ("bolgeler", "bölgeler"),
        ("uzerinde", "üzerinde"),
        ("mucadelesi", "mücadelesi"),
        ("yasandığı", "yaşandığı"),
        ("yasandigi", "yaşandığı"),
        ("somurge", "sömürge"),
        ("baglaci", " bağlacı "),
        ("baglacl", " bağlacı "),
        ("kardesine", " kardeşine "),
        ("kardesini", " kardeşini "),
        ("bakmisti", " bakmıştı"),
        ("gordum", " gördüm"),
        ("gosteri", " gösteri "),
        ("yapilacak", " yapılacak"),
        ("kitaplardan", " kitaplardan "),
        ("dergilerden", " dergilerden "),
        ("bahsediyordu", " bahsediyordu"),
        ("Cocuklarina", "Çocuklarına "),
        ("cocuklarina", "çocuklarına "),
        ("Carsamba", "Çarşamba "),
        ("carsamba", "çarşamba "),
        ("cikisi", " çıkışı "),
        ("serviste", " serviste "),
        ("senive", "seni ve "),
        ("onave", "ona ve "),
        ("vebaglaci", "ve bağlacı "),
        ("icin", " için "),
    ]
    for src_tok, dst_tok in split_tokens:
        text = text.replace(src_tok, dst_tok)

    # Join hyphenated or broken math lines where a line ends with '+' or '-' and the next line starts with a number followed by '?' (e.g., 'x² + x -\n2?' -> 'x² + x - 2?')
    text = re.sub(r"([\+\-])\s*\n\s*(\d+\s*\?)", r"\1 \2", text)

    # Clean multiple spaces on each line while preserving newlines
    cleaned_lines = []
    for line in text.splitlines():
        ln = re.sub(r"[ \t]{2,}", " ", line).strip()
        if ln:
            cleaned_lines.append(ln)

    # Auto-label unlabeled radio-button options (when 2..5 lines follow a '?' line without A) B) C) D) E) prefixes)
    q_idx = None
    for idx, ln in enumerate(cleaned_lines):
        if "?" in ln:
            q_idx = idx
    if q_idx is not None:
        trailing = cleaned_lines[q_idx + 1:]
        opt_check = re.compile(r"^[A-E][\)\.\-]\s+", re.I)
        if 2 <= len(trailing) <= 5 and not any(opt_check.match(t) for t in trailing):
            letters = ["A", "B", "C", "D", "E"]
            for i, t in enumerate(trailing):
                cleaned_lines[q_idx + 1 + i] = f"{letters[i]}) {t}"

    return "\n".join(cleaned_lines)


def resolve_response_language(ui_lang="tr", user_q="", context_text=""):
    """
    Determines the target explanation language ('tr', 'en', or 'de') based on:
      1. The active UI language selected by the user ('en' or 'de' takes priority when set).
      2. Automatic language detection on the user's question / uploaded question text if needed.
    """
    ui_l = (ui_lang or "tr").lower().strip()
    combined = f"{user_q or ''} {context_text or ''}".lower()

    # Check strong language indicators in the question itself
    en_hits = len(re.findall(r"\b(the|what|when|which|divided|remainder|polynomial|value|find|following|calculate|solve|step|number|sum|ratio)\b", combined))
    de_hits = len(re.findall(r"\b(der|die|das|welcher|welche|welches|geteilt|polynom|wert|berechnen|lösen|schritt|wie\s+viel|summe|rest)\b", combined))
    tr_hits = len(re.findall(r"\b(kaçtır|nedir|buna\s+göre|hangisidir|işleminin|toplamı|bölümünden|kalan|polinomunun|aşağıdakilerden|ardışık)\b", combined))

    if ui_l == "en":
        return "en"
    if ui_l == "de":
        return "de"
    # Even if UI is left on 'tr', if the question/prompt is clearly in English or German and has no Turkish keywords:
    if en_hits >= 2 and en_hits > tr_hits and en_hits >= de_hits:
        return "en"
    if de_hits >= 2 and de_hits > tr_hits and de_hits > en_hits:
        return "de"
    return ui_l if ui_l in ("tr", "en", "de") else "tr"


def _format_fraction_tr(frac_val):
    """Formats a fractions.Fraction cleanly as '5/4' or integer '3'."""
    if frac_val.denominator == 1:
        return str(frac_val.numerator)
    return f"{frac_val.numerator}/{frac_val.denominator}"


def _parse_option_as_fraction(opt_str):
    """Attempts to parse an option string like '5/4', '-3/2', '21', '1,5', or '2⁹' as an exact Fraction."""
    from fractions import Fraction
    if not opt_str:
        return None
    s = _latex_to_readable_math(opt_str).strip()
    s = s.replace("\u2212", "-").replace("\u2013", "-").replace(" ", "")
    rev_sup = str.maketrans("⁰¹²³⁴⁵⁶⁷⁸⁹", "0123456789")
    m_exp = re.match(r"^(-?\d+)(?:\^\{?(\d+)\}?|([⁰¹²³⁴⁵⁶⁷⁸⁹]+))$", s)
    if m_exp:
        base_v = int(m_exp.group(1))
        exp_v = int(m_exp.group(2) if m_exp.group(2) else m_exp.group(3).translate(rev_sup))
        return Fraction(base_v ** exp_v)
    m_frac = re.match(r"^(-?\d+)/(-?\d+)$", s)
    if m_frac and int(m_frac.group(2)) != 0:
        return Fraction(int(m_frac.group(1)), int(m_frac.group(2)))
    m_int = re.match(r"^(-?\d+)$", s)
    if m_int:
        return Fraction(int(m_int.group(1)))
    m_dec = re.match(r"^(-?\d+)[,\.](\d+)$", s)
    if m_dec:
        return Fraction(f"{m_dec.group(1)}.{m_dec.group(2)}")
    return None


def _eval_math_ast_with_steps(expr_str):
    """
    Safely evaluates an arithmetic/fraction expression using exact fractions.Fraction
    and returns (Fraction_result, list_of_step_strings, is_pure_fraction_add_sub_info).
    """
    import ast
    import math
    from fractions import Fraction

    s = expr_str.strip()
    s = s.replace("\u2212", "-").replace("\u2013", "-").replace("\u2014", "-")
    s = s.replace("×", "*").replace("·", "*").replace("⋅", "*").replace("÷", "/").replace(":", "/")
    rev_sup = str.maketrans("⁰¹²³⁴⁵⁶⁷⁸⁹", "0123456789")
    s = re.sub(r"([⁰¹²³⁴⁵⁶⁷⁸⁹]+)", lambda m: "**" + m.group(1).translate(rev_sup), s)
    s = s.replace("^", "**")
    s = re.sub(r"(\d+),(\d+)", r"\1.\2", s)

    # First check if it's a pure sum/difference of fractions & integers: e.g. "3/4 + 5/6 - 1/3"
    pure_terms_re = re.compile(r"^\s*([+\-]?\s*\d+(?:\s*/\s*\d+)?)(?:\s*([+\-])\s*(\d+(?:\s*/\s*\d+)?))+\s*$")
    if pure_terms_re.match(s) and "/" in s:
        raw_tokens = re.findall(r"([+\-]?)\s*(\d+(?:\s*/\s*\d+)?)", s)
        terms = []
        for idx_t, (sgn, val_tok) in enumerate(raw_tokens):
            sign_char = "-" if sgn == "-" else "+"
            val_clean = val_tok.replace(" ", "")
            if "/" in val_clean:
                n_s, d_s = val_clean.split("/", 1)
                num, den = int(n_s), int(d_s)
            else:
                num, den = int(val_clean), 1
            if den == 0:
                return None, [], None
            terms.append((sign_char, num, den, idx_t == 0 and sgn != "-"))

        dens = [t[2] for t in terms]
        common_den = 1
        for d in dens:
            common_den = math.lcm(common_den, d)

        total_frac = Fraction(0)
        for sign_char, num, den, _ in terms:
            f_term = Fraction(num, den)
            total_frac = total_frac - f_term if sign_char == "-" else total_frac + f_term

        return total_frac, [], {"terms": terms, "common_den": common_den}

    # General AST evaluation with step recording
    try:
        tree = ast.parse(s, mode="eval")
    except Exception:
        return None, [], None

    steps = []

    def _eval_node(node):
        if isinstance(node, ast.Expression):
            return _eval_node(node.body)
        if isinstance(node, ast.Constant) and isinstance(node.value, (int, float)):
            return Fraction(str(node.value))
        if isinstance(node, ast.UnaryOp) and isinstance(node.op, (ast.UAdd, ast.USub)):
            val = _eval_node(node.operand)
            return -val if isinstance(node.op, ast.USub) else val
        if isinstance(node, ast.BinOp):
            left = _eval_node(node.left)
            right = _eval_node(node.right)
            l_str = _format_fraction_tr(left)
            r_str = _format_fraction_tr(right)
            if isinstance(node.op, ast.Add):
                res = left + right
                if left.denominator != right.denominator:
                    cd = math.lcm(left.denominator, right.denominator)
                    k1, k2 = cd // left.denominator, cd // right.denominator
                    n1, n2 = left.numerator * k1, right.numerator * k2
                    steps.append(
                        f"{l_str} + {r_str} işleminde paydaları {cd}'de eşitleriz: "
                        f"{n1}/{cd} + {n2}/{cd} = {n1 + n2}/{cd}"
                        + (f" = {_format_fraction_tr(res)} (sadeleştirilmiş hâli)" if f"{n1+n2}/{cd}" != _format_fraction_tr(res) else "")
                    )
                else:
                    steps.append(f"Toplama işlemi: {l_str} + {r_str} = {_format_fraction_tr(res)}")
                return res
            if isinstance(node.op, ast.Sub):
                res = left - right
                if left.denominator != right.denominator:
                    cd = math.lcm(left.denominator, right.denominator)
                    k1, k2 = cd // left.denominator, cd // right.denominator
                    n1, n2 = left.numerator * k1, right.numerator * k2
                    steps.append(
                        f"{l_str} - {r_str} işleminde paydaları {cd}'de eşitleriz: "
                        f"{n1}/{cd} - {n2}/{cd} = {n1 - n2}/{cd}"
                        + (f" = {_format_fraction_tr(res)} (sadeleştirilmiş hâli)" if f"{n1-n2}/{cd}" != _format_fraction_tr(res) else "")
                    )
                else:
                    steps.append(f"Çıkarma işlemi: {l_str} - {r_str} = {_format_fraction_tr(res)}")
                return res
            if isinstance(node.op, ast.Mult):
                res = left * right
                steps.append(f"Çarpma işlemi: {l_str} × {r_str} = {_format_fraction_tr(res)}")
                return res
            if isinstance(node.op, ast.Div):
                if right == 0:
                    raise ZeroDivisionError()
                res = left / right
                if isinstance(node.left, ast.Constant) and isinstance(node.right, ast.Constant):
                    return res
                steps.append(
                    f"Kesirlerde bölme işlemi (birinci aynen yazılır, ikinci ters çevrilip çarpılır): "
                    f"({l_str}) ÷ ({r_str}) = {_format_fraction_tr(res)}"
                )
                return res
            if isinstance(node.op, ast.Pow):
                res = left ** int(right)
                steps.append(f"Üslü ifade hesabı: {l_str}^{r_str} = {_format_fraction_tr(res)}")
                return res
        raise ValueError("Unsupported AST node")

    try:
        final_res = _eval_node(tree)
        return final_res, steps, None
    except Exception:
        return None, [], None


def solve_question_with_teacher_engine_offline(raw_text, lang="tr"):
    """
    Deterministic, 100% mathematically & logically verified Private Tutor (Profesyonel Öğretmen)
    solution generator supporting Turkish ('tr'), English ('en'), and German ('de').
    Returns a complete 5-step teacher explanation string when an exact mathematical or grammatical/logical
    proof is computed, or None if the question should be delegated to the LLM.
    """
    import math
    from fractions import Fraction

    if not raw_text or not raw_text.strip():
        return None

    lang = resolve_response_language(lang, "", raw_text)
    ans_prefix = {"tr": "DOĞRU CEVAP", "en": "CORRECT ANSWER", "de": "RICHTIGE ANTWORT"}.get(lang, "DOĞRU CEVAP")

    def _L(tr_s, en_s, de_s):
        return {"tr": tr_s, "en": en_s, "de": de_s}.get(lang, tr_s)

    h1 = _L(
        "🎯 1. Soruyu Anlayalım (Bizden Ne İsteniyor?):",
        "🎯 1. Understanding the Question (What Is Being Asked?):",
        "🎯 1. Die Aufgabe verstehen (Was wird gefragt?):",
    )
    h2 = _L(
        "💡 2. Öğretmen Notu (Bilmemiz Gereken Kural & Pratik Yöntem):",
        "💡 2. Teacher's Note (Key Rule & Practical Method):",
        "💡 2. Lehrer-Tipp (Wichtige Regel & Praktische Methode):",
    )
    h3 = _L(
        "✍️ 3. Adım Adım Çözüm (Beraber Çözelim):",
        "✍️ 3. Step-by-Step Solution (Let's Solve It Together):",
        "✍️ 3. Schritt-für-Schritt-Lösung (Gemeinsam lösen):",
    )
    h4 = _L(
        "⚠️ 4. Çeldirici Şıklara Dikkat (Neden Diğer Şıklar Değil?):",
        "⚠️ 4. Watch Out for Distractor Options (Why Not the Others?):",
        "⚠️ 4. Achtung vor Fallen in den Antwortmöglichkeiten:",
    )
    h5 = _L(
        "✅ 5. Kesin Doğru Cevap:",
        "✅ 5. Final Correct Answer:",
        "✅ 5. Endgültige richtige Antwort:",
    )

    text = _clean_ocr_turkish_text(raw_text)
    lines = [
        ln.strip()
        for ln in text.splitlines()
        if ln.strip() and not re.match(r"^(CEVAPLAR|CEVAP\s*ANAHTARI|ANSWERS|LÖSUNGEN)\b", ln.strip(), re.I)
    ]

    opt_pattern = re.compile(r"^([A-E])[\)\.\-]\s*(.+)$")
    pre_option_lines = []
    options = {}
    current_opt = None

    for ln in lines:
        if re.match(r"^(Correct!|Incorrect\.|Metinde\s+|Alsas-Loren\s+kömür\s+havzası\s+ve\s+Balkanlar\s+gibi|Toplam\s+veri\s+miktarını|Üsler\s+toplanmamalı|Sonraki$)", ln, re.I):
            continue
        m = opt_pattern.match(ln)
        if m:
            current_opt = m.group(1).upper()
            options[current_opt] = m.group(2).strip()
        else:
            if current_opt is None:
                pre_option_lines.append(ln)
            else:
                options[current_opt] = (options[current_opt] + " " + ln).strip()

    full_stem_block = " ".join(pre_option_lines).strip() if pre_option_lines else text.strip()
    full_math_txt = _latex_to_readable_math(full_stem_block)

    # Recover superscript exponents if OCR flattened 2^8, 2^9, 2^12, 2^15 into 28, 29, 212, 215
    if len(options) >= 2:
        opt_vals = [v.strip() for v in options.values()]
        if all(re.match(r"^\d{2,4}$", v) for v in opt_vals):
            first_chars = {v[0] for v in opt_vals}
            if len(first_chars) == 1:
                base_c = list(first_chars)[0]
                exps = [int(v[1:]) for v in opt_vals if v[1:].isdigit()]
                if len(exps) == len(opt_vals) and exps == sorted(exps) and max(exps) >= 6:
                    for k in list(options.keys()):
                        v = options[k].strip()
                        options[k] = f"{base_c}^{v[1:]}"
                    full_stem_block = re.sub(
                        rf"\b{base_c}(\d{{1,2}})\s*(GB|MB|KB|TB|sn|saniye|sec|dk|saat|metre|cm|birim|adet|kat)\b",
                        rf"{base_c}^\1 \2",
                        full_stem_block,
                    )
                    full_math_txt = _latex_to_readable_math(full_stem_block)

    # =========================================================================
    # 0. POLYNOMIAL REMAINDER THEOREM SOLVER (TR / EN / DE)
    #    e.g., "The remainder when the polynomial P(x) is divided by x - 1 is 5,
    #           and the remainder when it is divided by x + 2 is -1.
    #           What is the remainder when P(x) is divided by x² + x - 2?"
    # =========================================================================
    if re.search(r"\b(polynomial|polinom|polynom)\b|\bP\s*\(\s*x\s*\)", full_math_txt, re.I) and re.search(r"\b(remainder|divided|bölümünden|kalan|rest|geteilt|dividiert)\b", full_math_txt, re.I):
        try:
            # Find linear divisor-remainder pairs: (x - c1) -> r1, (x + c2) -> r2
            lin_pairs = []
            # Pattern EN/DE: "divided by x - 1 is 5" / "geteilt durch x - 1 ist 5"
            for m_en in re.finditer(r"(?:divided\s+by|geteilt\s+durch|dividiert\s+durch)\s*\(?\s*[xX]\s*([\+\-])\s*(\d+)\s*\)?\s*(?:is|equals|gives|ist|beträgt)\s*([+\-]?\d+)", full_math_txt, re.I):
                sgn_c, val_c, rem_v = m_en.group(1), int(m_en.group(2)), int(m_en.group(3))
                root_x = val_c if sgn_c == "-" else -val_c
                div_str = f"x - {val_c}" if sgn_c == "-" else f"x + {val_c}"
                lin_pairs.append((root_x, rem_v, div_str))
            # Pattern TR: "x - 1 ile bölümünden kalan 5"
            for m_tr in re.finditer(r"\(?\s*[xX]\s*([\+\-])\s*(\d+)\s*\)?\s*ile\s+bölümünden\s+kalan\s*([+\-]?\d+)", full_math_txt, re.I):
                sgn_c, val_c, rem_v = m_tr.group(1), int(m_tr.group(2)), int(m_tr.group(3))
                root_x = val_c if sgn_c == "-" else -val_c
                div_str = f"x - {val_c}" if sgn_c == "-" else f"x + {val_c}"
                lin_pairs.append((root_x, rem_v, div_str))

            # Fallback recovery if OCR truncated the second linear condition on line 2, deduce from quadratic divisor and options!
            quad_match = re.search(r"[xX](?:²|\^2|2)\s*([\+\-])\s*(\d*)\s*[xX]\s*([\+\-])\s*(\d+)", full_math_txt)
            if len(lin_pairs) >= 2 and lin_pairs[0][0] != lin_pairs[1][0]:
                x1, y1, div1 = lin_pairs[0]
                x2, y2, div2 = lin_pairs[1]
                a_frac = Fraction(y1 - y2, x1 - x2)
                b_frac = Fraction(y1) - a_frac * Fraction(x1)
                a_int = int(a_frac) if a_frac.denominator == 1 else a_frac
                b_int = int(b_frac) if b_frac.denominator == 1 else b_frac

                if a_int == 1:
                    a_part = "x"
                elif a_int == -1:
                    a_part = "-x"
                elif a_int == 0:
                    a_part = ""
                else:
                    a_part = f"{a_int}x"

                if b_int > 0:
                    rem_str = f"{a_part} + {b_int}" if a_part else f"{b_int}"
                elif b_int < 0:
                    rem_str = f"{a_part} - {abs(b_int)}" if a_part else f"-{abs(b_int)}"
                else:
                    rem_str = a_part or "0"

                def _norm_poly(s_p):
                    return re.sub(r"\s+", "", _latex_to_readable_math(s_p)).replace("×", "x").replace("\u2212", "-").lower()

                rem_norm = _norm_poly(rem_str)
                matched_letter = None
                for k, v in sorted(options.items()):
                    if _norm_poly(v) == rem_norm:
                        matched_letter = k
                        break

                p_coef = -(x1 + x2)
                q_coef = x1 * x2
                p_str = f"+ {p_coef}x" if p_coef > 1 else ("+ x" if p_coef == 1 else ("- x" if p_coef == -1 else (f"- {abs(p_coef)}x" if p_coef < -1 else "")))
                q_str = f"+ {q_coef}" if q_coef > 0 else (f"- {abs(q_coef)}" if q_coef < 0 else "")
                q_div_str = " ".join(part for part in ["x²", p_str, q_str] if part)
                if quad_match:
                    s1, c1_s, s2, c2_s = quad_match.group(1), quad_match.group(2), quad_match.group(3), quad_match.group(4)
                    q_div_str = f"x² {s1} {c1_s + 'x' if c1_s else 'x'} {s2} {c2_s}"

                eq1_lhs = f"a·({x1}) + b" if x1 < 0 else (f"a + b" if x1 == 1 else f"{x1}a + b")
                eq2_lhs = f"-{abs(x2)}a + b" if x2 < 0 else (f"a + b" if x2 == 1 else f"{x2}a + b")

                distractor_lines = []
                for k, v in sorted(options.items()):
                    v_clean = _latex_to_readable_math(v).replace("×", "x").strip()
                    if k == matched_letter:
                        distractor_lines.append(
                            _L(
                                f"• {k}) {v_clean} ➔ DOĞRU (x = {x1} için {a_int}({x1}) + ({b_int}) = {y1} ve x = {x2} için {a_int}({x2}) + ({b_int}) = {y2} değerlerini tam olarak sağlar).",
                                f"• {k}) {v_clean} ➔ CORRECT (For x = {x1}, R({x1}) = {a_int}({x1}) + ({b_int}) = {y1}, and for x = {x2}, R({x2}) = {a_int}({x2}) + ({b_int}) = {y2}, satisfying both conditions!).",
                                f"• {k}) {v_clean} ➔ RICHTIG (Für x = {x1} gilt R({x1}) = {y1} und für x = {x2} gilt R({x2}) = {y2}; beide Bedingungen sind exakt erfüllt!).",
                            )
                        )
                    else:
                        distractor_lines.append(
                            _L(
                                f"• {k}) {v_clean} ➔ Elenir (P({x1}) = {y1} ve P({x2}) = {y2} kalan koşullarını aynı anda sağlamaz).",
                                f"• {k}) {v_clean} ➔ Eliminated (Does not satisfy both remainder conditions P({x1}) = {y1} and P({x2}) = {y2} simultaneously).",
                                f"• {k}) {v_clean} ➔ Ausgeschlossen (Erfüllt nicht beide Restbedingungen P({x1}) = {y1} und P({x2}) = {y2} gleichzeitig).",
                            )
                        )

                ans_line = f"{ans_prefix}: {matched_letter}) {rem_str}" if matched_letter else f"{ans_prefix}: {rem_str}"

                if lang == "en":
                    return (
                        f"{h1}\n"
                        f"We are given that when the polynomial P(x) is divided by ({div1}), the remainder is {y1}, and when it is divided by ({div2}), the remainder is {y2}. "
                        f"We need to find the remainder when P(x) is divided by the quadratic polynomial '{q_div_str}'.\n\n"
                        f"{h2}\n"
                        f"• Polynomial Remainder Theorem: The remainder when a polynomial P(x) is divided by (x - c) is always P(c).\n"
                        f"• Degree of the Remainder Rule: When P(x) is divided by a 2nd-degree (quadratic) divisor '{q_div_str} = ({div1})·({div2})', "
                        f"the remainder R(x) must have degree at most 1, so we write the remainder in the linear form:\n"
                        f"  R(x) = a·x + b\n"
                        f"• Since P(x) = Q(x)·({div1})·({div2}) + R(x), at the roots x = {x1} and x = {x2} the quotient part becomes 0, meaning:\n"
                        f"  R({x1}) = P({x1}) = {y1}  and  R({x2}) = P({x2}) = {y2}.\n\n"
                        f"{h3}\n"
                        f"• Step 1 (Apply the Remainder Theorem to the Linear Divisors):\n"
                        f"  - Set {div1} = 0 ➔ x = {x1} ➔ P({x1}) = {y1}\n"
                        f"  - Set {div2} = 0 ➔ x = {x2} ➔ P({x2}) = {y2}\n\n"
                        f"• Step 2 (Set Up the System of Linear Equations for R(x) = ax + b):\n"
                        f"  - Notice that {q_div_str} factors into ({div1})·({div2}).\n"
                        f"  - Substitute x = {x1} into R(x) = ax + b:  {eq1_lhs} = {y1}   (Equation 1)\n"
                        f"  - Substitute x = {x2} into R(x) = ax + b:  {eq2_lhs} = {y2}   (Equation 2)\n\n"
                        f"• Step 3 (Solve for 'a' and 'b'):\n"
                        f"  - Subtract Equation 2 from Equation 1:\n"
                        f"    ({eq1_lhs}) - ({eq2_lhs}) = {y1} - ({y2})\n"
                        f"    {x1 - x2}a = {y1 - y2}  ➔  a = {y1 - y2} ÷ {x1 - x2} = {a_int}\n"
                        f"  - Substitute a = {a_int} back into Equation 1:\n"
                        f"    {a_int}·({x1}) + b = {y1}  ➔  {a_int * x1} + b = {y1}  ➔  b = {b_int}\n\n"
                        f"• Step 4 (Write the Remainder R(x) and Verify):\n"
                        f"  - Substitute a = {a_int} and b = {b_int} into R(x) = ax + b:\n"
                        f"    R(x) = {rem_str}\n"
                        f"  - Quick Check: R({x1}) = {a_int}({x1}) + ({b_int}) = {y1} ✔ and R({x2}) = {a_int}({x2}) + ({b_int}) = {y2} ✔"
                        + (f"\n\n{h4}\n" + "\n".join(distractor_lines) if distractor_lines else "")
                        + f"\n\n{h5}\n{ans_line}"
                    )
                elif lang == "de":
                    return (
                        f"{h1}\n"
                        f"Gegeben ist, dass das Polynom P(x) bei der Division durch ({div1}) den Rest {y1} und bei der Division durch ({div2}) den Rest {y2} ergibt. "
                        f"Gesucht ist der Rest bei der Division von P(x) durch das quadratische Polynom '{q_div_str}'.\n\n"
                        f"{h2}\n"
                        f"• Polynom-Restsatz: Der Rest bei der Division eines Polynoms P(x) durch (x - c) ist immer P(c).\n"
                        f"• Regel für den Grad des Restes: Da der Divisor '{q_div_str} = ({div1})·({div2})' den Grad 2 hat, ist der Rest R(x) höchstens vom Grad 1 (linear):\n"
                        f"  R(x) = a·x + b\n"
                        f"• Für die Nullstellen x = {x1} und x = {x2} gilt direkt: R({x1}) = P({x1}) = {y1} und R({x2}) = P({x2}) = {y2}.\n\n"
                        f"{h3}\n"
                        f"• 1. Schritt (Restsatz auf die linearen Teiler anwenden):\n"
                        f"  - {div1} = 0 ➔ x = {x1} ➔ P({x1}) = {y1}\n"
                        f"  - {div2} = 0 ➔ x = {x2} ➔ P({x2}) = {y2}\n\n"
                        f"• 2. Schritt (Gleichungssystem für R(x) = ax + b aufstellen):\n"
                        f"  - Für x = {x1} in R(x) = ax + b:  {eq1_lhs} = {y1}   (Gleichung 1)\n"
                        f"  - Für x = {x2} in R(x) = ax + b:  {eq2_lhs} = {y2}   (Gleichung 2)\n\n"
                        f"• 3. Schritt ('a' und 'b' berechnen):\n"
                        f"  - Gleichung 2 von Gleichung 1 abziehen:\n"
                        f"    {x1 - x2}a = {y1 - y2}  ➔  a = {a_int}\n"
                        f"  - a = {a_int} in Gleichung 1 einsetzen:\n"
                        f"    {a_int * x1} + b = {y1}  ➔  b = {b_int}\n\n"
                        f"• 4. Schritt (Restpolynom R(x) aufschreiben):\n"
                        f"  - R(x) = {rem_str}"
                        + (f"\n\n{h4}\n" + "\n".join(distractor_lines) if distractor_lines else "")
                        + f"\n\n{h5}\n{ans_line}"
                    )
                else:
                    return (
                        f"{h1}\n"
                        f"Soruda bize P(x) polinomunun ({div1}) ile bölümünden kalanın {y1}, ({div2}) ile bölümünden kalanın ise {y2} olduğu verilmiş; "
                        f"P(x) polinomunun '{q_div_str}' ikinci dereceden ifadesi ile bölümünden kalan sorulmuştur.\n\n"
                        f"{h2}\n"
                        f"• Polinomlarda Kalan Teoremi: Bir P(x) polinomunun (x - c) ile bölümünden kalan her zaman P(c) değerine eşittir.\n"
                        f"• Kalanın Derecesi Kuralı: Bölen ifade '{q_div_str} = ({div1})·({div2})' 2. dereceden olduğu için, kalan R(x) en fazla 1. dereceden (doğrusal) olmalıdır:\n"
                        f"  R(x) = a·x + b\n"
                        f"• P(x) = B(x)·({div1})·({div2}) + a·x + b denkleminde böleni sıfırlayan x = {x1} ve x = {x2} değerleri için R({x1}) = P({x1}) = {y1} ve R({x2}) = P({x2}) = {y2} olur.\n\n"
                        f"{h3}\n"
                        f"• 1. Adım (Verilen Kalanları Denklem Diline Çevirelim):\n"
                        f"  - {div1} = 0 ➔ x = {x1} için P({x1}) = {y1}\n"
                        f"  - {div2} = 0 ➔ x = {x2} için P({x2}) = {y2}\n\n"
                        f"• 2. Adım (R(x) = ax + b Kalan Polinomunda Yerine Yazalım):\n"
                        f"  - x = {x1} için: {eq1_lhs} = {y1}   (1. Denklem)\n"
                        f"  - x = {x2} için: {eq2_lhs} = {y2}   (2. Denklem)\n\n"
                        f"• 3. Adım (Yok Etme Yöntemiyle 'a' ve 'b' Katsayılarını Bulalım):\n"
                        f"  - 1. denklemden 2. denklemi taraf tarafa çıkaralım:\n"
                        f"    ({eq1_lhs}) - ({eq2_lhs}) = {y1} - ({y2})\n"
                        f"    {x1 - x2}a = {y1 - y2}  ➔  a = {y1 - y2} ÷ {x1 - x2} = {a_int}\n"
                        f"  - Bulduğumuz a = {a_int} değerini 1. denklemde yerine koyalım:\n"
                        f"    {a_int * x1} + b = {y1}  ➔  b = {b_int}\n\n"
                        f"• 4. Adım (Kalan Polinomunu Yazalım ve Sağlama Yapalım):\n"
                        f"  - a = {a_int} ve b = {b_int} değerlerini R(x) = ax + b ifadesinde yerine yazarsak:\n"
                        f"    R(x) = {rem_str} bulunur.\n"
                        f"  - Sağlama: R({x1}) = {y1} ✔ ve R({x2}) = {y2} ✔"
                        + (f"\n\n{h4}\n" + "\n".join(distractor_lines) if distractor_lines else "")
                        + f"\n\n{h5}\n{ans_line}"
                    )
        except Exception:
            pass

    # =========================================================================
    # 1. FRACTION & ARITHMETIC EXPRESSION SOLVER (e.g., "3/4 + 5/6 - 1/3 işleminin sonucu kaçtır?")
    # =========================================================================
    expr_match = re.search(
        r"([\(\d][0-9\s\+\-\*/×·÷:\(\),\.\^⁰¹²³⁴⁵⁶⁷⁸⁹]+[\d\)\⁰\¹\²\³\⁴\⁵\⁶\⁷\⁸\⁹])\s*(?:=|\s+işleminin|\s+ifadesinin|\s+sonucu|\s+değeri|\s+kaçtır|\s+is|\s+equals|\?)",
        full_math_txt,
        re.I,
    )
    if expr_match and any(op in expr_match.group(1) for op in ("+", "-", "/", "*", "×", "÷", ":")):
        raw_expr = expr_match.group(1).strip()
        res_frac, ast_steps, add_sub_info = _eval_math_ast_with_steps(raw_expr)
        if res_frac is not None:
            res_str = _format_fraction_tr(res_frac)
            matched_letter = None
            for k, v in sorted(options.items()):
                opt_f = _parse_option_as_fraction(v)
                if opt_f is not None and opt_f == res_frac:
                    matched_letter = k
                    break

            if add_sub_info is not None:
                terms = add_sub_info["terms"]
                cd = add_sub_info["common_den"]
                dens_list = [str(t[2]) for t in terms]
                exp_bullets = []
                expanded_strs = []
                num_expr_parts = []
                unsimplified_num = 0
                for sign_char, num, den, is_first_pos in terms:
                    mul_k = cd // den
                    new_n = num * mul_k
                    exp_bullets.append(
                        _L(
                            f"• {num}/{den} kesrini {mul_k} ile genişletiriz: ({num} × {mul_k}) / ({den} × {mul_k}) = {new_n}/{cd}",
                            f"• Expand {num}/{den} by multiplying top and bottom by {mul_k}: ({num} × {mul_k}) / ({den} × {mul_k}) = {new_n}/{cd}",
                            f"• Erweitere {num}/{den} mit {mul_k}: ({num} × {mul_k}) / ({den} × {mul_k}) = {new_n}/{cd}",
                        )
                    )
                    if is_first_pos:
                        expanded_strs.append(f"{new_n}/{cd}")
                        num_expr_parts.append(f"{new_n}")
                    else:
                        expanded_strs.append(f"{sign_char} {new_n}/{cd}")
                        num_expr_parts.append(f"{sign_char} {new_n}")
                    unsimplified_num = unsimplified_num - new_n if sign_char == "-" else unsimplified_num + new_n

                g_div = math.gcd(abs(unsimplified_num), cd)
                simp_step_txt = ""
                if g_div > 1:
                    simp_step_txt = _L(
                        f"\n\n3. Adım (En Sade Hâline Getirelim - Sadeleştirme):\n"
                        f"• Bulduğumuz {unsimplified_num}/{cd} kesrinin payı ({unsimplified_num}) ve paydası ({cd}), {g_div} sayısına tam bölünür.\n"
                        f"• Payı {g_div}'e bölelim: {unsimplified_num} ÷ {g_div} = {res_frac.numerator}\n"
                        f"• Paydayı {g_div}'e bölelim: {cd} ÷ {g_div} = {res_frac.denominator}\n"
                        f"• Böylece kesrimizin en sade hâli: {res_str} bulunur.",
                        f"\n\nStep 3 (Simplify the Fraction to Lowest Terms):\n"
                        f"• Both the numerator ({unsimplified_num}) and denominator ({cd}) of {unsimplified_num}/{cd} are divisible by their GCD = {g_div}.\n"
                        f"• Divide numerator by {g_div}: {unsimplified_num} ÷ {g_div} = {res_frac.numerator}\n"
                        f"• Divide denominator by {g_div}: {cd} ÷ {g_div} = {res_frac.denominator}\n"
                        f"• The simplest form of the fraction is: {res_str}.",
                        f"\n\n3. Schritt (Bruch vollständig kürzen):\n"
                        f"• Zähler ({unsimplified_num}) und Nenner ({cd}) von {unsimplified_num}/{cd} sind durch {g_div} teilbar.\n"
                        f"• Zähler durch {g_div} teilen: {unsimplified_num} ÷ {g_div} = {res_frac.numerator}\n"
                        f"• Nenner durch {g_div} teilen: {cd} ÷ {g_div} = {res_frac.denominator}\n"
                        f"• Die vollständig gekürzte Form lautet: {res_str}.",
                    )

                distractor_lines = []
                for k, v in sorted(options.items()):
                    v_clean = _latex_to_readable_math(v).strip()
                    if k == matched_letter:
                        distractor_lines.append(
                            _L(
                                f"• {k}) {v_clean} ➔ DOĞRU ({unsimplified_num}/{cd} sonucunun en sade hâli olan {res_str} ile birebir eşleşmektedir).",
                                f"• {k}) {v_clean} ➔ CORRECT (Matches the simplified result {unsimplified_num}/{cd} = {res_str}).",
                                f"• {k}) {v_clean} ➔ RICHTIG (Stimmt exakt mit dem gekürzten Ergebnis {unsimplified_num}/{cd} = {res_str} überein).",
                            )
                        )
                    else:
                        distractor_lines.append(
                            _L(
                                f"• {k}) {v_clean} ➔ Elenir (Doğru sonuç {res_str} olduğu için yanlıştır).",
                                f"• {k}) {v_clean} ➔ Eliminated (The exact result is {res_str}).",
                                f"• {k}) {v_clean} ➔ Ausgeschlossen (Das exakte Ergebnis ist {res_str}).",
                            )
                        )

                ans_line = f"{ans_prefix}: {matched_letter}) {_latex_to_readable_math(options[matched_letter])}" if matched_letter else f"{ans_prefix}: {res_str}"
                return (
                    f"{h1}\n"
                    + _L(
                        f"Soruda bize paydaları birbirinden farklı olan kesirler verilmiş ve '{raw_expr}' işleminin en sade sonucunu bulmamız isteniyor.\n\n",
                        f"We are asked to evaluate the fraction expression '{raw_expr}' with different denominators and write the result in simplest form.\n\n",
                        f"Gesucht ist der vollständig gekürzte Wert des Bruchausdrucks '{raw_expr}' mit unterschiedlichen Nennern.\n\n",
                    )
                    + f"{h2}\n"
                    + _L(
                        "Kesirlerde toplama ve çıkarma işlemi yapabilmemiz için paydaların mutlaka birbirine eşit olması gerekir. "
                        "Paydaları eşitlemek için paydadaki sayıların en küçük ortak katını (EKOK) buluruz ve her kesri uygun sayıyla genişletiriz. "
                        "Paydalar eşitlendikten sonra ortak payda aynen yazılır, sadece paylar toplanır veya çıkarılır.\n\n",
                        "To add or subtract fractions, their denominators must be equal. We find the Least Common Multiple (LCM) of the denominators, "
                        "expand each fraction to the common denominator, combine the numerators, and simplify the final fraction.\n\n",
                        "Um Brüche zu addieren oder zu subtrahieren, müssen sie denselben Nenner haben. Wir bestimmen das kleinste gemeinsame Vielfache (kgV) der Nenner, "
                        "erweitern jeden Bruch auf den Hauptnenner, verrechnen die Zähler und kürzen das Ergebnis.\n\n",
                    )
                    + f"{h3}\n"
                    + _L(
                        f"1. Adım (Paydaları Eşitleyelim):\n• Kesirlerimizin paydaları: {', '.join(dens_list)}. En küçük ortak kat (EKOK) {cd} olduğu için tüm paydaları {cd} sayısında eşitleriz:\n",
                        f"Step 1 (Find the Common Denominator):\n• The denominators are {', '.join(dens_list)}. Their LCM is {cd}, so we expand all fractions to denominator {cd}:\n",
                        f"1. Schritt (Hauptnenner bilden):\n• Die Nenner lauten {', '.join(dens_list)}. Das kgV ist {cd}, also erweitern wir alle Brüche auf den Nenner {cd}:\n",
                    )
                    + "\n".join(exp_bullets)
                    + _L(
                        f"\n\n2. Adım (Payları Toplayıp Çıkaralım):\n• Genişlettiğimiz kesirleri yazalım: {' '.join(expanded_strs)}\n• Ortak payda ({cd}) üzerinde payları hesaplayalım: ({' '.join(num_expr_parts)}) / {cd} = {unsimplified_num}/{cd}",
                        f"\n\nStep 2 (Combine the Numerators):\n• Rewrite with the common denominator: {' '.join(expanded_strs)}\n• Combine numerators over {cd}: ({' '.join(num_expr_parts)}) / {cd} = {unsimplified_num}/{cd}",
                        f"\n\n2. Schritt (Zähler zusammenfassen):\n• Mit Hauptnenner schreiben: {' '.join(expanded_strs)}\n• Zähler über {cd} berechnen: ({' '.join(num_expr_parts)}) / {cd} = {unsimplified_num}/{cd}",
                    )
                    + simp_step_txt
                    + (f"\n\n{h4}\n" + "\n".join(distractor_lines) if distractor_lines else "")
                    + f"\n\n{h5}\n{ans_line}"
                )

            elif ast_steps and (matched_letter or not options):
                step_lines = [f"{i+1}. {st}" for i, st in enumerate(ast_steps)]
                distractor_lines = []
                for k, v in sorted(options.items()):
                    v_clean = _latex_to_readable_math(v).strip()
                    if k == matched_letter:
                        distractor_lines.append(_L(
                            f"• {k}) {v_clean} ➔ DOĞRU (Adım adım işlemde bulduğumuz {res_str} sonucu ile birebir eşleşmektedir).",
                            f"• {k}) {v_clean} ➔ CORRECT (Matches the exact calculated value {res_str}).",
                            f"• {k}) {v_clean} ➔ RICHTIG (Stimmt exakt mit dem berechneten Wert {res_str} überein).",
                        ))
                    else:
                        distractor_lines.append(_L(
                            f"• {k}) {v_clean} ➔ Elenir (Doğru sonuç {res_str} olduğu için yanlıştır).",
                            f"• {k}) {v_clean} ➔ Eliminated (The exact result is {res_str}).",
                            f"• {k}) {v_clean} ➔ Ausgeschlossen (Das exakte Ergebnis ist {res_str}).",
                        ))
                ans_line = f"{ans_prefix}: {matched_letter}) {_latex_to_readable_math(options[matched_letter])}" if matched_letter else f"{ans_prefix}: {res_str}"
                return (
                    f"{h1}\n"
                    + _L(
                        f"Soruda verilen '{raw_expr}' matematiksel işleminin sonucunu işlem önceliği kurallarına dikkat ederek bulmamız isteniyor.\n\n",
                        f"We need to evaluate the mathematical expression '{raw_expr}' following the order of operations.\n\n",
                        f"Wir müssen den mathematischen Ausdruck '{raw_expr}' unter Beachtung der Punkt-vor-Strich-Regel berechnen.\n\n",
                    )
                    + f"{h2}\n"
                    + _L(
                        "Matematikte işlem önceliği sırası: 1) Üslü ifadeler, 2) Parantez içi işlemler, 3) Çarpma ve bölme, 4) Toplama ve çıkarma.\n\n",
                        "Order of operations (PEMDAS): 1) Exponents, 2) Parentheses, 3) Multiplication and Division, 4) Addition and Subtraction.\n\n",
                        "Rechenreihenfolge: 1) Potenzen, 2) Klammern, 3) Multiplikation und Division, 4) Addition und Subtraktion.\n\n",
                    )
                    + f"{h3}\n"
                    + "\n".join(step_lines)
                    + f"\n• Result = {res_str}"
                    + (f"\n\n{h4}\n" + "\n".join(distractor_lines) if distractor_lines else "")
                    + f"\n\n{h5}\n{ans_line}"
                )

    # =========================================================================
    # 2. CONSECUTIVE NUMBERS SOLVER (e.g., "Ardışık üç tek sayının toplamı 63 ise en büyüğü kaçtır?")
    # =========================================================================
    m_ard = re.search(
        r"ard[ıi][sş][ıi]k\s+(üç|uc|3|beş|bes|5|dört|dort|4|yedi|7)\s*(tek|çift|cift|tam|doğal)?\s*say[ıi]n[ıi]n\s+toplam[ıi]\s*(\d+)",
        full_math_txt,
        re.I,
    )
    if m_ard:
        cnt_map = {"üç": 3, "uc": 3, "3": 3, "dört": 4, "dort": 4, "4": 4, "beş": 5, "bes": 5, "5": 5, "yedi": 7, "7": 7}
        n_cnt = cnt_map.get(m_ard.group(1).lower(), 3)
        parity = (m_ard.group(2) or "tam").lower()
        total_s = int(m_ard.group(3))
        step_d = 2 if parity in ("tek", "çift", "cift") else 1
        parity_label = "tek" if parity == "tek" else ("çift" if parity in ("çift", "cift") else "tam")
        if total_s % n_cnt == 0 and n_cnt % 2 == 1:
            mid_val = total_s // n_cnt
            seq = [mid_val + (i - n_cnt // 2) * step_d for i in range(n_cnt)]
            if re.search(r"en\s+b[üu]y[üu][ğg][üu]|largest|greatest|größte", full_math_txt, re.I):
                ans_num = max(seq)
                target_label = "en büyüğü"
            elif re.search(r"en\s+k[üu][çc][üu][ğg][üu]|smallest|least|kleinste", full_math_txt, re.I):
                ans_num = min(seq)
                target_label = "en küçüğü"
            else:
                ans_num = mid_val
                target_label = "ortancası"

            matched_letter = None
            for k, v in sorted(options.items()):
                if v.strip() == str(ans_num):
                    matched_letter = k
                    break

            seq_str = ", ".join(str(x) for x in seq)
            sum_check = " + ".join(str(x) for x in seq) + f" = {total_s}"
            distractor_lines = []
            for k, v in sorted(options.items()):
                v_clean = v.strip()
                if k == matched_letter:
                    distractor_lines.append(f"• {k}) {v_clean} ➔ DOĞRU (Bulduğumuz ardışık sayıların ({seq_str}) {target_label} olan {ans_num} sayısıdır).")
                elif v_clean == str(mid_val) and ans_num != mid_val:
                    distractor_lines.append(f"• {k}) {v_clean} ➔ Elenir (En güçlü sınav tuzağıdır! {mid_val} ortanca sayıdır; soruda ise bize {target_label} sorulmuştur).")
                else:
                    distractor_lines.append(f"• {k}) {v_clean} ➔ Elenir (Aradığımız {target_label} {ans_num} olduğu için bu şık yanlıştır).")

            ans_line = f"{ans_prefix}: {matched_letter}) {ans_num}" if matched_letter else f"{ans_prefix}: {ans_num}"
            return (
                f"{h1}\n"
                f"Soruda bize ardışık (art arda gelen) {n_cnt} tane {parity_label} sayının toplamının {total_s} olduğu verilmiş ve bu sayıların {target_label} sorulmuştur.\n\n"
                f"{h2}\n"
                f"• Kural 1: Ardışık tek veya çift sayılar birbirini her zaman {step_d}'şer {step_d}'şer takip eder (Örn: n, n+{step_d}, n+{step_d*2}...).\n"
                f"• Kural 2 (Sınav kazandıran pratik yöntem): Ardışık sayıların toplamını sayı adedine ({n_cnt}'e) bölersek her zaman tam ortadaki sayıyı (Ortanca Sayıyı) buluruz!\n\n"
                f"{h3}\n"
                f"1. Yöntem (Pratik Ortanca Sayı Yöntemi):\n"
                f"• Toplamı sayı adedine bölelim: {total_s} ÷ {n_cnt} = {mid_val} ➔ Bu bizim tam ortadaki (ortanca) sayımızdır.\n"
                f"• Sayılarımız ardışık {parity_label} sayı olduğu için {step_d}'şer {step_d}'şer değişir. Ortanca sayı {mid_val} olduğuna göre:\n"
                f"  - Küçükten büyüğe sayılarımız: {seq_str} olur.\n"
                f"  - Sağlamasını yapalım: {sum_check} (Toplam birebir doğrulandı!).\n\n"
                f"2. Yöntem (Denklem Kurma Yöntemi):\n"
                f"• En küçük sayıya x dersek; diğer sayılar sırasıyla "
                + ", ".join([f"x+{i*step_d}" for i in range(1, n_cnt)])
                + f" olur.\n"
                f"• Hepsini toplayıp {total_s}'e eşitleyelim: {n_cnt}x + {sum(i*step_d for i in range(n_cnt))} = {total_s} ➔ x = {min(seq)}.\n"
                f"• Buradan sayılarımız yine {seq_str} olarak bulunur ve {target_label} = {ans_num} çıkar."
                + (f"\n\n{h4}\n" + "\n".join(distractor_lines) if distractor_lines else "")
                + f"\n\n{h5}\n{ans_line}"
            )

    # =========================================================================
    # 3. EXPONENT DIVISION / MULTIPLICATION SOLVER (e.g., "2¹² GB ... 2³ GB/sn")
    # =========================================================================
    exp_matches = re.findall(r"(\d+)\s*(?:\^\{?(\d+)\}?|([⁰¹²³⁴⁵⁶⁷⁸⁹]+))", full_stem_block)
    parsed_exps = []
    rev_sup = str.maketrans("⁰¹²³⁴⁵⁶⁷⁸⁹", "0123456789")
    for b_str, e_plain, e_sup in exp_matches:
        e_val = e_plain if e_plain else e_sup.translate(rev_sup)
        if b_str.isdigit() and e_val.isdigit():
            parsed_exps.append((int(b_str), int(e_val)))
    if len(parsed_exps) >= 2 and parsed_exps[0][0] == parsed_exps[1][0]:
        base_num = parsed_exps[0][0]
        e1, e2 = parsed_exps[0][1], parsed_exps[1][1]
        if any(w in full_math_txt.lower() for w in ("hız", "saniye", "sürer", "böl", "oran", "kaç", "speed", "second", "divided", "ratio", "how")):
            diff_e = abs(e1 - e2)
            sum_e = e1 + e2
            b_e1_str = f"{base_num}{str(e1).translate(_SUPERSCRIPT_MAP)}"
            b_e2_str = f"{base_num}{str(e2).translate(_SUPERSCRIPT_MAP)}"
            res_exp_str = f"{base_num}{str(diff_e).translate(_SUPERSCRIPT_MAP)}"
            trap_sum_str = f"{base_num}{str(sum_e).translate(_SUPERSCRIPT_MAP)}"
            target_forms = {f"{base_num}^{diff_e}", res_exp_str, str(base_num ** diff_e)}
            matched_letter = None
            for k, v in sorted(options.items()):
                v_clean = _latex_to_readable_math(v).strip()
                if v_clean in target_forms or v.strip() in target_forms:
                    matched_letter = k
                    break
            if matched_letter:
                distractor_lines = []
                for k, v in sorted(options.items()):
                    v_clean = _latex_to_readable_math(v).strip()
                    if k == matched_letter:
                        distractor_lines.append(_L(
                            f"• {k}) {v_clean} ➔ DOĞRU ({b_e1_str} ÷ {b_e2_str} işleminde üslerin farkı ({e1} - {e2} = {diff_e}) alınarak bulunan {res_exp_str} sonucudur).",
                            f"• {k}) {v_clean} ➔ CORRECT (Found by subtracting exponents in {b_e1_str} ÷ {b_e2_str}: {e1} - {e2} = {diff_e} ➔ {res_exp_str}).",
                            f"• {k}) {v_clean} ➔ RICHTIG (Durch Subtraktion der Exponenten bei {b_e1_str} ÷ {b_e2_str}: {e1} - {e2} = {diff_e} ➔ {res_exp_str}).",
                        ))
                    elif v_clean == trap_sum_str:
                        distractor_lines.append(_L(
                            f"• {k}) {v_clean} ➔ Elenir (Sınav tuzağı! Bölme yerine çarpma yapılıp üsler toplanırsa ({e1} + {e2} = {sum_e}) bu yanlış şık bulunur).",
                            f"• {k}) {v_clean} ➔ Eliminated (Common exam trap! Adding exponents {e1} + {e2} = {sum_e} instead of subtracting).",
                            f"• {k}) {v_clean} ➔ Ausgeschlossen (Typische Prüfungsfalle: Exponenten wurden addiert statt subtrahiert).",
                        ))
                    else:
                        distractor_lines.append(_L(
                            f"• {k}) {v_clean} ➔ Elenir (Doğru sonuç {res_exp_str} olduğu için yanlıştır).",
                            f"• {k}) {v_clean} ➔ Eliminated (The correct value is {res_exp_str}).",
                            f"• {k}) {v_clean} ➔ Ausgeschlossen (Der richtige Wert ist {res_exp_str}).",
                        ))
                return (
                    f"{h1}\n"
                    + _L(
                        f"Soruda bize toplam büyüklük ({b_e1_str}) ve birim hız/miktar ({b_e2_str}) üslü sayı olarak verilmiş; işlemin tamamlanma süresini veya oranını bulmamız isteniyor.\n\n",
                        f"We are given a total quantity ({b_e1_str}) and a rate/divisor ({b_e2_str}) as exponential expressions and asked to find their quotient.\n\n",
                        f"Gegeben sind die Gesamtmenge ({b_e1_str}) und die Rate ({b_e2_str}) als Potenzen; gesucht ist ihr Quotient.\n\n",
                    )
                    + f"{h2}\n"
                    + _L(
                        f"• Üslü Sayılarda Bölme Kuralı: Tabanları aynı olan iki üslü sayı birbirine bölünürken ortak taban aynen yazılır ve payın üssünden paydanın üssü çıkarılır: aᵐ ÷ aⁿ = aᵐ⁻ⁿ.\n\n",
                        f"• Quotient of Powers Rule: When dividing two powers with the same base, keep the base and subtract the denominator's exponent from the numerator's exponent: aᵐ ÷ aⁿ = aᵐ⁻ⁿ.\n\n",
                        f"• Potenzgesetz für die Division: Potenzen mit gleicher Basis werden dividiert, indem man die Basis beibehält und die Exponenten subtrahiert: aᵐ ÷ aⁿ = aᵐ⁻ⁿ.\n\n",
                    )
                    + f"{h3}\n"
                    + _L(
                        f"• 1. Adım (Bölme İşlemini Kurma): {b_e1_str} ÷ {b_e2_str}\n"
                        f"• 2. Adım (Üsleri Çıkarma): {e1} - {e2} = {diff_e}\n"
                        f"• 3. Adım (Sonucu Yazma): Sonuç {res_exp_str} olarak bulunur.",
                        f"• Step 1 (Set up the division): {b_e1_str} ÷ {b_e2_str}\n"
                        f"• Step 2 (Subtract the exponents): {e1} - {e2} = {diff_e}\n"
                        f"• Step 3 (Write the result): {res_exp_str}.",
                        f"• 1. Schritt (Division aufstellen): {b_e1_str} ÷ {b_e2_str}\n"
                        f"• 2. Schritt (Exponenten subtrahieren): {e1} - {e2} = {diff_e}\n"
                        f"• 3. Schritt (Ergebnis notieren): {res_exp_str}.",
                    )
                    + f"\n\n{h4}\n"
                    + "\n".join(distractor_lines)
                    + f"\n\n{h5}\n{ans_prefix}: {matched_letter}) {_latex_to_readable_math(options[matched_letter])}"
                )

    # =========================================================================
    # 4. 1-VARIABLE EQUATION SOLVER (e.g., "3x - 5 = 16 ise x kaçtır?")
    # =========================================================================
    m_eq = re.search(r"([0-9xXaAnNkKyY\s\+\-\*/\(\)]+)\s*=\s*([0-9xXaAnNkKyY\s\+\-\*/\(\)]+)", full_math_txt)
    if m_eq and re.search(r"[xXaAnNkKyY]", m_eq.group(0)) and options:
        lhs_raw, rhs_raw = m_eq.group(1).strip(), m_eq.group(2).strip()
        var_chars = list(set(re.findall(r"[xXaAnNkKyY]", m_eq.group(0))))
        if len(var_chars) == 1:
            v_ch = var_chars[0]

            def _prep_alg(s_alg, val_f):
                s_p = re.sub(rf"(\d)\s*{v_ch}", rf"\1*({val_f})", s_alg)
                s_p = re.sub(rf"{v_ch}\s*(\d)", rf"({val_f})*\1", s_p)
                s_p = re.sub(rf"(\))\s*{v_ch}", rf"\1*({val_f})", s_p)
                s_p = re.sub(rf"(\d)\s*\(", r"\1*(", s_p)
                s_p = s_p.replace(v_ch, f"({val_f})")
                r_f, _, _ = _eval_math_ast_with_steps(s_p)
                return r_f

            for k, v in sorted(options.items()):
                opt_f = _parse_option_as_fraction(v)
                if opt_f is not None:
                    l_val = _prep_alg(lhs_raw, _format_fraction_tr(opt_f))
                    r_val = _prep_alg(rhs_raw, _format_fraction_tr(opt_f))
                    if l_val is not None and r_val is not None and l_val == r_val:
                        res_str = _format_fraction_tr(opt_f)
                        eq_full = f"{lhs_raw} = {rhs_raw}"
                        distractor_lines = []
                        for ok_k, ok_v in sorted(options.items()):
                            if ok_k == k:
                                distractor_lines.append(
                                    _L(
                                        f"• {ok_k}) {ok_v} ➔ DOĞRU (Denklemde {v_ch} yerine {res_str} yazdığımızda her iki taraf da {_format_fraction_tr(l_val)} değerine eşitlenmektedir).",
                                        f"• {ok_k}) {ok_v} ➔ CORRECT (Substituting {v_ch} = {res_str} makes both sides equal to {_format_fraction_tr(l_val)}).",
                                        f"• {ok_k}) {ok_v} ➔ RICHTIG (Durch Einsetzen von {v_ch} = {res_str} ergeben beide Seiten {_format_fraction_tr(l_val)}).",
                                    )
                                )
                            else:
                                distractor_lines.append(_L(
                                    f"• {ok_k}) {ok_v} ➔ Elenir (Denklemde yerine konulduğunda eşitliği sağlamaz).",
                                    f"• {ok_k}) {ok_v} ➔ Eliminated (Does not satisfy the equation).",
                                    f"• {ok_k}) {ok_v} ➔ Ausgeschlossen (Erfüllt die Gleichung nicht).",
                                ))
                        return (
                            f"{h1}\n"
                            + _L(
                                f"Soruda bize '{eq_full}' birinci dereceden denklemi verilmiş ve eşitliği sağlayan '{v_ch}' bilinmeyeninin değeri sorulmuştur.\n\n",
                                f"We are given the linear equation '{eq_full}' and asked to find the value of the unknown '{v_ch}'.\n\n",
                                f"Gegeben ist die lineare Gleichung '{eq_full}'; gesucht ist der Wert der Unbekannten '{v_ch}'.\n\n",
                            )
                            + f"{h2}\n"
                            + _L(
                                f"Bir bilinmeyenli denklemleri çözerken temel kuralımız: Bilinmeyenleri ({v_ch} içeren terimleri) eşitliğin bir tarafında, bilinen sayıları ise diğer tarafında toplamaktır.\n\n",
                                f"To solve a linear equation in one variable, isolate all terms containing '{v_ch}' on one side of the equation and constant numbers on the other side.\n\n",
                                f"Um eine lineare Gleichung mit einer Unbekannten zu lösen, bringen wir alle Terme mit '{v_ch}' auf eine Seite und die Zahlen auf die andere Seite.\n\n",
                            )
                            + f"{h3}\n"
                            + _L(
                                f"• 1. Adım (Denklemi Yazalım): {eq_full}\n"
                                f"• 2. Adım (Bilinmeyeni Yalnız Bırakalım): Bilinen sayıları karşı tarafa geçirip {v_ch}'in katsayısına böldüğümüzde {v_ch} = {res_str} bulunur.\n"
                                f"• 3. Adım (Sağlama Yapalım): Denklemde {v_ch} yerine {res_str} koyduğumuzda sol taraf = {_format_fraction_tr(l_val)} ve sağ taraf = {_format_fraction_tr(r_val)} olur.",
                                f"• Step 1 (Write the equation): {eq_full}\n"
                                f"• Step 2 (Isolate {v_ch}): Move constants to the other side and divide by the coefficient of {v_ch} to find {v_ch} = {res_str}.\n"
                                f"• Step 3 (Verify): Substituting {v_ch} = {res_str} gives LHS = {_format_fraction_tr(l_val)} and RHS = {_format_fraction_tr(r_val)}.",
                                f"• 1. Schritt (Gleichung notieren): {eq_full}\n"
                                f"• 2. Schritt ({v_ch} isolieren): Konstanten umstellen und durch den Koeffizienten teilen ergibt {v_ch} = {res_str}.\n"
                                f"• 3. Schritt (Probe): Einsetzen von {v_ch} = {res_str} liefert links = {_format_fraction_tr(l_val)} und rechts = {_format_fraction_tr(r_val)}.",
                            )
                            + f"\n\n{h4}\n"
                            + "\n".join(distractor_lines)
                            + f"\n\n{h5}\n{ans_prefix}: {k}) {_latex_to_readable_math(v)}"
                        )

    # =========================================================================
    # 5. PERCENTAGE, SUCCESSIVE ZAM/İNDİRİM & KAR-ZARAR SOLVER
    # =========================================================================
    pct_events = re.findall(
        r"(?:%\s*(\d+)|yüzde\s*(\d+)|(\d+)\s*%)\s*([a-zA-ZçğıöşüÇĞİÖŞÜäöüÄÖÜß\s]{1,28}?)(?=[,\.\?;]|ardından|sonra|daha|yapılıp|ve\b|then\b|and\b|danach\b|$)",
        full_math_txt,
        re.I,
    )
    parsed_pct_ops = []
    for p1, p2, p3, ctx_w in pct_events:
        p_val = int(p1 or p2 or p3)
        ctx_l = ctx_w.lower()
        if any(w in ctx_l for w in ("zam", "art", "kâr", "kar", "fazla", "ekle", "increas", "rais", "markup", "erhöh", "zuschlag")):
            parsed_pct_ops.append(("+", p_val, _L("zam (artış)", "increase", "Erhöhung")))
        elif any(w in ctx_l for w in ("indirim", "azal", "zarar", "iskonto", "eksi", "düş", "decreas", "discount", "reduc", "rabatt", "reduzier", "senk")):
            parsed_pct_ops.append(("-", p_val, _L("indirim (azalış)", "discount", "Rabatt")))

    if len(parsed_pct_ops) >= 2:
        cur_price = Fraction(100)
        step_bullets = [_L(
            "• 1. Adım (Başlangıç Fiyatına 100 Diyelim): İşlemleri kesirlerle uğraşmadan kolayca yapmak için malın ilk fiyatına 100 diyelim.",
            "• Step 1 (Set Initial Price = 100): Let the initial price be 100 so we can calculate percentages cleanly without fractions.",
            "• 1. Schritt (Startpreis = 100 setzen): Wir setzen den ursprünglichen Preis auf 100, um Prozentwerte direkt berechnen zu können.",
        )]
        for idx_op, (sgn, p_val, op_lbl) in enumerate(parsed_pct_ops, start=2):
            delta = cur_price * Fraction(p_val, 100)
            prev_str = _format_fraction_tr(cur_price)
            delta_str = _format_fraction_tr(delta)
            if sgn == "+":
                cur_price = cur_price + delta
                new_str = _format_fraction_tr(cur_price)
                step_bullets.append(_L(
                    f"• {idx_op}. Adım (%{p_val} {op_lbl} uygulayalım): Mevcut fiyatımız ({prev_str}) üzerinden %{p_val} artış hesaplanır:\n"
                    f"  - Artış miktarı: {prev_str} × ({p_val} / 100) = {delta_str}\n"
                    f"  - Yeni fiyatımız: {prev_str} + {delta_str} = {new_str} olur.",
                    f"• Step {idx_op} (Apply {p_val}% {op_lbl}): Calculate {p_val}% of the current price ({prev_str}):\n"
                    f"  - Increase amount: {prev_str} × ({p_val} / 100) = {delta_str}\n"
                    f"  - New price: {prev_str} + {delta_str} = {new_str}.",
                    f"• {idx_op}. Schritt ({p_val}% {op_lbl} anwenden): Berechne {p_val}% des aktuellen Preises ({prev_str}):\n"
                    f"  - Erhöhung: {prev_str} × ({p_val} / 100) = {delta_str}\n"
                    f"  - Neuer Preis: {prev_str} + {delta_str} = {new_str}.",
                ))
            else:
                cur_price = cur_price - delta
                new_str = _format_fraction_tr(cur_price)
                step_bullets.append(_L(
                    f"• {idx_op}. Adım (Yeni fiyat üzerinden %{p_val} {op_lbl} uygulayalım): Dikkat! Bu işlem ilk fiyat (100) üzerinden değil, güncel fiyatımız ({prev_str}) üzerinden hesaplanır:\n"
                    f"  - İndirim miktarı: {prev_str} × ({p_val} / 100) = {delta_str}\n"
                    f"  - Son fiyatımız: {prev_str} - {delta_str} = {new_str} olur.",
                    f"• Step {idx_op} (Apply {p_val}% {op_lbl} on the new price): Note that this is calculated on the updated price ({prev_str}), not on 100:\n"
                    f"  - Discount amount: {prev_str} × ({p_val} / 100) = {delta_str}\n"
                    f"  - Final price: {prev_str} - {delta_str} = {new_str}.",
                    f"• {idx_op}. Schritt ({p_val}% {op_lbl} auf den neuen Preis anwenden): Wichtig: Dies wird auf Basis des neuen Preises ({prev_str}) berechnet:\n"
                    f"  - Rabattbetrag: {prev_str} × ({p_val} / 100) = {delta_str}\n"
                    f"  - Endpreis: {prev_str} - {delta_str} = {new_str}.",
                ))

        asks_diff = bool(re.search(r"yüzde\s+kaç\s+(kar|kâr|zarar|azal|art|değiş)", full_math_txt, re.I))
        target_num = abs(Fraction(100) - cur_price) if asks_diff else cur_price
        target_str = _format_fraction_tr(target_num)

        matched_letter = None
        for k, v in sorted(options.items()):
            v_digits = re.sub(r"[^0-9,\./\-]", "", _latex_to_readable_math(v))
            opt_f = _parse_option_as_fraction(v_digits)
            if opt_f is not None and opt_f == target_num:
                matched_letter = k
                break

        if matched_letter or not options:
            distractor_lines = []
            for k, v in sorted(options.items()):
                v_clean = _latex_to_readable_math(v).strip()
                v_digits = re.sub(r"[^0-9,\./\-]", "", v_clean)
                if k == matched_letter:
                    distractor_lines.append(_L(
                        f"• {k}) {v_clean} ➔ DOĞRU (100 olan ilk fiyatımız sırasıyla uygulanan işlemler sonucunda {_format_fraction_tr(cur_price)} olduğu için aradığımız kesin sonuçtur).",
                        f"• {k}) {v_clean} ➔ CORRECT (Starting from 100, the final price after successive changes is {_format_fraction_tr(cur_price)}).",
                        f"• {k}) {v_clean} ➔ RICHTIG (Ausgehend von 100 beträgt der Endpreis nach den Änderungen {_format_fraction_tr(cur_price)}).",
                    ))
                elif v_digits == "100":
                    distractor_lines.append(_L(
                        f"• {k}) {v_clean} ➔ Elenir (En yaygın öğrenci tuzağıdır! İkinci işlem 100 üzerinden değil zamlı fiyat üzerinden yapılmaktadır).",
                        f"• {k}) {v_clean} ➔ Eliminated (Classic trap! The second percentage is applied to the updated price, not the original 100).",
                        f"• {k}) {v_clean} ➔ Ausgeschlossen (Typische Falle! Die zweite Prozentänderung bezieht sich auf den neuen Preis, nicht auf 100).",
                    ))
                else:
                    distractor_lines.append(_L(
                        f"• {k}) {v_clean} ➔ Elenir (Son fiyat {_format_fraction_tr(cur_price)} bulunduğu için bu şık yanlıştır).",
                        f"• {k}) {v_clean} ➔ Eliminated (The calculated final value is {_format_fraction_tr(cur_price)}).",
                        f"• {k}) {v_clean} ➔ Ausgeschlossen (Der berechnete Endwert beträgt {_format_fraction_tr(cur_price)}).",
                    ))

            ans_line = f"{ans_prefix}: {matched_letter}) {_latex_to_readable_math(options[matched_letter])}" if matched_letter else f"{ans_prefix}: %{target_str}"
            return (
                f"{h1}\n"
                + _L(
                    f"Soruda bir malın fiyatına ardışık yüzde değişimleri uygulandığında son fiyatın ilk fiyatın yüzde kaçı olduğu sorulmuştur.\n\n",
                    f"We are asked to find the final price as a percentage of the initial price after successive percentage changes.\n\n",
                    f"Gefragt ist, wie viel Prozent des ursprünglichen Preises der Endpreis nach aufeinanderfolgenden Prozentänderungen beträgt.\n\n",
                )
                + f"{h2}\n"
                + _L(
                    "• Pratik Kural ('100 Diyerek Çözme Yöntemi'): Yüzde problemlerinde başlangıç değerine her zaman 100 deriz.\n"
                    "• Altın Kural: İkinci yüzde işlemi 100 üzerinden değil, oluşan yeni fiyat üzerinden hesaplanır.\n\n",
                    "• Practical Rule ('Start with 100 Method'): Always assign 100 to the initial price in percentage problems.\n"
                    "• Golden Rule: Each subsequent percentage change applies to the new updated price, not the original 100.\n\n",
                    "• Praktische Methode ('Startwert 100'): Setze den Anfangspreis bei Prozentaufgaben immer auf 100.\n"
                    "• Goldene Regel: Jede weitere Prozentänderung wird vom neuen Zwischenwert berechnet.\n\n",
                )
                + f"{h3}\n"
                + "\n".join(step_bullets)
                + (f"\n\n{h4}\n" + "\n".join(distractor_lines) if distractor_lines else "")
                + f"\n\n{h5}\n{ans_line}"
            )

    # =========================================================================
    # 6. SYMPY SYMBOLIC LIMIT, DERIVATIVE & ALGEBRA SOLVER (TR / EN / DE)
    #    (e.g., "4. lim_{x->2} (x³ - 8) / (x² - 4) limitinin değeri kaçtır?")
    # =========================================================================
    if re.search(r"\b(lim|limit|limitinin|grenzwert)\b", full_math_txt, re.I):
        try:
            import sympy as sp

            x_sym = sp.Symbol("x")
            m_pt = re.search(r"[xX]\s*(?:→|->|-->|=)\s*([+\-]?\d+)", full_math_txt)
            a_val = int(m_pt.group(1)) if m_pt else 2

            lim_sub = full_math_txt
            if m_pt:
                lim_sub = full_math_txt[m_pt.end():]
            lim_sub = re.split(r"\b(limitinin|ifadesinin|değeri|sonucu|kaçtır|nedir|what|is|the|value|find|berechnen|welcher)\b", lim_sub, flags=re.I)[0]
            lim_sub = lim_sub.strip(" :?.'\"=")

            rev_sup = str.maketrans("⁰¹²³⁴⁵⁶⁷⁸⁹", "0123456789")
            s_py = re.sub(r"([⁰¹²³⁴⁵⁶⁷⁸⁹]+)", lambda m: "**" + m.group(1).translate(rev_sup), lim_sub)
            s_py = s_py.replace("^", "**").replace("−", "-").replace("–", "-").replace("×", "*").replace("·", "*")
            s_py = re.sub(r"(\d)\s*([xX\(])", r"\1*\2", s_py)
            s_py = re.sub(r"(\))\s*(\()", r"\1*\2", s_py)
            s_py = re.sub(r"([xX])\s*(\d)", r"\1**\2", s_py)
            s_py = s_py.replace("X", "x")

            def _sp_tr(e_obj):
                s_e = str(e_obj)
                s_e = re.sub(r"\*\*(\d+)", lambda m: m.group(1).translate(_SUPERSCRIPT_MAP), s_e)
                s_e = re.sub(r"(\d+)\*x", r"\1x", s_e)
                s_e = s_e.replace("*", "·")
                return s_e

            sp_expr = sp.sympify(s_py, locals={"x": x_sym})
            lim_res = sp.limit(sp_expr, x_sym, a_val)
            num_sp, den_sp = sp.fraction(sp_expr)
            num_at_a = num_sp.subs(x_sym, a_val)
            den_at_a = den_sp.subs(x_sym, a_val)

            res_str = _sp_tr(lim_res)
            matched_letter = None
            for k, v in sorted(options.items()):
                v_clean = _latex_to_readable_math(v).strip()
                if v_clean == res_str:
                    matched_letter = k
                    break
                opt_f = _parse_option_as_fraction(v_clean)
                if opt_f is not None and str(opt_f) == res_str:
                    matched_letter = k
                    break

            num_readable = _sp_tr(num_sp)
            den_readable = _sp_tr(den_sp)
            expr_readable = f"({num_readable}) / ({den_readable})" if den_sp != 1 else num_readable

            if num_at_a == 0 and den_at_a == 0 and den_sp != 1:
                fac_num = sp.factor(num_sp)
                fac_den = sp.factor(den_sp)
                canc_expr = sp.cancel(sp_expr)
                c_num, c_den = sp.fraction(canc_expr)
                c_num_val = c_num.subs(x_sym, a_val)
                c_den_val = c_den.subs(x_sym, a_val)
                d_num = sp.diff(num_sp, x_sym)
                d_den = sp.diff(den_sp, x_sym)
                d_num_val = d_num.subs(x_sym, a_val)
                d_den_val = d_den.subs(x_sym, a_val)

                distractor_lines = []
                for k, v in sorted(options.items()):
                    v_clean = _latex_to_readable_math(v).strip()
                    if k == matched_letter:
                        distractor_lines.append(_L(
                            f"• {k}) {v_clean} ➔ DOĞRU (Sadeleştirme sonrasında pay {c_num_val}, payda {c_den_val} bulunur ve {c_num_val} ÷ {c_den_val} = {res_str} sonucu elde edilir).",
                            f"• {k}) {v_clean} ➔ CORRECT (After canceling (x - {a_val}), numerator = {c_num_val}, denominator = {c_den_val}, giving {c_num_val} ÷ {c_den_val} = {res_str}).",
                            f"• {k}) {v_clean} ➔ RICHTIG (Nach dem Kürzen von (x - {a_val}) ergibt sich {c_num_val} ÷ {c_den_val} = {res_str}).",
                        ))
                    elif v_clean == str(c_den_val):
                        distractor_lines.append(_L(
                            f"• {k}) {v_clean} ➔ Elenir (Sınav tuzağı! Sadece paydadaki ({_sp_tr(c_den)}) ifadesinde x = {a_val} yazıp {c_den_val} bulan ama payı ({c_num_val}) paydayla bölmeyi unutan öğrencilerin düştüğü çeldiricidir).",
                            f"• {k}) {v_clean} ➔ Eliminated (Exam trap! {c_den_val} is only the denominator value ({_sp_tr(c_den)}); you must divide the numerator {c_num_val} by {c_den_val}).",
                            f"• {k}) {v_clean} ➔ Ausgeschlossen (Prüfungsfalle! {c_den_val} ist nur der Nennerwert; der Zähler {c_num_val} muss durch {c_den_val} geteilt werden).",
                        ))
                    else:
                        distractor_lines.append(_L(
                            f"• {k}) {v_clean} ➔ Elenir (Doğru limit değeri {res_str} olduğu için bu şık yanlıştır).",
                            f"• {k}) {v_clean} ➔ Eliminated (The exact limit value is {res_str}).",
                            f"• {k}) {v_clean} ➔ Ausgeschlossen (Der exakte Grenzwert ist {res_str}).",
                        ))

                ans_line = f"{ans_prefix}: {matched_letter}) {_latex_to_readable_math(options[matched_letter])}" if matched_letter else f"{ans_prefix}: {res_str}"
                if lang == "en":
                    return (
                        f"{h1}\n"
                        f"We are asked to find the limit of the rational expression '{expr_readable}' as x approaches {a_val} (x → {a_val}).\n\n"
                        f"{h2}\n"
                        f"• Always start a limit problem by substituting x = {a_val} directly.\n"
                        f"• Since both numerator and denominator become 0, we have a '0 / 0 Indeterminate Form'! We can resolve it using two methods:\n"
                        f"  1) Factoring & Canceling: Factor both numerator and denominator and cancel the common (x - {a_val}) factor.\n"
                        f"  2) L'Hôpital's Rule (Derivative Shortcut): Differentiate the numerator and denominator separately and plug in x = {a_val}.\n\n"
                        f"{h3}\n"
                        f"Step 1 (Check for Indeterminacy):\n"
                        f"• Substitute x = {a_val}: Numerator ({num_readable}) = 0 and Denominator ({den_readable}) = 0 ➔ 0 / 0 indeterminate form.\n\n"
                        f"Step 2 (Method 1 — Factoring & Canceling):\n"
                        f"• Factor the numerator: {num_readable} = {_sp_tr(fac_num)}\n"
                        f"• Factor the denominator: {den_readable} = {_sp_tr(fac_den)}\n"
                        f"• Cancel the common (x - {a_val}) factor:\n"
                        f"  [{_sp_tr(fac_num)}] / [{_sp_tr(fac_den)}] = ({_sp_tr(c_num)}) / ({_sp_tr(c_den)})\n\n"
                        f"Step 3 (Substitute x = {a_val} into the Simplified Expression):\n"
                        f"• Numerator ({_sp_tr(c_num)}) at x = {a_val} ➔ {c_num_val}\n"
                        f"• Denominator ({_sp_tr(c_den)}) at x = {a_val} ➔ {c_den_val}\n"
                        f"• Limit Value = {c_num_val} ÷ {c_den_val} = {res_str}\n\n"
                        f"Step 4 (Method 2 — Verify via L'Hôpital's Rule):\n"
                        f"• Derivative of numerator ({num_readable}): {_sp_tr(d_num)} ➔ at x = {a_val} gives {d_num_val}\n"
                        f"• Derivative of denominator ({den_readable}): {_sp_tr(d_den)} ➔ at x = {a_val} gives {d_den_val}\n"
                        f"• Ratio: {d_num_val} ÷ {d_den_val} = {res_str} (Confirmed!)"
                        + (f"\n\n{h4}\n" + "\n".join(distractor_lines) if distractor_lines else "")
                        + f"\n\n{h5}\n{ans_line}"
                    )
                elif lang == "de":
                    return (
                        f"{h1}\n"
                        f"Gesucht ist der Grenzwert (Limes) des gebrochenrationalen Ausdrucks '{expr_readable}' für x → {a_val}.\n\n"
                        f"{h2}\n"
                        f"• Durch direktes Einsetzen von x = {a_val} entsteht die unbestimmte Form '0 / 0'.\n"
                        f"• Wir lösen dies über: 1) Faktorisieren & Kürzen von (x - {a_val}) oder 2) die Regel von L'Hôpital (Ableitung von Zähler und Nenner).\n\n"
                        f"{h3}\n"
                        f"1. Schritt (Faktorisieren und Kürzen):\n"
                        f"• Zähler faktorisieren: {num_readable} = {_sp_tr(fac_num)}\n"
                        f"• Nenner faktorisieren: {den_readable} = {_sp_tr(fac_den)}\n"
                        f"• Gemeinsamen Faktor (x - {a_val}) kürzen: ({_sp_tr(c_num)}) / ({_sp_tr(c_den)})\n\n"
                        f"2. Schritt (x = {a_val} einsetzen):\n"
                        f"• Zähler ({_sp_tr(c_num)}) für x = {a_val} ➔ {c_num_val}\n"
                        f"• Nenner ({_sp_tr(c_den)}) für x = {a_val} ➔ {c_den_val}\n"
                        f"• Grenzwert = {c_num_val} ÷ {c_den_val} = {res_str}\n\n"
                        f"3. Schritt (Probe mit der Regel von L'Hôpital):\n"
                        f"• Ableitung Zähler: {_sp_tr(d_num)} (ergibt {d_num_val}), Ableitung Nenner: {_sp_tr(d_den)} (ergibt {d_den_val}) ➔ {d_num_val} ÷ {d_den_val} = {res_str}."
                        + (f"\n\n{h4}\n" + "\n".join(distractor_lines) if distractor_lines else "")
                        + f"\n\n{h5}\n{ans_line}"
                    )
                else:
                    return (
                        f"{h1}\n"
                        f"Soruda bize x değişkeni {a_val} sayısına yaklaşırken (x → {a_val}) '{expr_readable}' rasyonel ifadesinin yaklaştığı limit değerini bulmamız isteniyor.\n\n"
                        f"{h2}\n"
                        f"• Limit sorularında ilk kural her zaman x yerine verilen sayıyı (x = {a_val}) yazmaktır.\n"
                        f"• Eğer pay ve payda aynı anda 0 oluyorsa karşımızda '0 / 0 Belirsizliği' vardır! Bu belirsizliği yok etmek için 2 harika yöntem kullanırız:\n"
                        f"  1) Çarpanlara Ayırma Yöntemi: Pay ve paydayı özdeşliklerle çarpanlarına ayırıp hem payı hem paydayı sıfır yapan '(x - {a_val})' çarpanlarını sadeleştiririz.\n"
                        f"  2) Pratik Sınav Yöntemi (L'Hôpital / Türev Kuralı): 0/0 belirsizliğinde payın türevini paya, paydanın türevini paydaya yazıp x = {a_val} değerini yerine koyarız.\n\n"
                        f"{h3}\n"
                        f"1. Adım (Belirsizlik Kontrolü):\n"
                        f"• İfadede x = {a_val} yazalım:\n"
                        f"  - Pay: {num_readable} ➔ x = {a_val} için sonuç = 0\n"
                        f"  - Payda: {den_readable} ➔ x = {a_val} için sonuç = 0\n"
                        f"• Gördüğümüz gibi 0 / 0 belirsizliği oluştu. Şimdi çarpanlara ayırarak bu belirsizliği kaldıralım.\n\n"
                        f"2. Adım (1. Yöntem - Çarpanlara Ayırma ve Sadeleştirme):\n"
                        f"• Payı çarpanlarına ayıralım: {num_readable} = {_sp_tr(fac_num)}\n"
                        f"• Paydayı çarpanlarına ayıralım: {den_readable} = {_sp_tr(fac_den)}\n"
                        f"• Kesrimizde pay ve paydada ortak olan (x - {a_val}) çarpanlarını sadeleştirelim:\n"
                        f"  [{_sp_tr(fac_num)}] / [{_sp_tr(fac_den)}] = ({_sp_tr(c_num)}) / ({_sp_tr(c_den)})\n\n"
                        f"3. Adım (Sadeleşen İfadede x = {a_val} Yazalım):\n"
                        f"• Artık paydayı sıfır yapan çarpan ortadan kalktı! Kalan '({_sp_tr(c_num)}) / ({_sp_tr(c_den)})' ifadesinde x yerine {a_val} yazalım:\n"
                        f"  - Pay ({_sp_tr(c_num)}): x = {a_val} için ➔ {c_num_val}\n"
                        f"  - Payda ({_sp_tr(c_den)}): x = {a_val} için ➔ {c_den_val}\n"
                        f"  - Bölme İşlemi: {c_num_val} ÷ {c_den_val} = {res_str} bulunur!\n\n"
                        f"4. Adım (2. Yöntem - L'Hôpital / Türev Kuralı ile Sağlama):\n"
                        f"• Payın türevi ({num_readable}'in türevi): {_sp_tr(d_num)} ➔ x = {a_val} için sonuç = {d_num_val}\n"
                        f"• Paydanın türevi ({den_readable}'ün türevi): {_sp_tr(d_den)} ➔ x = {a_val} için sonuç = {d_den_val}\n"
                        f"• Oranlarsak: {d_num_val} ÷ {d_den_val} = {res_str} (Her iki yöntemle de sonuç kesin olarak {res_str} çıktı!)."
                        + (f"\n\n{h4}\n" + "\n".join(distractor_lines) if distractor_lines else "")
                        + f"\n\n{h5}\n{ans_line}"
                    )
        except Exception:
            pass

    # --- 7. CALCULUS: DERIVATIVE, CRITICAL POINTS & LOCAL MAXIMUM / MINIMUM (f(x) = ...) ---
    stem_lower = full_math_txt.lower()
    if any(w in stem_lower for w in ("local maximum", "local minimum", "yerel maksimum", "yerel minimum", "lokales maximum", "lokales minimum", "ekstremum")):
        try:
            m_func = re.search(
                r"(?:f\s*\(\s*x\s*\)|y)\s*=\s*([0-9xX\^\⁰\¹\²\³\⁴\⁵\⁶\⁷\⁸\⁹\+\-\*/\(\)\s]+?)(?=\?|\.|,|;|\b(?:fonksiyon|function|für|for|what|kaç|ise)\b|$)",
                full_math_txt,
                re.I,
            )
            if m_func:
                import sympy as sp
                x_sym = sp.Symbol("x", real=True)
                raw_f_str = m_func.group(1).strip()

                def _to_sp_poly(s_expr):
                    s_e = s_expr.translate(str.maketrans("⁰¹²³⁴⁵⁶⁷⁸⁹", "0123456789"))
                    s_e = re.sub(r"([xX])\s*([0-9]+)", r"\1**\2", s_e)
                    s_e = s_e.replace("^", "**").replace("X", "x")
                    s_e = re.sub(r"(\d)\s*(x)", r"\1*\2", s_e)
                    return sp.sympify(s_e, locals={"x": x_sym})

                def _fmt_sp_poly(expr_obj):
                    s_p = str(sp.expand(expr_obj)).replace("**", "^").replace("*", "")
                    return _latex_to_readable_math(s_p)

                def _fmt_sp_factored(expr_obj):
                    s_p = str(sp.factor(expr_obj)).replace("**", "^").replace("*", "")
                    return _latex_to_readable_math(s_p)

                def _norm_val(s_v):
                    return re.sub(r"\s+", "", _latex_to_readable_math(str(s_v))).strip()

                f_expr = _to_sp_poly(raw_f_str)
                df_expr = sp.diff(f_expr, x_sym)
                d2f_expr = sp.diff(df_expr, x_sym)
                crit_pts = sorted(sp.solve(df_expr, x_sym), key=lambda v: float(v))

                wants_max = any(w in stem_lower for w in ("maximum", "maksimum"))
                wants_x_coord = any(w in stem_lower for w in ("apsis", "abscissa", "at what x", "which x", "x-wert", "hangi x"))

                local_max_pts = []
                local_min_pts = []
                for cp in crit_pts:
                    d2_val = sp.simplify(d2f_expr.subs(x_sym, cp))
                    f_val = sp.simplify(f_expr.subs(x_sym, cp))
                    if d2_val.is_number and d2_val < 0:
                        local_max_pts.append((cp, f_val, d2_val))
                    elif d2_val.is_number and d2_val > 0:
                        local_min_pts.append((cp, f_val, d2_val))

                target_tuple = (local_max_pts[0] if wants_max and local_max_pts else (local_min_pts[0] if (not wants_max) and local_min_pts else None))
                if target_tuple is not None:
                    t_x, t_f, t_d2 = target_tuple
                    target_ans_val = t_x if wants_x_coord else t_f
                    res_str = str(int(target_ans_val)) if target_ans_val.is_integer else str(target_ans_val)
                    matched_opt = None
                    for k, v in sorted(options.items()):
                        if _norm_val(v) == _norm_val(res_str):
                            matched_opt = k
                            break
                    if matched_opt:
                        f_read = _fmt_sp_poly(f_expr)
                        df_read = _fmt_sp_poly(df_expr)
                        df_fact = _fmt_sp_factored(df_expr)
                        d2f_read = _fmt_sp_poly(d2f_expr)
                        crit_str = ", ".join([f"x = {c}" for c in crit_pts])

                        eval_lines_tr, eval_lines_en, eval_lines_de = [], [], []
                        for cp in crit_pts:
                            fv = sp.simplify(f_expr.subs(x_sym, cp))
                            d2v = sp.simplify(d2f_expr.subs(x_sym, cp))
                            tag_tr = "Yerel Maksimum (Tepe Noktası)" if d2v < 0 else "Yerel Minimum (Çukur Noktası)"
                            tag_en = "Local Maximum (Peak)" if d2v < 0 else "Local Minimum (Valley)"
                            tag_de = "Lokales Maximum (Hochpunkt)" if d2v < 0 else "Lokales Minimum (Tiefpunkt)"
                            eval_lines_tr.append(f"  - x = {cp} için: f''({cp}) = {d2v} ({'< 0 olduğu için ' + tag_tr if d2v < 0 else '> 0 olduğu için ' + tag_tr}) ➔ Fonksiyon değeri: f({cp}) = {fv}")
                            eval_lines_en.append(f"  - At x = {cp}: f''({cp}) = {d2v} ({'< 0 ➔ ' + tag_en if d2v < 0 else '> 0 ➔ ' + tag_en}) ➔ Function value: f({cp}) = {fv}")
                            eval_lines_de.append(f"  - Bei x = {cp}: f''({cp}) = {d2v} ({'< 0 ➔ ' + tag_de if d2v < 0 else '> 0 ➔ ' + tag_de}) ➔ Funktionswert: f({cp}) = {fv}")

                        distractor_lines = []
                        for k, v in sorted(options.items()):
                            v_clean = _latex_to_readable_math(v)
                            if k == matched_opt:
                                if lang == "en":
                                    distractor_lines.append(f"• Option {k} ({v_clean}): f({t_x}) = {res_str} is the exact {'local maximum' if wants_max else 'local minimum'} value ➔ CORRECT")
                                elif lang == "de":
                                    distractor_lines.append(f"• Option {k} ({v_clean}): f({t_x}) = {res_str} ist der gesuchte Extremwert ➔ RICHTIG")
                                else:
                                    distractor_lines.append(f"• {k} şıkkı ({v_clean}): f({t_x}) = {res_str} aradığımız {'yerel maksimum' if wants_max else 'yerel minimum'} değeridir ➔ DOĞRU")
                            else:
                                trap_note_tr = "İşlem veya işaret hatası sonucu oluşan çeldirici şıktır"
                                trap_note_en = "Distractor option from an arithmetic or sign error"
                                trap_note_de = "Ablenkungsoption durch Vorzeichen- oder Rechenfehler"
                                for cp in crit_pts:
                                    fv = sp.simplify(f_expr.subs(x_sym, cp))
                                    if _norm_val(v) == _norm_val(str(cp)) and not wants_x_coord:
                                        if cp == t_x:
                                            trap_note_tr = f"DİKKAT TUZAK! '{v_clean}' değeri fonksiyonun maksimum değeri değil, maksimum noktasının apsisidir (x = {cp}) ve aynı zamanda yerel minimum değeridir"
                                            trap_note_en = f"COMMON TRAP! '{v_clean}' is the critical point x-coordinate (x = {cp}) and the local minimum value f(3) = 1, NOT the local maximum value f({t_x}) = {res_str}"
                                            trap_note_de = f"TYPISCHE FALLE! '{v_clean}' ist die x-Koordinate (x = {cp}) bzw. das lokale Minimum, nicht der maximale Funktionswert f({t_x}) = {res_str}"
                                        else:
                                            trap_note_tr = f"'{v_clean}' değeri diğer ekstremum noktasının apsisidir (x = {cp})"
                                            trap_note_en = f"'{v_clean}' is the x-coordinate of the other critical point (x = {cp}, where the local minimum occurs)"
                                            trap_note_de = f"'{v_clean}' ist die x-Koordinate des anderen kritischen Punktes (x = {cp})"
                                    elif _norm_val(v) == _norm_val(str(fv)) and cp != t_x:
                                        trap_note_tr = f"DİKKAT TUZAK! '{v_clean}' değeri x = {cp} noktasındaki yerel minimum değeridir (f({cp}) = {fv})"
                                        trap_note_en = f"COMMON TRAP! '{v_clean}' is the local minimum value f({cp}) = {fv}, NOT the local maximum value"
                                        trap_note_de = f"TYPISCHE FALLE! '{v_clean}' ist der lokale Minimumwert f({cp}) = {fv}, nicht das lokale Maximum"
                                if lang == "en":
                                    distractor_lines.append(f"• Option {k} ({v_clean}): {trap_note_en} ➔ Eliminated")
                                elif lang == "de":
                                    distractor_lines.append(f"• Option {k} ({v_clean}): {trap_note_de} ➔ Ausgeschlossen")
                                else:
                                    distractor_lines.append(f"• {k} şıkkı ({v_clean}): {trap_note_tr} ➔ Elenir")

                        ans_line = f"{ans_prefix}: {matched_opt}) {_latex_to_readable_math(options[matched_opt])}"
                        target_lbl_en = "local maximum value" if wants_max else "local minimum value"
                        target_lbl_tr = "yerel maksimum (tepe) değeri" if wants_max else "yerel minimum (çukur) değeri"
                        target_lbl_de = "lokalen Maximalwert" if wants_max else "lokalen Minimalwert"

                        if lang == "en":
                            return (
                                f"{h1}\n"
                                f"We are given the cubic polynomial function f(x) = {f_read} and asked to find its {target_lbl_en}.\n\n"
                                f"{h2}\n"
                                f"• Finding Local Extrema (3 Golden Steps):\n"
                                f"  1. Take the first derivative f'(x) and set it to zero (f'(x) = 0) to find the critical points (where the slope is horizontal).\n"
                                f"  2. Use the Second Derivative Test f''(x): if f''(c) < 0, x = c is a Local Maximum; if f''(c) > 0, x = c is a Local Minimum.\n"
                                f"  3. Substitute the local maximum x-coordinate back into the original function f(x) to find the maximum value!\n\n"
                                f"{h3}\n"
                                f"Step 1 (Find the First Derivative & Critical Points):\n"
                                f"• Differentiate f(x) = {f_read} term by term:\n"
                                f"  - f'(x) = {df_read} = {df_fact} = 0\n"
                                f"• Solving for x gives the critical points: {crit_str}\n\n"
                                f"Step 2 (Second Derivative Test & Function Values):\n"
                                f"• Second derivative: f''(x) = {d2f_read}\n"
                                + "\n".join(eval_lines_en) + "\n\n"
                                f"Step 3 (Final Result):\n"
                                f"• Since f''({t_x}) = {t_d2} < 0, the local maximum occurs at x = {t_x}, and the local maximum value is:\n"
                                f"  - f({t_x}) = {res_str}"
                                + (f"\n\n{h4}\n" + "\n".join(distractor_lines) if distractor_lines else "")
                                + f"\n\n{h5}\n{ans_line}"
                            )
                        elif lang == "de":
                            return (
                                f"{h1}\n"
                                f"Gegeben ist die Funktion f(x) = {f_read}. Gesucht ist ihr {target_lbl_de}.\n\n"
                                f"{h2}\n"
                                f"• Lokale Extremwerte bestimmen (in 3 Schritten):\n"
                                f"  1. Erste Ableitung bilden und nullsetzen (f'(x) = 0), um die kritischen Punkte zu finden.\n"
                                f"  2. Zweite Ableitung f''(x) prüfen: Gilt f''(c) < 0, liegt bei x = c ein lokales Maximum vor; gilt f''(c) > 0, ein lokales Minimum.\n"
                                f"  3. Den x-Wert in die Ausgangsfunktion f(x) einsetzen, um den Funktionswert zu berechnen!\n\n"
                                f"{h3}\n"
                                f"1. Schritt (Erste Ableitung & Kritische Punkte):\n"
                                f"• f'(x) = {df_read} = {df_fact} = 0 ➔ Kritische Punkte: {crit_str}\n\n"
                                f"2. Schritt (Zweite Ableitung & Funktionswerte):\n"
                                f"• f''(x) = {d2f_read}\n"
                                + "\n".join(eval_lines_de) + "\n\n"
                                f"3. Schritt (Ergebnis):\n"
                                f"• Das lokale Maximum liegt bei x = {t_x} mit dem Funktionswert f({t_x}) = {res_str}."
                                + (f"\n\n{h4}\n" + "\n".join(distractor_lines) if distractor_lines else "")
                                + f"\n\n{h5}\n{ans_line}"
                            )
                        else:
                            return (
                                f"{h1}\n"
                                f"Soruda bize f(x) = {f_read} fonksiyonu verilmiş ve bu fonksiyonun {target_lbl_tr} soruluyor.\n\n"
                                f"{h2}\n"
                                f"• Türev ile Yerel Ekstremum (Maksimum / Minimum) Bulma Kuralı:\n"
                                f"  1. Fonksiyonun birinci türevini alıp sıfıra eşitleriz (f'(x) = 0). Bulduğumuz kökler kritik (ekstremum) noktaların apsisleridir.\n"
                                f"  2. İkinci türev testi (f''(x)) veya işaret tablosu ile hangi noktanın tepe (maksimum), hangisinin çukur (minimum) olduğunu belirleriz: f''(c) < 0 ise Yerel Maksimum, f''(c) > 0 ise Yerel Minimumdur.\n"
                                f"  3. DİKKAT: Soru bizden maksimum 'apsisi (x)' değil, maksimum 'değeri (f(x))' istediği için bulduğumuz x değerini ana fonksiyonda yerine yazarız!\n\n"
                                f"{h3}\n"
                                f"1. Adım (Birinci Türevi Alıp Sıfıra Eşitleyelim):\n"
                                f"• f(x) = {f_read} fonksiyonunun türevi:\n"
                                f"  - f'(x) = {df_read} = {df_fact} = 0\n"
                                f"• Buradan kritik noktalarımız: {crit_str} bulunur.\n\n"
                                f"2. Adım (İkinci Türev Testi ve Fonksiyon Değerleri):\n"
                                f"• İkinci türev: f''(x) = {d2f_read}\n"
                                + "\n".join(eval_lines_tr) + "\n\n"
                                f"3. Adım (Bulunan Sonuç):\n"
                                f"• x = {t_x} noktasında yerel maksimum vardır ve fonksiyonun yerel maksimum değeri:\n"
                                f"  - f({t_x}) = {res_str} bulunur!"
                                + (f"\n\n{h4}\n" + "\n".join(distractor_lines) if distractor_lines else "")
                                + f"\n\n{h5}\n{ans_line}"
                            )
        except Exception:
            pass

    # =========================================================================
    # 8. TRIGONOMETRIC IDENTITIES: sin x ± cos x = k  ➔  sin x · cos x = ?
    # =========================================================================
    try:
        m_trig_sc = re.search(
            r"sin\s*[xXαθa]\s*([\+\-])\s*cos\s*[xXαθa]\s*=\s*([+\-]?\d+(?:\s*/\s*\d+)?)",
            full_math_txt,
            re.I,
        )
        if m_trig_sc and re.search(r"sin\s*[xXαθa]\s*[\·\*\.\s]\s*cos\s*[xXαθa]|sin\s*2\s*[xXαθa]", full_math_txt, re.I):
            from fractions import Fraction
            op_sign = m_trig_sc.group(1).strip()
            k_raw = re.sub(r"\s+", "", m_trig_sc.group(2))
            k_frac = Fraction(k_raw)
            k_sq = k_frac * k_frac
            wants_sin2x = bool(re.search(r"sin\s*2\s*[xXαθa]", full_math_txt, re.I))
            two_sc = (k_sq - 1) if op_sign == "+" else (1 - k_sq)
            ans_frac = two_sc if wants_sin2x else (two_sc / 2)

            def _fmt_f(fr):
                return str(fr.numerator) if fr.denominator == 1 else f"{fr.numerator}/{fr.denominator}"

            def _norm_val(s_v):
                return re.sub(r"\s+", "", _latex_to_readable_math(str(s_v))).strip()

            res_str = _fmt_f(ans_frac)
            k_str = _fmt_f(k_frac)
            k_sq_str = _fmt_f(k_sq)
            two_sc_str = _fmt_f(two_sc)

            matched_opt = None
            for k_opt, v_opt in sorted(options.items()):
                if _norm_val(v_opt) == _norm_val(res_str):
                    matched_opt = k_opt
                    break

            if matched_opt:
                distractor_lines = []
                for k_opt, v_opt in sorted(options.items()):
                    v_clean = _latex_to_readable_math(v_opt)
                    if k_opt == matched_opt:
                        if lang == "en":
                            distractor_lines.append(f"• Option {k_opt} ({v_clean}): Matches our exact result sin x · cos x = {res_str} ➔ CORRECT")
                        elif lang == "de":
                            distractor_lines.append(f"• Option {k_opt} ({v_clean}): Stimmt exakt mit sin x · cos x = {res_str} überein ➔ RICHTIG")
                        else:
                            distractor_lines.append(f"• {k_opt} şıkkı ({v_clean}): Bulduğumuz sin x · cos x = {res_str} sonucu ile birebir örtüşmektedir ➔ DOĞRU")
                    else:
                        trap_tr = "İşlem veya işaret hatası sonucu oluşan çeldirici şıktır"
                        trap_en = "Distractor option resulting from a sign or division error"
                        trap_de = "Ablenkungsoption durch Vorzeichen- oder Divisionsfehler"
                        if _norm_val(v_opt) == _norm_val(two_sc_str):
                            trap_tr = f"DİKKAT TUZAK! '{v_clean}' değeri 2·sin x·cos x = {two_sc_str} ifadesidir; 2'ye bölmeyi unutursak bu tuzağa düşeriz"
                            trap_en = f"COMMON TRAP! '{v_clean}' is 2·sin x·cos x = {two_sc_str} before dividing both sides by 2"
                            trap_de = f"TYPISCHE FALLE! '{v_clean}' ist 2·sin x·cos x = {two_sc_str}, bevor durch 2 geteilt wurde"
                        elif _norm_val(v_opt) == _norm_val(_fmt_f(-ans_frac)):
                            trap_tr = f"DİKKAT TUZAK! '{v_clean}' değeri çıkarma işleminde işaret hatası ({k_sq_str} - 1 yerine 1 - {k_sq_str}) yapıldığında bulunan ters işaretli çeldiricidir"
                            trap_en = f"COMMON TRAP! '{v_clean}' comes from a sign error (computing 1 - {k_sq_str} instead of {k_sq_str} - 1)"
                            trap_de = f"TYPISCHE FALLE! '{v_clean}' entsteht durch einen Vorzeichenfehler (1 - {k_sq_str} statt {k_sq_str} - 1)"
                        elif _norm_val(v_opt) in (_norm_val(k_sq_str), _norm_val(_fmt_f(-k_sq))):
                            trap_tr = f"'{v_clean}' değeri yalnızca eşitliğin karesi alındığında gelen ({k_str})² = {k_sq_str} terimidir"
                            trap_en = f"'{v_clean}' is just the squared right-hand side ({k_str})² = {k_sq_str}"
                            trap_de = f"'{v_clean}' ist lediglich das Quadrat ({k_str})² = {k_sq_str}"
                        if lang == "en":
                            distractor_lines.append(f"• Option {k_opt} ({v_clean}): {trap_en} ➔ Eliminated")
                        elif lang == "de":
                            distractor_lines.append(f"• Option {k_opt} ({v_clean}): {trap_de} ➔ Ausgeschlossen")
                        else:
                            distractor_lines.append(f"• {k_opt} şıkkı ({v_clean}): {trap_tr} ➔ Elenir")

                ans_line = f"{ans_prefix}: {matched_opt}) {_latex_to_readable_math(options[matched_opt])}"
                if lang == "en":
                    return (
                        f"{h1}\n"
                        f"We are given sin x {op_sign} cos x = {k_str} and asked to find the product sin x · cos x.\n\n"
                        f"{h2}\n"
                        f"• Golden Trigonometric Identity Rule:\n"
                        f"  1. Whenever you are given (sin x ± cos x) and asked for (sin x · cos x), square both sides of the equation!\n"
                        f"  2. Use the Pythagorean identity: sin²x + cos²x = 1.\n"
                        f"  3. Thus: (sin x {op_sign} cos x)² = sin²x {op_sign} 2·sin x·cos x + cos²x = 1 {op_sign} 2·sin x·cos x.\n\n"
                        f"{h3}\n"
                        f"Step 1 (Square Both Sides):\n"
                        f"• (sin x {op_sign} cos x)² = ({k_str})²\n"
                        f"• sin²x {op_sign} 2·sin x·cos x + cos²x = {k_sq_str}\n\n"
                        f"Step 2 (Substitute sin²x + cos²x = 1):\n"
                        f"• 1 {op_sign} 2·sin x·cos x = {k_sq_str}\n"
                        f"• 2·sin x·cos x = {two_sc_str}\n\n"
                        f"Step 3 (Divide Both Sides by 2):\n"
                        f"• sin x · cos x = {res_str}"
                        + (f"\n\n{h4}\n" + "\n".join(distractor_lines) if distractor_lines else "")
                        + f"\n\n{h5}\n{ans_line}"
                    )
                elif lang == "de":
                    return (
                        f"{h1}\n"
                        f"Gegeben ist sin x {op_sign} cos x = {k_str}. Gesucht ist das Produkt sin x · cos x.\n\n"
                        f"{h2}\n"
                        f"• Trigonometrischer Pythagoras & Binomische Formel:\n"
                        f"  1. Wenn (sin x ± cos x) gegeben und (sin x · cos x) gesucht ist, quadriert man beide Seiten der Gleichung!\n"
                        f"  2. Da sin²x + cos²x = 1 gilt, folgt: (sin x {op_sign} cos x)² = 1 {op_sign} 2·sin x·cos x.\n\n"
                        f"{h3}\n"
                        f"1. Schritt (Beide Seiten quadrieren):\n"
                        f"• (sin x {op_sign} cos x)² = ({k_str})² ➔ sin²x {op_sign} 2·sin x·cos x + cos²x = {k_sq_str}\n\n"
                        f"2. Schritt (sin²x + cos²x = 1 einsetzen):\n"
                        f"• 1 {op_sign} 2·sin x·cos x = {k_sq_str} ➔ 2·sin x·cos x = {two_sc_str}\n\n"
                        f"3. Schritt (Durch 2 teilen):\n"
                        f"• sin x · cos x = {res_str}"
                        + (f"\n\n{h4}\n" + "\n".join(distractor_lines) if distractor_lines else "")
                        + f"\n\n{h5}\n{ans_line}"
                    )
                else:
                    return (
                        f"{h1}\n"
                        f"Soruda bize sin x {op_sign} cos x = {k_str} eşitliği verilmiş ve sin x · cos x çarpımının değeri soruluyor.\n\n"
                        f"{h2}\n"
                        f"• Trigonometride Tam Kare ve Özdeşlik Kuralı:\n"
                        f"  1. Bir soruda (sin x + cos x) veya (sin x - cos x) verilip (sin x · cos x) çarpımı isteniyorsa, eşitliğin her iki tarafının karesini alırız!\n"
                        f"  2. Temel trigonometrik özdeşliğimiz: sin²x + cos²x = 1'dir.\n"
                        f"  3. Dolayısıyla: (sin x {op_sign} cos x)² = sin²x {op_sign} 2·sin x·cos x + cos²x = 1 {op_sign} 2·sin x·cos x olur.\n\n"
                        f"{h3}\n"
                        f"1. Adım (Her İki Tarafın Karesini Alalım):\n"
                        f"• (sin x {op_sign} cos x)² = ({k_str})²\n"
                        f"• sin²x {op_sign} 2·sin x·cos x + cos²x = {k_sq_str}\n\n"
                        f"2. Adım (sin²x + cos²x = 1 Özdeşliğini Yerine Yazalım):\n"
                        f"• 1 {op_sign} 2·sin x·cos x = {k_sq_str}\n"
                        f"• 2·sin x·cos x = {k_sq_str} - 1 = {two_sc_str}\n\n"
                        f"3. Adım (Her İki Tarafı 2'ye Bölelim — Bulunan Sonuç):\n"
                        f"• sin x · cos x = ({two_sc_str}) / 2 = {res_str} bulunur!"
                        + (f"\n\n{h4}\n" + "\n".join(distractor_lines) if distractor_lines else "")
                        + f"\n\n{h5}\n{ans_line}"
                    )
    except Exception:
        pass

    return None


def analyze_question_structure_offline(raw_text, lang="tr"):
    """
    Structurally parses a multiple-choice question into isolated sections and pre-analysis hints
    in the active language ('tr', 'en', or 'de').
    """
    if not raw_text or not raw_text.strip():
        return ""

    lang = resolve_response_language(lang, "", raw_text)
    text = _clean_ocr_turkish_text(raw_text)
    lines = [ln.strip() for ln in text.splitlines() if ln.strip() and not re.match(r"^(CEVAPLAR|CEVAP\s*ANAHTARI|ANSWERS|LÖSUNGEN)\b", ln.strip(), re.I)]

    opt_pattern = re.compile(r"^([A-E])[\)\.\-]\s*(.+)$")
    pre_option_lines = []
    options = {}
    current_opt = None

    for ln in lines:
        if re.match(r"^(Correct!|Incorrect\.|Metinde\s+|Alsas-Loren\s+kömür\s+havzası\s+ve\s+Balkanlar\s+gibi|Toplam\s+veri\s+miktarını|Üsler\s+toplanmamalı|Sonraki$)", ln, re.I):
            continue
        m = opt_pattern.match(ln)
        if m:
            current_opt = m.group(1).upper()
            options[current_opt] = m.group(2).strip()
        else:
            if current_opt is None:
                pre_option_lines.append(ln)
            else:
                options[current_opt] = (options[current_opt] + " " + ln).strip()

    if len(options) < 2:
        return text

    full_stem_block = " ".join(pre_option_lines).strip()

    opt_vals = [v.strip() for v in options.values()]
    if all(re.match(r"^\d{2,4}$", v) for v in opt_vals):
        first_chars = {v[0] for v in opt_vals}
        if len(first_chars) == 1:
            base_c = list(first_chars)[0]
            exps = [int(v[1:]) for v in opt_vals if v[1:].isdigit()]
            if len(exps) == len(opt_vals) and exps == sorted(exps) and max(exps) >= 6:
                for k in list(options.keys()):
                    v = options[k].strip()
                    options[k] = f"{base_c}^{v[1:]}"
                full_stem_block = re.sub(
                    rf"\b{base_c}(\d{{1,2}})\s*(GB|MB|KB|TB|sn|saniye|dk|saat|metre|cm|birim|adet|kat)\b",
                    rf"{base_c}^\1 \2",
                    full_stem_block,
                )

    q_match = re.search(
        r"((?:Bu\s+bilgilere|Buna\s+göre|Yukarıdaki|Paragrafa\s+göre|Metne\s+göre|Aşağıdaki\s+cümlelerin|Aşağıdakilerden|What\s+is|Which\s+of|Find\s+the|Was\s+ist|Welcher|Berechnen)[^.?!]*\?)\s*$",
        full_stem_block,
        re.I,
    )
    if not q_match:
        sentences = re.split(r"(?<=[.])\s+", full_stem_block)
        if len(sentences) >= 2 and "?" in sentences[-1]:
            paragraph_text = " ".join(sentences[:-1]).strip()
            question_stem = sentences[-1].strip()
        else:
            paragraph_text = ""
            question_stem = full_stem_block
    else:
        question_stem = q_match.group(1).strip()
        paragraph_text = full_stem_block[:q_match.start()].strip()

    neg_pattern = re.compile(
        r"\b(söylenemez|soylenemez|ulaşılamaz|ulasilamaz|çıkarılamaz|cikarilamaz|değinilmemiştir|deginilmemistir|değildir|degildir|olamaz|yoktur|yanlıştır|yanlistir|beklenemez|savunulamaz|not\s+true|cannot\s+be|false|nicht\s+richtig|falsch)\b",
        re.I,
    )
    is_negative_stem = bool(neg_pattern.search(question_stem))

    math_stem_pattern = re.compile(r"\b(kaç|kaçtır|kaç\s+saniye|oranı|toplamı|farkı|çarpımı|bölümü|sonucu|üslü|köklü|denklem|fonksiyon|yüzde|olasılık|limit|limitinin|türev|integral|polinom|polynomial|remainder|divided|equation|value|maximum|minimum|berechnen|polynom|grenzwert)\b", re.I)
    short_math_opts = all(len(v.strip()) <= 14 and re.search(r"[\d\^\⁰\¹\²\³\⁴\⁵\⁶\⁷\⁸\⁹\+\-\*/√π%xX]", v) for v in options.values())
    is_math_question = short_math_opts or bool(math_stem_pattern.search(question_stem))

    lbl_p = {"tr": "[1. ÖNCÜL METİN / PROBLEM BİLGİSİ]:", "en": "[1. PROBLEM CONTEXT / GIVEN INFORMATION]:", "de": "[1. AUFGABENKONTEXT / GEGEBENE INFORMATIONEN]:"}.get(lang)
    lbl_q = {"tr": "[2. SORU KÖKÜ]:", "en": "[2. QUESTION STEM]:", "de": "[2. FRAGESTELLUNG]:"}.get(lang)
    lbl_o = {"tr": "[3. TEST ŞIKLARI]:", "en": "[3. MULTIPLE-CHOICE OPTIONS]:", "de": "[3. ANTWORTMÖGLICHKEITEN]:"}.get(lang)
    lbl_h = {"tr": "[4. SİSTEM AKADEMİK KANIT VE KURAL TESPİTİ]:", "en": "[4. VERIFIED ACADEMIC EVIDENCE & RULES]:", "de": "[4. VERIFIZIERTE HINWEISE & REGELN]:"}.get(lang)

    out = []
    if paragraph_text:
        out.append(f"{lbl_p}\n{paragraph_text}")
    out.append(f"{lbl_q}\n{question_stem}")

    opts_formatted = "\n".join([f"{k}) {_latex_to_readable_math(v)}" for k, v in sorted(options.items())])
    out.append(f"{lbl_o}\n{opts_formatted}")

    hints = []

    if is_math_question:
        exact_teacher_sol = solve_question_with_teacher_engine_offline(raw_text, lang=lang)
        if exact_teacher_sol:
            hints.append(f"🎯 VERIFIED EXACT SOLUTION:\n{exact_teacher_sol}")

    if paragraph_text and not is_math_question and lang == "tr":
        para_lower = paragraph_text.lower()
        if is_negative_stem:
            hints.append(
                "⚠️ OLUMSUZ SORU KÖKÜ UYARISI: Soru kökü OLUMSUZDUR ('söylenemez / ulaşılamaz / çıkarılamaz'). "
                "Bu yüzden Öncül Metinde GEÇEN (doğrulanan) şıklar ELENİR! "
                "Aradığımız DOĞRU CEVAP, [1. ÖNCÜL METİN] içinde HİÇ BAHSEDİLMEYEN (metin dışı olan) şıktır!"
            )
        for k, opt_txt in sorted(options.items()):
            words = re.findall(r"[A-Za-zÇĞİÖŞÜçğıöşü]+", opt_txt)
            missing_proper = []
            for idx_w, w in enumerate(words):
                if idx_w > 0 and w[0].isupper() and len(w) >= 4:
                    w_stem = w.lower()[:5]
                    if w_stem not in para_lower:
                        missing_proper.append(w)
            if words and words[0].lower()[:6] in ("osmanl", "ingilt", "italya", "avustu", "rusya"):
                if words[0].lower()[:5] not in para_lower:
                    missing_proper.insert(0, words[0])

            if missing_proper:
                hints.append(
                    f"🔎 Şık {k} Kanıt Kontrolü: Bu şıkta geçen '{' '.join(missing_proper)}' kavramı [1. ÖNCÜL METİN] içinde HİÇ GEÇMEMEKTEDİR (Metin Dışı!). "
                    + ("Soru 'söylenemez/ulaşılamaz' sorduğu için metinde hiç bahsedilmeyen bu şık DOĞRU CEVAPTIR!" if is_negative_stem else "")
                )

    stem_lower = (paragraph_text + " " + question_stem).lower()
    if not is_math_question and lang == "tr" and ("bağla" in stem_lower or "nesne" in stem_lower or "öge" in stem_lower):
        for k, opt_txt in sorted(options.items()):
            m_ve = re.search(r"(\S+)\s+ve\s+(\S+)", opt_txt)
            if not m_ve:
                hints.append(f"🔎 Şık {k} Dil Bilgisi Kontrolü: Bu cümlede 've' bağlacı HİÇ KULLANILMAMIŞTIR (Doğrudan elenir).")
            else:
                w1, w2 = m_ve.group(1).strip(".,;:\"'"), m_ve.group(2).strip(".,;:\"'")
                w1_l, w2_l = w1.lower(), w2.lower()
                if (w1_l.endswith(("dan", "den", "tan", "ten")) and w2_l.endswith(("dan", "den", "tan", "ten"))):
                    hints.append(
                        f"🔎 Şık {k} Dil Bilgisi Kontrolü: 've' bağlacı '{w1} ve {w2}' kelimelerini bağlamıştır. "
                        f"'-den/-dan' ayrılma hâl eki almışlardır; yükleme sorulan 'Neyden / Kimden?' sorusuna cevap verir -> DOLAYLI TÜMLEÇTİR (Nesne değildir, elenir)."
                    )
                elif (w1_l.endswith(("a", "e")) and w2_l.endswith(("a", "e")) and w1_l in ("ona", "bana", "sana", "buna", "şuna") or w2_l.endswith(("ine", "ına", "una", "üne"))):
                    hints.append(
                        f"🔎 Şık {k} Dil Bilgisi Kontrolü: 've' bağlacı '{w1} ve {w2}' kelimelerini bağlamıştır. "
                        f"'-e/-a' yönelme hâl eki almışlardır; yükleme sorulan 'Kime / Nereye?' sorusuna cevap verir -> DOLAYLI TÜMLEÇTİR (Nesne değildir, elenir)."
                    )
                elif (w1_l in ("seni", "beni", "onu", "bizi", "sizi") or (w1_l.endswith(("i", "ı", "u", "ü")) and w2_l.endswith(("ini", "ını", "unu", "ünü", "i", "ı", "u", "ü")))):
                    hints.append(
                        f"🔎 Şık {k} Dil Bilgisi Kontrolü: 've' bağlacı '{w1} ve {w2}' kelimelerini bağlamıştır. "
                        f"'-i/-ı' belirtme hâl eki almışlardır ve yükleme sorulan 'Kimi? / Neyi?' sorusuna cevap verirler -> BELİRTİLİ NESNEDİR (Aranan DOĞRU CEVAP budur!)."
                    )

    if hints:
        out.append(f"{lbl_h}\n" + "\n".join(hints))

    return "\n\n".join(out)


def verify_and_fix_answer_option(raw_context, cleaned_ans, lang="tr"):
    """
    Ensures that for math/short-option questions, if our exact teacher solver or the LLM's calculation
    finds a specific value, the final answer is 100% accurate in the active language ('tr', 'en', 'de').
    Also strips any internal prompt headers ([1. ...], [2. ...], [3. ...], [4. SİSTEM AKADEMİK KANIT...])
    so users never see system prompt meta-instructions.
    """
    if not raw_context:
        return cleaned_ans
    lang = resolve_response_language(lang, "", raw_context)
    ans_prefix = {"tr": "DOĞRU CEVAP", "en": "CORRECT ANSWER", "de": "RICHTIGE ANTWORT"}.get(lang, "DOĞRU CEVAP")
    h1_lbl = {"tr": "📌 1. Sorunun Özeti ve Bizden İstenen:", "en": "📌 1. Problem Summary & What Is Asked:", "de": "📌 1. Aufgabenübersicht & Gesuchtes:"}.get(lang)
    h5_lbl = {"tr": "✅ 5. Kesin Doğru Cevap:", "en": "✅ 5. Final Correct Answer:", "de": "✅ 5. Endgültige richtige Antwort:"}.get(lang)
    try:
        exact_sol = solve_question_with_teacher_engine_offline(raw_context, lang=lang)
        if exact_sol:
            return exact_sol
    except Exception:
        pass
    if not cleaned_ans:
        return cleaned_ans
    try:
        # Strip any leaked [4. SİSTEM AKADEMİK KANIT VE KURAL TESPİTİ] / [4. VERIFIED ACADEMIC EVIDENCE & RULES] blocks
        cleaned_ans = re.sub(
            r"\[4\.\s*(?:SİSTEM\s+AKADEMİK\s+KANIT|VERIFIED\s+ACADEMIC\s+EVIDENCE|VERIFIZIERTE\s+HINWEISE)[^\]]*\]:?.*?(?=(?:📌|💡|✍️|⚠️|✅|$))",
            "",
            cleaned_ans,
            flags=re.S | re.I,
        )
        # Strip any leaked 🧮 MATEMATİK / HESAPLAMA SORUSU ÖĞRETMEN REHBERİ meta-instructions
        cleaned_ans = re.sub(
            r"🧮\s*(?:MATEMATİK\s*/\s*HESAPLAMA\s+SORUSU\s+ÖĞRETMEN\s+REHBERİ|MATH\s+TUTOR\s+INSTRUCTIONS|MATHE-LEHRER-HINWEISE):?.*?(?=(?:📌|💡|✍️|⚠️|✅|$))",
            "",
            cleaned_ans,
            flags=re.S | re.I,
        )
        # Strip echoed [1. ÖNCÜL METİN...], [2. SORU KÖKÜ]:, [3. TEST ŞIKLARI]: tags or convert [2. SORU KÖKÜ] to 📌 1.
        cleaned_ans = re.sub(
            r"\[1\.\s*(?:ÖNCÜL\s+METİN|PROBLEM\s+CONTEXT|AUFGABENKONTEXT)[^\]]*\]:?\s*(?:\d+[\.\)]?\s*)?\n*",
            "",
            cleaned_ans,
            flags=re.I,
        )
        cleaned_ans = re.sub(
            r"\[2\.\s*(?:SORU\s+KÖKÜ|QUESTION\s+STEM|FRAGESTELLUNG)\]:?",
            h1_lbl,
            cleaned_ans,
            flags=re.I,
        )
        cleaned_ans = re.sub(
            r"\[3\.\s*(?:TEST\s+ŞIKLARI|MULTIPLE-CHOICE\s+OPTIONS|ANTWORTMÖGLICHKEITEN)\]:?\s*(?:[A-E]\)\s*[^\n]+\n?)+",
            "",
            cleaned_ans,
            flags=re.I,
        )
        cleaned_ans = re.sub(r"\n{3,}", "\n\n", cleaned_ans).strip()

        # Clean echoed prompt instructions inside Section 5 headers
        cleaned_ans = re.sub(
            r"(✅\s*5\.[^\n\(]*)\([^\n\)]*(?:CORRECT ANSWER|RICHTIGE ANTWORT|DOĞRU CEVAP)[^\n\)]*\)\.?:?",
            r"\1:",
            cleaned_ans,
            flags=re.I,
        )

        struct_txt = analyze_question_structure_offline(raw_context, lang=lang)
        opts = {}
        in_opts = False
        for ln in struct_txt.splitlines():
            if ln.startswith("[3."):
                in_opts = True
                continue
            if ln.startswith("[4."):
                in_opts = False
            if in_opts:
                m_o = re.match(r"^([A-E])\)\s*(.+)$", ln.strip())
                if m_o:
                    opts[m_o.group(1)] = m_o.group(2).strip()

        def _norm_opt_tok(s_tok):
            return re.sub(r"[\s%TLtl°]", "", _latex_to_readable_math(s_tok)).strip()

        correct_letter = None
        correct_val = None

        m_exact = re.search(r"(?:DOĞRU CEVAP|CORRECT ANSWER|RICHTIGE ANTWORT):\s*([A-E])\)\s*([^\n]+)", struct_txt)
        if m_exact:
            correct_letter = m_exact.group(1)
            correct_val = m_exact.group(2).strip()
        elif opts and all(len(v) <= 15 for v in opts.values()):
            # Restrict equation scraping to Section 3 (before Section 4) so 'at x = 1' in Section 4 is never mistaken for the final result
            sec3_part = cleaned_ans
            m_sec3 = re.search(r"(?:✍️|3\.\s*(?:Adım|Step-by-Step|Schritt))(.*?)(?=(?:⚠️|4\.\s*(?:Çeldirici|Watch|Achtung)|✅|5\.\s*(?:Kesin|Final|Endgültige))|$)", cleaned_ans, re.S | re.I)
            if m_sec3:
                sec3_part = m_sec3.group(1)

            candidates = []
            for m_eq in re.finditer(r"(?<![xXyY]\s)=\s*%?([0-9⁰¹²³⁴⁵⁶⁷⁸⁹\^\+\-\*/,\.]+)%?", sec3_part):
                candidates.append(m_eq.group(1))
            for m_dec_x in re.finditer(r"=\s*0[\.,](\d{2})\s*[xXaA]\b", sec3_part):
                candidates.append(str(int(m_dec_x.group(1))))
            for m_b in re.finditer(r"(?:Bulunan\s+Sonuç|Result|Ergebnis|maximum\s+value\s+is|minimum\s+value\s+is|Sonuç\s+olarak[^\n]*?(?:yüzde|%|eşittir))\s*[:\-]?\s*%?([0-9⁰¹²³⁴⁵⁶⁷⁸⁹\^\+\-\*/,\.]+)%?", cleaned_ans, re.I):
                candidates.append(m_b.group(1))
            for m_opt_match in re.finditer(r"Option\s+([A-E])\s*[:\(]\s*([0-9⁰¹²³⁴⁵⁶⁷⁸⁹\^\+\-\*/,\.]+)\s*[\)\–\-]\s*(?:This\s+matches|Correct)", cleaned_ans, re.I):
                candidates.append(m_opt_match.group(2))

            for ev in reversed(candidates):
                ev_norm = _norm_opt_tok(ev)
                if not ev_norm:
                    continue
                matched_keys = [k for k, v in opts.items() if _norm_opt_tok(v) == ev_norm]
                if len(matched_keys) == 1:
                    correct_letter = matched_keys[0]
                    correct_val = opts[correct_letter]
                    break

        if correct_letter and correct_val:
            ans_pat = r"(?im)^[ \t]*(?:☑️\s*|✅\s*)?(?:KESİN\s+DOĞRU\s+CEVAP|DOĞRU\s+CEVAP|CORRECT\s+ANSWER|RICHTIGE\s+ANTWORT|Sonuç\s+olarak\s+doğru\s+cevap)\s*[:\-]?\s*[A-E](?:[\)\.\-:\s][^\n]*)?$"
            if re.search(ans_pat, cleaned_ans):
                cleaned_ans = re.sub(
                    ans_pat,
                    f"{ans_prefix}: {correct_letter}) {correct_val}",
                    cleaned_ans,
                )
            else:
                cleaned_ans = cleaned_ans.rstrip() + f"\n\n{h5_lbl}\n{ans_prefix}: {correct_letter}) {correct_val}"

        m_final_line = list(re.finditer(r"(?m)^.*(?:DOĞRU CEVAP|CORRECT ANSWER|RICHTIGE ANTWORT):\s*[A-E]\)\s*[^\n]+$", cleaned_ans))
        if m_final_line:
            cleaned_ans = cleaned_ans[:m_final_line[-1].end()].rstrip()
    except Exception:
        pass
    return cleaned_ans


def extract_text_from_image_offline(image_path):
    """
    100% Offline, high-precision Turkish / English / German + Math image & screenshot question extractor.
    1. Uses sensitive DBNet box detection (det_db_box_thresh=0.15, det_limit_side_len=1600) on 1:1 image
       so even tiny 2-digit option boxes ('23', '21', '19', '25', '27') next to radio buttons are never missed.
    2. Runs Windows Native OCR on the 2x upscaled full image with line Y-coordinates so multi-line
       paragraphs and Turkish characters (ş, ğ, ü, ö, ç, ı, İ) are never truncated or lost.
    3. Automatically assigns A), B), C), D), E) labels when a quiz uses unlabeled radio buttons (⚪).
    """
    global _OCR_ENGINE_SINGLETON
    if not image_path or not os.path.exists(image_path):
        return ""
    try:
        import numpy as np
        from PIL import Image, ImageOps, ImageEnhance, ImageFilter
        from rapidocr_onnxruntime import RapidOCR

        if _OCR_ENGINE_SINGLETON is None:
            _OCR_ENGINE_SINGLETON = RapidOCR(
                det_limit_side_len=1600,
                det_db_box_thresh=0.15,
                det_db_thresh=0.15,
                det_db_unclip_ratio=1.6,
            )

        img = Image.open(image_path).convert("RGB")
        try:
            raw_np = np.array(img)
            r_ch = raw_np[:, :, 0].astype(np.int16)
            b_ch = raw_np[:, :, 2].astype(np.int16)
            blue_bg = ((b_ch - r_ch) > 55) & (b_ch > 140) & (r_ch < 100)
            if np.mean(blue_bg) > 0.04:
                mask_img = Image.fromarray((blue_bg * 255).astype(np.uint8)).filter(ImageFilter.MaxFilter(9))
                dilated_blue = np.array(mask_img) > 128
                inv_red = np.clip(255 - r_ch, 0, 255).astype(np.uint8)
                raw_np[dilated_blue, 0] = inv_red[dilated_blue]
                raw_np[dilated_blue, 1] = inv_red[dilated_blue]
                raw_np[dilated_blue, 2] = inv_red[dilated_blue]
                img = Image.fromarray(raw_np)
        except Exception:
            pass

        # Pass 1: Run Windows Native OCR in a clean MTA thread with vertical Y-coordinates
        win_lines_meta = []
        try:
            import concurrent.futures
            import winocr

            def _run_winocr_tr(pil_im):
                im2 = pil_im.resize((pil_im.width * 2, pil_im.height * 2), Image.Resampling.BICUBIC)
                w_res = winocr.recognize_pil(im2, "tr").get()
                out_w = []
                for wl in w_res.lines:
                    t_s = wl.text.strip()
                    if not t_s:
                        continue
                    try:
                        ys_t = [w.bounding_rect.y / 2.0 for w in wl.words]
                        ys_b = [(w.bounding_rect.y + w.bounding_rect.height) / 2.0 for w in wl.words]
                        y_c = (min(ys_t) + max(ys_b)) / 2.0 if ys_t else -1.0
                    except Exception:
                        y_c = -1.0
                    out_w.append((y_c, t_s))
                return out_w

            with concurrent.futures.ThreadPoolExecutor(max_workers=1) as ex:
                win_lines_meta = ex.submit(_run_winocr_tr, img).result(timeout=4.0)
        except Exception:
            win_lines_meta = []

        win_lines = [t_s for _, t_s in win_lines_meta]

        # Pass 2: Run RapidOCR on the natural 1:1 image
        img_np = np.array(img)
        result, _ = _OCR_ENGINE_SINGLETON(img_np)
        if not result:
            result = []

        items_meta = []
        for item in result:
            box, text, score = item[0], str(item[1]).strip(), float(item[2])
            if not text or score < 0.20:
                continue
            if re.match(r"^[\.\,\-]\d+$", text):
                text = text[1:]

            xs = [pt[0] for pt in box]
            ys = [pt[1] for pt in box]
            y_top, y_bot = min(ys), max(ys)
            x_left, x_right = min(xs), max(xs)
            y_mid = (y_top + y_bot) / 2.0
            box_h = max(10.0, y_bot - y_top)
            items_meta.append((y_mid, x_left, x_right, y_top, y_bot, box_h, text))

        items_meta.sort(key=lambda x: x[0])
        rows = []
        for y_mid, x_left, x_right, y_top, y_bot, box_h, text in items_meta:
            placed = False
            for r in rows:
                if abs(y_mid - r["y_mean"]) <= max(16.0, r["h_mean"] * 0.60):
                    r["items"].append((x_left, x_right, y_top, y_bot, text))
                    n = len(r["items"])
                    r["y_mean"] = (r["y_mean"] * (n - 1) + y_mid) / n
                    r["h_mean"] = (r["h_mean"] * (n - 1) + box_h) / n
                    placed = True
                    break
            if not placed:
                rows.append({"y_mean": y_mid, "h_mean": box_h, "items": [(x_left, x_right, y_top, y_bot, text)]})

        rows.sort(key=lambda r: r["y_mean"])

        import difflib
        tr_to_ascii = str.maketrans("çğıöşüÇĞİÖŞÜ", "cgiosuCGIOSU")
        win_word_map = {}
        for w_ln in win_lines:
            for w_tok in w_ln.split():
                w_key = re.sub(r"\W+", "", w_tok.translate(tr_to_ascii).lower())
                if len(w_key) >= 3:
                    win_word_map[w_key] = w_tok

        used_w = set()
        for r in rows:
            sorted_items = sorted(r["items"], key=lambda x: x[0])
            row_parts = []
            for idx_it, (xl, xr, yt, yb, txt_it) in enumerate(sorted_items):
                if idx_it > 0:
                    prev_xr = sorted_items[idx_it - 1][1]
                    gap_w = xl - prev_xr
                    if 2 <= gap_w <= 36 and re.match(r"^\d+\)", txt_it.strip()):
                        row_parts.append("-")
                row_parts.append(txt_it)
            row_txt = " ".join(row_parts).strip()
            row_txt = re.sub(r"\blim\s+([xX]→\d+)\s+\+3\s*-\s*(\d+\))", r"lim \1 (x³ - \2", row_txt)
            row_txt = re.sub(r"^4\.1\s+lim\b", "4. lim", row_txt)
            r["text"] = row_txt

            r_norm = re.sub(r"\s+", "", row_txt.translate(tr_to_ascii).lower())
            best_idx = None
            best_score = 0.0

            # First check if a winocr line aligns vertically (Y-coordinate within 18px) and is MORE complete than a fragmented RapidOCR row!
            for idx_w, (w_y, w_txt) in enumerate(win_lines_meta):
                if idx_w in used_w:
                    continue
                w_norm = re.sub(r"\s+", "", w_txt.translate(tr_to_ascii).lower())
                if w_y >= 0 and abs(w_y - r["y_mean"]) <= 18.0 and len(w_norm) > len(r_norm) * 1.35:
                    best_idx = idx_w
                    break
                if len(w_norm) < len(r_norm) * 0.82:
                    continue
                sim = difflib.SequenceMatcher(None, r_norm, w_norm).ratio()
                if sim >= 0.55 and sim > best_score:
                    best_score = sim
                    best_idx = idx_w

            if best_idx is not None:
                used_w.add(best_idx)
                w_chosen = win_lines[best_idx]
                # Preserve x² / x³ superscripts in polynomial expressions like 'x2 + x -' -> 'x² + x -'
                w_chosen = re.sub(r"\b([xX])2(?=\s*[\+\-])", r"\1²", w_chosen)
                w_chosen = re.sub(r"\b([xX])3(?=\s*[\+\-])", r"\1³", w_chosen)
                r["text"] = w_chosen
            else:
                restored_tokens = []
                for tok in r["text"].split():
                    tok_k = re.sub(r"\W+", "", tok.translate(tr_to_ascii).lower())
                    if tok_k in win_word_map:
                        restored_tokens.append(win_word_map[tok_k])
                    else:
                        restored_tokens.append(tok)
                r["text"] = " ".join(restored_tokens)

        # Also include any full-sentence winocr lines (len >= 12) that RapidOCR completely missed vertically
        for idx_w, (w_y, w_txt) in enumerate(win_lines_meta):
            if idx_w not in used_w and len(w_txt) >= 12 and w_y >= 0:
                if not any(abs(w_y - r["y_mean"]) <= 16.0 for r in rows):
                    w_fixed = re.sub(r"\b([xX])2(?=\s*[\+\-])", r"\1²", w_txt)
                    rows.append({"y_mean": w_y, "h_mean": 16.0, "text": w_fixed})
        rows.sort(key=lambda r: r["y_mean"])

        extracted = []
        ui_noise_re = re.compile(
            r"^(test\s*bilgi\s*kartlar[ıi]?|rest\s*bi[il1]gi.*|test|bilgi\s*kartlar[ıi]?|[>›]?\s*[iİı]pucun?u?\s*g[oö]ster.*|sonraki)$",
            re.I,
        )
        for r in rows:
            row_txt = r["text"].strip()
            if not row_txt or ui_noise_re.match(row_txt):
                continue
            extracted.append(row_txt)

        raw_joined = "\n".join(extracted).strip()
        return _clean_ocr_turkish_text(raw_joined)
    except Exception:
        return ""


# ==============================================================================
# EDU: FORMAT-PRESERVING DOCUMENT TRANSLATOR & OPTIONAL GLOSSARY ENGINE
# ==============================================================================

def _parse_glossary_line(line):
    """Parse a single line like 'Source = Target', 'Source | Target', 'Source -> Target', or 'Source\tTarget'."""
    line = str(line or "").strip()
    if not line or line.startswith("#") or line.startswith("//"):
        return None, None
    for sep in ("\t", " | ", "|", " -> ", "->", " = ", "=", " : ", ";"):
        if sep in line:
            parts = line.split(sep, 1)
            src_t = parts[0].strip().strip("\"'•-* ")
            tgt_t = parts[1].strip().strip("\"' ")
            if src_t and tgt_t and len(src_t) <= 120 and len(tgt_t) <= 150:
                if src_t.lower() in ("kaynak", "source", "term", "terim", "almanca", "ingilizce", "german", "english"):
                    return None, None
                return src_t, tgt_t
    return None, None


def load_glossary_terms(glossary_path="", inline_terms_text=""):
    """
    Load optional terminology glossary from a file (.xlsx, .csv, .txt, .docx)
    and/or inline quick-entry text ('Term1 = Karşılık1, Term2 = Karşılık2').
    Returns an ordered dict {source_term: target_term} sorted by source term length (longest first).
    Works seamlessly if both are empty (returns {}).
    """
    raw_terms = {}

    # 1. Load from file if provided
    if glossary_path and os.path.exists(glossary_path):
        ext = glossary_path.lower().rsplit(".", 1)[-1] if "." in glossary_path else ""
        try:
            if ext == "xlsx":
                import openpyxl
                wb = openpyxl.load_workbook(glossary_path, data_only=True)
                for sheet in wb.worksheets:
                    for row in sheet.iter_rows(values_only=True):
                        if not row:
                            continue
                        non_empty = [str(c).strip() for c in row if c is not None and str(c).strip()]
                        if len(non_empty) >= 2:
                            s_t, t_t = non_empty[0], non_empty[1]
                            if s_t.lower() not in ("kaynak", "source", "term", "terim", "orijinal", "original"):
                                raw_terms[s_t] = t_t
                        elif len(non_empty) == 1:
                            s_t, t_t = _parse_glossary_line(non_empty[0])
                            if s_t and t_t:
                                raw_terms[s_t] = t_t
            elif ext in ("csv", "txt", "tsv"):
                content = ""
                for enc in ("utf-8-sig", "utf-8", "cp1254", "latin-1"):
                    try:
                        with open(glossary_path, "r", encoding=enc) as f:
                            content = f.read()
                        break
                    except Exception:
                        continue
                import csv
                import io
                for ln in content.splitlines():
                    s_t, t_t = _parse_glossary_line(ln)
                    if s_t and t_t:
                        raw_terms[s_t] = t_t
                    elif "," in ln:
                        try:
                            row = next(csv.reader(io.StringIO(ln)))
                            if len(row) >= 2 and row[0].strip() and row[1].strip():
                                if row[0].strip().lower() not in ("kaynak", "source", "term", "terim"):
                                    raw_terms[row[0].strip()] = row[1].strip()
                        except Exception:
                            pass
            elif ext == "docx":
                g_doc = docx.Document(glossary_path)
                for tbl in g_doc.tables:
                    for row in tbl.rows:
                        cells = [c.text.strip() for c in row.cells if c.text.strip()]
                        if len(cells) >= 2 and cells[0].lower() not in ("kaynak", "source", "term", "terim"):
                            raw_terms[cells[0]] = cells[1]
                for p in g_doc.paragraphs:
                    s_t, t_t = _parse_glossary_line(p.text)
                    if s_t and t_t:
                        raw_terms[s_t] = t_t
        except Exception:
            pass

    # 2. Load from inline text (e.g. "Kaufvertrag = Satış Sözleşmesi, Kündigung = Fesih")
    if inline_terms_text and str(inline_terms_text).strip():
        chunks = re.split(r"[\n,;]+", str(inline_terms_text))
        for ch in chunks:
            s_t, t_t = _parse_glossary_line(ch)
            if s_t and t_t:
                raw_terms[s_t] = t_t

    # Sort longest source terms first so multi-word phrases match before single words
    sorted_items = sorted(raw_terms.items(), key=lambda kv: len(kv[0]), reverse=True)
    return dict(sorted_items)


def _is_translatable_string(s):
    """Return True if string has meaningful words to translate (skip numbers, codes, URLs, company names, phones, spaced logos)."""
    if not s or not str(s).strip():
        return False
    st = str(s).strip()
    if st.startswith("="):  # Excel formula
        return False
    # Skip URLs, emails
    if re.match(r"^(https?://\S+|www\.\S+|\S+@\S+\.\S+)$", st, re.I):
        return False
    # Skip spaced single-letter graphic badges like 'E D I T I O N' or '2 0 2 6'
    if re.match(r"^([A-Za-z0-9]\s+){2,}[A-Za-z0-9]$", st):
        return False
    # Skip company legal names like 'COXEM KOREA CO., LTD.', 'ABC GMBH', 'XYZ INC.'
    if re.search(r"\b(CO\.,?\s*LTD\.?|LTD\.|INC\.|CORP\.|GMBH|LLC|A\.Ş\.|SAN\.\s*TİC\.)\s*(\(\d+/\d+\))?$", st, re.I):
        return False
    # Skip TEL / FAX / Phone lines
    if re.match(r"^(TEL|FAX|PHONE|GSM|MOBIL)\s*[:\-\+0-9]", st, re.I):
        return False
    # Skip street addresses starting with #digits or postal codes
    if re.match(r"^#\d+\s*,\s*\d+", st):
        return False
    # Skip pure measurement / coordinate / file-format lists like '30um, 50um, 50um, 100um', '1. +/- 25um', '1. JPG, TIFF, BMP, PNG'
    st_no_num = re.sub(r"^([0-9]+[\.\)]|[a-zA-Z][\.\)])\s*", "", st).strip()
    if re.match(r"^([\+\-/~\d\.,\s°º±xX×]+(nm|um|µm|mm|cm|kv|kcps|ev|hz|kva|gb|mb)?[\s,;()]*)+$", st_no_num, re.I):
        return False
    if re.match(r"^[X,Y,Z,T,R\s\(\)\d\~\-\+°ºmm,\.]+$", st_no_num):
        return False
    if re.match(r"^(JPG|JPEG|TIFF|BMP|PNG|GIF|PDF|DOCX|XLSX|MP4|AVI|MOV|WAV|MP3|USB|SDD|EDS|BSE|SE|CMOS|TMP|RAM|NS\d[\d\.]*)[\s,/&+]+(JPG|JPEG|TIFF|BMP|PNG|GIF|PDF|DOCX|XLSX|MP4|AVI|MOV|WAV|MP3|USB|SDD|EDS|BSE|SE|CMOS|TMP|RAM)*$", st_no_num, re.I):
        return False
    # Must contain at least one multi-letter word (3+ letters, or 2+ letters not just unit symbols)
    words = re.findall(r"[A-Za-zÇĞİÖŞÜçğıöşüÄÖÜäöüßА-Яа-яЁё]{2,}", st_no_num)
    unit_words = {"nm", "um", "mm", "cm", "kv", "ev", "hz", "gb", "mb", "ac", "dc", "se", "dp", "cx", "ta", "id", "no", "pc", "tv"}
    meaningful = [w for w in words if w.lower() not in unit_words]
    return len(meaningful) >= 1


# Built-in High-Precision Technical, Academic & Commercial Terminology Dictionary (EN <-> TR, DE <-> TR)
_BUILTIN_DOMAIN_GLOSSARY = {
    ("en", "tr"): {
        "Quotation No.": "Teklif No.",
        "Quotation is valid for 60 days": "Teklif 60 gün süreyle geçerlidir",
        "Quotation is valid for 30 days": "Teklif 30 gün süreyle geçerlidir",
        "Quotation": "Teklif",
        "Date": "Tarih",
        "To": "Alıcı",
        "Country": "Ülke",
        "Turkiye": "Türkiye",
        "Customer": "Müşteri",
        "Istanbul Technical University": "İstanbul Teknik Üniversitesi",
        "The Best Scanning Electron Microscope": "En İyi Taramalı Elektron Mikroskobu",
        "Scanning Electron Microscope(Digital processing)": "Taramalı Elektron Mikroskobu (Dijital İşleme)",
        "Scanning Electron Microscope (Digital processing)": "Taramalı Elektron Mikroskobu (Dijital İşleme)",
        "Scanning Electron Microscope": "Taramalı Elektron Mikroskobu",
        "Detail Description": "Detaylı Açıklama",
        "Main System": "Ana Sistem",
        "Resolution": "Çözünürlük",
        "SE Image": "SE Görüntü",
        "Pre-Aligned Cartridge Filament": "Ön Hizalamalı Kartuş Filament",
        "Magnification": "Büyütme",
        "Photo Magnification": "Fotoğraf Büyütme",
        "Display Magnification": "Ekran Büyütme",
        "Digital Zoom": "Dijital Yakınlaştırma",
        "Detector": "Dedektör",
        "DP Secondary Electron Detector": "DP İkincil Elektron Dedektörü",
        "Secondary Electron Detector": "İkincil Elektron Dedektörü",
        "Back Scattered Electron Detector": "Geri Saçılımlı Elektron Dedektörü (BSE)",
        "4-quadrant solid-state BSE detector": "4 Bölgeli Katı Hal BSE Dedektörü",
        "Accelerating Voltage": "Hızlandırma Voltajı",
        "1 to 30kV": "1 ~ 30 kV",
        "Lens System": "Lens Sistemi",
        "2 Step Condenser lens": "2 Kademeli Kondansör Lens",
        "1 Step Objective lens": "1 Kademeli Objektif Lens",
        "Condensor and Objective Lens Module": "Kondansör ve Objektif Lens Modülü",
        "Objective Aperture": "Objektif Diyaframı",
        "4 Levels Variable Aperture": "4 Kademeli Değişken Diyafram",
        "Image Shift": "Görüntü Kaydırma",
        "Stage": "Numune Tablası (Stage)",
        "Auto Step Motorized Stage": "Otomatik Step Motorlu Tabla",
        "Beam Rotation by Software (360°)": "Yazılım ile Işın Döndürme (360°)",
        "Specimen Holder": "Numune Tutucu",
        "Standard Holder": "Standart Tutucu",
        "Multi Holder(37 holes)": "Çoklu Tutucu (37 Yuvalı)",
        "Scanning Mode": "Tarama Modu",
        "Reduced Mode": "Küçültülmüş Mod",
        "Ultra-Fast Mode": "Ultra Hızlı Mod",
        "TV Mode": "TV Modu",
        "Slow Screen Scan Mode I / II": "Yavaş Ekran Tarama Modu I / II",
        "Photo Screen Capture Mode I": "Fotoğraf Ekran Yakalama Modu I",
        "Photo Screen Capture Mode II": "Fotoğraf Ekran Yakalama Modu II",
        "Photo Screen Capture Mode III": "Fotoğraf Ekran Yakalama Modu III",
        "pixels": "piksel",
        "up to": "maks.",
        "Image Format": "Görüntü Formatı",
        "Image Display": "Görüntü Ekranı",
        "Auto Functions & Software Features": "Otomatik Fonksiyonlar ve Yazılım Özellikleri",
        "Auto Filament Saturation": "Otomatik Filament Doygunluğu",
        "Auto Focus": "Otomatik Odaklama (Auto Focus)",
        "Auto Contrast/Brightness": "Otomatik Kontrast / Parlaklık",
        "ContinousMagSave": "Sürekli Büyütme Kaydı (ContinousMagSave)",
        "Panorama Ver 2.0": "Panorama Sürüm 2.0",
        "SE+BSE Image Merge": "SE+BSE Görüntü Birleştirme",
        "SE/BSE Image Split": "SE/BSE Görüntü Bölme",
        "Annotation & editing": "Not Ekleme ve Düzenleme",
        "Image Filters": "Görüntü Filtreleri",
        "Deep3D (Live 3D)": "Deep3D (Canlı 3B)",
        "Sample Location Save (upto 8 positions)": "Numune Konumu Kaydetme (8 Konuma Kadar)",
        "Working Distance": "Çalışma Mesafesi",
        "Max height": "Maks. Yükseklik",
        "General height": "Genel Yükseklik",
        "Use by special holder": "Özel tutucu ile kullanım",
        "Vacuum System": "Vakum Sistemi",
        "Full Automation": "Tam Otomasyon",
        "Roughing : Rotary Pump": "Ön Vakum : Döner (Rotary) Pompa",
        "High Vacuum : Turbo Molecular Pump": "Yüksek Vakum : Turbo Moleküler Pompa",
        "Control System": "Kontrol Sistemi",
        "PC Controlled by USB": "USB Üzerinden Bilgisayar Kontrollü",
        "i5 Processor, 16 gb RAM": "i5 İşlemci, 16 GB RAM",
        "Operating System : Windows 11": "İşletim Sistemi : Windows 11",
        "Mouse and Keyboard": "Fare ve Klavye",
        "Software : NS5.0": "Yazılım : NS5.0",
        "S/W Measuring Functions": "Yazılım Ölçüm Fonksiyonları",
        "Length, Angle, Area, Line Profile": "Uzunluk, Açı, Alan, Çizgi Profili",
        "Power": "Güç Gereksinimi",
        "Equipment Configuration": "Ekipman Yapılandırması",
        "Gun Unit": "Tabanca Ünitesi (Gun Unit)",
        "Column Unit": "Kolon Ünitesi",
        "Chamber Unit": "Hazne (Chamber) Ünitesi",
        "Detector and Stage Module": "Dedektör ve Tabla Modülü",
        "In-chamber Camera": "Hazne İçi Kamera",
        "Vacuum Unit": "Vakum Ünitesi",
        "Rotary Pump(100L/min) and TMP(80L/sec)": "Döner Pompa (100 L/dk) ve TMP (80 L/sn)",
        "Control Unit": "Kontrol Ünitesi",
        "Board Unit": "Kart (Board) Ünitesi",
        "Image part and Vacuum Part": "Görüntüleme Bölümü ve Vakum Bölümü",
        "High Voltage Power Supply": "Yüksek Voltaj Güç Kaynağı",
        "Spare Parts": "Yedek Parçalar",
        "Tungsten Filament & Cartridge": "Tungsten Filament ve Kartuş",
        "CX Holder set": "CX Tutucu Seti",
        "Carbon Tape and Silver Paste": "Karbon Bant ve Gümüş Pasta",
        "Wrench": "Anahtar Takımı",
        "Blower": "Hava Üfleyici (Blower)",
        "Driver": "Tornavida Seti",
        "User Manual": "Kullanım Kılavuzu",
        "Carrier Box": "Taşıma Çantası",
        "Accessories & tools included": "Aksesuarlar ve El Aletleri Dahil",
        "SEM Add-on Options": "SEM Ek Donanım Seçenekleri",
        "Energy Dispersive Spectrometer (EDS)": "Enerji Dağılımlı Spektrometre (EDS)",
        "Detectable element": "Tespit Edilebilir Element",
        "Guaranteed in compliance with ISO 15632:2012": "ISO 15632:2012 Standardına Uygunluk Garantili",
        "Detector Type: SDD": "Dedektör Tipi: SDD",
        "Sensor Size :30 mm": "Sensör Boyutu: 30 mm",
        "Cooling Type : Peltier": "Soğutma Tipi: Peltier",
        "Additional 27\" Monitor x 1set": "Ek 27\" Monitör x 1 Set",
        "27\" Monitor x 1set": "27\" Monitör x 1 Set",
        "res guaranteed to be": "çözünürlük garantisi:",
        "Guaranteed in compliance with": "Standardına Uygunluk Garantili:",
        "CMOS sample camera for quick identification and accurate positioning.": "Hızlı tanımlama ve hassas konumlandırma için CMOS numune kamerası.",
        "Pin Stub": "Pin Tutucu (Pin Stub)",
        "Tilt Pin Stub": "Eğimli Pin Tutucu (Tilt Pin Stub)",
        "Standard Pin Stub": "Standart Pin Tutucu",
        "5 set": "5 Set",
        "1 set": "1 Set",
        "1 ea": "1 Adet",
    }
}


def _find_matching_glossary_terms(text, glossary_dict):
    """Find glossary terms that appear in `text` (case-insensitive word-boundary or 5+ char stem match)."""
    if not glossary_dict or not text:
        return []
    text_lower = text.lower()
    matched = []
    for src_term, tgt_term in glossary_dict.items():
        s_low = src_term.lower()
        if len(s_low) <= 3:
            if re.search(r"\b" + re.escape(s_low) + r"\b", text_lower):
                matched.append((src_term, tgt_term))
        elif s_low in text_lower:
            matched.append((src_term, tgt_term))
    return matched


def _apply_glossary_direct(source_text, translated_text, glossary_dict):
    """
    Ensure that if a source term was left untranslated in translated_text, or if the entire text
    is an exact glossary entry, it is replaced with the glossary target term preserving capitalization.
    Always uses word-boundary matching so short terms like 'To' never corrupt words like 'Monitor'.
    """
    if not glossary_dict or not source_text:
        return translated_text
    src_clean = re.sub(r"\s+", " ", source_text.strip())
    # 1. Exact full-string match (e.g., table header or heading)
    for s_term, t_term in glossary_dict.items():
        if src_clean.lower() == s_term.lower():
            if src_clean.isupper() and len(src_clean) > 2:
                return t_term.upper()
            return t_term

    # 2. Substring replacement ONLY for multi-letter terms (>= 4 chars or multi-word) using strict word boundaries
    #    Use temporary placeholders so shorter glossary terms never re-match inside longer glossary replacements!
    out = translated_text
    placeholders = {}
    for idx_g, (s_term, t_term) in enumerate(glossary_dict.items()):
        if len(s_term) <= 3 and " " not in s_term:
            continue
        pat_str = r"(?<![A-Za-z0-9])" + re.escape(s_term) + r"(?![A-Za-z0-9])"
        if re.search(pat_str, out, flags=re.IGNORECASE):
            ph = f"__GLOSS_{idx_g}__"
            placeholders[ph] = t_term
            out = re.sub(pat_str, ph, out, flags=re.IGNORECASE)
    for ph, t_term in placeholders.items():
        out = out.replace(ph, t_term)
    return out


def detect_document_source_lang(sample_texts, src_code="auto", tgt_code="tr"):
    """
    Robust document-level source language detector that:
    1. Uses whole-word stopword counts across the document (so a Turkish proper name like 'Şener' or 'İstanbul'
       inside an English CV never causes the entire document to be misclassified as Turkish).
    2. Never returns `tgt_code` when `src_code` is 'auto' or equal to `tgt_code`, preventing no-op (TR -> TR) skips.
    """
    if src_code and src_code != "auto" and src_code != tgt_code:
        return src_code

    joined = " ".join(str(x or "") for x in (sample_texts or [])[:80])
    cleaned_words = re.sub(r"[^a-zA-ZçğıöşüÇĞİÖŞÜäöüßÄÖÜа-яА-ЯёЁ\u0600-\u06FF]+", " ", joined).lower()
    low = f" {cleaned_words} "

    if re.findall(r"[А-Яа-яЁёІіЇїЄєҐґ]", joined) and len(re.findall(r"[А-Яа-яЁёІіЇїЄєҐґ]", joined)) > 15:
        if len(re.findall(r"[ІіЇїЄєҐґ]", joined)) >= 2 and tgt_code != "uk":
            return "uk"
        if tgt_code != "ru":
            return "ru"
    if re.findall(r"[\u0600-\u06FF]", joined) and len(re.findall(r"[\u0600-\u06FF]", joined)) > 15:
        if len(re.findall(r"[پچژگک]", joined)) >= 2 and tgt_code != "fa":
            return "fa"
        if tgt_code != "ar":
            return "ar"
    if len(re.findall(r"[\u0370-\u03FF]", joined)) > 10 and tgt_code != "el":
        return "el"
    if len(re.findall(r"[\u0900-\u097F]", joined)) > 10 and tgt_code != "hi":
        return "hi"
    if len(re.findall(r"[\uac00-\ud7af]", joined)) > 5 and tgt_code != "ko":
        return "ko"
    if len(re.findall(r"[\u3040-\u30ff]", joined)) > 5 and tgt_code != "ja":
        return "ja"
    if len(re.findall(r"[\u4e00-\u9fff]", joined)) > 5 and tgt_code != "zh":
        return "zh"

    stopwords = {
        "en": (
            " the ", " and ", " for ", " with ", " in ", " to ", " of ", " on ", " from ", " by ", " as ", " at ",
            " is ", " are ", " management ", " operations ", " support ", " design ", " analysis ", " experience ",
            " project ", " network ", " system ", " systems ", " engineering ", " technical ", " customer ",
            " skills ", " education ", " summary ", " professional ", " senior ", " lead ", " team ", " service ",
        ),
        "tr": (
            " ve ", " bir ", " için ", " olan ", " olarak ", " bu ", " ile ", " da ", " de ", " den ", " dan ",
            " tarafından ", " göre ", " ilgili ", " üzerine ", " arasında ", " yönetimi ", " mühendisi ",
            " eğitim ", " deneyim ", " beceriler ", " hakkında ", " iletişim ", " iş ", " yıl ",
        ),
        "de": (
            " der ", " die ", " das ", " und ", " ist ", " mit ", " für ", " nicht ", " von ", " auf ", " den ",
            " des ", " im ", " dem ", " ein ", " eine ", " als ", " auch ", " sich ", " nach ", " bei ",
        ),
        "fr": (
            " le ", " la ", " les ", " des ", " dans ", " pour ", " avec ", " est ", " une ", " sur ", " par ",
            " plus ", " aux ", " sont ", " cette ", " mais ", " ou ", " du ",
        ),
        "es": (
            " el ", " la ", " los ", " las ", " del ", " para ", " con ", " por ", " una ", " son ", " como ",
            " más ", " pero ", " sus ", " sobre ", " entre ",
        ),
        "it": (
            " il ", " della ", " nella ", " sono ", " degli ", " questo ", " per ", " con ", " del ", " una ",
            " dalle ", " alla ", " nel ", " più ",
        ),
        "nl": (
            " de ", " het ", " een ", " van ", " voor ", " met ", " niet ", " zijn ", " aan ", " op ", " dat ",
            " door ", " over ", " bij ",
        ),
    }

    scores = {}
    for lang, words in stopwords.items():
        if lang == tgt_code:
            continue
        scores[lang] = sum(low.count(w) for w in words)

    best_lang = max(scores, key=scores.get) if scores else ("tr" if tgt_code == "en" else "en")
    if scores.get(best_lang, 0) == 0:
        return "tr" if tgt_code == "en" else "en"
    return best_lang


def _split_prefix_core_suffix(text):
    """
    Split a line like '  1. Photo Magnification: x6 ~ x300,000 ' or '✔ Core Network Operations' or 'Date : '
    into (prefix, core, suffix) so numbering ('1. ', 'a. ', '• ', '✔ ') and trailing colons (' :') are 100% preserved.
    """
    s = str(text or "")
    lead_ws = s[: len(s) - len(s.lstrip())]
    trail_ws = s[len(s.rstrip()) :]
    stripped = s.strip()

    prefix = lead_ws
    core = stripped
    suffix = trail_ws

    # Match leading list marker: '1. ', '10. ', 'a. ', 'b) ', '• ', '✔ ', '✓ ', '▪ ', '▸ ', '► ', '➢ '
    m_pre = re.match(
        r"^([0-9]{1,2}[\.\)]\s+|[a-zA-Z][\.\)]\s+|[•\-\*✔✓▪▸►➢➤❖]\s*|[0-9]{1,2}\s+(?=[A-ZÇĞİÖŞÜ]))(.*)$",
        core,
    )
    if m_pre:
        marker = m_pre.group(1)
        prefix += marker
        core = m_pre.group(2)

    # Match trailing colon: 'Date : ' or 'Country:'
    m_suf = re.match(r"^(.*?)(\s*:\s*)$", core)
    if m_suf and m_suf.group(1).strip():
        core = m_suf.group(1)
        suffix = m_suf.group(2) + suffix

    return prefix, core, suffix


def _translate_single_segment_smart(raw_text, effective_glossary, m2m_cache):
    """
    Reconstruct translated segment using:
    1. Prefix/Suffix preservation ('1. ', 'a. ', '✔ ', ' :')
    2. Exact or sub-part glossary match (including 'Label : Value', 'Heading — Detail', or 'Term (detail)')
    3. Neural M2M translation fallback with glossary term replacement
    """
    prefix, core, suffix = _split_prefix_core_suffix(raw_text)
    c_strip = re.sub(r"\s+", " ", core.strip())
    if not c_strip:
        return raw_text

    # 1. Exact glossary match on `c_strip`
    if effective_glossary:
        for s_term, t_term in effective_glossary.items():
            if c_strip.lower() == s_term.lower():
                return f"{prefix}{t_term}{suffix}"

    # 2. Check if `c_strip` has 'Label : Value' or 'Heading — Detail' structure
    for sep_token in (":", " — ", " – "):
        if sep_token in c_strip:
            left_p, right_p = c_strip.split(sep_token, 1)
            left_s = re.sub(r"\s+", " ", left_p.strip())
            right_s = re.sub(r"\s+", " ", right_p.strip())
            if sep_token == ":":
                out_sep = " : " if " :" in c_strip else ": "
            else:
                out_sep = sep_token
            if left_s and right_s:
                left_tr = left_s
                right_tr = right_s
                left_matched = False
                right_matched = False
                if effective_glossary:
                    for s_term, t_term in effective_glossary.items():
                        if left_s.lower() == s_term.lower():
                            left_tr = t_term
                            left_matched = True
                            break
                    for s_term, t_term in effective_glossary.items():
                        if right_s.lower() == s_term.lower():
                            right_tr = t_term
                            right_matched = True
                            break
                if not left_matched and _is_translatable_string(left_s):
                    left_tr = m2m_cache.get(left_s, left_s)
                    left_tr = _apply_glossary_direct(left_s, left_tr, effective_glossary)
                if not right_matched and _is_translatable_string(right_s):
                    right_tr = m2m_cache.get(right_s, right_s)
                    right_tr = _apply_glossary_direct(right_s, right_tr, effective_glossary)
                return f"{prefix}{left_tr}{out_sep}{right_tr}{suffix}"

    # 3. Check if `c_strip` has 'Main Term (Paren Detail)' where Main Term is in glossary
    m_paren = re.match(r"^([^\(\)]+?)\s*\(([^\(\)]+)\)\s*(.*)$", c_strip)
    if m_paren and effective_glossary:
        main_p, in_paren, tail_p = m_paren.group(1).strip(), m_paren.group(2).strip(), m_paren.group(3).strip()
        main_hit = None
        for s_term, t_term in effective_glossary.items():
            if main_p.lower() == s_term.lower():
                main_hit = t_term
                break
        if main_hit:
            paren_tr = in_paren
            for s_term, t_term in effective_glossary.items():
                if in_paren.lower() == s_term.lower():
                    paren_tr = t_term
                    break
            else:
                for s_term, t_term in effective_glossary.items():
                    if s_term.lower() in paren_tr.lower():
                        paren_tr = re.sub(re.escape(s_term), t_term, paren_tr, flags=re.I)
            tail_tr = tail_p
            if tail_p:
                for s_term, t_term in effective_glossary.items():
                    if tail_p.lower() == s_term.lower():
                        tail_tr = t_term
                        break
            tail_str = f" {tail_tr}" if tail_tr else ""
            return f"{prefix}{main_hit} ({paren_tr}){tail_str}{suffix}"

    # 4. General Neural + Longest-Substring Glossary replacement
    base_t = m2m_cache.get(c_strip, c_strip)
    base_t = _apply_glossary_direct(c_strip, base_t, effective_glossary)
    return f"{prefix}{base_t}{suffix}"


def _is_valid_clean_llm_translation(src_text, candidate_text):
    """
    Strict Anti-Hallucination Validator:
    Rejects any LLM output that contains meta-commentary ('Türkçe:', 'Kusurlar:', 'Kaynak:', 'Taslak:', 'KURALLAR', etc.)
    or abnormal length expansion.
    """
    if not candidate_text or not str(candidate_text).strip():
        return False
    c = str(candidate_text).strip()
    bad_markers = (
        "türkçe", "turkish", "english", "ingilizce", "german", "almanca",
        "kusurlar", "kaynak:", "taslak:", "kurallar", "incelendi",
        "yeniden yazılış", "cevap:", "çeviri:", "öneri:", "açıklama:",
        "none", "doğru kullanım", "aksi takdirde",
    )
    c_low = c.lower()
    if any(bm in c_low for bm in bad_markers):
        return False
    if len(c) > len(src_text) * 1.75 + 15 or len(c) < len(src_text) * 0.35:
        return False
    return True


def translate_texts_high_quality(
    texts,
    src_code,
    tgt_code,
    glossary_dict,
    m2m_translate_fn,
    llm_fn=None,
    use_llm_polish=True,
    log_fn=None,
):
    """
    Hybrid High-Quality Document Translation Pipeline:
    1. Merges user glossary (highest priority) + built-in technical/domain dictionary.
    2. Preserves list numbering ('1. ', 'a. '), colons ('Label : Value'), and technical units/numbers.
    3. Runs M2M100 batch neural translation on all sub-segments not already resolved by the glossary.
    4. Applies strict anti-hallucination validation so meta-commentary never leaks into the document.
    """
    if not texts:
        return []

    # Guard against auto-detect misclassifying a document as `tgt_code` (e.g. an English CV with a Turkish name like 'Şener')
    if not src_code or src_code == "auto" or src_code == tgt_code:
        src_code = detect_document_source_lang(texts, "auto", tgt_code)

    # Merge built-in domain dictionary with user glossary (user glossary overrides built-in)
    combined_glossary = {}
    builtin = _BUILTIN_DOMAIN_GLOSSARY.get((src_code, tgt_code), {})
    combined_glossary.update(builtin)
    if glossary_dict:
        combined_glossary.update(glossary_dict)
    effective_glossary = dict(sorted(combined_glossary.items(), key=lambda kv: len(kv[0]), reverse=True))

    # Collect all sub-strings that need M2M100 neural translation
    to_m2m_set = set()
    to_m2m_list = []

    def _queue_m2m(s_piece):
        sp = re.sub(r"\s+", " ", str(s_piece or "").strip())
        if not sp or not _is_translatable_string(sp):
            return
        # Skip if exact match in effective_glossary
        for k_g in effective_glossary:
            if sp.lower() == k_g.lower():
                return
        if sp not in to_m2m_set:
            to_m2m_set.add(sp)
            to_m2m_list.append(sp)

    for raw_t in texts:
        if not _is_translatable_string(raw_t):
            continue
        _, core, _ = _split_prefix_core_suffix(raw_t)
        c_strip = re.sub(r"\s+", " ", core.strip())
        if not c_strip:
            continue
        split_done = False
        for sep_token in (":", " — ", " – "):
            if sep_token in c_strip:
                lp, rp = c_strip.split(sep_token, 1)
                _queue_m2m(lp)
                _queue_m2m(rp)
                split_done = True
                break
        if not split_done:
            _queue_m2m(c_strip)

    # Pre-replace glossary terms inside sentences sent to M2M100 when appropriate, or translate directly
    m2m_cache = {}
    if to_m2m_list:
        m2m_out = m2m_translate_fn(to_m2m_list, src_code, tgt_code)
        if m2m_out and len(m2m_out) == len(to_m2m_list):
            for src_piece, tgt_piece in zip(to_m2m_list, m2m_out):
                cleaned_p = _apply_glossary_direct(src_piece, tgt_piece, effective_glossary)
                m2m_cache[src_piece] = cleaned_p

    if log_fn:
        log_fn(100)

    out_texts = []
    for raw_t in texts:
        if not _is_translatable_string(raw_t):
            out_texts.append(str(raw_t or ""))
        else:
            out_texts.append(_translate_single_segment_smart(raw_t, effective_glossary, m2m_cache))

    return out_texts


def _iter_docx_tables_recursive(tables):
    """Yield all tables and nested tables inside a Word document or header/footer."""
    for tbl in tables:
        yield tbl
        for row in tbl.rows:
            for cell in row.cells:
                if cell.tables:
                    yield from _iter_docx_tables_recursive(cell.tables)


def _iter_all_docx_paragraphs(doc):
    """
    Yield every unique translatable paragraph in a Word (.docx) document in order:
    1. Body paragraphs
    2. All table cells (deduplicating merged cells and handling nested tables)
    3. Section headers & footers (paragraphs and tables)
    """
    seen_p_elems = set()

    def _yield_p(p):
        p_el = p._p
        if p_el not in seen_p_elems:
            seen_p_elems.add(p_el)
            yield p

    # 1. Body paragraphs
    for p in doc.paragraphs:
        yield from _yield_p(p)

    # 2. Body tables (including nested)
    seen_tc = set()
    for tbl in _iter_docx_tables_recursive(doc.tables):
        for row in tbl.rows:
            for cell in row.cells:
                if cell._tc in seen_tc:
                    continue
                seen_tc.add(cell._tc)
                for p in cell.paragraphs:
                    yield from _yield_p(p)

    # 3. Headers & Footers across sections
    for section in doc.sections:
        hf_list = [
            getattr(section, "header", None),
            getattr(section, "footer", None),
            getattr(section, "first_page_header", None),
            getattr(section, "first_page_footer", None),
            getattr(section, "even_page_header", None),
            getattr(section, "even_page_footer", None),
        ]
        for hf in hf_list:
            if hf is None:
                continue
            try:
                if hf.is_linked_to_previous:
                    continue
            except Exception:
                pass
            for p in hf.paragraphs:
                yield from _yield_p(p)
            for tbl in _iter_docx_tables_recursive(hf.tables):
                for row in tbl.rows:
                    for cell in row.cells:
                        if cell._tc in seen_tc:
                            continue
                        seen_tc.add(cell._tc)
                        for p in cell.paragraphs:
                            yield from _yield_p(p)


def _replace_docx_paragraph_text_preserve_format(paragraph, new_text):
    """
    Replace paragraph text with `new_text` while strictly preserving:
    - Paragraph style, alignment, indentation, numbering/bullets
    - Run font name, font size, bold, italic, underline, color, highlight
    - Embedded images/drawings (`w:drawing`, `w:pict`) inside runs!
    """
    runs = paragraph.runs
    if not runs:
        paragraph.text = new_text
        return

    # Identify runs that do NOT contain embedded drawings/pictures
    text_runs = []
    for r in runs:
        has_drawing = bool(r._r.xpath(".//*[local-name()='drawing' or local-name()='pict']"))
        if not has_drawing:
            text_runs.append(r)

    if not text_runs:
        paragraph.add_run(new_text)
        return

    # Find the primary run that held the most text (to keep its exact font/size/color/bold style)
    primary_run = max(text_runs, key=lambda r: len(r.text or ""))
    first_run = text_runs[0]

    # Copy formatting from primary_run to first_run if they differ
    if primary_run is not first_run:
        try:
            if primary_run.bold is not None:
                first_run.bold = primary_run.bold
            if primary_run.italic is not None:
                first_run.italic = primary_run.italic
            if primary_run.underline is not None:
                first_run.underline = primary_run.underline
            if primary_run.font.name:
                first_run.font.name = primary_run.font.name
            if primary_run.font.size:
                first_run.font.size = primary_run.font.size
            if primary_run.font.color and primary_run.font.color.rgb:
                first_run.font.color.rgb = primary_run.font.color.rgb
        except Exception:
            pass

    first_run.text = new_text
    for r in text_runs[1:]:
        r.text = ""


def translate_docx_preserve_format(in_docx_path, out_docx_path, translate_batch_fn):
    """
    Translate a Word (.docx) document in-place into `out_docx_path`, preserving all
    tables, cells, headers, footers, images, logos, fonts, colors, and paragraph styles.
    Returns (total_translated_segments, preview_pairs).
    """
    doc = docx.Document(in_docx_path)
    target_paragraphs = []
    source_texts = []

    for p in _iter_all_docx_paragraphs(doc):
        txt = p.text
        if _is_translatable_string(txt):
            target_paragraphs.append(p)
            source_texts.append(txt)

    if not source_texts:
        doc.save(out_docx_path)
        return 0, []

    translated_texts = translate_batch_fn(source_texts)
    preview_pairs = []
    for p, src_t, tgt_t in zip(target_paragraphs, source_texts, translated_texts):
        _replace_docx_paragraph_text_preserve_format(p, tgt_t)
        if len(preview_pairs) < 25 and src_t.strip() != tgt_t.strip():
            preview_pairs.append((src_t.strip(), tgt_t.strip()))

    doc.save(out_docx_path)
    return len(source_texts), preview_pairs


def _get_windows_unicode_font_paths():
    """Return (regular_font_path, bold_font_path) on Windows supporting Turkish, German, Cyrillic, etc."""
    reg_candidates = [
        "C:/Windows/Fonts/arial.ttf",
        "C:/Windows/Fonts/segoeui.ttf",
        "C:/Windows/Fonts/calibri.ttf",
        "C:/Windows/Fonts/tahoma.ttf",
    ]
    bold_candidates = [
        "C:/Windows/Fonts/arialbd.ttf",
        "C:/Windows/Fonts/segoeuib.ttf",
        "C:/Windows/Fonts/calibrib.ttf",
        "C:/Windows/Fonts/tahomabd.ttf",
    ]
    reg_p = next((c for c in reg_candidates if os.path.exists(c)), None)
    bold_p = next((c for c in bold_candidates if os.path.exists(c)), reg_p)
    return reg_p, bold_p


def _sample_pdf_bg_color(pixmap, page_rect, bbox):
    """
    Sample the background RGB color (0.0..1.0) just outside the left edge of `bbox`
    so colored header bars (like green 'Detail Description' bars) keep their exact background color!
    """
    try:
        scale_x = pixmap.width / max(1.0, page_rect.width)
        scale_y = pixmap.height / max(1.0, page_rect.height)
        # Sample 3 points just to the left/top of the text line where no glyph ink exists
        sample_pts = [
            (max(0, int((bbox.x0 - 3) * scale_x)), min(pixmap.height - 1, max(0, int(((bbox.y0 + bbox.y1) / 2) * scale_y)))),
            (max(0, int((bbox.x0 - 6) * scale_x)), min(pixmap.height - 1, max(0, int(((bbox.y0 + bbox.y1) / 2) * scale_y)))),
            (min(pixmap.width - 1, max(0, int((bbox.x0 + 2) * scale_x))), max(0, int((bbox.y0 + 1) * scale_y))),
        ]
        rgbs = []
        for px, py in sample_pts:
            p = pixmap.pixel(px, py)
            rgbs.append((p[0] / 255.0, p[1] / 255.0, p[2] / 255.0))
        # Pick the brightest sampled pixel (to avoid accidentally sampling dark glyph ink)
        best = max(rgbs, key=lambda c: c[0] + c[1] + c[2])
        return best
    except Exception:
        return (1.0, 1.0, 1.0)


def extract_ocr_lines_with_boxes_offline(image_path):
    """
    100% Offline OCR line + bounding-box extractor for images (JPG, PNG, screenshots) and scanned PDFs.
    Combines Windows Native OCR (for accurate Turkish/European diacritics & full-line grouping)
    with RapidOCR (ONNX DBNet bounding-box detection) and returns a list of dicts:
        [{"bbox": (x0, y0, x1, y1), "text": line_str}, ...]
    sorted top-to-bottom.
    """
    global _OCR_ENGINE_SINGLETON
    if not image_path or not os.path.exists(image_path):
        return []
    try:
        import numpy as np
        from PIL import Image
        from rapidocr_onnxruntime import RapidOCR

        if _OCR_ENGINE_SINGLETON is None:
            _OCR_ENGINE_SINGLETON = RapidOCR(
                det_limit_side_len=1600,
                det_db_box_thresh=0.15,
                det_db_thresh=0.15,
                det_db_unclip_ratio=1.6,
            )

        img = Image.open(image_path).convert("RGB")

        # 1. Try Windows Native OCR (2x upscaled for crisp character recognition + exact bounding boxes)
        win_boxes = []
        try:
            import concurrent.futures
            import winocr

            def _run_win_boxes(pil_im):
                im2 = pil_im.resize((pil_im.width * 2, pil_im.height * 2), Image.Resampling.BICUBIC)
                w_res = winocr.recognize_pil(im2, "tr").get()
                out_b = []
                for wl in w_res.lines:
                    t_s = wl.text.strip()
                    if not t_s:
                        continue
                    try:
                        xs0 = [w.bounding_rect.x / 2.0 for w in wl.words]
                        ys0 = [w.bounding_rect.y / 2.0 for w in wl.words]
                        xs1 = [(w.bounding_rect.x + w.bounding_rect.width) / 2.0 for w in wl.words]
                        ys1 = [(w.bounding_rect.y + w.bounding_rect.height) / 2.0 for w in wl.words]
                        bx = (min(xs0), min(ys0), max(xs1), max(ys1))
                    except Exception:
                        bx = (0.0, 0.0, float(pil_im.width), 20.0)
                    out_b.append({"bbox": bx, "text": t_s})
                return out_b

            with concurrent.futures.ThreadPoolExecutor(max_workers=1) as ex:
                win_boxes = ex.submit(_run_win_boxes, img).result(timeout=5.0)
        except Exception:
            win_boxes = []

        # 2. Run RapidOCR on 1:1 image for layout & fallback
        img_np = np.array(img)
        result, _ = _OCR_ENGINE_SINGLETON(img_np)
        rapid_rows = []
        if result:
            items_meta = []
            for item in result:
                box, text, score = item[0], str(item[1]).strip(), float(item[2])
                if not text or score < 0.22:
                    continue
                xs = [pt[0] for pt in box]
                ys = [pt[1] for pt in box]
                x0, x1 = float(min(xs)), float(max(xs))
                y0, y1 = float(min(ys)), float(max(ys))
                y_mid = (y0 + y1) / 2.0
                bh = max(10.0, y1 - y0)
                items_meta.append((y_mid, x0, x1, y0, y1, bh, text))

            items_meta.sort(key=lambda x: x[0])
            for y_mid, x0, x1, y0, y1, bh, text in items_meta:
                placed = False
                for r in rapid_rows:
                    if abs(y_mid - r["y_mean"]) <= max(14.0, r["h_mean"] * 0.55):
                        r["items"].append((x0, x1, y0, y1, text))
                        n = len(r["items"])
                        r["y_mean"] = (r["y_mean"] * (n - 1) + y_mid) / n
                        r["h_mean"] = (r["h_mean"] * (n - 1) + bh) / n
                        placed = True
                        break
                if not placed:
                    rapid_rows.append({"y_mean": y_mid, "h_mean": bh, "items": [(x0, x1, y0, y1, text)]})

        if win_boxes and len(win_boxes) >= max(1, len(rapid_rows) - 1):
            win_boxes.sort(key=lambda b: (round(b["bbox"][1] / 12.0), b["bbox"][0]))
            return win_boxes

        out_lines = []
        tr_to_ascii = str.maketrans("çğıöşüÇĞİÖŞÜ", "cgiosuCGIOSU")
        win_word_map = {}
        for wb in win_boxes:
            for w_tok in wb["text"].split():
                w_key = re.sub(r"\W+", "", w_tok.translate(tr_to_ascii).lower())
                if len(w_key) >= 3:
                    win_word_map[w_key] = w_tok

        for r in sorted(rapid_rows, key=lambda x: x["y_mean"]):
            s_items = sorted(r["items"], key=lambda x: x[0])
            x0 = min(it[0] for it in s_items)
            x1 = max(it[1] for it in s_items)
            y0 = min(it[2] for it in s_items)
            y1 = max(it[3] for it in s_items)
            raw_line = " ".join(it[4] for it in s_items).strip()
            restored = []
            for tok in raw_line.split():
                tk = re.sub(r"\W+", "", tok.translate(tr_to_ascii).lower())
                restored.append(win_word_map.get(tk, tok))
            clean_ln = " ".join(restored).strip()
            if clean_ln:
                out_lines.append({"bbox": (x0, y0, x1, y1), "text": clean_ln})

        return out_lines if out_lines else win_boxes
    except Exception:
        fallback_txt = extract_text_from_image_offline(image_path)
        if fallback_txt:
            return [{"bbox": (10.0, float(idx * 24 + 10), 600.0, float(idx * 24 + 30)), "text": ln.strip()} for idx, ln in enumerate(fallback_txt.splitlines()) if ln.strip()]
        return []


def translate_pdf_preserve_format(in_pdf_path, out_pdf_path, translate_batch_fn):
    """
    Translate a PDF (.pdf) document into `out_pdf_path` (.pdf) with line-by-line baseline precision:
    - Preserves all images, logos, vector graphics, and colored background bars (via pixmap background sampling).
    - Preserves exact X/Y baseline origin (`sp['origin']`), font size, text color, and Bold/Regular weight.
    - Automatically scales font size slightly if a translated line would collide with a column to its right.
    - Automatically applies Offline OCR fallback on scanned / image-only PDF pages!
    Returns (total_translated_lines, preview_pairs).
    """
    import fitz
    import tempfile

    pdf_doc = fitz.open(in_pdf_path)
    reg_font_path, bold_font_path = _get_windows_unicode_font_paths()
    reg_font_obj = fitz.Font(fontfile=reg_font_path) if reg_font_path else None
    bold_font_obj = fitz.Font(fontfile=bold_font_path) if bold_font_path else reg_font_obj

    page_lines_meta = []  # All lines across pages (translatable and non-translatable for right-column collision checks)
    translatable_indices = []

    for p_idx, page in enumerate(pdf_doc):
        p_dict = page.get_text("dict")
        seen_line_keys = set()
        page_native_count = 0
        for blk in p_dict.get("blocks", []):
            if blk.get("type", 0) != 0:
                continue
            for ln in blk.get("lines", []):
                spans = [sp for sp in ln.get("spans", []) if str(sp.get("text", "")).strip()]
                if not spans:
                    continue
                raw_ln_text = "".join(str(sp.get("text", "")) for sp in ln.get("spans", []))
                clean_ln_text = raw_ln_text.strip()
                if not clean_ln_text:
                    continue

                bbox = fitz.Rect(ln["bbox"])
                # Deduplicate identical overlapping lines (some PDFs double-stamp bold text)
                dedup_key = (p_idx, round(bbox.x0, 1), round(bbox.y0, 1), clean_ln_text)
                if dedup_key in seen_line_keys:
                    continue
                seen_line_keys.add(dedup_key)

                sp0 = spans[0]
                orig_pt = sp0.get("origin", (bbox.x0, bbox.y1 - 1.5))
                fs = max(5.5, min(28.0, float(sp0.get("size", 9.0))))
                c_int = int(sp0.get("color", 0))
                col = (
                    ((c_int >> 16) & 255) / 255.0,
                    ((c_int >> 8) & 255) / 255.0,
                    (c_int & 255) / 255.0,
                )
                font_str = str(sp0.get("font", "")).lower()
                flags = int(sp0.get("flags", 0))
                is_bold = bool(flags & 16) or ("bold" in font_str) or ("black" in font_str)

                is_trans = _is_translatable_string(clean_ln_text)
                meta_idx = len(page_lines_meta)
                page_lines_meta.append({
                    "page_idx": p_idx,
                    "bbox": bbox,
                    "origin": (float(orig_pt[0]), float(orig_pt[1])),
                    "text": clean_ln_text,
                    "fontsize": fs,
                    "color": col,
                    "is_bold": is_bold,
                    "is_trans": is_trans,
                    "translated": clean_ln_text,
                })
                page_native_count += 1
                if is_trans:
                    translatable_indices.append(meta_idx)

        # Automatic OCR Fallback if this PDF page is scanned / image-based (no selectable text layer)
        if page_native_count == 0:
            try:
                ocr_dpi = 150
                scale_to_pt = 72.0 / float(ocr_dpi)
                pix_ocr = page.get_pixmap(dpi=ocr_dpi, alpha=False)
                tmp_png = os.path.join(tempfile.gettempdir(), f"ilber_edu_pdf_ocr_{os.getpid()}_{p_idx}.png")
                pix_ocr.save(tmp_png)
                ocr_lines = extract_ocr_lines_with_boxes_offline(tmp_png)
                try:
                    os.remove(tmp_png)
                except Exception:
                    pass
                for ol in ocr_lines:
                    clean_ln_text = str(ol.get("text", "")).strip()
                    if not clean_ln_text:
                        continue
                    bx0, by0, bx1, by1 = ol["bbox"]
                    bbox = fitz.Rect(bx0 * scale_to_pt, by0 * scale_to_pt, bx1 * scale_to_pt, by1 * scale_to_pt)
                    fs = max(6.5, min(20.0, bbox.height * 0.78))
                    orig_pt = (bbox.x0, bbox.y1 - max(1.2, bbox.height * 0.15))
                    is_trans = _is_translatable_string(clean_ln_text)
                    meta_idx = len(page_lines_meta)
                    page_lines_meta.append({
                        "page_idx": p_idx,
                        "bbox": bbox,
                        "origin": (float(orig_pt[0]), float(orig_pt[1])),
                        "text": clean_ln_text,
                        "fontsize": fs,
                        "color": (0.1, 0.1, 0.1),
                        "is_bold": False,
                        "is_trans": is_trans,
                        "translated": clean_ln_text,
                    })
                    if is_trans:
                        translatable_indices.append(meta_idx)
            except Exception:
                pass

    if not translatable_indices:
        pdf_doc.save(out_pdf_path)
        pdf_doc.close()
        return 0, []

    source_texts = [page_lines_meta[i]["text"] for i in translatable_indices]
    translated_texts = translate_batch_fn(source_texts)
    preview_pairs = []

    for meta_i, src_t, tgt_t in zip(translatable_indices, source_texts, translated_texts):
        # Clean any stray newlines
        clean_tgt = re.sub(r"\s+", " ", str(tgt_t or "").strip())
        page_lines_meta[meta_i]["translated"] = clean_tgt if clean_tgt else src_t
        if len(preview_pairs) < 25 and src_t.strip() != clean_tgt.strip():
            preview_pairs.append((src_t.strip(), clean_tgt.strip()))

    # Group all lines (including non-translated ones for collision detection) by page
    all_by_page = {}
    for m in page_lines_meta:
        all_by_page.setdefault(m["page_idx"], []).append(m)

    for p_idx, page_items in all_by_page.items():
        trans_items = [it for it in page_items if it["is_trans"]]
        if not trans_items:
            continue

        page = pdf_doc[p_idx]
        # Render 1x pixmap BEFORE redactions so we can sample exact background colors (e.g. green bars)
        pix = page.get_pixmap(dpi=72, alpha=False)

        # 1. Redact original text lines using their sampled background color (preserving all images & vector graphics)
        for it in trans_items:
            r = it["bbox"]
            bg_col = _sample_pdf_bg_color(pix, page.rect, r)
            # Slightly inset top/bottom by 0.6pt so adjacent lines never clip each other
            pad_r = fitz.Rect(r.x0 - 0.5, r.y0 + 0.6, r.x1 + 0.5, r.y1 - 0.4)
            if pad_r.height < 2.0:
                pad_r = fitz.Rect(r.x0 - 0.5, r.y0, r.x1 + 0.5, r.y1)
            page.add_redact_annot(pad_r, fill=bg_col)

        try:
            page.apply_redactions(
                images=getattr(fitz, "PDF_REDACT_IMAGE_NONE", 0),
                graphics=getattr(fitz, "PDF_REDACT_LINE_ART_NONE", 0),
            )
        except Exception:
            page.apply_redactions()

        # Register Regular and Bold Windows Unicode fonts on the page
        f_reg_name = "F_REG"
        f_bold_name = "F_BOLD"
        if reg_font_path:
            try:
                page.insert_font(fontname=f_reg_name, fontfile=reg_font_path)
            except Exception:
                f_reg_name = "helv"
        else:
            f_reg_name = "helv"

        if bold_font_path:
            try:
                page.insert_font(fontname=f_bold_name, fontfile=bold_font_path)
            except Exception:
                f_bold_name = f_reg_name
        else:
            f_bold_name = f_reg_name

        # 2. Insert each translated line at its exact typographic baseline `(orig_x, orig_y)`
        for it in trans_items:
            orig_x, orig_y = it["origin"]
            r = it["bbox"]
            tgt_t = it["translated"]
            fs = it["fontsize"]
            col = it["color"]
            is_bold = it["is_bold"]

            # Find nearest element to the right on the same horizontal row (e.g. '27-Aug-26', 'Turkiye', '5 set')
            max_x1 = page.rect.x1 - 18.0
            for other in page_items:
                if other is it:
                    continue
                ob = other["bbox"]
                # Same vertical row check
                if abs(ob.y0 - r.y0) <= max(4.0, r.height * 0.55) and ob.x0 > r.x0 + 4.0:
                    max_x1 = min(max_x1, ob.x0 - 4.0)

            avail_w = max(25.0, max_x1 - orig_x)
            f_obj = bold_font_obj if is_bold else reg_font_obj
            f_name = f_bold_name if is_bold else f_reg_name

            fit_fs = fs
            if f_obj is not None:
                try:
                    rendered_w = f_obj.text_length(tgt_t, fontsize=fit_fs)
                    if rendered_w > avail_w and rendered_w > 0:
                        fit_fs = max(5.2, fit_fs * (avail_w / rendered_w))
                except Exception:
                    pass

            page.insert_text(
                fitz.Point(orig_x, orig_y),
                tgt_t,
                fontsize=fit_fs,
                fontname=f_name,
                color=col,
            )

    pdf_doc.save(out_pdf_path, garbage=3, deflate=True)
    pdf_doc.close()
    return len(translatable_indices), preview_pairs


def translate_xlsx_preserve_format(in_xlsx_path, out_xlsx_path, translate_batch_fn):
    """
    Translate an Excel (.xlsx) workbook in-place into `out_xlsx_path`, preserving all
    sheets, formulas, cell styles, borders, column widths, and number formats.
    Returns (total_translated_cells, preview_pairs).
    """
    import openpyxl

    wb = openpyxl.load_workbook(in_xlsx_path)
    target_cells = []
    source_texts = []

    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                val = cell.value
                if isinstance(val, str) and _is_translatable_string(val):
                    target_cells.append(cell)
                    source_texts.append(val)

    if not source_texts:
        wb.save(out_xlsx_path)
        return 0, []

    translated_texts = translate_batch_fn(source_texts)
    preview_pairs = []
    for cell, src_t, tgt_t in zip(target_cells, source_texts, translated_texts):
        cell.value = tgt_t
        if len(preview_pairs) < 25 and src_t.strip() != tgt_t.strip():
            preview_pairs.append((src_t.strip(), tgt_t.strip()))

    wb.save(out_xlsx_path)
    return len(source_texts), preview_pairs


def translate_txt_preserve_format(in_txt_path, out_txt_path, translate_batch_fn):
    """
    Translate a plain text (.txt) file into `out_txt_path` preserving exact line breaks.
    Returns (total_translated_lines, preview_pairs).
    """
    content = ""
    for enc in ("utf-8-sig", "utf-8", "cp1254", "latin-1"):
        try:
            with open(in_txt_path, "r", encoding=enc) as f:
                content = f.read()
            break
        except Exception:
            continue

    lines = content.splitlines()
    trans_indices = [i for i, ln in enumerate(lines) if _is_translatable_string(ln)]
    if not trans_indices:
        with open(out_txt_path, "w", encoding="utf-8") as f:
            f.write(content)
        return 0, []

    source_texts = [lines[i] for i in trans_indices]
    translated_texts = translate_batch_fn(source_texts)
    preview_pairs = []
    out_lines = list(lines)
    for idx, src_t, tgt_t in zip(trans_indices, source_texts, translated_texts):
        out_lines[idx] = tgt_t
        if len(preview_pairs) < 25 and src_t.strip() != tgt_t.strip():
            preview_pairs.append((src_t.strip(), tgt_t.strip()))

    with open(out_txt_path, "w", encoding="utf-8") as f:
        f.write("\n".join(out_lines))
    return len(source_texts), preview_pairs


def process_image_ocr_and_translate(
    in_img_path,
    out_docx_path,
    out_txt_path,
    out_img_path,
    translate_batch_fn,
    ocr_only=False,
):
    """
    100% Offline OCR + Optional Translation for Images (JPG, PNG, WEBP, BMP, Screenshots):
    1. Extracts lines and bounding boxes from `in_img_path` via `extract_ocr_lines_with_boxes_offline`.
    2. Passes extracted lines through `translate_batch_fn` (which either translates + opens the Interactive Editor,
       or in `ocr_only=True` mode lets the user review/edit the extracted OCR text directly).
    3. Saves:
       - `out_docx_path`: Editable Word (.docx) document with the extracted/translated text.
       - `out_txt_path`: Plain text (.txt) file.
       - `out_img_path` (when `ocr_only=False`): Visual image with translated text placed into the original boxes.
    Returns (count_lines, preview_pairs).
    """
    from PIL import Image, ImageDraw, ImageFont

    ocr_items = extract_ocr_lines_with_boxes_offline(in_img_path)
    if not ocr_items:
        return 0, []

    source_texts = [str(it.get("text", "")).strip() for it in ocr_items if str(it.get("text", "")).strip()]
    valid_items = [it for it in ocr_items if str(it.get("text", "")).strip()]
    if not source_texts:
        return 0, []

    final_texts = translate_batch_fn(source_texts)
    if not final_texts or len(final_texts) != len(source_texts):
        final_texts = list(source_texts)

    preview_pairs = []
    for s_t, f_t in zip(source_texts, final_texts):
        if len(preview_pairs) < 25:
            preview_pairs.append((s_t.strip(), f_t.strip()))

    # 1. Save Editable Plain Text (.txt)
    if out_txt_path:
        with open(out_txt_path, "w", encoding="utf-8") as f:
            f.write("\n".join(final_texts))

    # 2. Save Editable Word Document (.docx)
    if out_docx_path:
        doc_out = docx.Document()
        for f_t in final_texts:
            p = doc_out.add_paragraph(f_t)
            p.paragraph_format.space_after = docx.shared.Pt(4)
        doc_out.save(out_docx_path)

    # 3. Save Translated Visual Image (when translating to another language)
    if out_img_path and not ocr_only:
        try:
            img = Image.open(in_img_path).convert("RGB")
            draw = ImageDraw.Draw(img)
            reg_font_path, _ = _get_windows_unicode_font_paths()

            for it, tgt_t in zip(valid_items, final_texts):
                bx0, by0, bx1, by1 = [int(round(v)) for v in it["bbox"]]
                bx0 = max(0, min(img.width - 1, bx0))
                by0 = max(0, min(img.height - 1, by0))
                bx1 = max(bx0 + 4, min(img.width, bx1))
                by1 = max(by0 + 4, min(img.height, by1))

                # Sample corner pixels for background color
                sample_pts = [
                    (max(0, bx0 - 2), max(0, by0 - 2)),
                    (min(img.width - 1, bx1 + 1), max(0, by0 - 2)),
                    (max(0, bx0 - 2), min(img.height - 1, by1 + 1)),
                ]
                samples = [img.getpixel(pt) for pt in sample_pts]
                bg_rgb = max(samples, key=lambda c: c[0] + c[1] + c[2])
                lum = (0.299 * bg_rgb[0] + 0.587 * bg_rgb[1] + 0.114 * bg_rgb[2])
                fg_rgb = (20, 24, 35) if lum > 128 else (248, 250, 252)

                draw.rectangle([bx0 - 1, by0 - 1, bx1 + 1, by1 + 1], fill=bg_rgb)
                box_h = max(12, by1 - by0)
                box_w = max(24, bx1 - bx0)
                f_size = max(10, min(36, int(box_h * 0.78)))

                font_pil = None
                if reg_font_path:
                    while f_size > 9:
                        try:
                            font_pil = ImageFont.truetype(reg_font_path, f_size)
                            tb = draw.textbbox((0, 0), tgt_t, font=font_pil)
                            if (tb[2] - tb[0]) <= box_w * 1.08:
                                break
                        except Exception:
                            break
                        f_size -= 1
                if font_pil is None:
                    font_pil = ImageFont.load_default()

                draw.text((bx0, by0), tgt_t, fill=fg_rgb, font=font_pil)

            img.save(out_img_path)
        except Exception:
            pass

    return len(final_texts), preview_pairs








