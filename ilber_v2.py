import os
import sys
import io
import subprocess

# 1. Ensure sys.stdout and sys.stderr are never None in PyInstaller --noconsole builds
#    (prevents tqdm / huggingface / model loaders from crashing on a clean PC)
if sys.stdout is None:
    sys.stdout = io.StringIO()
if sys.stderr is None:
    sys.stderr = io.StringIO()

# 2. Register bundled application directory in PATH so bundled ffmpeg.exe / ffprobe.exe
#    work out-of-the-box on any computer without installing anything externally
_app_dirs = [os.path.dirname(os.path.abspath(__file__))]
if getattr(sys, "frozen", False):
    _app_dirs.insert(0, os.path.dirname(os.path.abspath(sys.executable)))
    if hasattr(sys, "_MEIPASS"):
        _app_dirs.insert(0, sys._MEIPASS)
        if sys._MEIPASS not in sys.path:
            sys.path.append(sys._MEIPASS)
os.environ["PATH"] = os.pathsep.join(_app_dirs) + os.pathsep + os.environ.get("PATH", "")

try:
    import importlib.metadata as _imeta
    _orig_imeta_version = _imeta.version
    _FALLBACK_PKG_VERSIONS = {
        "torch": "2.6.0",
        "torchaudio": "2.6.0",
        "numpy": "1.26.4",
        "transformers": "4.44.2",
        "tokenizers": "0.19.1",
        "huggingface-hub": "0.24.6",
        "huggingface_hub": "0.24.6",
        "safetensors": "0.4.4",
        "tqdm": "4.66.5",
        "regex": "2024.7.24",
        "requests": "2.32.3",
        "packaging": "24.1",
        "filelock": "3.15.4",
        "pyyaml": "6.0.2",
        "sentencepiece": "0.2.0",
    }
    def _safe_imeta_version(pkg_name):
        try:
            return _orig_imeta_version(pkg_name)
        except Exception as e:
            key = str(pkg_name).lower().replace("_", "-")
            if key in _FALLBACK_PKG_VERSIONS:
                return _FALLBACK_PKG_VERSIONS[key]
            raise e
    _imeta.version = _safe_imeta_version
except Exception:
    pass

try:
    import certifi
    if os.path.exists(certifi.where()):
        os.environ.setdefault("SSL_CERT_FILE", certifi.where())
        os.environ.setdefault("REQUESTS_CA_BUNDLE", certifi.where())
except Exception:
    pass

# 3. Prevent any child console window (e.g. ffmpeg called by whisper) from flashing on Windows
if os.name == "nt":
    _orig_popen = subprocess.Popen
    class _NoWindowPopen(_orig_popen):
        def __init__(self, *args, **kwargs):
            if "creationflags" not in kwargs:
                kwargs["creationflags"] = 0x08000000
            else:
                kwargs["creationflags"] |= 0x08000000
            super().__init__(*args, **kwargs)
    subprocess.Popen = _NoWindowPopen

import docx
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


def _lazy_import_whisper():
    """Import numba/torch/whisper on demand so startup UI opens in <0.8 seconds."""
    if "whisper" in sys.modules:
        return sys.modules["whisper"]
    try:
        import numpy as np
        import numpy.core._multiarray_umath  # Ensure NumPy C-API is loaded before torch._C initializes
    except Exception:
        pass
    try:
        import torch
    except Exception:
        pass
    try:
        import types as _types
        import numba as _numba
        if not hasattr(_numba, "jit"):
            def _fallback_jit(*args, **kwargs):
                if len(args) == 1 and callable(args[0]):
                    return args[0]
                return lambda fn: fn
            _numba.jit = _fallback_jit
    except Exception:
        import types as _types
        _fake_numba = _types.ModuleType("numba")
        def _fallback_jit(*args, **kwargs):
            if len(args) == 1 and callable(args[0]):
                return args[0]
            return lambda fn: fn
        _fake_numba.jit = _fallback_jit
        sys.modules["numba"] = _fake_numba
    import whisper as _wh
    return _wh

from ilber_core_utils import (
    resource_path,
    open_path_crossplatform,
    format_ts_hms,
    format_ts_srt,
    format_ts_vtt,
    clean_markdown_text,
    add_formatted_text_to_doc,
    enforce_subtitle_rules,
    write_srt_file,
    write_vtt_file,
    acoustic_diarize_segments,
    detect_anonymization_candidates,
    run_edu_4stage_pipeline,
    save_bilingual_word_doc,
    generate_field_ai_suggestions,
    export_field_outputs,
    extract_text_from_image_offline,
    _clean_ocr_turkish_text,
    resolve_response_language,
    analyze_question_structure_offline,
    solve_question_with_teacher_engine_offline,
    verify_and_fix_answer_option,
    load_glossary_terms,
    translate_texts_high_quality,
    translate_docx_preserve_format,
    translate_pdf_preserve_format,
    translate_xlsx_preserve_format,
    translate_txt_preserve_format,
)
from ilber_ui_dialogs import InteractiveSegmentEditor, FieldResearchStudio, ScreenSnipperOverlay, DocumentTranslationEditor
from ilber_license import SupabaseLicenseManager, PACKAGE_TIERS, generate_signed_license_key


SUPPORTED_MEDIA_EXTS = (
    ("Medya Dosyaları (Ses & Video)", "*.mp3 *.wav *.m4a *.ogg *.flac *.mp4 *.mov *.mkv *.webm"),
    ("Ses Dosyaları", "*.mp3 *.wav *.m4a *.ogg *.flac"),
    ("Video Dosyaları", "*.mp4 *.mov *.mkv *.webm"),
    ("Tüm Dosyalar", "*.*"),
)

SUPPORTED_DOC_EXTS = (
    ("Belgeler & Görseller (.docx, .pdf, .xlsx, .txt, .jpg, .png)", "*.docx *.pdf *.xlsx *.txt *.png *.jpg *.jpeg *.bmp *.webp"),
    ("Word Belgeleri (*.docx)", "*.docx"),
    ("PDF Belgeleri (*.pdf)", "*.pdf"),
    ("Excel Tabloları (*.xlsx)", "*.xlsx"),
    ("Görsel / Ekran Görüntüsü - OCR (*.png, *.jpg, *.webp)", "*.png *.jpg *.jpeg *.bmp *.webp"),
    ("Metin Dosyaları (*.txt)", "*.txt"),
    ("Tüm Dosyalar", "*.*"),
)

SUPPORTED_GLOSSARY_EXTS = (
    ("Terim Sözlüğü Dosyaları (.xlsx, .csv, .txt, .docx)", "*.xlsx *.csv *.txt *.tsv *.docx"),
    ("Excel Sözlük (*.xlsx)", "*.xlsx"),
    ("CSV / Metin Sözlük (*.csv, *.txt)", "*.csv *.txt *.tsv"),
    ("Word Sözlük (*.docx)", "*.docx"),
    ("Tüm Dosyalar", "*.*"),
)

SUPPORTED_CHAT_EXTS = (
    ("Görsel, PDF, Ses ve Video", "*.png *.jpg *.jpeg *.bmp *.webp *.pdf *.mp3 *.wav *.m4a *.ogg *.flac *.mp4 *.mov *.mkv *.webm"),
    ("Görsel / Fotoğraf Dosyaları", "*.png *.jpg *.jpeg *.bmp *.webp"),
    ("PDF Belgeleri", "*.pdf"),
    ("Tüm Dosyalar", "*.*"),
)

LANG_MAP = {
    "Auto": "auto",
    "OCR Only (Extract Text)": "ocr_only",
    "Sadece OCR (Metne Dönüştür)": "ocr_only",
    "Turkish": "tr",
    "English": "en",
    "German": "de",
    "Spanish": "es",
    "French": "fr",
    "Italian": "it",
    "Russian": "ru",
    "Arabic": "ar",
    "Chinese": "zh",
    "Japanese": "ja",
    "Korean": "ko",
    "Portuguese": "pt",
    "Dutch": "nl",
}

TARGET_LANGS = [
    "English", "Turkish", "German", "Spanish", "French",
    "Italian", "Russian", "Arabic", "Portuguese", "Dutch",
]

EDU_TARGET_LANGS = [
    "English", "Turkish", "German", "Spanish", "French",
    "Italian", "Russian", "Arabic", "Portuguese", "Dutch",
    "OCR Only (Extract Text)",
]


UI_DICT = {
    "tr": {
        "btn_back": "← Geri Dön",
        "btn_file": "Dosya Seç",
        "btn_start_proc": "İşlemi Başlat",
        "btn_processing": "⏳ İşleniyor...",
        "btn_record": "🎙️ Mikrofon Kaydını Başlat",
        "btn_record_stop": "⏹️ Kaydı Bitir ve Deşifre Et",
        "log_wait_mic": "🎙️ Mikrofon kaydı başladı. Konuşabilirsiniz...",
        "log_err_file": "❌ Lütfen önce bir ses veya video dosyası seçin!",
        "log_loading": "[1] Yapay Zeka Paketi Yükleniyor... (İlk yükleme birkaç saniye sürebilir)",
        "log_loaded": "✓ Yapay zeka paketi yüklendi.",
        "nav_home": "🏠 Ana Sayfa",
        "nav_chat": "🤖 AI Sohbet & PDF (Ekstra)",
        "nav_files": "📁 Dosyalarım",
        "nav_history": "⏱️ Geçmiş",
        "nav_settings": "⚙️ Ayarlar",
        "banner_title": "ILBER ile sınırları kaldırın.",
        "badge_tags": "%100 GİZLİ · İNTERNETSİZ · ABONELİKSİZ",
        "banner_sub": "Ses ve videolarınızı yerel yapay zeka ile metne, ders notuna, çeviriye ve akademik veriye dönüştürün.",
        "c1_t": "MIC · Aynı Dilde Deşifre",
        "c1_d": "Ses ve videolarınızı kaynak dilde zaman damgalı metne dönüştürün. Word, TXT, SRT ve VTT olarak dışa aktarın.",
        "c1_b": "MIC  Word/TXT  SRT/VTT",
        "c2_t": "EDU · Format Korumalı Belge & Doküman Çevirici",
        "c2_d": "Word (.docx), PDF (.pdf), Excel (.xlsx) ve TXT belgelerinizi sayfa düzenini, tablolarını ve formatını birebir koruyarak çevirin. İsteğe bağlı Terim Sözlüğü destekler.",
        "c2_b": "EDU  DOCX/PDF/XLSX  Terim Sözlüğü",
        "c3_t": "GLOBE · Video Çeviri Stüdyosu",
        "c3_d": "Videolarınız için standartlara uygun (42 krk/2 satır) altyazı üretin, çoklu hedef dile çapraz çevirin ve videoya gömün.",
        "c3_b": "GLOBE  Çapraz Çeviri  MP4/SRT/VTT",
        "c4_t": "FIELD · Araştırmacı Araç Kutusu",
        "c4_d": "Nitel araştırmalar için ses üzerinden konuşmacı ayrımı, saha notları ve NVivo/MAXQDA çıktısı.",
        "c4_b": "FIELD  Diarization  NVivo/MAXQDA",
        "c5_t": "EDU · Format Korumalı Belge & Doküman Çevirici",
        "c5_d": "Word (.docx), PDF (.pdf), Excel (.xlsx) ve TXT belgelerinizi sayfa düzenini, tablolarını ve formatını birebir koruyarak çevirin. İsteğe bağlı Terim Sözlüğü destekler.",
        "mic_note": "ℹ️ Bilgi: Amatör araştırmacılar MIC modülünü kullanabilir; ancak konuşmacı ayrımı ve saha notları yalnızca FIELD segmentindedir.",
        "src_lang_lbl": "Kaynak Dil (Otomatik / Elle):",
        "edu_trans_lbl": "Çeviri & Çıktı Biçimi (Çevrimdışı Motor):",
        "edu_trans_none": "Çeviri Yok (Yalnızca Kaynak Dil)",
        "edu_trans_bi": "Çift Dilli Word Tablosu (Kaynak + Çeviri Yan Yana)",
        "tgt_lang_lbl": "Hedef Dil:",
        "globe_multi_lbl": "Hedef Diller (Çoklu Seçim & Doğrudan Çapraz Çeviri):",
        "globe_burn_mp4": "Altyazıyı Videoya Göm (.mp4)",
        "field_meta_title": "Görüşme Bilgileri Formu (Saha Künyesi)",
        "field_date": "Tarih:",
        "field_loc": "Yer:",
        "field_pcode": "Katılımcı Kodu:",
        "field_interviewer": "Görüşmeci:",
        "field_spk_count": "Konuşmacı Sayısı:",
        "field_ctx": "Bağlam / Gözlem Notu:",
        "field_trans_chk": "Deşifreyi Başka Dile Çevir (Çevrimdışı Motor)",
        "field_ethics_title": "Çeviri ve Etik Kurul Onay Uyarısı (FIELD)",
        "field_ethics_msg": "Çeviri işlemi seçildi. Ses dosyası bilgisayarınızda kalır ve çeviri tamamen yerel olarak yapılır.\n\nEtik kurul onayınızın çeviri işlemine izin verdiğinden emin misiniz?",
        "title_files": "Dosyalarım", "search_placeholder": "Dosya ara...", "btn_upload": "+ Dosya Yükle", "msg_uploaded": "✔ Yüklendi!",
        "quick_access": "Hızlı Erişim", "folder_trans": "Dökümler", "folder_ups": "Yüklemeler", "folder_subs": "Altyazılar", "files_count": "dosya",
        "recent_files": "Son Dosyalar", "no_files": "ILBER_Workspace içinde dosya bulunamadı.",
        "title_history": "İşlem Geçmişi", "stat_trans": "Toplam Deşifre", "stat_subs": "Altyazı & Çeviri", "stat_storage": "Kullanılan Alan",
        "timeline_title": "Son İşlemler", "no_history": "Henüz işlem yok. Geçmişi görmek için bir dosya işleyin.",
        "act_processed": "Dosya İşlendi", "act_transcribed": "Ses Metne Dönüştürüldü", "act_extracted": "Sesten MP3 Çıkarıldı",
        "act_srt": "Altyazı Üretildi (SRT/VTT)", "act_translated": "Video Çevrildi", "file_prefix": "Dosya:",
        "title_settings": "Ayarlar", "pop_edit_title": "Düzenleme Ekranı", "pop_edit_msg": "Metni aşağıda kontrol edip düzenleyebilirsiniz (Ham Markdown işaretleri temizlenmiştir):", "pop_save": "✓ Kaydet ve Dışa Aktar", "pop_report_title": "EDU Ders Özeti ve Tekrar Soruları Kontrolü", "pop_qa_title": "Soru-Cevap / Röportaj Raporu:", "set_sub_pref": "Altyazı Tercihleri", "set_font_size": "Yazı Boyutu", "set_pos": "Ekran Konumu",
        "set_sub_style": "Altyazı Stili", "style_classic": "Klasik Beyaz", "style_yellow": "Sarı Kalın Çerçeveli", "style_black": "Siyah Kutu",
        "set_gen_pref": "Genel Tercihler", "set_def_lang": "Varsayılan Çeviri Dili",
        "set_data_mgmt": "Veri Yönetimi", "set_data_desc": "Çalışma alanındaki tüm dosyaları (Dökümler, Yüklemeler, Altyazılar) kalıcı olarak sil.",
        "set_clear_btn": "Tüm Verileri Temizle", "warn_title": "Uyarı", "warn_desc": "Üretilen tüm dosyaları silmek istediğinize emin misiniz? Bu işlem geri alınamaz.",
        "succ_title": "Başarılı", "succ_desc": "Çalışma alanı başarıyla temizlendi!",
        "chat_placeholder": "Yapay zekaya istediğinizi sorun (veya PDF/Video yükleyip içeriğini sorun)...",
        "chat_send": "Sor",
        "chat_memorize": "Dosyayı (PDF / Video / Ses) Hafızaya Al",
        "btn_open_extra": "Aç →",
        "mic_hint_idle": "Mikrofonla doğrudan ses kaydedip anında metne dönüştürebilirsiniz.",
        "mic_hint_rec": "🔴 Mikrofon kaydı sürüyor... Konuşmanız bitince butona tekrar basın.",
        "mic_hint_done": "✔ Mikrofon kaydı tamamlandı, metne dönüştürülüyor...",
        "busy_thinking": "⏳ Yapay zeka çalışıyor, lütfen bekleyin...",
        "busy_chat": "⏳ Yapay zeka sorunuzu analiz ediyor ve yanıt üretiyor, lütfen bekleyin...",
        "btn_wait_short": "⏳ Bekleyin...",
        "sec_unit": "sn",
        "log_trans_loading": "      [Çevrimdışı Çeviri] Yapay Zeka Çeviri Paketi yükleniyor...",
        "log_trans_loaded": "      ✓ Yapay Zeka Çeviri Paketi hazır.",
        "log_trans_prog": "%100 Çevrimdışı Çeviri",
    },
    "en": {
        "btn_back": "← Go Back",
        "btn_file": "Select File",
        "btn_start_proc": "Start Process",
        "btn_processing": "⏳ Processing...",
        "btn_record": "🎙️ Start Microphone Recording",
        "btn_record_stop": "⏹️ Stop & Transcribe",
        "log_wait_mic": "🎙️ Microphone recording started. You can speak now...",
        "log_err_file": "❌ Please select an audio or video file first!",
        "log_loading": "[1] Loading AI Package... (May take a few seconds)",
        "log_loaded": "✓ AI package loaded.",
        "nav_home": "🏠 Home",
        "nav_chat": "🎓 EDU · AI Tutor & PDF",
        "nav_files": "📁 My Files",
        "nav_history": "⏱️ History",
        "nav_settings": "⚙️ Settings",
        "banner_title": "Remove borders with ILBER.",
        "badge_tags": "100% OFFLINE · PRIVATE · NO SUBSCRIPTION",
        "banner_sub": "Transform your audio and video into transcripts, verified lecture notes, translations, and qualitative research data locally.",
        "c1_t": "MIC · Same-Language Transcription",
        "c1_d": "Transcribe audio or video in its source language with timestamps. Export as Word, TXT, SRT, and VTT.",
        "c1_b": "MIC  Word/TXT  SRT/VTT",
        "c2_t": "EDU · Format-Preserving Document Translator",
        "c2_d": "Translate Word (.docx), PDF (.pdf), Excel (.xlsx), and TXT documents while preserving exact layout, tables, and formatting. Optional Terminology Glossary support.",
        "c2_b": "EDU  DOCX/PDF/XLSX  Glossary Support",
        "c3_t": "GLOBE · Video Translation Studio",
        "c3_d": "Generate standard subtitles (42 chars/2 lines), cross-translate into multiple languages, and burn into MP4.",
        "c3_b": "GLOBE  Cross-Translate  MP4/SRT/VTT",
        "c4_t": "FIELD · Researcher Toolbox",
        "c4_d": "Audio-based speaker diarization, field notes, and NVivo/MAXQDA exports for qualitative researchers.",
        "c4_b": "FIELD  Diarization  NVivo/MAXQDA",
        "c5_t": "EDU · Format-Preserving Document Translator",
        "c5_d": "Translate Word (.docx), PDF (.pdf), Excel (.xlsx), and TXT documents while preserving exact layout, tables, and formatting. Optional Terminology Glossary support.",
        "mic_note": "ℹ️ Note: Amateur researchers can use MIC; however, speaker diarization and field notes are in the FIELD segment.",
        "src_lang_lbl": "Source Language (Auto / Manual):",
        "edu_trans_lbl": "Translation & Output Mode (Offline Engine):",
        "edu_trans_none": "No Translation (Source Language Only)",
        "edu_trans_bi": "Bilingual Word Table (Source + Translation Side-by-Side)",
        "tgt_lang_lbl": "Target Language:",
        "globe_multi_lbl": "Target Languages (Multi-Select & Direct Cross-Translation):",
        "globe_burn_mp4": "Burn Subtitles into Video (.mp4)",
        "field_meta_title": "Interview Metadata Form (Field Header)",
        "field_date": "Date:",
        "field_loc": "Location:",
        "field_pcode": "Participant Code:",
        "field_interviewer": "Interviewer:",
        "field_spk_count": "Speaker Count:",
        "field_ctx": "Context / Observation Note:",
        "field_trans_chk": "Translate Transcript (Offline Engine)",
        "field_ethics_title": "Translation & Ethics Board Warning (FIELD)",
        "field_ethics_msg": "Translation selected. The audio file never leaves your computer and translation is performed locally.\n\nPlease confirm that your ethics approval permits translation.",
        "title_files": "My Files", "search_placeholder": "Search files...", "btn_upload": "+ Upload File", "msg_uploaded": "✔ Uploaded!",
        "quick_access": "Quick Access", "folder_trans": "Transcripts", "folder_ups": "Uploads", "folder_subs": "Subtitles", "files_count": "files",
        "recent_files": "Recent Files", "no_files": "No files found in ILBER_Workspace.",
        "title_history": "History & Activity", "stat_trans": "Total Transcriptions", "stat_subs": "Subtitles & Translations", "stat_storage": "Storage Used",
        "timeline_title": "Recent Activity Timeline", "no_history": "No activity yet. Process a file to see your history here.",
        "act_processed": "Processed File", "act_transcribed": "Audio Transcribed to Text", "act_extracted": "Audio Extracted from Video",
        "act_srt": "Generated Subtitles (SRT/VTT)", "act_translated": "Video Translated & Subbed", "file_prefix": "File:",
        "title_settings": "Settings", "pop_edit_title": "Edit Screen", "pop_edit_msg": "Review and edit the text below (raw Markdown symbols have been cleaned):", "pop_save": "✓ Save and Export", "pop_report_title": "EDU Lecture Summary & Review Questions", "pop_qa_title": "Q&A / Interview Report:", "set_sub_pref": "Subtitle Preferences", "set_font_size": "Font Size", "set_pos": "Screen Position",
        "set_sub_style": "Subtitle Style", "style_classic": "Classic White", "style_yellow": "Yellow Outlined", "style_black": "Black Box",
        "set_gen_pref": "General Preferences", "set_def_lang": "Default Translation Language",
        "set_data_mgmt": "Data Management", "set_data_desc": "Permanently delete all workspace files (Transcripts, Uploads, Subtitles).",
        "set_clear_btn": "Clear Workspace Data", "warn_title": "Warning", "warn_desc": "Are you sure you want to delete all generated files? This cannot be undone.",
        "succ_title": "Success", "succ_desc": "Workspace cleared successfully!",
        "chat_placeholder": "Ask AI anything (or upload a PDF/Video and ask about its content)...",
        "chat_send": "Ask",
        "chat_memorize": "Memorize File (PDF / Video / Audio)",
        "btn_open_extra": "Open →",
        "mic_hint_idle": "Record directly from your microphone and transcribe into text immediately.",
        "mic_hint_rec": "🔴 Recording in progress... Click the button again when finished speaking.",
        "mic_hint_done": "✔ Recording completed, converting speech to text...",
        "busy_thinking": "⏳ AI is working, please wait...",
        "busy_chat": "⏳ AI is analyzing your question and generating a response, please wait...",
        "btn_wait_short": "⏳ Please wait...",
        "sec_unit": "s",
        "log_trans_loading": "      [Offline Translation] Loading AI Translation Package...",
        "log_trans_loaded": "      ✓ AI Translation Package ready.",
        "log_trans_prog": "100% Offline Translation",
    },
    "de": {
        "btn_back": "← Zurück",
        "btn_file": "Datei Auswählen",
        "btn_start_proc": "Prozess Starten",
        "btn_processing": "⏳ Wird bearbeitet...",
        "btn_record": "🎙️ Mikrofonaufnahme Starten",
        "btn_record_stop": "⏹️ Stopp & Transkribieren",
        "log_wait_mic": "🎙️ Mikrofonaufnahme gestartet. Sie können jetzt sprechen...",
        "log_err_file": "❌ Bitte wählen Sie zuerst eine Audio- oder Videodatei aus!",
        "log_loading": "[1] KI-Paket wird geladen... (Kann einige Sekunden dauern)",
        "log_loaded": "✓ KI-Paket geladen.",
        "nav_home": "🏠 Startseite",
        "nav_chat": "🎓 EDU · KI-Lehrer & PDF",
        "nav_files": "📁 Meine Dateien",
        "nav_history": "⏱️ Verlauf",
        "nav_settings": "⚙️ Einstellungen",
        "banner_title": "Grenzen überwinden mit ILBER.",
        "badge_tags": "100% OFFLINE · PRIVAT · OHNE ABO",
        "banner_sub": "Verwandeln Sie Audio- und Videodaten lokal in Transkripte, Vorlesungsnotizen, Übersetzungen und Forschungsdaten.",
        "c1_t": "MIC · Gleiche-Sprache-Transkription",
        "c1_d": "Transkribieren Sie Audio oder Video in der Ausgangssprache mit Zeitstempeln. Export als Word, TXT, SRT und VTT.",
        "c1_b": "MIC  Word/TXT  SRT/VTT",
        "c2_t": "EDU · Formatgetreuer Dokumentübersetzer",
        "c2_d": "Übersetzen Sie Word (.docx), PDF (.pdf), Excel (.xlsx) und TXT-Dokumente unter vollständiger Beibehaltung von Layout, Tabellen und Formatierung. Optional mit Fachglossar.",
        "c2_b": "EDU  DOCX/PDF/XLSX  Glossar (Optional)",
        "c3_t": "GLOBE · Video-Übersetzungsstudio",
        "c3_d": "Erstellen Sie Standard-Untertitel (42 Zeichen/2 Zeilen), Direktübersetzung in mehrere Sprachen & MP4-Export.",
        "c3_b": "GLOBE  Direktübersetzung  MP4/SRT/VTT",
        "c4_t": "FIELD · Forscher-Werkzeugkasten",
        "c4_d": "Audio-Sprechererkennung, Feldnotizen und NVivo/MAXQDA-Export für qualitative Forschung.",
        "c4_b": "FIELD  Diarisierung  NVivo/MAXQDA",
        "c5_t": "EDU · Formatgetreuer Dokumentübersetzer",
        "c5_d": "Übersetzen Sie Word (.docx), PDF (.pdf), Excel (.xlsx) und TXT-Dokumente unter vollständiger Beibehaltung von Layout, Tabellen und Formatierung. Optional mit Fachglossar.",
        "mic_note": "ℹ️ Hinweis: Amateurforscher können MIC nutzen; Sprechererkennung und Feldnotizen befinden sich jedoch im Segment FIELD.",
        "src_lang_lbl": "Ausgangssprache (Auto / Manuell):",
        "edu_trans_lbl": "Übersetzung & Ausgabeformat (Offline-Engine):",
        "edu_trans_none": "Keine Übersetzung (Nur Ausgangssprache)",
        "edu_trans_bi": "Zweisprachige Word-Tabelle (Quelle + Übersetzung)",
        "tgt_lang_lbl": "Zielsprache:",
        "globe_multi_lbl": "Zielsprachen (Mehrfachauswahl & Direktübersetzung):",
        "globe_burn_mp4": "Untertitel in Video einbrennen (.mp4)",
        "field_meta_title": "Interview-Metadatenformular",
        "field_date": "Datum:",
        "field_loc": "Ort:",
        "field_pcode": "Teilnehmercode:",
        "field_interviewer": "Interviewer:",
        "field_spk_count": "Sprecheranzahl:",
        "field_ctx": "Kontext / Beobachtungsnotiz:",
        "field_trans_chk": "Transkript übersetzen (Offline-Engine)",
        "field_ethics_title": "Ethikkommission-Hinweis zur Übersetzung (FIELD)",
        "field_ethics_msg": "Übersetzung ausgewählt. Die Audiodatei verbleibt auf Ihrem Computer und die Übersetzung erfolgt vollständig lokal.\n\nBitte bestätigen Sie, dass Ihr Ethikvotum dies erlaubt.",
        "title_files": "Meine Dateien", "search_placeholder": "Dateien suchen...", "btn_upload": "+ Datei Hochladen", "msg_uploaded": "✔ Hochgeladen!",
        "quick_access": "Schnellzugriff", "folder_trans": "Transkripte", "folder_ups": "Uploads", "folder_subs": "Untertitel", "files_count": "dateien",
        "recent_files": "Letzte Dateien", "no_files": "Keine Dateien in ILBER_Workspace gefunden.",
        "title_history": "Verlauf & Aktivität", "stat_trans": "Gesamte Transkripte", "stat_subs": "Untertitel & Übersetzungen", "stat_storage": "Speicherplatz",
        "timeline_title": "Letzte Aktivitäten", "no_history": "Noch keine Aktivität. Verarbeiten Sie eine Datei.",
        "act_processed": "Datei Verarbeitet", "act_transcribed": "Audio zu Text transkribiert", "act_extracted": "Audio aus Video extrahiert",
        "act_srt": "Untertitel erstellt (SRT/VTT)", "act_translated": "Video übersetzt", "file_prefix": "Datei:",
        "title_settings": "Einstellungen", "pop_edit_title": "Bearbeitungsbildschirm", "pop_edit_msg": "Bitte prüfen und bearbeiten Sie den Text unten (Markdown-Symbole wurden entfernt):", "pop_save": "✓ Speichern und Exportieren", "pop_report_title": "EDU Vorlesungszusammenfassung & Fragen", "pop_qa_title": "Q&A / Interviewbericht:", "set_sub_pref": "Untertitel Präferenzen", "set_font_size": "Schriftgröße", "set_pos": "Bildschirmposition",
        "set_sub_style": "Untertitel Stil", "style_classic": "Klassisch Weiß", "style_yellow": "Gelb Umrandet", "style_black": "Schwarze Box",
        "set_gen_pref": "Allgemeine Einstellungen", "set_def_lang": "Standard Übersetzungssprache",
        "set_data_mgmt": "Datenverwaltung", "set_data_desc": "Löschen Sie alle Arbeitsbereichsdateien dauerhaft.",
        "set_clear_btn": "Arbeitsbereich bereinigen", "warn_title": "Warnung", "warn_desc": "Sind Sie sicher, dass Sie alle Dateien löschen möchten? Dies kann nicht rückgängig gemacht werden.",
        "succ_title": "Erfolg", "succ_desc": "Arbeitsbereich erfolgreich bereinigt!",
        "chat_placeholder": "Fragen Sie die KI etwas (oder laden Sie ein PDF/Video hoch)...",
        "chat_send": "Fragen",
        "chat_memorize": "Datei (PDF / Video / Audio) ins Gedächtnis laden",
        "btn_open_extra": "Öffnen →",
        "mic_hint_idle": "Direkt über das Mikrofon aufnehmen und sofort in Text umwandeln.",
        "mic_hint_rec": "🔴 Aufnahme läuft... Klicken Sie erneut auf die Schaltfläche, wenn Sie fertig sind.",
        "mic_hint_done": "✔ Aufnahme abgeschlossen, Sprache wird in Text umgewandelt...",
        "busy_thinking": "⏳ KI arbeitet, bitte warten...",
        "busy_chat": "⏳ KI analysiert Ihre Frage und erstellt eine Antwort, bitte warten...",
        "btn_wait_short": "⏳ Bitte warten...",
        "sec_unit": "Sek.",
        "log_trans_loading": "      [Offline-Übersetzung] KI-Übersetzungspaket wird geladen...",
        "log_trans_loaded": "      ✓ KI-Übersetzungspaket bereit.",
        "log_trans_prog": "100% Offline-Übersetzung",
    },
}

# --- CONSTANTS ---
PURPLE      = ("#00BCD4", "#00E5FF")
PURPLE_DARK = ("#00838F", "#008B99")
BG          = ("#EBF4F6", "#161324")
CARD        = ("#FFFFFF", "#1E1A33")
CARD2       = ("#D1E8EB", "#262142")
PURPLE_BG   = ("#E0F7FA", "#1A162B")
TEXT        = ("#1F2937", "#FFFFFF")
TEXT_MUTED  = ("#64748B", "#8B949E")
RED         = ("#EF4444", "#FF1744")

ctk.set_appearance_mode("light")


class OutputRedirector:
    def __init__(self, app, logbox):
        self.app = app
        self.logbox = logbox

    def write(self, text):
        if getattr(self.app, "cancel_requested", False):
            raise RuntimeError("__USER_CANCELLED__")
        s = text.strip()
        if not s:
            return
        # Hide internal technical CLI messages from third-party libraries
        if "Detecting language using up to the first" in s or "Use `--language`" in s:
            return
        if s.startswith("Detected language:"):
            lang_name = s.split(":", 1)[-1].strip()
            s = self.app.L(
                f"      ✓ Algılanan Konuşma Dili: {lang_name}",
                f"      ✓ Detected Speech Language: {lang_name}",
                f"      ✓ Erkannte Sprache: {lang_name}",
            )
        self.app.after(0, lambda t=s: self.app._log(self.logbox, t))

    def flush(self):
        if getattr(self.app, "cancel_requested", False):
            raise RuntimeError("__USER_CANCELLED__")


class ILBERTechV2(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.model = None
        self.trans_model = None
        self.trans_tokenizer = None
        self.llm_instance = None
        self.is_processing = False
        self.cancel_requested = False
        self._current_run_id = 0
        self.is_recording = False
        self.ui_lang = "en"
        self.audio_frames = []

        self.pages = {}
        self.log_boxes = {}
        self.page_controls = {}
        self.user_settings = {
            "sub_size": "20",
            "sub_style": "Classic White",
            "sub_pos": "Bottom",
            "def_lang": "English",
        }

        self.license_mgr = SupabaseLicenseManager()
        if self.license_mgr.try_auto_login(sync_cloud=False):
            self.on_login_success()
            threading.Thread(target=self._bg_sync_license_on_startup, daemon=True).start()
        else:
            self._init_login()

    def _bg_sync_license_on_startup(self):
        """Non-blocking cloud license & plan sync after the main UI window is already open."""
        try:
            ok = self.license_mgr.sync_session_in_background()
            if ok:
                self.after(0, self.refresh_license_ui)
            elif not self.license_mgr.is_authenticated:
                self.after(0, self._init_login)
        except Exception:
            pass

    def t(self, key):
        return UI_DICT.get(self.ui_lang, UI_DICT["en"]).get(key, key)

    def L(self, tr_text, en_text, de_text):
        if self.ui_lang == "en":
            return en_text
        if self.ui_lang == "de":
            return de_text
        return tr_text

    def _init_login(self):
        self.withdraw()
        if hasattr(self, "login_window") and self.login_window.winfo_exists():
            self.login_window.destroy()

        self.login_window = ctk.CTkToplevel(self)
        self.login_window.title(self.L(
            "ILBERTech Platform v2.0 — Giriş & Ücretsiz Deneme",
            "ILBERTech Platform v2.0 — Sign In & Free Trial",
            "ILBERTech Platform v2.0 — Anmeldung & Testversion",
        ))
        self.login_window.geometry("940x590")
        self.login_window.resizable(False, False)
        self.login_window.configure(fg_color="#FFFFFF")
        self.login_window.protocol("WM_DELETE_WINDOW", self.destroy)
        try:
            ico_p = resource_path("ilber_logo.ico")
            if os.path.exists(ico_p):
                self.login_window.iconbitmap(ico_p)
        except Exception:
            pass

        # Left Branding Panel
        left_panel = ctk.CTkFrame(self.login_window, width=410, fg_color="#141B2D", corner_radius=0)
        left_panel.pack(side="left", fill="y")
        left_panel.pack_propagate(False)

        brand_box = ctk.CTkFrame(left_panel, fg_color="transparent")
        brand_box.pack(expand=True, padx=36)

        logo_row = ctk.CTkFrame(brand_box, fg_color="transparent")
        logo_row.pack(anchor="w", pady=(0, 14))
        try:
            logo_p = resource_path("ilber_logo.jpg")
            if os.path.exists(logo_p):
                logo_img = ctk.CTkImage(light_image=Image.open(logo_p), size=(54, 54))
                lbl_l = ctk.CTkLabel(logo_row, image=logo_img, text="")
                lbl_l._img_ref = logo_img
                lbl_l.pack(side="left", padx=(0, 14))
        except Exception:
            pass
        ctk.CTkLabel(logo_row, text="ILBERTech", font=ctk.CTkFont(family="Segoe UI", size=30, weight="bold"), text_color="#FFFFFF").pack(side="left")

        ctk.CTkLabel(
            brand_box,
            text="MIC · EDU · GLOBE · FIELD",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color="#38BDF8",
        ).pack(anchor="w", pady=(0, 14))

        ctk.CTkLabel(
            brand_box,
            text=self.L(
                "• %100 Offline & KVKK Uyumlu Yapay Zeka\n"
                "• 🎙️ MIC ($9.99): Ses Kaydı & Deşifre\n"
                "• 🎓 EDU ($19.99): MIC + Belge Çevirici & OCR\n"
                "• 🌐 GLOBE ($39.99): MIC + EDU + Video Altyazı Stüdyosu\n"
                "• 🔬 FIELD PRO ($79.99): Tüm 4 Modül + Saha Analizi\n\n"
                "🎁 Yeni hesaplara tüm modüllerde geçerli\n"
                "5 ÜCRETSİZ DENEME HAKKI tanımlanır!",
                "• 100% Offline & Privacy-First Local AI\n"
                "• 🎙️ MIC ($9.99): Audio Recording & Transcription\n"
                "• 🎓 EDU ($19.99): MIC + Format-Preserving Document Translator & OCR\n"
                "• 🌐 GLOBE ($39.99): MIC + EDU + Video Subtitle Studio\n"
                "• 🔬 FIELD PRO ($79.99): All 4 Modules + Qualitative Research\n\n"
                "🎁 New accounts automatically receive\n"
                "5 FREE TRIAL CREDITS across all modules!",
                "• 100% Offline & Datenschutz-KI\n"
                "• 🎙️ MIC ($9.99): Aufnahme & Transkription\n"
                "• 🎓 EDU ($19.99): MIC + Dokumentübersetzer & OCR\n"
                "• 🌐 GLOBE ($39.99): MIC + EDU + Video-Untertitelstudio\n"
                "• 🔬 FIELD PRO ($79.99): Alle 4 Module + Feldforschung\n\n"
                "🎁 Neue Konten erhalten 5 KOSTENLOSE TESTVORGÄNGE!",
            ),
            font=ctk.CTkFont(family="Segoe UI", size=13),
            text_color="#CBD5E1",
            justify="left",
        ).pack(anchor="w", pady=(0, 20))

        dev_box = ctk.CTkFrame(brand_box, fg_color="#1E293B", corner_radius=10)
        dev_box.pack(fill="x", pady=(8, 0))
        ctk.CTkLabel(
            dev_box,
            text=f"💻 {self.L('Cihaz Kodu', 'Device Code', 'Gerätecode')}: {self.license_mgr.device_code}",
            font=ctk.CTkFont(family="Consolas", size=13, weight="bold"),
            text_color="#F8FAFC",
        ).pack(padx=14, pady=8)

        # Right Form Panel
        right_panel = ctk.CTkFrame(self.login_window, fg_color="#FFFFFF", corner_radius=0)
        right_panel.pack(side="right", fill="both", expand=True)

        form = ctk.CTkFrame(right_panel, fg_color="transparent", width=390)
        form.place(relx=0.5, rely=0.5, anchor="center")

        ctk.CTkLabel(
            form,
            text=self.L("Hesabınıza Giriş Yapın", "Sign In to Your Account", "Bei Ihrem Konto anmelden"),
            font=ctk.CTkFont(family="Segoe UI", size=26, weight="bold"),
            text_color="#0F172A",
        ).pack(anchor="w")
        ctk.CTkLabel(
            form,
            text=self.L(
                "Supabase bulut hesabınızla giriş yapın veya ücretsiz deneme başlatın.",
                "Sign in with your cloud account or start a free trial.",
                "Melden Sie sich mit Ihrem Cloud-Konto an oder starten Sie eine Testversion.",
            ),
            font=ctk.CTkFont(family="Segoe UI", size=13),
            text_color="#64748B",
        ).pack(anchor="w", pady=(4, 18))

        ctk.CTkLabel(
            form,
            text=self.L("E-posta Adresi", "Email Address", "E-Mail-Adresse"),
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color="#1E293B",
        ).pack(anchor="w", pady=(0, 4))
        email_entry = ctk.CTkEntry(form, placeholder_text="name@company.com", width=390, height=42, corner_radius=8, border_color="#CBD5E1", fg_color="#F8FAFC", text_color="#0F172A")
        email_entry.pack(pady=(0, 12))

        ctk.CTkLabel(
            form,
            text=self.L("Şifre", "Password", "Passwort"),
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color="#1E293B",
        ).pack(anchor="w", pady=(0, 4))
        pass_entry = ctk.CTkEntry(form, placeholder_text="••••••••", show="•", width=390, height=42, corner_radius=8, border_color="#CBD5E1", fg_color="#F8FAFC", text_color="#0F172A")
        pass_entry.pack(pady=(0, 8))

        status_lbl = ctk.CTkLabel(form, text="", font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"), text_color="#DC2626", wraplength=380, justify="left")
        status_lbl.pack(anchor="w", pady=(2, 8))

        def do_signin(*args):
            em = email_entry.get().strip()
            pw = pass_entry.get().strip()
            status_lbl.configure(text=self.L("🔄 Sunucuya bağlanılıyor...", "🔄 Connecting to server...", "🔄 Verbindung wird hergestellt..."), text_color="#2563EB")
            self.login_window.update_idletasks()
            ok, msg = self.license_mgr.sign_in(em, pw)
            if ok:
                self.on_login_success()
            else:
                status_lbl.configure(text=f"❌ {msg}", text_color="#DC2626")

        def do_signup():
            em = email_entry.get().strip()
            pw = pass_entry.get().strip()
            status_lbl.configure(text=self.L("🔄 Hesabınız oluşturuluyor...", "🔄 Creating account & assigning 5 Trial Credits...", "🔄 Konto wird erstellt..."), text_color="#2563EB")
            self.login_window.update_idletasks()
            ok, msg = self.license_mgr.sign_up_trial(em, pw)
            if ok:
                messagebox.showinfo(self.L("Hoş Geldiniz!", "Welcome!", "Willkommen!"), msg)
                self.on_login_success()
            else:
                status_lbl.configure(text=f"❌ {msg}", text_color="#DC2626")

        def do_instant_trial():
            self.license_mgr.activate_package_with_key("ILBER-TRIAL")
            self.on_login_success()

        btn_login = ctk.CTkButton(
            form,
            text=self.L("🔐 Giriş Yap", "🔐 Sign In", "🔐 Anmelden"),
            width=390,
            height=44,
            corner_radius=8,
            fg_color="#2563EB",
            hover_color="#1D4ED8",
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            command=do_signin,
        )
        btn_login.pack(pady=(4, 8))

        btn_signup = ctk.CTkButton(
            form,
            text=self.L("✨ Yeni Hesap Oluştur & 5 Ücretsiz Deneme Başlat", "✨ Create Account & Start 5 Free Trials", "✨ Konto erstellen & 5 Testvorgänge starten"),
            width=390,
            height=42,
            corner_radius=8,
            fg_color="#10B981",
            hover_color="#059669",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            command=do_signup,
        )
        btn_signup.pack(pady=(0, 8))

        btn_demo = ctk.CTkButton(
            form,
            text=self.L("⚡ Hızlı Test Girişi (Ücretsiz Deneme — 5 Hak)", "⚡ Quick Trial Access (5 Free Credits)", "⚡ Schneller Testzugang (5 Credits)"),
            width=390,
            height=38,
            corner_radius=8,
            fg_color="#F1F5F9",
            text_color="#334155",
            hover_color="#E2E8F0",
            border_width=1,
            border_color="#CBD5E1",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            command=do_instant_trial,
        )
        btn_demo.pack(pady=(2, 0))

        pass_entry.bind("<Return>", do_signin)

    def on_login_success(self):
        if hasattr(self, "login_window") and self.login_window.winfo_exists():
            self.login_window.destroy()
        try:
            if os.name == "nt":
                import ctypes
                ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("ILBERTech.Platform.v2")
            ico_p = resource_path("ilber_logo.ico")
            if os.path.exists(ico_p):
                self.iconbitmap(ico_p)
                self.after(200, lambda: self.iconbitmap(ico_p))
        except Exception:
            pass
        self.title("ILBERTech Platform v2.1 — MIC · EDU · GLOBE · FIELD")
        self.geometry("1320x840")
        self.resizable(True, True)
        self.configure(fg_color=BG)
        for widget in self.winfo_children():
            widget.destroy()
        self.pages.clear()
        self.log_boxes.clear()
        self.page_controls.clear()
        self._build_dashboard()
        self.deiconify()

    def refresh_license_ui(self):
        """Refresh sidebar package badge and home module lock/unlock badges dynamically."""
        try:
            if hasattr(self, "sidebar_plan_lbl") and self.sidebar_plan_lbl.winfo_exists():
                self.sidebar_plan_lbl.configure(text=self.license_mgr.get_plan_display_name(self.ui_lang))
            if hasattr(self, "sidebar_email_lbl") and self.sidebar_email_lbl.winfo_exists():
                self.sidebar_email_lbl.configure(text=self.license_mgr.session.get("email") or "Kullanıcı")
            if "home" in self.pages and self.pages["home"].winfo_exists():
                self._build_home()
        except Exception:
            pass

    def open_package_upgrade_modal(self, locked_pid=None):
        """Open the 4-Tier Package & Email-Based Supabase Cloud Authority Modal ($9.99 / $19.99 / $39.99 / $79.99)."""
        if hasattr(self, "_pkg_modal") and self._pkg_modal.winfo_exists():
            self._pkg_modal.lift()
            self._pkg_modal.focus_force()
            return

        modal = ctk.CTkToplevel(self)
        self._pkg_modal = modal
        modal.title(self.L(
            "ILBERTech — Paketler & Hesap Yetkisi",
            "ILBERTech — Packages & Account Access",
            "ILBERTech — Pakete & Kontozugriff",
        ))
        modal.geometry("920x610")
        modal.resizable(False, False)
        modal.configure(fg_color="#F8FAFC")
        modal.attributes("-topmost", True)
        try:
            ico_p = resource_path("ilber_logo.ico")
            if os.path.exists(ico_p):
                modal.iconbitmap(ico_p)
        except Exception:
            pass

        user_email = self.license_mgr.session.get("email") or "deneme@ilbertech.com"

        # Top Header Banner
        hdr = ctk.CTkFrame(modal, fg_color="#141B2D", corner_radius=0, height=84)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)

        hdr_left = ctk.CTkFrame(hdr, fg_color="transparent")
        hdr_left.pack(side="left", padx=28, pady=14)
        ctk.CTkLabel(
            hdr_left,
            text=self.L("💎 ILBERTech Modüler Paket & Yetki Merkezi", "💎 ILBERTech Modular Package & Access Center", "💎 ILBERTech Paket- & Zugriffszentrum"),
            font=ctk.CTkFont(family="Segoe UI", size=21, weight="bold"),
            text_color="#FFFFFF",
        ).pack(anchor="w")
        ctk.CTkLabel(
            hdr_left,
            text=f"📧 {self.L('Hesap:', 'Account:', 'Konto:')} {user_email}   |   {self.L('Aktif Yetki:', 'Active Plan:', 'Aktives Paket:')} {self.license_mgr.get_plan_display_name(self.ui_lang)}",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color="#38BDF8",
        ).pack(anchor="w", pady=(3, 0))

        # Locked Module Warning Banner (if triggered by clicking a locked module)
        if locked_pid:
            req_map = {
                "tool_ses_kaydet": ("Paket 1: MIC ($9.99)", "#9333EA"),
                "tool_ai_ozet": ("Paket 2: EDU ($19.99)", "#2563EB"),
                "tool_video_cevir": ("Paket 3: GLOBE ($39.99)", "#10B981"),
                "tool_ai_rapor": ("Paket 4: FIELD PRO ($79.99)", "#EF4444"),
            }
            req_name, _ = req_map.get(locked_pid, ("Üst Paket", "#DC2626"))
            warn_bar = ctk.CTkFrame(modal, fg_color="#FEF2F2", border_width=1, border_color="#EF4444", corner_radius=12)
            warn_bar.pack(fill="x", padx=24, pady=(12, 0))
            if self.license_mgr.session.get("plan") == "trial" and int(self.license_mgr.session.get("trial_remaining", 0)) <= 0:
                w_txt = self.L(
                    "🔒 5 Ücretsiz Deneme Hakkınız tamamlandı! Çalışmaya kesintisiz devam etmek için e-posta hesabınıza tanımlı paketi yükseltin.",
                    "🔒 Your 5 Free Trial operations have been used! Upgrade the package assigned to your email account to continue.",
                    "🔒 Ihre 5 kostenlosen Testvorgänge sind aufgebraucht! Bitte führen Sie ein Upgrade durch.",
                )
            else:
                w_txt = self.L(
                    f"🔒 Bu modül mevcut paketinizde kilitlidir. Açmak için hesabınızın en az '{req_name}' yetkisine sahip olması gerekir.",
                    f"🔒 This module is locked on your current plan. Your account needs '{req_name}' or higher.",
                    f"🔒 Dieses Modul ist in Ihrem aktuellen Tarif gesperrt. '{req_name}' erforderlich.",
                )
            ctk.CTkLabel(warn_bar, text=w_txt, font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"), text_color="#991B1B", wraplength=840, justify="left").pack(padx=16, pady=10, anchor="w")

        # 4 Package Cards Grid
        cards_frame = ctk.CTkFrame(modal, fg_color="transparent")
        cards_frame.pack(fill="x", padx=20, pady=(16, 12))
        for i in range(4):
            cards_frame.grid_columnconfigure(i, weight=1)

        pkg_specs = [
            (
                "paket_1",
                self.L("🎙️ PAKET 1: MIC", "🎙️ PACKAGE 1: MIC", "🎙️ PAKET 1: MIC"),
                "$9.99",
                "#9333EA",
                "#F4EBFF",
                self.L(
                    "✔ MIC (Ses Kaydı & Deşifre)\n🔒 EDU Kilitli\n🔒 GLOBE Kilitli\n🔒 FIELD Kilitli",
                    "✔ MIC (Audio & Transcription)\n🔒 EDU Locked\n🔒 GLOBE Locked\n🔒 FIELD Locked",
                    "✔ MIC (Aufnahme & Transkription)\n🔒 EDU Gesperrt\n🔒 GLOBE Gesperrt\n🔒 FIELD Gesperrt",
                ),
            ),
            (
                "paket_2",
                self.L("🎓 PAKET 2: EDU", "🎓 PACKAGE 2: EDU", "🎓 PAKET 2: EDU"),
                "$19.99",
                "#2563EB",
                "#EBF3FF",
                self.L(
                    "✔ MIC (Ses & Deşifre)\n✔ EDU (Belge Çevirici & OCR)\n🔒 GLOBE Kilitli\n🔒 FIELD Kilitli",
                    "✔ MIC (Audio & Transcription)\n✔ EDU (Document Translator & OCR)\n🔒 GLOBE Locked\n🔒 FIELD Locked",
                    "✔ MIC (Aufnahme & Transkription)\n✔ EDU (Dokumentübersetzer & OCR)\n🔒 GLOBE Gesperrt\n🔒 FIELD Gesperrt",
                ),
            ),
            (
                "paket_3",
                self.L("🌐 PAKET 3: GLOBE", "🌐 PACKAGE 3: GLOBE", "🌐 PAKET 3: GLOBE"),
                "$39.99",
                "#10B981",
                "#E6FBF2",
                self.L(
                    "✔ MIC (Ses & Deşifre)\n✔ EDU (Belge Çevirici & OCR)\n✔ GLOBE (Video & Altyazı)\n🔒 FIELD Kilitli",
                    "✔ MIC (Audio & Transcription)\n✔ EDU (Document Translator & OCR)\n✔ GLOBE (Video & Subtitle Studio)\n🔒 FIELD Locked",
                    "✔ MIC (Aufnahme & Transkription)\n✔ EDU (Dokumentübersetzer & OCR)\n✔ GLOBE (Video & Untertitel)\n🔒 FIELD Gesperrt",
                ),
            ),
            (
                "paket_4",
                self.L("🔬 PAKET 4: FIELD PRO", "🔬 PACKAGE 4: FIELD PRO", "🔬 PAKET 4: FIELD PRO"),
                "$79.99",
                "#EF4444",
                "#FEEBEB",
                self.L(
                    "✔ MIC (Ses & Deşifre)\n✔ EDU (Belge Çevirici & OCR)\n✔ GLOBE (Video & Altyazı)\n✔ FIELD (Saha & Mülakat)",
                    "✔ MIC (Audio & Transcription)\n✔ EDU (Document Translator & OCR)\n✔ GLOBE (Video & Subtitle Studio)\n✔ FIELD (Qualitative Research)",
                    "✔ MIC (Aufnahme & Transkription)\n✔ EDU (Dokumentübersetzer & OCR)\n✔ GLOBE (Video & Untertitel)\n✔ FIELD (Feldforschung)",
                ),
            ),
        ]

        def on_buy_package(pcode, ptitle, pprice):
            messagebox.showinfo(
                f"{ptitle} ({pprice})",
                self.L(
                    f"Kayıtlı E-Posta Adresiniz: {user_email}\n"
                    f"Seçilen Paket: {ptitle} ({pprice})\n\n"
                    f"Ödeme işleminiz tamamlandığında (veya yönetici tarafından '{user_email}' adresinize Supabase üzerinden yetki tanımlandığında) aşağıdaki '🔄 Yetkimi / Paketimi Güncelle' butonuna basmanız yeterlidir!",
                    f"Registered Email: {user_email}\n"
                    f"Selected Plan: {ptitle} ({pprice})\n\n"
                    f"Once your payment is completed (or after your plan is assigned on Supabase), simply click '🔄 Check & Activate My Package' below!",
                    f"Registrierte E-Mail: {user_email}\n"
                    f"Ausgewähltes Paket: {ptitle} ({pprice})\n\n"
                    f"Klicken Sie nach der Freischaltung unten auf '🔄 Mein Paket aktualisieren'!",
                ),
            )

        cur_plan = self.license_mgr.session.get("plan", "trial")
        for idx, (pcode, ptitle, pprice, pcol, pbg, pfeatures) in enumerate(pkg_specs):
            is_cur = (cur_plan == pcode)
            pc = ctk.CTkFrame(
                cards_frame,
                fg_color=pbg if is_cur else "#FFFFFF",
                border_width=2 if is_cur else 1,
                border_color=pcol if is_cur else "#CBD5E1",
                corner_radius=14,
            )
            pc.grid(row=0, column=idx, padx=6, sticky="nsew")
            ctk.CTkLabel(pc, text=ptitle, font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"), text_color=pcol).pack(anchor="w", padx=14, pady=(14, 2))
            ctk.CTkLabel(pc, text=pprice, font=ctk.CTkFont(family="Segoe UI", size=26, weight="bold"), text_color="#0F172A").pack(anchor="w", padx=14, pady=(0, 8))
            ctk.CTkLabel(pc, text=pfeatures, font=ctk.CTkFont(family="Segoe UI", size=12), text_color="#334155", justify="left").pack(anchor="w", padx=14, pady=(0, 12))

            if is_cur:
                ctk.CTkButton(
                    pc,
                    text=self.L("✓ AKTİF PAKETİNİZ", "✓ ACTIVE PLAN", "✓ AKTIVES PAKET"),
                    height=36,
                    state="disabled",
                    fg_color=pcol,
                    text_color_disabled="#FFFFFF",
                    font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
                ).pack(fill="x", padx=12, pady=(0, 14))
            else:
                ctk.CTkButton(
                    pc,
                    text=self.L(f"💳 Satın Al ({pprice})", f"💳 Upgrade ({pprice})", f"💳 Kaufen ({pprice})"),
                    height=36,
                    fg_color=pcol,
                    hover_color="#1E293B",
                    font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
                    command=lambda c=pcode, t=ptitle, pr=pprice: on_buy_package(c, t, pr),
                ).pack(fill="x", padx=12, pady=(0, 14))

        # Supabase Email Authority Sync Box
        act_card = ctk.CTkFrame(modal, fg_color="#FFFFFF", border_width=1, border_color="#CBD5E1", corner_radius=14)
        act_card.pack(fill="x", padx=24, pady=(8, 14))

        ctk.CTkLabel(
            act_card,
            text=self.L(
                f"☁️ E-Posta Bazlı Otomatik Yetki Sistemi ({user_email})",
                f"☁️ Email-Based Cloud License System ({user_email})",
                f"☁️ E-Mail-basiertes Cloud-Lizenzsystem ({user_email})",
            ),
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            text_color="#0F172A",
        ).pack(anchor="w", padx=20, pady=(14, 4))

        ctk.CTkLabel(
            act_card,
            text=self.L(
                "E-posta adresinize Supabase üzerinden paket yetkisi tanımlandığında (veya LemonSqueezy ödemeniz onaylandığında) hiçbir lisans kodu girmenize gerek yoktur. Aşağıdaki butona tıkladığınız anda paketiniz otomatik olarak açılır:",
                "Once your package is assigned to your email on Supabase (or after LemonSqueezy checkout), no license key is needed. Click the button below to unlock your modules automatically:",
                "Sobald Ihr Paket Ihrer E-Mail zugewiesen wurde, klicken Sie unten, um Ihre Module sofort freizuschalten:",
            ),
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color="#475569",
            wraplength=830,
            justify="left",
        ).pack(anchor="w", padx=20, pady=(0, 10))

        msg_lbl = ctk.CTkLabel(act_card, text="", font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"), text_color="#059669")
        msg_lbl.pack(anchor="w", padx=20, pady=(0, 6))

        def sync_from_supabase():
            msg_lbl.configure(
                text=self.L(
                    f"🔄 '{user_email}' için Supabase bulut yetkisi kontrol ediliyor...",
                    f"🔄 Checking Supabase cloud license for '{user_email}'...",
                    f"🔄 Cloud-Lizenz für '{user_email}' wird geprüft...",
                ),
                text_color="#2563EB",
            )
            modal.update_idletasks()
            ok, msg = self.license_mgr.refresh_from_cloud()
            if ok:
                msg_lbl.configure(text=f"✔ {msg}", text_color="#059669")
                self.refresh_license_ui()
            else:
                msg_lbl.configure(text=f"⚠️ {msg}", text_color="#D97706")

        ctk.CTkButton(
            act_card,
            text=self.L(
                "🔄 Yetkimi / Paketimi Güncelle (Supabase Buluttan Sorgula)",
                "🔄 Check & Activate My Package (Sync from Supabase Cloud)",
                "🔄 Mein Paket aktualisieren (Supabase Cloud-Sync)",
            ),
            height=44,
            fg_color="#10B981",
            hover_color="#059669",
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            command=sync_from_supabase,
        ).pack(fill="x", padx=20, pady=(0, 16))


    def _get_theme_defs(self):
        return {
            "1 · Executive Slate": {
                "sidebar_bg": "#0B0F19",
                "sidebar_card": "#161F33",
                "sidebar_active": "#1E293B",
                "canvas_bg": "#F1F5F9",
                "hero_bg": "#0F172A",
                "hero_border": "#1E293B",
                "hero_pill_bg": "#1E293B",
                "hero_pill_fg": "#38BDF8",
                "hero_title": "#FFFFFF",
                "hero_sub": "#94A3B8",
                "wave_cols": ("#38BDF8", "#6366F1", "#818CF8"),
                "card_bg": "#FFFFFF",
                "card_border": "#CBD5E1",
                "card_shadow": "#E2E8F0",
                "card_title": "#0F172A",
                "card_desc": "#475569",
                "tag_border": "#E2E8F0",
                "axes": {
                    "tool_ses_kaydet": ("MIC", "#4F46E5", "#EEF2FF", "#4338CA"),
                    "tool_ai_ozet": ("EDU", "#0284C7", "#E0F2FE", "#0369A1"),
                    "tool_video_cevir": ("GLOBE", "#059669", "#D1FAE5", "#047857"),
                    "tool_ai_rapor": ("FIELD", "#D97706", "#FEF3C7", "#B45309"),
                },
            },
            "2 · Obsidian Dark": {
                "sidebar_bg": "#070A12",
                "sidebar_card": "#111827",
                "sidebar_active": "#1E293B",
                "canvas_bg": "#0B0F19",
                "hero_bg": "#111827",
                "hero_border": "#1F2937",
                "hero_pill_bg": "#1E1B4B",
                "hero_pill_fg": "#818CF8",
                "hero_title": "#F8FAFC",
                "hero_sub": "#94A3B8",
                "wave_cols": ("#6366F1", "#38BDF8", "#10B981"),
                "card_bg": "#111827",
                "card_border": "#1F2937",
                "card_shadow": "#070A12",
                "card_title": "#F8FAFC",
                "card_desc": "#94A3B8",
                "tag_border": "#334155",
                "axes": {
                    "tool_ses_kaydet": ("MIC", "#6366F1", "#1E1B4B", "#A5B4FC"),
                    "tool_ai_ozet": ("EDU", "#0EA5E9", "#0C4A6E", "#7DD3FC"),
                    "tool_video_cevir": ("GLOBE", "#10B981", "#064E3B", "#6EE7B7"),
                    "tool_ai_rapor": ("FIELD", "#F59E0B", "#451A03", "#FCD34D"),
                },
            },
            "3 · Swiss Minimal": {
                "sidebar_bg": "#18181B",
                "sidebar_card": "#27272A",
                "sidebar_active": "#27272A",
                "canvas_bg": "#F4F4F5",
                "hero_bg": "#FFFFFF",
                "hero_border": "#E4E4E7",
                "hero_pill_bg": "#18181B",
                "hero_pill_fg": "#FAFAFA",
                "hero_title": "#09090B",
                "hero_sub": "#52525B",
                "wave_cols": ("#18181B", "#52525B", "#71717A"),
                "card_bg": "#FFFFFF",
                "card_border": "#E4E4E7",
                "card_shadow": "#E4E4E7",
                "card_title": "#09090B",
                "card_desc": "#52525B",
                "tag_border": "#D4D4D8",
                "axes": {
                    "tool_ses_kaydet": ("MIC", "#18181B", "#F4F4F5", "#18181B"),
                    "tool_ai_ozet": ("EDU", "#27272A", "#F4F4F5", "#27272A"),
                    "tool_video_cevir": ("GLOBE", "#18181B", "#F4F4F5", "#18181B"),
                    "tool_ai_rapor": ("FIELD", "#27272A", "#F4F4F5", "#27272A"),
                },
            },
            "4 · Nordic Emerald": {
                "sidebar_bg": "#071C19",
                "sidebar_card": "#0D2E2A",
                "sidebar_active": "#134E4A",
                "canvas_bg": "#F0F7F6",
                "hero_bg": "#FFFFFF",
                "hero_border": "#99F6E4",
                "hero_pill_bg": "#CCFBF1",
                "hero_pill_fg": "#0F766E",
                "hero_title": "#042F2E",
                "hero_sub": "#334155",
                "wave_cols": ("#0D9488", "#0284C7", "#059669"),
                "card_bg": "#FFFFFF",
                "card_border": "#CBD5E1",
                "card_shadow": "#D1E2DF",
                "card_title": "#0F172A",
                "card_desc": "#475569",
                "tag_border": "#99F6E4",
                "axes": {
                    "tool_ses_kaydet": ("MIC", "#0D9488", "#F0FDFA", "#0F766E"),
                    "tool_ai_ozet": ("EDU", "#0284C7", "#F0F9FF", "#0369A1"),
                    "tool_video_cevir": ("GLOBE", "#059669", "#ECFDF5", "#047857"),
                    "tool_ai_rapor": ("FIELD", "#4F46E5", "#EEF2FF", "#4338CA"),
                },
            },
            "5 · Royal Gold LEX": {
                "sidebar_bg": "#0A1120",
                "sidebar_card": "#152036",
                "sidebar_active": "#1E2D4A",
                "canvas_bg": "#F8F7F4",
                "hero_bg": "#0F172A",
                "hero_border": "#B45309",
                "hero_pill_bg": "#292018",
                "hero_pill_fg": "#F59E0B",
                "hero_title": "#FFFFFF",
                "hero_sub": "#CBD5E1",
                "wave_cols": ("#F59E0B", "#D97706", "#38BDF8"),
                "card_bg": "#FFFFFF",
                "card_border": "#E2DDD5",
                "card_shadow": "#E5E0D8",
                "card_title": "#0F172A",
                "card_desc": "#475569",
                "tag_border": "#E2DDD5",
                "axes": {
                    "tool_ses_kaydet": ("MIC", "#1E3A8A", "#EFF6FF", "#1E3A8A"),
                    "tool_ai_ozet": ("EDU", "#B45309", "#FEF3C7", "#92400E"),
                    "tool_video_cevir": ("GLOBE", "#0F766E", "#F0FDFA", "#0F766E"),
                    "tool_ai_rapor": ("FIELD", "#7C2D12", "#FFEDD5", "#7C2D12"),
                },
            },
        }

    def _make_vector_module_icon(self, kind, bg_hex, fg_hex="#FFFFFF", size=50):
        """Render a 4x supersampled anti-aliased vector icon (replacing cartoonish system emojis)."""
        from PIL import ImageDraw
        S = size * 4
        im = Image.new("RGBA", (S, S), (0, 0, 0, 0))
        d = ImageDraw.Draw(im)
        d.rounded_rectangle([0, 0, S - 1, S - 1], radius=int(S * 0.26), fill=bg_hex)
        w = int(S * 0.068)
        cx, cy = S // 2, S // 2
        if kind == "MIC":
            d.rounded_rectangle([cx - 24, cy - 54, cx + 24, cy + 6], radius=24, outline=fg_hex, width=w)
            d.line([cx - 10, cy - 30, cx + 10, cy - 30], fill=fg_hex, width=max(4, w - 4))
            d.line([cx - 10, cy - 15, cx + 10, cy - 15], fill=fg_hex, width=max(4, w - 4))
            d.arc([cx - 44, cy - 38, cx + 44, cy + 30], start=0, end=180, fill=fg_hex, width=w)
            d.line([cx, cy + 30, cx, cy + 54], fill=fg_hex, width=w)
            d.line([cx - 26, cy + 54, cx + 26, cy + 54], fill=fg_hex, width=w)
        elif kind == "EDU":
            d.rounded_rectangle([cx - 46, cy - 50, cx + 20, cy + 34], radius=12, outline=fg_hex, width=w)
            d.line([cx - 30, cy - 26, cx + 4, cy - 26], fill=fg_hex, width=max(5, w - 3))
            d.line([cx - 30, cy - 8, cx + 4, cy - 8], fill=fg_hex, width=max(5, w - 3))
            d.line([cx - 30, cy + 10, cx - 8, cy + 10], fill=fg_hex, width=max(5, w - 3))
            d.rounded_rectangle([cx - 6, cy - 8, cx + 50, cy + 52], radius=12, fill=bg_hex, outline=fg_hex, width=w)
            d.line([cx + 10, cy + 12, cx + 34, cy + 12], fill=fg_hex, width=max(5, w - 3))
            d.line([cx + 10, cy + 30, cx + 34, cy + 30], fill=fg_hex, width=max(5, w - 3))
        elif kind == "GLOBE":
            r = 48
            d.ellipse([cx - r, cy - r - 6, cx + r, cy + r - 6], outline=fg_hex, width=w)
            d.arc([cx - 24, cy - r - 6, cx + 24, cy + r - 6], start=0, end=360, fill=fg_hex, width=max(5, w - 3))
            d.line([cx - r + 4, cy - 6, cx + r - 4, cy - 6], fill=fg_hex, width=max(5, w - 3))
            d.rounded_rectangle([cx - 40, cy + 34, cx + 40, cy + 54], radius=7, fill=fg_hex)
        elif kind == "FIELD":
            for idx_b, bh in enumerate([26, 52, 78, 40, 68, 30]):
                bx = cx - 50 + idx_b * 20
                d.rounded_rectangle([bx - 4, cy - bh // 2, bx + 4, cy + bh // 2], radius=4, fill=fg_hex)
            d.ellipse([cx - 54, cy - 54, cx - 32, cy - 32], outline=fg_hex, width=max(5, w - 3))
            d.ellipse([cx + 32, cy - 54, cx + 54, cy - 32], outline=fg_hex, width=max(5, w - 3))
        pil_res = im.resize((size, size), Image.Resampling.LANCZOS)
        return ctk.CTkImage(light_image=pil_res, dark_image=pil_res, size=(size, size))

    def _build_dashboard(self):
        if not hasattr(self, "active_theme_name"):
            self.active_theme_name = "1 · Executive Slate"
        themes = self._get_theme_defs()
        th = themes.get(self.active_theme_name, themes["1 · Executive Slate"])

        self.configure(fg_color=th["canvas_bg"])

        # Left Sidebar
        self.sidebar = ctk.CTkFrame(self, width=245, fg_color=th["sidebar_bg"], corner_radius=0)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        brand_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        brand_frame.pack(pady=(32, 22), fill="x")

        logo_row = ctk.CTkFrame(brand_frame, fg_color="transparent")
        logo_row.pack(anchor="w", padx=25)

        try:
            logo_p = resource_path("ilber_logo.jpg")
            if os.path.exists(logo_p):
                logo_img = ctk.CTkImage(light_image=Image.open(logo_p), size=(45, 45))
                ctk.CTkLabel(logo_row, image=logo_img, text="").pack(side="left", padx=(0, 12))
        except Exception:
            pass

        ctk.CTkLabel(logo_row, text="ILBER", font=ctk.CTkFont(family="Segoe UI", size=28, weight="bold"), text_color="#F8FAFC").pack(side="left")
        ctk.CTkLabel(brand_frame, text="Enterprise AI Suite", font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"), text_color="#94A3B8").pack(anchor="w", padx=27, pady=(4, 0))

        self.nav_btns = {}
        nav_items = [
            (self.t("nav_home"), "home"),
            (self.t("nav_files"), "files"),
            (self.t("nav_history"), "history"),
            (self.t("nav_settings"), "settings"),
        ]
        for nav, pid in nav_items:
            bg_col = th["sidebar_active"] if pid == "home" else "transparent"
            b = ctk.CTkButton(
                self.sidebar,
                text=nav,
                fg_color=bg_col,
                text_color="#FFFFFF",
                hover_color=th["sidebar_active"],
                anchor="w",
                font=ctk.CTkFont(family="Segoe UI", size=14),
                command=lambda p=pid: self.show_page(p),
            )
            b.pack(fill="x", padx=10, pady=3)
            self.nav_btns[pid] = b

        def change_ui_lang(choice):
            lmap = {"🇬🇧 English": "en", "🇹🇷 Türkçe": "tr", "🇩🇪 Deutsch": "de"}
            new_l = lmap.get(choice, "en")
            if self.ui_lang == new_l:
                return
            prev_page = getattr(self, "current_page", "home")
            saved_ctx = getattr(self, "video_context", "")
            saved_hist = list(getattr(self, "chat_history", []))
            self.ui_lang = new_l
            for widget in self.winfo_children():
                widget.destroy()
            self.pages.clear()
            self.log_boxes.clear()
            self.page_controls.clear()
            self._build_dashboard()
            self.video_context = saved_ctx
            self.chat_history = saved_hist
            if prev_page in self.pages:
                self.show_page(prev_page)

        def change_ui_theme(choice):
            if getattr(self, "active_theme_name", "") == choice:
                return
            self.active_theme_name = choice
            prev_page = getattr(self, "current_page", "home")
            for widget in self.winfo_children():
                widget.destroy()
            self.pages.clear()
            self.log_boxes.clear()
            self.page_controls.clear()
            self._build_dashboard()
            if prev_page in self.pages:
                self.show_page(prev_page)

        lang_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        lang_frame.pack(side="bottom", fill="x", pady=(4, 14), padx=14)
        rev_map = {"en": "🇬🇧 English", "tr": "🇹🇷 Türkçe", "de": "🇩🇪 Deutsch"}
        opt = ctk.CTkOptionMenu(
            lang_frame,
            values=["🇬🇧 English", "🇹🇷 Türkçe", "🇩🇪 Deutsch"],
            command=change_ui_lang,
            fg_color=th["sidebar_active"],
            text_color="#FFFFFF",
            button_color=th["sidebar_active"],
            button_hover_color="#334155",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            height=30,
        )
        opt.set(rev_map.get(self.ui_lang, "🇬🇧 English"))
        opt.pack(fill="x", pady=(4, 0))

        theme_opt = ctk.CTkOptionMenu(
            lang_frame,
            values=list(themes.keys()),
            command=change_ui_theme,
            fg_color=th["sidebar_card"],
            text_color="#38BDF8",
            button_color=th["sidebar_card"],
            button_hover_color="#334155",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            height=30,
        )
        theme_opt.set(self.active_theme_name)
        theme_opt.pack(fill="x")

        # Sidebar Account & Package Status Box
        lic_box = ctk.CTkFrame(self.sidebar, fg_color=th["sidebar_card"], corner_radius=12, border_width=1, border_color="#334155")
        lic_box.pack(side="bottom", fill="x", padx=14, pady=(4, 4))

        self.sidebar_email_lbl = ctk.CTkLabel(
            lic_box,
            text=self.license_mgr.session.get("email") or "deneme@ilbertech.com",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color="#94A3B8",
        )
        self.sidebar_email_lbl.pack(anchor="w", padx=12, pady=(8, 0))

        self.sidebar_plan_lbl = ctk.CTkLabel(
            lic_box,
            text=self.license_mgr.get_plan_display_name(self.ui_lang),
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color="#38BDF8",
            wraplength=195,
            justify="left",
        )
        self.sidebar_plan_lbl.pack(anchor="w", padx=12, pady=(2, 6))

        ctk.CTkButton(
            lic_box,
            text=self.L("💎 Paketler & Yükselt", "💎 Packages & Upgrade", "💎 Pakete & Upgrade"),
            height=32,
            fg_color="#2563EB",
            hover_color="#1D4ED8",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            command=lambda: self.open_package_upgrade_modal(),
        ).pack(fill="x", padx=10, pady=(0, 5))

        def do_logout():
            self.license_mgr.logout()
            self._init_login()

        ctk.CTkButton(
            lic_box,
            text=self.L("🚪 Oturumu Kapat", "🚪 Sign Out", "🚪 Abmelden"),
            height=28,
            fg_color="transparent",
            text_color="#CBD5E1",
            hover_color="#334155",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            command=do_logout,
        ).pack(fill="x", padx=10, pady=(0, 8))

        # Center Main Area
        self.main_area = ctk.CTkFrame(self, fg_color=th["canvas_bg"])
        self.main_area.pack(side="left", fill="both", expand=True)

        self.pages["home"] = ctk.CTkScrollableFrame(self.main_area, fg_color="#F8FAFC")
        self._build_home()

        self.pages["files"] = ctk.CTkFrame(self.main_area, fg_color="#F8FAFC")
        self._build_files()

        self.pages["history"] = ctk.CTkFrame(self.main_area, fg_color="#F8FAFC")
        self._build_history()

        self.pages["settings"] = ctk.CTkFrame(self.main_area, fg_color="#F8FAFC")
        self._build_settings()

        # Build the 4 core segment pages (MIC, EDU [Format-Preserving Document Translator], GLOBE, FIELD)
        self.create_tool_page("tool_ses_kaydet", self.t("c1_t"), self.t("c1_d"), seg_type="MIC")
        self.create_tool_page("tool_ai_ozet", self.t("c2_t"), self.t("c2_d"), seg_type="EDU")
        self.create_tool_page("tool_video_cevir", self.t("c3_t"), self.t("c3_d"), seg_type="GLOBE")
        self.create_tool_page("tool_ai_rapor", self.t("c4_t"), self.t("c4_d"), seg_type="FIELD")

        self.show_page("home")

    def _build_settings(self):
        parent = self.pages["settings"]
        for widget in parent.winfo_children():
            widget.destroy()

        header = ctk.CTkFrame(parent, fg_color="transparent")
        header.pack(fill="x", padx=40, pady=(40, 20))
        ctk.CTkLabel(header, text=self.t("title_settings"), font=ctk.CTkFont(family="Segoe UI", size=28, weight="bold"), text_color="#111827").pack(side="left")

        scroll = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=40)

        def save_setting(key, val):
            self.user_settings[key] = val

        card1 = ctk.CTkFrame(scroll, fg_color="#FFFFFF", border_width=1, border_color="#E2E8F0", corner_radius=16)
        card1.pack(fill="x", pady=10)
        ctk.CTkLabel(card1, text=self.t("set_sub_pref"), font=ctk.CTkFont(family="Segoe UI", size=18, weight="bold"), text_color="#1E293B").pack(anchor="w", padx=25, pady=(25, 10))

        row1 = ctk.CTkFrame(card1, fg_color="transparent")
        row1.pack(fill="x", padx=25, pady=10)

        box1 = ctk.CTkFrame(row1, fg_color="transparent")
        box1.pack(side="left", padx=(0, 40))
        ctk.CTkLabel(box1, text=self.t("set_font_size"), text_color="#64748B", font=ctk.CTkFont(family="Segoe UI", size=13)).pack(anchor="w")
        cb_size = ctk.CTkOptionMenu(box1, values=["16", "20", "24", "28", "32"], command=lambda v: save_setting("sub_size", v), width=150, fg_color="#F8FAFC", text_color="#1E293B", button_color="#E2E8F0", button_hover_color="#CBD5E1")
        cb_size.set(self.user_settings["sub_size"])
        cb_size.pack(anchor="w", pady=(5, 0))

        box2 = ctk.CTkFrame(row1, fg_color="transparent")
        box2.pack(side="left")
        ctk.CTkLabel(box2, text=self.t("set_pos"), text_color="#64748B", font=ctk.CTkFont(family="Segoe UI", size=13)).pack(anchor="w")
        cb_pos = ctk.CTkOptionMenu(box2, values=["Bottom", "Top"], command=lambda v: save_setting("sub_pos", v), width=150, fg_color="#F8FAFC", text_color="#1E293B", button_color="#E2E8F0", button_hover_color="#CBD5E1")
        cb_pos.set(self.user_settings["sub_pos"])
        cb_pos.pack(anchor="w", pady=(5, 0))

        row2 = ctk.CTkFrame(card1, fg_color="transparent")
        row2.pack(fill="x", padx=25, pady=(20, 30))
        ctk.CTkLabel(row2, text=self.t("set_sub_style"), text_color="#64748B", font=ctk.CTkFont(family="Segoe UI", size=13)).pack(anchor="w", pady=(0, 10))

        style_frame = ctk.CTkFrame(row2, fg_color="transparent")
        style_frame.pack(fill="x")

        style_btns = []

        def select_style(btn_obj, s_name):
            save_setting("sub_style", s_name)
            for b in style_btns:
                b.configure(fg_color="#F8FAFC", border_color="#E2E8F0", text_color="#475569")
            btn_obj.configure(fg_color="#00BCD4", border_color="#00838F", text_color="#FFFFFF")

        styles = [("Classic White", "style_classic"), ("Yellow Outlined", "style_yellow"), ("Black Box", "style_black")]
        for internal_name, ui_key in styles:
            b = ctk.CTkButton(style_frame, text=self.t(ui_key), height=45, corner_radius=8, font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"))
            if internal_name == self.user_settings["sub_style"]:
                b.configure(fg_color="#00BCD4", border_color="#00838F", text_color="#FFFFFF", border_width=1)
            else:
                b.configure(fg_color="#F8FAFC", border_color="#E2E8F0", text_color="#475569", border_width=1)
            b.configure(command=lambda btn=b, name=internal_name: select_style(btn, name))
            b.pack(side="left", fill="x", expand=True, padx=(0, 10) if internal_name != styles[-1][0] else 0)
            style_btns.append(b)

        card2 = ctk.CTkFrame(scroll, fg_color="#FFFFFF", border_width=1, border_color="#E2E8F0", corner_radius=16)
        card2.pack(fill="x", pady=20)
        ctk.CTkLabel(card2, text=self.t("set_gen_pref"), font=ctk.CTkFont(family="Segoe UI", size=18, weight="bold"), text_color="#1E293B").pack(anchor="w", padx=25, pady=(25, 10))

        row3 = ctk.CTkFrame(card2, fg_color="transparent")
        row3.pack(fill="x", padx=25, pady=(10, 30))
        ctk.CTkLabel(row3, text=self.t("set_def_lang"), text_color="#64748B", font=ctk.CTkFont(family="Segoe UI", size=13)).pack(anchor="w")
        cb_lang = ctk.CTkOptionMenu(row3, values=TARGET_LANGS, command=lambda v: save_setting("def_lang", v), width=250, fg_color="#F8FAFC", text_color="#1E293B", button_color="#E2E8F0", button_hover_color="#CBD5E1")
        cb_lang.set(self.user_settings["def_lang"])
        cb_lang.pack(anchor="w", pady=(5, 0))

        card3 = ctk.CTkFrame(scroll, fg_color="#FFFFFF", border_width=1, border_color="#EF4444", corner_radius=16)
        card3.pack(fill="x", pady=10)
        row4 = ctk.CTkFrame(card3, fg_color="transparent")
        row4.pack(fill="x", padx=25, pady=25)
        ctk.CTkLabel(row4, text=self.t("set_data_mgmt"), font=ctk.CTkFont(family="Segoe UI", size=18, weight="bold"), text_color="#1E293B").pack(anchor="w")
        ctk.CTkLabel(row4, text=self.t("set_data_desc"), font=ctk.CTkFont(family="Segoe UI", size=13), text_color="#64748B").pack(anchor="w")

        def clear_workspace():
            if messagebox.askyesno(self.t("warn_title"), self.t("warn_desc")):
                ws = os.path.join(os.path.expanduser("~"), "Documents", "ILBER_Workspace")
                for s in ["Uploads", "Transcripts", "Subtitles", "Reports"]:
                    try:
                        shutil.rmtree(os.path.join(ws, s))
                        os.makedirs(os.path.join(ws, s), exist_ok=True)
                    except Exception:
                        pass
                messagebox.showinfo(self.t("succ_title"), self.t("succ_desc"))

        ctk.CTkButton(row4, text=self.t("set_clear_btn"), fg_color="#FEF2F2", text_color="#EF4444", hover_color="#FEE2E2", border_color="#EF4444", border_width=1, font=ctk.CTkFont(family="Segoe UI", weight="bold"), command=clear_workspace).pack(anchor="e")

    def _build_history(self):
        parent = self.pages["history"]
        for widget in parent.winfo_children():
            widget.destroy()

        workspace_dir = os.path.join(os.path.expanduser("~"), "Documents", "ILBER_Workspace")
        os.makedirs(workspace_dir, exist_ok=True)

        header = ctk.CTkFrame(parent, fg_color="transparent")
        header.pack(fill="x", padx=40, pady=(40, 20))
        ctk.CTkLabel(header, text=self.t("title_history"), font=ctk.CTkFont(family="Segoe UI", size=28, weight="bold"), text_color="#111827").pack(side="left")

        trans_count = 0
        sub_count = 0
        total_size = 0
        history_items = []

        for root, dirs, files in os.walk(workspace_dir):
            if "Uploads" in root or "Models" in root:
                continue
            for name in files:
                path = os.path.join(root, name)
                try:
                    stat = os.stat(path)
                    total_size += stat.st_size
                    action = self.t("act_processed")
                    icon = "⚡"
                    color = "#64748B"
                    bg_color = "#F1F5F9"
                    if name.endswith((".docx", ".txt", ".tsv")):
                        trans_count += 1
                        action = self.t("act_transcribed")
                        icon = "📝"
                        color = "#F59E0B"
                        bg_color = "#FFF7E6"
                    elif name.endswith((".srt", ".vtt")):
                        sub_count += 1
                        action = self.t("act_srt")
                        icon = "🎬"
                        color = "#2563EB"
                        bg_color = "#EBF3FF"
                    elif name.endswith(".mp4"):
                        trans_count += 1
                        action = self.t("act_translated")
                        icon = "🌐"
                        color = "#10B981"
                        bg_color = "#E6FBF2"
                    history_items.append((stat.st_mtime, action, name, path, icon, color, bg_color))
                except Exception:
                    pass

        stats_frame = ctk.CTkFrame(parent, fg_color="transparent")
        stats_frame.pack(fill="x", padx=40, pady=10)

        def make_stat_card(p_frame, title, val, icon_txt):
            glass_map = {
                self.t("stat_trans"): ("#FFF7E6", "#FDE68A", "#F59E0B"),
                self.t("stat_subs"): ("#EBF3FF", "#BFDBFE", "#2563EB"),
                self.t("stat_storage"): ("#E6FBF2", "#A7F3D0", "#10B981"),
            }
            bg_glass, shadow_color, icon_bg = glass_map.get(title, ("#FFFFFF", "#E2E8F0", "#64748B"))
            container = ctk.CTkFrame(p_frame, fg_color="transparent", height=100)
            container.pack(side="left", fill="x", expand=True, padx=(0, 20))
            container.pack_propagate(False)
            shadow = ctk.CTkFrame(container, fg_color=shadow_color, corner_radius=20)
            shadow.place(relwidth=1, relheight=0.9, y=10)
            c = ctk.CTkFrame(container, fg_color=bg_glass, border_width=1, border_color="#FFFFFF", corner_radius=20)
            c.place(relwidth=1, relheight=0.9, y=0)
            icon_box = ctk.CTkFrame(c, fg_color=icon_bg, width=40, height=40, corner_radius=12)
            icon_box.pack(side="left", padx=20)
            ctk.CTkLabel(icon_box, text=icon_txt, font=ctk.CTkFont(size=20), text_color="#FFFFFF").place(relx=0.5, rely=0.5, anchor="center")
            tf = ctk.CTkFrame(c, fg_color="transparent")
            tf.pack(side="left", pady=15, fill="both")
            ctk.CTkLabel(tf, text=val, font=ctk.CTkFont(family="Segoe UI", size=22, weight="bold"), text_color="#1E293B").pack(anchor="w", pady=(0, 2))
            ctk.CTkLabel(tf, text=title, font=ctk.CTkFont(family="Segoe UI", size=13), text_color="#64748B").pack(anchor="w")

        make_stat_card(stats_frame, self.t("stat_trans"), f"{trans_count}", "🎙️")
        make_stat_card(stats_frame, self.t("stat_subs"), f"{sub_count}", "🌍")
        make_stat_card(stats_frame, self.t("stat_storage"), f"{total_size / (1024*1024):.1f} MB", "💾")

        list_frame = ctk.CTkFrame(parent, fg_color="transparent")
        list_frame.pack(fill="both", expand=True, padx=40, pady=20)
        ctk.CTkLabel(list_frame, text=self.t("timeline_title"), font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold"), text_color="#475569").pack(anchor="w", pady=(0, 15))

        scroll = ctk.CTkScrollableFrame(list_frame, fg_color="#F8FAFC", corner_radius=0)
        scroll.pack(fill="both", expand=True)

        history_items.sort(key=lambda x: x[0], reverse=True)
        if not history_items:
            ctk.CTkLabel(scroll, text=self.t("no_history"), text_color="#94A3B8").pack(pady=40)

        def open_path(p):
            open_path_crossplatform(p)

        for mtime, action, name, path, icon, color, bg_color in history_items[:40]:
            row = ctk.CTkFrame(scroll, fg_color="#FFFFFF", border_width=1, border_color="#E2E8F0", corner_radius=12, height=75)
            row.pack(fill="x", pady=6, padx=5)
            row.pack_propagate(False)
            i_box = ctk.CTkFrame(row, fg_color=bg_color, width=45, height=45, corner_radius=12)
            i_box.pack(side="left", padx=15)
            i_box.pack_propagate(False)
            l_i = ctk.CTkLabel(i_box, text=icon, font=ctk.CTkFont(size=20), text_color=color)
            l_i.place(relx=0.5, rely=0.5, anchor="center")
            tf = ctk.CTkFrame(row, fg_color="transparent")
            tf.pack(side="left", padx=10, fill="y", pady=15)
            l_a = ctk.CTkLabel(tf, text=action, font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"), text_color="#1E293B")
            l_a.pack(anchor="w")
            l_n = ctk.CTkLabel(tf, text=f"{self.t('file_prefix')} {name}", font=ctk.CTkFont(family="Segoe UI", size=12), text_color="#64748B")
            l_n.pack(anchor="w")
            date_str = datetime.datetime.fromtimestamp(mtime).strftime("%b %d, %H:%M")
            l_d = ctk.CTkLabel(row, text=date_str, font=ctk.CTkFont(family="Segoe UI", size=13), text_color="#94A3B8")
            l_d.pack(side="right", padx=20)
            for w in [row, i_box, l_i, tf, l_a, l_n, l_d]:
                try:
                    w.bind("<Button-1>", lambda e, p=path: open_path(p))
                    w.configure(cursor="hand2")
                except Exception:
                    pass

    def _build_files(self):
        parent = self.pages["files"]
        for widget in parent.winfo_children():
            widget.destroy()

        workspace_dir = os.path.join(os.path.expanduser("~"), "Documents", "ILBER_Workspace")
        os.makedirs(workspace_dir, exist_ok=True)
        for s in ["Uploads", "Transcripts", "Subtitles", "Reports"]:
            os.makedirs(os.path.join(workspace_dir, s), exist_ok=True)

        header = ctk.CTkFrame(parent, fg_color="transparent")
        header.pack(fill="x", padx=40, pady=(40, 20))
        ctk.CTkLabel(header, text=self.t("title_files"), font=ctk.CTkFont(family="Segoe UI", size=28, weight="bold"), text_color="#111827").pack(side="left")

        search_box = ctk.CTkEntry(header, placeholder_text=self.t("search_placeholder"), width=250, height=40, corner_radius=20, border_color="#E2E8F0", fg_color="#FFFFFF")
        search_box.pack(side="right")

        upload_status = ctk.CTkLabel(header, text="", text_color="#10B981", font=ctk.CTkFont(family="Segoe UI", size=13))
        upload_status.pack(side="right", padx=15)

        def upload_action():
            filepath = filedialog.askopenfilename(filetypes=SUPPORTED_MEDIA_EXTS)
            if filepath:
                dest = os.path.join(workspace_dir, "Uploads", os.path.basename(filepath))
                shutil.copy2(filepath, dest)
                upload_status.configure(text=self.t("msg_uploaded"))
                self.after(3000, lambda: upload_status.configure(text=""))
                refresh_file_list()

        btn_upload = ctk.CTkButton(header, text=self.t("btn_upload"), font=ctk.CTkFont(family="Segoe UI", weight="bold"), fg_color="#2563EB", hover_color="#1D4ED8", height=40, corner_radius=20, command=upload_action)
        btn_upload.pack(side="right", padx=10)

        folder_frame = ctk.CTkFrame(parent, fg_color="transparent")
        folder_frame.pack(fill="x", padx=40, pady=10)
        ctk.CTkLabel(folder_frame, text=self.t("quick_access"), font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold"), text_color="#475569").pack(anchor="w", pady=(0, 15))
        folder_grid = ctk.CTkFrame(folder_frame, fg_color="transparent")
        folder_grid.pack(fill="x")

        def open_folder(sub):
            open_path_crossplatform(os.path.join(workspace_dir, sub))

        folders = [
            (self.t("folder_trans"), "Transcripts", "#F4EBFF", "#9333EA"),
            (self.t("folder_ups"), "Uploads", "#E6FAFD", "#06B6D4"),
            (self.t("folder_subs"), "Subtitles", "#EBF3FF", "#2563EB"),
        ]
        folder_labels = {}
        for title, sub, bg, color in folders:
            glass_map = {
                "Transcripts": ("#F4EBFF", "#E9D5FF"),
                "Uploads": ("#E6FAFD", "#A5F3FC"),
                "Subtitles": ("#EBF3FF", "#BFDBFE"),
            }
            bg_glass, shadow_color = glass_map.get(sub, ("#FFFFFF", "#E2E8F0"))
            container = ctk.CTkFrame(folder_grid, fg_color="transparent", height=110)
            container.pack(side="left", fill="x", expand=True, padx=(0, 20) if title != folders[-1][0] else 0)
            container.pack_propagate(False)
            shadow = ctk.CTkFrame(container, fg_color=shadow_color, corner_radius=20)
            shadow.place(relwidth=1, relheight=0.9, y=10)
            f_card = ctk.CTkFrame(container, fg_color=bg_glass, border_width=1, border_color="#FFFFFF", corner_radius=20)
            f_card.place(relwidth=1, relheight=0.9, y=0)
            icon = ctk.CTkFrame(f_card, fg_color=color, width=40, height=40, corner_radius=12)
            icon.pack(side="left", padx=20)
            ctk.CTkLabel(icon, text="📁", font=ctk.CTkFont(size=20), text_color="#FFFFFF").place(relx=0.5, rely=0.5, anchor="center")
            text_frame = ctk.CTkFrame(f_card, fg_color="transparent")
            text_frame.pack(side="left", padx=10, pady=25, fill="both")
            l1 = ctk.CTkLabel(text_frame, text=title, font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"), text_color="#1E293B")
            l1.pack(anchor="w")
            l2 = ctk.CTkLabel(text_frame, text=f"0 {self.t('files_count')}", font=ctk.CTkFont(family="Segoe UI", size=12), text_color="#64748B")
            l2.pack(anchor="w", pady=(2, 0))
            folder_labels[sub] = l2
            widgets = [container, shadow, f_card, icon, l1, l2, text_frame] + icon.winfo_children()
            for w in widgets:
                try:
                    w.bind("<Button-1>", lambda e, s=sub: open_folder(s))
                    w.configure(cursor="hand2")
                except Exception:
                    pass

        list_frame = ctk.CTkFrame(parent, fg_color="transparent")
        list_frame.pack(fill="both", expand=True, padx=40, pady=25)
        ctk.CTkLabel(list_frame, text=self.t("recent_files"), font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold"), text_color="#475569").pack(anchor="w", pady=(0, 12))

        scroll = ctk.CTkScrollableFrame(list_frame, fg_color="#FFFFFF", corner_radius=16, border_width=1, border_color="#E2E8F0")
        scroll.pack(fill="both", expand=True)

        def open_file(fpath):
            open_path_crossplatform(fpath)

        def refresh_file_list(*args):
            for sub, lbl in folder_labels.items():
                try:
                    lbl.configure(text=f"{len(os.listdir(os.path.join(workspace_dir, sub)))} {self.t('files_count')}")
                except Exception:
                    pass
            for widget in scroll.winfo_children():
                widget.destroy()
            query = search_box.get().lower()
            all_files = []
            for root, dirs, files in os.walk(workspace_dir):
                if "Models" in root:
                    continue
                for name in files:
                    if query and query not in name.lower():
                        continue
                    path = os.path.join(root, name)
                    try:
                        stat = os.stat(path)
                        all_files.append((path, name, stat.st_size, stat.st_mtime))
                    except Exception:
                        pass
            all_files.sort(key=lambda x: x[3], reverse=True)
            if not all_files:
                ctk.CTkLabel(scroll, text=self.t("no_files"), text_color="#94A3B8").pack(pady=40)
                return
            for path, name, size, mtime in all_files[:25]:
                row = ctk.CTkFrame(scroll, fg_color="transparent", height=55)
                row.pack(fill="x", pady=4, padx=10)
                row.pack_propagate(False)
                ext = name.split(".")[-1].lower() if "." in name else ""
                typ = "Video" if ext in ["mp4", "mkv", "mov", "webm"] else "Audio" if ext in ["mp3", "wav", "m4a", "ogg", "flac"] else "Subtitle" if ext in ["srt", "vtt"] else "Document"
                icon_char = "🎬" if typ == "Video" else "🎵" if typ == "Audio" else "📝"
                l_icon = ctk.CTkLabel(row, text=icon_char, font=ctk.CTkFont(size=20))
                l_icon.pack(side="left", padx=15)
                size_str = f"{size / (1024*1024):.1f} MB" if size > 1024*1024 else f"{size / 1024:.1f} KB"
                date_str = datetime.datetime.fromtimestamp(mtime).strftime("%b %d, %Y")
                l_name = ctk.CTkLabel(row, text=name, font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"), text_color="#1E293B", width=300, anchor="w")
                l_name.pack(side="left")
                ctk.CTkLabel(row, text=typ, font=ctk.CTkFont(family="Segoe UI", size=13), text_color="#64748B", width=100, anchor="w").pack(side="left")
                ctk.CTkLabel(row, text=size_str, font=ctk.CTkFont(family="Segoe UI", size=13), text_color="#64748B", width=100, anchor="w").pack(side="left")
                ctk.CTkLabel(row, text=date_str, font=ctk.CTkFont(family="Segoe UI", size=13), text_color="#64748B", width=120, anchor="w").pack(side="left")
                for w in [row, l_icon, l_name]:
                    try:
                        w.bind("<Button-1>", lambda e, p=path: open_file(p))
                        w.configure(cursor="hand2")
                    except Exception:
                        pass

        search_box.bind("<KeyRelease>", refresh_file_list)
        self.after(200, refresh_file_list)

    def _build_home(self):
        parent = self.pages["home"]
        for widget in parent.winfo_children():
            widget.destroy()

        themes = self._get_theme_defs()
        th = themes.get(getattr(self, "active_theme_name", "1 · Executive Slate"), themes["1 · Executive Slate"])

        banner = ctk.CTkFrame(parent, fg_color=th["hero_bg"], corner_radius=18, border_width=2, border_color=th["hero_border"])
        banner.pack(fill="x", padx=36, pady=(20, 4))

        left_side = ctk.CTkFrame(banner, fg_color="transparent")
        left_side.pack(side="left", fill="both", expand=True, padx=32, pady=22)

        badge_bg = ctk.CTkFrame(left_side, fg_color=th["hero_pill_bg"], corner_radius=12)
        badge_bg.pack(anchor="w", pady=(0, 10))
        ctk.CTkLabel(
            badge_bg,
            text=self.t("badge_tags"),
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color=th["hero_pill_fg"],
        ).pack(padx=12, pady=4)

        ctk.CTkLabel(
            left_side,
            text=self.t("banner_title"),
            font=ctk.CTkFont(family="Segoe UI", size=28, weight="bold"),
            text_color=th["hero_title"],
            justify="left",
            anchor="w",
        ).pack(anchor="w", fill="x")
        ctk.CTkLabel(
            left_side,
            text=self.t("banner_sub"),
            font=ctk.CTkFont(family="Segoe UI", size=13),
            text_color=th["hero_sub"],
            justify="left",
            anchor="w",
            wraplength=540,
        ).pack(anchor="w", fill="x", pady=(6, 0))

        right_side = ctk.CTkFrame(banner, fg_color="transparent", width=250, height=120)
        right_side.pack(side="right", padx=28, pady=15)
        right_side.pack_propagate(False)

        import tkinter as tk
        canvas = tk.Canvas(right_side, width=250, height=120, bg=th["hero_bg"], highlightthickness=0)
        canvas.pack(fill="both", expand=True)

        wave_bars = [16, 28, 44, 68, 88, 58, 102, 82, 48, 72, 96, 62, 38, 76, 52, 32, 18]
        center_y = 60
        w_cols = th["wave_cols"]
        for bi, bh in enumerate(wave_bars):
            bx = 16 + bi * 13
            col_b = w_cols[bi % len(w_cols)]
            canvas.create_line(bx, center_y - bh // 2, bx, center_y + bh // 2, fill=col_b, width=5, capstyle="round")

        grid_container = ctk.CTkFrame(parent, fg_color="transparent")
        grid_container.pack(fill="both", expand=True, pady=(8, 14))

        grid = ctk.CTkFrame(grid_container, fg_color="transparent")
        grid.pack(fill="both", expand=True, padx=32, pady=2)
        grid.grid_columnconfigure(0, weight=1, uniform="home_col")
        grid.grid_columnconfigure(1, weight=1, uniform="home_col")
        grid.grid_rowconfigure(0, weight=1)
        grid.grid_rowconfigure(1, weight=1)

        cards = [
            (self.t("c1_t"), self.t("c1_d"), "tool_ses_kaydet", self.t("c1_b")),
            (self.t("c2_t"), self.t("c2_d"), "tool_ai_ozet", self.t("c2_b")),
            (self.t("c3_t"), self.t("c3_d"), "tool_video_cevir", self.t("c3_b")),
            (self.t("c4_t"), self.t("c4_d"), "tool_ai_rapor", self.t("c4_b")),
        ]

        row, col = 0, 0
        for title, desc, pid, badges in cards:
            unlocked = self.license_mgr.is_module_unlocked(pid)
            v_kind, acc_col, pill_bg, pill_fg = th["axes"].get(pid, ("MIC", "#2563EB", "#EFF6FF", "#1D4ED8"))
            eff_col = acc_col if unlocked else "#64748B"

            c_frame = ctk.CTkFrame(
                grid,
                fg_color=th["card_bg"],
                border_width=2,
                border_color=th["card_border"],
                corner_radius=16,
            )
            c_frame.grid(row=row, column=col, padx=10, pady=10, sticky="nsew")

            # Left precision accent bar
            acc_bar = ctk.CTkFrame(c_frame, width=6, fg_color=eff_col, corner_radius=3)
            acc_bar.place(x=0, rely=0.1, relheight=0.8)

            top_row = ctk.CTkFrame(c_frame, fg_color="transparent")
            top_row.pack(fill="x", padx=22, pady=(18, 6))

            vec_img = self._make_vector_module_icon(v_kind, eff_col, "#FFFFFF", size=50)
            lbl_ic = ctk.CTkLabel(top_row, image=vec_img, text="")
            lbl_ic._img_ref = vec_img
            lbl_ic.pack(side="left")

            # Package / License Status Pill on Top-Right of Card
            b_txt, b_fg, b_bg = self.license_mgr.get_module_badge_info(pid, self.ui_lang)
            status_pill = ctk.CTkLabel(
                top_row,
                text=b_txt,
                font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
                text_color=b_fg,
                fg_color=b_bg,
                corner_radius=10,
                padx=10,
                pady=4,
            )
            status_pill.pack(side="right", pady=4)

            ctk.CTkLabel(
                c_frame,
                text=title,
                font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold"),
                text_color=th["card_title"] if unlocked else "#64748B",
                wraplength=420,
                justify="left",
                anchor="w",
            ).pack(anchor="w", padx=22, pady=(4, 0))

            ctk.CTkLabel(
                c_frame,
                text=desc,
                font=ctk.CTkFont(family="Segoe UI", size=12),
                text_color=th["card_desc"],
                wraplength=420,
                justify="left",
                anchor="w",
            ).pack(anchor="w", padx=22, pady=(4, 6))

            badge_frame = ctk.CTkFrame(c_frame, fg_color="transparent")
            badge_frame.pack(fill="x", padx=22, pady=(8, 16), side="bottom")
            for b in badges.split("  "):
                lbl = ctk.CTkLabel(
                    badge_frame,
                    text=b,
                    font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
                    text_color=pill_fg if unlocked else "#64748B",
                    fg_color=pill_bg,
                    corner_radius=8,
                    padx=10,
                    pady=4,
                )
                lbl.pack(side="left", padx=(0, 6))

            open_lbl = ctk.CTkLabel(
                badge_frame,
                text=self.L("Aç  →", "Open  →", "Öffnen  →"),
                font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
                text_color="#FFFFFF",
                fg_color=eff_col,
                corner_radius=8,
                padx=12,
                pady=4,
            )
            open_lbl.pack(side="right")

            widgets = [c_frame, top_row, lbl_ic, badge_frame, status_pill, open_lbl] + c_frame.winfo_children() + top_row.winfo_children() + badge_frame.winfo_children()
            for w in widgets:
                try:
                    w.bind("<Button-1>", lambda e, p=pid: self.show_page(p))
                    w.configure(cursor="hand2")
                except Exception:
                    pass

            col += 1
            if col > 1:
                col = 0
                row += 1

    def create_tool_page(self, page_id, title, desc, seg_type="MIC"):
        self.pages[page_id] = ctk.CTkFrame(self.main_area, fg_color="transparent")
        f = self.pages[page_id]
        ctrls = {"seg_type": seg_type}
        self.page_controls[page_id] = ctrls

        # Header
        top = ctk.CTkFrame(f, fg_color="transparent", height=48)
        top.pack(fill="x", padx=25, pady=(14, 4))
        ctk.CTkButton(top, text=self.t("btn_back"), width=90, fg_color=CARD, text_color=TEXT, hover_color=CARD2, command=lambda: self.show_page("home")).pack(side="left")
        ctk.CTkLabel(top, text=title, font=ctk.CTkFont(family="Segoe UI", size=21, weight="bold"), text_color=TEXT).pack(side="left", padx=18)

        ctk.CTkLabel(f, text=desc, text_color=TEXT_MUTED, font=ctk.CTkFont(family="Segoe UI", size=13), wraplength=960, justify="left").pack(anchor="w", padx=25, pady=(0, 8))

        # Segment-specific info note for MIC
        if seg_type == "MIC":
            note_box = ctk.CTkFrame(f, fg_color="#FEF3C7", border_width=1, border_color="#F59E0B", corner_radius=10)
            note_box.pack(fill="x", padx=25, pady=(0, 8))
            ctk.CTkLabel(note_box, text=self.t("mic_note"), text_color="#92400E", font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"), wraplength=940, justify="left").pack(anchor="w", padx=14, pady=6)

        # Dedicated Prominent Microphone Bar for MIC
        if seg_type == "MIC":
            mic_bar = ctk.CTkFrame(f, fg_color="#FFFFFF", corner_radius=12, border_width=1, border_color="#FECACA")
            mic_bar.pack(fill="x", padx=25, pady=(0, 6))
            rec_btn = ctk.CTkButton(
                mic_bar,
                text=self.t("btn_record"),
                fg_color=RED,
                hover_color="#B91C1C",
                width=320,
                height=42,
                corner_radius=10,
                font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
                command=lambda: self.toggle_recording(rec_btn, page_id),
            )
            rec_btn.pack(side="left", padx=14, pady=8)
            mic_status_lbl = ctk.CTkLabel(
                mic_bar,
                text=self.t("mic_hint_idle"),
                font=ctk.CTkFont(family="Segoe UI", size=13),
                text_color="#475569",
            )
            mic_status_lbl.pack(side="left", padx=10)
            ctrls["rec_btn"] = rec_btn
            ctrls["mic_status_lbl"] = mic_status_lbl

        # File & Source Language Row
        file_var = ctk.StringVar()
        src_var = ctk.StringVar(value="Auto")
        ctrls["file_var"] = file_var
        ctrls["src_var"] = src_var

        file_row = ctk.CTkFrame(f, fg_color="#FFFFFF", corner_radius=12, border_width=1, border_color="#E2E8F0")
        file_row.pack(fill="x", padx=25, pady=6)

        if seg_type == "CHAT":
            ext_filter = SUPPORTED_CHAT_EXTS
            entry_w = 210
        elif seg_type == "EDU":
            ext_filter = SUPPORTED_DOC_EXTS
            entry_w = 215
        else:
            ext_filter = SUPPORTED_MEDIA_EXTS
            entry_w = 440
        ctk.CTkEntry(file_row, textvariable=file_var, state="disabled", width=entry_w, height=36, fg_color="#F8FAFC").pack(side="left", padx=(10, 6), pady=10)
        ctk.CTkButton(
            file_row,
            text=self.L("📄 Belge / Görsel Seç", "📄 Select File / Image", "📄 Datei / Bild") if seg_type == "EDU" else self.t("btn_file"),
            width=142 if seg_type == "EDU" else 100,
            height=36,
            fg_color=PURPLE_DARK,
            hover_color=PURPLE,
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            command=lambda ef=ext_filter: file_var.set(
                filedialog.askopenfilename(
                    initialdir=os.path.join(os.path.expanduser("~"), "Documents", "ILBER_Workspace", "Uploads"),
                    filetypes=ef,
                )
            ),
        ).pack(side="left", padx=(0, 6), pady=10)

        if seg_type == "EDU":
            def on_edu_snip_ocr():
                ws_up = os.path.join(os.path.expanduser("~"), "Documents", "ILBER_Workspace", "Uploads")
                os.makedirs(ws_up, exist_ok=True)
                def _after_edu_snip(saved_png):
                    file_var.set(saved_png)
                    start_btn = ctrls.get("start_btn")
                    t_sel = ctrls.get("edu_tgt_var").get() if "edu_tgt_var" in ctrls else "Turkish"
                    s_sel = src_var.get() if src_var else "Auto"
                    self.start_pipeline(page_id, saved_png, s_sel, t_sel, start_btn)
                ScreenSnipperOverlay(self, save_dir=ws_up, ui_lang=self.ui_lang, on_capture=_after_edu_snip)

            def on_edu_paste_ocr(*args):
                from PIL import ImageGrab, Image
                try:
                    clip_data = ImageGrab.grabclipboard()
                except Exception:
                    clip_data = None
                ws_up = os.path.join(os.path.expanduser("~"), "Documents", "ILBER_Workspace", "Uploads")
                os.makedirs(ws_up, exist_ok=True)
                saved_png = None
                if isinstance(clip_data, Image.Image):
                    saved_png = os.path.join(ws_up, f"OCR_Pano_{datetime.datetime.now().strftime('%H%M%S')}.png")
                    clip_data.save(saved_png)
                elif isinstance(clip_data, list) and clip_data:
                    first_f = str(clip_data[0])
                    if os.path.exists(first_f):
                        saved_png = first_f
                if saved_png:
                    file_var.set(saved_png)
                    start_btn = ctrls.get("start_btn")
                    t_sel = ctrls.get("edu_tgt_var").get() if "edu_tgt_var" in ctrls else "Turkish"
                    s_sel = src_var.get() if src_var else "Auto"
                    self.start_pipeline(page_id, saved_png, s_sel, t_sel, start_btn)
                else:
                    lg = self.log_boxes.get(page_id)
                    if lg:
                        self._log(lg, self.L(
                            "⚠️ Panoda kopyalanmış bir resim/ekran görüntüsü bulunamadı. (Win+Shift+S ile ekran görüntüsü alabilir veya '📸 Ekran Kırp (OCR)' butonunu kullanabilirsiniz.)",
                            "⚠️ No screenshot or image found in clipboard.",
                            "⚠️ Kein Bild in der Zwischenablage gefunden.",
                        ))

            ctk.CTkButton(
                file_row,
                text=self.L("📸 Ekran Kırp (OCR)", "📸 Snip (OCR)", "📸 OCR-Snip"),
                height=36,
                width=130,
                fg_color="#DC2626",
                hover_color="#B91C1C",
                font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
                command=on_edu_snip_ocr,
            ).pack(side="left", padx=(0, 5), pady=10)

            ctk.CTkButton(
                file_row,
                text=self.L("📋 Pano (OCR)", "📋 Paste OCR", "📋 Einfügen"),
                height=36,
                width=102,
                fg_color="#0D9488",
                hover_color="#0F766E",
                font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
                command=on_edu_paste_ocr,
            ).pack(side="left", padx=(0, 6), pady=10)

        if seg_type == "CHAT":
            def on_snip_screenshot():
                ws_up = os.path.join(os.path.expanduser("~"), "Documents", "ILBER_Workspace", "Uploads")
                os.makedirs(ws_up, exist_ok=True)
                def _after_snip(saved_png):
                    file_var.set(saved_png)
                    start_btn = ctrls.get("start_btn")
                    self.start_pipeline(page_id, saved_png, "Auto", "Turkish", start_btn)
                ScreenSnipperOverlay(self, save_dir=ws_up, ui_lang=self.ui_lang, on_capture=_after_snip)

            def on_paste_clipboard_image(*args):
                import datetime
                from PIL import ImageGrab, Image
                try:
                    clip_data = ImageGrab.grabclipboard()
                except Exception:
                    clip_data = None
                ws_up = os.path.join(os.path.expanduser("~"), "Documents", "ILBER_Workspace", "Uploads")
                os.makedirs(ws_up, exist_ok=True)
                saved_png = None
                if isinstance(clip_data, Image.Image):
                    saved_png = os.path.join(ws_up, f"Pano_Goruntusu_{datetime.datetime.now().strftime('%H%M%S')}.png")
                    clip_data.save(saved_png)
                elif isinstance(clip_data, list) and clip_data:
                    first_f = str(clip_data[0])
                    if os.path.exists(first_f):
                        saved_png = first_f
                if saved_png:
                    file_var.set(saved_png)
                    start_btn = ctrls.get("start_btn")
                    self.start_pipeline(page_id, saved_png, "Auto", "Turkish", start_btn)
                    return

                # Also check if the user copied plain text (Ctrl+C) into the clipboard!
                try:
                    clip_txt = self.clipboard_get()
                except Exception:
                    clip_txt = ""
                lg = self.log_boxes.get(page_id)
                if clip_txt and len(clip_txt.strip()) >= 5:
                    cleaned_clip = _clean_ocr_turkish_text(clip_txt.strip())
                    self.video_context = cleaned_clip
                    self.chat_history = []
                    if lg:
                        self._log(lg, self.L(
                            f"\n📄 PANODAN YAPIŞTIRILAN METİN:\n--------------------------------------------------\n{cleaned_clip[:1200]}\n--------------------------------------------------\n ✔ İÇERİK HAFIZAYA ALINDI! Şimdi '🧠 Soruyu Adım Adım Çöz' butonuna basabilirsiniz.\n",
                            f"\n📄 PASTED TEXT FROM CLIPBOARD:\n--------------------------------------------------\n{cleaned_clip[:1200]}\n--------------------------------------------------\n ✔ CONTENT MEMORIZED! You can now click '🧠 Solve Step-by-Step'.\n",
                            f"\n📄 EINGEFÜGTER TEXT AUS DER ZWISCHENABLAGE:\n--------------------------------------------------\n{cleaned_clip[:1200]}\n--------------------------------------------------\n ✔ INHALT GESPEICHERT!\n"
                        ))
                else:
                    if lg:
                        self._log(lg, self.L(
                            "⚠️ Panoda kopyalanmış bir ekran görüntüsü veya metin bulunamadı. (Win+Shift+S ile ekran görüntüsü alabilir, Ctrl+C ile soru metnini kopyalayabilir veya '📸 Ekranı Kırp & Sor' butonunu kullanabilirsiniz.)",
                            "⚠️ No screenshot or text found in clipboard.",
                            "⚠️ Kein Screenshot oder Text in der Zwischenablage gefunden."
                        ))

            ctk.CTkButton(
                file_row,
                text=self.L("📸 Ekranı Kırp & Sor", "📸 Snip & Ask", "📸 Screenshot"),
                height=36,
                fg_color="#DC2626",
                hover_color="#B91C1C",
                font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
                command=on_snip_screenshot,
            ).pack(side="left", padx=(0, 6), pady=10)

            ctk.CTkButton(
                file_row,
                text=self.L("📋 Panodan Yapıştır", "📋 Paste", "📋 Einfügen"),
                height=36,
                fg_color="#0D9488",
                hover_color="#0F766E",
                font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
                command=on_paste_clipboard_image,
            ).pack(side="left", padx=(0, 6), pady=10)

            ctk.CTkButton(
                file_row,
                text=self.L("🧠 Soruyu Adım Adım Çöz", "🧠 Solve Step-by-Step", "🧠 Schrittweise lösen"),
                height=36,
                fg_color="#2563EB",
                hover_color="#1D4ED8",
                font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
                command=lambda: ctrls.get("solve_step_fn") and ctrls["solve_step_fn"](),
            ).pack(side="left", padx=(0, 8), pady=10)
            ctrls["paste_img_fn"] = on_paste_clipboard_image

        if seg_type != "CHAT":
            src_lbl_txt = self.L("Kaynak:", "Source:", "Quelle:") if seg_type == "EDU" else self.t("src_lang_lbl")
            ctk.CTkLabel(file_row, text=src_lbl_txt, font=ctk.CTkFont(size=12, weight="bold"), text_color="#334155").pack(side="left", padx=(4, 5))
            ctk.CTkOptionMenu(
                file_row,
                variable=src_var,
                values=[k for k in LANG_MAP.keys() if k != "Sadece OCR (Metne Dönüştür)"],
                width=95 if seg_type == "EDU" else 165,
                height=34,
                fg_color="#F1F5F9",
                text_color="#0F172A",
                button_color="#CBD5E1",
            ).pack(side="left", padx=(0, 6), pady=10)

        if seg_type == "EDU":
            ctk.CTkLabel(file_row, text=self.L("Hedef:", "Target:", "Ziel:"), font=ctk.CTkFont(size=12, weight="bold"), text_color="#1E3A8A").pack(side="left", padx=(2, 5))
            edu_tgt_var = ctk.StringVar(value=self.user_settings.get("def_lang", "Turkish"))
            ctrls["edu_tgt_var"] = edu_tgt_var
            ctk.CTkOptionMenu(
                file_row,
                variable=edu_tgt_var,
                values=EDU_TARGET_LANGS,
                width=145,
                height=34,
                fg_color="#DBEAFE",
                text_color="#1E3A8A",
                button_color="#93C5FD",
            ).pack(side="left", padx=(0, 10), pady=10)

        # EDU Options Row (Optional Terminology Glossary & High-Quality Format-Preserving Settings)
        if seg_type == "EDU":
            edu_box = ctk.CTkFrame(f, fg_color="#FFFFFF", corner_radius=12, border_width=1, border_color="#BFDBFE")
            edu_box.pack(fill="x", padx=25, pady=4)

            # Row 1: Optional Glossary File Selector + Quick Inline Terms
            g_row1 = ctk.CTkFrame(edu_box, fg_color="transparent")
            g_row1.pack(fill="x", padx=12, pady=(8, 4))

            ctk.CTkLabel(
                g_row1,
                text=self.L(
                    "📚 Terim Sözlüğü (İsteğe Bağlı):",
                    "📚 Terminology Glossary (Optional):",
                    "📚 Fachglossar (Optional):",
                ),
                font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
                text_color="#1E3A8A",
            ).pack(side="left", padx=(0, 6))

            glossary_file_var = ctk.StringVar(value="")
            ctrls["glossary_file_var"] = glossary_file_var

            ctk.CTkEntry(
                g_row1,
                textvariable=glossary_file_var,
                state="disabled",
                width=230,
                height=30,
                placeholder_text=self.L("Sözlük seçilmedi (Sözlüksüz çalışır)", "No glossary (Works without glossary)", "Kein Glossar ausgewählt"),
                fg_color="#F8FAFC",
            ).pack(side="left", padx=(0, 6))

            ctk.CTkButton(
                g_row1,
                text=self.L("📂 Sözlük Yükle (.xlsx/.csv/.txt/.docx)", "📂 Load Glossary", "📂 Glossar laden"),
                height=30,
                width=210,
                fg_color="#2563EB",
                hover_color="#1D4ED8",
                font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
                command=lambda: glossary_file_var.set(
                    filedialog.askopenfilename(
                        initialdir=os.path.join(os.path.expanduser("~"), "Documents", "ILBER_Workspace", "Uploads"),
                        filetypes=SUPPORTED_GLOSSARY_EXTS,
                    )
                ),
            ).pack(side="left", padx=(0, 4))

            ctk.CTkButton(
                g_row1,
                text=self.L("✖ Temizle", "✖ Clear", "✖ Löschen"),
                height=30,
                width=70,
                fg_color="#E2E8F0",
                hover_color="#CBD5E1",
                text_color="#334155",
                font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
                command=lambda: glossary_file_var.set(""),
            ).pack(side="left", padx=(0, 12))

            ctk.CTkLabel(
                g_row1,
                text=self.L("Hızlı Terim:", "Quick Terms:", "Schnellbegriffe:"),
                font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
                text_color="#334155",
            ).pack(side="left", padx=(0, 4))

            glossary_inline_entry = ctk.CTkEntry(
                g_row1,
                height=30,
                placeholder_text=self.L(
                    "Örn: Kaufvertrag = Satış Sözleşmesi, Kündigung = Fesih",
                    "e.g. Kaufvertrag = Sales Contract, Kündigung = Termination",
                    "z.B. Kaufvertrag = Satış Sözleşmesi",
                ),
                fg_color="#F8FAFC",
            )
            glossary_inline_entry.pack(side="left", fill="x", expand=True)
            ctrls["glossary_inline_entry"] = glossary_inline_entry

            # Row 2: High-Quality AI Redaction Checkbox + Format Preservation Guarantee Badge
            g_row2 = ctk.CTkFrame(edu_box, fg_color="transparent")
            g_row2.pack(fill="x", padx=12, pady=(2, 8))

            edu_hq_var = ctk.BooleanVar(value=True)
            ctrls["edu_hq_var"] = edu_hq_var
            ctk.CTkCheckBox(
                g_row2,
                text=self.L(
                    "✨ Maksimum Çeviri Kalitesi",
                    "✨ Maximum Translation Quality",
                    "✨ Maximale Übersetzungsqualität",
                ),
                variable=edu_hq_var,
                font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
                text_color="#0F172A",
            ).pack(side="left", padx=(0, 16))

            edu_edit_var = ctk.BooleanVar(value=True)
            ctrls["edu_edit_var"] = edu_edit_var
            ctk.CTkCheckBox(
                g_row2,
                text=self.L(
                    "✏️ Nihai Belge Oluşturulmadan Önce Düzenleme / Kontrol Ekranı Aç",
                    "✏️ Open Review & Edit Screen Before Finalizing Document",
                    "✏️ Vor dem Speichern Bearbeitungsfenster öffnen",
                ),
                variable=edu_edit_var,
                font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
                text_color="#1E3A8A",
            ).pack(side="left")

            ctk.CTkLabel(
                g_row2,
                text=self.L(
                    "✔ Giriş Formatı = Çıkış Formatı (.docx➔.docx · .pdf➔.pdf · .xlsx➔.xlsx)",
                    "✔ Input Format = Output Format (.docx➔.docx · .pdf➔.pdf · .xlsx➔.xlsx)",
                    "✔ Eingabeformat = Ausgabeformat (.docx➔.docx · .pdf➔.pdf · .xlsx➔.xlsx)",
                ),
                font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
                text_color="#059669",
            ).pack(side="right", padx=6)

        # GLOBE Options Row (Multi-Target Language Checkboxes + MP4 Burn toggle)
        if seg_type == "GLOBE":
            globe_box = ctk.CTkFrame(f, fg_color="#FFFFFF", corner_radius=12, border_width=1, border_color="#E2E8F0")
            globe_box.pack(fill="x", padx=25, pady=4)

            g_top = ctk.CTkFrame(globe_box, fg_color="transparent")
            g_top.pack(fill="x", padx=12, pady=(8, 4))
            ctk.CTkLabel(g_top, text=self.t("globe_multi_lbl"), font=ctk.CTkFont(size=12, weight="bold"), text_color="#065F46").pack(side="left")

            burn_mp4_var = ctk.BooleanVar(value=True)
            ctrls["burn_mp4_var"] = burn_mp4_var
            ctk.CTkCheckBox(g_top, text=self.t("globe_burn_mp4"), variable=burn_mp4_var, font=ctk.CTkFont(size=12, weight="bold"), text_color="#1E293B").pack(side="right", padx=10)

            g_langs = ctk.CTkFrame(globe_box, fg_color="transparent")
            g_langs.pack(fill="x", padx=12, pady=(2, 10))
            multi_lang_vars = {}
            def_t = self.user_settings.get("def_lang", "Turkish")
            for t_lang in TARGET_LANGS:
                v = ctk.BooleanVar(value=(t_lang == def_t))
                multi_lang_vars[t_lang] = v
                ctk.CTkCheckBox(g_langs, text=t_lang, variable=v, width=90, font=ctk.CTkFont(size=12)).pack(side="left", padx=4)
            ctrls["multi_lang_vars"] = multi_lang_vars

        # FIELD Metadata Form (Date, Location, Participant Code, Interviewer, Speaker Count, Context, Ethics Translation)
        if seg_type == "FIELD":
            field_box = ctk.CTkFrame(f, fg_color="#FFFFFF", corner_radius=12, border_width=1, border_color="#FECACA")
            field_box.pack(fill="x", padx=25, pady=4)

            r1 = ctk.CTkFrame(field_box, fg_color="transparent")
            r1.pack(fill="x", padx=12, pady=(8, 4))

            ctk.CTkLabel(r1, text=self.t("field_date"), font=ctk.CTkFont(size=12, weight="bold")).pack(side="left", padx=(0, 4))
            f_date = ctk.CTkEntry(r1, width=105, height=28, fg_color="#F8FAFC")
            f_date.insert(0, datetime.date.today().strftime("%Y-%m-%d"))
            f_date.pack(side="left", padx=(0, 12))

            ctk.CTkLabel(r1, text=self.t("field_loc"), font=ctk.CTkFont(size=12, weight="bold")).pack(side="left", padx=(0, 4))
            f_loc = ctk.CTkEntry(r1, width=130, height=28, placeholder_text="İstanbul / Saha", fg_color="#F8FAFC")
            f_loc.pack(side="left", padx=(0, 12))

            ctk.CTkLabel(r1, text=self.t("field_pcode"), font=ctk.CTkFont(size=12, weight="bold")).pack(side="left", padx=(0, 4))
            f_pcode = ctk.CTkEntry(r1, width=75, height=28, fg_color="#F8FAFC")
            f_pcode.insert(0, "P01")
            f_pcode.pack(side="left", padx=(0, 12))

            ctk.CTkLabel(r1, text=self.t("field_interviewer"), font=ctk.CTkFont(size=12, weight="bold")).pack(side="left", padx=(0, 4))
            f_inter = ctk.CTkEntry(r1, width=110, height=28, fg_color="#F8FAFC")
            f_inter.insert(0, "Görüşmeci")
            f_inter.pack(side="left", padx=(0, 12))

            ctk.CTkLabel(r1, text=self.t("field_spk_count"), font=ctk.CTkFont(size=12, weight="bold")).pack(side="left", padx=(0, 4))
            f_spks = ctk.CTkOptionMenu(r1, values=["2", "3", "4", "5"], width=70, height=28, fg_color="#FEE2E2", text_color="#991B1B")
            f_spks.set("2")
            f_spks.pack(side="left")

            r2 = ctk.CTkFrame(field_box, fg_color="transparent")
            r2.pack(fill="x", padx=12, pady=(4, 8))
            ctk.CTkLabel(r2, text=self.t("field_ctx"), font=ctk.CTkFont(size=12, weight="bold")).pack(side="left", padx=(0, 6))
            f_ctx = ctk.CTkEntry(r2, width=360, height=28, placeholder_text="Görüşme ortamı, gözlemler...", fg_color="#F8FAFC")
            f_ctx.pack(side="left", padx=(0, 16))

            f_tr_chk = ctk.BooleanVar(value=False)
            f_tr_tgt = ctk.StringVar(value="English")
            ctk.CTkCheckBox(r2, text=self.t("field_trans_chk"), variable=f_tr_chk, font=ctk.CTkFont(size=12, weight="bold"), text_color="#991B1B").pack(side="left", padx=(4, 8))
            ctk.CTkOptionMenu(r2, variable=f_tr_tgt, values=TARGET_LANGS, width=120, height=28, fg_color="#F1F5F9", text_color="#0F172A").pack(side="left")

            ctrls.update({
                "f_date": f_date,
                "f_loc": f_loc,
                "f_pcode": f_pcode,
                "f_inter": f_inter,
                "f_spks": f_spks,
                "f_ctx": f_ctx,
                "f_tr_chk": f_tr_chk,
                "f_tr_tgt": f_tr_tgt,
            })

        # Bottom Controls Container (packed at bottom FIRST so it is always visible on any screen size)
        bottom_area = ctk.CTkFrame(f, fg_color="transparent")
        bottom_area.pack(side="bottom", fill="x", padx=25, pady=(0, 14))

        # Live AI Thinking / Processing Status Banner (hidden when idle)
        status_banner = ctk.CTkFrame(bottom_area, fg_color="#EFF6FF", border_width=1, border_color="#93C5FD", corner_radius=10)
        status_lbl = ctk.CTkLabel(
            status_banner,
            text=self.t("busy_thinking"),
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color="#1D4ED8",
        )
        status_lbl.pack(side="left", padx=14, pady=6)
        banner_stop_btn = ctk.CTkButton(
            status_banner,
            text=self.L("⏹ Durdur", "⏹ Stop", "⏹ Stopp"),
            width=90,
            height=26,
            corner_radius=6,
            fg_color="#DC2626",
            hover_color="#B91C1C",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            command=lambda pid=page_id: self.stop_pipeline(pid),
        )
        banner_stop_btn.pack(side="right", padx=(4, 12), pady=6)
        status_pbar = ctk.CTkProgressBar(status_banner, mode="indeterminate", width=180, height=10, progress_color="#2563EB")
        status_pbar.pack(side="right", padx=10, pady=8)

        ctrls["status_banner"] = status_banner
        ctrls["status_lbl"] = status_lbl
        ctrls["status_pbar"] = status_pbar

        # Main Log & Output Box (wrap="word" enforced!)
        logbox = ctk.CTkTextbox(
            f,
            height=220,
            wrap="word",
            fg_color="#F9FAFB",
            text_color="#1F2937",
            border_width=1,
            border_color="#E5E7EB",
            font=ctk.CTkFont(family="Segoe UI", size=13),
        )
        logbox.pack(fill="both", expand=True, padx=25, pady=8)
        self.log_boxes[page_id] = logbox

        if seg_type == "CHAT":
            if not hasattr(self, "video_context"):
                self.video_context = ""
            if not hasattr(self, "chat_history"):
                self.chat_history = []
            chat_row = ctk.CTkFrame(bottom_area, fg_color="transparent")
            chat_row.pack(fill="x", pady=(4, 8))

            chat_entry = ctk.CTkEntry(
                chat_row,
                placeholder_text=self.t("chat_placeholder"),
                height=44,
                fg_color="#1F2937",
                text_color="white",
                font=ctk.CTkFont(size=14),
            )
            chat_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))

            send_btn = ctk.CTkButton(
                chat_row,
                text=self.t("chat_send"),
                height=44,
                width=110,
                fg_color="#4C1D95",
                hover_color="#5B21B6",
                font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            )
            send_btn.pack(side="left")

            def _send_question_to_chat(user_q, show_user_msg=True):
                if not user_q:
                    return
                if not self.license_mgr.is_module_unlocked(page_id):
                    self.open_package_upgrade_modal(locked_pid=page_id)
                    return
                if self.license_mgr.session.get("plan") == "trial":
                    rem = self.license_mgr.consume_trial_credit()
                    self.after(0, self.refresh_license_ui)
                    self.after(0, lambda r=rem: self._log(logbox, self.L(
                        f"⏳ [Ücretsiz Deneme Sürümü — Kalan İşlem Hakkınız: {r}/5]",
                        f"⏳ [Free Trial Version — Remaining Credits: {r}/5]",
                        f"⏳ [Kostenlose Testversion — Verbleibende Vorgänge: {r}/5]",
                    )))
                you_lbl = self.L("Siz", "You", "Sie")
                if show_user_msg:
                    self.after(0, lambda: self._log(logbox, f"\n{you_lbl}: {user_q}"))
                else:
                    self.after(0, lambda: self._log(logbox, ""))
                self.after(0, lambda: self._log(logbox, self.t("busy_thinking")))

                def process_chat():
                    self.after(0, lambda: send_btn.configure(state="disabled", text=self.t("btn_wait_short")))
                    self._start_busy_indicator(page_id, self.t("busy_chat"))
                    try:
                        active_lang = resolve_response_language(
                            self.ui_lang,
                            user_q if show_user_msg else "",
                            getattr(self, "video_context", ""),
                        )
                        if active_lang == "en":
                            prompt = (
                                "You are an experienced, patient, warm, and professional private tutor.\n"
                                "CRITICAL LANGUAGE RULE: You MUST respond 100% in ENGLISH. Never respond in Turkish!\n"
                                "Your goal is not just to state the right option, but to teach your student the underlying concept, practical rule, and common exam traps.\n"
                                "NEVER use raw LaTeX (\\( \\), \\[ \\], \\frac, ^{ }, $x$) or raw Markdown (** or #)! Write clean, readable math like '2¹² / 2³ = 2⁹' or 'R(x) = 2x + 3'.\n\n"
                                "ALWAYS STRUCTURE YOUR EXPLANATION WITH THESE 5 SECTIONS:\n"
                                "🎯 1. Understanding the Question (What Is Being Asked?)\n"
                                "💡 2. Teacher's Note (Key Rule & Practical Method)\n"
                                "✍️ 3. Step-by-Step Solution (Let's Solve It Together)\n"
                                "⚠️ 4. Watch Out for Distractor Options (Why Not the Others?)\n"
                                "✅ 5. Final Correct Answer:\n"
                                "CORRECT ANSWER: [Letter]) [Value]"
                            )
                        elif active_lang == "de":
                            prompt = (
                                "Du bist ein sehr erfahrener, geduldiger, herzlicher und professioneller Nachhilfelehrer.\n"
                                "WICHTIGE SPRACHREGEL: Antworte zu 100% auf DEUTSCH. Antworte niemals auf Türkisch!\n"
                                "Dein Ziel ist es, dem Schüler die Logik, die praktische Regel und typische Prüfungsfallen verständlich beizubringen.\n"
                                "Verwende NIEMALS LaTeX-Code (\\( \\), \\[ \\], \\frac, ^{ }, $x$) oder Markdown (** oder #)!\n\n"
                                "VERWENDE IMMER DIESE 5 ABSCHNITTE:\n"
                                "🎯 1. Die Aufgabe verstehen (Was wird gefragt?)\n"
                                "💡 2. Lehrer-Tipp (Wichtige Regel & Praktische Methode)\n"
                                "✍️ 3. Schritt-für-Schritt-Lösung (Gemeinsam lösen)\n"
                                "⚠️ 4. Achtung vor Fallen in den Antwortmöglichkeiten\n"
                                "✅ 5. Endgültige richtige Antwort:\n"
                                "RICHTIGE ANTWORT: [Buchstabe]) [Wert]"
                            )
                        else:
                            prompt = (
                                "Sen öğrencisine özel ders veren çok deneyimli, sabırlı, sıcak kanlı ve profesyonel bir öğretmensin.\n"
                                "Amacın sadece doğru şıkkı söylemek değil; karşındaki öğrencinin konunun mantığını, kuralını ve sınav tuzaklarını tam olarak öğrenmesini sağlamaktır.\n\n"
                                "ÖĞRETMEN ANLATIM KURALLARI (BU KURALLARA HARFİYEN UY):\n"
                                "1. ÜSLUP VE YAZIM:\n"
                                "   - Öğrencinle konuşur gibi anlaşılır, motive edici ve akıcı bir öğretmen dili kullan.\n"
                                "   - ASLA \\( \\), \\[ \\], \\frac{...}{...}, ^{...} veya $x$ gibi LaTeX kodları veya ham Markdown (** veya #) kullanma!\n"
                                "2. ÇÖZÜM ŞABLONU (HER ZAMAN BU 5 BAŞLIKLA ANLAT):\n"
                                "   🎯 1. Soruyu Anlayalım (Bizden Ne İsteniyor?)\n"
                                "   💡 2. Öğretmen Notu (Bilmemiz Gereken Kural & Pratik Yöntem)\n"
                                "   ✍️ 3. Adım Adım Çözüm (Beraber Çözelim)\n"
                                "   ⚠️ 4. Çeldirici Şıklara Dikkat (Neden Diğer Şıklar Değil?)\n"
                                "   ✅ 5. Kesin Doğru Cevap:\n"
                                "   DOĞRU CEVAP: [Şık]) [Değer]"
                            )

                        # First check if our 100% deterministic offline math/logic teacher engine solves the question directly
                        target_q_src = self.video_context[:12000] if getattr(self, "video_context", "") else user_q
                        try:
                            exact_teacher_ans = solve_question_with_teacher_engine_offline(target_q_src, lang=active_lang)
                        except Exception:
                            exact_teacher_ans = None
                        if exact_teacher_ans:
                            ans = exact_teacher_ans
                        else:
                            if getattr(self, "video_context", ""):
                                structured_doc = analyze_question_structure_offline(self.video_context[:12000], lang=active_lang)
                                if active_lang == "en":
                                    context_query = (
                                        f"{structured_doc}\n\n"
                                        f"[STUDENT'S QUESTION / REQUEST]:\n{user_q}\n\n"
                                        f"(IMPORTANT: Explain strictly in ENGLISH like a warm, professional private tutor. Do not use LaTeX.)"
                                    )
                                elif active_lang == "de":
                                    context_query = (
                                        f"{structured_doc}\n\n"
                                        f"[FRAGE DES SCHÜLERS]:\n{user_q}\n\n"
                                        f"(WICHTIG: Erkläre ausschließlich auf DEUTSCH wie ein professioneller Nachhilfelehrer. Verwende kein LaTeX.)"
                                    )
                                else:
                                    context_query = (
                                        f"{structured_doc}\n\n"
                                        f"[ÖĞRENCİNİN SORUSU / İSTEĞİ]:\n{user_q}\n\n"
                                        f"(ÖNEMLİ: Karşındaki öğrenciye profesyonel bir öğretmen gibi sıcak, anlaşılır ve öğretici bir dille Türkçe anlat. LaTeX kodu kullanma.)"
                                    )
                            else:
                                context_query = analyze_question_structure_offline(user_q, lang=active_lang)

                            hist = list(getattr(self, "chat_history", []))[-4:] if show_user_msg else []
                            raw_ans = self._run_llm(prompt, context_query, max_tokens=1100, temperature=0.15, history_messages=hist)
                            ans = clean_markdown_text(raw_ans)
                            if getattr(self, "video_context", ""):
                                ans = verify_and_fix_answer_option(self.video_context, ans, lang=active_lang)
                            else:
                                ans = verify_and_fix_answer_option(user_q, ans, lang=active_lang)

                        self.chat_history.append({"role": "user", "content": user_q})
                        self.chat_history.append({"role": "assistant", "content": ans})
                        self.after(0, lambda: self._log(logbox, f"ILBER AI:\n{ans}\n"))
                    except Exception as e:
                        err_msg = f"❌ Hata: {e}"
                        self.after(0, lambda m=err_msg: self._log(logbox, m))
                    finally:
                        self._stop_busy_indicator(page_id)
                        self.after(0, lambda: send_btn.configure(state="normal", text=self.t("chat_send")))

                threading.Thread(target=process_chat, daemon=True).start()

            def on_global_chat_send(*args):
                user_q = chat_entry.get().strip()
                if not user_q:
                    return
                chat_entry.delete(0, "end")
                _send_question_to_chat(user_q, show_user_msg=True)

            def on_solve_step_by_step():
                if not getattr(self, "video_context", "").strip():
                    self._log(logbox, self.L(
                        "⚠️ Henüz hafızaya alınmış bir soru veya belge yok. Lütfen önce '📸 Ekranı Kırp & Sor', '📋 Panodan Yapıştır' veya '📂 Dosya Seç' ile bir soru yükleyin.",
                        "⚠️ No question or document in memory yet. Please take a screenshot, paste an image, or select a file first.",
                        "⚠️ Noch keine Frage oder kein Dokument im Speicher. Bitte zuerst Screenshot aufnehmen oder Datei auswählen."
                    ))
                    return
                q_txt = self.L(
                    "Bu soruyu bir öğrenciye özel ders veren profesyonel bir öğretmen gibi adım adım, mantığını ve pratik kuralını öğreterek çöz.",
                    "Explain and solve this question step by step like a professional teacher tutoring a student, teaching the logic and practical rules.",
                    "Erklären und lösen Sie diese Frage Schritt für Schritt wie ein professioneller Lehrer im Einzelunterricht."
                )
                _send_question_to_chat(q_txt, show_user_msg=False)

            ctrls["solve_step_fn"] = on_solve_step_by_step
            send_btn.configure(command=on_global_chat_send)
            chat_entry.bind("<Return>", on_global_chat_send)

            action_row = ctk.CTkFrame(bottom_area, fg_color="transparent")
            action_row.pack(pady=(0, 2))
            start_btn = ctk.CTkButton(
                action_row,
                text=self.t("chat_memorize"),
                height=42,
                fg_color="#10B981",
                hover_color="#059669",
                font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
                command=lambda: self.start_pipeline(page_id, file_var.get(), src_var.get(), "Turkish", start_btn),
            )
            start_btn.pack(side="left", padx=(0, 10))
            stop_btn = ctk.CTkButton(
                action_row,
                text=self.L("⏹ İşlemi Durdur", "⏹ Stop Process", "⏹ Prozess stoppen"),
                height=42,
                width=155,
                state="disabled",
                fg_color="#94A3B8",
                hover_color="#B91C1C",
                text_color_disabled="#F1F5F9",
                font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
                command=lambda pid=page_id: self.stop_pipeline(pid),
            )
            stop_btn.pack(side="left")
            ctrls["start_btn"] = start_btn
            ctrls["stop_btn"] = stop_btn
            ctrls["action_row"] = action_row
        else:
            action_row = ctk.CTkFrame(bottom_area, fg_color="transparent")
            action_row.pack(pady=(4, 2))
            start_btn = ctk.CTkButton(
                action_row,
                text=self.t("btn_start_proc"),
                height=45,
                width=210,
                fg_color="#4C1D95",
                hover_color="#5B21B6",
                font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
                command=lambda: self.start_pipeline(page_id, file_var.get(), src_var.get(), "Turkish", start_btn),
            )
            start_btn.pack(side="left", padx=(0, 12))
            stop_btn = ctk.CTkButton(
                action_row,
                text=self.L("⏹ İşlemi Durdur", "⏹ Stop Process", "⏹ Prozess stoppen"),
                height=45,
                width=170,
                state="disabled",
                fg_color="#94A3B8",
                hover_color="#B91C1C",
                text_color_disabled="#F1F5F9",
                font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
                command=lambda pid=page_id: self.stop_pipeline(pid),
            )
            stop_btn.pack(side="left")
            ctrls["start_btn"] = start_btn
            ctrls["stop_btn"] = stop_btn
            ctrls["action_row"] = action_row

    def show_page(self, name):
        if name in ("tool_ses_kaydet", "tool_ai_ozet", "tool_video_cevir", "tool_ai_rapor"):
            if not self.license_mgr.is_module_unlocked(name):
                self.open_package_upgrade_modal(locked_pid=name)
                return
        self.current_page = name
        for p in self.pages.values():
            p.pack_forget()
        self.pages[name].pack(fill="both", expand=True)
        for pid, b in self.nav_btns.items():
            b.configure(fg_color="#252D41" if pid == name else "transparent")
        if name == "home":
            self._build_home()
        elif name == "files":
            self._build_files()
        elif name == "history":
            self._build_history()
        elif name == "settings":
            self._build_settings()

    def ask_user_edit_text(self, text_content, title_text=None):
        editor = ctk.CTkToplevel(self)
        editor.title(title_text if title_text else self.t("pop_edit_title"))
        editor.geometry("880x660")
        editor.attributes("-topmost", True)
        editor.protocol("WM_DELETE_WINDOW", lambda: on_save())

        ctk.CTkLabel(editor, text=self.t("pop_edit_msg"), font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold")).pack(pady=(12, 6), padx=20, anchor="w")

        textbox = ctk.CTkTextbox(editor, width=840, height=520, wrap="word", font=ctk.CTkFont(family="Segoe UI", size=13))
        textbox.pack(pady=6, padx=20, fill="both", expand=True)
        textbox.insert("1.0", clean_markdown_text(text_content))

        def on_save():
            self.edited_text = clean_markdown_text(textbox.get("1.0", "end-1c"))
            editor.destroy()
            if hasattr(self, "user_edit_done"):
                self.user_edit_done.set()

        ctk.CTkButton(editor, text=self.t("pop_save"), fg_color="#10B981", hover_color="#059669", height=44, font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"), command=on_save).pack(fill="x", padx=20, pady=12)

    def open_interactive_segment_editor_sync(self, segments, media_path, title_text, desc_text, btn_text=None):
        """Open InteractiveSegmentEditor on main UI thread and block worker thread until user saves."""
        done_event = threading.Event()
        result_holder = [segments]

        def _launch():
            def _on_done(edited_segs):
                result_holder[0] = edited_segs
                done_event.set()

            InteractiveSegmentEditor(
                self,
                segments=segments,
                media_path=media_path,
                ui_lang=self.ui_lang,
                title_text=title_text,
                desc_text=desc_text,
                btn_text=btn_text or self.t("pop_save"),
                on_complete=_on_done,
            )

        self.after(0, _launch)
        done_event.wait()
        return result_holder[0]

    def open_doc_translation_editor_sync(self, source_texts, translated_texts, doc_name=""):
        """Open DocumentTranslationEditor on main UI thread and block worker thread until user approves/saves."""
        import importlib
        import ilber_ui_dialogs as _iud
        importlib.reload(_iud)
        done_event = threading.Event()
        result_holder = [list(translated_texts)]

        def _launch():
            def _on_done(edited_list):
                result_holder[0] = edited_list
                done_event.set()

            _iud.DocumentTranslationEditor(
                self,
                source_texts=source_texts,
                translated_texts=translated_texts,
                doc_name=doc_name,
                ui_lang=self.ui_lang,
                on_complete=_on_done,
            )

        self.after(0, _launch)
        done_event.wait()
        return result_holder[0]

    def open_field_studio_sync(self, segments, suggestions, anon_candidates, media_path, default_p_code, default_interviewer):
        """Open FieldResearchStudio on main UI thread and block worker thread until researcher saves."""
        done_event = threading.Event()
        result_holder = [segments, suggestions, anon_candidates]

        def _launch():
            def _on_done(ed_segs, ed_suggs, ed_anons):
                result_holder[0] = ed_segs
                result_holder[1] = ed_suggs
                result_holder[2] = ed_anons
                done_event.set()

            FieldResearchStudio(
                self,
                segments=segments,
                suggestions=suggestions,
                anon_candidates=anon_candidates,
                media_path=media_path,
                ui_lang=self.ui_lang,
                default_p_code=default_p_code,
                default_interviewer=default_interviewer,
                on_complete=_on_done,
            )

        self.after(0, _launch)
        done_event.wait()
        return result_holder[0], result_holder[1], result_holder[2]

    def _configure_rich_log_tags(self, tb):
        if getattr(tb, "_rich_tags_ready", False):
            return
        try:
            tb.configure(padx=18, pady=14, spacing1=2, spacing2=3, spacing3=3)
            tb.tag_config("sec_1", font=("Segoe UI", 13, "bold"), foreground="#1E40AF", background="#DBEAFE", spacing1=12, spacing3=6, lmargin1=8, lmargin2=8)
            tb.tag_config("sec_2", font=("Segoe UI", 13, "bold"), foreground="#92400E", background="#FEF3C7", spacing1=12, spacing3=6, lmargin1=8, lmargin2=8)
            tb.tag_config("sec_3", font=("Segoe UI", 13, "bold"), foreground="#3730A3", background="#E0E7FF", spacing1=12, spacing3=6, lmargin1=8, lmargin2=8)
            tb.tag_config("sec_4", font=("Segoe UI", 13, "bold"), foreground="#9A3412", background="#FFEDD5", spacing1=12, spacing3=6, lmargin1=8, lmargin2=8)
            tb.tag_config("sec_5", font=("Segoe UI", 13, "bold"), foreground="#065F46", background="#D1FAE5", spacing1=12, spacing3=6, lmargin1=8, lmargin2=8)
            tb.tag_config("ans_banner", font=("Segoe UI", 15, "bold"), foreground="#065F46", background="#A7F3D0", spacing1=6, spacing3=12, lmargin1=12, lmargin2=12)
            tb.tag_config("step_title", font=("Segoe UI", 12, "bold"), foreground="#0F172A", spacing1=8, spacing3=3, lmargin1=12, lmargin2=20)
            tb.tag_config("math_line", font=("Segoe UI Semibold", 12), foreground="#1D4ED8", background="#F1F5F9", spacing1=2, spacing3=2, lmargin1=26, lmargin2=26)
            tb.tag_config("rule_bullet", font=("Segoe UI", 12), foreground="#334155", spacing1=3, spacing3=3, lmargin1=14, lmargin2=24)
            tb.tag_config("opt_correct", font=("Segoe UI", 12, "bold"), foreground="#047857", background="#ECFDF5", spacing1=4, spacing3=4, lmargin1=14, lmargin2=24)
            tb.tag_config("opt_elim", font=("Segoe UI", 11), foreground="#64748B", spacing1=2, spacing3=2, lmargin1=14, lmargin2=24)
            tb.tag_config("preview_hdr", font=("Segoe UI", 12, "bold"), foreground="#0F766E", background="#F0FDFA", spacing1=6, spacing3=4)
            tb.tag_config("ai_badge", font=("Segoe UI", 13, "bold"), foreground="#4C1D95", spacing1=10, spacing3=4)
            tb.tag_config("user_badge", font=("Segoe UI", 12, "bold"), foreground="#1E3A8A", background="#EFF6FF", spacing1=8, spacing3=4)
            tb.tag_config("body_txt", font=("Segoe UI", 12), foreground="#1E293B", spacing1=2, spacing2=3, spacing3=3, lmargin1=10, lmargin2=16)
            tb._rich_tags_ready = True
        except Exception:
            pass

    def _log(self, box, txt):
        tb = getattr(box, "_textbox", None)
        if tb is None:
            box.insert("end", txt + "\n")
            box.see("end")
            return
        self._configure_rich_log_tags(tb)
        for raw_ln in (txt + "\n").splitlines(True):
            s = raw_ln.strip()
            tag = "body_txt"
            if s.startswith("🎯") or re.match(r"^1\.\s+(Soruyu|Understanding|Die Aufgabe)", s):
                tag = "sec_1"
                raw_ln = "  " + s + "  \n"
            elif s.startswith("💡") or re.match(r"^2\.\s+(Öğretmen|Teacher|Lehrer)", s):
                tag = "sec_2"
                raw_ln = "  " + s + "  \n"
            elif s.startswith("✍️") or re.match(r"^3\.\s+(Adım|Step-by-Step|Schritt)", s):
                tag = "sec_3"
                raw_ln = "  " + s + "  \n"
            elif s.startswith("⚠️ 4.") or re.match(r"^4\.\s+(Çeldirici|Watch Out|Achtung)", s):
                tag = "sec_4"
                raw_ln = "  " + s + "  \n"
            elif s.startswith("✅") or re.match(r"^5\.\s+(Kesin|Final|Endgültige)", s):
                tag = "sec_5"
                raw_ln = "  " + s + "  \n"
            elif re.match(r"^(DOĞRU CEVAP|CORRECT ANSWER|RICHTIGE ANTWORT)\s*:", s, re.I):
                tag = "ans_banner"
                raw_ln = f"   🏆  {s}   \n"
            elif re.search(r"➔\s*(DOĞRU|CORRECT|RICHTIG)\b", s) or (re.match(r"^[•\-]?\s*(Option\s+[A-E]|[A-E]\s+şıkkı)", s, re.I) and re.search(r"\b(matches|correct|doğru|richtig)\b", s, re.I)):
                tag = "opt_correct"
                raw_ln = "  ✔ " + s.lstrip("•- ").strip() + "\n"
            elif re.search(r"➔\s*(Elenir|Eliminated|Ausgeschlossen)\b", s) or re.match(r"^[•\-]?\s*(Option\s+[A-E]|[A-E]\s+şıkkı)", s, re.I):
                tag = "opt_elim"
                raw_ln = "  ✖ " + s.lstrip("•- ").strip() + "\n"
            elif re.match(r"^(•\s*)?(Step\s*\d+|\d+\.\s*(Adım|Yöntem|Schritt)|🔹\s*\d+\.)", s, re.I):
                tag = "step_title"
            elif s.startswith(("📄", "✔ İÇERİK", "✔ CONTENT", "✔ INHALT")):
                tag = "preview_hdr"
            elif s.startswith("ILBER AI:"):
                tag = "ai_badge"
            elif re.match(r"^(Siz|You|Sie)\s*:", s):
                tag = "user_badge"
            elif s.startswith("-") and any(sym in s for sym in ("=", "➔", "÷", "×", "P(", "R(", "f(", "f'(", "lim")):
                tag = "math_line"
            elif not s.startswith(("•", "-", "Since", "So,", "Therefore")) and any(sym in s for sym in ("=", "➔", "÷")) and len(s) < 62:
                tag = "math_line"
            elif s.startswith("•"):
                tag = "rule_bullet"
            tb.insert("end", raw_ln, tag)
        box.see("end")

    def toggle_recording(self, btn, page_id):
        ctrls = self.page_controls.get(page_id, {})
        status_lbl = ctrls.get("mic_status_lbl")
        if not self.is_recording:
            self.is_recording = True
            btn.configure(text=self.t("btn_record_stop"), fg_color="#0F766E")
            if status_lbl:
                status_lbl.configure(text=self.t("mic_hint_rec"), text_color="#DC2626")
            self.audio_frames = []

            def record_thread():
                fs = 16000
                try:
                    import sounddevice as sd
                    with sd.InputStream(samplerate=fs, channels=1, dtype="int16", blocksize=2048) as stream:
                        while self.is_recording:
                            data, _ = stream.read(2048)
                            self.audio_frames.append(data.copy().reshape(-1))
                except Exception as e:
                    self.after(0, lambda err=str(e): self._log(self.log_boxes[page_id], f"❌ {err}"))

            self.rec_thread_obj = threading.Thread(target=record_thread, daemon=True)
            self.rec_thread_obj.start()
            self._log(self.log_boxes[page_id], self.t("log_wait_mic"))
        else:
            self.is_recording = False
            if hasattr(self, "rec_thread_obj") and self.rec_thread_obj.is_alive():
                self.rec_thread_obj.join(timeout=1.0)
            btn.configure(text=self.t("btn_record"), fg_color=RED)
            if status_lbl:
                status_lbl.configure(text=self.t("mic_hint_done"), text_color="#059669")

            if not self.audio_frames:
                self._log(self.log_boxes[page_id], "❌ Audio frame empty.")
                return

            audio_data = np.concatenate(self.audio_frames).astype(np.float32)
            # Normalize microphone gain so quiet laptop mics are heard clearly by Whisper
            peak = float(np.max(np.abs(audio_data))) if len(audio_data) else 0.0
            if peak > 50.0:
                audio_data = np.clip(audio_data * (26000.0 / peak), -32767.0, 32767.0)
            audio_int16 = audio_data.astype(np.int16)

            ws_up = os.path.join(os.path.expanduser("~"), "Documents", "ILBER_Workspace", "Uploads")
            os.makedirs(ws_up, exist_ok=True)
            out_wav = os.path.join(ws_up, f"Mikrofon_Kaydi_{datetime.datetime.now().strftime('%H%M%S')}.wav")
            from scipy.io.wavfile import write as write_wav
            write_wav(out_wav, 16000, audio_int16)
            if "file_var" in ctrls:
                ctrls["file_var"].set(out_wav)
            start_btn = ctrls.get("start_btn", btn)
            src_val = ctrls.get("src_var").get() if "src_var" in ctrls else "Auto (Otomatik)"
            self.start_pipeline(page_id, out_wav, src_val, "Turkish", start_btn)

    def _start_busy_indicator(self, page_id, base_msg=None):
        if base_msg is None:
            base_msg = self.t("busy_thinking")
        ctrls = self.page_controls.get(page_id, {})
        banner = ctrls.get("status_banner")
        lbl = ctrls.get("status_lbl")
        pbar = ctrls.get("status_pbar")
        if not banner or not lbl or not pbar:
            return
        ctrls["busy_active"] = True
        ctrls["busy_start_ts"] = time.time()
        ctrls["busy_base_msg"] = base_msg

        def _show_and_tick():
            if not ctrls.get("busy_active"):
                return
            try:
                before_w = ctrls.get("action_row") or ctrls.get("start_btn")
                banner.pack(fill="x", pady=(0, 6), before=before_w)
            except Exception:
                banner.pack(fill="x", pady=(0, 6))
            pbar.start()

            def _tick():
                if not ctrls.get("busy_active") or not lbl.winfo_exists():
                    return
                elapsed = int(time.time() - ctrls.get("busy_start_ts", time.time()))
                lbl.configure(text=f"{ctrls.get('busy_base_msg')} ({elapsed} {self.t('sec_unit')})")
                self.after(1000, _tick)

            _tick()

        self.after(0, _show_and_tick)

    def _stop_busy_indicator(self, page_id):
        ctrls = self.page_controls.get(page_id, {})
        ctrls["busy_active"] = False
        banner = ctrls.get("status_banner")
        pbar = ctrls.get("status_pbar")

        def _hide():
            try:
                if pbar and pbar.winfo_exists():
                    pbar.stop()
                if banner and banner.winfo_exists():
                    banner.pack_forget()
            except Exception:
                pass

        self.after(0, _hide)

    def start_pipeline(self, mode, filepath, src, tgt, btn):
        # If user is currently recording on MIC and clicks Start Process, stop recording & transcribe that recording
        if self.is_recording:
            ctrls = self.page_controls.get(mode, {})
            rec_btn = ctrls.get("rec_btn", btn)
            self.toggle_recording(rec_btn, mode)
            return

        if not self.license_mgr.is_module_unlocked(mode):
            self.open_package_upgrade_modal(locked_pid=mode)
            return

        logbox = self.log_boxes[mode]
        if not filepath:
            if mode == "tool_ai_ozet":
                self.after(0, lambda: self._log(logbox, self.L(
                    "❌ Lütfen önce çevrilecek bir belge (.docx, .pdf, .xlsx veya .txt) seçin!",
                    "❌ Please select a document (.docx, .pdf, .xlsx, or .txt) to translate first!",
                    "❌ Bitte wählen Sie zuerst ein Dokument (.docx, .pdf, .xlsx oder .txt) aus!",
                )))
            else:
                self.after(0, lambda: self._log(logbox, self.t("log_err_file")))
            return
        if self.is_processing:
            self.after(0, lambda: self._log(logbox, self.t("btn_processing")))
            return

        # Mandatory Ethics Board Warning for FIELD if translation is checked
        ctrls = self.page_controls.get(mode, {})
        if ctrls.get("seg_type") == "FIELD" and ctrls.get("f_tr_chk") and ctrls["f_tr_chk"].get():
            ok = messagebox.askyesno(self.t("field_ethics_title"), self.t("field_ethics_msg"))
            if not ok:
                return

        if self.license_mgr.session.get("plan") == "trial":
            rem = self.license_mgr.consume_trial_credit()
            self.after(0, self.refresh_license_ui)
            self.after(0, lambda r=rem: self._log(logbox, self.L(
                f"⏳ [Ücretsiz Deneme Sürümü — Kalan İşlem Hakkınız: {r}/5]",
                f"⏳ [Free Trial Version — Remaining Credits: {r}/5]",
                f"⏳ [Kostenlose Testversion — Verbleibende Vorgänge: {r}/5]",
            )))

        self.cancel_requested = False
        self._current_run_id = getattr(self, "_current_run_id", 0) + 1
        run_id = self._current_run_id
        self.is_processing = True
        btn.configure(state="disabled", text=self.t("btn_processing"))
        stop_btn = ctrls.get("stop_btn")
        if stop_btn and stop_btn.winfo_exists():
            stop_btn.configure(
                state="normal",
                fg_color="#DC2626",
                text=self.L("⏹ İşlemi Durdur", "⏹ Stop Process", "⏹ Prozess stoppen"),
            )
        self._start_busy_indicator(mode, self.t("busy_thinking"))
        self._log(logbox, "-" * 52)
        threading.Thread(target=self._run_engine, args=(mode, filepath, src, tgt, logbox, btn, run_id), daemon=True).start()

    def stop_pipeline(self, mode=None):
        """Immediately stops any running transcription, translation, OCR, or video processing across MIC/EDU/GLOBE/FIELD."""
        self.cancel_requested = True
        self.is_processing = False
        self._current_run_id = getattr(self, "_current_run_id", 0) + 1
        if self.is_recording:
            self.is_recording = False

        # Terminate any active child ffmpeg process if running
        try:
            if os.name == "nt":
                subprocess.run(["taskkill", "/F", "/IM", "ffmpeg.exe", "/T"], capture_output=True, creationflags=0x08000000)
            else:
                subprocess.run(["pkill", "-f", "ffmpeg"], capture_output=True)
        except Exception:
            pass

        target_modes = [mode] if mode else list(self.page_controls.keys())
        for m in target_modes:
            self._stop_busy_indicator(m)
            ctrls = self.page_controls.get(m, {})
            s_btn = ctrls.get("start_btn")
            st_btn = ctrls.get("stop_btn")
            if s_btn and s_btn.winfo_exists():
                s_btn.configure(
                    state="normal",
                    text=self.t("chat_memorize") if m == "tool_ai_chat" else self.t("btn_start_proc"),
                )
            if st_btn and st_btn.winfo_exists():
                st_btn.configure(
                    state="disabled",
                    fg_color="#94A3B8",
                    text=self.L("⏹ İşlemi Durdur", "⏹ Stop Process", "⏹ Prozess stoppen"),
                )

        active_log = self.log_boxes.get(mode) if mode else None
        if active_log and active_log.winfo_exists():
            self._log(active_log, self.L(
                "🛑 İşlem kullanıcı tarafından durduruldu. Yeni bir dosya seçebilir veya işlemi tekrar başlatabilirsiniz.",
                "🛑 Process stopped by user. You can select a new file or start again immediately.",
                "🛑 Prozess vom Benutzer gestoppt. Sie können sofort eine neue Datei auswählen oder neu starten.",
            ))

    def _run_llm(self, sys_prompt, user_text, max_tokens=800, temperature=0.1, history_messages=None):
        try:
            import urllib.request
            from llama_cpp import Llama

            workspace_dir = os.path.join(os.path.expanduser("~"), "Documents", "ILBER_Workspace", "Models")
            os.makedirs(workspace_dir, exist_ok=True)
            model_path = os.path.join(workspace_dir, "Qwen2.5-3B-Instruct-Q4_K_M.gguf")

            if not os.path.exists(model_path):
                url = "https://huggingface.co/bartowski/Qwen2.5-3B-Instruct-GGUF/resolve/main/Qwen2.5-3B-Instruct-Q4_K_M.gguf"
                urllib.request.urlretrieve(url, model_path)

            if self.llm_instance is None:
                self.llm_instance = Llama(
                    model_path=model_path,
                    n_ctx=8192,
                    n_threads=max(1, (os.cpu_count() or 4) - 1),
                    verbose=False,
                )

            # Preserve line breaks (essential for multiple-choice options A/B/C/D/E and tables!)
            short_text = user_text[:11000]

            messages = [{"role": "system", "content": sys_prompt}]
            if history_messages:
                for hm in history_messages[-6:]:
                    messages.append({"role": hm.get("role", "user"), "content": str(hm.get("content", ""))[:1800]})
            messages.append({"role": "user", "content": short_text})

            output = self.llm_instance.create_chat_completion(
                messages=messages,
                max_tokens=max_tokens,
                temperature=temperature,
                repeat_penalty=1.12,
            )
            return output["choices"][0]["message"]["content"]
        except Exception as e:
            return f"AI Hatasi: {str(e)}"

    def _ensure_m2m100_loaded(self, logbox):
        if self.trans_model is None:
            self.after(0, lambda: self._log(logbox, self.t("log_trans_loading")))
            from transformers import M2M100ForConditionalGeneration, M2M100Tokenizer
            self.trans_model = M2M100ForConditionalGeneration.from_pretrained("facebook/m2m100_418M")
            self.trans_tokenizer = M2M100Tokenizer.from_pretrained("facebook/m2m100_418M")
            self.after(0, lambda: self._log(logbox, self.t("log_trans_loaded")))

    def _translate_lines_m2m100(self, lines_arr, src_code, tgt_code, logbox, label_prefix=""):
        if not lines_arr:
            return []
        if src_code == tgt_code:
            return list(lines_arr)
        self._ensure_m2m100_loaded(logbox)
        supported_m2m = {"en", "tr", "de", "es", "fr", "it", "ru", "ar", "zh", "ja", "ko", "pt", "nl"}
        s_code = src_code if src_code in supported_m2m else "en"
        t_code = tgt_code if tgt_code in supported_m2m else "tr"
        if s_code == t_code:
            return list(lines_arr)

        self.trans_tokenizer.src_lang = s_code
        bos_id = self.trans_tokenizer.get_lang_id(t_code)

        # Split long multi-sentence paragraphs into individual sentences so M2M100 never drops sentences or loops
        flat_sentences = []
        line_sentence_ranges = []
        for raw_line in lines_arr:
            clean_line = raw_line.replace("\n", " ").strip()
            if not clean_line:
                line_sentence_ranges.append((0, 0))
                continue
            if len(clean_line) > 110:
                sents = [s.strip() for s in re.split(r"(?<=[.!?])\s+", clean_line) if s.strip()]
            else:
                sents = [clean_line]
            start_idx = len(flat_sentences)
            flat_sentences.extend(sents)
            line_sentence_ranges.append((start_idx, len(flat_sentences)))

        translated_flat = []
        chunk_size = 4
        total_sents = max(1, len(flat_sentences))
        for i in range(0, len(flat_sentences), chunk_size):
            if getattr(self, "cancel_requested", False):
                raise RuntimeError("__USER_CANCELLED__")
            batch = [s if s else " " for s in flat_sentences[i : i + chunk_size]]
            encoded = self.trans_tokenizer(batch, return_tensors="pt", padding=True, truncation=True, max_length=256)
            gen_tokens = self.trans_model.generate(
                **encoded,
                forced_bos_token_id=bos_id,
                max_new_tokens=256,
                no_repeat_ngram_size=3,
                repetition_penalty=1.25,
            )
            t_chunks = self.trans_tokenizer.batch_decode(gen_tokens, skip_special_tokens=True)
            translated_flat.extend([c.strip() for c in t_chunks])
            progress = min(100, int((i + chunk_size) / total_sents * 100))
            self.after(0, lambda p=progress, lp=label_prefix: self._log(logbox, f"      -> {lp} {self.t('log_trans_prog')} ({s_code.upper()} ➔ {t_code.upper()}): %{p}"))

        t_lines_arr = []
        for start_idx, end_idx in line_sentence_ranges:
            if start_idx == end_idx:
                t_lines_arr.append("")
            else:
                t_lines_arr.append(" ".join(translated_flat[start_idx:end_idx]))
        return t_lines_arr

    def _transcribe_media(self, filepath, src_selection, logbox, verbatim=False):
        whisper = _lazy_import_whisper()
        if self.model is None:
            self.after(0, lambda: self._log(logbox, self.t("log_loading")))
            try:
                self.model = whisper.load_model("medium")
            except Exception as e_load:
                # If offline on first run, check if any local speech model (.pt) is already cached on disk
                cache_dir = os.path.join(os.path.expanduser("~"), ".cache", "whisper")
                loaded_fallback = False
                for fallback_name in ("medium", "small", "base", "tiny"):
                    pt_file = os.path.join(cache_dir, f"{fallback_name}.pt")
                    if os.path.exists(pt_file) and os.path.getsize(pt_file) > 10 * 1024 * 1024:
                        try:
                            self.model = whisper.load_model(fallback_name)
                            loaded_fallback = True
                            break
                        except Exception:
                            pass
                if not loaded_fallback:
                    err_str = str(e_load)
                    if "11001" in err_str or "getaddrinfo" in err_str or "urlopen" in err_str:
                        raise RuntimeError(self.L(
                            "İLK KURULUM İNDİRMESİ GEREKLİ: Ses tanıma yapay zeka paketi bilgisayarınıza sadece İLK KULLANIMDA 1 kez indirilir ve sonrasında ömür boyu %100 internetsiz (offline) çalışır. Şu an bilgisayarınızda internet bağlantısı olmadığı için ilk kurulum paketi indirilemedi. Lütfen sadece bu ilk işlem için Wi-Fi / internet bağlantınızı açıp tekrar deneyin.",
                            "FIRST-TIME MODEL DOWNLOAD REQUIRED: The offline speech AI engine is downloaded only ONCE on first use and runs 100% offline forever after. Please connect to the internet once for this initial setup.",
                            "ERSTMALIGER DOWNLOAD ERFORDERLICH: Das KI-Sprachmodell wird nur beim allerersten Start einmalig heruntergeladen und läuft danach dauerhaft 100% offline. Bitte verbinden Sie sich einmalig mit dem Internet."
                        ))
                    raise
            self.after(0, lambda: self._log(logbox, self.t("log_loaded")))

        src_code = LANG_MAP.get(src_selection, "auto")
        whisper_kwargs = {
            "verbose": True,
            "condition_on_previous_text": not verbatim,
            "no_speech_threshold": 0.95,
            "logprob_threshold": None,
        }
        if src_code != "auto":
            whisper_kwargs["language"] = src_code

        old_stdout = sys.stdout
        sys.stdout = OutputRedirector(self, logbox)
        try:
            result = self.model.transcribe(filepath, **whisper_kwargs)
            if not result.get("segments") and src_code == "auto":
                # Fallback for very short microphone clips where auto-lang detection needs a hint
                whisper_kwargs["language"] = self.ui_lang if self.ui_lang in ("tr", "en", "de") else "tr"
                result = self.model.transcribe(filepath, **whisper_kwargs)
        finally:
            sys.stdout = old_stdout

        detected_lang = result.get("language", src_code if src_code != "auto" else "tr").lower()
        raw_segments = [
            {
                "start": float(s["start"]),
                "end": float(s["end"]),
                "text": str(s["text"]).strip(),
                "speaker": "",
                "note": "",
            }
            for s in result.get("segments", [])
            if str(s.get("text", "")).strip()
        ]
        if not raw_segments and str(result.get("text", "")).strip():
            raw_segments = [{
                "start": 0.0,
                "end": 3.0,
                "text": str(result.get("text", "")).strip(),
                "speaker": "",
                "note": "",
            }]
        return raw_segments, detected_lang

    def _run_engine(self, mode, filepath, src, tgt, logbox, btn, run_id=0):
        try:
            if getattr(self, "cancel_requested", False):
                raise RuntimeError("__USER_CANCELLED__")
            workspace_dir = os.path.join(os.path.expanduser("~"), "Documents", "ILBER_Workspace")
            for sub in ["Uploads", "Transcripts", "Subtitles", "Reports"]:
                os.makedirs(os.path.join(workspace_dir, sub), exist_ok=True)
            base_name = os.path.basename(filepath).rsplit(".", 1)[0]
            ctrls = self.page_controls.get(mode, {})

            # =========================================================================
            # 1. SEGMENT: MIC (Aynı Dilde Deşifre — Word, TXT, SRT, VTT)
            # =========================================================================
            if mode == "tool_ses_kaydet":
                self.after(0, lambda: self._log(logbox, self.L(
                    "[2] MIC: Kayıt yerel olarak zaman damgalı deşifre ediliyor...",
                    "[2] MIC: Transcribing recording locally with timestamps...",
                    "[2] MIC: Aufnahme wird lokal mit Zeitstempeln transkribiert..."
                )))
                segments, detected_lang = self._transcribe_media(filepath, src, logbox, verbatim=False)
                self.after(0, lambda d=detected_lang: self._log(logbox, self.L(
                    f"      ✓ Tespit edilen kaynak dil: {d.upper()}",
                    f"      ✓ Detected source language: {d.upper()}",
                    f"      ✓ Erkannte Ausgangssprache: {d.upper()}"
                )))

                self.after(0, lambda: self._log(logbox, self.L(
                    "[3] Düzenleme ekranı açılıyor (▶ butonuna tıklayarak o saniyeyi dinleyebilirsiniz)...",
                    "[3] Opening editor screen (click ▶ to listen to that exact second)...",
                    "[3] Bearbeitungsfenster wird geöffnet (klicken Sie auf ▶ zum Anhören)..."
                )))
                edited_segments = self.open_interactive_segment_editor_sync(
                    segments,
                    media_path=filepath,
                    title_text=f"MIC — {self.t('pop_edit_title')} ({base_name})",
                    desc_text=self.L(
                        "Her satırı düzenleyebilirsiniz. ▶ butonuna tıkladığınızda kaydın o saniyesi oynatılır.",
                        "You can edit every line. Click the ▶ button to play that exact audio slice.",
                        "Sie können jede Zeile bearbeiten. Klicken Sie auf ▶, um diesen Abschnitt abzuspielen."
                    ),
                )

                # Export 4 formats: DOCX, TXT, SRT, VTT
                self.after(0, lambda: self._log(logbox, self.L(
                    "[4] Çıktı dosyaları (Word, TXT, SRT, VTT) oluşturuluyor...",
                    "[4] Generating output files (Word, TXT, SRT, VTT)...",
                    "[4] Ausgabedateien (Word, TXT, SRT, VTT) werden erstellt..."
                )))
                srt_path = os.path.join(workspace_dir, "Subtitles", f"{base_name}_MIC.srt")
                vtt_path = os.path.join(workspace_dir, "Subtitles", f"{base_name}_MIC.vtt")
                txt_path = os.path.join(workspace_dir, "Transcripts", f"{base_name}_MIC_Metin.txt")
                docx_path = os.path.join(workspace_dir, "Transcripts", f"{base_name}_MIC_Desifre.docx")

                write_srt_file(edited_segments, srt_path)
                write_vtt_file(edited_segments, vtt_path)

                with open(txt_path, "w", encoding="utf-8") as tf:
                    for s in edited_segments:
                        tf.write(f"[{format_ts_hms(s['start'])}] {s['text'].strip()}\n")

                doc = docx.Document()
                doc.add_heading(f"MIC Transcript: {base_name}", 0)
                for s in edited_segments:
                    p = doc.add_paragraph()
                    r_ts = p.add_run(f"[{format_ts_hms(s['start'])}] ")
                    r_ts.bold = True
                    p.add_run(s["text"].strip())
                doc.save(docx_path)

                self.after(0, lambda: self._log(logbox, self.L(
                    "✔ MIC İşlemi Tamamlandı! Üretilen Dosyalar:",
                    "✔ MIC Process Completed! Generated Files:",
                    "✔ MIC-Prozess abgeschlossen! Erstellte Dateien:"
                )))
                self.after(0, lambda: self._log(logbox, f"   📝 Word (.docx): {docx_path}"))
                self.after(0, lambda: self._log(logbox, f"   📄 TXT (.txt): {txt_path}"))
                self.after(0, lambda: self._log(logbox, f"   🎬 SRT (.srt): {srt_path}"))
                self.after(0, lambda: self._log(logbox, f"   🌐 WebVTT (.vtt): {vtt_path}"))
                open_path_crossplatform(docx_path)

            # =========================================================================
            # 3. SEGMENT: GLOBE (Video Çeviri Stüdyosu — Çoklu Hedef Dil, 42 Krk/2 Satır, SRT/VTT/MP4/Çift Dilli Word)
            # =========================================================================
            elif mode == "tool_video_cevir":
                self.after(0, lambda: self._log(logbox, self.L(
                    "[2] GLOBE: Video/Ses yerel olarak deşifre ediliyor...",
                    "[2] GLOBE: Transcribing video/audio locally...",
                    "[2] GLOBE: Video/Audio wird lokal transkribiert..."
                )))
                raw_segments, detected_lang = self._transcribe_media(filepath, src, logbox, verbatim=False)
                src_code = LANG_MAP.get(src, "auto")
                if src_code == "auto":
                    src_code = detected_lang

                # Enforce GLOBE Subtitle Standards (max 42 chars/line, max 2 lines, 1s-7s duration)
                self.after(0, lambda: self._log(logbox, self.L(
                    "[3] Altyazı standartları uygulanıyor (Max 42 karakter/satır, max 2 satır, 1-7 sn süre)...",
                    "[3] Applying subtitle standards (Max 42 chars/line, max 2 lines, 1-7s duration)...",
                    "[3] Untertitelstandards werden angewendet (Max. 42 Zeichen/Zeile, max. 2 Zeilen, 1-7s Dauer)..."
                )))
                std_segments = enforce_subtitle_rules(raw_segments, max_chars_per_line=42, max_lines=2, min_dur=1.0, max_dur=7.0)

                # Pre-translation Editor with Video Frame Preview & Audio Player
                edited_src_segments = self.open_interactive_segment_editor_sync(
                    std_segments,
                    media_path=filepath,
                    title_text=self.L(
                        f"GLOBE Aşama 1: Kaynak Altyazı & Video Önizleme ({src_code.upper()})",
                        f"GLOBE Stage 1: Source Subtitles & Video Preview ({src_code.upper()})",
                        f"GLOBE Stufe 1: Ausgangs-Untertitel & Videovorschau ({src_code.upper()})"
                    ),
                    desc_text=self.L(
                        "Çeviri öncesi kaynak altyazıları video önizlemesi ve ses oynatıcı ile kontrol edin.",
                        "Review source subtitles with video frame preview and audio player before translation.",
                        "Prüfen Sie die Ausgangs-Untertitel vor der Übersetzung mit Videovorschau und Audio-Player."
                    ),
                )

                # Save source SRT and VTT
                src_srt_path = os.path.join(workspace_dir, "Subtitles", f"{base_name}_GLOBE_Kaynak_{src_code}.srt")
                src_vtt_path = os.path.join(workspace_dir, "Subtitles", f"{base_name}_GLOBE_Kaynak_{src_code}.vtt")
                write_srt_file(edited_src_segments, src_srt_path)
                write_vtt_file(edited_src_segments, src_vtt_path)

                # Determine selected target languages
                multi_vars = ctrls.get("multi_lang_vars", {})
                selected_tgt_names = [lname for lname, var in multi_vars.items() if var.get()]
                if not selected_tgt_names:
                    selected_tgt_names = [self.user_settings.get("def_lang", "Turkish")]

                burn_mp4 = bool(ctrls.get("burn_mp4_var").get()) if "burn_mp4_var" in ctrls else True
                src_lines = [s["text"].replace("\n", " ").strip() for s in edited_src_segments]

                for tgt_name in selected_tgt_names:
                    t_code = LANG_MAP.get(tgt_name, "tr")
                    self.after(0, lambda tn=tgt_name, tc=t_code: self._log(logbox, self.L(
                        f"[4] GLOBE Doğrudan Çapraz Çeviri ({src_code.upper()} ➔ {tc.upper()} - {tn})...",
                        f"[4] GLOBE Direct Cross-Translation ({src_code.upper()} ➔ {tc.upper()} - {tn})...",
                        f"[4] GLOBE Direktübersetzung ({src_code.upper()} ➔ {tc.upper()} - {tn})..."
                    )))
                    t_lines = self._translate_lines_m2m100(src_lines, src_code, t_code, logbox, label_prefix=f"GLOBE [{tgt_name}]")

                    raw_tgt_segs = []
                    for i, s in enumerate(edited_src_segments):
                        raw_tgt_segs.append({
                            "start": s["start"],
                            "end": s["end"],
                            "text": t_lines[i] if i < len(t_lines) else s["text"],
                        })
                    # Enforce 42 chars / 2 lines on translated text too
                    std_tgt_segs = enforce_subtitle_rules(raw_tgt_segs, max_chars_per_line=42, max_lines=2, min_dur=1.0, max_dur=7.0)

                    # Post-translation Editor with Video Preview
                    edited_tgt_segs = self.open_interactive_segment_editor_sync(
                        std_tgt_segs,
                        media_path=filepath,
                        title_text=self.L(
                            f"GLOBE Aşama 2: Çevrilmiş Altyazı Kontrolü ({tgt_name})",
                            f"GLOBE Stage 2: Review Translated Subtitles ({tgt_name})",
                            f"GLOBE Stufe 2: Übersetzte Untertitel prüfen ({tgt_name})"
                        ),
                        desc_text=self.L(
                            f"{tgt_name} altyazılarını dışa aktarmadan ve videoya gömmeden önce kontrol edin.",
                            f"Review the {tgt_name} subtitles before exporting and burning into the video.",
                            f"Prüfen Sie die {tgt_name}-Untertitel vor dem Exportieren und Einbrennen ins Video."
                        ),
                    )

                    out_srt = os.path.join(workspace_dir, "Subtitles", f"{base_name}_GLOBE_{t_code}.srt")
                    out_vtt = os.path.join(workspace_dir, "Subtitles", f"{base_name}_GLOBE_{t_code}.vtt")
                    out_bi_docx = os.path.join(workspace_dir, "Transcripts", f"{base_name}_GLOBE_Cift_Dilli_{src_code}_{t_code}.docx")

                    write_srt_file(edited_tgt_segs, out_srt)
                    write_vtt_file(edited_tgt_segs, out_vtt)
                    save_bilingual_word_doc(
                        edited_src_segments,
                        t_lines,
                        src_code.upper(),
                        tgt_name,
                        out_bi_docx,
                        title_text=f"GLOBE Bilingual Video Transcript ({src_code.upper()} - {tgt_name})",
                    )

                    self.after(0, lambda p=out_srt: self._log(logbox, f"   🎬 SRT: {p}"))
                    self.after(0, lambda p=out_vtt: self._log(logbox, f"   🌐 VTT: {p}"))
                    self.after(0, lambda p=out_bi_docx: self._log(logbox, f"   📝 Word (.docx): {p}"))

                    ext = filepath.lower().rsplit(".", 1)[-1]
                    if burn_mp4 and ext in ("mp4", "mov", "mkv", "webm", "avi"):
                        self.after(0, lambda tn=tgt_name: self._log(logbox, self.L(
                            f"[5] {tn} altyazısı videoya gömülüyor (.mp4)...",
                            f"[5] Burning {tn} subtitles into video (.mp4)...",
                            f"[5] {tn}-Untertitel werden in das Video eingebrannt (.mp4)..."
                        )))
                        out_vid = os.path.join(workspace_dir, "Transcripts", f"{base_name}_GLOBE_{t_code}.mp4")
                        srt_esc = out_srt.replace("\\", "/").replace(":", "\\:")

                        f_size = self.user_settings.get("sub_size", "20")
                        pos = "8" if self.user_settings.get("sub_pos", "Bottom") == "Top" else "2"
                        style_name = self.user_settings.get("sub_style", "Classic White")
                        style_str = f"FontSize={f_size},Alignment={pos},PrimaryColour=&H00FFFFFF,OutlineColour=&H00000000,BorderStyle=1,Outline=1,Shadow=1"
                        if style_name == "Yellow Outlined":
                            style_str = f"FontSize={f_size},Alignment={pos},PrimaryColour=&H0000FFFF,OutlineColour=&H00000000,BorderStyle=1,Outline=3,Shadow=2,FontName=Arial,Bold=1"
                        elif style_name == "Black Box":
                            style_str = f"FontSize={f_size},Alignment={pos},PrimaryColour=&H00FFFFFF,BackColour=&H80000000,BorderStyle=3,Outline=4,Shadow=0"

                        vf_arg = f"subtitles='{srt_esc}':force_style='{style_str}'"
                        creationflags = 0x08000000 if os.name == "nt" else 0
                        subprocess.run(
                            ["ffmpeg", "-y", "-i", filepath, "-vf", vf_arg, "-c:v", "libx264", "-preset", "ultrafast", "-c:a", "copy", out_vid],
                            capture_output=True,
                            creationflags=creationflags,
                        )
                        self.after(0, lambda p=out_vid: self._log(logbox, f"   🎥 MP4 Video: {p}"))

                self.after(0, lambda: self._log(logbox, self.L(
                    "✔ GLOBE İşlemi Başarıyla Tamamlandı!",
                    "✔ GLOBE Process Completed Successfully!",
                    "✔ GLOBE-Prozess erfolgreich abgeschlossen!"
                )))

            # =========================================================================
            # 4. SEGMENT: FIELD (Araştırmacılar & Saha Araç Kutusu — Diarization, Anonimleştirme, Saha Notu, NVivo/MAXQDA)
            # =========================================================================
            elif mode == "tool_ai_rapor":
                self.after(0, lambda: self._log(logbox, self.L(
                    "[2] FIELD: Araştırma kaydı birebir (duraksama ve ağız özellikleri korunarak) deşifre ediliyor...",
                    "[2] FIELD: Transcribing research recording verbatim (preserving pauses and dialect)...",
                    "[2] FIELD: Forschungsaufnahme wird wortgetreu transkribiert..."
                )))
                raw_segments, detected_lang = self._transcribe_media(filepath, src, logbox, verbatim=True)

                num_spks = int(ctrls.get("f_spks").get()) if "f_spks" in ctrls else 2
                p_code = ctrls.get("f_pcode").get().strip() if "f_pcode" in ctrls else "P01"
                interviewer_lbl = ctrls.get("f_inter").get().strip() if "f_inter" in ctrls else "Görüşmeci"

                self.after(0, lambda n=num_spks: self._log(logbox, self.L(
                    f"[3] FIELD: Ses frekansı ve akustik özellikler üzerinden {n} konuşmacı ayrıştırılıyor (Audio Diarization)...",
                    f"[3] FIELD: Separating {n} speakers via acoustic voice features (Audio Diarization)...",
                    f"[3] FIELD: {n} Sprecher werden über akustische Stimmmerkmale getrennt (Audio-Diarisierung)..."
                )))
                default_labels = ["Konuşmacı 1", "Konuşmacı 2", "Konuşmacı 3", "Konuşmacı 4", "Konuşmacı 5"][:num_spks]
                diarized_segments = acoustic_diarize_segments(filepath, raw_segments, num_speakers=num_spks, speaker_labels=default_labels)

                self.after(0, lambda: self._log(logbox, self.L(
                    "[4] FIELD: Olası ses tanıma hataları için AI kelime önerileri taranıyor...",
                    "[4] FIELD: Scanning AI word suggestions for potential misheard terms...",
                    "[4] FIELD: KI-Wortvorschläge werden gescannt..."
                )))
                anon_candidates = []
                ai_suggestions = generate_field_ai_suggestions(self._run_llm, diarized_segments)

                self.after(0, lambda: self._log(logbox, self.L(
                    "[5] FIELD Araştırmacı Stüdyosu açılıyor (Konuşmacı adlandırma, saha notları ve kelime önerileri)...",
                    "[5] Opening FIELD Research Studio (Speaker naming, field notes, and word suggestions)...",
                    "[5] FIELD Forschungsstudio wird geöffnet (Sprecherbenennung, Feldnotizen und Wortvorschläge)..."
                )))
                ed_segments, ed_suggestions, ed_anons = self.open_field_studio_sync(
                    diarized_segments,
                    ai_suggestions,
                    anon_candidates,
                    media_path=filepath,
                    default_p_code=p_code,
                    default_interviewer=interviewer_lbl,
                )

                total_dur_sec = ed_segments[-1]["end"] if ed_segments else 0.0
                metadata = {
                    "date": ctrls.get("f_date").get().strip() if "f_date" in ctrls else "-",
                    "location": ctrls.get("f_loc").get().strip() if "f_loc" in ctrls else "-",
                    "participant_code": p_code,
                    "interviewer": interviewer_lbl,
                    "duration": format_ts_hms(total_dur_sec),
                    "context": ctrls.get("f_ctx").get().strip() if "f_ctx" in ctrls else "-",
                }

                translated_lines = None
                tgt_lang_label = None
                if ctrls.get("f_tr_chk") and ctrls["f_tr_chk"].get():
                    tgt_lang_label = ctrls.get("f_tr_tgt").get() if "f_tr_tgt" in ctrls else "English"
                    tgt_code = LANG_MAP.get(tgt_lang_label, "en")
                    self.after(0, lambda: self._log(logbox, self.L(
                        f"[5.5] FIELD: Metin {tgt_lang_label} diline çevriliyor...",
                        f"[5.5] FIELD: Translating transcript into {tgt_lang_label}...",
                        f"[5.5] FIELD: Transkript wird in {tgt_lang_label} übersetzt..."
                    )))
                    src_lines = [s["text"] for s in ed_segments]
                    translated_lines = self._translate_lines_m2m100(src_lines, detected_lang, tgt_code, logbox, label_prefix="FIELD")

                self.after(0, lambda: self._log(logbox, self.L(
                    "[6] FIELD Çıktıları (Birebir Word, Okunabilir Word, NVivo/MAXQDA/ATLAS.ti) üretiliyor...",
                    "[6] Generating FIELD Outputs (Verbatim Word, Readable Word, NVivo/MAXQDA/ATLAS.ti)...",
                    "[6] FIELD-Ausgaben (Wortgetreu Word, Lesbar Word, NVivo/MAXQDA/ATLAS.ti) werden erstellt..."
                )))
                paths = export_field_outputs(
                    workspace_dir=workspace_dir,
                    base_name=base_name,
                    metadata=metadata,
                    segments=ed_segments,
                    suggestions=ed_suggestions,
                    anon_candidates=ed_anons,
                    translated_lines=translated_lines,
                    tgt_lang_label=tgt_lang_label,
                )

                self.after(0, lambda: self._log(logbox, self.L(
                    "✔ FIELD Araştırma Paketi Hazır! Üretilen Dosyalar:",
                    "✔ FIELD Research Package Ready! Generated Files:",
                    "✔ FIELD Forschungspaket bereit! Erstellte Dateien:"
                )))
                self.after(0, lambda: self._log(logbox, f"   📘 1. Verbatim (.docx): {paths['verbatim_docx']}"))
                self.after(0, lambda: self._log(logbox, f"   📗 2. Readable (.docx): {paths['readable_docx']}"))
                self.after(0, lambda: self._log(logbox, f"   📊 3. NVivo / MAXQDA / ATLAS.ti (.txt): {paths['qda_txt']}"))
                open_path_crossplatform(paths["verbatim_docx"])

            # =========================================================================
            # 2. SEGMENT: EDU — FORMAT KORUMALI & TERİM SÖZLÜĞÜ DESTEKLİ BELGE ÇEVİRİCİ
            # =========================================================================
            elif mode == "tool_ai_ozet":
                import importlib
                import ilber_core_utils as _icu
                importlib.reload(_icu)
                ext = filepath.lower().rsplit(".", 1)[-1] if "." in filepath else ""
                is_image_input = ext in ("png", "jpg", "jpeg", "bmp", "webp")
                if ext not in ("docx", "pdf", "xlsx", "txt", "png", "jpg", "jpeg", "bmp", "webp"):
                    self.after(0, lambda e=ext: self._log(logbox, self.L(
                        f"❌ Desteklenmeyen dosya formatı (.{e}). Lütfen Word (.docx), PDF (.pdf), Excel (.xlsx), Görsel (.jpg/.png/.webp) veya Metin (.txt) dosyası seçin.",
                        f"❌ Unsupported format (.{e}). Please select .docx, .pdf, .xlsx, .jpg/.png/.webp, or .txt.",
                        f"❌ Nicht unterstütztes Format (.{e}). Bitte wählen Sie .docx, .pdf, .xlsx, .jpg/.png oder .txt."
                    )))
                    return

                tgt_name = ctrls.get("edu_tgt_var").get() if "edu_tgt_var" in ctrls else self.user_settings.get("def_lang", "Turkish")
                tgt_code = LANG_MAP.get(tgt_name, "tr")
                is_ocr_only = (tgt_code == "ocr_only")
                src_code = LANG_MAP.get(src, "auto")

                # Load optional glossary (from file and/or inline entry)
                g_file = ctrls.get("glossary_file_var").get().strip() if "glossary_file_var" in ctrls else ""
                g_inline = ctrls.get("glossary_inline_entry").get().strip() if "glossary_inline_entry" in ctrls else ""
                use_hq_ai = bool(ctrls.get("edu_hq_var").get()) if "edu_hq_var" in ctrls else True

                glossary_dict = _icu.load_glossary_terms(g_file, g_inline)
                if glossary_dict:
                    sample_pairs = ", ".join([f"{k} ➔ {v}" for k, v in list(glossary_dict.items())[:6]])
                    self.after(0, lambda cnt=len(glossary_dict), sp=sample_pairs: self._log(logbox, self.L(
                        f"📚 [1] Terim Sözlüğü Yüklendi: {cnt} özel terim aktif ({sp}{'...' if cnt > 6 else ''})",
                        f"📚 [1] Terminology Glossary Loaded: {cnt} custom terms active ({sp}{'...' if cnt > 6 else ''})",
                        f"📚 [1] Fachglossar geladen: {cnt} Fachbegriffe aktiv ({sp}{'...' if cnt > 6 else ''})"
                    )))
                else:
                    self.after(0, lambda: self._log(logbox, self.L(
                        "ℹ️ [1] Terim sözlüğü yüklenmedi — Sözlüksüz doğrudan yüksek kaliteli işlem uygulanıyor.",
                        "ℹ️ [1] No glossary loaded — Running direct high-quality processing without glossary.",
                        "ℹ️ [1] Kein Glossar geladen — Direkte Verarbeitung ohne Glossar."
                    )))

                if is_image_input:
                    self.after(0, lambda e=ext.upper(), tn=tgt_name: self._log(logbox, self.L(
                        f"🔍 [2] {e} görseli / ekran görüntüsü üzerindeki yazılar OCR (Optik Karakter Tanıma) ile düzenlenebilir metne dönüştürülüyor ({tn})...",
                        f"🔍 [2] Extracting editable text from {e} image / screenshot via OCR ({tn})...",
                        f"🔍 [2] Text aus {e}-Bild wird per OCR extrahiert ({tn})..."
                    )))
                else:
                    self.after(0, lambda e=ext.upper(), tn=tgt_name: self._log(logbox, self.L(
                        f"📄 [2] {e} belgesi okunuyor ve orijinal sayfa/tablo/stil yapısı korunarak {tn} diline çevriliyor...",
                        f"📄 [2] Reading {e} document and translating into {tn} while preserving original layout, tables, and styles...",
                        f"📄 [2] {e}-Dokument wird gelesen und unter Beibehaltung des Originalformats in {tn} übersetzt..."
                    )))

                def _detect_source_lang_from_texts(sample_texts):
                    effective_tgt = "tr" if is_ocr_only else tgt_code
                    return _icu.detect_document_source_lang(sample_texts, src_code=src_code, tgt_code=effective_tgt)

                def _batch_translate_doc(texts_list):
                    if is_ocr_only:
                        draft_translations = list(texts_list)
                        self.after(0, lambda n_seg=len(texts_list): self._log(logbox, self.L(
                            f"      ✓ OCR Metin Tanıma Tamamlandı: Toplam {n_seg} satır/bölüm düzenlenebilir metne dönüştürüldü",
                            f"      ✓ OCR Text Recognition Completed: {n_seg} lines converted to editable text",
                            f"      ✓ OCR-Texterkennung abgeschlossen: {n_seg} Zeilen in bearbeitbaren Text umgewandelt"
                        )))
                    else:
                        actual_src = _detect_source_lang_from_texts(texts_list)
                        self.after(0, lambda s_c=actual_src, t_c=tgt_code, n_seg=len(texts_list): self._log(logbox, self.L(
                            f"      ✓ Algılanan Kaynak Dil: {s_c.upper()} ➔ Hedef Dil: {t_c.upper()} (Toplam {n_seg} metin bloğu)",
                            f"      ✓ Detected Source Language: {s_c.upper()} ➔ Target: {t_c.upper()} ({n_seg} text blocks)",
                            f"      ✓ Erkannte Ausgangssprache: {s_c.upper()} ➔ Ziel: {t_c.upper()} ({n_seg} Textblöcke)"
                        )))

                        def _m2m_runner(arr, sc, tc):
                            return self._translate_lines_m2m100(arr, sc, tc, logbox, label_prefix="EDU")

                        def _polish_log(pct):
                            self.after(0, lambda p=pct: self._log(logbox, self.L(
                                f"      -> ✨ AI Cümle Akıcılığı & Terim Sözlüğü Redaksiyonu: %{p}",
                                f"      -> ✨ AI Fluency & Terminology Polishing: {p}%",
                                f"      -> ✨ KI-Stil- & Glossar-Feinschliff: {p}%"
                            )))

                        draft_translations = _icu.translate_texts_high_quality(
                            texts=texts_list,
                            src_code=actual_src,
                            tgt_code=tgt_code,
                            glossary_dict=glossary_dict,
                            m2m_translate_fn=_m2m_runner,
                            llm_fn=self._run_llm if use_hq_ai else None,
                            use_llm_polish=use_hq_ai,
                            log_fn=_polish_log,
                        )

                    open_editor = bool(ctrls.get("edu_edit_var").get()) if "edu_edit_var" in ctrls else True
                    if open_editor and texts_list:
                        self.after(0, lambda: self._log(logbox, self.L(
                            "✏️ [3] Nihai belge oluşturulmadan önce Kontrol & Düzenleme Ekranı açılıyor (Lütfen açılan pencerede kontrol edip '✓ Onayla ve Nihai Belgeyi Oluştur' butonuna basın)...",
                            "✏️ [3] Opening Pre-Export Review & Editor Screen (Please review and click '✓ Approve & Generate Final Document')...",
                            "✏️ [3] Bearbeitungsfenster vor der Dokumenterstellung wird geöffnet..."
                        )))
                        draft_translations = self.open_doc_translation_editor_sync(
                            source_texts=texts_list,
                            translated_texts=draft_translations,
                            doc_name=f"{base_name}.{ext}",
                        )
                    return draft_translations

                if is_image_input:
                    tag_suf = "OCR_Metin" if is_ocr_only else f"OCR_Ceviri_{tgt_code.upper()}"
                    out_docx_path = os.path.join(workspace_dir, "Transcripts", f"{base_name}_EDU_{tag_suf}.docx")
                    out_txt_path = os.path.join(workspace_dir, "Transcripts", f"{base_name}_EDU_{tag_suf}.txt")
                    out_img_path = os.path.join(workspace_dir, "Transcripts", f"{base_name}_EDU_{tag_suf}.{ext}")
                    count_seg, preview_pairs = _icu.process_image_ocr_and_translate(
                        filepath,
                        out_docx_path=out_docx_path,
                        out_txt_path=out_txt_path,
                        out_img_path=out_img_path,
                        translate_batch_fn=_batch_translate_doc,
                        ocr_only=is_ocr_only,
                    )
                    out_path = out_docx_path
                else:
                    tag_suf = "OCR_Metin" if is_ocr_only else f"Ceviri_{tgt_code.upper()}"
                    out_filename = f"{base_name}_EDU_{tag_suf}.{ext}"
                    out_path = os.path.join(workspace_dir, "Transcripts", out_filename)

                    if ext == "docx":
                        count_seg, preview_pairs = _icu.translate_docx_preserve_format(filepath, out_path, _batch_translate_doc)
                    elif ext == "pdf":
                        count_seg, preview_pairs = _icu.translate_pdf_preserve_format(filepath, out_path, _batch_translate_doc)
                    elif ext == "xlsx":
                        count_seg, preview_pairs = _icu.translate_xlsx_preserve_format(filepath, out_path, _batch_translate_doc)
                    else:
                        count_seg, preview_pairs = _icu.translate_txt_preserve_format(filepath, out_path, _batch_translate_doc)

                if preview_pairs:
                    preview_lines = []
                    for idx_p, (s_p, t_p) in enumerate(preview_pairs[:10], 1):
                        s_short = s_p[:140] + ("..." if len(s_p) > 140 else "")
                        t_short = t_p[:140] + ("..." if len(t_p) > 140 else "")
                        if s_short == t_short:
                            preview_lines.append(f"  [{idx_p}] {t_short}")
                        else:
                            preview_lines.append(f"  [{idx_p}] {s_short}\n       ➔ {t_short}")
                    joined_prev = "\n".join(preview_lines)
                    self.after(0, lambda jp=joined_prev: self._log(logbox, self.L(
                        f"\n📄 SONUÇ ÖNİZLEMESİ (İlk Bölümler):\n--------------------------------------------------\n{jp}\n--------------------------------------------------",
                        f"\n📄 OUTPUT PREVIEW (First Segments):\n--------------------------------------------------\n{jp}\n--------------------------------------------------",
                        f"\n📄 ERGEBNISVORSCHAU:\n--------------------------------------------------\n{jp}\n--------------------------------------------------"
                    )))

                if is_image_input:
                    self.after(0, lambda c=count_seg: self._log(logbox, self.L(
                        f"\n✔ EDU Görsel OCR & Düzenlenebilir Belge Dönüşümü Tamamlandı! ({c} satır işlendi)",
                        f"\n✔ EDU Image OCR & Editable Document Conversion Completed! ({c} lines processed)",
                        f"\n✔ EDU Bild-OCR & Dokumentkonvertierung abgeschlossen! ({c} Zeilen)"
                    )))
                    self.after(0, lambda p=out_docx_path: self._log(logbox, f"   📘 Düzenlenebilir Word (.docx): {p}"))
                    self.after(0, lambda p=out_txt_path: self._log(logbox, f"   📄 Düz Metin (.txt): {p}"))
                    if not is_ocr_only and os.path.exists(out_img_path):
                        self.after(0, lambda p=out_img_path, e=ext.upper(): self._log(logbox, f"   🖼️ Çevrilmiş Görsel (.{ext}): {p}"))
                else:
                    self.after(0, lambda c=count_seg, e=ext.upper(): self._log(logbox, self.L(
                        f"\n✔ EDU Format Korumalı Belge Çevirisi Tamamlandı! ({c} bölüm çevrildi · Giriş: .{ext} ➔ Çıkış: .{ext})",
                        f"\n✔ EDU Format-Preserving Document Translation Completed! ({c} segments translated · Input: .{ext} ➔ Output: .{ext})",
                        f"\n✔ EDU Formatgetreue Dokumentübersetzung abgeschlossen! ({c} Abschnitte · .{ext} ➔ .{ext})"
                    )))
                    self.after(0, lambda p=out_path, e=ext.upper(): self._log(logbox, f"   📁 {e} Çıktı Dosyası: {p}"))
                open_path_crossplatform(out_path)

            elif mode == "tool_ai_chat":
                ext = filepath.lower().rsplit(".", 1)[-1]
                if ext in ("png", "jpg", "jpeg", "bmp", "webp"):
                    self.after(0, lambda: self._log(logbox, self.L(
                        "[2] Görsel / ekran görüntüsü üzerindeki metin ve sorular okunuyor...",
                        "[2] Reading text and questions from image / screenshot...",
                        "[2] Text und Fragen aus Bild / Screenshot werden gelesen..."
                    )))
                    transcript_text = extract_text_from_image_offline(filepath)
                elif ext == "pdf":
                    self.after(0, lambda: self._log(logbox, self.L(
                        "[2] PDF belgesi okunuyor...",
                        "[2] Reading PDF document...",
                        "[2] PDF-Dokument wird gelesen..."
                    )))
                    import pdfplumber
                    pdf_text = []
                    with pdfplumber.open(filepath) as pdf:
                        for idx, page in enumerate(pdf.pages):
                            pt = page.extract_text()
                            if pt and len(pt.strip()) >= 15:
                                pdf_text.append(pt)
                            else:
                                try:
                                    tmp_img_path = os.path.join(TEMP_DIR, f"pdf_page_ocr_{idx}.png")
                                    page.to_image(resolution=200).save(tmp_img_path, format="PNG")
                                    ocr_pt = extract_text_from_image_offline(tmp_img_path)
                                    if ocr_pt:
                                        pdf_text.append(ocr_pt)
                                except Exception:
                                    if pt:
                                        pdf_text.append(pt)
                    transcript_text = "\n".join(pdf_text)
                else:
                    self.after(0, lambda: self._log(logbox, self.L(
                        "[2] Medya dosyası deşifre edilerek hafızaya alınıyor...",
                        "[2] Transcribing media file into memory...",
                        "[2] Mediendatei wird transkribiert und ins Gedächtnis geladen..."
                    )))
                    segs, _ = self._transcribe_media(filepath, src, logbox, verbatim=False)
                    transcript_text = "\n".join([f"[{format_ts_hms(s['start'])}] {s['text']}" for s in segs])

                self.video_context = transcript_text
                self.chat_history = []
                preview_txt = (transcript_text or "").strip()
                if preview_txt:
                    short_preview = preview_txt[:1200] + ("..." if len(preview_txt) > 1200 else "")
                    self.after(0, lambda p=short_preview: self._log(logbox, self.L(
                        f"\n📄 OKUNAN İÇERİK ÖNİZLEMESİ:\n--------------------------------------------------\n{p}\n--------------------------------------------------",
                        f"\n📄 EXTRACTED CONTENT PREVIEW:\n--------------------------------------------------\n{p}\n--------------------------------------------------",
                        f"\n📄 VORSCHAU DES GELESENEN INHALTS:\n--------------------------------------------------\n{p}\n--------------------------------------------------"
                    )))
                    self.after(0, lambda: self._log(logbox, "\n=================================================="))
                    self.after(0, lambda: self._log(logbox, self.L(
                        " ✔ İÇERİK HAFIZAYA ALINDI!",
                        " ✔ CONTENT MEMORIZED!",
                        " ✔ INHALT INS GEDÄCHTNIS GELADEN!"
                    )))
                    self.after(0, lambda: self._log(logbox, self.L(
                        " Artık aşağıdaki sohbet kutusundan soru sorabilir veya '🧠 Soruyu Adım Adım Çöz & Açıkla' butonuna basabilirsiniz.",
                        " You can now ask any question below or click '🧠 Solve & Explain Step-by-Step'.",
                        " Sie können nun unten Fragen stellen oder auf '🧠 Frage Schritt für Schritt lösen & erklären' klicken."
                    )))
                    self.after(0, lambda: self._log(logbox, "==================================================\n"))
                else:
                    self.after(0, lambda: self._log(logbox, self.L(
                        "\n⚠️ UYARI: Seçilen görsel üzerinde okunabilir bir metin bulunamadı. Lütfen soruyu tam çerçeve içine alarak tekrar kırpın veya '📋 Panodan Yapıştır' butonunu kullanın.\n",
                        "\n⚠️ WARNING: No readable text was found in the selected image. Please snip the full question area again or use '📋 Paste Image'.\n",
                        "\n⚠️ WARNUNG: Im ausgewählten Bild wurde kein lesbarer Text gefunden.\n"
                    )))

        except Exception as e:
            if "__USER_CANCELLED__" not in str(e) and not getattr(self, "cancel_requested", False):
                err = f"❌ Error: {str(e)}"
                self.after(0, lambda msg=err: self._log(logbox, msg))
        finally:
            if run_id == getattr(self, "_current_run_id", 0):
                self.is_processing = False
                self._stop_busy_indicator(mode)
                ctrls_fin = self.page_controls.get(mode, {})
                st_btn = ctrls_fin.get("stop_btn")
                if st_btn:
                    self.after(0, lambda b=st_btn: b.winfo_exists() and b.configure(state="disabled", fg_color="#94A3B8"))
                if mode == "tool_ai_chat":
                    self.after(0, lambda: btn.configure(state="normal", text=self.t("chat_memorize")))
                else:
                    self.after(0, lambda: btn.configure(state="normal", text=self.t("btn_start_proc")))


if __name__ == "__main__":
    if "--selftest" in sys.argv:
        import numpy as np
        wh = _lazy_import_whisper()
        import torch
        from transformers import M2M100ForConditionalGeneration, M2M100Tokenizer
        t_arr = torch.from_numpy(np.zeros(16, dtype=np.float32))
        assert t_arr.shape[0] == 16, "torch.from_numpy bridge failed!"
        assert hasattr(M2M100ForConditionalGeneration, "from_pretrained"), "M2M100 class invalid!"
        ff_res = subprocess.run(["ffmpeg", "-version"], capture_output=True, text=True)
        assert ff_res.returncode == 0, f"Bundled ffmpeg failed: {ff_res.stderr}"
        print(f"✅ SELFTEST PASSED | NumPy={np.__version__} | PyTorch={torch.__version__} | FFmpeg=OK")
        sys.exit(0)

    app = ILBERTechV2()
    try:
        app.deiconify()
        app.lift()
        app.attributes("-topmost", True)
        app.after(600, lambda: app.attributes("-topmost", False))
        app.focus_force()
    except Exception:
        pass
    app.mainloop()

