import os
import re
import threading
import tkinter as tk
import customtkinter as ctk
from tkinter import messagebox
from ilber_core_utils import (
    AudioVideoPreviewer,
    format_ts_hms,
    format_ts_srt,
    clean_markdown_text,
)


DIALOG_I18N = {
    "tr": {
        "play_btn": "▶ Dinle",
        "stop_btn": "⏹ Ses Durdur",
        "play_hint": "💡 İpucu: Satırın solundaki mavi ▶ zaman damgası butonuna basarak kaydın o saniyesini dinleyebilir ve video karesini görebilirsiniz.",
        "video_preview": "Video Önizleme\n(▶ Butonuna Basın)",
        "no_video_preview": "Ses Dosyası",
        "col_time": "Zaman (Tıkla & Dinle)",
        "col_speaker": "Konuşmacı",
        "col_text": "Deşifre Metni (Düzenlenebilir)",
        "col_note": "Saha Notu (Word Yorumu)",
        "save_continue": "✓ Kaydet ve Devam Et",
        "tab_rows": "📋 Satır Satır Düzenle & Dinle",
        "tab_raw": "📝 Tüm Metin Görünümü",
        "field_title": "FIELD Araştırmacı Stüdyosu — Konuşmacı, Anonimleştirme & Saha Notları",
        "tab_field_trans": "1. Deşifre, Konuşmacılar & Saha Notları",
        "tab_field_sugg": "2. AI Kelime Önerileri (Kabul / Ret)",
        "tab_field_anon": "3. Anonimleştirme ([KİŞİ-1], [YER-1])",
        "bulk_rename": "Hızlı Konuşmacı Adlandırma:",
        "apply_rename": "Tümüne Uygula",
        "sugg_desc": "Orijinal birebir deşifre ASLA otomatik değiştirilmez. Kabul ettiğiniz öneriler yalnızca 'Okunabilir Metin' çıktısına uygulanır.",
        "no_sugg": "AI tarafından bariz bir yanlış duyulma tespit edilmedi.",
        "btn_accept": "✓ Kabul Et",
        "btn_reject": "✗ Reddet",
        "status_accepted": "KABUL EDİLDİ",
        "status_rejected": "REDDEDİLDİ",
        "anon_desc": "Etik kurul standartları için tespit edilen kişi, yer ve kurum adlarını onaylayın. Eşleştirme anahtarı ayrı bir gizli dosyaya kaydedilir.",
        "anon_add_orig": "Özel İsim Ekle (Örn: Ahmet):",
        "anon_add_btn": "+ Listeye Ekle",
        "page_prev": "◀ Önceki Sayfa",
        "page_next": "Sonraki Sayfa ▶",
        "page_lbl": "Sayfa",
    },
    "en": {
        "play_btn": "▶ Play",
        "stop_btn": "⏹ Stop Audio",
        "play_hint": "💡 Tip: Click the blue ▶ timestamp button on the left of any row to play that exact audio slice and preview the video frame.",
        "video_preview": "Video Preview\n(Click ▶ Button)",
        "no_video_preview": "Audio File",
        "col_time": "Timestamp (Click & Play)",
        "col_speaker": "Speaker",
        "col_text": "Transcript Text (Editable)",
        "col_note": "Field Note (Word Comment)",
        "save_continue": "✓ Save and Continue",
        "tab_rows": "📋 Line-by-Line Edit & Play",
        "tab_raw": "📝 Full Text View",
        "field_title": "FIELD Research Studio — Speakers, Anonymization & Field Notes",
        "tab_field_trans": "1. Transcript, Speakers & Field Notes",
        "tab_field_sugg": "2. AI Word Suggestions (Accept / Reject)",
        "tab_field_anon": "3. Anonymization ([PERSON-1], [PLACE-1])",
        "bulk_rename": "Quick Speaker Renaming:",
        "apply_rename": "Apply to All",
        "sugg_desc": "Original verbatim transcript is NEVER auto-modified. Accepted suggestions are applied only to the 'Readable Text' export.",
        "no_sugg": "No obvious misheard words detected by AI.",
        "btn_accept": "✓ Accept",
        "btn_reject": "✗ Reject",
        "status_accepted": "ACCEPTED",
        "status_rejected": "REJECTED",
        "anon_desc": "Approve detected person, location, and institution names for ethics compliance. The mapping key is saved in a separate file.",
        "anon_add_orig": "Add Custom Name:",
        "anon_add_btn": "+ Add to List",
        "page_prev": "◀ Prev Page",
        "page_next": "Next Page ▶",
        "page_lbl": "Page",
    },
    "de": {
        "play_btn": "▶ Abspielen",
        "stop_btn": "⏹ Audio Stopp",
        "play_hint": "💡 Tipp: Klicken Sie auf die blaue ▶ Zeitstempel-Schaltfläche links in einer Zeile, um genau diese Sekunde anzuhören.",
        "video_preview": "Video-Vorschau\n(▶ Klicken)",
        "no_video_preview": "Audiodatei",
        "col_time": "Zeitstempel (Klicken & Hören)",
        "col_speaker": "Sprecher",
        "col_text": "Transkripttext (Bearbeitbar)",
        "col_note": "Feldnotiz (Word-Kommentar)",
        "save_continue": "✓ Speichern und Weiter",
        "tab_rows": "📋 Zeilenweise Bearbeiten & Hören",
        "tab_raw": "📝 Volltextansicht",
        "field_title": "FIELD Forschungsstudio — Sprecher, Anonymisierung & Feldnotizen",
        "tab_field_trans": "1. Transkript, Sprecher & Feldnotizen",
        "tab_field_sugg": "2. KI-Wortvorschläge (Annehmen / Ablehnen)",
        "tab_field_anon": "3. Anonymisierung ([PERSON-1], [ORT-1])",
        "bulk_rename": "Sprecher umbenennen:",
        "apply_rename": "Auf alle anwenden",
        "sugg_desc": "Das Originaltranskript wird NIEMALS automatisch geändert. Angenommene Vorschläge gelten nur für die 'Lesbare Version'.",
        "no_sugg": "Keine offensichtlichen Hörfehler von der KI erkannt.",
        "btn_accept": "✓ Annehmen",
        "btn_reject": "✗ Ablehnen",
        "status_accepted": "ANGENOMMEN",
        "status_rejected": "ABGELEHNT",
        "anon_desc": "Bestätigen Sie erkannte Personen-, Orts- und Institutionsnamen. Die Zuordnungsliste wird separat gespeichert.",
        "anon_add_orig": "Eigenen Namen hinzufügen:",
        "anon_add_btn": "+ Hinzufügen",
        "page_prev": "◀ Vorherige",
        "page_next": "Nächste ▶",
        "page_lbl": "Seite",
    },
}


class InteractiveSegmentEditor(ctk.CTkToplevel):
    """
    Fast, zero-lag timestamped segment editor for MIC, EDU, and GLOBE.
    - Uses lightweight native tk.Text(wrap="word") rows so typing & scrolling have 0ms lag.
    - Clicking inside the text box only edits text (never blocks or triggers FFmpeg).
    - Clicking the blue ▶ button plays that exact audio slice and updates the video preview asynchronously.
    """
    PAGE_SIZE = 20

    def __init__(self, parent, segments, media_path=None, ui_lang="tr", title_text="", desc_text="", btn_text="", on_complete=None):
        super().__init__(parent)
        self.segments = [dict(s) for s in segments]
        self.media_path = media_path
        self.ui_lang = ui_lang if ui_lang in DIALOG_I18N else "tr"
        self.tr_dict = DIALOG_I18N[self.ui_lang]
        self.on_complete = on_complete
        self.previewer = AudioVideoPreviewer()
        self.current_page = 0
        self.row_widgets = {}  # seg_idx -> tk.Text

        self.title(title_text or self.tr_dict["tab_rows"])
        self.geometry("1040x720")
        self.configure(fg_color="#F8FAFC")
        self.attributes("-topmost", True)
        self.protocol("WM_DELETE_WINDOW", self._on_save_and_close)

        # Header & Video/Audio Preview Bar
        top_bar = ctk.CTkFrame(self, fg_color="#FFFFFF", corner_radius=12, border_width=1, border_color="#E2E8F0")
        top_bar.pack(fill="x", padx=20, pady=(15, 8))

        info_col = ctk.CTkFrame(top_bar, fg_color="transparent")
        info_col.pack(side="left", fill="both", expand=True, padx=15, pady=10)

        ctk.CTkLabel(info_col, text=title_text, font=ctk.CTkFont(family="Segoe UI", size=17, weight="bold"), text_color="#0F172A").pack(anchor="w")
        if desc_text:
            ctk.CTkLabel(info_col, text=desc_text, font=ctk.CTkFont(family="Segoe UI", size=13), text_color="#475569", wraplength=620, justify="left").pack(anchor="w", pady=(2, 4))

        ctk.CTkLabel(info_col, text=self.tr_dict["play_hint"], font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"), text_color="#2563EB", wraplength=620, justify="left").pack(anchor="w", pady=(2, 6))

        ctrl_row = ctk.CTkFrame(info_col, fg_color="transparent")
        ctrl_row.pack(anchor="w", pady=(2, 0))
        ctk.CTkButton(ctrl_row, text=self.tr_dict["stop_btn"], width=120, height=30, fg_color="#EF4444", hover_color="#DC2626", font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"), command=self.previewer.stop).pack(side="left")

        # Video Thumbnail Preview Box on Right
        self.preview_box = ctk.CTkFrame(top_bar, fg_color="#0F172A", width=220, height=124, corner_radius=8)
        self.preview_box.pack(side="right", padx=15, pady=10)
        self.preview_box.pack_propagate(False)
        self.preview_lbl = ctk.CTkLabel(self.preview_box, text=self.tr_dict["video_preview"], text_color="#94A3B8", font=ctk.CTkFont(family="Segoe UI", size=11))
        self.preview_lbl.place(relx=0.5, rely=0.5, anchor="center")

        # Pagination bar
        self.page_bar = ctk.CTkFrame(self, fg_color="transparent")
        self.page_bar.pack(fill="x", padx=20, pady=(0, 5))

        self.btn_prev = ctk.CTkButton(self.page_bar, text=self.tr_dict["page_prev"], width=120, height=30, fg_color="#E2E8F0", text_color="#1E293B", hover_color="#CBD5E1", command=self._prev_page)
        self.btn_prev.pack(side="left")
        self.lbl_page = ctk.CTkLabel(self.page_bar, text="", font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"), text_color="#334155")
        self.lbl_page.pack(side="left", padx=15)
        self.btn_next = ctk.CTkButton(self.page_bar, text=self.tr_dict["page_next"], width=120, height=30, fg_color="#E2E8F0", text_color="#1E293B", hover_color="#CBD5E1", command=self._next_page)
        self.btn_next.pack(side="left")

        # Main Scrollable Segment List
        self.scroll = ctk.CTkScrollableFrame(self, fg_color="#FFFFFF", corner_radius=12, border_width=1, border_color="#E2E8F0")
        self.scroll.pack(fill="both", expand=True, padx=20, pady=5)

        # Bottom Save Button
        save_label = btn_text if btn_text else self.tr_dict["save_continue"]
        ctk.CTkButton(self, text=save_label, fg_color="#10B981", hover_color="#059669", height=44, font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"), command=self._on_save_and_close).pack(fill="x", padx=20, pady=12)

        self._render_page()
        if self.segments:
            self._update_video_preview_async(self.segments[0]["start"])

    def _sync_current_page_edits(self):
        for seg_idx, tb in self.row_widgets.items():
            if 0 <= seg_idx < len(self.segments) and tb.winfo_exists():
                self.segments[seg_idx]["text"] = tb.get("1.0", "end-1c").strip()

    def _total_pages(self):
        return max(1, (len(self.segments) + self.PAGE_SIZE - 1) // self.PAGE_SIZE)

    def _prev_page(self):
        if self.current_page > 0:
            self._sync_current_page_edits()
            self.current_page -= 1
            self._render_page()

    def _next_page(self):
        if self.current_page + 1 < self._total_pages():
            self._sync_current_page_edits()
            self.current_page += 1
            self._render_page()

    def _render_page(self):
        for w in self.scroll.winfo_children():
            w.destroy()
        self.row_widgets.clear()

        total_p = self._total_pages()
        self.lbl_page.configure(text=f"{self.tr_dict['page_lbl']} {self.current_page + 1} / {total_p}  ({len(self.segments)} segment)")
        self.btn_prev.configure(state="normal" if self.current_page > 0 else "disabled")
        self.btn_next.configure(state="normal" if self.current_page + 1 < total_p else "disabled")

        start_i = self.current_page * self.PAGE_SIZE
        end_i = min(len(self.segments), start_i + self.PAGE_SIZE)

        for idx in range(start_i, end_i):
            seg = self.segments[idx]
            row = ctk.CTkFrame(self.scroll, fg_color="#F8FAFC", corner_radius=8, border_width=1, border_color="#E2E8F0")
            row.pack(fill="x", padx=6, pady=3)

            ts_label = f"▶ {format_ts_hms(seg['start'])} - {format_ts_hms(seg['end'])}"
            btn_play = ctk.CTkButton(
                row,
                text=ts_label,
                width=165,
                height=36,
                fg_color="#2563EB",
                hover_color="#1D4ED8",
                font=ctk.CTkFont(family="Consolas", size=12, weight="bold"),
                command=lambda s=seg["start"], e=seg["end"]: self._on_row_play(s, e),
            )
            btn_play.pack(side="left", padx=8, pady=6)

            # Fast native tk.Text with wrap="word" (zero canvas lag when typing or scrolling)
            tb = tk.Text(
                row,
                height=2,
                wrap="word",
                bg="#FFFFFF",
                fg="#0F172A",
                insertbackground="#0F172A",
                relief="solid",
                bd=1,
                padx=8,
                pady=4,
                font=("Segoe UI", 11),
            )
            tb.pack(side="left", fill="x", expand=True, padx=(0, 8), pady=6)
            tb.insert("1.0", seg.get("text", ""))
            self.row_widgets[idx] = tb

    def _on_row_play(self, start_sec, end_sec):
        self.previewer.play_segment(self.media_path, start_sec, end_sec)
        self._update_video_preview_async(start_sec)

    def _update_video_preview_async(self, time_sec):
        def _bg():
            pil_img = self.previewer.extract_frame_pil(self.media_path, time_sec, size=(216, 120))
            def _apply():
                if not self.winfo_exists() or not self.preview_lbl.winfo_exists():
                    return
                if pil_img is not None:
                    ctk_img = ctk.CTkImage(light_image=pil_img, size=pil_img.size)
                    self.preview_lbl.configure(image=ctk_img, text="")
                    self.preview_lbl._image_ref = ctk_img
                else:
                    self.preview_lbl.configure(image=None, text=f"{format_ts_hms(time_sec)}\n{self.tr_dict['no_video_preview']}")
            self.after(0, _apply)
        threading.Thread(target=_bg, daemon=True).start()

    def _on_save_and_close(self):
        self.previewer.stop()
        self._sync_current_page_edits()
        edited = self.segments
        self.destroy()
        if self.on_complete:
            self.on_complete(edited)


class FieldResearchStudio(ctk.CTkToplevel):
    """
    Fast qualitative research workstation for the FIELD segment:
    - Uses lightweight native tk.Text(wrap="word") rows for zero-lag editing.
    - Audio plays only when clicking the ▶ timestamp button.
    """
    PAGE_SIZE = 20

    def __init__(self, parent, segments, suggestions, anon_candidates, media_path=None, ui_lang="tr", default_p_code="P01", default_interviewer="Görüşmeci", on_complete=None):
        super().__init__(parent)
        self.segments = [dict(s) for s in segments]
        self.suggestions = [dict(x) for x in suggestions]
        self.anon_candidates = [dict(a) for a in anon_candidates]
        self.media_path = media_path
        self.ui_lang = ui_lang if ui_lang in DIALOG_I18N else "tr"
        self.tr_dict = DIALOG_I18N[self.ui_lang]
        self.on_complete = on_complete
        self.previewer = AudioVideoPreviewer()
        self.current_page = 0
        self.row_widgets = {}  # idx -> (spk_var, text_widget, note_entry)

        self.speaker_list = [default_interviewer, default_p_code, "P02", "P03", "P04", "Konuşmacı 1", "Konuşmacı 2", "Konuşmacı 3", "Konuşmacı 4"]
        seen_s = set()
        self.speaker_list = [x for x in self.speaker_list if not (x in seen_s or seen_s.add(x))]

        self.title(self.tr_dict["field_title"])
        self.geometry("1160x760")
        self.configure(fg_color="#F8FAFC")
        self.attributes("-topmost", True)
        self.protocol("WM_DELETE_WINDOW", self._on_save_and_close)

        self.tabview = ctk.CTkTabview(self, fg_color="#FFFFFF", segmented_button_selected_color="#EF4444", segmented_button_selected_hover_color="#DC2626")
        self.tabview.pack(fill="both", expand=True, padx=18, pady=(10, 8))

        self.tab_trans = self.tabview.add(self.tr_dict["tab_field_trans"])
        self.tab_sugg = self.tabview.add(self.tr_dict["tab_field_sugg"])
        self.anon_vars = []

        self._build_tab_transcript(default_interviewer, default_p_code)
        self._build_tab_suggestions()

        ctk.CTkButton(
            self,
            text=self.tr_dict["save_continue"],
            fg_color="#10B981",
            hover_color="#059669",
            height=45,
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            command=self._on_save_and_close,
        ).pack(fill="x", padx=18, pady=(0, 12))

    def _build_tab_transcript(self, default_interviewer, default_p_code):
        top_ctrl = ctk.CTkFrame(self.tab_trans, fg_color="#F1F5F9", corner_radius=10)
        top_ctrl.pack(fill="x", padx=10, pady=6)

        ctk.CTkLabel(top_ctrl, text=self.tr_dict["bulk_rename"], font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"), text_color="#1E293B").pack(side="left", padx=(12, 6), pady=8)

        self.map_spk1_var = ctk.StringVar(value=default_interviewer)
        self.map_spk2_var = ctk.StringVar(value=default_p_code)

        ctk.CTkLabel(top_ctrl, text="Konuşmacı 1 →", font=ctk.CTkFont(size=12), text_color="#475569").pack(side="left", padx=(4, 2))
        ctk.CTkEntry(top_ctrl, textvariable=self.map_spk1_var, width=110, height=28, fg_color="#FFFFFF").pack(side="left", padx=(0, 8))

        ctk.CTkLabel(top_ctrl, text="Konuşmacı 2 →", font=ctk.CTkFont(size=12), text_color="#475569").pack(side="left", padx=(4, 2))
        ctk.CTkEntry(top_ctrl, textvariable=self.map_spk2_var, width=90, height=28, fg_color="#FFFFFF").pack(side="left", padx=(0, 8))

        ctk.CTkButton(top_ctrl, text=self.tr_dict["apply_rename"], width=120, height=28, fg_color="#2563EB", hover_color="#1D4ED8", font=ctk.CTkFont(size=12, weight="bold"), command=self._apply_bulk_speaker_rename).pack(side="left", padx=6)

        ctk.CTkButton(top_ctrl, text=self.tr_dict["stop_btn"], width=105, height=28, fg_color="#EF4444", hover_color="#DC2626", font=ctk.CTkFont(size=12, weight="bold"), command=self.previewer.stop).pack(side="right", padx=10)

        page_bar = ctk.CTkFrame(self.tab_trans, fg_color="transparent")
        page_bar.pack(fill="x", padx=10, pady=(2, 4))
        self.btn_prev = ctk.CTkButton(page_bar, text=self.tr_dict["page_prev"], width=115, height=28, fg_color="#E2E8F0", text_color="#1E293B", command=self._prev_page)
        self.btn_prev.pack(side="left")
        self.lbl_page = ctk.CTkLabel(page_bar, text="", font=ctk.CTkFont(size=12, weight="bold"), text_color="#334155")
        self.lbl_page.pack(side="left", padx=12)
        self.btn_next = ctk.CTkButton(page_bar, text=self.tr_dict["page_next"], width=115, height=28, fg_color="#E2E8F0", text_color="#1E293B", command=self._next_page)
        self.btn_next.pack(side="left")

        self.scroll_trans = ctk.CTkScrollableFrame(self.tab_trans, fg_color="#F8FAFC", corner_radius=8, border_width=1, border_color="#E2E8F0")
        self.scroll_trans.pack(fill="both", expand=True, padx=10, pady=4)

        self._render_transcript_page()

    def _sync_current_page(self):
        for idx, (spk_var, tb, note_entry) in self.row_widgets.items():
            if 0 <= idx < len(self.segments):
                if tb.winfo_exists():
                    self.segments[idx]["text"] = tb.get("1.0", "end-1c").strip()
                if note_entry.winfo_exists():
                    self.segments[idx]["note"] = note_entry.get().strip()
                self.segments[idx]["speaker"] = spk_var.get().strip()

    def _apply_bulk_speaker_rename(self):
        self._sync_current_page()
        s1 = self.map_spk1_var.get().strip() or "Görüşmeci"
        s2 = self.map_spk2_var.get().strip() or "P01"
        for s in (s1, s2):
            if s not in self.speaker_list:
                self.speaker_list.insert(0, s)
        for seg in self.segments:
            if seg.get("speaker") in ("Konuşmacı 1", "Speaker 1"):
                seg["speaker"] = s1
            elif seg.get("speaker") in ("Konuşmacı 2", "Speaker 2"):
                seg["speaker"] = s2
        self._render_transcript_page()

    def _total_pages(self):
        return max(1, (len(self.segments) + self.PAGE_SIZE - 1) // self.PAGE_SIZE)

    def _prev_page(self):
        if self.current_page > 0:
            self._sync_current_page()
            self.current_page -= 1
            self._render_transcript_page()

    def _next_page(self):
        if self.current_page + 1 < self._total_pages():
            self._sync_current_page()
            self.current_page += 1
            self._render_transcript_page()

    def _render_transcript_page(self):
        for w in self.scroll_trans.winfo_children():
            w.destroy()
        self.row_widgets.clear()

        total_p = self._total_pages()
        self.lbl_page.configure(text=f"{self.tr_dict['page_lbl']} {self.current_page + 1} / {total_p}  ({len(self.segments)} blok)")
        self.btn_prev.configure(state="normal" if self.current_page > 0 else "disabled")
        self.btn_next.configure(state="normal" if self.current_page + 1 < total_p else "disabled")

        start_i = self.current_page * self.PAGE_SIZE
        end_i = min(len(self.segments), start_i + self.PAGE_SIZE)

        for idx in range(start_i, end_i):
            seg = self.segments[idx]
            row = ctk.CTkFrame(self.scroll_trans, fg_color="#FFFFFF", corner_radius=8, border_width=1, border_color="#E2E8F0")
            row.pack(fill="x", padx=4, pady=3)

            ts_txt = f"▶ {format_ts_hms(seg['start'])}"
            ctk.CTkButton(
                row,
                text=ts_txt,
                width=105,
                height=34,
                fg_color="#1E293B",
                hover_color="#334155",
                font=ctk.CTkFont(family="Consolas", size=12, weight="bold"),
                command=lambda s=seg["start"], e=seg["end"]: self.previewer.play_segment(self.media_path, s, e),
            ).pack(side="left", padx=6, pady=6)

            curr_spk = seg.get("speaker", "Konuşmacı 1")
            if curr_spk not in self.speaker_list:
                self.speaker_list.append(curr_spk)
            spk_var = ctk.StringVar(value=curr_spk)
            spk_menu = ctk.CTkOptionMenu(
                row,
                variable=spk_var,
                values=self.speaker_list,
                width=125,
                height=34,
                fg_color="#FEE2E2",
                text_color="#991B1B",
                button_color="#FECACA",
                button_hover_color="#FCA5A5",
                font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            )
            spk_menu.pack(side="left", padx=(0, 6), pady=6)

            tb = tk.Text(
                row,
                height=2,
                wrap="word",
                bg="#FFFFFF",
                fg="#0F172A",
                insertbackground="#0F172A",
                relief="solid",
                bd=1,
                padx=8,
                pady=4,
                font=("Segoe UI", 11),
            )
            tb.pack(side="left", fill="x", expand=True, padx=(0, 6), pady=6)
            tb.insert("1.0", seg.get("text", ""))

            note_entry = ctk.CTkEntry(
                row,
                width=220,
                height=34,
                placeholder_text="💬 Saha notu ekle...",
                fg_color="#FFFBEB",
                text_color="#92400E",
                border_color="#FDE68A",
                font=ctk.CTkFont(family="Segoe UI", size=12),
            )
            note_entry.pack(side="right", padx=(0, 8), pady=6)
            if seg.get("note"):
                note_entry.insert(0, seg.get("note"))

            self.row_widgets[idx] = (spk_var, tb, note_entry)

    def _build_tab_suggestions(self):
        ctk.CTkLabel(
            self.tab_sugg,
            text=self.tr_dict["sugg_desc"],
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color="#B45309",
            wraplength=980,
            justify="left",
        ).pack(anchor="w", padx=12, pady=8)

        self.scroll_sugg = ctk.CTkScrollableFrame(self.tab_sugg, fg_color="#F8FAFC", corner_radius=8)
        self.scroll_sugg.pack(fill="both", expand=True, padx=12, pady=6)
        self._render_suggestions()

    def _render_suggestions(self):
        for w in self.scroll_sugg.winfo_children():
            w.destroy()
        if not self.suggestions:
            ctk.CTkLabel(self.scroll_sugg, text=self.tr_dict["no_sugg"], text_color="#64748B", font=ctk.CTkFont(size=14)).pack(pady=40)
            return

        for i, sugg in enumerate(self.suggestions):
            card = ctk.CTkFrame(self.scroll_sugg, fg_color="#FFFFFF", corner_radius=8, border_width=1, border_color="#E2E8F0")
            card.pack(fill="x", padx=8, pady=5)

            ts = sugg.get("timestamp", "00:00:00")
            orig_w = sugg.get("original_word", "")
            sugg_w = sugg.get("suggested_word", "")
            reason = sugg.get("reason", "")
            status = sugg.get("status", "pending")

            lbl_txt = f"[{ts}]  '{orig_w}'  →  '{sugg_w}'   ({reason})"
            ctk.CTkLabel(card, text=lbl_txt, font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"), text_color="#1E293B").pack(side="left", padx=12, pady=10)

            status_lbl = ctk.CTkLabel(card, text="", font=ctk.CTkFont(size=12, weight="bold"))
            status_lbl.pack(side="right", padx=12)

            def set_status(idx=i, new_st="accepted", lbl=status_lbl):
                self.suggestions[idx]["status"] = new_st
                if new_st == "accepted":
                    lbl.configure(text=self.tr_dict["status_accepted"], text_color="#10B981")
                else:
                    lbl.configure(text=self.tr_dict["status_rejected"], text_color="#EF4444")

            if status == "accepted":
                status_lbl.configure(text=self.tr_dict["status_accepted"], text_color="#10B981")
            elif status == "rejected":
                status_lbl.configure(text=self.tr_dict["status_rejected"], text_color="#EF4444")

            ctk.CTkButton(card, text=self.tr_dict["btn_reject"], width=95, height=30, fg_color="#FEE2E2", text_color="#991B1B", hover_color="#FECACA", command=lambda idx=i, l=status_lbl: set_status(idx, "rejected", l)).pack(side="right", padx=4, pady=8)
            ctk.CTkButton(card, text=self.tr_dict["btn_accept"], width=95, height=30, fg_color="#D1FAE5", text_color="#065F46", hover_color="#A7F3D0", command=lambda idx=i, l=status_lbl: set_status(idx, "accepted", l)).pack(side="right", padx=4, pady=8)

    def _build_tab_anonymization(self):
        ctk.CTkLabel(
            self.tab_anon,
            text=self.tr_dict["anon_desc"],
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color="#1E293B",
            wraplength=980,
            justify="left",
        ).pack(anchor="w", padx=12, pady=8)

        add_bar = ctk.CTkFrame(self.tab_anon, fg_color="#F1F5F9", corner_radius=8)
        add_bar.pack(fill="x", padx=12, pady=4)
        ctk.CTkLabel(add_bar, text=self.tr_dict["anon_add_orig"], font=ctk.CTkFont(size=12, weight="bold"), text_color="#334155").pack(side="left", padx=10, pady=8)
        self.new_anon_orig = ctk.CTkEntry(add_bar, width=200, height=30, fg_color="#FFFFFF")
        self.new_anon_orig.pack(side="left", padx=6)
        self.new_anon_type = ctk.CTkOptionMenu(add_bar, values=["KİŞİ", "YER", "KURUM"], width=110, height=30, fg_color="#FFFFFF", text_color="#1E293B")
        self.new_anon_type.pack(side="left", padx=6)
        ctk.CTkButton(add_bar, text=self.tr_dict["anon_add_btn"], width=130, height=30, fg_color="#2563EB", command=self._add_custom_anon).pack(side="left", padx=8)

        self.scroll_anon = ctk.CTkScrollableFrame(self.tab_anon, fg_color="#F8FAFC", corner_radius=8)
        self.scroll_anon.pack(fill="both", expand=True, padx=12, pady=6)
        self.anon_vars = []
        self._render_anonymization_list()

    def _add_custom_anon(self):
        orig = self.new_anon_orig.get().strip()
        if not orig:
            return
        etype = self.new_anon_type.get()
        count = sum(1 for a in self.anon_candidates if a["type"] == etype) + 1
        self.anon_candidates.append({
            "original": orig,
            "code": f"[{etype}-{count}]",
            "type": etype,
            "approved": True,
        })
        self.new_anon_orig.delete(0, "end")
        self._render_anonymization_list()

    def _render_anonymization_list(self):
        for w in self.scroll_anon.winfo_children():
            w.destroy()
        self.anon_vars.clear()

        for i, item in enumerate(self.anon_candidates):
            row = ctk.CTkFrame(self.scroll_anon, fg_color="#FFFFFF", corner_radius=8, border_width=1, border_color="#E2E8F0")
            row.pack(fill="x", padx=8, pady=4)
            var = ctk.BooleanVar(value=item.get("approved", True))
            self.anon_vars.append((i, var))
            cb = ctk.CTkCheckBox(
                row,
                text=f"{item['original']}   ➔   {item['code']}   ({item['type']})",
                variable=var,
                font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
                text_color="#0F172A",
            )
            cb.pack(side="left", padx=12, pady=10)

    def _on_save_and_close(self):
        self.previewer.stop()
        self._sync_current_page()
        for idx, var in self.anon_vars:
            if 0 <= idx < len(self.anon_candidates):
                self.anon_candidates[idx]["approved"] = bool(var.get())
        segs = self.segments
        suggs = self.suggestions
        anons = self.anon_candidates
        self.destroy()
        if self.on_complete:
            self.on_complete(segs, suggs, anons)


class ScreenSnipperOverlay(tk.Toplevel):
    """
    Interactive full-screen snipping tool (like Windows Snipping Tool).
    Minimizes ILBER, displays the frozen screen on a full-screen canvas so there is zero DPI shift,
    lets the user drag a rectangle over any question/slide/table, saves the cropped image,
    restores ILBER, and invokes on_capture(saved_png_path).
    """
    def __init__(self, parent, save_dir, ui_lang="tr", on_capture=None):
        super().__init__(parent)
        self.parent = parent
        self.save_dir = save_dir
        self.ui_lang = ui_lang
        self.on_capture = on_capture
        self.start_x = None
        self.start_y = None
        self.rect_id = None
        self._bg_photo = None

        self.withdraw()
        try:
            self.parent.withdraw()
        except Exception:
            pass

        self.after(320, self._init_capture_screen)

    def _init_capture_screen(self):
        from PIL import ImageGrab, ImageTk, ImageEnhance

        try:
            self.full_img = ImageGrab.grab()
        except Exception:
            self._cancel()
            return

        self.deiconify()
        self.attributes("-fullscreen", True)
        self.attributes("-topmost", True)
        self.configure(bg="#000000", cursor="cross")

        self.canvas = tk.Canvas(self, bg="#000000", highlightthickness=0, cursor="cross")
        self.canvas.pack(fill="both", expand=True)
        self.update_idletasks()
        self.update()

        cw = max(1, self.canvas.winfo_width())
        ch = max(1, self.canvas.winfo_height())
        img_w, img_h = self.full_img.size
        if cw <= 10 or ch <= 10:
            cw, ch = img_w, img_h

        self.canvas_w = cw
        self.canvas_h = ch

        # Render darkened screenshot directly on canvas so mouse coordinates match what the user sees 1:1
        try:
            disp_img = self.full_img
            if (cw, ch) != (img_w, img_h):
                disp_img = self.full_img.resize((cw, ch))
            darkened = ImageEnhance.Brightness(disp_img).enhance(0.65)
            self._bg_photo = ImageTk.PhotoImage(darkened)
            self.canvas.create_image(0, 0, image=self._bg_photo, anchor="nw")
        except Exception:
            self.attributes("-alpha", 0.35)

        hint_map = {
            "tr": "📸 Soruyu çerçeve içine alıp bırakın (Kenarlardan biraz boşluk bırakın — İptal: ESC)",
            "en": "📸 Drag a box around the question and release (Press ESC to cancel)",
            "de": "📸 Ziehen Sie einen Rahmen um die Frage und lassen Sie los (ESC zum Abbrechen)",
        }
        hint_txt = hint_map.get(self.ui_lang, hint_map["tr"])
        self.canvas.create_rectangle(cw // 2 - 360, 18, cw // 2 + 360, 58, fill="#0F172A", outline="#38BDF8", width=2)
        self.canvas.create_text(cw // 2, 38, text=hint_txt, fill="#FFFFFF", font=("Segoe UI", 12, "bold"))

        self.canvas.bind("<ButtonPress-1>", self._on_press)
        self.canvas.bind("<B1-Motion>", self._on_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_release)
        self.bind("<Escape>", lambda e: self._cancel())
        self.focus_force()

    def _on_press(self, event):
        self.start_x = self.canvas.canvasx(event.x)
        self.start_y = self.canvas.canvasy(event.y)
        if self.rect_id:
            self.canvas.delete(self.rect_id)
        self.rect_id = self.canvas.create_rectangle(
            self.start_x, self.start_y, self.start_x, self.start_y,
            outline="#00E5FF", width=3
        )

    def _on_drag(self, event):
        if self.rect_id and self.start_x is not None:
            cur_x = self.canvas.canvasx(event.x)
            cur_y = self.canvas.canvasy(event.y)
            self.canvas.coords(self.rect_id, self.start_x, self.start_y, cur_x, cur_y)

    def _on_release(self, event):
        import datetime
        if self.start_x is None or self.start_y is None:
            self._cancel()
            return

        cur_x = self.canvas.canvasx(event.x)
        cur_y = self.canvas.canvasy(event.y)
        x1, x2 = sorted([self.start_x, cur_x])
        y1, y2 = sorted([self.start_y, cur_y])

        img_w, img_h = self.full_img.size
        cw = max(1, getattr(self, "canvas_w", img_w))
        ch = max(1, getattr(self, "canvas_h", img_h))
        scale_x = img_w / float(cw)
        scale_y = img_h / float(ch)

        # Add a 12px padding margin around user selection so first/last letters are never clipped
        pad = 12
        crop_box = (
            max(0, int(x1 * scale_x) - pad),
            max(0, int(y1 * scale_y) - pad),
            min(img_w, int(x2 * scale_x) + pad),
            min(img_h, int(y2 * scale_y) + pad),
        )

        self.destroy()
        try:
            self.parent.deiconify()
            self.parent.lift()
            self.parent.focus_force()
        except Exception:
            pass

        if (crop_box[2] - crop_box[0]) < 20 or (crop_box[3] - crop_box[1]) < 20:
            return

        cropped = self.full_img.crop(crop_box)
        os.makedirs(self.save_dir, exist_ok=True)
        out_path = os.path.join(self.save_dir, f"Ekran_Alintisi_{datetime.datetime.now().strftime('%H%M%S')}.png")
        cropped.save(out_path)
        if self.on_capture:
            self.on_capture(out_path)

    def _cancel(self):
        self.destroy()
        try:
            self.parent.deiconify()
            self.parent.lift()
            self.parent.focus_force()
        except Exception:
            pass


class DocumentTranslationEditor(ctk.CTkToplevel):
    """
    Interactive Pre-Export Document Translation Review & Edit Dialog for EDU.
    Allows the user to review and edit every translated line/paragraph side-by-side with the
    original source text (plus bulk Find & Replace) BEFORE writing the final .docx/.pdf/.xlsx/.txt!
    Deduplicates identical source lines so editing a repeated header/term once updates all occurrences.
    """
    PAGE_SIZE = 25

    def __init__(self, parent, source_texts, translated_texts, doc_name="", ui_lang="tr", on_complete=None):
        super().__init__(parent)
        self.raw_sources = list(source_texts)
        self.raw_translations = list(translated_texts)
        self.ui_lang = ui_lang if ui_lang in ("tr", "en", "de") else "tr"
        self.on_complete = on_complete
        self.current_page = 0
        self.row_widgets = {}  # unique_idx -> tk.Text

        # Build deduplicated items so repeated headers/values only need to be edited once
        self.unique_items = []  # list of dicts: {"src": str, "tgt": str, "indices": [int, ...]}
        seen_map = {}
        for idx, (s_txt, t_txt) in enumerate(zip(self.raw_sources, self.raw_translations)):
            key = s_txt.strip()
            if key in seen_map:
                self.unique_items[seen_map[key]]["indices"].append(idx)
            else:
                seen_map[key] = len(self.unique_items)
                self.unique_items.append({
                    "src": s_txt,
                    "tgt": t_txt,
                    "indices": [idx],
                })

        title_map = {
            "tr": f"EDU — Nihai Belge Öncesi Çeviri Kontrol & Düzenleme ({doc_name})",
            "en": f"EDU — Pre-Export Translation Review & Editor ({doc_name})",
            "de": f"EDU — Übersetzungsprüfung & Bearbeitung vor dem Export ({doc_name})",
        }
        self.title(title_map[self.ui_lang])
        self.geometry("1080x740")
        self.configure(fg_color="#F8FAFC")
        self.attributes("-topmost", True)
        self.protocol("WM_DELETE_WINDOW", self._on_save_and_close)

        # Header Banner
        hdr = ctk.CTkFrame(self, fg_color="#1E3A8A", corner_radius=0)
        hdr.pack(fill="x")
        hdr_title = {
            "tr": f"📝 Nihai Belge Oluşturulmadan Önce Çevirileri Kontrol Edin ve Düzenleyin ({len(self.unique_items)} Benzersiz Bölüm)",
            "en": f"📝 Review and Edit Translations Before Final Document Generation ({len(self.unique_items)} Unique Segments)",
            "de": f"📝 Übersetzungen vor der Dokumenterstellung prüfen und bearbeiten ({len(self.unique_items)} Abschnitte)",
        }[self.ui_lang]
        ctk.CTkLabel(
            hdr,
            text=hdr_title,
            font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold"),
            text_color="#FFFFFF",
        ).pack(anchor="w", padx=20, pady=(12, 2))

        hdr_sub = {
            "tr": "💡 Sağ sütundaki çevirileri dilediğiniz gibi değiştirebilirsiniz. Tekrarlanan satırları (örn. 4x) bir kez düzenlemeniz belgedeki tüm tekrarları otomatik günceller.",
            "en": "💡 Edit any translation in the right column. Editing a repeated item (e.g. 4x) automatically updates all occurrences in the document.",
            "de": "💡 Bearbeiten Sie die Übersetzungen in der rechten Spalte. Wiederholte Zeilen werden automatisch im gesamten Dokument aktualisiert.",
        }[self.ui_lang]
        ctk.CTkLabel(
            hdr,
            text=hdr_sub,
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color="#BFDBFE",
        ).pack(anchor="w", padx=20, pady=(0, 10))

        # Quick Find & Replace Bar
        fr_bar = ctk.CTkFrame(self, fg_color="#EFF6FF", border_width=1, border_color="#BFDBFE", corner_radius=10)
        fr_bar.pack(fill="x", padx=18, pady=(10, 4))

        ctk.CTkLabel(
            fr_bar,
            text={"tr": "🔍 Toplu Bul & Değiştir:", "en": "🔍 Bulk Find & Replace:", "de": "🔍 Suchen & Ersetzen:"}[self.ui_lang],
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color="#1E3A8A",
        ).pack(side="left", padx=(12, 6), pady=8)

        self.find_entry = ctk.CTkEntry(
            fr_bar,
            width=200,
            height=30,
            placeholder_text={"tr": "Çeviride aranan kelime...", "en": "Find word in translation...", "de": "Suchbegriff..."}[self.ui_lang],
            fg_color="#FFFFFF",
        )
        self.find_entry.pack(side="left", padx=(0, 6), pady=8)

        ctk.CTkLabel(fr_bar, text="➔", font=ctk.CTkFont(size=14, weight="bold"), text_color="#1E3A8A").pack(side="left", padx=2)

        self.repl_entry = ctk.CTkEntry(
            fr_bar,
            width=200,
            height=30,
            placeholder_text={"tr": "Yeni karşılığı...", "en": "Replace with...", "de": "Ersetzen durch..."}[self.ui_lang],
            fg_color="#FFFFFF",
        )
        self.repl_entry.pack(side="left", padx=(6, 8), pady=8)

        self.fr_status_lbl = ctk.CTkLabel(fr_bar, text="", font=ctk.CTkFont(size=12, weight="bold"), text_color="#059669")

        def _do_bulk_replace():
            self._sync_current_page_edits()
            f_str = self.find_entry.get().strip()
            r_str = self.repl_entry.get().strip()
            if not f_str:
                return
            changed_cnt = 0
            pat = re.compile(re.escape(f_str), re.IGNORECASE)
            for item in self.unique_items:
                if pat.search(item["tgt"]):
                    item["tgt"] = pat.sub(r_str, item["tgt"])
                    changed_cnt += 1
            self._render_page()
            msg = {
                "tr": f"✔ {changed_cnt} satırda değiştirildi",
                "en": f"✔ Replaced in {changed_cnt} rows",
                "de": f"✔ In {changed_cnt} Zeilen ersetzt",
            }[self.ui_lang]
            self.fr_status_lbl.configure(text=msg)

        ctk.CTkButton(
            fr_bar,
            text={"tr": "Tümünde Değiştir", "en": "Replace All", "de": "Alle ersetzen"}[self.ui_lang],
            width=130,
            height=30,
            fg_color="#2563EB",
            hover_color="#1D4ED8",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            command=_do_bulk_replace,
        ).pack(side="left", padx=(0, 10), pady=8)
        self.fr_status_lbl.pack(side="left", padx=6)

        # Column Headers
        col_hdr = ctk.CTkFrame(self, fg_color="#E2E8F0", corner_radius=8, height=32)
        col_hdr.pack(fill="x", padx=18, pady=(4, 2))
        ctk.CTkLabel(
            col_hdr,
            text={"tr": "  #    Orijinal Kaynak Metin (Referans)", "en": "  #    Original Source Text (Reference)", "de": "  #    Originaltext (Referenz)"}[self.ui_lang],
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color="#334155",
            width=460,
            anchor="w",
        ).pack(side="left", padx=10, pady=4)
        ctk.CTkLabel(
            col_hdr,
            text={"tr": "✏️ Çeviri Metni (Doğrudan Düzenleyebilirsiniz)", "en": "✏️ Translated Text (Editable)", "de": "✏️ Übersetzung (Bearbeitbar)"}[self.ui_lang],
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color="#1E3A8A",
            anchor="w",
        ).pack(side="left", fill="x", expand=True, padx=10, pady=4)

        # Bottom Pagination & Confirm Bar (packed before scroll frame so it's always visible)
        bot_bar = ctk.CTkFrame(self, fg_color="#FFFFFF", border_width=1, border_color="#CBD5E1", corner_radius=0)
        bot_bar.pack(side="bottom", fill="x")

        nav_frame = ctk.CTkFrame(bot_bar, fg_color="transparent")
        nav_frame.pack(side="left", padx=18, pady=10)

        self.prev_btn = ctk.CTkButton(
            nav_frame,
            text={"tr": "◀ Önceki Sayfa", "en": "◀ Prev Page", "de": "◀ Zurück"}[self.ui_lang],
            width=120,
            height=34,
            fg_color="#E2E8F0",
            text_color="#1E293B",
            hover_color="#CBD5E1",
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self._prev_page,
        )
        self.prev_btn.pack(side="left", padx=(0, 8))

        self.page_lbl = ctk.CTkLabel(nav_frame, text="1 / 1", font=ctk.CTkFont(size=13, weight="bold"), text_color="#1E293B", width=90)
        self.page_lbl.pack(side="left", padx=4)

        self.next_btn = ctk.CTkButton(
            nav_frame,
            text={"tr": "Sonraki Sayfa ▶", "en": "Next Page ▶", "de": "Weiter ▶"}[self.ui_lang],
            width=120,
            height=34,
            fg_color="#E2E8F0",
            text_color="#1E293B",
            hover_color="#CBD5E1",
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self._next_page,
        )
        self.next_btn.pack(side="left", padx=(8, 0))

        save_btn = ctk.CTkButton(
            bot_bar,
            text={
                "tr": "✓ Onayla ve Nihai Belgeyi Oluştur",
                "en": "✓ Approve & Generate Final Document",
                "de": "✓ Bestätigen & Dokument erstellen",
            }[self.ui_lang],
            height=42,
            width=320,
            fg_color="#10B981",
            hover_color="#059669",
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            command=self._on_save_and_close,
        )
        save_btn.pack(side="right", padx=18, pady=10)

        # Scrollable Rows Area
        self.scroll = ctk.CTkScrollableFrame(self, fg_color="#FFFFFF", border_width=1, border_color="#E2E8F0", corner_radius=10)
        self.scroll.pack(fill="both", expand=True, padx=18, pady=6)

        self._render_page()

    def _sync_current_page_edits(self):
        for u_idx, txt_widget in list(self.row_widgets.items()):
            try:
                if txt_widget.winfo_exists():
                    val = txt_widget.get("1.0", "end-1c").strip()
                    self.unique_items[u_idx]["tgt"] = val
            except Exception:
                pass

    def _render_page(self):
        for w in self.scroll.winfo_children():
            w.destroy()
        self.row_widgets.clear()

        total_items = len(self.unique_items)
        total_pages = max(1, (total_items + self.PAGE_SIZE - 1) // self.PAGE_SIZE)
        self.current_page = max(0, min(self.current_page, total_pages - 1))

        start_i = self.current_page * self.PAGE_SIZE
        end_i = min(total_items, start_i + self.PAGE_SIZE)

        self.page_lbl.configure(text=f"{self.current_page + 1} / {total_pages}")
        self.prev_btn.configure(state="normal" if self.current_page > 0 else "disabled")
        self.next_btn.configure(state="normal" if self.current_page < total_pages - 1 else "disabled")

        for u_idx in range(start_i, end_i):
            item = self.unique_items[u_idx]
            src_str = item["src"]
            tgt_str = item["tgt"]
            occ_count = len(item["indices"])

            row_f = ctk.CTkFrame(
                self.scroll,
                fg_color="#F8FAFC" if u_idx % 2 == 0 else "#FFFFFF",
                corner_radius=8,
                border_width=1,
                border_color="#E2E8F0",
            )
            row_f.pack(fill="x", padx=6, pady=3)

            badge_txt = f"#{u_idx + 1}" + (f" ({occ_count}x)" if occ_count > 1 else "")
            ctk.CTkLabel(
                row_f,
                text=badge_txt,
                width=62,
                font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
                text_color="#2563EB" if occ_count > 1 else "#64748B",
                anchor="w",
            ).pack(side="left", padx=(8, 4), pady=6)

            # Source text label (read-only selectable/clean)
            ctk.CTkLabel(
                row_f,
                text=src_str,
                width=390,
                wraplength=380,
                justify="left",
                anchor="w",
                font=ctk.CTkFont(family="Segoe UI", size=12),
                text_color="#334155",
            ).pack(side="left", padx=(4, 10), pady=6)

            # Editable target text box
            line_h = max(2, min(5, (max(len(src_str), len(tgt_str)) // 55) + 1))
            t_box = tk.Text(
                row_f,
                height=line_h,
                wrap="word",
                font=("Segoe UI", 11),
                bg="#FFFFFF",
                fg="#0F172A",
                relief="solid",
                bd=1,
                padx=6,
                pady=4,
            )
            t_box.insert("1.0", tgt_str)
            t_box.pack(side="left", fill="x", expand=True, padx=(4, 8), pady=5)
            self.row_widgets[u_idx] = t_box

    def _prev_page(self):
        self._sync_current_page_edits()
        if self.current_page > 0:
            self.current_page -= 1
            self._render_page()

    def _next_page(self):
        self._sync_current_page_edits()
        total_pages = max(1, (len(self.unique_items) + self.PAGE_SIZE - 1) // self.PAGE_SIZE)
        if self.current_page < total_pages - 1:
            self.current_page += 1
            self._render_page()

    def _on_save_and_close(self):
        self._sync_current_page_edits()
        final_list = list(self.raw_translations)
        for item in self.unique_items:
            edited_val = item["tgt"]
            for orig_idx in item["indices"]:
                # Preserve leading/trailing whitespace of original segment if any
                orig_s = str(self.raw_sources[orig_idx] or "")
                lead_ws = orig_s[: len(orig_s) - len(orig_s.lstrip())]
                trail_ws = orig_s[len(orig_s.rstrip()) :]
                final_list[orig_idx] = f"{lead_ws}{edited_val.strip()}{trail_ws}"
        self.destroy()
        if self.on_complete:
            self.on_complete(final_list)


