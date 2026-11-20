"""
ILBERTech Platform — FIELD Qualitative Research Project Workbench
Project-based qualitative research management:
- Portable project folders under ILBER_Workspace/Projects/FIELD/<ProjectName>/
- Multi-interview queue with participant codes (K01, K02... or focus group multi-codes)
- Block transcript editor with per-block audio playback, low-confidence word flags,
  quote tagging with themes, line/file field notes, anonymization suggestions, and Review approval
- Project-wide Quotes (.docx + audio jump), Coded Participant Profiles (Word + Excel, custom fields, consent)
- Project search, external Word/TXT transcript import, Word/PDF custom reports, Method Summary, NVivo/MAXQDA export
- Password-encrypted Ethics Key (_ETHICS_KEY_DO_NOT_SHARE.ilberkey) strictly excluded from all exports
"""

import os
import re
import json
import hmac
import hashlib
import shutil
import zipfile
import datetime
import threading
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk

import docx
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_COLOR_INDEX

from ilber_core_utils import (
    open_path_crossplatform,
    format_ts_hms,
    add_word_comment_or_note,
)


# ---------------------------------------------------------------------------
# Password-Encrypted Ethics Key Helpers (PBKDF2-HMAC-SHA256 + Keystream + HMAC)
# ---------------------------------------------------------------------------

def _derive_key_and_mac(password: str, salt: bytes):
    dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, 120_000, dklen=64)
    return dk[:32], dk[32:]


def _xor_keystream(data: bytes, enc_key: bytes, nonce: bytes) -> bytes:
    out = bytearray(len(data))
    counter = 0
    pos = 0
    while pos < len(data):
        block = hashlib.sha256(enc_key + nonce + counter.to_bytes(8, "big")).digest()
        chunk = min(32, len(data) - pos)
        for i in range(chunk):
            out[pos + i] = data[pos + i] ^ block[i]
        pos += chunk
        counter += 1
    return bytes(out)


def save_encrypted_ethics_key(filepath: str, password: str, payload: dict):
    salt = os.urandom(16)
    nonce = os.urandom(16)
    enc_key, mac_key = _derive_key_and_mac(password, salt)
    raw = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    ct = _xor_keystream(raw, enc_key, nonce)
    tag = hmac.new(mac_key, salt + nonce + ct, hashlib.sha256).digest()
    blob = {
        "v": 1,
        "salt": salt.hex(),
        "nonce": nonce.hex(),
        "tag": tag.hex(),
        "ct": ct.hex(),
    }
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(blob, f, indent=2)


def load_encrypted_ethics_key(filepath: str, password: str) -> dict:
    with open(filepath, "r", encoding="utf-8") as f:
        blob = json.load(f)
    salt = bytes.fromhex(blob["salt"])
    nonce = bytes.fromhex(blob["nonce"])
    tag = bytes.fromhex(blob["tag"])
    ct = bytes.fromhex(blob["ct"])
    enc_key, mac_key = _derive_key_and_mac(password, salt)
    expected_tag = hmac.new(mac_key, salt + nonce + ct, hashlib.sha256).digest()
    if not hmac.compare_digest(tag, expected_tag):
        raise ValueError("Parola hatalı veya dosya bozulmuş / Invalid password.")
    raw = _xor_keystream(ct, enc_key, nonce)
    return json.loads(raw.decode("utf-8"))


# ---------------------------------------------------------------------------
# Anonymization Candidate Detector (Person / Place / Contact)
# ---------------------------------------------------------------------------

_NON_ENTITY_STARTERS = {
    "Ben", "Sen", "Biz", "Siz", "Onlar", "Bu", "Şu", "Evet", "Hayır", "Ama", "Fakat",
    "Çünkü", "Yani", "İşte", "Sonra", "Önce", "Burada", "Orada", "Şimdi", "Zaten",
    "Mesela", "Örneğin", "Tabii", "Aslında", "Bence", "Böyle", "Şöyle", "Nasıl",
    "Neden", "Niye", "Kim", "Hangi", "Kendi", "Hep", "Çok", "Daha", "Bir", "İki",
    "The", "And", "But", "Yes", "No", "Well", "Okay", "Then", "When", "Where",
    "What", "Why", "How", "This", "That", "There", "Here", "They", "We", "You",
    "Konuşmacı", "Görüşmeci", "Speaker", "Interviewer",
}

_PLACE_HINTS = {
    "İstanbul", "Ankara", "İzmir", "Bursa", "Antalya", "Adana", "Konya", "Maltepe",
    "Kadıköy", "Beşiktaş", "Üsküdar", "Bağcılar", "Fatih", "Bakırköy", "Beyoğlu",
    "Türkiye", "Almanya", "Berlin", "Londra", "Paris", "Mahalle", "Mahallesi",
    "İlçe", "İlçesi", "Köy", "Köyü", "Hastane", "Hastanesi", "Üniversitesi", "Okulu",
}


def detect_anonymization_candidates(segments):
    """Scan transcript segments for potential person names, places, and phone/email."""
    found = {}  # original -> {"original": str, "category": str, "replacement": str, "count": int}
    person_idx = 1
    place_idx = 1
    contact_idx = 1

    email_re = re.compile(r"\b[\w\.-]+@[\w\.-]+\.\w{2,}\b")
    phone_re = re.compile(r"\b(?:\+90|0)?\s*5\d{2}[\s\-]?\d{3}[\s\-]?\d{2}[\s\-]?\d{2}\b")
    proper_re = re.compile(r"\b([A-ZÇĞİÖŞÜ][a-zçğıöşü]{2,}(?:\s+[A-ZÇĞİÖŞÜ][a-zçğıöşü]{2,})?)\b")

    for seg in segments or []:
        txt = str(seg.get("text", ""))
        for m in email_re.findall(txt):
            if m not in found:
                found[m] = {"original": m, "category": "İletişim", "replacement": f"[İLETİŞİM_{contact_idx}]", "count": 0}
                contact_idx += 1
            found[m]["count"] += 1
        for m in phone_re.findall(txt):
            if m not in found:
                found[m] = {"original": m, "category": "Telefon", "replacement": f"[TELEFON_{contact_idx}]", "count": 0}
                contact_idx += 1
            found[m]["count"] += 1
        for m in proper_re.findall(txt):
            first_w = m.split()[0]
            if first_w in _NON_ENTITY_STARTERS or m in _NON_ENTITY_STARTERS:
                continue
            if m not in found:
                words = set(m.split())
                if words & _PLACE_HINTS or any(h in m for h in ("Mahalle", "İlçe", "Hastane", "Üniversite", "Okul", "Cadde", "Sokak")):
                    found[m] = {"original": m, "category": "Yer / Kurum", "replacement": f"[YER_{place_idx}]", "count": 0}
                    place_idx += 1
                else:
                    found[m] = {"original": m, "category": "Kişi", "replacement": f"[KİŞİ_{person_idx}]", "count": 0}
                    person_idx += 1
            found[m]["count"] += 1

    return sorted(found.values(), key=lambda x: x["count"], reverse=True)


# ---------------------------------------------------------------------------
# FIELD Project Manager (All paths relative for 100% portability)
# ---------------------------------------------------------------------------

class FieldProjectManager:
    DEFAULT_PROFILE_FIELDS = ["Yaş", "Cinsiyet", "Meslek / Rol", "Eğitim", "Bölge / Mahalle"]

    def __init__(self, project_dir: str):
        self.project_dir = os.path.abspath(project_dir)
        self.interviews_dir = self.project_dir
        os.makedirs(self.interviews_dir, exist_ok=True)
        self.project_json_path = os.path.join(self.project_dir, "project.json")
        self.ethics_file_path = os.path.join(self.project_dir, "_ETHICS_KEY_DO_NOT_SHARE.ilberkey")
        self.data = self._load_or_init()

    def _load_or_init(self):
        if os.path.isfile(self.project_json_path):
            try:
                with open(self.project_json_path, "r", encoding="utf-8") as f:
                    d = json.load(f)
                d.setdefault("project_name", os.path.basename(self.project_dir))
                d.setdefault("language", "Auto")
                d.setdefault("speaker_count", 2)
                d.setdefault("profile_fields", list(self.DEFAULT_PROFILE_FIELDS))
                d.setdefault("participants", {})
                d.setdefault("interviews", [])
                return d
            except Exception:
                pass
        d = {
            "project_name": os.path.basename(self.project_dir),
            "created_at": datetime.datetime.now().isoformat(),
            "language": "Auto",
            "speaker_count": 2,
            "profile_fields": list(self.DEFAULT_PROFILE_FIELDS),
            "participants": {},
            "interviews": [],
        }
        self.save(d)
        return d

    def save(self, data=None):
        if data is not None:
            self.data = data
        with open(self.project_json_path, "w", encoding="utf-8") as f:
            json.dump(self.data, f, ensure_ascii=False, indent=2)
        # Keep .ilber_meta.json synced for My Projects & History cards
        try:
            meta_p = os.path.join(self.project_dir, ".ilber_meta.json")
            with open(meta_p, "w", encoding="utf-8") as mf:
                json.dump({
                    "section": "FIELD",
                    "project": self.data.get("project_name", os.path.basename(self.project_dir)),
                    "updated_at": datetime.datetime.now().isoformat(),
                }, mf, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def next_participant_code(self) -> str:
        existing = set(self.data.get("participants", {}).keys())
        for iv in self.data.get("interviews", []):
            for c in iv.get("participant_codes", []):
                existing.add(c)
        n = 1
        while f"K{n:02d}" in existing:
            n += 1
        return f"K{n:02d}"

    def ensure_participant(self, code: str):
        code = code.strip().upper()
        if not code:
            return
        parts = self.data.setdefault("participants", {})
        if code not in parts:
            parts[code] = {
                "code": code,
                "consent": "Onam Alındı",
                "fields": {fld: "" for fld in self.data.get("profile_fields", [])},
            }

    def abs_path(self, rel_path: str) -> str:
        if not rel_path:
            return ""
        if os.path.isabs(rel_path):
            return rel_path
        return os.path.join(self.project_dir, rel_path.replace("/", os.sep))

    def rel_path(self, full_path: str) -> str:
        return os.path.relpath(full_path, self.project_dir).replace("\\", "/")

    def add_interview_record(self, src_media_path: str, participant_codes: list, speaker_count: int, is_focus_group: bool = False):
        codes = [c.strip().upper() for c in participant_codes if c.strip()]
        if not codes:
            codes = [self.next_participant_code()]
        for c in codes:
            self.ensure_participant(c)

        code_tag = "_".join(codes[:3])
        iv_folder_name = code_tag
        iv_dir = os.path.join(self.interviews_dir, iv_folder_name)
        if os.path.exists(iv_dir):
            stamp = datetime.datetime.now().strftime("%H%M%S")
            iv_folder_name = f"{code_tag}_{stamp}"
            iv_dir = os.path.join(self.interviews_dir, iv_folder_name)
        os.makedirs(iv_dir, exist_ok=True)

        dest_media = os.path.join(iv_dir, os.path.basename(src_media_path))
        if os.path.abspath(src_media_path) != os.path.abspath(dest_media):
            shutil.copy2(src_media_path, dest_media)

        iv_id = iv_folder_name
        rec = {
            "id": iv_id,
            "folder_rel": self.rel_path(iv_dir),
            "audio_rel": self.rel_path(dest_media),
            "transcript_rel": self.rel_path(os.path.join(iv_dir, "transcript.json")),
            "notes_rel": self.rel_path(os.path.join(iv_dir, "field_notes.txt")),
            "docx_rel": self.rel_path(os.path.join(iv_dir, f"{code_tag}_Desifre.docx")),
            "participant_codes": codes,
            "is_focus_group": bool(is_focus_group or len(codes) > 1),
            "speaker_count": int(speaker_count),
            "status": "queued",  # queued | ready | reviewed
            "duration_sec": 0.0,
            "general_note": "",
            "created_at": datetime.datetime.now().isoformat(),
        }
        self.data["interviews"].append(rec)
        self.save()
        return rec

    def delete_interview_record(self, iv_id: str):
        """Remove an interview from project.json and delete its folder from disk."""
        ivs = self.data.get("interviews", [])
        target = None
        remaining = []
        for iv in ivs:
            if iv.get("id") == iv_id and target is None:
                target = iv
            else:
                remaining.append(iv)
        self.data["interviews"] = remaining
        self.save()
        if target and target.get("folder_rel"):
            fdir = self.abs_path(target["folder_rel"])
            if os.path.isdir(fdir):
                shutil.rmtree(fdir, ignore_errors=True)

    def clear_queued_interviews(self) -> int:
        """Delete all interviews still in 'queued' status."""
        q_ids = [iv["id"] for iv in self.data.get("interviews", []) if iv.get("status") == "queued"]
        for qid in q_ids:
            self.delete_interview_record(qid)
        return len(q_ids)

    def import_external_transcript(self, doc_path: str, participant_codes: list, optional_audio_path: str = ""):
        """Import a Word (.docx) or TXT (.txt) transcript written elsewhere into the project."""
        codes = [c.strip().upper() for c in participant_codes if c.strip()]
        if not codes:
            codes = [self.next_participant_code()]
        for c in codes:
            self.ensure_participant(c)

        code_tag = "_".join(codes[:3])
        base_slug = re.sub(r"[^A-Za-z0-9_\-çğıöşüÇĞİÖŞÜ]+", "_", os.path.basename(doc_path).rsplit(".", 1)[0]).strip("_")[:32] or "IceAktarim"
        stamp = datetime.datetime.now().strftime("%H%M%S")
        iv_folder_name = f"{code_tag}_{base_slug}_{stamp}"
        iv_dir = os.path.join(self.interviews_dir, iv_folder_name)
        os.makedirs(iv_dir, exist_ok=True)

        audio_rel = ""
        if optional_audio_path and os.path.isfile(optional_audio_path):
            dest_audio = os.path.join(iv_dir, os.path.basename(optional_audio_path))
            shutil.copy2(optional_audio_path, dest_audio)
            audio_rel = self.rel_path(dest_audio)

        # Parse paragraphs/lines into blocks
        lines = []
        ext = doc_path.lower().rsplit(".", 1)[-1]
        if ext == "docx":
            d = docx.Document(doc_path)
            for p in d.paragraphs:
                t = p.text.strip()
                if t:
                    lines.append(t)
            for tbl in d.tables:
                for row in tbl.rows:
                    cells = [c.text.strip() for c in row.cells if c.text.strip()]
                    if cells:
                        lines.append(" | ".join(cells))
        else:
            with open(doc_path, "r", encoding="utf-8", errors="replace") as f:
                lines = [ln.strip() for ln in f if ln.strip()]

        ts_re = re.compile(r"^\[(\d{2}:\d{2}(?::\d{2})?)\]\s*(?:([^:]+):\s*)?(.*)$")
        segments = []
        cur_t = 0.0
        for ln in lines:
            m = ts_re.match(ln)
            if m:
                parts = [int(x) for x in m.group(1).split(":")]
                st = parts[0] * 60 + parts[1] if len(parts) == 2 else parts[0] * 3600 + parts[1] * 60 + parts[2]
                spk = (m.group(2) or codes[0]).strip()
                txt = m.group(3).strip()
            else:
                st = cur_t
                spk = codes[0]
                txt = ln
                cur_t += 5.0
            segments.append({
                "start": float(st),
                "end": float(st + 5.0),
                "speaker": spk,
                "text": txt,
                "low_conf_words": [],
                "is_quote": False,
                "quote_theme": "",
                "note": "",
            })

        rec = {
            "id": iv_folder_name,
            "folder_rel": self.rel_path(iv_dir),
            "audio_rel": audio_rel,
            "transcript_rel": self.rel_path(os.path.join(iv_dir, "transcript.json")),
            "notes_rel": self.rel_path(os.path.join(iv_dir, "field_notes.txt")),
            "docx_rel": self.rel_path(os.path.join(iv_dir, f"{code_tag}_Desifre.docx")),
            "participant_codes": codes,
            "is_focus_group": len(codes) > 1,
            "speaker_count": max(1, len(codes)),
            "status": "ready",
            "duration_sec": float(segments[-1]["end"]) if segments else 0.0,
            "general_note": "",
            "created_at": datetime.datetime.now().isoformat(),
        }
        self.save_interview_transcript(rec, segments, general_note="")
        self.data["interviews"].append(rec)
        self.save()
        return rec

    def delete_interview_record(self, iv_id: str):
        """Remove an interview from project.json and delete its folder from disk."""
        ivs = self.data.get("interviews", [])
        target = None
        remaining = []
        for iv in ivs:
            if iv.get("id") == iv_id and target is None:
                target = iv
            else:
                remaining.append(iv)
        self.data["interviews"] = remaining
        self.save()
        if target and target.get("folder_rel"):
            fdir = self.abs_path(target["folder_rel"])
            if os.path.isdir(fdir):
                shutil.rmtree(fdir, ignore_errors=True)

    def clear_queued_interviews(self) -> int:
        """Delete all interviews still in 'queued' status."""
        q_ids = [iv["id"] for iv in self.data.get("interviews", []) if iv.get("status") == "queued"]
        for qid in q_ids:
            self.delete_interview_record(qid)
        return len(q_ids)

    def load_interview_transcript(self, iv_rec: dict) -> list:
        tp = self.abs_path(iv_rec.get("transcript_rel", ""))
        if tp and os.path.isfile(tp):
            try:
                with open(tp, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        return []

    def save_interview_transcript(self, iv_rec: dict, segments: list, general_note: str = None):
        if general_note is not None:
            iv_rec["general_note"] = general_note
        if segments:
            iv_rec["duration_sec"] = max(float(iv_rec.get("duration_sec", 0.0)), float(segments[-1].get("end", 0.0)))
        tp = self.abs_path(iv_rec["transcript_rel"])
        os.makedirs(os.path.dirname(tp), exist_ok=True)
        with open(tp, "w", encoding="utf-8") as f:
            json.dump(segments, f, ensure_ascii=False, indent=2)

        # Save field_notes.txt inside the interview folder
        np_path = self.abs_path(iv_rec["notes_rel"])
        with open(np_path, "w", encoding="utf-8") as nf:
            nf.write(f"Proje: {self.data.get('project_name')}\n")
            nf.write(f"Katılımcı Kodları: {', '.join(iv_rec.get('participant_codes', []))}\n")
            nf.write(f"Durum: {iv_rec.get('status')}\n")
            nf.write("=" * 55 + "\n")
            nf.write(f"GENEL SAHA NOTU:\n{iv_rec.get('general_note', '').strip() or '(Not girilmedi)'}\n")
            nf.write("=" * 55 + "\nSATIR NOTLARI:\n")
            for s in segments:
                if str(s.get("note", "")).strip():
                    nf.write(f"[{format_ts_hms(s['start'])}] ({s.get('speaker', '')}): {s['note'].strip()}\n")

        # Export interview Word (.docx) with highlighted quotes + Word comments for field notes
        self.export_interview_docx(iv_rec, segments)
        # Export this participant's quotes directly inside their own subfolder (K01/, K02/...)
        self.export_interview_quotes_docx(iv_rec, segments)
        self.save()

    def export_interview_quotes_docx(self, iv_rec: dict, segments: list) -> str:
        """Export marked quotes for a single interview directly inside its own participant subfolder."""
        codes_str = "_".join(iv_rec.get("participant_codes", [])) or iv_rec["id"]
        iv_dir = self.abs_path(iv_rec["folder_rel"])
        os.makedirs(iv_dir, exist_ok=True)
        out_path = os.path.join(iv_dir, f"{codes_str}_Alintilar.docx")
        quotes = [s for s in (segments or []) if s.get("is_quote")]
        doc = docx.Document()
        doc.add_heading(f"FIELD Alıntılar — {codes_str}", 0)
        doc.add_paragraph(f"Proje: {self.data.get('project_name')} | Toplam Alıntı: {len(quotes)}")
        if not quotes:
            doc.add_paragraph("Bu görüşmede alıntı olarak işaretlenmiş satır bulunmuyor.")
        else:
            tbl = doc.add_table(rows=1, cols=4)
            tbl.style = "Table Grid"
            hdr = tbl.rows[0].cells
            hdr[0].text = "Zaman Damgası"
            hdr[1].text = "Konuşmacı"
            hdr[2].text = "Alıntı Metni"
            hdr[3].text = "Saha Yorumu"
            for q in quotes:
                r = tbl.add_row().cells
                r[0].text = format_ts_hms(float(q.get("start", 0.0)))
                r[1].text = str(q.get("speaker", "") or codes_str)
                r[2].text = str(q.get("text", ""))
                r[3].text = str(q.get("note", ""))
        doc.save(out_path)
        return out_path

    def export_interview_docx(self, iv_rec: dict, segments: list):
        docx_path = self.abs_path(iv_rec["docx_rel"])
        os.makedirs(os.path.dirname(docx_path), exist_ok=True)
        doc = docx.Document()
        codes_str = ", ".join(iv_rec.get("participant_codes", []))
        doc.add_heading(f"FIELD Deşifre — {codes_str} ({iv_rec['id']})", 0)

        meta_p = doc.add_paragraph()
        r_m = meta_p.add_run(
            f"Proje: {self.data.get('project_name')} | Katılımcı: {codes_str} | "
            f"Durum: {'Kontrol Edildi' if iv_rec.get('status') == 'reviewed' else 'İnceleniyor'}\n"
        )
        r_m.italic = True

        if iv_rec.get("general_note", "").strip():
            gp = doc.add_paragraph()
            rg = gp.add_run(f"📌 Genel Saha Notu: {iv_rec['general_note'].strip()}")
            rg.bold = True
            rg.font.color.rgb = RGBColor(0x92, 0x40, 0x0E)

        tbl = doc.add_table(rows=1, cols=3)
        tbl.style = "Table Grid"
        hdr = tbl.rows[0].cells
        hdr[0].text = "Zaman"
        hdr[1].text = "Konuşmacı"
        hdr[2].text = "Deşifre Metni (Alıntılar Sarı Vurgulu)"

        for s in segments:
            row = tbl.add_row().cells
            row[0].text = format_ts_hms(float(s.get("start", 0.0)))
            row[1].text = str(s.get("speaker", "") or codes_str)
            p_txt = row[2].paragraphs[0]
            run_txt = p_txt.add_run(str(s.get("text", "")))
            if s.get("is_quote"):
                run_txt.font.highlight_color = WD_COLOR_INDEX.YELLOW
                run_txt.bold = True
                if s.get("quote_theme", "").strip():
                    r_th = p_txt.add_run(f"  [Tema: {s['quote_theme'].strip()}]")
                    r_th.font.color.rgb = RGBColor(0x1D, 0x4E, 0xD8)
                    r_th.bold = True
            note_s = str(s.get("note", "")).strip()
            if note_s:
                add_word_comment_or_note(p_txt, note_s, author="Araştırmacı")

        doc.save(docx_path)
        return docx_path

    def get_all_quotes(self) -> list:
        quotes = []
        for iv in self.data.get("interviews", []):
            segs = self.load_interview_transcript(iv)
            codes_str = ", ".join(iv.get("participant_codes", []))
            for idx, s in enumerate(segs):
                if s.get("is_quote"):
                    quotes.append({
                        "interview_id": iv["id"],
                        "audio_rel": iv.get("audio_rel", ""),
                        "participant": s.get("speaker") or codes_str,
                        "participant_codes": iv.get("participant_codes", []),
                        "start": float(s.get("start", 0.0)),
                        "end": float(s.get("end", 0.0)),
                        "text": str(s.get("text", "")),
                        "theme": str(s.get("quote_theme", "")).strip() or "Genel",
                        "note": str(s.get("note", "")).strip(),
                        "seg_idx": idx,
                    })
        return quotes

    def export_all_quotes_docx(self) -> str:
        out_path = os.path.join(self.project_dir, "Alintilar_Quotes.docx")
        quotes = self.get_all_quotes()
        doc = docx.Document()
        doc.add_heading(f"FIELD Alıntılar (Quotes) — {self.data.get('project_name')}", 0)
        doc.add_paragraph(f"Toplam Alıntı: {len(quotes)} | Oluşturma: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}")

        if not quotes:
            doc.add_paragraph("Henüz alıntı olarak işaretlenmiş satır bulunmuyor.")
        else:
            tbl = doc.add_table(rows=1, cols=5)
            tbl.style = "Table Grid"
            hdr = tbl.rows[0].cells
            hdr[0].text = "Katılımcı Kodu"
            hdr[1].text = "Zaman Damgası"
            hdr[2].text = "Tema"
            hdr[3].text = "Alıntı Metni"
            hdr[4].text = "Saha Notu"
            for q in quotes:
                r = tbl.add_row().cells
                r[0].text = q["participant"]
                r[1].text = format_ts_hms(q["start"])
                r[2].text = q["theme"]
                r[3].text = q["text"]
                r[4].text = q["note"]
        doc.save(out_path)
        return out_path

    def export_participant_profiles(self):
        """Export coded participant profiles (NO real names) to Word (.docx) and Excel (.xlsx)."""
        fields = self.data.get("profile_fields", [])
        parts = self.data.get("participants", {})
        ordered_codes = sorted(parts.keys())

        # 1. Word (.docx)
        docx_path = os.path.join(self.project_dir, "Katilimci_Profilleri.docx")
        doc = docx.Document()
        doc.add_heading(f"Katılımcı Profilleri (Anonim Kodlu) — {self.data.get('project_name')}", 0)
        doc.add_paragraph("Not: Etik kurallar gereği bu dosyada yalnızca katılımcı kodları yer alır; gerçek isimler içermez.")
        cols = ["Katılımcı Kodu", "Onam Durumu"] + fields
        tbl = doc.add_table(rows=1, cols=len(cols))
        tbl.style = "Table Grid"
        for i, h in enumerate(cols):
            tbl.rows[0].cells[i].text = h
        for code in ordered_codes:
            p = parts[code]
            row = tbl.add_row().cells
            row[0].text = code
            row[1].text = p.get("consent", "Onam Alındı")
            for i, fld in enumerate(fields):
                row[2 + i].text = str(p.get("fields", {}).get(fld, ""))
        doc.save(docx_path)

        # 2. Excel (.xlsx)
        import openpyxl
        xlsx_path = os.path.join(self.project_dir, "Katilimci_Profilleri.xlsx")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Katilimci Profilleri"
        ws.append(cols)
        for code in ordered_codes:
            p = parts[code]
            ws.append([code, p.get("consent", "Onam Alındı")] + [str(p.get("fields", {}).get(fld, "")) for fld in fields])
        wb.save(xlsx_path)
        return docx_path, xlsx_path

    def search_project(self, query: str) -> list:
        q = (query or "").strip().lower()
        if not q:
            return []
        hits = []
        for iv in self.data.get("interviews", []):
            codes_str = ", ".join(iv.get("participant_codes", []))
            if q in iv.get("general_note", "").lower():
                hits.append({
                    "interview_id": iv["id"],
                    "audio_rel": iv.get("audio_rel", ""),
                    "participant": codes_str,
                    "start": 0.0,
                    "kind": "Genel Saha Notu",
                    "snippet": iv["general_note"],
                })
            for s in self.load_interview_transcript(iv):
                txt = str(s.get("text", ""))
                note = str(s.get("note", ""))
                theme = str(s.get("quote_theme", ""))
                if q in txt.lower() or q in note.lower() or q in theme.lower():
                    hits.append({
                        "interview_id": iv["id"],
                        "audio_rel": iv.get("audio_rel", ""),
                        "participant": s.get("speaker") or codes_str,
                        "start": float(s.get("start", 0.0)),
                        "kind": "Alıntı" if s.get("is_quote") else "Deşifre / Not",
                        "snippet": f"{txt}" + (f" [Not: {note}]" if note else ""),
                    })
        return hits

    def get_method_summary_stats(self) -> dict:
        ivs = self.data.get("interviews", [])
        parts = self.data.get("participants", {})
        total_dur = sum(float(iv.get("duration_sec", 0.0)) for iv in ivs)
        reviewed_cnt = sum(1 for iv in ivs if iv.get("status") == "reviewed")
        consent_counts = {}
        for p in parts.values():
            c = p.get("consent", "Onam Alındı")
            consent_counts[c] = consent_counts.get(c, 0) + 1
        return {
            "project_name": self.data.get("project_name"),
            "language": self.data.get("language", "Turkish"),
            "interview_count": len(ivs),
            "reviewed_count": reviewed_cnt,
            "participant_count": len(parts),
            "total_duration_sec": total_dur,
            "avg_duration_sec": (total_dur / len(ivs)) if ivs else 0.0,
            "quote_count": len(self.get_all_quotes()),
            "consent_counts": consent_counts,
        }

    def export_method_summary_docx(self) -> str:
        st = self.get_method_summary_stats()
        out_path = os.path.join(self.project_dir, "Yontem_Ozeti.docx")
        doc = docx.Document()
        doc.add_heading(f"Yöntem ve Veri Toplama Özeti — {st['project_name']}", 0)
        rows = [
            ("Proje Dili", st["language"]),
            ("Toplam Görüşme / Kayıt Sayısı", str(st["interview_count"])),
            ("Kontrol Edilmiş (Onaylı) Görüşme", f"{st['reviewed_count']} / {st['interview_count']}"),
            ("Toplam Katılımcı Sayısı (Kodlu)", str(st["participant_count"])),
            ("Toplam Kayıt Süresi", format_ts_hms(st["total_duration_sec"])),
            ("Ortalama Görüşme Süresi", format_ts_hms(st["avg_duration_sec"])),
            ("İşaretlenen Alıntı Sayısı", str(st["quote_count"])),
            ("Onam Durumu Dağılımı", ", ".join(f"{k}: {v}" for k, v in st["consent_counts"].items()) or "-"),
        ]
        tbl = doc.add_table(rows=1, cols=2)
        tbl.style = "Table Grid"
        tbl.rows[0].cells[0].text = "Gösterge"
        tbl.rows[0].cells[1].text = "Değer"
        for k, v in rows:
            c = tbl.add_row().cells
            c[0].text = k
            c[1].text = v
        doc.save(out_path)
        return out_path

    def export_custom_report(self, selected_quotes: list, selected_notes: list, title: str, fmt: str = "docx") -> str:
        stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_title = title.strip() or f"FIELD Rapor — {self.data.get('project_name')}"
        out_path = os.path.join(self.project_dir, f"FIELD_Rapor_{stamp}.docx")
        doc = docx.Document()
        doc.add_heading(safe_title, 0)
        doc.add_paragraph(f"Proje: {self.data.get('project_name')} | Tarih: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}")
        if selected_quotes:
            doc.add_heading("Seçilen Alıntılar", level=1)
            for q in selected_quotes:
                p = doc.add_paragraph()
                r = p.add_run(f"[{q['participant']} · {format_ts_hms(q['start'])} · {q['theme']}] ")
                r.bold = True
                p.add_run(q["text"])
        if selected_notes:
            doc.add_heading("Seçilen Saha Notları", level=1)
            for n in selected_notes:
                p = doc.add_paragraph()
                r = p.add_run(f"[{n['label']}] ")
                r.bold = True
                p.add_run(n["note"])
        doc.save(out_path)
        return out_path

    def export_nvivo_maxqda_project(self) -> str:
        out_path = os.path.join(self.project_dir, f"{self.data.get('project_name')}_NVivo_MAXQDA.txt")
        with open(out_path, "w", encoding="utf-8") as f:
            f.write("Dosya\tKatilimci\tZaman\tKonusmaci\tAlinti\tTema\tSahaNotu\tMetin\n")
            for iv in self.data.get("interviews", []):
                codes_str = ",".join(iv.get("participant_codes", []))
                for s in self.load_interview_transcript(iv):
                    line = "\t".join([
                        iv["id"],
                        codes_str,
                        format_ts_hms(float(s.get("start", 0.0))),
                        str(s.get("speaker", "")),
                        "EVET" if s.get("is_quote") else "HAYIR",
                        str(s.get("quote_theme", "")).replace("\t", " "),
                        str(s.get("note", "")).replace("\t", " ").replace("\n", " "),
                        str(s.get("text", "")).replace("\t", " ").replace("\n", " "),
                    ])
                    f.write(line + "\n")
        return out_path

    def export_portable_zip(self, dest_zip_path: str) -> str:
        with zipfile.ZipFile(dest_zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for root, _, files in os.walk(self.project_dir):
                for fn in files:
                    fp = os.path.join(root, fn)
                    arc = os.path.join(os.path.basename(self.project_dir), os.path.relpath(fp, self.project_dir))
                    zf.write(fp, arc)
        return dest_zip_path


# ---------------------------------------------------------------------------
# Interactive FIELD Project Workbench Window
# ---------------------------------------------------------------------------

class FieldProjectWorkspaceWindow(ctk.CTkToplevel):
    """Full project-based qualitative research studio for FIELD."""

    def __init__(self, master_app, project_dir: str, transcribe_fn=None, diarize_fn=None, lang_map=None):
        super().__init__(master_app)
        self.app = master_app
        self.mgr = FieldProjectManager(project_dir)
        self.transcribe_fn = transcribe_fn
        self.diarize_fn = diarize_fn
        self.lang_map = lang_map or {"Auto": "auto", "Turkish": "tr", "English": "en", "German": "de"}

        self.title(f"FIELD Proje Çalışma Alanı — {self.mgr.data.get('project_name')}")
        self.geometry("1220x820")
        self.minsize(1060, 700)
        self.configure(fg_color="#F8FAFC")

        self._active_iv_id = None
        self._block_widgets = []
        self._unlocked_ethics = None
        self._ethics_pw = None
        self._is_queue_running = False

        self._build_ui()
        self.refresh_all()

    # ---- Audio Playback Helper (Cross-Platform via FFmpeg + sounddevice) ----
    def _play_audio_slice(self, audio_rel: str, start_sec: float, duration_sec: float = 12.0):
        abs_audio = self.mgr.abs_path(audio_rel)
        if not abs_audio or not os.path.isfile(abs_audio):
            messagebox.showwarning("Ses Bulunamadı", "Bu görüşmeye ait ses dosyası bulunamadı.")
            return

        def _worker():
            try:
                import numpy as np
                import sounddevice as sd
                from scipy.io import wavfile
                import tempfile

                sd.stop()
                tmp_wav = os.path.join(tempfile.gettempdir(), "ilber_field_slice.wav")
                creationflags = 0x08000000 if os.name == "nt" else 0
                subprocess.run(
                    [
                        "ffmpeg", "-y", "-ss", str(max(0.0, float(start_sec))),
                        "-t", str(duration_sec), "-i", abs_audio,
                        "-ar", "16000", "-ac", "1", tmp_wav,
                    ],
                    capture_output=True,
                    creationflags=creationflags,
                )
                if os.path.isfile(tmp_wav):
                    fs, data = wavfile.read(tmp_wav)
                    sd.play(data, fs)
            except Exception as ex:
                self.after(0, lambda m=str(ex): messagebox.showerror("Ses Oynatma Hatası", m))

        threading.Thread(target=_worker, daemon=True).start()

    def _stop_audio(self):
        try:
            import sounddevice as sd
            sd.stop()
        except Exception:
            pass

    # ---- Main Layout ----
    def _build_ui(self):
        # Header Bar
        hdr = ctk.CTkFrame(self, fg_color="#0F172A", height=64, corner_radius=0)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)

        self.hdr_title = ctk.CTkLabel(
            hdr,
            text=f"🔬 FIELD Proje: {self.mgr.data.get('project_name')}",
            font=ctk.CTkFont(family="Segoe UI", size=18, weight="bold"),
            text_color="#FFFFFF",
        )
        self.hdr_title.pack(side="left", padx=20)

        self.hdr_meta = ctk.CTkLabel(
            hdr,
            text=f"Dil: {self.mgr.data.get('language')} · Konuşmacı: {self.mgr.data.get('speaker_count')}",
            font=ctk.CTkFont(size=12),
            text_color="#94A3B8",
        )
        self.hdr_meta.pack(side="left", padx=10)

        ctk.CTkButton(
            hdr, text="⏹ Sesi Durdur", width=110, height=32, fg_color="#DC2626", hover_color="#B91C1C",
            font=ctk.CTkFont(size=11, weight="bold"), command=self._stop_audio
        ).pack(side="right", padx=(6, 16))

        ctk.CTkButton(
            hdr, text="📦 Taşınabilir Paket (.zip)", width=165, height=32, fg_color="#2563EB", hover_color="#1D4ED8",
            font=ctk.CTkFont(size=11, weight="bold"), command=self._on_export_zip
        ).pack(side="right", padx=6)

        ctk.CTkButton(
            hdr, text="📂 Klasörü Aç", width=110, height=32, fg_color="#334155", hover_color="#475569",
            font=ctk.CTkFont(size=11, weight="bold"), command=lambda: open_path_crossplatform(self.mgr.project_dir)
        ).pack(side="right", padx=6)

        # Clean Spacious Navigation Bar (replacing cramped default tab bar)
        nav_bar = ctk.CTkFrame(self, fg_color="#FFFFFF", height=52, corner_radius=0, border_width=1, border_color="#E2E8F0")
        nav_bar.pack(fill="x")
        nav_bar.pack_propagate(False)

        self._tab_frames = {}
        self._tab_btns = {}
        tab_defs = [
            ("queue", "1. Görüşmeler & Kuyruk"),
            ("studio", "2. Deşifre Stüdyosu"),
            ("quotes", "3. Alıntılar"),
            ("profiles", "4. Katılımcı Profilleri"),
            ("search", "5. Arama & Rapor"),
            ("ethics", "6. Etik Anahtar"),
        ]
        self._tab_container = ctk.CTkFrame(self, fg_color="#F8FAFC", corner_radius=0)
        self._tab_container.pack(fill="both", expand=True, padx=14, pady=10)

        for key, lbl in tab_defs:
            frame = ctk.CTkFrame(self._tab_container, fg_color="#F8FAFC", corner_radius=0)
            self._tab_frames[key] = frame
            btn = ctk.CTkButton(
                nav_bar,
                text=lbl,
                height=34,
                corner_radius=10,
                border_width=1,
                border_color="#CBD5E1",
                font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
                command=lambda k=key: self._switch_tab(k),
            )
            btn.pack(side="left", padx=6, pady=9)
            self._tab_btns[key] = btn

        self.t_queue = self._tab_frames["queue"]
        self.t_studio = self._tab_frames["studio"]
        self.t_quotes = self._tab_frames["quotes"]
        self.t_profiles = self._tab_frames["profiles"]
        self.t_search = self._tab_frames["search"]
        self.t_ethics = self._tab_frames["ethics"]

        self._build_tab_queue()
        self._build_tab_studio()
        self._build_tab_quotes()
        self._build_tab_profiles()
        self._build_tab_search_reports()
        self._build_tab_ethics()
        self._switch_tab("queue")

    def _switch_tab(self, key: str):
        for k, fr in self._tab_frames.items():
            fr.pack_forget()
            b = self._tab_btns[k]
            if k == key:
                b.configure(fg_color="#2563EB", text_color="#FFFFFF", border_color="#2563EB", hover_color="#1D4ED8")
            else:
                b.configure(fg_color="#F1F5F9", text_color="#334155", border_color="#CBD5E1", hover_color="#E2E8F0")
        self._tab_frames[key].pack(fill="both", expand=True)

    def refresh_all(self):
        self.hdr_title.configure(text=f"🔬 FIELD Proje: {self.mgr.data.get('project_name')}")
        self.hdr_meta.configure(text=f"Dil: {self.mgr.data.get('language')} · Konuşmacı: {self.mgr.data.get('speaker_count')}")
        self._refresh_queue_list()
        self._refresh_studio_selector()
        self._refresh_quotes_list()
        self._refresh_profiles_table()
        self._refresh_method_summary()

    # =========================================================================
    # TAB 1: GÖRÜŞMELER & KUYRUK
    # =========================================================================
    def _build_tab_queue(self):
        guide = ctk.CTkFrame(self.t_queue, fg_color="#EFF6FF", corner_radius=10, border_width=1, border_color="#BFDBFE")
        guide.pack(fill="x", padx=8, pady=(4, 2))
        ctk.CTkLabel(
            guide,
            text="Adım 1: Ses/video ekle   ➔   Adım 2: 'Kuyruğu İşle'ye bas   ➔   Adım 3: 'Stüdyoda Aç' ile alıntı & not işaretleyip onayla",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#1E3A8A",
        ).pack(anchor="w", padx=12, pady=7)

        top = ctk.CTkFrame(self.t_queue, fg_color="#FFFFFF", corner_radius=12, border_width=1, border_color="#E2E8F0")
        top.pack(fill="x", padx=8, pady=6)

        r1 = ctk.CTkFrame(top, fg_color="transparent")
        r1.pack(fill="x", padx=12, pady=(8, 4))

        ctk.CTkLabel(r1, text="Proje Dili:", font=ctk.CTkFont(size=12, weight="bold")).pack(side="left", padx=(0, 6))
        self.q_lang_var = ctk.StringVar(value=self.mgr.data.get("language", "Auto"))
        ctk.CTkOptionMenu(
            r1, variable=self.q_lang_var, values=list(self.lang_map.keys()), width=130, height=30,
            command=self._on_project_settings_changed
        ).pack(side="left", padx=(0, 14))

        ctk.CTkLabel(r1, text="Konuşmacı Sayısı:", font=ctk.CTkFont(size=12, weight="bold")).pack(side="left", padx=(0, 6))
        self.q_spk_var = ctk.StringVar(value=str(self.mgr.data.get("speaker_count", 2)))
        ctk.CTkOptionMenu(
            r1, variable=self.q_spk_var, values=["1", "2", "3", "4", "5", "6", "8", "10"], width=70, height=30,
            command=self._on_project_settings_changed
        ).pack(side="left", padx=(0, 14))

        self.btn_run_queue = ctk.CTkButton(
            r1, text="▶ Kuyruğu İşle", width=140, height=32, fg_color="#7C3AED", hover_color="#6D28D9",
            font=ctk.CTkFont(size=12, weight="bold"), command=self._on_run_queue
        )
        self.btn_run_queue.pack(side="right", padx=(6, 0))

        ctk.CTkButton(
            r1, text="🧹 Kuyruğu Temizle", width=145, height=32, fg_color="#E2E8F0", hover_color="#CBD5E1", text_color="#334155",
            font=ctk.CTkFont(size=11, weight="bold"), command=self._on_clear_queue
        ).pack(side="right", padx=6)

        r2 = ctk.CTkFrame(top, fg_color="transparent")
        r2.pack(fill="x", padx=12, pady=(2, 8))

        ctk.CTkButton(
            r2, text="➕ Görüşme Kaydı Ekle (Tek / Toplu)", height=32, fg_color="#0F766E", hover_color="#115E59",
            font=ctk.CTkFont(size=12, weight="bold"), command=self._on_add_interviews
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            r2, text="📥 Hazır Word/TXT Döküm Ekle", height=32, fg_color="#2563EB", hover_color="#1D4ED8",
            font=ctk.CTkFont(size=12, weight="bold"), command=self._on_import_external_doc
        ).pack(side="left", padx=4)

        self.q_status_lbl = ctk.CTkLabel(self.t_queue, text="", font=ctk.CTkFont(size=12, weight="bold"), text_color="#2563EB")
        self.q_status_lbl.pack(anchor="w", padx=12, pady=(2, 4))

        self.q_scroll = ctk.CTkScrollableFrame(self.t_queue, fg_color="#FFFFFF", corner_radius=12, border_width=1, border_color="#E2E8F0")
        self.q_scroll.pack(fill="both", expand=True, padx=8, pady=4)

    def _on_project_settings_changed(self, *args):
        self.mgr.data["language"] = self.q_lang_var.get()
        try:
            self.mgr.data["speaker_count"] = int(self.q_spk_var.get())
        except Exception:
            pass
        self.mgr.save()
        self.hdr_meta.configure(text=f"Dil: {self.mgr.data.get('language')} · Konuşmacı: {self.mgr.data.get('speaker_count')}")

    def _on_add_interviews(self):
        files = filedialog.askopenfilenames(
            title="Görüşme Ses/Video Kayıtlarını Seçin",
            filetypes=[("Medya Dosyaları", "*.mp3 *.wav *.m4a *.ogg *.flac *.mp4 *.mov *.mkv *.webm"), ("Tüm Dosyalar", "*.*")],
        )
        if not files:
            return
        files = list(files)
        # Ask participant codes (auto-increment for 1-on-1, or comma-separated for focus group)
        suggested = self.mgr.next_participant_code()
        dlg = ctk.CTkInputDialog(
            text=(
                f"{len(files)} dosya seçildi.\n"
                f"Katılımcı kodunu yazın (örn: {suggested} — toplu yüklemede otomatik artar;\n"
                f"Odak grup ise virgülle ayırın: K01, K02, K03):"
            ),
            title="Katılımcı Kodu / Odak Grup",
        )
        raw_code = dlg.get_input()
        if raw_code is None:
            return
        raw_code = raw_code.strip() or suggested
        spk_cnt = int(self.q_spk_var.get())

        if "," in raw_code:
            codes = [c.strip().upper() for c in raw_code.split(",") if c.strip()]
            for fp in files:
                self.mgr.add_interview_record(fp, codes, max(spk_cnt, len(codes)), is_focus_group=True)
        else:
            m = re.match(r"^([A-Za-z]+)(\d+)$", raw_code)
            prefix = m.group(1).upper() if m else "K"
            start_num = int(m.group(2)) if m else 1
            for idx, fp in enumerate(files):
                code = f"{prefix}{start_num + idx:02d}" if (m or len(files) > 1) else raw_code.upper()
                self.mgr.add_interview_record(fp, [code], spk_cnt, is_focus_group=False)

        self.refresh_all()

    def _on_import_external_doc(self):
        doc_fp = filedialog.askopenfilename(
            title="Başka Yöntemle Yazılmış Döküm Seçin (.docx / .txt)",
            filetypes=[("Word / Metin", "*.docx *.txt")],
        )
        if not doc_fp:
            return
        audio_fp = ""
        if messagebox.askyesno("Ses Kaydı Bağla", "Bu döküme ait bir ses/video kaydını da proje klasörüne kopyalamak ister misiniz?"):
            audio_fp = filedialog.askopenfilename(
                title="İlgili Ses/Video Kaydını Seçin",
                filetypes=[("Medya", "*.mp3 *.wav *.m4a *.ogg *.flac *.mp4 *.mov *.mkv *.webm")],
            ) or ""
        suggested = self.mgr.next_participant_code()
        dlg = ctk.CTkInputDialog(text=f"Katılımcı kodu (örn: {suggested} veya K01, K02):", title="Katılımcı Kodu")
        raw_c = (dlg.get_input() or suggested).strip()
        codes = [c.strip().upper() for c in raw_c.split(",") if c.strip()]
        self.mgr.import_external_transcript(doc_fp, codes, optional_audio_path=audio_fp)
        self.refresh_all()

    def _on_run_queue(self):
        if self._is_queue_running:
            return
        queued = [iv for iv in self.mgr.data.get("interviews", []) if iv.get("status") == "queued"]
        if not queued:
            messagebox.showinfo("Kuyruk Boş", "Kuyrukta bekleyen (işlenmemiş) görüşme bulunmuyor.")
            return
        if not self.transcribe_fn or not self.diarize_fn:
            messagebox.showerror("Motor Hatası", "Ses çözümleme motoru bağlı değil.")
            return

        self._is_queue_running = True
        self.btn_run_queue.configure(state="disabled", text="⏳ Kuyruk İşleniyor...")

        def _worker():
            try:
                lang_sel = self.mgr.data.get("language", "Turkish")
                for idx, iv in enumerate(queued, 1):
                    codes_str = ", ".join(iv.get("participant_codes", []))
                    self.after(0, lambda i=idx, n=len(queued), c=codes_str: self.q_status_lbl.configure(
                        text=f"⏳ Kuyruk İşleniyor [{i}/{n}]: {c} — Deşifre & Konuşmacı Ayrımı..."
                    ))
                    audio_abs = self.mgr.abs_path(iv["audio_rel"])
                    raw_segs, _det = self.transcribe_fn(audio_abs, lang_sel)
                    spk_n = int(iv.get("speaker_count", 2))
                    if iv.get("is_focus_group") and len(iv.get("participant_codes", [])) >= 2:
                        spk_labels = list(iv["participant_codes"])[:spk_n]
                    else:
                        spk_labels = ["Görüşmeci"] + list(iv.get("participant_codes", ["K01"])) + [f"Konuşmacı {k}" for k in range(3, 12)]
                        spk_labels = spk_labels[:spk_n]
                    diarized = self.diarize_fn(audio_abs, raw_segs, spk_n, spk_labels)
                    blocks = []
                    for s in diarized:
                        txt = str(s.get("text", "")).strip()
                        # Flag very short or uncertain fragments as low confidence hint
                        low_words = list(s.get("low_conf_words") or [])
                        blocks.append({
                            "start": float(s.get("start", 0.0)),
                            "end": float(s.get("end", 0.0)),
                            "speaker": str(s.get("speaker", spk_labels[0] if spk_labels else "")),
                            "text": txt,
                            "low_conf_words": low_words,
                            "is_quote": False,
                            "quote_theme": "",
                            "note": "",
                        })
                    iv["status"] = "ready"
                    self.mgr.save_interview_transcript(iv, blocks, general_note=iv.get("general_note", ""))
                    self.after(0, self.refresh_all)
                self.after(0, lambda: self.q_status_lbl.configure(text="✅ Kuyruktaki tüm görüşmeler işlendi!"))
            except Exception as ex:
                self.after(0, lambda m=str(ex): self.q_status_lbl.configure(text=f"❌ Hata: {m}"))
            finally:
                self._is_queue_running = False
                self.after(0, lambda: self.btn_run_queue.configure(state="normal", text="▶ Kuyruğu İşle"))

        threading.Thread(target=_worker, daemon=True).start()

    def _refresh_queue_list(self):
        for w in self.q_scroll.winfo_children():
            w.destroy()
        ivs = self.mgr.data.get("interviews", [])
        if not ivs:
            ctk.CTkLabel(self.q_scroll, text="Henüz görüşme eklenmedi. Yukarıdan kayıt veya Word/TXT döküm ekleyin.", text_color="#64748B").pack(pady=40)
            return
        status_map = {
            "queued": ("⏳ Kuyrukta", "#FEF3C7", "#92400E"),
            "ready": ("📝 İnceleniyor", "#E0F2FE", "#0369A1"),
            "reviewed": ("✅ Kontrol Edildi", "#D1FAE5", "#065F46"),
        }
        for iv in ivs:
            row = ctk.CTkFrame(self.q_scroll, fg_color="#F8FAFC", corner_radius=10, border_width=1, border_color="#E2E8F0")
            row.pack(fill="x", padx=8, pady=4)
            st_lbl, st_bg, st_fg = status_map.get(iv.get("status", "queued"), ("📝", "#F1F5F9", "#334155"))
            codes_str = ", ".join(iv.get("participant_codes", []))
            fg_tag = " [Odak Grup]" if iv.get("is_focus_group") else ""
            ctk.CTkLabel(row, text=f" {codes_str}{fg_tag} ", fg_color="#EDE9FE", text_color="#5B21B6", corner_radius=6, font=ctk.CTkFont(size=12, weight="bold")).pack(side="left", padx=10, pady=10)
            ctk.CTkLabel(row, text=iv["id"], font=ctk.CTkFont(size=12, weight="bold"), text_color="#0F172A").pack(side="left", padx=6)
            dur_s = format_ts_hms(float(iv.get("duration_sec", 0.0)))
            ctk.CTkLabel(row, text=f"⏱ {dur_s}", font=ctk.CTkFont(size=11), text_color="#64748B").pack(side="left", padx=10)
            ctk.CTkLabel(row, text=f" {st_lbl} ", fg_color=st_bg, text_color=st_fg, corner_radius=6, font=ctk.CTkFont(size=11, weight="bold")).pack(side="left", padx=6)

            ctk.CTkButton(
                row, text="🗑️ Sil", width=65, height=28, fg_color="#FEE2E2", hover_color="#FECACA", text_color="#991B1B",
                font=ctk.CTkFont(size=11, weight="bold"),
                command=lambda iv_id=iv["id"]: self._on_delete_interview(iv_id)
            ).pack(side="right", padx=(4, 10), pady=8)

            ctk.CTkButton(
                row, text="📂 Klasör", width=80, height=28, fg_color="#E2E8F0", hover_color="#CBD5E1", text_color="#1E293B",
                font=ctk.CTkFont(size=11, weight="bold"),
                command=lambda r=iv: open_path_crossplatform(self.mgr.abs_path(r["folder_rel"]))
            ).pack(side="right", padx=4, pady=8)

            ctk.CTkButton(
                row, text="📘 Word", width=80, height=28, fg_color="#EFF6FF", hover_color="#DBEAFE", text_color="#1D4ED8",
                font=ctk.CTkFont(size=11, weight="bold"),
                command=lambda r=iv: open_path_crossplatform(self.mgr.abs_path(r["docx_rel"]))
            ).pack(side="right", padx=4, pady=8)

            ctk.CTkButton(
                row, text="🎙️ Stüdyoda Aç", width=115, height=28, fg_color="#2563EB", hover_color="#1D4ED8",
                font=ctk.CTkFont(size=11, weight="bold"),
                command=lambda iv_id=iv["id"]: self._open_in_studio(iv_id)
            ).pack(side="right", padx=4, pady=8)

    def _on_delete_interview(self, iv_id: str):
        if messagebox.askyesno("Görüşmeyi Sil", f"'{iv_id}' görüşmesi ve klasörü silinsin mi?"):
            self.mgr.delete_interview_record(iv_id)
            self.refresh_all()

    def _on_clear_queue(self):
        q_cnt = sum(1 for iv in self.mgr.data.get("interviews", []) if iv.get("status") == "queued")
        if not q_cnt:
            messagebox.showinfo("Kuyruk Temiz", "Kuyrukta bekleyen kayıt yok.")
            return
        if messagebox.askyesno("Kuyruğu Temizle", f"Kuyruktaki {q_cnt} adet işlenmemiş kayıt silinsin mi?"):
            self.mgr.clear_queued_interviews()
            self.refresh_all()

    def _open_in_studio(self, iv_id: str):
        self._active_iv_id = iv_id
        self.st_iv_var.set(iv_id)
        self._load_studio_blocks()
        self._switch_tab("studio")

    # =========================================================================
    # TAB 2: DEŞİFRE & İNCELEME STÜDYOSU (BLOKLAR + SOLDA SES + ALINTI + NOT)
    # =========================================================================
    def _build_tab_studio(self):
        bar = ctk.CTkFrame(self.t_studio, fg_color="#FFFFFF", corner_radius=12, border_width=1, border_color="#E2E8F0")
        bar.pack(fill="x", padx=8, pady=6)

        ctk.CTkLabel(bar, text="Görüşme:", font=ctk.CTkFont(size=12, weight="bold")).pack(side="left", padx=(10, 6), pady=8)
        self.st_iv_var = ctk.StringVar(value="")
        self.st_iv_menu = ctk.CTkOptionMenu(bar, variable=self.st_iv_var, values=[""], width=240, height=30, command=lambda _: self._load_studio_blocks())
        self.st_iv_menu.pack(side="left", padx=4, pady=8)

        ctk.CTkButton(
            bar, text="🕵️ Anonimleştirme Öner", height=30, fg_color="#0F766E", hover_color="#115E59",
            font=ctk.CTkFont(size=11, weight="bold"), command=self._on_suggest_anonymization
        ).pack(side="left", padx=6)

        ctk.CTkButton(
            bar, text="💾 Kaydet & Word Güncelle", height=30, fg_color="#2563EB", hover_color="#1D4ED8",
            font=ctk.CTkFont(size=11, weight="bold"), command=lambda: self._save_studio_blocks(mark_reviewed=False)
        ).pack(side="right", padx=(4, 10))

        ctk.CTkButton(
            bar, text="✅ Kontrol Edildi Onayla", height=30, fg_color="#059669", hover_color="#047857",
            font=ctk.CTkFont(size=11, weight="bold"), command=lambda: self._save_studio_blocks(mark_reviewed=True)
        ).pack(side="right", padx=4)

        # General File Field Note
        gn_row = ctk.CTkFrame(self.t_studio, fg_color="#FFFBEB", corner_radius=10, border_width=1, border_color="#FDE68A")
        gn_row.pack(fill="x", padx=8, pady=(0, 6))
        ctk.CTkLabel(gn_row, text="📌 Dosya Genel Saha Notu:", font=ctk.CTkFont(size=12, weight="bold"), text_color="#92400E").pack(side="left", padx=10, pady=6)
        self.st_general_note = ctk.CTkEntry(gn_row, height=30, fg_color="#FFFFFF", placeholder_text="Görüşmenin tamamı hakkında genel gözlem ve saha notu...")
        self.st_general_note.pack(side="left", fill="x", expand=True, padx=(0, 10), pady=6)

        self.st_scroll = ctk.CTkScrollableFrame(self.t_studio, fg_color="#FFFFFF", corner_radius=12, border_width=1, border_color="#E2E8F0")
        self.st_scroll.pack(fill="both", expand=True, padx=8, pady=4)

    def _refresh_studio_selector(self):
        ids = [iv["id"] for iv in self.mgr.data.get("interviews", [])]
        if not ids:
            self.st_iv_menu.configure(values=["(Görüşme Yok)"])
            self.st_iv_var.set("(Görüşme Yok)")
            return
        self.st_iv_menu.configure(values=ids)
        if self._active_iv_id not in ids:
            self._active_iv_id = ids[0]
        self.st_iv_var.set(self._active_iv_id)
        self._load_studio_blocks()

    def _get_active_interview(self):
        iv_id = self.st_iv_var.get()
        for iv in self.mgr.data.get("interviews", []):
            if iv["id"] == iv_id:
                return iv
        return None

    def _load_studio_blocks(self):
        for w in self.st_scroll.winfo_children():
            w.destroy()
        self._block_widgets = []
        iv = self._get_active_interview()
        if not iv:
            ctk.CTkLabel(self.st_scroll, text="İncelemek için önce bir görüşme ekleyin veya kuyruğu çalıştırın.", text_color="#64748B").pack(pady=40)
            return
        self._active_iv_id = iv["id"]
        self.st_general_note.delete(0, "end")
        self.st_general_note.insert(0, iv.get("general_note", ""))

        segs = self.mgr.load_interview_transcript(iv)
        if not segs:
            ctk.CTkLabel(self.st_scroll, text="Bu görüşme henüz deşifre edilmedi (Durum: Kuyrukta).", text_color="#94A3B8").pack(pady=40)
            return

        for idx, s in enumerate(segs):
            blk = ctk.CTkFrame(self.st_scroll, fg_color="#F8FAFC", corner_radius=10, border_width=1, border_color="#E2E8F0")
            blk.pack(fill="x", padx=6, pady=4)

            # Left column: Audio Play + Timestamp + Speaker
            left = ctk.CTkFrame(blk, fg_color="transparent", width=155)
            left.pack(side="left", fill="y", padx=8, pady=6)
            left.pack_propagate(False)

            audio_rel = iv.get("audio_rel", "")
            ctk.CTkButton(
                left, text=f"▶ {format_ts_hms(float(s.get('start', 0.0)))}", width=135, height=28,
                fg_color="#0F766E" if audio_rel else "#94A3B8", hover_color="#115E59",
                font=ctk.CTkFont(size=11, weight="bold"),
                command=lambda ar=audio_rel, st=float(s.get("start", 0.0)): self._play_audio_slice(ar, st)
            ).pack(pady=(2, 4))

            spk_e = ctk.CTkEntry(left, width=135, height=26, fg_color="#FFFFFF", font=ctk.CTkFont(size=11, weight="bold"))
            spk_e.insert(0, str(s.get("speaker", "")))
            spk_e.pack(pady=2)

            # Center column: Editable Text + Low-Confidence Hint
            mid = ctk.CTkFrame(blk, fg_color="transparent")
            mid.pack(side="left", fill="both", expand=True, padx=6, pady=6)
            txt_box = ctk.CTkTextbox(mid, height=58, wrap="word", fg_color="#FFFFFF", border_width=1, border_color="#CBD5E1", font=ctk.CTkFont(size=12))
            txt_box.insert("1.0", str(s.get("text", "")))
            txt_box.pack(fill="x", expand=True)
            if s.get("low_conf_words"):
                ctk.CTkLabel(
                    mid,
                    text=f"⚠️ Düşük güvenli kelimeler: {', '.join(s['low_conf_words'])}",
                    font=ctk.CTkFont(size=10, weight="bold"),
                    text_color="#B45309",
                ).pack(anchor="w", pady=(2, 0))

            # Right column: Quote toggle + Theme + Line Field Note
            right = ctk.CTkFrame(blk, fg_color="transparent", width=260)
            right.pack(side="right", fill="y", padx=8, pady=6)
            right.pack_propagate(False)

            q_var = ctk.BooleanVar(value=bool(s.get("is_quote", False)))
            top_r = ctk.CTkFrame(right, fg_color="transparent")
            top_r.pack(fill="x")
            ctk.CTkCheckBox(top_r, text="⭐ Alıntı", variable=q_var, width=85, font=ctk.CTkFont(size=11, weight="bold"), text_color="#B45309").pack(side="left")
            theme_e = ctk.CTkEntry(top_r, height=24, width=145, placeholder_text="Tema (örn: Göç)", fg_color="#FFFFFF", font=ctk.CTkFont(size=11))
            theme_e.insert(0, str(s.get("quote_theme", "")))
            theme_e.pack(side="right")

            note_e = ctk.CTkEntry(right, height=28, placeholder_text="Satır saha notu...", fg_color="#FFFBEB", font=ctk.CTkFont(size=11))
            note_e.insert(0, str(s.get("note", "")))
            note_e.pack(fill="x", pady=(6, 0))

            self._block_widgets.append({
                "base": s,
                "speaker": spk_e,
                "text": txt_box,
                "is_quote": q_var,
                "theme": theme_e,
                "note": note_e,
            })

    def _collect_studio_segments(self):
        updated = []
        for bw in self._block_widgets:
            s = dict(bw["base"])
            s["speaker"] = bw["speaker"].get().strip()
            s["text"] = bw["text"].get("1.0", "end-1c").strip()
            s["is_quote"] = bool(bw["is_quote"].get())
            s["quote_theme"] = bw["theme"].get().strip()
            s["note"] = bw["note"].get().strip()
            updated.append(s)
        return updated

    def _save_studio_blocks(self, mark_reviewed=False):
        iv = self._get_active_interview()
        if not iv or not self._block_widgets:
            return
        segs = self._collect_studio_segments()
        if mark_reviewed:
            iv["status"] = "reviewed"
        self.mgr.save_interview_transcript(iv, segs, general_note=self.st_general_note.get().strip())
        self.refresh_all()
        msg = "Görüşme 'Kontrol Edildi' olarak onaylandı ve Word/Alıntı dosyaları güncellendi." if mark_reviewed else "Deşifre, alıntılar ve saha notları kaydedildi."
        messagebox.showinfo("Kaydedildi", msg)

    def _on_suggest_anonymization(self):
        if not self._block_widgets:
            return
        segs = self._collect_studio_segments()
        cands = detect_anonymization_candidates(segs)
        if not cands:
            messagebox.showinfo("Anonimleştirme", "Metinde kişi veya yer adı adayı tespit edilmedi.")
            return

        pop = ctk.CTkToplevel(self)
        pop.title("🕵️ Anonimleştirme Önerileri (Kişi / Yer)")
        pop.geometry("620x480")
        pop.attributes("-topmost", True)
        ctk.CTkLabel(pop, text="Aşağıdaki kişi, yer ve iletişim bilgilerini kodlamak istediklerinizi seçin:", font=ctk.CTkFont(size=13, weight="bold")).pack(padx=16, pady=10, anchor="w")

        sc = ctk.CTkScrollableFrame(pop, fg_color="#FFFFFF")
        sc.pack(fill="both", expand=True, padx=16, pady=6)
        vars_list = []
        for c in cands:
            r = ctk.CTkFrame(sc, fg_color="#F8FAFC")
            r.pack(fill="x", pady=3, padx=4)
            v = ctk.BooleanVar(value=True)
            ctk.CTkCheckBox(r, text=f"{c['original']}  ➔  {c['replacement']}  ({c['category']} · {c['count']}x)", variable=v, font=ctk.CTkFont(size=12)).pack(side="left", padx=8, pady=6)
            vars_list.append((v, c))

        def _apply():
            repls = [(c["original"], c["replacement"]) for v, c in vars_list if v.get()]
            for bw in self._block_widgets:
                t = bw["text"].get("1.0", "end-1c")
                for orig, rep in repls:
                    t = re.sub(rf"\b{re.escape(orig)}\b", rep, t)
                bw["text"].delete("1.0", "end")
                bw["text"].insert("1.0", t)
            pop.destroy()
            messagebox.showinfo("Uygulandı", f"{len(repls)} ifade metinde anonimleştirildi. Kalıcı olması için 'Kaydet' butonuna basın.")

        ctk.CTkButton(pop, text="✔ Seçilenleri Metne Uygula", height=36, fg_color="#059669", command=_apply).pack(pady=12)

    # =========================================================================
    # TAB 3: ALINTILAR (QUOTES + AUDIO JUMP + DOCX)
    # =========================================================================
    def _build_tab_quotes(self):
        top = ctk.CTkFrame(self.t_quotes, fg_color="#FFFFFF", corner_radius=12, border_width=1, border_color="#E2E8F0")
        top.pack(fill="x", padx=8, pady=6)
        ctk.CTkLabel(top, text="Tema Filtresi:", font=ctk.CTkFont(size=12, weight="bold")).pack(side="left", padx=(12, 6), pady=8)
        self.qt_filter_entry = ctk.CTkEntry(top, width=220, height=30, placeholder_text="Tema veya katılımcı ara...")
        self.qt_filter_entry.pack(side="left", padx=4, pady=8)
        self.qt_filter_entry.bind("<KeyRelease>", lambda _: self._refresh_quotes_list())

        ctk.CTkButton(
            top, text="📘 Alıntılar Word (.docx) Oluştur & Aç", height=32, fg_color="#2563EB", hover_color="#1D4ED8",
            font=ctk.CTkFont(size=12, weight="bold"),
            command=lambda: open_path_crossplatform(self.mgr.export_all_quotes_docx())
        ).pack(side="right", padx=12, pady=8)

        self.qt_scroll = ctk.CTkScrollableFrame(self.t_quotes, fg_color="#FFFFFF", corner_radius=12, border_width=1, border_color="#E2E8F0")
        self.qt_scroll.pack(fill="both", expand=True, padx=8, pady=4)

    def _refresh_quotes_list(self):
        for w in self.qt_scroll.winfo_children():
            w.destroy()
        q_filter = self.qt_filter_entry.get().strip().lower()
        quotes = self.mgr.get_all_quotes()
        if q_filter:
            quotes = [q for q in quotes if q_filter in q["theme"].lower() or q_filter in q["participant"].lower() or q_filter in q["text"].lower()]
        if not quotes:
            ctk.CTkLabel(self.qt_scroll, text="Henüz alıntı işaretlenmedi. Deşifre Stüdyosu'ndan satırları ⭐ Alıntı olarak işaretleyin.", text_color="#64748B").pack(pady=40)
            return
        for q in quotes:
            row = ctk.CTkFrame(self.qt_scroll, fg_color="#FFFBEB", corner_radius=10, border_width=1, border_color="#FDE68A")
            row.pack(fill="x", padx=6, pady=4)
            hdr = ctk.CTkFrame(row, fg_color="transparent")
            hdr.pack(fill="x", padx=10, pady=(6, 2))
            ctk.CTkLabel(hdr, text=f" {q['participant']} ", fg_color="#EDE9FE", text_color="#5B21B6", corner_radius=6, font=ctk.CTkFont(size=11, weight="bold")).pack(side="left", padx=(0, 6))
            ctk.CTkLabel(hdr, text=f"🏷️ {q['theme']}", fg_color="#DBEAFE", text_color="#1E40AF", corner_radius=6, font=ctk.CTkFont(size=11, weight="bold")).pack(side="left", padx=4)
            ctk.CTkButton(
                hdr, text=f"▶ {format_ts_hms(q['start'])} Dinle", width=115, height=26, fg_color="#0F766E", hover_color="#115E59",
                font=ctk.CTkFont(size=11, weight="bold"),
                command=lambda ar=q["audio_rel"], st=q["start"]: self._play_audio_slice(ar, st)
            ).pack(side="right")
            ctk.CTkLabel(row, text=q["text"], font=ctk.CTkFont(size=12), text_color="#0F172A", wraplength=1050, justify="left").pack(anchor="w", padx=12, pady=(2, 6))

    # =========================================================================
    # TAB 4: KATILIMCI PROFİLLERİ (KODLU + ÖZEL ALANLAR + ONAM + WORD/EXCEL)
    # =========================================================================
    def _build_tab_profiles(self):
        top = ctk.CTkFrame(self.t_profiles, fg_color="#FFFFFF", corner_radius=12, border_width=1, border_color="#E2E8F0")
        top.pack(fill="x", padx=8, pady=6)

        ctk.CTkButton(top, text="+ Katılımcı Ekle", width=120, height=30, fg_color="#0F766E", command=self._on_add_participant).pack(side="left", padx=8, pady=8)
        ctk.CTkButton(top, text="+ Alan (Sütun) Ekle", width=130, height=30, fg_color="#2563EB", command=self._on_add_profile_field).pack(side="left", padx=4, pady=8)
        ctk.CTkButton(top, text="- Alan Sil", width=90, height=30, fg_color="#64748B", command=self._on_remove_profile_field).pack(side="left", padx=4, pady=8)

        ctk.CTkButton(
            top, text="📊 Excel (.xlsx) Çıkar", width=140, height=30, fg_color="#059669",
            font=ctk.CTkFont(size=11, weight="bold"), command=lambda: self._on_export_profiles("xlsx")
        ).pack(side="right", padx=8, pady=8)
        ctk.CTkButton(
            top, text="📘 Word (.docx) Çıkar", width=140, height=30, fg_color="#1D4ED8",
            font=ctk.CTkFont(size=11, weight="bold"), command=lambda: self._on_export_profiles("docx")
        ).pack(side="right", padx=4, pady=8)
        ctk.CTkButton(
            top, text="💾 Profilleri Kaydet", width=130, height=30, fg_color="#7C3AED",
            font=ctk.CTkFont(size=11, weight="bold"), command=self._save_profiles_from_ui
        ).pack(side="right", padx=4, pady=8)

        self.prof_scroll = ctk.CTkScrollableFrame(self.t_profiles, fg_color="#FFFFFF", corner_radius=12, border_width=1, border_color="#E2E8F0")
        self.prof_scroll.pack(fill="both", expand=True, padx=8, pady=4)
        self._prof_entries = {}

    def _refresh_profiles_table(self):
        for w in self.prof_scroll.winfo_children():
            w.destroy()
        self._prof_entries = {}
        fields = self.mgr.data.get("profile_fields", [])
        parts = self.mgr.data.get("participants", {})
        if not parts:
            ctk.CTkLabel(self.prof_scroll, text="Henüz katılımcı yok. Görüşme eklediğinizde veya '+ Katılımcı Ekle' dediğinizde burada kodlu olarak açılır.", text_color="#64748B").pack(pady=40)
            return

        hdr = ctk.CTkFrame(self.prof_scroll, fg_color="#F1F5F9")
        hdr.pack(fill="x", padx=6, pady=4)
        ctk.CTkLabel(hdr, text="Kod", width=70, font=ctk.CTkFont(size=12, weight="bold")).pack(side="left", padx=6, pady=6)
        ctk.CTkLabel(hdr, text="Onam Durumu", width=140, font=ctk.CTkFont(size=12, weight="bold")).pack(side="left", padx=6)
        for fld in fields:
            ctk.CTkLabel(hdr, text=fld, width=130, font=ctk.CTkFont(size=12, weight="bold")).pack(side="left", padx=6)

        for code in sorted(parts.keys()):
            p = parts[code]
            row = ctk.CTkFrame(self.prof_scroll, fg_color="#F8FAFC")
            row.pack(fill="x", padx=6, pady=3)
            ctk.CTkLabel(row, text=code, width=70, font=ctk.CTkFont(size=12, weight="bold"), text_color="#5B21B6").pack(side="left", padx=6, pady=6)
            c_var = ctk.StringVar(value=p.get("consent", "Onam Alındı"))
            ctk.CTkOptionMenu(row, variable=c_var, values=["Onam Alındı", "Sözlü Onam", "Bekliyor", "Geri Çekildi"], width=140, height=28).pack(side="left", padx=6)
            fld_map = {}
            for fld in fields:
                e = ctk.CTkEntry(row, width=130, height=28, fg_color="#FFFFFF")
                e.insert(0, str(p.get("fields", {}).get(fld, "")))
                e.pack(side="left", padx=6)
                fld_map[fld] = e
            self._prof_entries[code] = (c_var, fld_map)

    def _save_profiles_from_ui(self):
        for code, (c_var, fld_map) in self._prof_entries.items():
            self.mgr.ensure_participant(code)
            self.mgr.data["participants"][code]["consent"] = c_var.get()
            self.mgr.data["participants"][code]["fields"] = {k: e.get().strip() for k, e in fld_map.items()}
        self.mgr.save()
        self._refresh_method_summary()
        messagebox.showinfo("Kaydedildi", "Katılımcı profilleri kaydedildi (yalnızca kodlu, gerçek isim içermez).")

    def _on_add_participant(self):
        code = self.mgr.next_participant_code()
        self.mgr.ensure_participant(code)
        self.mgr.save()
        self._refresh_profiles_table()

    def _on_add_profile_field(self):
        dlg = ctk.CTkInputDialog(text="Eklenecek profil alanı adı (örn: Deneyim Yılı):", title="Profil Alanı Ekle")
        val = (dlg.get_input() or "").strip()
        if val and val not in self.mgr.data["profile_fields"]:
            self.mgr.data["profile_fields"].append(val)
            self.mgr.save()
            self._refresh_profiles_table()

    def _on_remove_profile_field(self):
        fields = self.mgr.data.get("profile_fields", [])
        if not fields:
            return
        dlg = ctk.CTkInputDialog(text=f"Silinecek alan adını yazın ({', '.join(fields)}):", title="Alan Sil")
        val = (dlg.get_input() or "").strip()
        if val in fields:
            fields.remove(val)
            self.mgr.save()
            self._refresh_profiles_table()

    def _on_export_profiles(self, fmt: str):
        self._save_profiles_from_ui()
        docx_p, xlsx_p = self.mgr.export_participant_profiles()
        open_path_crossplatform(xlsx_p if fmt == "xlsx" else docx_p)

    # =========================================================================
    # TAB 5: ARAMA, RAPOR, YÖNTEM ÖZETİ & NVIVO/MAXQDA
    # =========================================================================
    def _build_tab_search_reports(self):
        # Method Summary Box
        ms_box = ctk.CTkFrame(self.t_search, fg_color="#FFFFFF", corner_radius=12, border_width=1, border_color="#E2E8F0")
        ms_box.pack(fill="x", padx=8, pady=6)
        self.ms_lbl = ctk.CTkLabel(ms_box, text="", font=ctk.CTkFont(size=12, weight="bold"), text_color="#0F172A", justify="left")
        self.ms_lbl.pack(side="left", padx=14, pady=10)

        ctk.CTkButton(
            ms_box, text="📊 NVivo / MAXQDA (.txt)", height=32, fg_color="#7C3AED",
            font=ctk.CTkFont(size=11, weight="bold"),
            command=lambda: open_path_crossplatform(self.mgr.export_nvivo_maxqda_project())
        ).pack(side="right", padx=10, pady=8)

        ctk.CTkButton(
            ms_box, text="📘 Yöntem Özeti (.docx)", height=32, fg_color="#2563EB",
            font=ctk.CTkFont(size=11, weight="bold"),
            command=lambda: open_path_crossplatform(self.mgr.export_method_summary_docx())
        ).pack(side="right", padx=4, pady=8)

        # Search + Custom Report Bar
        sr_bar = ctk.CTkFrame(self.t_search, fg_color="#FFFFFF", corner_radius=12, border_width=1, border_color="#E2E8F0")
        sr_bar.pack(fill="x", padx=8, pady=4)
        self.sr_entry = ctk.CTkEntry(sr_bar, width=320, height=32, placeholder_text="Proje içinde kelime, tema veya not ara...")
        self.sr_entry.pack(side="left", padx=10, pady=8)
        self.sr_entry.bind("<Return>", lambda _: self._on_search_project())
        ctk.CTkButton(sr_bar, text="🔍 Ara", width=80, height=32, fg_color="#0F766E", command=self._on_search_project).pack(side="left", padx=4)

        ctk.CTkButton(sr_bar, text="📕 Seçili PDF Rapor", height=32, fg_color="#DC2626", font=ctk.CTkFont(size=11, weight="bold"), command=lambda: self._on_export_custom_report("pdf")).pack(side="right", padx=10)
        ctk.CTkButton(sr_bar, text="📘 Seçili Word Rapor", height=32, fg_color="#1D4ED8", font=ctk.CTkFont(size=11, weight="bold"), command=lambda: self._on_export_custom_report("docx")).pack(side="right", padx=4)

        self.sr_scroll = ctk.CTkScrollableFrame(self.t_search, fg_color="#FFFFFF", corner_radius=12, border_width=1, border_color="#E2E8F0")
        self.sr_scroll.pack(fill="both", expand=True, padx=8, pady=4)
        self._report_quote_vars = []

    def _refresh_method_summary(self):
        st = self.mgr.get_method_summary_stats()
        consent_str = " · ".join(f"{k}: {v}" for k, v in st["consent_counts"].items()) or "-"
        self.ms_lbl.configure(
            text=f"📊 Yöntem Özeti: {st['interview_count']} Görüşme ({st['reviewed_count']} Kontrol Edildi) · "
                 f"Katılımcı: {st['participant_count']} · Toplam Süre: {format_ts_hms(st['total_duration_sec'])} "
                 f"(Ort: {format_ts_hms(st['avg_duration_sec'])}) · Alıntı: {st['quote_count']}\n"
                 f"Onam Durumu: {consent_str}"
        )
        # Also populate selectable quotes/notes for custom report when not searching
        if not self.sr_entry.get().strip():
            self._render_report_selector()

    def _render_report_selector(self):
        for w in self.sr_scroll.winfo_children():
            w.destroy()
        self._report_quote_vars = []
        quotes = self.mgr.get_all_quotes()
        ctk.CTkLabel(self.sr_scroll, text="Rapor oluşturmak için dahil edilecek Alıntıları seçin (veya yukarıdan proje içi arama yapın):", font=ctk.CTkFont(size=12, weight="bold"), text_color="#334155").pack(anchor="w", padx=10, pady=6)
        if not quotes:
            ctk.CTkLabel(self.sr_scroll, text="Henüz işaretlenmiş alıntı yok.", text_color="#94A3B8").pack(pady=20)
            return
        for q in quotes:
            v = ctk.BooleanVar(value=True)
            ctk.CTkCheckBox(
                self.sr_scroll,
                text=f"[{q['participant']} · {format_ts_hms(q['start'])} · {q['theme']}] {q['text'][:110]}",
                variable=v,
                font=ctk.CTkFont(size=12),
            ).pack(anchor="w", padx=12, pady=3)
            self._report_quote_vars.append((v, q))

    def _on_search_project(self):
        q = self.sr_entry.get().strip()
        if not q:
            self._render_report_selector()
            return
        for w in self.sr_scroll.winfo_children():
            w.destroy()
        hits = self.mgr.search_project(q)
        if not hits:
            ctk.CTkLabel(self.sr_scroll, text=f"'{q}' için eşleşme bulunamadı.", text_color="#64748B").pack(pady=30)
            return
        for h in hits:
            row = ctk.CTkFrame(self.sr_scroll, fg_color="#F8FAFC", corner_radius=8)
            row.pack(fill="x", padx=8, pady=3)
            ctk.CTkButton(
                row, text=f"▶ {format_ts_hms(h['start'])}", width=85, height=26, fg_color="#0F766E",
                command=lambda ar=h["audio_rel"], st=h["start"]: self._play_audio_slice(ar, st)
            ).pack(side="left", padx=8, pady=6)
            ctk.CTkLabel(row, text=f"[{h['participant']} · {h['kind']}] {h['snippet'][:140]}", font=ctk.CTkFont(size=12), anchor="w").pack(side="left", fill="x", expand=True, padx=6)

    def _on_export_custom_report(self, fmt: str):
        sel_q = [q for v, q in self._report_quote_vars if v.get()]
        notes = []
        for iv in self.mgr.data.get("interviews", []):
            if iv.get("general_note", "").strip():
                notes.append({"label": f"{iv['id']} Genel Not", "note": iv["general_note"].strip()})
        out_p = self.mgr.export_custom_report(sel_q, notes, title=f"FIELD Rapor — {self.mgr.data.get('project_name')}", fmt=fmt)
        open_path_crossplatform(out_p)

    # =========================================================================
    # TAB 6: ETİK ANAHTAR (PAROLALI KOD–İSİM EŞLEŞTİRME)
    # =========================================================================
    def _build_tab_ethics(self):
        warn = ctk.CTkLabel(
            self.t_ethics,
            text="🔐 ETİK GÜVENLİK: Katılımcı Kod–İsim eşleştirmesi parolalı ayrı dosyada (_ETHICS_KEY_DO_NOT_SHARE.ilberkey) saklanır ve HİÇBİR Word/Excel/PDF/NVivo çıktısına dahil edilmez.",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#991B1B",
            wraplength=1050,
            justify="left",
        )
        warn.pack(anchor="w", padx=12, pady=10)

        self.eth_body = ctk.CTkFrame(self.t_ethics, fg_color="#FFFFFF", corner_radius=12, border_width=1, border_color="#FECACA")
        self.eth_body.pack(fill="both", expand=True, padx=10, pady=6)
        self._render_ethics_locked()

    def _render_ethics_locked(self):
        for w in self.eth_body.winfo_children():
            w.destroy()
        exists = os.path.isfile(self.mgr.ethics_file_path)
        msg = "Parolayı girerek Etik Kod–İsim Anahtarının kilidini açın:" if exists else "Bu proje için henüz Etik Anahtar oluşturulmadı. Yeni bir parola belirleyin:"
        ctk.CTkLabel(self.eth_body, text=msg, font=ctk.CTkFont(size=13, weight="bold")).pack(pady=(50, 12))
        pw_e = ctk.CTkEntry(self.eth_body, show="*", width=260, height=36, placeholder_text="Parola...")
        pw_e.pack(pady=6)

        def _unlock():
            pw = pw_e.get()
            if len(pw) < 4:
                messagebox.showwarning("Zayıf Parola", "Lütfen en az 4 karakterli bir parola girin.")
                return
            if exists:
                try:
                    self._unlocked_ethics = load_encrypted_ethics_key(self.mgr.ethics_file_path, pw)
                    self._ethics_pw = pw
                    self._render_ethics_unlocked()
                except Exception as ex:
                    messagebox.showerror("Erişim Reddedildi", str(ex))
            else:
                self._unlocked_ethics = {c: "" for c in sorted(self.mgr.data.get("participants", {}).keys())}
                self._ethics_pw = pw
                save_encrypted_ethics_key(self.mgr.ethics_file_path, pw, self._unlocked_ethics)
                self._render_ethics_unlocked()

        ctk.CTkButton(self.eth_body, text="🔓 Kilidi Aç / Oluştur", width=180, height=36, fg_color="#DC2626", command=_unlock).pack(pady=12)

    def _render_ethics_unlocked(self):
        for w in self.eth_body.winfo_children():
            w.destroy()
        top = ctk.CTkFrame(self.eth_body, fg_color="transparent")
        top.pack(fill="x", padx=12, pady=8)
        ctk.CTkLabel(top, text="🔓 Etik Kod–İsim Eşleştirme Tablosu (Şifreli)", font=ctk.CTkFont(size=14, weight="bold"), text_color="#065F46").pack(side="left")

        eth_entries = {}
        def _save_eth():
            payload = {c: e.get().strip() for c, e in eth_entries.items()}
            save_encrypted_ethics_key(self.mgr.ethics_file_path, self._ethics_pw, payload)
            self._unlocked_ethics = payload
            messagebox.showinfo("Şifrelendi", "Kod–isim eşleştirmesi parola ile şifrelenerek kaydedildi.")

        def _lock_eth():
            self._unlocked_ethics = None
            self._ethics_pw = None
            self._render_ethics_locked()

        ctk.CTkButton(top, text="🔒 Kilitle", width=90, height=30, fg_color="#64748B", command=_lock_eth).pack(side="right", padx=6)
        ctk.CTkButton(top, text="💾 Şifreli Kaydet", width=130, height=30, fg_color="#059669", command=_save_eth).pack(side="right", padx=6)

        sc = ctk.CTkScrollableFrame(self.eth_body, fg_color="#F8FAFC")
        sc.pack(fill="both", expand=True, padx=12, pady=8)
        all_codes = sorted(set(self.mgr.data.get("participants", {}).keys()) | set((self._unlocked_ethics or {}).keys()))
        if not all_codes:
            ctk.CTkLabel(sc, text="Henüz katılımcı kodu yok.", text_color="#64748B").pack(pady=30)
            return
        for c in all_codes:
            r = ctk.CTkFrame(sc, fg_color="#FFFFFF")
            r.pack(fill="x", pady=3, padx=6)
            ctk.CTkLabel(r, text=c, width=80, font=ctk.CTkFont(size=12, weight="bold"), text_color="#5B21B6").pack(side="left", padx=10, pady=6)
            e = ctk.CTkEntry(r, width=320, height=30, placeholder_text="Gerçek Ad Soyad (yalnızca şifreli anahtarda kalır)")
            e.insert(0, (self._unlocked_ethics or {}).get(c, ""))
            e.pack(side="left", padx=8, pady=6)
            eth_entries[c] = e

    # ---- Portable Export ----
    def _on_export_zip(self):
        dest = filedialog.asksaveasfilename(
            title="Taşınabilir FIELD Proje Paketi Kaydet",
            defaultextension=".zip",
            initialfile=f"{self.mgr.data.get('project_name')}_FIELD_Proje.zip",
            filetypes=[("ZIP Arşivi", "*.zip")],
        )
        if dest:
            self.mgr.export_portable_zip(dest)
            messagebox.showinfo("Hazır", f"Proje tüm ses kayıtları, deşifreler, notlar ve şifreli etik anahtarıyla paketlendi:\n{dest}")
