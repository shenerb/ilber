import os
import sys
import json
import time
import uuid
import hmac
import hashlib
import urllib.request
import urllib.error
import subprocess
import customtkinter as ctk
from tkinter import messagebox

SUPABASE_URL = "https://kkzgilkisvuflodpulnl.supabase.co"
SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImtremdpbGtpc3Z1ZmxvZHB1bG5sIiwicm9sZSI6ImFub24iLCJpYXQiOjE3OTE0NjQ4NzMsImV4cCI6MjEwNzA0MDg3M30.xZAP0R90TdyRZsBsB9lM9IwO3wWre0CHKDAQpt_2zMM"
LICENSE_SECRET_SALT = "ILBERTech_v2_Master_Signature_2026_#9981"

AUTH_SESSION_PATH = os.path.join(os.path.expanduser("~"), "Documents", "ILBER_Workspace", "auth_session.json")

# Page IDs in ilber_v2.py:
#   MIC   -> "tool_ses_kaydet"
#   EDU   -> "tool_ai_ozet"
#   GLOBE -> "tool_video_cevir"
#   FIELD -> "tool_ai_rapor"

MODULE_CODE_TO_PID = {
    "MIC": "tool_ses_kaydet",
    "EDU": "tool_ai_ozet",
    "GLOBE": "tool_video_cevir",
    "FIELD": "tool_ai_rapor",
}

PID_TO_MODULE_CODE = {v: k for k, v in MODULE_CODE_TO_PID.items()}

PACKAGE_TIERS = {
    "trial": {
        "code": "trial",
        "price": "$0.00",
        "modules": ["MIC", "EDU", "GLOBE", "FIELD"],
        "tr_name": "⏳ Ücretsiz Deneme Sürümü (Trial)",
        "en_name": "⏳ Free Trial Version",
        "de_name": "⏳ Kostenlose Testversion (Trial)",
        "tr_desc": "5 ücretsiz işlem hakkı ile tüm modülleri (MIC, EDU, GLOBE, FIELD) test edebilirsiniz.",
        "en_desc": "Test all 4 modules (MIC, EDU, GLOBE, FIELD) with 5 free trial actions.",
        "de_desc": "Testen Sie alle 4 Module mit 5 kostenlosen Aktionen.",
        "badge_color": "#D97706",
    },
    "paket_1": {
        "code": "paket_1",
        "price": "$9.99",
        "modules": ["MIC"],
        "tr_name": "🥉 Paket 1 · MIC ($9.99)",
        "en_name": "🥉 Package 1 · MIC ($9.99)",
        "de_name": "🥉 Paket 1 · MIC ($9.99)",
        "tr_desc": "Yalnızca MIC (Aynı Dilde Ses & Video Deşifre - Word/TXT/SRT/VTT) modülü sınırsız açıktır.",
        "en_desc": "Unlimited access to MIC (Same-Language Audio & Video Transcription).",
        "de_desc": "Unbegrenzter Zugriff auf das Modul MIC (Transkription).",
        "badge_color": "#9333EA",
    },
    "paket_2": {
        "code": "paket_2",
        "price": "$19.99",
        "modules": ["MIC", "EDU"],
        "tr_name": "🥈 Paket 2 · EDU ($19.99)",
        "en_name": "🥈 Package 2 · EDU ($19.99)",
        "de_name": "🥈 Paket 2 · EDU ($19.99)",
        "tr_desc": "MIC + EDU (Format Korumalı & Terim Sözlüğü Destekli Belge/Doküman Çevirici: Word, PDF, Excel, TXT) sınırsız açıktır.",
        "en_desc": "Unlimited access to MIC + EDU (Format-Preserving & Glossary-Supported Document Translator: Word, PDF, Excel, TXT).",
        "de_desc": "Unbegrenzter Zugriff auf MIC + EDU (Formatgetreuer Dokumentübersetzer mit Glossar-Unterstützung: Word, PDF, Excel, TXT).",
        "badge_color": "#2563EB",
    },
    "paket_3": {
        "code": "paket_3",
        "price": "$39.99",
        "modules": ["MIC", "EDU", "GLOBE"],
        "tr_name": "🥇 Paket 3 · GLOBE ($39.99)",
        "en_name": "🥇 Package 3 · GLOBE ($39.99)",
        "de_name": "🥇 Paket 3 · GLOBE ($39.99)",
        "tr_desc": "MIC + EDU + GLOBE (Çoklu Dil Video Çeviri Stüdyosu & MP4 Altyazı Gömme) sınırsız açıktır.",
        "en_desc": "Unlimited access to MIC + EDU + GLOBE (Multi-Language Video Translation Studio & MP4 Burn-in).",
        "de_desc": "Unbegrenzter Zugriff auf MIC + EDU + GLOBE (Video-Übersetzungsstudio).",
        "badge_color": "#10B981",
    },
    "paket_4": {
        "code": "paket_4",
        "price": "$79.99",
        "modules": ["MIC", "EDU", "GLOBE", "FIELD"],
        "tr_name": "💎 Paket 4 · FIELD PRO ($79.99)",
        "en_name": "💎 Package 4 · FIELD PRO ($79.99)",
        "de_name": "💎 Paket 4 · FIELD PRO ($79.99)",
        "tr_desc": "4 Modülün Tamamı (MIC + EDU + GLOBE + FIELD Akademik Konuşmacı Ayrımı & NVivo/MAXQDA) sınırsız açıktır!",
        "en_desc": "All 4 Modules Unlocked (MIC + EDU + GLOBE + FIELD Academic Diarization & NVivo/MAXQDA)!",
        "de_desc": "Alle 4 Module freigeschaltet (MIC + EDU + GLOBE + FIELD Forschungspaket)!",
        "badge_color": "#EF4444",
    },
}


def get_hardware_uuid():
    """Returns a deterministic hardware UUID string on both Windows and macOS without flashing any console window."""
    try:
        if os.name == "nt":
            import winreg
            key = winreg.OpenKey(
                winreg.HKEY_LOCAL_MACHINE,
                r"SOFTWARE\Microsoft\Cryptography",
                0,
                winreg.KEY_READ | winreg.KEY_WOW64_64KEY,
            )
            val, _ = winreg.QueryValueEx(key, "MachineGuid")
            winreg.CloseKey(key)
            if val and len(str(val).strip()) >= 8:
                return str(val).strip().upper()
        elif sys.platform == "darwin":
            out = subprocess.check_output(
                ["ioreg", "-rd1", "-c", "IOPlatformExpertDevice"],
                timeout=4,
            ).decode(errors="ignore")
            for line in out.splitlines():
                if "IOPlatformUUID" in line and "=" in line:
                    uuid_val = line.split("=", 1)[1].strip().strip('"').strip()
                    if len(uuid_val) >= 8:
                        return uuid_val.upper()
    except Exception:
        pass
    try:
        if os.name == "nt":
            cmd = 'powershell -NoProfile -Command "(Get-CimInstance -Class Win32_ComputerSystemProduct).UUID"'
            out = subprocess.check_output(cmd, shell=True, creationflags=0x08000000, timeout=4).decode(errors="ignore").strip()
            if out and len(out) >= 8:
                return out.upper()
    except Exception:
        pass
    return str(uuid.getnode()).upper()


def get_short_device_code(hw_uuid=None):
    """Produces a clean, customer-friendly device code like ILB-8F4A-992C from the hardware UUID."""
    hw = (hw_uuid or get_hardware_uuid()).strip().upper()
    digest = hashlib.sha256((hw + LICENSE_SECRET_SALT).encode("utf-8")).hexdigest().upper()
    return f"ILB-{digest[:4]}-{digest[4:8]}"


def generate_signed_license_key(device_code, plan_code):
    """Generates a deterministic offline-verifiable license key bound to a customer's device_code and plan_code."""
    dc = device_code.strip().upper()
    pc = plan_code.strip().lower()
    tier_tag = {
        "paket_1": "P1",
        "paket_2": "P2",
        "paket_3": "P3",
        "paket_4": "P4",
    }.get(pc, "P4")
    msg = f"{dc}|{tier_tag}|{LICENSE_SECRET_SALT}".encode("utf-8")
    sig = hmac.new(LICENSE_SECRET_SALT.encode("utf-8"), msg, hashlib.sha256).hexdigest().upper()[:8]
    return f"ILBER-{tier_tag}-{sig[:4]}-{sig[4:8]}"


def verify_signed_license_key(license_key, device_code):
    """
    Verifies a license key or test activation code.
    Returns (is_valid, plan_code, error_msg).
    """
    k = (license_key or "").strip().upper()
    dc = (device_code or get_short_device_code()).strip().upper()

    # Instant test / demo codes for partner & developer testing
    demo_map = {
        "ILBER-TRIAL": "trial",
        "ILBER-PAKET1": "paket_1",
        "ILBER-MIC": "paket_1",
        "ILBER-PAKET2": "paket_2",
        "ILBER-EDU": "paket_2",
        "ILBER-PAKET3": "paket_3",
        "ILBER-GLOBE": "paket_3",
        "ILBER-PAKET4": "paket_4",
        "ILBER-FIELD": "paket_4",
        "ILBER-PRO": "paket_4",
    }
    if k in demo_map:
        return True, demo_map[k], ""

    for p_code in ("paket_1", "paket_2", "paket_3", "paket_4"):
        expected = generate_signed_license_key(dc, p_code)
        if hmac.compare_digest(k, expected):
            return True, p_code, ""

    return False, None, "Geçersiz veya başka bir cihaza ait lisans anahtarı!"


class SupabaseLicenseManager:
    """
    Handles Supabase Authentication, Hardware ID (hwid) locking, Trial counter,
    and 4-tier Modular Package permissions with offline session caching.
    """
    def __init__(self):
        self.hwid = get_hardware_uuid()
        self.device_code = get_short_device_code(self.hwid)
        self.email = ""
        self.user_id = ""
        self.access_token = ""
        self.refresh_token = ""
        self.plan = "trial"
        self.modules = ["MIC", "EDU", "GLOBE", "FIELD"]
        self.trial_remaining = 5
        self.is_authenticated = False

    def _supabase_request(self, method, path, body=None, token=None, timeout=8):
        url = f"{SUPABASE_URL}{path}"
        headers = {
            "apikey": SUPABASE_ANON_KEY,
            "Content-Type": "application/json",
        }
        if token:
            headers["Authorization"] = f"Bearer {token}"
        data_bytes = json.dumps(body).encode("utf-8") if body is not None else None
        req = urllib.request.Request(url, data=data_bytes, headers=headers, method=method)
        try:
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                raw = resp.read().decode("utf-8", errors="ignore")
                return resp.status, (json.loads(raw) if raw else {})
        except urllib.error.HTTPError as e:
            raw = e.read().decode("utf-8", errors="ignore")
            try:
                err_json = json.loads(raw) if raw else {}
            except Exception:
                err_json = {"msg": raw}
            return e.code, err_json
        except Exception as e:
            return 0, {"error": str(e)}

    def _apply_user_metadata(self, user_obj):
        self.user_id = str(user_obj.get("id", ""))
        self.email = str(user_obj.get("email", self.email))
        meta = user_obj.get("user_metadata") or {}

        plan = str(meta.get("plan", "")).strip().lower()
        if plan not in PACKAGE_TIERS:
            plan = "trial"
        self.plan = plan

        if "trial_remaining" in meta:
            try:
                self.trial_remaining = max(0, int(meta.get("trial_remaining", 5)))
            except Exception:
                self.trial_remaining = 5
        else:
            self.trial_remaining = 5

        custom_mods = meta.get("modules")
        if isinstance(custom_mods, list) and custom_mods:
            self.modules = [str(m).upper() for m in custom_mods]
        else:
            self.modules = list(PACKAGE_TIERS[self.plan]["modules"])

    def save_local_session(self):
        try:
            os.makedirs(os.path.dirname(AUTH_SESSION_PATH), exist_ok=True)
            payload = {
                "email": self.email,
                "user_id": self.user_id,
                "hwid": self.hwid,
                "device_code": self.device_code,
                "access_token": self.access_token,
                "refresh_token": self.refresh_token,
                "plan": self.plan,
                "modules": self.modules,
                "trial_remaining": self.trial_remaining,
                "saved_at": int(time.time()),
            }
            sig_str = f"{self.email}|{self.hwid}|{self.plan}|{','.join(self.modules)}|{self.trial_remaining}|{LICENSE_SECRET_SALT}"
            payload["signature"] = hashlib.sha256(sig_str.encode("utf-8")).hexdigest()
            with open(AUTH_SESSION_PATH, "w", encoding="utf-8") as f:
                json.dump(payload, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def clear_local_session(self):
        self.is_authenticated = False
        self.access_token = ""
        self.refresh_token = ""
        try:
            if os.path.exists(AUTH_SESSION_PATH):
                os.remove(AUTH_SESSION_PATH)
        except Exception:
            pass

    @property
    def session(self):
        return {
            "logged_in": self.is_authenticated,
            "email": self.email,
            "user_id": self.user_id,
            "plan": self.plan,
            "modules": self.modules,
            "trial_remaining": self.trial_remaining,
            "device_code": self.device_code,
        }

    def _save_cache(self):
        self.save_local_session()

    def logout(self):
        self.clear_local_session()

    def activate_package_with_key(self, license_key):
        ok, plan_code, err_msg = verify_signed_license_key(license_key, self.device_code)
        if not ok:
            return False, err_msg
        trial_rem = 5 if plan_code == "trial" else self.trial_remaining
        self.is_authenticated = True
        if not self.email:
            self.email = "deneme@ilbertech.com"
        return self.update_plan_on_supabase(plan_code, trial_rem=trial_rem)

    def refresh_from_cloud(self):
        return self.refresh_from_supabase()

    def try_auto_login(self, sync_cloud=False):
        """
        Checks saved local session first (instant ~1ms startup).
        If sync_cloud=True, refreshes from Supabase synchronously; otherwise caller can sync in a background thread.
        Returns bool.
        """
        if not os.path.exists(AUTH_SESSION_PATH):
            return False
        try:
            with open(AUTH_SESSION_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            return False

        saved_hwid = str(data.get("hwid", "")).strip().upper()
        if saved_hwid and saved_hwid != self.hwid:
            return False

        email = str(data.get("email", ""))
        plan = str(data.get("plan", "trial"))
        mods = data.get("modules", ["MIC", "EDU", "GLOBE", "FIELD"])
        t_rem = int(data.get("trial_remaining", 5))
        expected_sig = hashlib.sha256(
            f"{email}|{saved_hwid}|{plan}|{','.join(mods)}|{t_rem}|{LICENSE_SECRET_SALT}".encode("utf-8")
        ).hexdigest()

        if data.get("signature") != expected_sig:
            return False

        self.email = email
        self.user_id = str(data.get("user_id", ""))
        self.access_token = str(data.get("access_token", ""))
        self.refresh_token = str(data.get("refresh_token", ""))
        self.plan = plan if plan in PACKAGE_TIERS else "trial"
        self.modules = mods
        self.trial_remaining = t_rem
        self.is_authenticated = True

        if sync_cloud:
            return self.sync_session_in_background()
        return True

    def sync_session_in_background(self):
        """Refreshes token and syncs `user_licenses` table in a non-blocking background thread."""
        try:
            if self.refresh_token:
                st, resp = self._supabase_request(
                    "POST",
                    "/auth/v1/token?grant_type=refresh_token",
                    {"refresh_token": self.refresh_token},
                    timeout=4,
                )
                if st == 200 and resp.get("access_token"):
                    self.access_token = resp["access_token"]
                    self.refresh_token = resp.get("refresh_token", self.refresh_token)
                    user_obj = resp.get("user") or {}
                    meta = user_obj.get("user_metadata") or {}
                    db_hwid = str(meta.get("hwid", "")).strip().upper()
                    if db_hwid and db_hwid != self.hwid:
                        self.clear_local_session()
                        return False
                    self._apply_user_metadata(user_obj)

            if self.email:
                self._sync_with_user_licenses_table()
            self.save_local_session()
            return True
        except Exception:
            return False

    def sign_in(self, email, password, remember=True):
        """Authenticates against Supabase Auth REST API and enforces Hardware ID lock."""
        if not email or not password:
            return False, "Lütfen e-posta adresinizi ve şifrenizi girin."
        st, resp = self._supabase_request(
            "POST",
            "/auth/v1/token?grant_type=password",
            {"email": email, "password": password},
        )
        if st == 0:
            return False, "İnternet bağlantısı kurulamadı. İlk giriş için internet gereklidir."
        if st != 200 or not resp.get("access_token"):
            err_desc = resp.get("error_description") or resp.get("msg") or "E-posta veya şifre hatalı!"
            if "Invalid login credentials" in str(err_desc):
                err_desc = "E-posta veya şifre hatalı! Henüz hesabınız yoksa 'Yeni Hesap Oluştur' butonunu kullanın."
            elif "Email not confirmed" in str(err_desc):
                err_desc = "E-posta adresiniz henüz onaylanmamış. Lütfen e-posta kutunuzu kontrol edin."
            return False, str(err_desc)

        self.access_token = resp["access_token"]
        self.refresh_token = resp.get("refresh_token", "")
        user_obj = resp.get("user") or {}
        meta = user_obj.get("user_metadata") or {}
        db_hwid = str(meta.get("hwid", "")).strip().upper()

        if db_hwid and db_hwid != self.hwid:
            return False, f"Bu hesap başka bir bilgisayara ({meta.get('device_code', 'Farklı Cihaz')}) kilitlenmiştir!"

        # If first login on this account, bind hwid and initialize trial if no plan exists
        need_update = False
        new_meta = dict(meta)
        if not db_hwid:
            new_meta["hwid"] = self.hwid
            new_meta["device_code"] = self.device_code
            need_update = True
        if "plan" not in new_meta:
            new_meta["plan"] = "trial"
            new_meta["modules"] = ["MIC", "EDU", "GLOBE", "FIELD"]
            new_meta["trial_remaining"] = 5
            need_update = True

        if need_update:
            st_u, resp_u = self._supabase_request(
                "PUT",
                "/auth/v1/user",
                {"data": new_meta},
                token=self.access_token,
            )
            if st_u == 200 and resp_u.get("user_metadata"):
                user_obj = resp_u
            else:
                user_obj["user_metadata"] = new_meta

        self._apply_user_metadata(user_obj)
        self._sync_with_user_licenses_table()
        self.is_authenticated = True
        if remember:
            self.save_local_session()
        return True, "Giriş başarılı!"

    def _sync_with_user_licenses_table(self):
        """
        Syncs the user's plan and trial_remaining by EMAIL with the `public.user_licenses` table in Supabase.
        - If a row for `self.email` exists in `public.user_licenses`, its `plan` and `trial_remaining` take priority!
        - If no row exists yet, inserts a new row for `self.email` so the admin can see and edit it in Supabase Table Editor.
        """
        if not self.email:
            return False
        import urllib.parse
        enc_email = urllib.parse.quote(self.email.strip().lower())
        tok = self.access_token or SUPABASE_ANON_KEY
        st, rows = self._supabase_request(
            "GET",
            f"/rest/v1/user_licenses?email=eq.{enc_email}&select=*",
            token=tok,
            timeout=5,
        )
        if st == 200 and isinstance(rows, list):
            if len(rows) > 0:
                row = rows[0]
                db_hwid = str(row.get("hwid") or "").strip().upper()
                # Bind hwid if empty in table
                if not db_hwid and self.hwid:
                    self._supabase_request(
                        "PATCH",
                        f"/rest/v1/user_licenses?email=eq.{enc_email}",
                        {"hwid": self.hwid, "device_code": self.device_code},
                        token=tok,
                        timeout=4,
                    )
                db_plan = str(row.get("plan") or "trial").strip().lower()
                if db_plan in PACKAGE_TIERS:
                    self.plan = db_plan
                    self.modules = list(PACKAGE_TIERS[db_plan]["modules"])
                if row.get("trial_remaining") is not None:
                    try:
                        self.trial_remaining = max(0, int(row.get("trial_remaining", 5)))
                    except Exception:
                        pass
                return True
            else:
                # Row doesn't exist yet -> automatically insert this email into public.user_licenses!
                new_row = {
                    "email": self.email.strip().lower(),
                    "plan": self.plan,
                    "trial_remaining": self.trial_remaining,
                    "device_code": self.device_code,
                    "hwid": self.hwid,
                }
                self._supabase_request(
                    "POST",
                    "/rest/v1/user_licenses",
                    new_row,
                    token=tok,
                    timeout=5,
                )
                return True
        return False

    def admin_grant_plan_by_email(self, target_email, plan_code, trial_rem=5):
        """
        Admin helper: Grants `plan_code` ('trial', 'paket_1', 'paket_2', 'paket_3', 'paket_4')
        to `target_email` in Supabase `public.user_licenses`.
        """
        em = (target_email or "").strip().lower()
        if not em or "@" not in em:
            return False, "Lütfen geçerli bir müşteri e-posta adresi girin."
        if plan_code not in PACKAGE_TIERS:
            return False, "Geçersiz paket kodu."
        import urllib.parse
        enc_email = urllib.parse.quote(em)
        tok = self.access_token or SUPABASE_ANON_KEY

        # Check if email row exists
        st_get, rows = self._supabase_request(
            "GET",
            f"/rest/v1/user_licenses?email=eq.{enc_email}&select=*",
            token=tok,
            timeout=6,
        )
        if st_get == 404 or (isinstance(rows, dict) and "user_licenses" in str(rows)):
            return False, "Supabase'de 'user_licenses' tablosu henüz oluşturulmamış. Lütfen aşağıdaki 1-Adımlık SQL Kurulum kodunu Supabase SQL Editor'de çalıştırın."

        if st_get == 200 and isinstance(rows, list) and len(rows) > 0:
            st_up, _ = self._supabase_request(
                "PATCH",
                f"/rest/v1/user_licenses?email=eq.{enc_email}",
                {"plan": plan_code, "trial_remaining": trial_rem},
                token=tok,
                timeout=6,
            )
            if st_up in (200, 204):
                return True, f"✔ '{em}' adresine '{PACKAGE_TIERS[plan_code]['tr_name']}' yetkisi başarıyla tanımlandı!"
        else:
            st_ins, resp_ins = self._supabase_request(
                "POST",
                "/rest/v1/user_licenses",
                {"email": em, "plan": plan_code, "trial_remaining": trial_rem},
                token=tok,
                timeout=6,
            )
            if st_ins in (200, 201, 204):
                return True, f"✔ '{em}' e-postası tabloya eklendi ve '{PACKAGE_TIERS[plan_code]['tr_name']}' yetkisi verildi!"
            return False, f"Supabase Hatası ({st_ins}): {resp_ins}"
        return False, "Yetki güncellenemedi."

    def sign_up_trial(self, email, password, remember=True):
        """Creates a new user account on Supabase with 5 free trial credits bound to this PC's hwid."""
        if not email or "@" not in email or len(password) < 6:
            return False, "Lütfen geçerli bir e-posta ve en az 6 karakterli bir şifre girin."
        email = email.strip().lower()
        init_meta = {
            "hwid": self.hwid,
            "device_code": self.device_code,
            "plan": "trial",
            "modules": ["MIC", "EDU", "GLOBE", "FIELD"],
            "trial_remaining": 5,
            "created_at_ts": int(time.time()),
        }
        st, resp = self._supabase_request(
            "POST",
            "/auth/v1/signup",
            {"email": email, "password": password, "data": init_meta},
        )
        if st == 0:
            return False, "İnternet bağlantısı kurulamadı."
        if st not in (200, 201):
            err_desc = resp.get("msg") or resp.get("error_description") or resp.get("message") or "Kayıt oluşturulamadı."
            if "already registered" in str(err_desc).lower():
                return self.sign_in(email, password, remember=remember)
            return False, f"Kayıt Hatası: {err_desc}"

        # If Supabase returns an active session immediately
        if resp.get("access_token"):
            self.access_token = resp["access_token"]
            self.refresh_token = resp.get("refresh_token", "")
            user_obj = resp.get("user") or {"email": email, "user_metadata": init_meta}
            self._apply_user_metadata(user_obj)
            self._sync_with_user_licenses_table()
            self.is_authenticated = True
            if remember:
                self.save_local_session()
            return True, "🎉 Hesabınız oluşturuldu ve 5 Ücretsiz Deneme Hakkınız tanımlandı!"

        # Try immediate sign-in; if email confirmation is required by Supabase settings, start local trial and register in user_licenses table
        ok_in, msg_in = self.sign_in(email, password, remember=remember)
        if ok_in:
            return True, "🎉 Hesabınız oluşturuldu ve 5 Ücretsiz Deneme Hakkınız tanımlandı!"
        self.email = email
        self.plan = "trial"
        self.modules = ["MIC", "EDU", "GLOBE", "FIELD"]
        self.trial_remaining = 5
        self._sync_with_user_licenses_table()
        self.is_authenticated = True
        self.save_local_session()
        return True, "🎉 Hesabınız oluşturuldu ve 5 Ücretsiz Deneme Hakkınız başlatıldı!"

    def refresh_from_supabase(self):
        """Pulls the latest plan by EMAIL from `public.user_licenses` (and `auth.users`) on Supabase."""
        if not self.email:
            return False, "Sorgulama için e-posta adresinizle giriş yapmış olmalısınız."
        synced_table = self._sync_with_user_licenses_table()
        if self.access_token and not synced_table:
            st, resp = self._supabase_request("GET", "/auth/v1/user", token=self.access_token)
            if st == 200 and resp.get("id"):
                self._apply_user_metadata(resp)
                synced_table = True
        if synced_table:
            self.save_local_session()
            return True, f"Yetkiniz Supabase'den güncellendi: {self.get_plan_display_name('tr')}"
        return False, "Supabase'den yetki bilgisi alınamadı (internet bağlantınızı kontrol edin)."

    def update_plan_on_supabase(self, new_plan_code, custom_modules=None, trial_rem=None):
        """Updates the user's package tier on Supabase (`user_licenses` + `auth.users`) and local cache."""
        if new_plan_code not in PACKAGE_TIERS:
            return False, "Geçersiz paket kodu."
        self.plan = new_plan_code
        self.modules = list(custom_modules) if custom_modules else list(PACKAGE_TIERS[new_plan_code]["modules"])
        if trial_rem is not None:
            self.trial_remaining = trial_rem
        elif new_plan_code == "trial":
            self.trial_remaining = 5

        new_meta = {
            "hwid": self.hwid,
            "device_code": self.device_code,
            "plan": self.plan,
            "modules": self.modules,
            "trial_remaining": self.trial_remaining,
        }
        if self.access_token:
            self._supabase_request("PUT", "/auth/v1/user", {"data": new_meta}, token=self.access_token)
        if self.email:
            import urllib.parse
            enc_email = urllib.parse.quote(self.email.strip().lower())
            tok = self.access_token or SUPABASE_ANON_KEY
            self._supabase_request(
                "PATCH",
                f"/rest/v1/user_licenses?email=eq.{enc_email}",
                {"plan": self.plan, "trial_remaining": self.trial_remaining},
                token=tok,
                timeout=4,
            )
        self.save_local_session()
        return True, f"Paketiniz başarıyla etkinleştirildi: {PACKAGE_TIERS[self.plan]['tr_name']}"

    def consume_trial_credit(self):
        """If the user is on the trial plan, decrements trial_remaining by 1, syncs to Supabase, and returns remaining credits."""
        if self.plan != "trial":
            return self.trial_remaining
        self.trial_remaining = max(0, self.trial_remaining - 1)
        new_meta = {
            "hwid": self.hwid,
            "device_code": self.device_code,
            "plan": self.plan,
            "modules": self.modules,
            "trial_remaining": self.trial_remaining,
        }
        if self.access_token:
            self._supabase_request("PUT", "/auth/v1/user", {"data": new_meta}, token=self.access_token, timeout=4)
        if self.email:
            import urllib.parse
            enc_email = urllib.parse.quote(self.email.strip().lower())
            tok = self.access_token or SUPABASE_ANON_KEY
            self._supabase_request(
                "PATCH",
                f"/rest/v1/user_licenses?email=eq.{enc_email}",
                {"trial_remaining": self.trial_remaining},
                token=tok,
                timeout=4,
            )
        self.save_local_session()
        return self.trial_remaining

    def is_module_unlocked(self, page_id):
        """Returns True if the user's current package (or active Trial) allows opening `page_id`."""
        mod_code = PID_TO_MODULE_CODE.get(page_id)
        if not mod_code:
            return True
        if self.plan == "trial":
            return self.trial_remaining > 0
        return mod_code in self.modules

    def get_module_badge_info(self, page_id, ui_lang="tr"):
        """
        Returns (badge_text, text_color, bg_color) for the 4 Home cards.
        """
        mod_code = PID_TO_MODULE_CODE.get(page_id, "MIC")
        req_tier_price = {
            "MIC": "$9.99",
            "EDU": "$19.99",
            "GLOBE": "$39.99",
            "FIELD": "$79.99",
        }.get(mod_code, "$19.99")

        if self.plan == "trial":
            if self.trial_remaining > 0:
                txt = {
                    "tr": f"⏳ DENEME ({self.trial_remaining} Hak)",
                    "en": f"⏳ TRIAL ({self.trial_remaining} Left)",
                    "de": f"⏳ TEST ({self.trial_remaining} übrig)",
                }.get(ui_lang, f"⏳ DENEME ({self.trial_remaining} Hak)")
                return txt, "#B45309", "#FEF3C7"
            else:
                txt = {
                    "tr": f"🔒 DENEME BİTTİ · {req_tier_price}",
                    "en": f"🔒 TRIAL ENDED · {req_tier_price}",
                    "de": f"🔒 TEST BEENDET · {req_tier_price}",
                }.get(ui_lang, f"🔒 DENEME BİTTİ · {req_tier_price}")
                return txt, "#991B1B", "#FEE2E2"

        if mod_code in self.modules:
            txt = {
                "tr": "✓ LİSANSLI · AÇIK",
                "en": "✓ LICENSED · ACTIVE",
                "de": "✓ LIZENZIERT · AKTIV",
            }.get(ui_lang, "✓ LİSANSLI · AÇIK")
            return txt, "#065F46", "#D1FAE5"
        else:
            txt = {
                "tr": f"🔒 KİLİTLİ · Yükselt ({req_tier_price})",
                "en": f"🔒 LOCKED · Upgrade ({req_tier_price})",
                "de": f"🔒 GESPERRT · Upgrade ({req_tier_price})",
            }.get(ui_lang, f"🔒 KİLİTLİ · Yükselt ({req_tier_price})")
            return txt, "#991B1B", "#FEE2E2"

    def get_plan_display_name(self, ui_lang="tr"):
        info = PACKAGE_TIERS.get(self.plan, PACKAGE_TIERS["trial"])
        name_key = f"{ui_lang}_name" if f"{ui_lang}_name" in info else "tr_name"
        base = info[name_key]
        if self.plan == "trial":
            rem_lbl = {
                "tr": f"({self.trial_remaining} Hak Kaldı)",
                "en": f"({self.trial_remaining} Uses Left)",
                "de": f"({self.trial_remaining} übrig)",
            }.get(ui_lang, f"({self.trial_remaining} Hak)")
            return f"{base} {rem_lbl}"
        return base

