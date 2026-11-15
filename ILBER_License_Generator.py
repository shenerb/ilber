import customtkinter as ctk
from tkinter import messagebox
from ilber_license import SupabaseLicenseManager

SUPABASE_SETUP_SQL = """CREATE TABLE IF NOT EXISTS public.user_licenses (
    email TEXT PRIMARY KEY,
    plan TEXT NOT NULL DEFAULT 'trial',
    trial_remaining INT NOT NULL DEFAULT 5,
    device_code TEXT,
    hwid TEXT,
    updated_at TIMESTAMPTZ DEFAULT NOW()
);
ALTER TABLE public.user_licenses DISABLE ROW LEVEL SECURITY;
GRANT ALL ON TABLE public.user_licenses TO anon, authenticated, service_role;"""


class SupabaseAdminPanel(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.mgr = SupabaseLicenseManager()
        self.title("ILBERTech — E-Posta İle Supabase Yetki Yönetim Paneli")
        self.geometry("680x540")
        self.resizable(False, False)
        self.configure(fg_color="#0F172A")

        hdr = ctk.CTkFrame(self, fg_color="#1E293B", corner_radius=0, height=74)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        ctk.CTkLabel(
            hdr,
            text="☁️ ILBERTech — E-Posta İle Supabase Yetki Yönetimi",
            font=ctk.CTkFont(family="Segoe UI", size=20, weight="bold"),
            text_color="#F8FAFC",
        ).pack(anchor="w", padx=24, pady=(14, 0))
        ctk.CTkLabel(
            hdr,
            text="Müşterinin e-posta adresine Supabase üzerinden anında paket yetkisi tanımlayın (Lisans kodu gerekmez)",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color="#38BDF8",
        ).pack(anchor="w", padx=24)

        # 1-Time SQL Setup Box
        sql_box = ctk.CTkFrame(self, fg_color="#1E293B", corner_radius=12, border_width=1, border_color="#334155")
        sql_box.pack(fill="x", padx=24, pady=(16, 8))
        ctk.CTkLabel(
            sql_box,
            text="🛠️ İlk Kurulum (Sadece 1 Kez Yapılır): Supabase 'SQL Editor'de 'user_licenses' tablosunu oluşturmak için:",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color="#FDE68A",
        ).pack(anchor="w", padx=16, pady=(10, 4))

        def copy_sql():
            self.clipboard_clear()
            self.clipboard_append(SUPABASE_SETUP_SQL)
            messagebox.showinfo(
                "SQL Kopyalandı!",
                "Tablo kurulum SQL kodu panoya kopyalandı!\n\n"
                "Şimdi Supabase sol menüsünden 'SQL Editor'e tıklayın, yapıştırın (Ctrl+V) ve 'Run' butonuna basın.\n"
                "Bunu sadece 1 kez yapmanız yeterlidir!",
            )

        ctk.CTkButton(
            sql_box,
            text="📋 1-Defalık Supabase Tablo Kurulum SQL Kodunu Kopyala",
            height=34,
            fg_color="#334155",
            hover_color="#475569",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            command=copy_sql,
        ).pack(fill="x", padx=16, pady=(0, 10))

        # Main Email Permission Form
        body = ctk.CTkFrame(self, fg_color="#1E293B", corner_radius=14, border_width=1, border_color="#334155")
        body.pack(fill="both", expand=True, padx=24, pady=(6, 20))

        ctk.CTkLabel(
            body,
            text="1. Müşterinin Kayıtlı E-Posta Adresi:",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color="#E2E8F0",
        ).pack(anchor="w", padx=20, pady=(18, 6))

        self.email_entry = ctk.CTkEntry(
            body,
            placeholder_text="ornek@gmail.com",
            height=42,
            font=ctk.CTkFont(family="Segoe UI", size=15),
            fg_color="#0F172A",
            text_color="#FFFFFF",
            border_color="#475569",
        )
        if self.mgr.email:
            self.email_entry.insert(0, self.mgr.email)
        self.email_entry.pack(fill="x", padx=20, pady=(0, 14))

        ctk.CTkLabel(
            body,
            text="2. E-Postaya Tanımlanacak Paket Yetkisi:",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color="#E2E8F0",
        ).pack(anchor="w", padx=20, pady=(0, 6))

        self.pkg_map = {
            "⏳ Ücretsiz Deneme (Trial — 5 Hak)": "trial",
            "🎙️ Paket 1: MIC ($9.99) — Sadece MIC Açık": "paket_1",
            "🎓 Paket 2: EDU ($19.99) — MIC + EDU Açık": "paket_2",
            "🌐 Paket 3: GLOBE ($39.99) — MIC + EDU + GLOBE Açık": "paket_3",
            "🔬 Paket 4: FIELD PRO ($79.99) — Tüm 4 Modül Açık": "paket_4",
        }
        self.pkg_opt = ctk.CTkOptionMenu(
            body,
            values=list(self.pkg_map.keys()),
            height=42,
            fg_color="#2563EB",
            button_color="#1D4ED8",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
        )
        self.pkg_opt.set("🎓 Paket 2: EDU ($19.99) — MIC + EDU Açık")
        self.pkg_opt.pack(fill="x", padx=20, pady=(0, 18))

        self.status_lbl = ctk.CTkLabel(
            body,
            text="",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color="#34D399",
            wraplength=600,
        )
        self.status_lbl.pack(padx=20, pady=(0, 8))

        ctk.CTkButton(
            body,
            text="☁️ Supabase'de Bu E-Postaya Yetki Ver",
            height=46,
            fg_color="#10B981",
            hover_color="#059669",
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            command=self.grant_permission,
        ).pack(fill="x", padx=20, pady=(4, 14))

    def grant_permission(self):
        em = self.email_entry.get().strip().lower()
        plan_code = self.pkg_map.get(self.pkg_opt.get(), "paket_2")
        self.status_lbl.configure(text="🔄 Supabase'e bağlanılıyor...", text_color="#38BDF8")
        self.update_idletasks()
        ok, msg = self.mgr.admin_grant_plan_by_email(em, plan_code, trial_rem=5)
        if ok:
            self.status_lbl.configure(text=msg, text_color="#34D399")
            messagebox.showinfo("Başarılı!", f"{msg}\n\nKullanıcı programda '🔄 Yetkimi / Paketimi Güncelle' butonuna bastığında yeni paketi anında açılacaktır!")
        else:
            self.status_lbl.configure(text=f"❌ {msg}", text_color="#F87171")


if __name__ == "__main__":
    app = SupabaseAdminPanel()
    app.mainloop()
