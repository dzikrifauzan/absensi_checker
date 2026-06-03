"""
config.py — Konfigurasi AbsensiChecker
Isi semua nilai yang diawali TODO sebelum dijalankan.
"""

import os
from dotenv import load_dotenv

# Load .env otomatis
load_dotenv()

CONFIG = {
    # ──────────────────────────────────────────────
    # HR PORTAL
    # ──────────────────────────────────────────────

    # URL halaman login HR portal
    "HR_PORTAL_URL": os.getenv("HR_PORTAL_URL", "https://hrportalDummy.co.id/login"),

    # URL halaman absensi/kehadiran individu
    # (URL setelah login, biasanya beda dari login)
    "ATTENDANCE_PAGE_URL": os.getenv(
        "ATTENDANCE_PAGE_URL",
        "https://hrportalDummy.co.id/attendance/individual"  # TODO: ganti
    ),

    # Kredensial login
    "HR_USERNAME": os.getenv("HR_USERNAME", ""),   # TODO: isi username kamu
    "HR_PASSWORD": os.getenv("HR_PASSWORD", ""),   # TODO: isi password kamu

    # ── Selector elemen halaman login (sesuaikan dengan portal) ──
    # Cek via F12 > Inspector di browser
    "LOGIN_FIELD_USERNAME": os.getenv("LOGIN_FIELD_USERNAME", "username"),   # name attr input
    "LOGIN_FIELD_PASSWORD": os.getenv("LOGIN_FIELD_PASSWORD", "password"),   # name attr input
    "LOGIN_BUTTON_CSS"    : os.getenv("LOGIN_BUTTON_CSS", "button[type=submit]"),

    # ── Selector filter periode di halaman absensi ──
    "FIELD_YEAR"       : os.getenv("FIELD_YEAR", "year"),    # name attr input tahun
    "FIELD_MONTH"      : os.getenv("FIELD_MONTH", "month"),  # name attr input bulan
    "SEARCH_BUTTON_CSS": os.getenv("SEARCH_BUTTON_CSS", "button[type=submit]"),

    # ──────────────────────────────────────────────
    # EMAIL (SMTP)
    # ──────────────────────────────────────────────

    # Untuk Gmail: aktifkan "App Password" di Google Account
    # (bukan password biasa jika 2FA aktif)
    "SMTP_HOST"    : os.getenv("SMTP_HOST", "smtp.gmail.com"),
    "SMTP_PORT"    : int(os.getenv("SMTP_PORT", "587")),
    "SMTP_USER"    : os.getenv("SMTP_USER", ""),        # TODO: email pengirim
    "SMTP_PASSWORD": os.getenv("SMTP_PASSWORD", ""),    # TODO: app password
    "FROM_EMAIL"   : os.getenv("FROM_EMAIL", ""),       # TODO: sama dengan SMTP_USER
    "TO_EMAIL"     : os.getenv("TO_EMAIL", ""),         # TODO: email tujuan (bisa sama)

    # ──────────────────────────────────────────────
    # PENYIMPANAN LOKAL
    # ──────────────────────────────────────────────

    # Folder tempat file XLS didownload
    "DOWNLOAD_DIR": os.getenv("DOWNLOAD_DIR", "/home/ubuntu/absensi_checker/downloads"),
}
