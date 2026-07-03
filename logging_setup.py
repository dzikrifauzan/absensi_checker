"""
Konfigurasi logging terpusat untuk semua modul AbsensiChecker.
Import `log` dari sini di modul lain, jangan panggil basicConfig lagi.
"""

__version__ = "1.0.0"

import logging
from pathlib import Path

Path("logs").mkdir(exist_ok=True)  # auto-create biar gak crash kalau folder belum ada

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("logs/absensi_checker.log"),
        logging.StreamHandler()
    ]
)

log = logging.getLogger("absensi_checker")