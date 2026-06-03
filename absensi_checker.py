"""
╔══════════════════════════════════════════════════════════╗
║         AUTOMATION CEK ABSENSI HR PORTAL                 ║
║  - Login ke HR Portal via Selenium                       ║
║  - Download laporan Kehadiran Individu (XLS)             ║
║  - Analisis: Mangkir, Sakit > 6 hari, Keterlambatan      ║
║  - Kirim email ringkasan bulanan                         ║
╚══════════════════════════════════════════════════════════╝
"""

import os
import time
import logging
import smtplib
import traceback
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from pathlib import Path

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service

from config import CONFIG
from analyzer import AbsensiAnalyzer

# ─── Setup Logging ────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("logs/absensi_checker.log"),
        logging.StreamHandler()
    ]
)
log = logging.getLogger(__name__)


# ─── Selenium HR Portal Scraper ───────────────────────────
class HRPortalScraper:
    def __init__(self):
        self.driver = None
        self.download_dir = Path(CONFIG["DOWNLOAD_DIR"]).resolve()
        self.download_dir.mkdir(parents=True, exist_ok=True)
        self.screenshot_dir = Path("screenshots")
        self.screenshot_dir.mkdir(exist_ok=True)

    def _shot(self, step: str):
        path = self.screenshot_dir / f"{step}.png"
        self.driver.save_screenshot(str(path))

    def _init_driver(self):
        """Inisialisasi Chrome headless dengan anti-detection."""
        chrome_options = Options()
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--window-size=1920,1080")

        # Anti-bot detection
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_argument(
            "user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0.0.0 Safari/537.36"
        )
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option("useAutomationExtension", False)

        # Auto-download tanpa popup
        prefs = {
            "download.default_directory": str(self.download_dir),
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        }
        chrome_options.add_experimental_option("prefs", prefs)

        self.driver = webdriver.Chrome(options=chrome_options)
        self.driver.execute_cdp_cmd(
            "Page.addScriptToEvaluateOnNewDocument",
            {"source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"}
        )
        self.driver.implicitly_wait(10)
        log.info("Chrome driver berhasil diinisialisasi.")

    def _wait(self, seconds=1):
        time.sleep(seconds)

    def login(self):
        """Login ke HR Portal."""
        log.info(f"Membuka HR Portal: {CONFIG['HR_PORTAL_URL']}")
        self.driver.get(CONFIG["HR_PORTAL_URL"])
        self._wait(5)

        wait = WebDriverWait(self.driver, 30)

        username_field = wait.until(
            EC.visibility_of_element_located((By.ID, CONFIG["LOGIN_FIELD_USERNAME"]))
        )
        password_field = wait.until(
            EC.visibility_of_element_located((By.ID, CONFIG["LOGIN_FIELD_PASSWORD"]))
        )

        username_field.clear()
        username_field.send_keys(CONFIG["HR_USERNAME"])
        self._wait(1)
        password_field.clear()
        password_field.send_keys(CONFIG["HR_PASSWORD"])
        self._wait(1)

        login_btn = wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, CONFIG["LOGIN_BUTTON_CSS"]))
        )
        login_btn.click()
        self._wait(5)

        log.info(f"Login berhasil. URL: {self.driver.current_url}")

    def navigate_to_attendance(self, year: int, month: int):
        """
        Set periode via jQuery datepicker API.
        getParam() membaca: $("#periode-text").datepicker("getFormattedDate", "yyyy-mm-dd")
        Jadi harus set via datepicker('update'), bukan send_keys biasa.
        """
        periode_date = f"{year}-{month:02d}-01"  # format yyyy-mm-dd untuk JS Date

        log.info(f"Navigasi ke halaman absensi {year}-{month:02d}")
        self.driver.get(CONFIG["ATTENDANCE_PAGE_URL"])
        self._wait(3)

        wait = WebDriverWait(self.driver, 20)
        wait.until(EC.visibility_of_element_located((By.ID, "periode-text")))

        # ── Set datepicker via jQuery API ──────────────────
        # getParam() pakai: $("#periode-text").datepicker("getFormattedDate", "yyyy-mm-dd")
        # Jadi harus update lewat datepicker('update'), bukan isi field manual
        self.driver.execute_script(
            "var d = new Date(arguments[0]); $('#periode-text').datepicker('update', d);",
            periode_date
        )
        self._wait(1)

        # Verifikasi getParam() return periode yang benar
        param = self.driver.execute_script("return JSON.stringify(getParam())")
        log.info(f"getParam setelah set: {param}")

        # Klik tombol Cari via JS (hindari overlap elemen)
        search_btn = wait.until(EC.presence_of_element_located((By.ID, "search-button")))
        self.driver.execute_script("arguments[0].click()", search_btn)
        self._wait(5)

        # Verifikasi data sudah reload ke periode yang benar
        param_after = self.driver.execute_script("return JSON.stringify(getParam())")
        log.info(f"getParam setelah search: {param_after}")
        log.info(f"Halaman absensi {year}-{month:02d} berhasil dimuat.")

    def download_excel(self) -> Path:
        """Klik tombol Unduh dan tunggu file terdownload."""
        wait = WebDriverWait(self.driver, 15)

        # Catat file yang sudah ada SEBELUM download
        existing_files = set(self.download_dir.glob("Kehadiran_Individu_*.xls*"))
        log.info(f"File existing sebelum download: {[f.name for f in existing_files]}")

        download_btn = wait.until(
            EC.element_to_be_clickable((By.ID, "download-button"))
        )
        self.driver.execute_script("arguments[0].click()", download_btn)
        log.info("Tombol Unduh diklik, menunggu download...")

        downloaded_file = self._wait_for_download(existing_files=existing_files, timeout=30)

        log.info(f"File berhasil diunduh: {downloaded_file}")
        return downloaded_file

    def _wait_for_download(self, existing_files: set = None, timeout=30) -> Path:
        """Tunggu sampai file BARU .xls/.xlsx muncul di folder download."""
        existing_files = existing_files or set()
        deadline = time.time() + timeout
        while time.time() < deadline:
            current_files = set(self.download_dir.glob("Kehadiran_Individu_*.xls*"))
            new_files = current_files - existing_files
            if new_files:
                newest = max(new_files, key=lambda f: f.stat().st_mtime)
                if not str(newest).endswith(".crdownload"):
                    return newest
            time.sleep(1)
        raise TimeoutError("File download tidak selesai dalam waktu yang ditentukan.")

    def close(self):
        if self.driver:
            self.driver.quit()
            log.info("Browser ditutup.")


# ─── Email Sender ─────────────────────────────────────────
class EmailSender:
    def __init__(self):
        self.smtp_host = CONFIG["SMTP_HOST"]
        self.smtp_port = CONFIG["SMTP_PORT"]
        self.smtp_user = CONFIG["SMTP_USER"]
        self.smtp_password = CONFIG["SMTP_PASSWORD"]
        self.from_email = CONFIG["FROM_EMAIL"]
        self.to_email = CONFIG["TO_EMAIL"]

    def send_report(self, subject: str, html_body: str, attachment_path: Path = None):
        msg = MIMEMultipart("alternative")
        msg["Subject"] = subject
        msg["From"] = self.from_email
        msg["To"] = self.to_email

        msg.attach(MIMEText(html_body, "html"))

        if attachment_path and attachment_path.exists():
            with open(attachment_path, "rb") as f:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition",
                f"attachment; filename={attachment_path.name}"
            )
            msg.attach(part)

        with smtplib.SMTP(self.smtp_host, self.smtp_port) as server:
            server.starttls()
            server.login(self.smtp_user, self.smtp_password)
            server.sendmail(self.from_email, self.to_email, msg.as_string())

        log.info(f"Email terkirim ke {self.to_email}")


# ─── Main Orchestrator ────────────────────────────────────
def run_monthly_check(year: int = None, month: int = None):
    if year is None or month is None:
        today = datetime.today()
        first_of_month = today.replace(day=1)
        last_month = first_of_month - timedelta(days=1)
        year = last_month.year
        month = last_month.month

    period_label = f"{year}-{month:02d}"
    log.info(f"════ Memulai pengecekan absensi periode {period_label} ════")

    scraper = HRPortalScraper()
    try:
        scraper._init_driver()
        scraper.login()
        scraper.navigate_to_attendance(year, month)
        xls_file = scraper.download_excel()
    except Exception as e:
        log.error(f"Gagal saat scraping: {e}\n{traceback.format_exc()}")
        raise
    finally:
        scraper.close()

    log.info("Menganalisis data absensi...")
    analyzer = AbsensiAnalyzer(xls_file)
    result = analyzer.analyze()

    subject = f"[Absensi] Laporan Kehadiran {period_label}"
    html_body = build_email_html(result, period_label)

    sender = EmailSender()
    sender.send_report(subject, html_body, attachment_path=xls_file)

    # Hapus file XLS setelah email terkirim
    try:
        xls_file.unlink()
        log.info(f"File {xls_file.name} dihapus setelah email terkirim.")
    except Exception as e:
        log.warning(f"Gagal hapus file: {e}")

    log.info(f"════ Selesai. Laporan dikirim untuk periode {period_label} ════")
    return result


def build_email_html(result: dict, period_label: str) -> str:
    def badge(text, color):
        return f'<span style="background:{color};color:white;padding:2px 8px;border-radius:4px;font-size:12px;">{text}</span>'

    warnings = result.get("warnings", [])
    warning_rows = ""
    for w in warnings:
        warning_rows += f"<tr><td>⚠️</td><td>{w}</td></tr>"
    if not warning_rows:
        warning_rows = "<tr><td colspan='2' style='color:green;'>✅ Tidak ada peringatan</td></tr>"

    late_rows = ""
    for item in result.get("late_days", []):
        late_rows += f"""
        <tr>
          <td>{item['tanggal']}</td>
          <td>{item['jam_masuk_standar']}</td>
          <td>{item['jam_masuk_aktual']}</td>
          <td>{item['telat_menit']} menit</td>
        </tr>"""
    if not late_rows:
        late_rows = "<tr><td colspan='4' style='color:green;'>Tidak ada keterlambatan</td></tr>"

    html = f"""
    <html><body style="font-family:Arial,sans-serif;max-width:700px;margin:auto;color:#333;">
      <h2 style="background:#1a73e8;color:white;padding:16px;border-radius:6px;">
        📋 Laporan Kehadiran Individu — {period_label}
      </h2>
      <table style="width:100%;border-collapse:collapse;margin-bottom:20px;">
        <tr><td style="padding:8px;width:50%;">
          <b>Noreg</b><br><span style="font-size:18px;">{result.get('noreg','—')}</span>
        </td><td style="padding:8px;">
          <b>Periode</b><br><span style="font-size:18px;">{period_label}</span>
        </td></tr>
      </table>
      <h3 style="border-bottom:2px solid #1a73e8;">📊 Ringkasan</h3>
      <table style="width:100%;border-collapse:collapse;">
        <tr style="background:#f1f3f4;">
          <td style="padding:10px;">Hari Kerja (PLA/shift)</td>
          <td style="padding:10px;text-align:right;font-weight:bold;">{result.get('total_workdays',0)} hari</td>
        </tr>
        <tr>
          <td style="padding:10px;">Hadir</td>
          <td style="padding:10px;text-align:right;color:green;font-weight:bold;">{result.get('hadir',0)} hari</td>
        </tr>
        <tr style="background:#f1f3f4;">
          <td style="padding:10px;">Mangkir</td>
          <td style="padding:10px;text-align:right;">
            {result.get('mangkir',0)} hari
            {"&nbsp;" + badge("⚠️ PERHATIAN","#e53935") if result.get('mangkir',0) > 0 else ""}
          </td>
        </tr>
        <tr>
          <td style="padding:10px;">Sakit</td>
          <td style="padding:10px;text-align:right;">
            {result.get('sakit',0)} hari
            {"&nbsp;" + badge("⚠️ > 6 HARI","#e53935") if result.get('sakit',0) > 6 else ""}
          </td>
        </tr>
        <tr style="background:#f1f3f4;">
          <td style="padding:10px;">Izin</td>
          <td style="padding:10px;text-align:right;">{result.get('izin',0)} hari</td>
        </tr>
        <tr>
          <td style="padding:10px;">Terlambat</td>
          <td style="padding:10px;text-align:right;">{result.get('total_late_days',0)} hari
            (total {result.get('total_late_minutes',0)} menit)
          </td>
        </tr>
      </table>
      <h3 style="border-bottom:2px solid #e53935;margin-top:24px;">🚨 Peringatan</h3>
      <table style="width:100%;border-collapse:collapse;">
        {warning_rows}
      </table>
      <h3 style="border-bottom:2px solid #f9a825;margin-top:24px;">🕐 Detail Keterlambatan</h3>
      <table style="width:100%;border-collapse:collapse;font-size:13px;">
        <tr style="background:#f1f3f4;font-weight:bold;">
          <th style="padding:8px;text-align:left;">Tanggal</th>
          <th style="padding:8px;text-align:left;">Standar Masuk</th>
          <th style="padding:8px;text-align:left;">Aktual Masuk</th>
          <th style="padding:8px;text-align:left;">Selisih</th>
        </tr>
        {late_rows}
      </table>
      <p style="color:#888;font-size:12px;margin-top:30px;">
        Email ini dikirim otomatis oleh AbsensiChecker 🤖 — {datetime.now().strftime('%d %b %Y %H:%M')}
      </p>
    </body></html>
    """
    return html


# ─── Entry Point ──────────────────────────────────────────
if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Cek absensi HR Portal")
    parser.add_argument("--year", type=int, help="Tahun (default: bulan lalu)")
    parser.add_argument("--month", type=int, help="Bulan 1-12 (default: bulan lalu)")
    parser.add_argument(
        "--analyze-only",
        type=str,
        metavar="FILE",
        help="Analisis file XLS lokal tanpa login (untuk testing)"
    )
    args = parser.parse_args()

    if args.analyze_only:
        from analyzer import AbsensiAnalyzer
        import glob

        target = args.analyze_only
        if "*" in target or target == "latest":
            pattern = target if "*" in target else "Kehadiran_Individu_*.xls*"
            matches = glob.glob(pattern)
            if not matches:
                print(f"❌ Tidak ada file yang cocok dengan pattern: {pattern}")
                exit(1)
            target = max(matches, key=os.path.getmtime)
            print(f"📂 File terbaru ditemukan: {target}")

        analyzer = AbsensiAnalyzer(Path(target))
        result = analyzer.analyze()
        print("\n" + "═"*50)
        print("HASIL ANALISIS ABSENSI")
        print("═"*50)
        for k, v in result.items():
            if k not in ("late_days",):
                print(f"  {k:30s}: {v}")
        print("\nKeterlambatan:")
        for item in result.get("late_days", []):
            print(f"  {item['tanggal']} → telat {item['telat_menit']} menit")
        print("\nPeringatan:")
        for w in result.get("warnings", []):
            print(f"  ⚠️  {w}")
    else:
        run_monthly_check(year=args.year, month=args.month)
