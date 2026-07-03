"""
Scraper Selenium untuk login & download laporan Kehadiran Individu
dari HR Portal.
"""

__version__ = "1.0.0"

import time
from pathlib import Path

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options

from config import CONFIG
from logging_setup import log


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