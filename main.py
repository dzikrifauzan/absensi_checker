"""
╔══════════════════════════════════════════════════════════╗
║         AUTOMATION CEK ABSENSI HR PORTAL                 ║
║  - Login ke HR Portal via Selenium                       ║
║  - Download laporan Kehadiran Individu (XLS)             ║
║  - Analisis: Mangkir, Sakit > 6 hari, Keterlambatan       ║
║  - Kirim email ringkasan bulanan                          ║
╚══════════════════════════════════════════════════════════╝

Entry point utama. Jalankan dengan:
    python main.py --year 2026 --month 6
    python main.py --analyze-only latest
    python main.py --version

CHANGELOG
  v1.4  Split awal dari absensi_checker.py monolitik (ikut versi analyzer)
  v1.5  --version flag + print_versions(). Rapiin _print_analysis_result
        supaya hadir_dengan_catatan & pending_corrections tampil terformat,
        bukan raw dict dump
"""

__version__ = "1.5.0"

import os
import traceback
from datetime import datetime, timedelta
from pathlib import Path

import config as config_module
import analyzer as analyzer_module
import scraper as scraper_module
import email_sender as email_sender_module
import email_template as email_template_module
import logging_setup as logging_setup_module
from analyzer import AbsensiAnalyzer
from scraper import HRPortalScraper
from email_sender import EmailSender
from email_template import build_email_html
from logging_setup import log


def print_versions():
    print(f"AbsensiChecker v{__version__}")
    for mod in (config_module, analyzer_module, scraper_module, email_sender_module,
                email_template_module, logging_setup_module):
        print(f"  - {mod.__name__:16s} v{getattr(mod, '__version__', '?')}")


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


def _print_analysis_result(result: dict):
    detail_keys = ("late_days", "hadir_dengan_catatan", "pending_corrections", "no_scan_days")

    print("\n" + "═" * 50)
    print("HASIL ANALISIS ABSENSI")
    print("═" * 50)
    for k, v in result.items():
        if k not in detail_keys:
            print(f"  {k:30s}: {v}")

    print("\nCatatan Kehadiran:")
    catatan = result.get("hadir_dengan_catatan", [])
    if catatan:
        for item in catatan:
            print(f"  {item['tanggal']} → {item['keterangan']}")
    else:
        print("  (tidak ada)")

    pending = result.get("pending_corrections", [])
    if pending:
        print("\nKoreksi Belum Disetujui:")
        for item in pending:
            print(f"  {item['tanggal']} → kode asli {item['kode_asli']}, status: {item['status']}")

    print("\nPeringatan:")
    for w in result.get("warnings", []):
        print(f"  ⚠️  {w}")


def _run_analyze_only(target: str):
    import glob

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
    _print_analysis_result(result)


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
    parser.add_argument(
        "--version",
        action="store_true",
        help="Tampilkan versi AbsensiChecker & tiap modul, lalu keluar"
    )
    args = parser.parse_args()

    if args.version:
        print_versions()
    elif args.analyze_only:
        _run_analyze_only(args.analyze_only)
    else:
        run_monthly_check(year=args.year, month=args.month)