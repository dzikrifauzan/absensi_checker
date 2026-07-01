"""
analyzer.py — Parser & Analyzer laporan Kehadiran Individu dari HR Portal.

Struktur file XLS yang diharapkan (sesuai sample Kehadiran_Individu_*.xls):
  - Row 7  : Periode
  - Row 8  : Noreg
  - Row 10 : Header grup kolom
  - Row 11 : Sub-header
  - Row 12 : Sub-sub-header
  - Row 13+ : Data per tanggal

Kolom yang dipakai (0-indexed, dikonfirmasi dari file asli):
  col[0]  : Tanggal
  col[1]  : Kode Shift (FRE1=libur, PLA1=kerja, dll)
  col[2]  : Jam Masuk Standar
  col[3]  : Jam Keluar Standar
  col[4]  : Jam Masuk ARS (aktual, sebelum koreksi)
  col[5]  : Jam Keluar ARS (aktual, sebelum koreksi)
  col[6]  : Kode Masuk (sebelum koreksi)
  col[7]  : Kode Keluar (sebelum koreksi)
  col[8]  : Koreksi Jam Masuk (kalau ada pengajuan koreksi)
  col[9]  : Koreksi Jam Keluar
  col[10] : Koreksi Kode Masuk
  col[11] : Koreksi Kode Keluar
  col[12] : Alasan / Keterangan koreksi
  col[13] : Status Persetujuan koreksi (mis. "Approved by DpH")

Kalau koreksi ada dan sudah Approved, kode & jam ARS pada baris tsb
DIGANTIKAN oleh nilai koreksi sebelum dievaluasi (lihat analyze()).
"""

import re
import logging
from pathlib import Path
from datetime import datetime, time

import pandas as pd

log = logging.getLogger(__name__)

# ─── Kode Absensi Toyota HR Portal ───────────────────────
KODE_ABSENSI = {
    # PERMOHONAN (Pengajuan Sebelum/Terencana)
    "02" : "Ijin terlambat kurang 3 jam",
    "03" : "Ijin terlambat lebih 3 jam",
    "04" : "Ijin pulang cepat kurang 3 jam",
    "05" : "Ijin pulang cepat lebih 3 jam",
    "06" : "Ijin meninggalkan kerja",
    "07" : "Cuti tahunan",
    "14" : "Ijin pernikahan sendiri",
    "15" : "Ijin menikahkan anak",
    "16" : "Ijin khitan/baptis anak",
    "17" : "Ijin tugas negara",
    "18" : "Ijin Kewajiban Ibadah (Haji)",
    "26" : "Ijin tidak masuk kerja tanpa upah",
    "32" : "Training",
    "33" : "Dinas luar/Perjalanan luar kota",
    "45" : "MPP (Masa Persiapan Pensiun)",
    "46" : "Pulang Cepat dengan Instruksi",
    "47" : "Datang Terlambat dengan Instruksi",
    "48" : "Kerja dari Rumah",

    # PEMBERITAHUAN (Pengajuan Setelah Masuk Kerja)
    "07A": "Cuti Tahunan",
    "11" : "Sakit",
    "12" : "Sakit kecelakaan kerja",
    "13" : "Sakit opname/CDP/keguguran",
    "19" : "Ijin ortu/mertua sakit (opname)",
    "20" : "Istri melahirkan/keguguran",
    "21" : "Istri/Suami/Anak sakit keras (opname)",
    "22" : "Istri/Suami/Anak/Orang Tua meninggal dunia",
    "23" : "Mertua/Menantu meninggal dunia",
    "24" : "Saudara kandung/Orang serumah meninggal dunia",
    "25" : "Kakek/Nenek/Cucu/Ipar meninggal dunia",
    "40" : "Tidak masuk kerja karena banjir",
    "44" : "Lupa Absen",
    "32A": "Training",
    "33A": "Dinas luar/Perjalanan luar kota",
    "45A": "MPP (Masa Persiapan Pensiun)",
    "26A": "Ijin tidak masuk kerja tanpa upah",
    "46A": "Pulang Cepat dengan Instruksi",
    "47A": "Datang Terlambat dengan Instruksi",
    "48A": "Kerja dari Rumah",

    # Kode hadir normal
    "01" : "Hadir",

    # Kode mangkir
    "37" : "Alpha / Tidak masuk tanpa keterangan",

    # Kode shift
    "FRE1": "Hari Libur / Day Off",
    "PLA1": "Shift Standar (Kerja)",
}

# Kode MANGKIR — tidak masuk tanpa keterangan
# (tidak ada kode khusus di portal, ditandai dengan tidak ada scan & tidak ada kode)
KODE_MANGKIR = {"37"}  # Alpha / tidak masuk tanpa keterangan

# Kode SAKIT
KODE_SAKIT = {"11", "12", "13"}

# Kode IZIN (berbagai jenis izin resmi)
KODE_IZIN = {
    "02", "03", "04", "05", "06",
    "07", "07A",                        # cuti
    "14", "15", "16", "17", "18",
    "19", "20", "21", "22", "23", "24", "25",
    "32", "32A", "33", "33A",          # training/dinas
    "40", "44", "45", "45A",
    "46", "46A", "47", "47A", "48", "48A",
}

# Kode TERLAMBAT dengan instruksi (sudah diizinkan, tidak dihitung pelanggaran)
KODE_TERLAMBAT_IZIN = {"02", "03", "47", "47A"}

# Kode shift hari LIBUR (tidak dihitung hari kerja)
KODE_LIBUR_SHIFT = {"FRE1", "H", "LIB"}


def parse_time_str(val):
    """
    Parse waktu dari cell XLS menjadi datetime.time.

    Menangani beberapa bentuk representasi:
    - Sudah berupa datetime.time (langsung dipakai)
    - Berupa datetime.datetime / pandas.Timestamp — termasuk kasus umum
      xlrd/pandas yang membaca cell bertipe "time" sebagai datetime dengan
      tanggal dummy, misal '1899-12-30 07:15:00'
    - String 'HH:MM' atau 'HH:MM:SS', termasuk yang diawali tanggal dummy
      di atas (dicari dengan re.search, bukan re.match, supaya tidak gagal
      hanya karena polanya tidak persis di awal string)
    """
    if pd.isna(val) or val == "" or val is None:
        return None

    if isinstance(val, time):
        return val
    if isinstance(val, (datetime, pd.Timestamp)):
        return val.time()

    s = str(val).strip()
    # Cari pola HH:MM di mana saja dalam string, bukan cuma di awal,
    # supaya string semacam '1899-12-30 07:15:00' tetap berhasil di-parse.
    match = re.search(r"(\d{1,2}):(\d{2})", s)
    if match:
        h, m = int(match.group(1)), int(match.group(2))
        if 0 <= h <= 23 and 0 <= m <= 59:
            return time(h, m)
    return None


def time_diff_minutes(t1: time, t2: time) -> int:
    """Hitung selisih waktu dalam menit (t2 - t1). Positif = t2 lebih lambat."""
    dt1 = datetime.combine(datetime.today(), t1)
    dt2 = datetime.combine(datetime.today(), t2)
    return int((dt2 - dt1).total_seconds() / 60)


class AbsensiAnalyzer:
    def __init__(self, xls_path: Path):
        self.xls_path = Path(xls_path)
        self.df_raw = None
        self.info = {}

    def _load(self):
        """Load file XLS mentah."""
        log.info(f"Membaca file: {self.xls_path}")
        self.df_raw = pd.read_excel(
            self.xls_path,
            engine="xlrd",
            header=None
        )

    def _extract_header_info(self):
        """Ambil metadata: periode, noreg dari baris header."""
        df = self.df_raw

        def clean_val(raw):
            # Strip spasi dan titik dua di depan, misal ': 02437272' -> '02437272'
            return re.sub(r"^[\s:]+", "", str(raw)).strip()

        # Periode ada di baris index 7, kolom 1
        try:
            periode_raw = clean_val(df.iloc[7, 1])
            m = re.search(r"(\d{4})-(\d{2})", periode_raw)
            if m:
                self.info["tahun"] = int(m.group(1))
                self.info["bulan"] = int(m.group(2))
            self.info["periode_raw"] = periode_raw
        except Exception as e:
            log.warning(f"Gagal baca periode: {e}")
            self.info["periode_raw"] = "—"

        # Noreg di baris index 8, kolom 1
        try:
            self.info["noreg"] = clean_val(df.iloc[8, 1])
        except Exception:
            self.info["noreg"] = "—"

        log.info(f"Noreg: {self.info.get('noreg')}, Periode: {self.info.get('periode_raw')}")

    def _extract_data_rows(self) -> pd.DataFrame:
        """
        Ambil baris data (mulai dari row 13 ke bawah).
        Return DataFrame dengan kolom yang sudah diberi nama.
        """
        df = self.df_raw

        # Data mulai dari baris index 13
        data = df.iloc[13:].copy()
        data = data.reset_index(drop=True)

        # Rename kolom sesuai posisi (berdasarkan struktur file sample)
        col_map = {
            0: "tanggal",
            1: "shift_kode",
            2: "std_masuk",           # Jam masuk standar
            3: "std_keluar",          # Jam keluar standar
            4: "ars_masuk",           # Jam masuk aktual (mesin absen), sebelum koreksi
            5: "ars_keluar",          # Jam keluar aktual, sebelum koreksi
            6: "kode_masuk",          # Kode absensi masuk, sebelum koreksi
            7: "kode_keluar",         # Kode absensi keluar, sebelum koreksi
            8: "koreksi_ars_masuk",   # Jam masuk hasil koreksi (kalau ada pengajuan)
            9: "koreksi_ars_keluar",  # Jam keluar hasil koreksi
            10: "koreksi_kode_masuk",  # Kode masuk hasil koreksi
            11: "koreksi_kode_keluar",  # Kode keluar hasil koreksi
            12: "alasan",             # Alasan pengajuan koreksi
            13: "status_koreksi",     # Status persetujuan koreksi
        }
        # Hanya ambil kolom yang relevan
        available_cols = [c for c in col_map.keys() if c < len(data.columns)]
        data = data[available_cols].rename(columns=col_map)

        # Buang baris kosong (tanggal NaN)
        data = data[data["tanggal"].notna()].copy()
        data = data[data["tanggal"].astype(str).str.strip() != "nan"].copy()

        return data

    def analyze(self) -> dict:
        """
        Analisis utama:
        - Hitung hari kerja, hadir, mangkir, sakit, izin
        - Deteksi keterlambatan
        - Generate warnings
        """
        self._load()
        self._extract_header_info()
        rows = self._extract_data_rows()

        total_workdays = 0
        hadir = 0
        mangkir = 0
        sakit = 0
        izin = 0
        late_days = []
        no_scan_days = []
        pending_corrections = []  # koreksi yang belum di-approve, dilaporkan terpisah

        def _clean(val):
            s = str(val).strip().upper()
            return "" if s in ("", "NAN", "NONE") else s

        for _, row in rows.iterrows():
            tanggal   = str(row.get("tanggal", "")).strip()
            shift     = str(row.get("shift_kode", "")).strip().upper()
            std_in    = parse_time_str(row.get("std_masuk"))
            kode_in_raw  = _clean(row.get("kode_masuk", ""))
            kode_out_raw = _clean(row.get("kode_keluar", ""))
            ars_in_raw   = parse_time_str(row.get("ars_masuk"))

            koreksi_kode_in  = _clean(row.get("koreksi_kode_masuk", ""))
            koreksi_kode_out = _clean(row.get("koreksi_kode_keluar", ""))
            koreksi_ars_in   = parse_time_str(row.get("koreksi_ars_masuk"))
            status_koreksi   = _clean(row.get("status_koreksi", ""))
            alasan           = str(row.get("alasan", "")).strip()
            alasan           = "" if alasan.lower() == "nan" else alasan

            ada_koreksi = bool(koreksi_kode_in or koreksi_kode_out)
            koreksi_approved = ada_koreksi and "APPROV" in status_koreksi

            if ada_koreksi and not koreksi_approved:
                # Ada pengajuan koreksi tapi belum disetujui — jangan langsung
                # dipakai, tapi jangan didiamkan juga. Tetap pakai kode asli
                # untuk perhitungan, dan laporkan sebagai catatan terpisah.
                pending_corrections.append({
                    "tanggal": tanggal,
                    "kode_asli": kode_in_raw or kode_out_raw,
                    "status": status_koreksi or "Belum diketahui",
                    "alasan": alasan,
                })

            if koreksi_approved:
                kode_in, kode_out = koreksi_kode_in, koreksi_kode_out
                ars_in = koreksi_ars_in if koreksi_ars_in is not None else ars_in_raw
            else:
                kode_in, kode_out = kode_in_raw, kode_out_raw
                ars_in = ars_in_raw

            # Skip hari libur (shift FRE1, H, dll)
            if shift in KODE_LIBUR_SHIFT:
                continue

            total_workdays += 1

            # Tentukan status kehadiran hari ini
            kode = kode_in if kode_in else kode_out

            if kode in KODE_SAKIT:
                sakit += 1
                continue

            if kode in KODE_IZIN:
                izin += 1
                continue

            if kode in KODE_MANGKIR:
                mangkir += 1
                continue

            # Tidak ada scan & tidak ada kode izin/sakit → kemungkinan mangkir
            if ars_in is None and kode not in ("01",) and kode not in KODE_IZIN and kode not in KODE_SAKIT:
                no_scan_days.append(tanggal)
                mangkir += 1
                continue

            hadir += 1

            # Cek keterlambatan
            # Skip kode terlambat yang sudah diizinkan (02, 03, 47, 47A)
            if kode in KODE_TERLAMBAT_IZIN:
                continue

            if std_in and ars_in:
                MAX_LATE_MINUTES = 240  # lebih dari 4 jam = anomali scan
                diff = time_diff_minutes(std_in, ars_in)
                if diff > MAX_LATE_MINUTES:
                    import logging as _l
                    _l.getLogger(__name__).warning(
                        f"{tanggal}: scan masuk {ars_in.strftime('%H:%M')} "
                        f"terlalu jauh ({diff} menit) — kemungkinan anomali, dilewati"
                    )
                elif diff > 0:
                    late_days.append({
                        "tanggal": tanggal,
                        "jam_masuk_standar": std_in.strftime("%H:%M"),
                        "jam_masuk_aktual": ars_in.strftime("%H:%M"),
                        "telat_menit": diff
                    })

        total_late_minutes = sum(d["telat_menit"] for d in late_days)

        # ── Generate Warnings ──────────────────────────────
        warnings = []

        if mangkir > 0:
            detail = ""
            if no_scan_days:
                detail = f" (tidak scan: {', '.join(no_scan_days)})"
            warnings.append(
                f"Terdapat {mangkir} hari MANGKIR/Alpha{detail}"
            )

        if sakit > 6:
            warnings.append(
                f"Sakit {sakit} hari — melebihi batas 6 hari (perlu surat dokter/verifikasi HR)"
            )

        if len(late_days) >= 3:
            warnings.append(
                f"Terlambat {len(late_days)} kali (total {total_late_minutes} menit)"
            )

        if pending_corrections:
            tgl_list = ", ".join(p["tanggal"] for p in pending_corrections)
            warnings.append(
                f"Ada {len(pending_corrections)} pengajuan koreksi yang BELUM di-approve "
                f"({tgl_list}) — kode asli masih dipakai sampai disetujui HR"
            )

        result = {
            "noreg"              : self.info.get("noreg", "—"),
            "periode"            : self.info.get("periode_raw", "—"),
            "total_workdays"     : total_workdays,
            "hadir"              : hadir,
            "mangkir"            : mangkir,
            "sakit"              : sakit,
            "izin"               : izin,
            "total_late_days"    : len(late_days),
            "total_late_minutes" : total_late_minutes,
            "late_days"          : late_days,
            "no_scan_days"       : no_scan_days,
            "pending_corrections": pending_corrections,
            "warnings"           : warnings,
        }

        log.info(
            f"Hasil: workdays={total_workdays}, hadir={hadir}, "
            f"mangkir={mangkir}, sakit={sakit}, izin={izin}, "
            f"terlambat={len(late_days)}x"
        )
        return result