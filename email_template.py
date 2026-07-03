"""
Builder HTML untuk isi email laporan kehadiran bulanan.

CHANGELOG
  v1.0  Versi awal (icon-heavy, multi-warna)
  v1.1  Simplifikasi tampilan (single column, minim warna) + tombol
        "Buka HR Portal" yang link ke CONFIG["ATTENDANCE_PAGE_URL"]
  v1.2  Ganti tabel "Detail Keterlambatan" (cuma dari late_days) jadi
        "Catatan Kehadiran" yang pakai hadir_dengan_catatan — lebih
        lengkap karena mencakup lupa absen & telat/pulang cepat yang
        sudah diizinkan, bukan cuma telat tanpa kode. Tambah section
        "Koreksi Belum Disetujui" dari pending_corrections biar nggak
        kelewat di email
"""

__version__ = "1.2.0"

from datetime import datetime

from config import CONFIG

BORDER = "#e5e7eb"
GRAY = "#6b7280"
DARK = "#111827"
RED = "#b91c1c"
AMBER = "#b45309"


def _row(label, value, highlight=False):
    color = RED if highlight else DARK
    weight = "700" if highlight else "600"
    return f"""
        <tr>
          <td style="padding:8px 0;border-bottom:1px solid {BORDER};color:{GRAY};font-size:14px;">{label}</td>
          <td style="padding:8px 0;border-bottom:1px solid {BORDER};text-align:right;font-size:14px;font-weight:{weight};color:{color};">{value}</td>
        </tr>"""


def build_email_html(result: dict, period_label: str) -> str:
    mangkir = result.get("mangkir", 0)
    sakit = result.get("sakit", 0)
    portal_url = CONFIG.get("ATTENDANCE_PAGE_URL") or CONFIG.get("HR_PORTAL_URL", "#")

    warnings = result.get("warnings", [])
    if warnings:
        warning_html = "".join(f'<li style="margin-bottom:6px;">{w}</li>' for w in warnings)
        warning_block = f"""
        <div style="margin-top:24px;padding:14px 16px;background:#fef2f2;border:1px solid #fecaca;border-radius:6px;">
          <div style="font-size:13px;font-weight:700;color:{RED};margin-bottom:6px;">Peringatan</div>
          <ul style="margin:0;padding-left:18px;font-size:13px;color:{DARK};">{warning_html}</ul>
        </div>"""
    else:
        warning_block = ""

    catatan = result.get("hadir_dengan_catatan", [])
    if catatan:
        catatan_rows = "".join(
            f"""
            <tr>
              <td style="padding:6px 0;border-bottom:1px solid {BORDER};font-size:13px;white-space:nowrap;">{item['tanggal']}</td>
              <td style="padding:6px 0 6px 12px;border-bottom:1px solid {BORDER};font-size:13px;color:{GRAY};">{item['keterangan']}</td>
            </tr>"""
            for item in catatan
        )
        catatan_block = f"""
        <div style="margin-top:24px;">
          <div style="font-size:13px;font-weight:700;color:{DARK};margin-bottom:8px;">Catatan Kehadiran</div>
          <table role="presentation" width="100%" cellpadding="0" cellspacing="0">
            {catatan_rows}
          </table>
        </div>"""
    else:
        catatan_block = ""

    pending = result.get("pending_corrections", [])
    if pending:
        pending_rows = "".join(
            f"""
            <tr>
              <td style="padding:6px 0;border-bottom:1px solid {BORDER};font-size:13px;white-space:nowrap;">{item['tanggal']}</td>
              <td style="padding:6px 0 6px 12px;border-bottom:1px solid {BORDER};font-size:13px;color:{GRAY};">kode asli {item['kode_asli']} · status: {item['status']}</td>
            </tr>"""
            for item in pending
        )
        pending_block = f"""
        <div style="margin-top:24px;padding:14px 16px;background:#fffbeb;border:1px solid #fde68a;border-radius:6px;">
          <div style="font-size:13px;font-weight:700;color:{AMBER};margin-bottom:8px;">Koreksi Belum Disetujui</div>
          <table role="presentation" width="100%" cellpadding="0" cellspacing="0">
            {pending_rows}
          </table>
        </div>"""
    else:
        pending_block = ""

    html = f"""
    <html>
    <body style="margin:0;padding:24px;background:#f5f5f5;font-family:-apple-system,Helvetica,Arial,sans-serif;color:{DARK};">
      <table role="presentation" width="100%" cellpadding="0" cellspacing="0">
        <tr>
          <td align="center">
            <table role="presentation" width="560" cellpadding="0" cellspacing="0"
                   style="background:#ffffff;border:1px solid {BORDER};border-radius:8px;padding:28px;">

              <tr>
                <td style="font-size:12px;color:{GRAY};text-transform:uppercase;letter-spacing:.04em;">
                  Laporan Kehadiran · {period_label}
                </td>
              </tr>
              <tr>
                <td style="font-size:18px;font-weight:700;padding-top:4px;padding-bottom:20px;">
                  Noreg {result.get('noreg', '—')}
                </td>
              </tr>

              <tr>
                <td>
                  <table role="presentation" width="100%" cellpadding="0" cellspacing="0">
                    {_row("Hari kerja", f"{result.get('total_workdays', 0)} hari")}
                    {_row("Hadir", f"{result.get('hadir', 0)} hari")}
                    {_row("Mangkir", f"{mangkir} hari", highlight=mangkir > 0)}
                    {_row("Sakit", f"{sakit} hari", highlight=sakit > 6)}
                    {_row("Izin", f"{result.get('izin', 0)} hari")}
                    {_row("Terlambat", f"{result.get('total_late_days', 0)} hari ({result.get('total_late_minutes', 0)} menit)")}
                  </table>
                </td>
              </tr>

              {warning_block}
              {catatan_block}
              {pending_block}

              <tr>
                <td style="padding-top:28px;">
                  <a href="{portal_url}"
                     style="display:inline-block;padding:12px 24px;background:{DARK};color:#ffffff;
                            font-size:14px;font-weight:600;text-decoration:none;border-radius:6px;">
                    Buka HR Portal
                  </a>
                </td>
              </tr>

              <tr>
                <td style="padding-top:24px;font-size:11px;color:{GRAY};">
                  Dikirim otomatis oleh AbsensiChecker · {datetime.now().strftime('%d %b %Y %H:%M')}
                </td>
              </tr>

            </table>
          </td>
        </tr>
      </table>
    </body>
    </html>
    """
    return html