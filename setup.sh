#!/bin/bash
# setup.sh — Jalankan SEKALI untuk setup di VPS
# chmod +x setup.sh && ./setup.sh

set -e
echo "════════════════════════════════════════"
echo "  Setup AbsensiChecker di VPS"
echo "════════════════════════════════════════"

# 1. Install system dependencies
echo "[1/5] Install system packages..."
sudo apt-get update -qq
sudo apt-get install -y -qq \
    python3 python3-pip python3-venv \
    chromium-browser chromium-chromedriver \
    cron

# 2. Buat virtual environment
echo "[2/5] Buat Python virtual environment..."
python3 -m venv venv
source venv/bin/activate

# 3. Install Python packages
echo "[3/5] Install Python packages..."
pip install -r requirements.txt --quiet

# 4. Buat folder downloads
echo "[4/5] Buat folder downloads..."
mkdir -p downloads
mkdir -p logs

# 5. Setup .env
if [ ! -f .env ]; then
    cp .env.example .env
    echo ""
    echo "⚠️  File .env sudah dibuat dari template."
    echo "    Edit file .env dan isi kredensial kamu!"
    echo "    nano .env"
else
    echo "[5/5] File .env sudah ada, skip."
fi

echo ""
echo "✅ Setup selesai!"
echo ""
echo "════ LANGKAH SELANJUTNYA ════════════════"
echo ""
echo "1. Edit konfigurasi:"
echo "   nano .env"
echo ""
echo "2. Test analisis file lokal (tanpa login):"
echo "   source venv/bin/activate"
echo "   python absensi_checker.py --analyze-only Kehadiran_Individu_*.xls"
echo ""
echo "3. Test run lengkap (dengan login):"
echo "   python absensi_checker.py --year 2026 --month 3"
echo ""
echo "4. Pasang cron job (otomatis tiap tanggal 2, jam 08:00):"
echo "   crontab -e"
echo "   Tambahkan baris:"
echo "   0 8 2 * * /home/absensi_checker/run_monthly.sh >> /home/absensi_checker/logs/cron.log 2>&1"
echo ""
