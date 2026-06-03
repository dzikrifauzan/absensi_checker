#!/bin/bash
# run_monthly.sh — Dijalankan oleh cron tiap bulan
# Otomatis load .env dan jalankan checker

# Lokasi project (sesuaikan jika beda)
PROJECT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$PROJECT_DIR"

# Load environment variables dari .env
if [ -f .env ]; then
    export $(grep -v '^#' .env | xargs)
fi

# Aktifkan virtual environment
source venv/bin/activate

# Jalankan checker (default: bulan lalu)
echo "[$(date '+%Y-%m-%d %H:%M:%S')] Menjalankan AbsensiChecker..."
python absensi_checker.py

EXIT_CODE=$?
if [ $EXIT_CODE -eq 0 ]; then
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] Selesai dengan sukses."
else
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] ERROR: Exit code $EXIT_CODE"
fi

exit $EXIT_CODE
