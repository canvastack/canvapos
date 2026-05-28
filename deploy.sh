#!/bin/bash
# Deploy CanvaPOS ke Google Apps Script via clasp
# Usage: ./deploy.sh [env]
#   env: production | staging | development (default: production)

set -e

ENV="${1:-production}"
echo "🚀 Deploy CanvaPOS ke environment: $ENV"

# Validasi
if ! command -v clasp &> /dev/null; then
    echo "❌ clasp tidak terinstall. Install: npm install -g @google/clasp"
    exit 1
fi

if [ ! -f ".clasp.json" ]; then
    echo "❌ .clasp.json tidak ditemukan. Buat dulu dengan scriptId."
    exit 1
fi

# Cek appsscript.json
if [ ! -f "src/appsscript.json" ]; then
    echo "❌ src/appsscript.json tidak ditemukan."
    exit 1
fi

# Push
echo "📦 Push kode ke Google Apps Script..."
clasp push

echo "✅ Deploy selesai! Buka spreadsheet:"
echo "   clasp open"
