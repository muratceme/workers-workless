#!/bin/bash
# Gerisi Bende - macOS / Linux başlatıcısı.
# macOS: dosyaya sağ tıklayıp "Aç"ı seçin (ilk seferde güvenlik onayı ister).
# Çalışmazsa Terminal'de:  bash BASLAT.command
cd "$(dirname "$0")" || exit 1
PY=""
for c in python3 python; do
  if command -v "$c" >/dev/null 2>&1 && "$c" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)' 2>/dev/null; then
    PY="$c"; break
  fi
done
if [ -z "$PY" ]; then
  echo "Bu paketi çalıştırmak için Python 3.10 veya üstü gerekiyor ve bulunamadı."
  echo "https://www.python.org/downloads/ adresinden kurup bu dosyayı yeniden açın."
  if command -v open >/dev/null 2>&1; then open "https://www.python.org/downloads/"; fi
  read -r -p "Kapatmak için Enter'a basın..."
  exit 1
fi
"$PY" _baslat.py
