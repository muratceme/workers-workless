"""
belge.py — Workers / Workless ortak belge okuma katmanı

PDF, DOCX ve TXT dosyalarından düz metin çıkarır (kaynağı: library/_ortak/belge.py; paket klasörlerindeki
kopyalar CI tarafından bununla aynı tutulur). DOCX ek kütüphane olmadan okunur; PDF için pypdf gerekir.
"""
from __future__ import annotations

import logging
import zipfile
from pathlib import Path
from xml.etree import ElementTree

DESTEKLENEN = {".pdf", ".docx", ".txt", ".md"}


class BelgeHatasi(RuntimeError):
    """Kullanıcıya gösterilecek, anlaşılır hata."""


def pdf_oku(yol: Path) -> str:
    try:
        from pypdf import PdfReader
    except ImportError as h:
        raise BelgeHatasi("PDF okumak için pypdf gerekli: pip install -r requirements.txt") from h
    logging.getLogger("pypdf").setLevel(logging.ERROR)
    return "\n".join((sayfa.extract_text() or "") for sayfa in PdfReader(str(yol)).pages)


def docx_oku(yol: Path) -> str:
    """DOCX bir ZIP + XML paketidir; paragraflar ve tablo hücreleri sırasıyla okunur."""
    ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    with zipfile.ZipFile(yol) as z:
        kok = ElementTree.fromstring(z.read("word/document.xml"))
    satirlar = []
    for p in kok.iter(f"{ns}p"):
        satirlar.append("".join(t.text or "" for t in p.iter(f"{ns}t")))
    return "\n".join(satirlar)


def txt_oku(yol: Path) -> str:
    for kodlama in ("utf-8-sig", "cp1254", "latin-1"):
        try:
            return yol.read_text(encoding=kodlama)
        except UnicodeDecodeError:
            continue
    return ""


def metin_oku(yol: Path) -> str:
    uzanti = yol.suffix.lower()
    if uzanti == ".doc":
        raise BelgeHatasi(f"{yol.name}: eski .doc biçimi desteklenmiyor; Word'de 'Farklı Kaydet → .docx' yapın.")
    if uzanti not in DESTEKLENEN:
        raise BelgeHatasi(f"{yol.name}: desteklenmeyen dosya türü ({uzanti}). Desteklenenler: {', '.join(sorted(DESTEKLENEN))}")
    if uzanti == ".pdf":
        metin = pdf_oku(yol)
        if len(metin.strip()) < 50:
            raise BelgeHatasi(f"{yol.name}: PDF'ten metin çıkarılamadı (taranmış görüntü olabilir; önce OCR uygulayın).")
        return metin
    if uzanti == ".docx":
        return docx_oku(yol)
    return txt_oku(yol)
