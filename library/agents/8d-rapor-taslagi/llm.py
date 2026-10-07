"""
llm.py — Workers / Workless ortak yapay zekâ sağlayıcı katmanı

Tüm agent'lar bu dosyayı kullanır (kaynağı: library/_ortak/llm.py; agent klasörlerindeki
kopyalar CI tarafından bununla aynı tutulur).

Ayarlar .env dosyasından okunur:
    WW_PROVIDER   anthropic (varsayılan) | openai | ollama
    WW_MODEL      model adı (anthropic için varsayılan: claude-opus-5-5)
    WW_EFFORT     anthropic için düşünme/çaba düzeyi: low | medium | high (varsayılan: medium)
    ANTHROPIC_API_KEY / OPENAI_API_KEY   ilgili sağlayıcının anahtarı
    OLLAMA_HOST   yerel Ollama adresi (varsayılan: http://localhost:11434)
"""
from __future__ import annotations

import json
import os
import re
import sys
import urllib.error
import urllib.request

VARSAYILAN_ANTHROPIC_MODEL = "claude-opus-5-5"


class LLMHatasi(RuntimeError):
    """Kullanıcıya gösterilecek, anlaşılır hata."""


KULLANIM = {"istek": 0, "girdi_token": 0, "cikti_token": 0}


def ayarlar() -> dict:
    saglayici = os.getenv("WW_PROVIDER", "anthropic").strip().lower()
    if saglayici not in {"anthropic", "openai", "ollama"}:
        raise LLMHatasi(f"WW_PROVIDER '{saglayici}' tanınmadı. anthropic, openai veya ollama yazın.")
    model = os.getenv("WW_MODEL", "").strip()
    if not model:
        if saglayici != "anthropic":
            raise LLMHatasi(f"{saglayici} için .env dosyasında WW_MODEL belirtmelisiniz.")
        model = VARSAYILAN_ANTHROPIC_MODEL
    return {"saglayici": saglayici, "model": model, "effort": os.getenv("WW_EFFORT", "medium").strip().lower()}


# ----------------------------------------------------------------------------
# Kişisel veri maskeleme (KVKK) — yardımcıdır, garanti değildir
# ----------------------------------------------------------------------------

_MASKELER = [
    (re.compile(r"\bTR\d{2}(?:\s?\d{4}){5}\s?\d{2}\b", re.I), "[IBAN]"),
    (re.compile(r"[\w.+-]+@[\w-]+(?:\.[\w-]+)+"), "[E-POSTA]"),
    (re.compile(r"(?:\+?90[\s-]?)?\(?0?5\d{2}\)?[\s.-]?\d{3}[\s.-]?\d{2}[\s.-]?\d{2}"), "[TELEFON]"),
    # Sabit hat / 0850: başında 0 veya +90 şart (sipariş no gibi 10 haneli sayılar maskelenmesin)
    (re.compile(r"(?:\+90[\s-]?|\b0)\(?[2-48]\d{2}\)?[\s.-]?\d{3}[\s.-]?\d{2}[\s.-]?\d{2}\b"), "[TELEFON]"),
    (re.compile(r"\b[1-9]\d{10}\b"), "[TCKN]"),
    (re.compile(r"(?:https?://)?(?:www\.)?linkedin\.com/in/[\w\-%]+/?", re.I), "[LINKEDIN]"),
]


def maskele(metin: str) -> str:
    for desen, yerine in _MASKELER:
        metin = desen.sub(yerine, metin)
    return metin


# ----------------------------------------------------------------------------
# Onay
# ----------------------------------------------------------------------------

def onay_al(aciklama: str, evet: bool = False) -> None:
    """Veri dışarı gönderilmeden önce kullanıcıdan açık onay ister."""
    a = ayarlar()
    anahtar = {"anthropic": "ANTHROPIC_API_KEY", "openai": "OPENAI_API_KEY"}.get(a["saglayici"])
    if anahtar and not os.getenv(anahtar, "").strip():
        raise LLMHatasi(f"{anahtar} bulunamadı. .env.example dosyasını .env olarak kopyalayıp kendi anahtarınızı yazın.")
    hedef = "bilgisayarınızdaki Ollama" if a["saglayici"] == "ollama" else f"{a['saglayici']} API'sine"
    print(f"\n[?] {aciklama}\n    Gönderilecek yer: {hedef} · model: {a['model']}")
    if a["saglayici"] != "ollama":
        print("    API ücretleri ve gönderilen verinin uygunluğu (KVKK) sizin sorumluluğunuzdadır.")
    if evet:
        print("    --evet verildi, onay sorulmadan devam ediliyor.\n")
        return
    if not sys.stdin or not sys.stdin.isatty():
        raise LLMHatasi("Etkileşimsiz çalışmada onay için --evet parametresini verin.")
    if input("    Devam edilsin mi? [e/H] ").strip().lower() not in {"e", "evet", "y", "yes"}:
        raise LLMHatasi("Kullanıcı iptal etti. Hiçbir veri gönderilmedi.")
    print()


# ----------------------------------------------------------------------------
# Yapılandırılmış (JSON şemalı) istek
# ----------------------------------------------------------------------------

def json_iste(sistem: str, kullanici: str, sema: dict, max_tokens: int = 16000) -> dict:
    """Modelden `sema` ile uyumlu bir JSON nesnesi ister ve sözlük olarak döndürür."""
    a = ayarlar()
    KULLANIM["istek"] += 1
    if a["saglayici"] == "anthropic":
        return _anthropic(a, sistem, kullanici, sema, max_tokens)
    if a["saglayici"] == "openai":
        return _openai(a, sistem, kullanici, sema, max_tokens)
    return _ollama(a, sistem, kullanici, sema)


def _anthropic(a, sistem, kullanici, sema, max_tokens):
    try:
        import anthropic
    except ImportError as h:
        raise LLMHatasi("anthropic paketi kurulu değil: pip install -r requirements.txt") from h
    if not os.getenv("ANTHROPIC_API_KEY"):
        raise LLMHatasi("ANTHROPIC_API_KEY bulunamadı. .env dosyasına kendi anahtarınızı ekleyin.")

    client = anthropic.Anthropic()
    try:
        yanit = client.beta.messages.create(
            model=a["model"],
            max_tokens=max_tokens,
            system=sistem,
            messages=[{"role": "user", "content": kullanici}],
            output_config={"effort": a["effort"], "format": {"type": "json_schema", "schema": sema}},
            # Güvenlik sınıflandırıcısı isteği reddederse sunucu tarafında önerilen modele düşer
            betas=["server-side-fallback-2026-07-01"],
            fallbacks="default",
        )
    except anthropic.AuthenticationError as h:
        raise LLMHatasi("ANTHROPIC_API_KEY geçersiz.") from h
    except anthropic.PermissionDeniedError as h:
        raise LLMHatasi("API anahtarının bu modele erişim izni yok.") from h
    except anthropic.NotFoundError as h:
        raise LLMHatasi(f"Model bulunamadı: {a['model']} (WW_MODEL ayarını kontrol edin).") from h
    except anthropic.RateLimitError as h:
        raise LLMHatasi("Hız/kota sınırına takıldınız. Biraz bekleyip tekrar deneyin.") from h
    except anthropic.BadRequestError as h:
        raise LLMHatasi(f"İstek reddedildi: {h.message}") from h
    except anthropic.APIStatusError as h:
        raise LLMHatasi(f"Sağlayıcı hatası ({h.status_code}). Daha sonra tekrar deneyin.") from h
    except anthropic.APIConnectionError as h:
        raise LLMHatasi("Anthropic API'sine bağlanılamadı. İnternet bağlantınızı kontrol edin.") from h

    KULLANIM["girdi_token"] += yanit.usage.input_tokens
    KULLANIM["cikti_token"] += yanit.usage.output_tokens
    if yanit.stop_reason == "refusal":
        raise LLMHatasi("Model bu isteği güvenlik nedeniyle yanıtlamadı.")
    if yanit.stop_reason == "max_tokens":
        raise LLMHatasi("Yanıt uzunluk sınırına takıldı; girdiyi küçültüp tekrar deneyin.")
    metin = next((b.text for b in yanit.content if b.type == "text"), "")
    return json.loads(metin)


def _openai(a, sistem, kullanici, sema, max_tokens):
    try:
        import openai
    except ImportError as h:
        raise LLMHatasi("openai paketi kurulu değil: pip install openai") from h
    if not os.getenv("OPENAI_API_KEY"):
        raise LLMHatasi("OPENAI_API_KEY bulunamadı. .env dosyasına kendi anahtarınızı ekleyin.")
    try:
        yanit = openai.OpenAI().chat.completions.create(
            model=a["model"],
            messages=[{"role": "system", "content": sistem}, {"role": "user", "content": kullanici}],
            response_format={"type": "json_schema", "json_schema": {"name": "cikti", "strict": True, "schema": sema}},
            max_completion_tokens=max_tokens,
        )
    except openai.AuthenticationError as h:
        raise LLMHatasi("OPENAI_API_KEY geçersiz.") from h
    except openai.APIError as h:
        raise LLMHatasi(f"OpenAI hatası: {h}") from h
    if yanit.usage:
        KULLANIM["girdi_token"] += yanit.usage.prompt_tokens
        KULLANIM["cikti_token"] += yanit.usage.completion_tokens
    return json.loads(yanit.choices[0].message.content)


def _ollama(a, sistem, kullanici, sema):
    adres = os.getenv("OLLAMA_HOST", "http://localhost:11434").rstrip("/")
    govde = json.dumps({
        "model": a["model"],
        "messages": [{"role": "system", "content": sistem}, {"role": "user", "content": kullanici}],
        "format": sema,
        "stream": False,
        "options": {"temperature": 0},
    }).encode()
    istek = urllib.request.Request(f"{adres}/api/chat", data=govde, headers={"Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(istek, timeout=600) as r:
            veri = json.loads(r.read())
    except urllib.error.URLError as h:
        raise LLMHatasi(f"Ollama'ya ulaşılamadı ({adres}). Ollama çalışıyor mu?") from h
    KULLANIM["girdi_token"] += veri.get("prompt_eval_count", 0)
    KULLANIM["cikti_token"] += veri.get("eval_count", 0)
    return json.loads(veri["message"]["content"])


def kullanim_ozeti() -> str:
    a = ayarlar()
    return (f"{a['saglayici']} · {a['model']} · {KULLANIM['istek']} istek · "
            f"{KULLANIM['girdi_token']:,} girdi + {KULLANIM['cikti_token']:,} çıktı token").replace(",", ".")
