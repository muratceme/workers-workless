"""
Kütüphane kalite kapısı. CI her push'ta çalıştırır; yerelde de çalıştırabilirsiniz.

    python scripts/kontrol.py              # standart kontrolü
    python scripts/kontrol.py --senkronla  # ortak dosyaları kaynaklarından kopyala
    python scripts/kontrol.py --test       # + her görevin testlerini çalıştır

Kontroller:
  - Klasör standardı (task.json, README.md, requirements.txt, ana dosya, tests/, ornek_veri/)
  - Tüm bağımlılıklar `paket==sürüm` biçiminde sabitlenmiş mi
  - Kod blokları internete/kabuğa erişen modül kullanıyor mu (yasak)
  - Agent'larda .env.example var mı, gerçek .env yanlışlıkla eklenmiş mi
  - ortak_dosyalar kopyaları kaynaklarıyla birebir aynı mı
  - çift tıkla başlatıcı (scripts/baslatici) derleniyor ve ağ modülü kullanmıyor mu
  - kod bloklarının örnek çıktı önizlemeleri (data/onizlemeler.json) içerikle güncel mi
"""
from __future__ import annotations

import argparse
import ast
import filecmp
import json
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path

KOK = Path(__file__).resolve().parents[1]
LIB = KOK / "library"
YASAK_MODULLER = {"socket", "urllib", "http", "requests", "httpx", "aiohttp", "ftplib", "smtplib",
                  "telnetlib", "paramiko", "subprocess", "webbrowser", "xmlrpc"}
YASAK_CAGRILAR = {("os", "system"), ("os", "popen")}
PIN = re.compile(r"^[A-Za-z0-9_.\-\[\]]+==[\w.\-+]+$")


def gorevler():
    # "istekler": Sizden Gelenler paketleri; kod / agent kuralları task.json'daki 'tur' alanına göre uygulanır
    for tur in ("code-blocks", "agents", "istekler"):
        d = LIB / tur
        if d.is_dir():
            for k in sorted(d.iterdir()):
                if k.is_dir() and not k.name.startswith("_"):
                    yield tur, k


def ayni_mi(a: Path, b: Path) -> bool:
    if a.is_file() and b.is_file():
        return a.read_bytes() == b.read_bytes()
    if a.is_dir() and b.is_dir():
        c = filecmp.dircmp(a, b, ignore=["__pycache__"])
        return not (c.left_only or c.right_only or c.diff_files or c.funny_files) and all(
            ayni_mi(a / s, b / s) for s in c.common_dirs)
    return False


def kopyala(kaynak: Path, hedef: Path) -> None:
    if hedef.is_dir():
        shutil.rmtree(hedef)
    elif hedef.exists():
        hedef.unlink()
    if kaynak.is_dir():
        shutil.copytree(kaynak, hedef, ignore=shutil.ignore_patterns("__pycache__", "cikti"))
    else:
        shutil.copy2(kaynak, hedef)


def ag_kullanimi(py: Path) -> list[str]:
    bulunan = []
    agac = ast.parse(py.read_text(encoding="utf-8"))
    for dugum in ast.walk(agac):
        if isinstance(dugum, ast.Import):
            bulunan += [a.name for a in dugum.names if a.name.split(".")[0] in YASAK_MODULLER]
        elif isinstance(dugum, ast.ImportFrom) and dugum.module and dugum.module.split(".")[0] in YASAK_MODULLER:
            bulunan.append(dugum.module)
        elif isinstance(dugum, ast.Call) and isinstance(dugum.func, ast.Attribute) and isinstance(dugum.func.value, ast.Name):
            if (dugum.func.value.id, dugum.func.attr) in YASAK_CAGRILAR:
                bulunan.append(f"{dugum.func.value.id}.{dugum.func.attr}()")
    return bulunan


def kontrol_et(tur: str, klasor: Path, senkronla: bool) -> list[str]:
    h = []
    ad = f"{tur}/{klasor.name}"
    manifest = klasor / "task.json"
    if not manifest.exists():
        return [f"{ad}: task.json yok"]
    m = json.loads(manifest.read_text(encoding="utf-8"))

    if m.get("id") != klasor.name:
        h.append(f"{ad}: task.json id '{m.get('id')}' klasör adıyla aynı değil")
    beklenen = {"code-blocks": {"kod"}, "agents": {"agent"}, "istekler": {"kod", "agent"}}[tur]
    if m.get("tur") not in beklenen:
        h.append(f"{ad}: task.json 'tur' alanı hatalı")
    kod = m.get("tur") == "kod"
    for alan in ("ad", "surum", "python", "ana_dosya", "calistir", "girdiler", "ciktilar"):
        if alan not in m:
            h.append(f"{ad}: task.json '{alan}' alanı eksik")

    for kaynak_rel, hedef_rel in [(v, k) for k, v in m.get("ortak_dosyalar", {}).items()]:
        kaynak, hedef = LIB / kaynak_rel, klasor / hedef_rel
        if not kaynak.exists():
            h.append(f"{ad}: ortak kaynak yok: library/{kaynak_rel}")
        elif senkronla:
            kopyala(kaynak, hedef)
        elif not ayni_mi(kaynak, hedef):
            h.append(f"{ad}: '{hedef_rel}' kaynağıyla (library/{kaynak_rel}) aynı değil -> python scripts/kontrol.py --senkronla")

    gerekli = ["README.md", "requirements.txt", m.get("ana_dosya", "?"), "ornek_veri"]
    if not kod:
        gerekli += ["prompt.md", ".env.example"]
    for f in gerekli:
        if not (klasor / f).exists():
            h.append(f"{ad}: '{f}' eksik")
    if not list((klasor / "tests").glob("test_*.py")):
        h.append(f"{ad}: tests/test_*.py yok")
    if (klasor / ".env").exists():
        h.append(f"{ad}: gerçek .env dosyası depoya eklenmemeli")

    req = klasor / "requirements.txt"
    if req.exists():
        for satir in req.read_text(encoding="utf-8").splitlines():
            s = satir.split("#")[0].strip()
            if s and not PIN.match(s):
                h.append(f"{ad}: bağımlılık sabitlenmemiş: '{s}' (paket==sürüm olmalı)")

    if kod:
        if m.get("internet") is not False:
            h.append(f"{ad}: kod blokları için task.json 'internet': false olmalı")
        for py in klasor.rglob("*.py"):
            if "tests" in py.relative_to(klasor).parts:
                continue
            for y in ag_kullanimi(py):
                h.append(f"{ad}: {py.relative_to(klasor)} yasak modül/çağrı kullanıyor: {y}")
    return h


def test_et(klasor: Path) -> bool:
    ortam = {**os.environ, "PYTHONIOENCODING": "utf-8"}
    r = subprocess.run([sys.executable, "-m", "unittest", "discover", "-s", "tests"], cwd=klasor,
                       capture_output=True, text=True, env=ortam, encoding="utf-8")
    if r.returncode != 0:
        print(r.stdout[-3000:], r.stderr[-3000:])
        return False
    # Komut satırı yardımının çalıştığını da doğrula (argparse yardım metni hataları vb.)
    ana = json.loads((klasor / "task.json").read_text(encoding="utf-8"))["ana_dosya"]
    h = subprocess.run([sys.executable, ana, "--help"], cwd=klasor, capture_output=True, text=True, env=ortam, encoding="utf-8")
    if h.returncode != 0:
        print(f"{klasor.name}: '{ana} --help' başarısız:", h.stderr[-1500:])
        return False
    return True


def baslatici_kontrol() -> list[str]:
    """Her ZIP'e eklenen çift tıkla başlatıcı (scripts/baslatici): derlenmeli, ağ modülü kullanmamalı.
    subprocess'e izin var (pip ve paketi çalıştırmak için gerekir)."""
    h = []
    d = KOK / "scripts" / "baslatici"
    for f in ("BASLAT.bat", "BASLAT.command", "_baslat.py"):
        if not (d / f).exists():
            h.append(f"baslatici: '{f}' eksik")
    py = d / "_baslat.py"
    if py.exists():
        try:
            ast.parse(py.read_text(encoding="utf-8"))
        except SyntaxError as e:
            h.append(f"baslatici: _baslat.py derlenmiyor: {e}")
        else:
            for y in ag_kullanimi(py):
                if y != "subprocess":
                    h.append(f"baslatici: _baslat.py yasak modül/çağrı kullanıyor: {y}")
    return h


def onizleme_kontrol() -> list[str]:
    """Sitedeki örnek çıktı önizlemeleri (data/onizlemeler.json) paket içeriğiyle güncel mi?"""
    sys.path.insert(0, str(KOK / "scripts"))
    from onizleme import anahtarlar, imza  # noqa: E402
    yol = KOK / "data" / "onizlemeler.json"
    if not yol.exists():
        return ["önizleme yok: data/onizlemeler.json -> scripts/onizleme.py çalıştırın"]
    d = json.loads(yol.read_text(encoding="utf-8"))
    h = []
    for anahtar in anahtarlar():
        v = d.get(anahtar)
        ad = anahtar if "/" in anahtar else f"code-blocks/{anahtar}"
        if not v or "hata" in v:
            h.append(f"{ad}: önizleme yok -> scripts/onizleme.py {anahtar}")
        elif v.get("imza") != imza(anahtar):
            h.append(f"{ad}: önizleme eski (kod / örnek veri değişmiş) -> scripts/onizleme.py {anahtar}")
    return h


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    p = argparse.ArgumentParser()
    p.add_argument("--senkronla", action="store_true")
    p.add_argument("--test", action="store_true")
    a = p.parse_args()

    hatalar, sayi = [], 0
    for tur, klasor in gorevler():
        sayi += 1
        hatalar += kontrol_et(tur, klasor, a.senkronla)
    hatalar += baslatici_kontrol()
    hatalar += onizleme_kontrol()
    for hata in hatalar:
        print(f"[X] {hata}")
    print(f"[{'OK' if not hatalar else 'X'}] {sayi} görev klasörü kontrol edildi, {len(hatalar)} sorun")

    if a.test and not hatalar:
        basarisiz = [f"{t}/{k.name}" for t, k in gorevler() if not test_et(k)]
        for b in basarisiz:
            print(f"[X] testler başarısız: {b}")
        print(f"[{'OK' if not basarisiz else 'X'}] testler: {sayi - len(basarisiz)}/{sayi} geçti")
        return 1 if basarisiz else 0
    return 1 if hatalar else 0


if __name__ == "__main__":
    sys.exit(main())
