"""Gerisi Bende başlatıcısı: terminal bilmeden çalıştırmak için.

BASLAT.bat (Windows) ya da BASLAT.command (macOS / Linux) bu dosyayı çalıştırır:
  - ilk seferde pakete özel bir sanal ortam (.venv) kurar ve requirements.txt'yi yükler,
  - "örnek veriyle dene" ya da "kendi dosyalarımla çalıştır" seçeneklerini sunar,
  - kendi dosyalarınız için task.json'daki örnek komuttan dosya argümanlarını bulur ve
    dosyayı pencereye sürükleyip bırakmanızı ister,
  - bitince çıktı klasörünü açar.
"""
import json, os, shlex, shutil, subprocess, sys
from pathlib import Path

K = Path(__file__).resolve().parent
M = json.loads((K / "task.json").read_text(encoding="utf-8"))
VENV = K / ".venv"
UZANTI = (".xlsx", ".xls", ".xlsm", ".xlsb", ".ods", ".csv", ".tsv", ".pdf", ".txt", ".log", ".docx", ".doc", ".odt", ".rtf",
          ".json", ".jsonl", ".ndjson", ".yaml", ".yml", ".md", ".eml", ".msg", ".png", ".jpg", ".jpeg", ".zip", ".xml",
          ".html", ".htm", ".ics", ".pptx", ".sql", ".dat")


def vpy():
    return VENV / ("Scripts/python.exe" if os.name == "nt" else "bin/python")


def kur():
    if not vpy().exists():
        print("\nİlk çalıştırma: bu paket için ayrı bir Python ortamı kuruluyor (yalnızca bir kez)...")
        r = subprocess.run([sys.executable, "-m", "venv", str(VENV)], stdin=subprocess.DEVNULL)
        if r.returncode != 0:
            print("\n[!] Python ortamı kurulamadı."); bekle(); sys.exit(1)
    req = K / "requirements.txt"
    damga = VENV / ".kuruldu"
    icerik = req.read_text(encoding="utf-8") if req.exists() else ""
    if icerik.strip() and (not damga.exists() or damga.read_text(encoding="utf-8") != icerik):
        print("Gerekli kütüphaneler yükleniyor (internet gerekir, 1–2 dakika sürebilir)...")
        r = subprocess.run([str(vpy()), "-m", "pip", "install", "-q", "--disable-pip-version-check", "-r", str(req)], stdin=subprocess.DEVNULL)
        if r.returncode != 0:
            print("\n[!] Kütüphaneler yüklenemedi. İnternet bağlantınızı kontrol edip tekrar deneyin.")
            bekle(); sys.exit(1)
        damga.write_text(icerik, encoding="utf-8")


def bekle():
    try:
        input("\nKapatmak için Enter'a basın...")
    except EOFError:
        pass


def ac(yol: Path):
    if os.environ.get("GB_SESSIZ"):  # otomatik testlerde pencere açma
        print(f"[açılacaktı] {yol}"); return
    try:
        if os.name == "nt":
            os.startfile(str(yol))  # noqa: S606 — kullanıcının kendi klasörünü açar
        elif sys.platform == "darwin":
            subprocess.run(["open", str(yol)])
        else:
            subprocess.run(["xdg-open", str(yol)])
    except Exception:
        print(f"Çıktı klasörü: {yol}")


def env_hazir() -> bool:
    """AI agent'lar için .env (API anahtarı) dosyasını hazırlar."""
    ornek = K / ".env.example"
    if not ornek.exists():
        return True
    env = K / ".env"
    if not env.exists():
        shutil.copy(ornek, env)
        print("\nBu bir AI agent: kendi yapay zekâ API anahtarınız gerekir.")
        print("Açılan .env dosyasına anahtarınızı yazıp kaydedin, sonra buraya dönün.")
        ac(env)
        try:
            input("Kaydettikten sonra Enter'a basın...")
        except EOFError:
            pass
    return True


def argumanlar():
    """task.json'daki örnek komuttan (python main.py --a x --b y z) bayrak → değer listesi çıkarır.
    Bir bayrağı izleyen bütün bayraksız sözcükler o bayrağın değerleridir (--girdi a.log b.log, --donem 01.01 30.06)."""
    t = shlex.split(M.get("calistir", ""), posix=True)
    if t[:1] == ["python"]:
        t = t[1:]
    if t[:1] == [M["ana_dosya"]]:
        t = t[1:]
    ciftler, i = [], 0
    while i < len(t):
        if t[i].startswith("--"):
            degerler = []
            i += 1
            while i < len(t) and not t[i].startswith("--"):
                degerler.append(t[i]); i += 1
            ciftler.append((t[i - 1 - len(degerler)], degerler))
        else:
            i += 1  # bayrağa bağlı olmayan sözcük: örnek komutta beklenmez
    return ciftler


def dosya_mi(deger: str) -> bool:
    d = deger.lower()
    return d.endswith(UZANTI) or "/" in d or "\\" in d


def sor(istem: str) -> str:
    try:
        return input(istem).strip().strip('"').strip("'").strip()
    except EOFError:
        return ""


def kendi_verim():
    komut = []
    print("\nHer dosya için: dosyayı bu pencereye sürükleyip bırakın ve Enter'a basın.")
    print("Boş bırakırsanız o seçenek kullanılmaz. Ayrıntılı girdi standardı README.md dosyasında.\n")
    for bayrak, degerler in argumanlar():
        if not degerler:
            komut.append(bayrak); continue
        ad = bayrak.lstrip("-").replace("-", " ")
        if len(degerler) > 1:
            # çok değerli seçenek: her değer ayrı sorulur. Dosya listesinde boş bırakılan atlanır (hepsi boşsa
            # seçenek verilmez, paket örnek veriyle çalışır); sayı / metin listesinde boş bırakılınca örnek değer kalır.
            dosya = any(dosya_mi(d) for d in degerler)
            print(f"  {ad}: {len(degerler)} değer (örnek: {' '.join(degerler)})")
            girilen = []
            for k, d in enumerate(degerler, 1):
                g = sor(f"    {k}. değer [{'boş = atla' if dosya else d}]: ")
                if g:
                    girilen.append(g)
                elif not dosya:
                    girilen.append(d)
            if girilen:
                komut += [bayrak, *girilen]
            continue
        deger = degerler[0]
        if dosya_mi(deger):
            g = sor(f"  {ad} (örnek: {deger}): ")
            if g:
                komut += [bayrak, g]
        else:
            komut += [bayrak, sor(f"  {ad} [{deger}]: ") or deger]
    return komut


def calistir(ekler):
    ortam = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
    print("\nÇalışıyor...\n" + "-" * 60)
    r = subprocess.run([str(vpy()), M["ana_dosya"], *ekler], cwd=K, env=ortam)
    print("-" * 60)
    if r.returncode == 0:
        print("\n✓ Bitti.")
        cikti = K / "cikti"
        ac(cikti if cikti.exists() else K)
    else:
        print("\n[!] Bir hata oluştu. Yukarıdaki mesajı okuyun; takılırsanız sitedeki 'Yardım al' bölümünden yazın.")


def main():
    try:
        sys.stdout.reconfigure(line_buffering=True)  # alt süreç çıktısıyla sıra karışmasın
    except Exception:
        pass
    print("=" * 60)
    print(f"  Gerisi Bende · {M['ad']}")
    print("=" * 60)
    kur()
    env_hazir()
    while True:
        print("\n  1) Örnek veriyle dene")
        print("  2) Kendi dosyalarımla çalıştır")
        print("  3) Çıkış")
        try:
            s = input("\nSeçiminiz (1/2/3): ").strip()
        except EOFError:
            s = "3"
        if s == "1":
            calistir([])
        elif s == "2":
            calistir(kendi_verim())
        elif s == "3" or s == "":
            break


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        pass
    bekle()
