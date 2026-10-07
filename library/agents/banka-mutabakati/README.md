# Banka Mutabakatı · AI Agent

> Finans & Muhasebe › Muhasebe Elemanı · Workers / Workless

Mutabakatın **eşleştirme kısmını yapay zekâya bırakmaz**: kod bloğunun deterministik motoru
çalışır, sonuç her seferinde aynıdır. Yapay zekâ yalnızca eşleşmeyen **açık kalemleri**
yorumlar:

- olası neden (masraf, faiz, zamanlama, tutar hatası, mükerrer kayıt…)
- karşı taraftaki ilişkili kalem
- düzeltme için **yevmiye kaydı taslağı** (ör. `780 Finansman Giderleri (B) / 102 Bankalar (A)`)
- güven düzeyi ve genel değerlendirme

Raporda yapay zekâ sütunları mor, **İnsan Onayı** sütunu sarıdır.

## Gizlilik

- Eşleşen kayıtlar **hiç gönderilmez**; yalnızca açık kalemlerin tarih, açıklama ve tutarı gider.
- Açıklamalardaki IBAN, TCKN, e-posta ve telefon maskelenir.
- Göndermeden önce onayınız istenir. Veri hiç çıkmasın istiyorsanız `WW_PROVIDER=ollama`.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env
```

`.env` içine kendi API anahtarınızı yazın (bkz. `.env.example`). API ücretleri size aittir.

## Kullanım

```bash
# Örnek veriyle deneyin
python agent.py

# Kendi dosyalarınız
python agent.py --banka ekstre.xlsx --defter muavin_102.xlsx --cikti cikti/mutabakat_ai.xlsx
```

Girdi standardı kod bloğuyla aynıdır: `tarih`, `aciklama` ve `tutar` (giriş +, çıkış −) ya da
`borc` + `alacak` sütunları; `.csv` veya `.xlsx`. Ayrıntı ve şablonlar: `ornek_veri/`.

| Parametre | Açıklama |
|---|---|
| `--banka`, `--defter` | Girdi dosyaları |
| `--cikti` | Excel çıktısı (varsayılan `cikti/mutabakat_ai.xlsx`) |
| `--gun-toleransi` | Tutar+tarih eşleşmesinde gün farkı (varsayılan 3) |
| `--evet` | Onay sorusunu atla |

## Testler

Testler gerçek API'yi çağırmaz:

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Yevmiye önerileri taslaktır. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
