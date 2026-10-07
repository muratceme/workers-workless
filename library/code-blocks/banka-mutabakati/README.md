# Banka Mutabakatı · Kod Bloğu

> Finans & Muhasebe › Muhasebe Elemanı · Workers / Workless

Banka ekstresini muhasebe defteriyle (102 Bankalar muavini) eşleştirir; eşleşenleri,
iki taraftaki açık kalemleri ve **kural tabanlı olası nedenleri** tek Excel raporunda verir.
İnternete bağlanmaz, API anahtarı gerekmez.

## Nasıl eşleştirir?

1. **Referans:** iki açıklamada ortak dekont/fatura numarası (6+ hane) ve aynı tutar varsa, tarih farkı ne olursa olsun eşler.
2. **Tutar + Tarih:** aynı tutar ve tarih farkı ≤ `--gun-toleransi`; birden fazla aday varsa en yakın tarih, sonra en benzer açıklama seçilir.

Açık kalemler için olası nedenler: banka masrafı/BSMV, faiz, zamanlama farkı, tutar hatası
(hane kayması), mükerrer kayıt, bankaya henüz yansımamış işlem, deftere atlanmış kayıt.

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

# Örnek veriyle deneyin
python main.py

# Kendi dosyalarınız
python main.py --banka ekstre.xlsx --defter muavin_102.xlsx --gun-toleransi 3
```

## Girdi standardı

`.csv` (virgül, noktalı virgül veya sekme) ya da `.xlsx`. İlk satır başlık olmalı.

| Sütun | Kabul edilen başlıklar |
|---|---|
| tarih | Tarih, İşlem Tarihi, Valör, Fiş Tarihi, Date |
| aciklama | Açıklama, İşlem Açıklaması, Detay, Description |
| tutar **veya** borc + alacak | Tutar, İşlem Tutarı, Amount · Borç, Alacak, Debit, Credit |

İşaret kuralı (tek `tutar` sütunu varsa): **hesaba giriş +, çıkış −**.
Borç/alacak sütunları varsa: banka için `alacak − borç`, defter (102) için `borç − alacak`.
Sayılar `1.234,56` veya `1,234.56` biçiminde olabilir. Boş şablon: `ornek_veri/sablon_*.csv`.

> Bankanızın ekstresinde borç/alacak anlamı ters ise, **Özet** sayfasındaki toplamlar
> hemen ters işaretli çıkar; bu durumda sütun başlıklarının yerini değiştirmeniz yeterlidir.

## Çıktı

- **Özet:** sayılar, toplamlar, fark ve açık kalemlerle açıklanan fark (ikisi eşitse mutabakat tamdır).
- **Eşleşenler**, **Açık - Banka**, **Açık - Defter**.

## Sınırlamalar

- Bire-çok eşleşme (tek EFT ↔ birden fazla fatura) şu an desteklenmez; açık kalem olarak görünür.
- Nedenler kural tabanlı tahmindir. Muhasebe kaydı yapmadan önce kontrol edin.
  Yevmiye kaydı önerisi ve serbest metin yorumu için bu görevin **AI Agent** sürümüne bakın.

## Testler

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
