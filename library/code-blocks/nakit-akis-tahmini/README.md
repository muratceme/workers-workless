# Nakit Akış Tahmini · Kod Bloğu

> Finans › Finans Uzmanı · Workers / Workless

Vadeli alacak ve borçlar, alınan/verilen çekler, kredi taksitleri ve tekrarlayan ödemelerden (maaş, kira,
vergi, SGK) **haftalık nakit akış tahmini** (varsayılan 13 hafta) üretir. Minimum nakit eşiğinin altına düşen
haftaları kırmızıyla gösterir. İnternete bağlanmaz.

## Varsayımlar (değiştirilebilir)

- `--tahsilat-gecikmesi`: "Müşteri" kategorisindeki tahsilatlar vadeden bu kadar gün sonra gerçekleşir.
- `--gecikmis-tahsil-orani`: başlangıçtan önce vadesi geçmiş alacakların bu yüzdesi 1. haftada tahsil edilir (varsayılan %50).
- Vadesi geçmiş borçlar 1. haftada ödenir.
- Tekrarlayan kalemlerin (Haftalık/Aylık) başlangıçtan önceki dönemleri tahmine alınmaz.

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                             # örnek kalemlerle dener
python main.py --girdi nakit_kalemleri.xlsx --acilis 1250000 --baslangic 05.10.2026 --minimum 250000 --tahsilat-gecikmesi 7
```

Sütunlar: **Tür** (Giriş/Çıkış), **Tarih**, **Tutar** zorunlu; Kategori (ör. Müşteri, Tedarikçi, Maaş, Kira,
Kredi, Vergi, SGK, Çek), Açıklama, Tekrar (Haftalık/Aylık), Bitiş isteğe bağlı.

## Testler

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
