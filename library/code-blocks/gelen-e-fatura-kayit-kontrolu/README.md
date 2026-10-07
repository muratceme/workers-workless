# Gelen e-Fatura Kayıt Kontrolü · Kod Bloğu

> Muhasebe › Muhasebe Elemanı · Mali Müşavirlik Bürosu › Muhasebe Elemanı · Workers / Workless

GİB veya özel entegratör portalından alınan **gelen e-fatura / e-arşiv** listesini muhasebe **alış
kayıtlarıyla** karşılaştırır. Deftere işlenmemiş, mükerrer işlenmiş, tutarı farklı ve portalda karşılığı
olmayan kayıtları ayrı sayfalarda gösterir. İnternete bağlanmaz.

> Form Ba-Bs bildirimi, 565 Sıra No'lu VUK Genel Tebliği ile 01.10.2024'ten itibaren kaldırılmıştır.
> e-Belge verileri İdarece doğrudan izlenebildiğinden, gelen e-faturaların eksiksiz ve doğru kaydedildiğini
> dönem sonunda kontrol etmek daha da önemli hale gelmiştir.

## Eşleştirme

1. Fatura numarası + satıcı VKN (boşluk/tire ve büyük-küçük harf farkı yok sayılır)
2. Fatura numarası (kayıtta VKN yoksa)
3. Satıcı VKN + vergiler dahil tutar + tarih farkı ≤ `--gun-toleransi` → belge numarası **yanlış girilmiş**
   kayıtları bulur (Eşleşenler sayfasında sarı)

Dövizli e-faturalar faturadaki kurla TL'ye çevrilir. Aynı faturanın ZIP içindeki ikinci kopyası bir kez sayılır.

## Çıktı (`cikti/efatura_kayit_kontrolu.xlsx`)

| Sayfa | Anlamı |
|---|---|
| Deftere İşlenmemiş | Portalda var, muhasebede yok — KDV indirimi ve gider kaybı riski |
| Tutarı Farklı | Eşleşti ama kayıt tutarı farklı |
| Portalda Yok | Muhasebede var, e-fatura listesinde yok — kâğıt fatura, SMM, ithalat, yanlış dönem veya mükerrer kayıt olabilir |
| Mükerrer Kayıtlar | Aynı belge no + VKN birden çok kez kaydedilmiş |
| Eşleşenler | Tüm eşleşmeler ve yöntemi |

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                                         # örnek veriyle dener
python main.py --efatura "C:/Gelen/2026-09" --kayitlar muavin_alislar.xlsx
python main.py --efatura portal_listesi.xlsx --kayitlar alis_kayitlari.csv --gun-toleransi 5
```

Muhasebe kaydı sütunları (başlıklar esnek): **Belge No / Fatura No** ve **Tutar** zorunlu; Tarih, Cari VKN /
Vergi No, Cari Unvan, Açıklama isteğe bağlı. Portal listesi için de aynı başlık esnekliği geçerlidir
(ör. "Gönderici VKN/TCKN", "Vergiler Dahil Tutar").

## Testler

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
