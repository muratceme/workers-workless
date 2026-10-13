# Tahsilat ve Avukat Performans Raporu · Kod Bloğu

> Krediler İzleme ve Takip › Yasal Takip Uzmanı · Workers / Workless

Yasal takipteki dosyaların **tahsilatını, masrafını ve net sonucunu** sözleşmeli avukat ve dosya yaşı bazında
raporlar; uzun süredir tahsilat gelmeyen ve masrafı tahsilatını aşan dosyaları işaretler. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Dosya yaşı (ay) | Takip tarihinden durum tarihine kadar geçen tam ay sayısı |
| Yaş dilimi | 0–6, 7–12, 13–24, 25+ ay (`--dilimler`) |
| Tahsilat oranı % | Tahsilat ÷ takip tutarı × 100 |
| Net tahsilat | Tahsilat − masraf |
| Masraf / tahsilat % | Masraf ÷ tahsilat × 100 |
| Tahsilatsız eski dosya | Durumu açık, yaşı `--eski` (12) ay ve üstü, tahsilatı sıfır |

Grup oranları toplamlar üzerinden hesaplanır; dosya oranlarının basit ortalaması değildir.

**Örnek:** Bir avukatın dört dosyasında takip tutarı 1.150.000, tahsilat 590.000 → tahsilat oranı **%51,30**; masraf
32.000 → net **558.000**, masraf / tahsilat **%5,42**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                            # örnek veri, durum tarihi 12.10.2026
python main.py --dosyalar dosyalar.xlsx
python main.py --dosyalar dosyalar.xlsx --dilimler "12;24;36" --eski 18
```

Zorunlu sütunlar: Dosya No, Takip Tutarı. Durum boşsa dosya açık sayılır; "Kapalı", "İnfaz" ve "Aciz" yazılan dosyalar
açık dosya sayısına girmez.

## Çıktı

`cikti/tahsilat_avukat_performans.xlsx`:

- `Avukatlar`: tahsilat oranına göre sıralı; dosya sayısı, takip tutarı, tahsilat, masraf, net, ortalama dosya yaşı.
- `Dosya Yaşı`: aynı göstergeler yaş dilimi bazında.
- `Avukat × Yaş`: her avukatın her yaş dilimindeki dosya sayısı ve tahsilat oranı.
- `Dosyalar` ve `Uyarılar`.

## Dikkat

- **Avukatları ham oranla karşılaştırmayın.** Tahsilat oranı dosyanın yaşına, tutarına, teminatına ve borçlunun
  durumuna bağlıdır; zor dosyaları alan avukat düşük görünür. Adil karşılaştırma için `Avukat × Yaş` sayfasında aynı
  yaş dilimine bakın ve dosya sayısı az olan hücreleri yorumlamayın.
- **Tahsilat ve takip tutarı aynı tanımda olmalı.** Takip tutarına faiz ve masrafı dahil edip etmediğiniz, tahsilatı
  brüt mü net mi yazdığınız bütün dosyalarda aynı olmalıdır. Tahsilat, dosya açıldığından bugüne toplam tahsilattır;
  dönemsel performans için dönem tahsilatını ayrı dosyada çalıştırın.
- **Masraf tanımı sizindir.** Harç, vekâlet ücreti, bilirkişi ve satış masraflarından hangilerini yazdığınız net
  tahsilatı değiştirir; karşı taraftan tahsil edilecek masraflar ayrıca izlenmelidir.
- **Tutarlar tek para biriminde olmalı.**
- **Kapsam dışı:** dosya aşamaları ve süreleri (haciz, satış, itiraz), teminatın paraya çevrilmesi, vekâlet ücreti
  hakedişi, karşılık ve aktiften silme.
- **Gizlilik:** Rapor borçlu ve dosya bilgisi içerir; yalnız yetkili kişilerle paylaşın.
- **Örnek veri:** Avukatlar, dosyalar ve tutarlar kurgusaldır.

## Testler

- Dosya yaşı (ay sınırı, gün farkı, ileri tarih), yaş dilimi sınırları, özel dilimler, tarihi olmayan dosya.
- Tahsilat oranı, net, masraf oranı; toplamlar üzerinden grup oranı; avukat sıralaması ve avukat × yaş tablosu.
- Tahsilatsız eski dosya eşiği ve durum koşulu, masrafı aşan dosya, okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
