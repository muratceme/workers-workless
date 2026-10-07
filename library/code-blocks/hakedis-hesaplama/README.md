# Hakediş Hesaplama · Kod Bloğu

> İnşaat › Teknik Ofis › Teknik Ofis Mühendisi · Workers / Workless

Birim fiyatlı sözleşmenin **poz listesi** ve dönem **metrajlarından** hakediş icmalini hazırlar:
kümülatif metraj özeti, *toplam − önceki = bu hakediş* tablosu, KDV, KDV tevkifatı, stopaj, teminat ve
avans kesintileri, ödenecek net tutar ve tüm hakedişlerin geçmişi. İnternete bağlanmaz.

## Vergi ve kesinti kuralları

| Kalem | Varsayılan | Dayanak |
|---|---|---|
| KDV | %20 | `kdv_orani` |
| KDV tevkifatı | **4/10**; alıcı belirlenmiş alıcıysa ya da KDV dahil sözleşme bedeli ≥ 5.000.000 TL ise (`otomatik`) | KDV Genel Uygulama Tebliği I/C-2.1.3.2.1 (Seri No: 35 ile 01.03.2021'den beri) |
| Stopaj | **%5**, yıllara yaygın inşaat ve onarım işlerinde (`yillara_yaygin: true`) | GVK md. 94/3, KVK md. 15; 3491 sayılı CBK. Demiryolu, gemi ve nükleer santral inşaatında %1 |
| Damga vergisi | 0 (sözleşmeniz hakedişten kesilmesini öngörüyorsa oranı girin) | `damga_vergisi_orani` |
| Teminat kesintisi | 0 | `teminat_kesintisi_orani` |
| Avans mahsubu | Bu hakediş × avans / sözleşme bedeli; toplamı avansı aşmaz | `avans_tutari` |
| Fiyat farkı | Hesaplanmış tutar girilir (bu paket fiyat farkı formülünü uygulamaz) | `fiyat_farki` |

Tutarlar **kümülatif** hesaplanır, bu hakediş önceki hakedişlerden farkla bulunur; yuvarlama farkı birikmez.

## Kontroller

- Kümülatif miktarı **sözleşme miktarını aşan pozlar** (iş artışı / onaylı fazla metraj gerekir)
- Toplam iş tutarı sözleşme bedelini aşarsa iş artışı uyarısı (kamu işlerinde 4735 sayılı Kanun md. 24)
- Metrajda olup **poz listesinde olmayan** pozlar (yeni birim fiyat gerekir) hesaba katılmaz ve listelenir

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                                   # örnek sözleşmeyle 3 no'lu hakediş
python main.py --sozlesme sozlesme.json --pozlar pozlar.xlsx --metraj metraj.xlsx --hakedis 4
```

`ornek_veri/sozlesme.json` dosyasını kendi sözleşmenize göre kopyalayıp düzenleyin.

## Çıktı

`Hakediş İcmali` · `Metraj Özeti` · `Hakediş Geçmişi` · `Bilgi`

> Sonuçlar sözleşme şartlarına ve güncel mevzuata göre mali müşavir ve kontrol mühendisi tarafından
> teyit edilmelidir. Oranlar değişirse `sozlesme.json` üzerinden güncelleyin.

## Testler

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
