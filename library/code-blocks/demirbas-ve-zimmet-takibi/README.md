# Demirbaş ve Zimmet Takibi · Kod Bloğu

> İdari İşler › İdari İşler Sorumlusu · Workers / Workless

Demirbaş kaydınızı sayım listesiyle karşılaştırır: hangi demirbaş **yerinde**, hangisi **başka konumda**, hangisi
**bulunamadı**, hangisi **kayıtta yok** gösterir. Konum ve zimmetli kişi bazında özet verir. İnternete bağlanmaz.

## Nasıl hesaplar?

| Sonuç | Kural |
|---|---|
| Yerinde | Sayımda bulundu ve bulunduğu konum kayıttaki konumla aynı |
| Farklı konumda | Sayımda bulundu ama konum kayıttakinden farklı (ya da kayıtta konum yok) |
| Bulunamadı | Kayıtta aktif, sayım listesinde yok |
| Mükerrer sayım | Aynı demirbaş numarası birden çok farklı konumda sayılmış |
| Kayıt dışı | Sayım listesinde var, demirbaş kaydında yok |
| Kayıttan çıkmış, sayıldı | Durumu Hurda / Satıldı / Devredildi / Kayıp / Pasif olan demirbaş sayımda bulundu |
| Sayım oranı % | Sayımda bulunan aktif demirbaş ÷ aktif demirbaş × 100 |

Eşleştirme demirbaş numarasıyla yapılır; büyük-küçük harf, boşluk ve tire farkı dikkate alınmaz (`DMB-0001` =
`dmb 0001`). Konum karşılaştırması da harf ve noktalama farkına bakmaz.

**Örnek:** 10 aktif demirbaştan 9'u sayımda bulundu → sayım oranı **%90**; bulunamayan projeksiyon cihazının
kayıtlı değeri 32.000,00.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                  # örnek veri
python main.py --demirbaslar demirbaslar.xlsx --sayim sayim.xlsx
python main.py --demirbaslar demirbaslar.xlsx                   # sayımsız: envanter ve zimmet özeti
```

Zorunlu sütun iki dosyada da Demirbaş No. Durum boşsa demirbaş aktif sayılır.

## Çıktı

`cikti/demirbas_zimmet_takibi.xlsx`:

- `Demirbaşlar`: kaydın tamamı, sayım sonucu ve sayımda bulunduğu yer.
- `Sayım Farkları`: yalnız farklar; "Açıklama" sütununu siz doldurursunuz (sayım dosyası verildiyse).
- `Konum Özeti`, `Zimmet Özeti`: adet, değer, yerinde / farklı konumda / bulunamayan sayısı ve bulunamayan değer.
- `Uyarılar`: bulunamayan, mükerrer, kayıt dışı, zimmetsiz ve konumu yazılmamış demirbaşlar.

## Dikkat

- **"Bulunamadı" kayıp demek değildir.** Sayım listesinde olmadığını gösterir; sayılmamış bir oda ya da okunmamış
  bir etiket de aynı sonucu verir. Kayıttan düşmeden önce yerinde arayın.
- **Değer, kayda yazdığınız tutardır.** Kod amortisman, net defter değeri ya da rayiç değer hesaplamaz; muhasebe
  sabit kıymet kayıtlarıyla mutabakat ayrı bir iştir.
- **Zimmet tutanağının yerini tutmaz.** Kimde ne olduğunu gösterir; imzalı zimmet ve iade tutanaklarını ayrıca
  saklayın.
- **Kapsam dışı:** adet bazlı (numarasız) sarf malzemesi, bilgisayar ve telefon için garanti ve yenileme takibi
  (bunun için Donanım Zimmet Envanteri paketine bakın), amortisman ve sigorta.
- **Kişisel veri:** Zimmetli kişi adları yer alır; raporu yalnız ilgili kişilerle paylaşın.
- **Örnek veri:** Demirbaşlar, kişiler ve değerler kurgusaldır.

## Testler

- Her sonuç türü: yerinde, farklı konumda, bulunamadı, mükerrer sayım, kayıt dışı, kayıttan çıkmış ama sayılmış.
- Numara ve konum eşleştirmesinde yazım farkı; konumu boş kayıt ve konumu boş sayım satırı.
- Sayım oranı, konum ve zimmet özeti, sayımsız çalışma, mükerrer kayıt numarası, okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
