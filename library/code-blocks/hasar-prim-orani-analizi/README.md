# Hasar/Prim Oranı Analizi · Kod Bloğu

> Teknik (Risk Kabul / Underwriting) › Teknik Uzman (Underwriter) · Workers / Workless

Poliçe ve hasar listesinden **müşteri, acente ve branş bazında hasar/prim (H/P) oranını** hesaplar; yenilemesi yaklaşan
poliçeler için eşiklere dayalı bir öneri yazar. İnternete bağlanmaz.

## Nasıl hesaplar?

```
Kazanılmış prim   = prim × (rapor tarihine kadar geçen gün ÷ poliçe gün sayısı)
                    poliçe başlamadıysa 0, süresi dolduysa primin tamamı
Gerçekleşen hasar = ödenen + muallak (bekleyen)
H/P %             = gerçekleşen hasar ÷ kazanılmış prim × 100
```

- Hasar tarihi rapor tarihinden sonra olan hasarlar alınmaz.
- Örnek: 73.000 TL primli, 01.04.2026 – 01.04.2027 poliçe; rapor tarihi 01.10.2026 → 183 ÷ 365 → kazanılmış prim
  36.600 TL. Gerçekleşen hasar 45.000 TL ise H/P = %122,95.

**Yenileme önerisi** müşterinin tüm poliçelerinin toplam H/P oranına göre verilir:

| Grup | Koşul | Öneri |
|---|---|---|
| Olumlu | H/P ≤ `--iyi` (varsayılan %40) | Mevcut şartlarla yenileme değerlendirilebilir |
| İzlenecek | `--iyi` < H/P ≤ `--kotu` (varsayılan %80) | Fiyat ve muafiyetler gözden geçirilmeli |
| Olumsuz | H/P > `--kotu` | Şart ve fiyat revizyonu veya yenilememe değerlendirilmeli |
| Veri yetersiz | Kazanılmış prim 0 | Oran hesaplanamadı |

Hasar adedi `--sik-hasar` (varsayılan 3) ve üzerindeyse öneriye hasar sıklığı notu eklenir. "Yenilemesi Yaklaşanlar",
bitiş tarihi rapor tarihinden sonraki `--yenileme-gun` (varsayılan 60) gün içinde olan poliçelerdir.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri (rapor tarihi 01.10.2026)
python main.py --policeler policeler.xlsx --hasarlar hasarlar.xlsx --tarih 31.12.2026 --iyi 50 --kotu 90
```

| Dosya | Sütunlar |
|---|---|
| Poliçeler | Poliçe No, Müşteri, Acente, Branş, Başlangıç, Bitiş, Prim |
| Hasarlar | Hasar No, Poliçe No, Hasar Tarihi, Ödenen, Muallak |

## Çıktı

`cikti/hasar_prim_analizi.xlsx`:

- `Özet`: toplam yazılan ve kazanılmış prim, hasarlar, H/P.
- `Müşteri Bazında`: H/P, grup ve yenileme önerisi.
- `Acente Bazında`, `Branş Bazında`: aynı göstergeler.
- `Yenilemesi Yaklaşanlar`: poliçe ve müşteri H/P'si, öneri; boş "Karar" sütunu.
- `Poliçeler`: poliçe bazında kazanılmış prim ve hasarlar.
- `Uyarılar`: poliçesi bulunmayan hasarlar, okunamayan satırlar.

## Dikkat

- **Basit H/P oranıdır.** Gerçekleşmiş ama bildirilmemiş hasarlar (IBNR), hasar masrafları, rücu ve sovtaj gelirleri,
  reasürans payı, komisyon ve genel giderler hesaba girmez. Aktüeryal veya mali tablo H/P oranıyla aynı değildir.
- **Muallak tahmindir.** Dosyadaki muallak tutar değiştikçe oran değişir.
- **Kısa geçmiş yanıltır.** Yeni başlamış poliçede tek hasar oranı çok yükseltir; tek yıllık veri yerine mümkünse
  birkaç yılın poliçelerini birlikte verin.
- **Eşikler şirket politikasıdır.** %40 ve %80 yasal veya sektörel bir ölçüt değildir; branşa göre farklı eşik
  gerekiyorsa branşları ayrı çalıştırın.
- **Öneri karar değildir.** Yenileme şartlarını yetkili underwriter belirler.
- **Zeyiller:** Her poliçe tek satır olmalıdır; prim artış / azalış zeyillerini poliçe primine ekleyin.
- **Örnek veri:** Müşteriler, acenteler, primler ve hasarlar kurgusaldır.

## Testler

- Kazanılmış prim: başlamamış, süresi dolmuş ve kısmi dönem.
- Müşteri, acente, branş ve toplam H/P; kazanılmış primi olmayan müşteri.
- Öneri eşikleri (sınır değerler), sık hasar notu, yenilemesi yaklaşan poliçeler.
- Okuma hataları ve örnek veri uçtan uca.

```bash
python -m unittest discover -s tests
```

## Sorumluluk

Bu paketi indiren ve çalıştıran kişi aşağıdakileri kabul etmiş sayılır:

- **Tüm risk kullanıcıdadır.** Paket ücretsizdir ve "olduğu gibi" sunulur; doğruluk, eksiksizlik veya belirli bir amaca
  uygunluk konusunda hiçbir garanti verilmez. Paketi kendi bilgisayarınızda, kendi verinizle ve kendi kararınızla
  çalıştırırsınız. Hatalı veya eksik sonuç, kaçırılan süre, vergi / SGK cezası, veri kaybı, mali kayıp ve diğer her türlü
  doğrudan ya da dolaylı zarardan yazarlar ve katkıda bulunanlar sorumlu tutulamaz.
- **Sonuçları kontrol etmek sizin işinizdir.** Çıktılar karar desteği içindir. Beyan vermeden, ödeme yapmadan, sözleşme
  imzalamadan, personel / kredi / hasar / tedavi kararı almadan önce sonuçları yetkin bir kişi kontrol etmelidir.
- **Danışmanlık değildir.** Hukuki, mali, vergisel, tıbbi veya yatırım danışmanlığı yerine geçmez.
- **Mevzuat değişir.** Koddaki süre, oran, eşik ve kurallar yazıldığı tarihteki bilgiye dayanır ve eskiyebilir; özel
  durumları kapsamayabilir. Güncelliğini kontrol etmek kullanıcının sorumluluğundadır.
- **Örnek veri kurgusaldır.** Kişi, kurum, tutar ve tarihler gerçek değildir; kendi verinizle sonuçlar farklı olur.
- **Kişisel veri.** Verinizin KVKK, şirket politikaları ve sektör düzenlemelerine uygun işlenmesinden, saklanmasından ve
  paylaşılmasından siz sorumlusunuz.

Ayrıntılar: `LICENSE` (MIT) ve `SORUMLULUK_REDDI.md`.
