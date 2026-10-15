# Stok Yaşlandırma Raporu · Kod Bloğu

> Lojistik ve Depo › Depo Sorumlusu · Workers / Workless

Stok listenizi **son hareket tarihine göre yaşlandırır**: hangi stok ne kadar süredir çıkış görmedi, hareketsiz
stokta ne kadar para bağlı, hangi depo ve kategoride yoğunlaşıyor. İnternete bağlanmaz.

## Nasıl hesaplar?

| Kavram | Kural |
|---|---|
| Stok tutarı | Miktar × birim maliyet |
| Yaş (gün) | Durum tarihi − son çıkış tarihi; son çıkış yoksa son giriş tarihi |
| `--esas hareket` | Son çıkış ve son girişten yeni olan tarih esas alınır |
| Yaş dilimi | 0–30, 31–60, 61–90, 91–180, 181–365, 366+ gün (`--dilimler`) |
| Durum | Yaş `--hareketsiz` (180) ve üstü: Hareketsiz · `--yavas` (90) ve üstü: Yavaş · altı: Normal |
| Pay % | Dilimin ya da durumun tutarı ÷ toplam stok tutarı × 100 |

Miktarı sıfır olan satırlar rapora girmez; eksi miktar uyarılır ve dışarıda bırakılır.

**Örnek:** 12 adet × 7.500 = **90.000,00**; son çıkış 15.02.2026, durum tarihi 12.10.2026 → yaş **239 gün**, dilim
181–365, durum **Hareketsiz**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                            # örnek veri, durum tarihi 12.10.2026
python main.py --stok stok.xlsx
python main.py --stok stok.xlsx --dilimler "30;90;180;365" --yavas 60 --hareketsiz 120 --esas hareket
```

Zorunlu sütunlar: Stok Kodu, Miktar. Aynı stok birden çok depodaysa her depo ayrı satır olmalıdır.

## Çıktı

`cikti/stok_yaslandirma.xlsx`:

- `Özet`: yaş dilimlerine ve duruma göre kalem sayısı, tutar ve pay.
- `Depolar`, `Kategoriler`: toplam tutar, yavaş dönen ve hareketsiz tutar, hareketsiz payı.
- `Hareketsiz Stok`: tutara göre büyükten küçüğe; "Karar" sütununu siz doldurursunuz (iade, indirimli satış, hurda).
- `Stoklar`: bütün kalemler, en yaşlıdan başlayarak.
- `Uyarılar`: hiç çıkışı olmayan stok, tarihi ya da maliyeti eksik satırlar, eksi stok.

## Dikkat

- **Yaş, son hareketin yaşıdır; partinin yaşı değil.** Geçen hafta tek bir adet çıkış görmüş stokun tamamı "yeni"
  görünür. Parti ya da giriş tarihi bazlı (FIFO katmanlı) yaşlandırma bu paketin kapsamında değildir.
- **Eşikler yönetim tercihidir.** 90 ve 180 gün varsayılandır; mevsimlik ürün, yedek parça ve hızlı tüketim için
  aynı eşik anlamlı olmaz. Kategoriye göre ayrı çalıştırmayı düşünün.
- **Değer düşüklüğü karşılığı hesaplanmaz.** Rapor hareketsiz stoku gösterir; net gerçekleşebilir değer tespiti ve
  karşılık ayrılması muhasebe değerlendirmesidir.
- **Tutar, yazdığınız birim maliyet kadardır.** Maliyet yönteminiz (ortalama, FIFO) ve maliyetin güncelliği sonucu
  doğrudan etkiler. Tutarlar tek para biriminde olmalıdır.
- **Kritik yedek parça hareketsiz görünebilir.** Nadiren kullanılan ama bulundurulması gereken parçaları kararda
  ayrı tutun.
- **Kapsam dışı:** stok devir hızı, stokta kalma süresi tahmini, son kullanma tarihi takibi, sipariş önerisi.
- **Örnek veri:** Stoklar, depolar ve maliyetler kurgusaldır.

## Testler

- Yaş dilimi ve durum sınırları (30/31, 89/90, 179/180 gün), özel dilim ve eşikler, ileri tarihli hareket.
- Son çıkışı olmayan stok, iki tarihi de boş satır, `--esas hareket`, sıfır ve eksi miktar, maliyeti boş satır.
- Dilim, durum, depo ve kategori özetleri, pay yüzdeleri, hareketsiz liste sırası, okuma ve komut satırı.

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
