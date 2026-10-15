# Müşteri Haftalık Durum Raporu · Kod Bloğu

> Müşteri Temsilciliği (Merchandising) › Production Merchandiser · Workers / Workless

Siparişlerin **üretim aşaması ve sevk tarihi durumunu**, müşteriye gönderilecek haftalık rapor biçimine çevirir:
müşteri bazında özet, sipariş sipariş durum satırları ve Türkçe veya İngilizce rapor metni. İnternete bağlanmaz.

## Nasıl hesaplar?

| Durum (sırayla denenir) | Koşul |
|---|---|
| Sevk edildi | Aşama "Sevk edildi" veya sevk edilen adet ≥ sipariş adedi |
| Kısmi sevk | 0 < sevk edilen adet < sipariş adedi |
| Gecikmeli | Güncel (tahmini) sevk tarihi planlanandan sonra; veya güncel tarih yok ve planlanan tarih geçmiş |
| Riskli | Planlanan sevke `--risk-gun` (varsayılan 7) gün veya daha az kaldı ve sipariş `--risk-asama` (varsayılan Dikim) veya öncesinde |
| Zamanında | Diğerleri |

- **Gecikme günü:** güncel sevk − planlanan sevk; güncel tarih yoksa rapor tarihi − planlanan sevk.
- **Aşama sırası:** Kumaş bekleniyor → Kesim → Dikim → Yıkama → Ütü-paket → Final kontrol → Sevk edildi. Aşama adı bu
  sözcüklerden birini içeriyorsa tanınır; tanınmayan aşama (ör. Nakış) rapora yazıldığı gibi alınır ama risk
  değerlendirmesine girmez.
- **Rapor metni:** `--dil en` ile aşama ve durum adları İngilizce yazılır. Güncel tarihi olmayan gecikmeli siparişte
  tarih yerine "bildirilecek" / "to be confirmed" yazılır.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri (rapor tarihi 12.10.2026)
python main.py --siparisler siparisler.xlsx --dil en --imza "Best regards, ..."
python main.py --risk-gun 10 --risk-asama Yıkama
```

Sütunlar: Sipariş No, Planlanan Sevk (zorunlu); Müşteri, Model, Renk, Sipariş Adedi, Aşama, Tahmini Sevk, Sevk Edilen
Adet, Not.

## Çıktı

`cikti/haftalik_durum_raporu.xlsx`:

- `Sipariş Durumu`: müşteri ve öncelik sırasıyla siparişler; kalan gün, gecikme, durum; boş "İç Not" sütunu.
- `Müşteri Özeti`: müşteri başına sipariş sayısı ve durum dağılımı.
- `Uyarılar`: tanınmayan aşamalar, okunamayan tarihler.

`cikti/haftalik_durum_raporu.txt`: müşteri bazında rapor metni (e-postaya yapıştırılabilir).

Örnek metin:

```
=== Deneme Giyim · Haftalık Durum Raporu · 12.10.2026 ===
Toplam 4 sipariş: 1 gecikmeli, 1 riskli, 1 zamanında, 1 sevk edildi.

PO-503 · GM-2002 · Lacivert · 500 adet
  Aşama: Yıkama · Planlanan sevk: 15.10.2026 · Güncel sevk: 22.10.2026 · Durum: Gecikmeli (7 gün)
  Not: Yıkama tekrarı yapılıyor
```

## Dikkat

- **Göndermeden önce okuyun.** Metin taslaktır; "Not" sütunundaki her şey müşteriye giden metne yazılır. Müşteriye
  gitmemesi gereken açıklamaları Excel'deki "İç Not" sütununa yazın, girdi dosyasındaki "Not" sütununa değil.
- **Gecikme taahhüt değildir.** Güncel sevk tarihi sizin girdiğiniz tahmindir; kod üretim hızından tarih hesaplamaz.
- **Risk kuralı kabadır.** Yalnız kalan güne ve aşamaya bakar; sipariş adedi, hat kapasitesi ve kalan iş miktarı
  dikkate alınmaz.
- **Tek rapor dosyası, birden çok müşteri:** Metin dosyasında her müşterinin bölümü ayrıdır; bir müşteriye
  başkasının bölümünü göndermeyin.
- **Örnek veri:** Müşteriler, siparişler ve tarihler kurgusaldır.

## Testler

- Aşama tanıma; durum kuralları ve sınırları (sevk günü, 7. ve 8. gün, risk aşaması).
- Müşteri özeti, sıralama, Türkçe ve İngilizce rapor metni.
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
