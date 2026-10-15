# İade Gelen EFT/FAST Analizi · Kod Bloğu

> Bankacılık Operasyonları › Ödeme Sistemleri Operasyon Uzmanı · Workers / Workless

İade gelen ödemeleri **hata türü, şube ve kanala** göre sınıflandırır; tekrarlayan hata nedenlerini ve iade oranı
yüksek birimleri raporlar. İnternete bağlanmaz.

## Nasıl hesaplar?

1. **Hata türü:** İade açıklaması büyük-küçük harf ve Türkçe karakter ayırmadan taranır; ilk eşleşen türe atanır.

   | Hata türü | Örnek ifadeler | Nerede önlenir |
   |---|---|---|
   | Alıcı adı uyuşmuyor | alıcı adı, isim uyuşmuyor, alıcı ismi | Gönderen müşteri |
   | Hesap kapalı | hesap kapalı, kapanmış hesap | Gönderen müşteri |
   | IBAN / hesap numarası hatalı | IBAN hatalı, geçersiz IBAN, hesap bulunamadı | Şube girişi |
   | Blokeli / kısıtlı hesap | bloke, kısıtlı, haciz, tedbir | Alıcı taraf |
   | Mükerrer gönderim | mükerrer, çift gönderim | Şube girişi |
   | Alıcı talebiyle iade | alıcı talebi, kabul edilmedi | Alıcı taraf |
   | Tutar / limit | limit, tutar aşımı | Şube girişi |
   | Açıklama / zorunlu alan eksik | zorunlu alan, eksik bilgi, kimlik no | Şube girişi |
   | Para birimi / hesap türü uyumsuz | para birimi, döviz hesabı | Gönderen müşteri |
   | Teknik / sistemsel | teknik, sistem, zaman aşımı | Sistem |
   | Diğer | hiçbiri eşleşmezse | Belirsiz |

2. **Pareto:** Türler adede göre sıralanır; iadelerin ilk %80'ini oluşturanlar "A" olarak işaretlenir.
3. **İade oranı:** Gönderilen işlem adedi dosyası verilirse şube ve kanal bazında iade ÷ gönderilen. Oranı genel
   ortalamanın 2 katını aşan ve en az 5 iadesi olan şube için uyarı verilir.
4. **Tekrar:** Aynı gönderen müşteriden veya aynı alıcı IBAN'ına `--tekrar` (varsayılan 2) ve üzeri iade listelenir.
   Aynı neden en az 3 kez tekrarlamışsa uyarı verilir.
5. **İade süresi:** İade tarihi − gönderim tarihi (takvim günü); tür bazında ortalaması alınır.

"Diğer" payı %15'i aşarsa en sık görülen açıklamalar uyarıda listelenir.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek veri
python main.py --iadeler iadeler.xlsx --gonderilen gonderilen.xlsx
python main.py --iadeler iadeler.xlsx --hata-turleri hata_turleri.csv --tekrar 3
```

| Dosya | Sütunlar |
|---|---|
| İadeler | **İade Açıklaması** (zorunlu); Referans, Gönderim Tarihi, İade Tarihi, Şube, Kanal, Sistem, Gönderen Müşteri, Alıcı, Alıcı IBAN, Tutar, İade Kodu, Kullanıcı |
| Gönderilen (isteğe bağlı) | Şube, Kanal, Gönderilen Adet |
| Hata türleri (isteğe bağlı) | Hata Türü, Anahtar İfadeler (virgülle), Kaynak, Öneri |

Kendi hata türleriniz yerleşik listeden önce denenir.

## Çıktı

`cikti/iade_analizi.xlsx`:

- `Hata Türleri`: adet, pay, kümülatif pay, Pareto sınıfı, tutar, ortalama iade süresi, iade kodları, öneri.
- `Şube Bazında`, `Kanal Bazında`: iade adedi, oranı, giriş kaynaklı iade sayısı, en sık neden.
- `Tekrarlayanlar`: tekrar eden gönderen ve alıcılar; giriş kaynaklı iadelerde kullanıcı sayımı.
- `Günlük`: günlük iade adedi (ortalamanın 2 katını aşan günler işaretli).
- `İadeler`: her iadenin atandığı hata türü.

## Dikkat

- **Sınıflandırma açıklama metnine dayanır.** İade kodlarının anlamı ödeme sistemine ve bankaya göre değişir; kod
  sabit bir kod tablosu içermez. Açıklama boş veya genel yazılmışsa iade "Diğer" olur.
- **"Nerede önlenir" bir yönlendirmedir**, sorumluluk tespiti değildir. Örneğin IBAN hatası müşterinin verdiği
  bilgiden de kaynaklanabilir.
- **Kullanıcı sayımı** eğitim ihtiyacını göstermek içindir; işlem hacmine oranlanmadan performans ölçüsü olarak
  kullanılmamalıdır.
- **Kişisel veri:** İade listesi müşteri, alıcı ve IBAN bilgisi içerir; çıktıyı yalnız yetkili kişilerle paylaşın.
- **Örnek veri:** İade kodları, açıklamalar, müşteriler ve IBAN'lar kurgusaldır.

## Testler

- Açıklamadan hata türü bulma; kullanıcı tanımlı türlerin önceliği.
- Pay, kümülatif pay ve Pareto; iade oranı ve yüksek oran uyarısı.
- Tekrarlayan gönderen / alıcı; "Diğer" oranı uyarısı; kullanıcı sayımı.
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
