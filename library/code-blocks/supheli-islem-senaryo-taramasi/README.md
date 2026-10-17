# Şüpheli İşlem Senaryo Taraması · Kod Bloğu

> Bankacılık › Kurumsal Uyum › Uyum Uzmanı · Gerisi Bende

İşlem verisini yapılandırılabilir **senaryolarla** tarar ve uyum biriminin inceleyeceği, öncelik sıralı bir
**inceleme listesi** çıkarır. İnternete bağlanmaz.

> **Bu araç bir işlemin şüpheli olup olmadığına karar vermez.** Yalnızca incelenmesi gereken örüntüleri öne
> çıkarır; karar, uyum görevlisinin değerlendirmesiyle verilir. Şüpheli işlem bildirimi yapıldığı veya yapılacağı
> bilgisi, işleme taraf olanlar dahil kimseye açıklanamaz (5549 sayılı Kanun md. 4/2).

## Senaryolar

| Kod | Senaryo | Örüntü |
|---|---|---|
| S1 | Parçalı işlem | Eşiğin hemen altında, kısa sürede tekrarlanan nakit işlemler; toplamı eşiği aşıyor |
| S2 | Hızlı geçiş | Gelen fonun 48 saat içinde büyük oranda çıkması (geçiş hesabı) |
| S3 | Nakit yoğunluğu | Girişlerde nakdin payı ve tutarı yüksek |
| S4 | Riskli ülke | Karşı taraf ülkesi kullanıcının verdiği listede |
| S5 | Çok sayıda gönderen | Kısa sürede çok sayıda farklı kişiden para girişi (hesap kiralama / toplama hesabı) |
| S6 | Profil dışı tutar | Beyan gelirine veya müşterinin kendi giriş/çıkış geçmişine göre olağandışı işlem |
| S7 | Uyuyan hesap | Uzun hareketsizlikten sonra ani yüksek hareket |
| S8 | Anahtar kelime | Açıklamada veya karşı taraf adında listedeki kelimeler (ör. yasa dışı bahis terimleri) |
| S9 | Yuvarlak tutar | Kısa sürede çok sayıda yuvarlak tutarlı işlem |

**Ayarlar (`senaryolar.json`)**
- **Eşikler ve ağırlıklar:** Her senaryonun eşikleri, zaman pencereleri ve ağırlığı bu dosyada tanımlıdır.
  Senaryolar tek tek kapatılabilir.
- **Kodda yasal eşik yoktur:** MASAK eşikleri ve tebliğleri değiştiği, ayrıca yükümlü grubuna göre farklılaştığı
  için hiçbir yasal eşik koda yazılmamıştır. Dosyadaki değerler **örnektir**; kurumunuzun risk politikasına ve
  güncel düzenlemelere göre değiştirin.

**Müşteri puanı**
- **Hesaplama:** Alarm veren her farklı senaryonun ağırlığı toplanır (en çok 100). Aynı senaryonun tekrarı puanı
  artırmaz.
- **Öncelik:** 60 ve üzeri yüksek, 30 ve üzeri orta öncelik sayılır.

**Diğer ayrıntılar**
- **Döviz:** Dövizli işlemler `--kur` ile TL'ye çevrilir.
- **Riskli ülke listesi:** FATF'nin güncel listelerine ve kurumunuzun ülke risk politikasına göre kendiniz
  verirsiniz. FATF listeleri yılda üç kez güncellenir. Örnekteki liste Haziran 2026 genel kurulunun çağrı
  listesidir.

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                          # kurgusal örnek veriyle (6 müşteri, 125 işlem)
python main.py --islemler islemler.xlsx --musteriler musteriler.xlsx \
               --senaryolar senaryolar.json --riskli-ulkeler ulkeler.csv --kur USD=41,20 EUR=48,10
```

**Girdi dosyaları**
- **İşlemler:** İşlem No, Tarih (saatli), Müşteri No, İşlem Tipi, Tutar, Para Birimi, Karşı Taraf, Karşı Taraf
  Ülke ve Açıklama sütunlarını içerir.
- **İşlem yönü:** İşlem Tipi'ndeki "Nakit yatırma", "Gelen EFT", "Giden FAST", "SWIFT giden" gibi ifadelerden
  çıkarılır. Çıkarılamıyorsa `Yön` sütunu (B/A, Giriş/Çıkış) ekleyin.
- **Müşteriler:** Müşteri No, Ad/Unvan, Tip, Meslek, Beyan Aylık Gelir ve Risk Profili sütunlarını içerir. Bu dosya
  isteğe bağlıdır; verilmezse S6 yalnız işlem geçmişine göre çalışır.

## Çıktı

- **Müşteri Öncelik:** Puan, öncelik ve senaryolar listelenir. İnceleyen, karar (Kapat / İzle / Bildirim) ve
  gerekçe için boş sütunlar bulunur.
- **Alarmlar:** Her alarmın açıklaması ve ilgili işlem numaraları.
- **Alarmlı İşlemler:** İnceleme için işlem dökümü.
- **Senaryo Tanımları:** Kullanılan senaryolar ve ayarları.

## Dikkat

- **Kurallar ve yapay zekâ:** Kural tabanlı tarama kurumun tüm risklerini kapsamaz. Senaryoları kendi müşteri
  kitlenize göre ayarlayın ve yanlış alarm oranını izleyin.
- **Kişisel veri:** İşlem ve müşteri verisi kişisel veridir. Araç internete bağlanmaz. Çıktı dosyalarını yetkisiz
  kişilerle paylaşmayın.

## Testler

```bash
python -m unittest discover -s tests
```

## Kaynaklar ve dayanaklar

Kontrol tarihi: 15.10.2026. Mevzuat sık değişir; paketteki süre, oran ve kuralları kullanmadan önce güncel resmî metinle
karşılaştırın. "Karşılaştırılan maddeler" dışındaki atıflar yalnız yol göstericidir.

| Kaynak | Bağlantı | Durum |
|---|---|---|
| 5549 sayılı Suç Gelirlerinin Aklanmasının Önlenmesi Hakkında Kanun | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.5549.pdf) | Resmî metinle karşılaştırılan maddeler: 4 |

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
