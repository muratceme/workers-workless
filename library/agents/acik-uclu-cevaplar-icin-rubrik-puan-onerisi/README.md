# Açık Uçlu Cevaplar İçin Rubrik Puan Önerisi · AI Agent

> Eğitim › Akademik › Ölçme Değerlendirme Uzmanı · Workers / Workless

Açık uçlu sınav cevaplarını **rubriğe (dereceli puanlama anahtarı)** göre puanlamak için öneri hazırlar:
- her ölçüt için rubrikteki bir puan düzeyi,
- öğrencinin cevabından kelimesi kelimesine kanıt ve gerekçe,
- öğrenciye kısa geri bildirim.

**Puanı kod toplar; son kararı öğretmen verir.** Aynı cevaba aynı puanın verilip verilmediğini kod denetler.

## Nasıl çalışır?

1. **Hazırlık (kod):** Sorular, rubrik ve cevaplar okunur.
   - Öğrenci adı ve numarası modele gönderilmez; telefon, e-posta, TCKN maskelenir.
   - Boş cevap modele gönderilmez; kod 0 puan verir.
2. **Puan önerisi (model):** Her ölçüt için rubrik düzeyi, kanıt ve gerekçe; geri bildirim ve gerekirse öğretmene not.
3. **Denetim (kod):** Aşağıdakilerden biri varsa yalnız o cevap modele geri gönderilir (varsayılan en fazla 2 tur):
   - Puanlanmayan, rubrikte olmayan ya da tekrar eden ölçüt
   - Rubrikte olmayan puan (ara puan)
   - En düşük düzeyin üstündeki puanın kanıtı cevapta aynen geçmiyor
   - Gerekçe boş
   - **Tutarlılık:** aynı soruya birebir aynı cevabı veren öğrencilere farklı ölçüt puanı
4. **Hesap (kod):**
   - cevap başına toplam ve azami puan,
   - öğrenci × soru tablosu,
   - kazanım bazında sınıf başarı yüzdesi.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                              # örnek 2 soru, 3 öğrenci (1 boş cevap)
python agent.py --sorular sorular.xlsx --rubrik rubrik.xlsx --cevaplar cevaplar.xlsx
```

Rubrik örneği:

```
Soru;Ölçüt;Puan;Tanım
1;Etmenler;0;Doğru etmen yok
1;Etmenler;1;Bir ya da iki etmen doğru
1;Etmenler;2;Üç etmen de doğru (ışık, su, karbondioksit)
```

## Çıktı

`cikti/rubrik_puan_onerileri.xlsx`:
- `Puan Önerileri`: cevap, ölçüt puanları ve kanıtları, önerilen puan, **boş "Öğretmen Puanı" ve "Öğretmen Açıklaması"
  sütunları**, geri bildirim, öğretmene not, son kontrol (mor: model).
- `Öğrenci × Soru`: önerilen puanlarla tablo.
- `Kazanım Özeti`: kazanım bazında alınan / azami puan ve başarı yüzdesi.
- `Kontrol`

## Dikkat

- **Puanlar öneridir.** Not kararını öğretmen verir; "Öğretmen Puanı" sütununu doldurun. Sonuç öğrenciye ve veliye
  öğretmen onayı olmadan bildirilmemelidir.
- **Rubrik kalitesi sonucu belirler.** Düzey tanımları açık ve gözlenebilir olmalı. Rubrikte karşılığı olmayan doğru
  cevaplar "öğretmene not" olarak gelir.
- **Kişisel veri:** Öğrenci numarasını gerekirse takma numarayla verin. Cevap metninde ad geçiyorsa çıkarın.
- **Örnek veri:** Sorular, rubrik ve cevaplar kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk turda 103 numaralı öğrencinin, 101 ile birebir aynı cevabına farklı puan
verir. Kodun tutarlılık denetimiyle bunu yakalayıp düzelttirdiği doğrulanır. Ayrıca şunlar doğrulanır:
- ara puan, kanıt, gerekçe ve ölçüt denetimleri,
- toplam ve kazanım yüzdesi,
- boş cevabın gönderilmemesi,
- öğrenci numarasının gönderilmemesi.

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
- **Yapay zekâ hata yapabilir.** Model yanlış, eksik veya uydurma bilgi üretebilir. Kod kontrolleri bazı hataları
  yakalar, hepsini yakalayamaz. Her taslağı okuyup onaylamadan kullanmayın, göndermeyin, imzalamayın.
- **Maskeleme sınırlıdır.** Yalnız telefon, e-posta, TCKN, IBAN ve LinkedIn adresi kalıpları maskelenir. Ad soyad,
  adres, doğum tarihi, sağlık bilgisi ve serbest metindeki diğer kişisel veriler modele gidebilir. Hangi verinin bulut
  sağlayıcısına gönderildiğinden siz sorumlusunuz; gizli veride yerel model (Ollama) kullanın.
- **API ücretleri, kotalar ve API anahtarının güvenliği** size aittir.

Ayrıntılar: `LICENSE` (MIT) ve `SORUMLULUK_REDDI.md`.
