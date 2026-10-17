# Öğrenci Gelişim Raporu Taslağı · AI Agent

> Akademik › Rehber Öğretmen · Gerisi Bende

Not, devamsızlık ve öğretmen gözlemlerinden **veli görüşmesi için öğrenci gelişim raporu taslağı** hazırlar. Puan
değişimlerini ve ortalamayı **kod hesaplar**. **Öğrenci adı ve okul numarası yapay zekâ modeline gönderilmez.**

## Nasıl çalışır?

1. **Hazırlık (kod):** Öğrenci başına şunlar hesaplanır:
   - Ders puanları ve önceki döneme göre değişim
   - Bu dönem ortalaması
   - Belirgin düşüşler (`--dusus` puan ve üzeri, varsayılan 10)

   Gözlemler G1, G2 … diye numaralanır; gözlemde geçen öğrenci adı `{OGRENCI}` ile değiştirilir.
2. **Taslak (model):** Güçlü yönler, gelişime açık alanlar, öneriler, veli görüşmesi soruları ve genel değerlendirme
   yazılır. Her madde dayandığı ders ya da gözlemi gösterir.
3. **Denetim (kod):** Aşağıdakilerden biri varsa yalnız o öğrencinin taslağı modele geri gönderilir (varsayılan en
   fazla 2 tur):
   - Dayanak öğrencinin dersi değil ya da gözlem numarası geçersiz
   - Öğrencinin puan, değişim, ortalama ve devamsızlık değerlerinde olmayan bir sayı var
   - Ortalamanın altındaki bir ders güçlü yön olarak gösterilmiş
   - Belirgin düşüş gösteren bir ders gelişime açık alanlarda yok
   - Tanı koyan ya da etiketleyen ifade var ("dikkat eksikliği", "hiperaktif", "disleksi", "tembel")
   - Başka öğrencilerle kıyas var ("sınıfın en …", "arkadaşlarından geri")
4. **Raporlar (kod):** Ad, sınıf ve numara geri yazılır; kontrolden geçen her öğrenci için .txt taslak üretilir.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                              # örnek sınıf
python agent.py --notlar notlar.xlsx --devamsizlik devamsizlik.xlsx --gozlemler gozlemler.txt --dusus 15
```

Not tablosunu e-Okul'dan aldığınız puanlarla hazırlayabilirsiniz: Öğrenci No, Ad Soyad, Sınıf, Ders, Önceki Dönem,
Bu Dönem. Gözlem dosyasında her satır `1021: gözlem (öğretmen branşı)` biçimindedir.

## Çıktı

- `cikti/gelisim_raporlari_<ogrenci_no>.txt`: öğrenci başına taslak.
- `cikti/gelisim_raporlari.xlsx` → `Özet`: ortalama, belirgin düşüş, devamsızlık, güçlü yönler, gelişime açık alanlar,
  kontrol sonucu.

## Dikkat

- **Taslaktır.** Rehber öğretmen gözden geçirmeden veliyle paylaşılmamalıdır.
- **Tanı konulmaz.** Rapor gözleme dayalıdır. Özel eğitim değerlendirmesi gerekiyorsa süreç okul idaresi ve veliyle
  birlikte, mevzuattaki yönlendirme yoluyla yürütülür.
- **Kişisel veri.** Öğrenci adı ve numarası modele gönderilmez. Puanlar, devamsızlık ve gözlemler gönderilir; gözlemlere
  başka öğrencilerin adlarını yazmayın. Çıktıları yalnız yetkili personel kullanmalıdır.
- **Örnek veri:** Öğrenciler, puanlar ve gözlemler kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk taslakta bilerek:
- ortalamanın altındaki Fizik'i güçlü yön yazar,
- 13 puan düşen Matematik'i gelişime açık alanlardan çıkarır,
- ortalamayı 79 yazar ve "dikkat eksikliği olabilir" der.

Kodun bunları yakaladığı, yalnız sorunlu öğrencinin taslağını yeniden yazdırdığı ve öğrenci adı / numarasının modele
gönderilmediği doğrulanır.

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
