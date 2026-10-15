# Çalışan Bağlılık Anketi Analizi · AI Agent

> İnsan Kaynakları › İnsan Kaynakları Müdürü · Workers / Workless

Anket puanlarını boyut ve birim bazında analiz eder. Açık uçlu yorumları temalara ayırır ve kurumun ele alması
gereken **temel sorun başlıklarını** özetler. Mobbing, iş güvenliği gibi **hassas yorumları** ayrı bir sayfada
toplar.

## Nasıl çalışır?

1. **Kod (puanlar):**
   - Likert (1–5) sorular boyutlara toplanır. Ters maddeler 6 − puan olarak çevrilir; örnek: "İş yüküm fazla".
   - Boyut ve birim bazında ortalama, olumlu (4–5) ve olumsuz (1–2) oranı hesaplanır.
   - 0–10 tavsiye sorusundan eNPS hesaplanır: destekçi (9–10) % − kötüleyen (0–6) %.
   - **Anonimlik:** `--min-yanit` (varsayılan 5) altındaki birimler "Diğer (küçük birimler)" altında birleştirilir.
     Bu grup da küçükse birim tablolarında gösterilmez.
   - Şirket ortalamasının 15 puan altındaki birim skorları ve olumlu oranı %40'ın altındaki boyutlar işaretlenir.
   - Birim çalışan sayıları verilirse katılım oranı hesaplanır.
2. **Model (yorumlar):**
   - Yorumlar maskelenerek gönderilir: e-posta, telefon, TCKN ve "Ahmet Bey" gibi unvanlı adlar. Birim bilgisi
     gönderilmez.
   - Model her yorumu sabit tema listesine göre etiketler; duygu, öneri ve hassas durum çıkarır.
   - Model her yorum için birebir kısa bir alıntı verir. Kod alıntının yorumda geçtiğini doğrular; geçmiyorsa
     alıntıyı çıkarır.
3. **Kod + model (sorun başlıkları):**
   - Kod tema sayılarını hesaplar. Model bu özetten 3–6 sorun başlığı yazar; her başlık yorum kimliklerine
     dayanmalıdır.
   - Geçerli kimliğe dayanmayan başlık rapora alınmaz. Başlıklardaki sayılar koddan gelir.
4. **Güvenlik ağı:** Kod ayrıca anahtar kelimeyle hassas yorum taraması yapar (bağırma, hakaret, taciz, kaza, fren,
   iş güvenliği...). Modelin kaçırdığı yorumlar da insan incelemesine düşer.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                    # örnek: 46 yanıt, 5 birim, 17 yorum
python agent.py --anket anket.xlsx --sorular sorular.xlsx --min-yanit 5
python agent.py --anket anket.xlsx --sorular sorular.xlsx --calisan-sayisi Üretim=25 Depo=12 Satış=10
python agent.py --anket anket.xlsx --sorular sorular.xlsx --yorum-yok           # yalnız puan analizi, API gerekmez
```

| Dosya | Sütunlar |
|---|---|
| Anket | Yanıt No, Birim, Kıdem (isteğe bağlı), soru kodu sütunları (S01, S02…), başlığında "görüş", "yorum" veya "öneri" geçen sütunlar |
| Sorular | Soru Kodu, Soru, Boyut, Ters Madde (Evet / Hayır). Tavsiye sorusunun boyutu `eNPS` yazılır (0–10). |

Form araçlarından (Google Forms, Microsoft Forms) alınan çıktıda soru başlıklarını soru kodlarıyla değiştirin.

## Çıktı

`baglilik_anketi_analizi.xlsx`:
- `Özet`: boyut skorları (renkli), yanıt sayısı, katılım, eNPS.
- `Hassas Yorumlar`: model değerlendirmesi ve anahtar kelime; boş "Yapılacak / Sorumlu" sütunu.
- `Birim × Boyut`: olumlu oran ısı haritası, eNPS, yanıt ve katılım.
- `Sorular`: soru bazında ortalama ve olumlu / olumsuz oran.
- `Sorun Başlıkları` (model): açıklama, dayanak yorum sayısı, doğrulanmış alıntılar, olası aksiyon; boş
  "Sorumlu / Karar" sütunu.
- `Temalar`: tema × birim yorum sayıları.
- `Yorumlar`: maskeli metin, temalar, duygu, öneri, alıntı, kontrol.
- `Uyarılar`.

## Dikkat

- **Anonimlik:** Rapor, küçük birimlerdeki yanıtlayanların tanınmasını önlemek için birleştirme yapar. Raporu
  paylaşırken de bunu koruyun. Yorumlardan yazanı bulmaya çalışmayın.
- **Hassas yorumlar:** Mobbing, ayrımcılık ve iş güvenliği bildirimleri ilgili sürece yönlendirilmelidir: etik
  kurul, İSG kurulu veya işyeri hekimi / İSG uzmanı. Model ve anahtar kelime taraması bir ön elemedir; tüm yorumları
  bir insan okumalıdır.
- **Model çıktısı:** Temalar ve başlıklar taslaktır. Aksiyon kararları yönetimindir.
- **Kişisel veri:** Maskeleme yardımcıdır, garanti değildir. Yorumların API'ye gönderilmesinin KVKK açısından
  uygunluğu (açık rıza / aydınlatma, yurt dışına aktarım) sizin sorumluluğunuzdadır. Yerel model için
  `WW_PROVIDER=ollama` kullanabilirsiniz.

## Testler

Testler gerçek API çağırmaz. Şunlar test edilir:
- Puan analizi: ters madde, eNPS, geçersiz puan, anonim birleştirme, katılım.
- Maskeleme ve anahtar kelime taraması.
- Uydurma tema ve kimliklerin elenmesi; alıntı doğrulaması; dayanaksız başlığın çıkarılması.
- Birim bilgisinin modele gönderilmemesi, Excel ve CLI.

```bash
python -m unittest discover -s tests
```

## Kaynaklar ve dayanaklar

Kontrol tarihi: 15.10.2026. Mevzuat sık değişir; paketteki süre, oran ve kuralları kullanmadan önce güncel resmî metinle
karşılaştırın. "Karşılaştırılan maddeler" dışındaki atıflar yalnız yol göstericidir.

| Kaynak | Bağlantı | Durum |
|---|---|---|
| 6698 sayılı Kişisel Verilerin Korunması Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.6698.pdf) | Resmî metinle karşılaştırılan maddeler: 13 |

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
