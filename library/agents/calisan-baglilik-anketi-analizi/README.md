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

---
Bu paket "olduğu gibi" sunulur. API kullanım ücretleri ve modele gönderilen verinin hukuka uygunluğu
kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
