# Uygulama Destek E-postası Sınıflandırma · AI Agent

> Sizden Gelenler › Abonelikli mobil uygulama · Gerisi Bende

Abonelikli bir mobil uygulamanın destek kutusuna gelen e-postaları **kategori, duygu ve aciliyete** ayırır,
**öncelik (P1–P3)** ve **ilk yanıt hedefi** hesaplar, kısa özet ve müşteriye **cevap taslağı** yazar. İade, hesap /
veri silme, yasal tehdit ve hesap güvenliği gibi hassas e-postaları **"Onay Bekleyenler"** sayfasına ayırır:
bunlarda son karar her zaman ekibindir. Agent hiçbir e-posta göndermez; yalnız taslak yazar.

## Nasıl çalışır?

1. **Kod (yapay zekâ yok):**
   - E-postaları okur: `.csv` / `.xlsx` dökümü, `.mbox` dosyası (ör. Google Takeout ile indirilen posta kutusu) ya da
     `.eml` dosyalarının bulunduğu klasör. Cevaplardaki alıntılanan eski yazışmayı ("... wrote:", "> ..." satırları) atar.
   - **Hassas konuları kendisi bulur:** iade / geri ödeme, hesap ya da veri silme (KVKK), yasal süreç (avukat, tüketici
     hakem heyeti, BTK, CİMER), "ödeme alındı ama abonelik açılmadı", hesap güvenliği (izinsiz giriş, şifre değişmiş).
   - Gönderen adresi modele gitmez; e-posta, telefon, IBAN ve TCKN maskelenir.
2. **Model:** Kategori, duygu, aciliyet, tek cümlelik özet ve cevap taslağı yazar. Söz vermemesi ("iade ettik",
   "hesabınız silindi") ve müşteriden şifre / kart bilgisi istememesi söylenir.
3. **Kod kuralları modelin üstündedir:**
   - İade bulunan e-postanın kategorisi model ne derse desin **İade talebi**, veri silme bulunanınki **Hesap / veri
     silme (KVKK)** olur.
   - İade, veri silme, yasal süreç ve güvenlik bulunan e-posta **insan onayı gerektirir**; model bunu kaldıramaz.
   - Onay gerektiren e-postanın taslağında söz verme ifadesi varsa taslak **kırmızı** işaretlenir ve terminalde uyarı
     verilir.

| Öncelik | Ne zaman? | İlk yanıt hedefi (varsayılan) |
|---|---|---|
| P1 | Ödeme alındı / abonelik açılmadı, hesap güvenliği, yasal tehdit; ya da ödeme / giriş sorunu ve yüksek aciliyet | 4 saat |
| P2 | Diğer onay gerektirenler (iade, veri silme), hata / çökme, yüksek aciliyet | 24 saat |
| P3 | Diğerleri (soru, öneri, geri bildirim) | 72 saat |

Saatleri `sla.json` ile değiştirebilirsiniz: `{ "P1": 2, "P2": 12, "P3": 48 }`.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                  # örnek 12 e-posta (kurgusal)
python agent.py --girdi destek_epostalari.csv --sla sla.json
python agent.py --girdi Destek.mbox --urun "burç ve fal uygulaması, Plus aboneliği"
python agent.py --girdi eml_klasoru/
```

`--urun` modele ürününüzü kısaca tanıtır; cevap taslakları buna göre yazılır.

## Çıktı

- `cikti/destek_epostalari.xlsx`: `Onay Bekleyenler` · `Tüm E-postalar` (önceliğe göre sıralı; mor sütunlar yapay
  zekâ, "Onay" sütunu sizin) · `Özet`
- `cikti/uygulama-destek-epostasi.ozet.json`: **Uygulama Yönetim Özeti** paketinin okuduğu günlük talep, P1 ve onay
  sayıları.

## Canlı bağlantı ve zamanlama

Bu paket dosyadan çalışır; posta kutunuza bağlanmaz. E-postaların her sabah otomatik çekilmesi (Gmail / Outlook /
Zendesk), taslakların destek aracına yazılması ve zamanlama n8n / Make gibi bir otomasyon aracıyla kurulur; bu
**kurulum hizmeti** kapsamındadır.

## Testler

Testler gerçek API çağırmaz. Sahte model bilerek iadeyi "Abonelik / ödeme" sayar ve "ücretiniz iade edilecek" diye
söz verir. Kodun kategoriyi düzelttiği, onaya ayırdığı, sözü yakaladığı, gönderen adresinin ve telefonun modele
gitmediği doğrulanır.

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
