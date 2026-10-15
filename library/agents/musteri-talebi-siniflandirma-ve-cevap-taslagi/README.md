# Müşteri Talebi Sınıflandırma ve Cevap Taslağı · AI Agent

> Müşteri Hizmetleri ve Çağrı Merkezi › Müşteri Temsilcisi · Workers / Workless

E-posta, form ve sosyal medyadan gelen müşteri taleplerini **kategori, aciliyet, duygu ve ilgili ekibe** göre
sınıflandırır. Her talep için **yalnızca şirketinizin bilgi bankasına dayanan** bir cevap taslağı hazırlar.
Hiçbir cevap otomatik gönderilmez; her taslak için "Temsilci Onayı" sütunu vardır.

## Nasıl çalışır?

1. **Kod (yapay zekâ yok):**
   - Sipariş numaralarını bulur.
   - Aynı göndericinin tekrar eden taleplerini bağlar.
   - Kural etiketleri koyar: **KVKK başvurusu**, **hukuki risk** (avukat, dava, tüketici hakem heyeti,
     şikâyet sitesi, CİMER), **acil**, **tekrar yazıyor**.
   - E-posta, telefon, IBAN ve TCKN'yi maskeler. Gönderici adresi modele hiç gönderilmez.
2. **Model:** Her talep sınıflandırılır ve bilgi bankasına göre cevap taslağı yazılır. Model sipariş durumu,
   tarih, tutar, tazminat gibi bilgi bankasında olmayan bilgileri uyduramaz. Bilgi yoksa veya bir işlem
   gerekiyorsa talep **insan gerekli** olarak işaretlenir ve yapılacak işlem yazılır.
3. **Kod kuralları modelin üstündedir:**
   - KVKK ve hukuki risk etiketli talepler her zaman insana yönlendirilir ve en az "yüksek" aciliyet alır.
   - Tekrar eden talepler "düşük" aciliyette kalmaz.
   - KVKK başvurularına **30 günlük yasal cevap süresi** (KVKK md. 13) yazılır.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                      # örnek taleplerle
python agent.py --girdi talepler.xlsx --bilgi bilgi_bankasi.md
python agent.py --girdi ./gelen_kutusu --bilgi bilgi_bankasi.md      # .eml / .txt klasörü
python agent.py --girdi talepler.xlsx --bilgi bb.md --kategoriler kategoriler.json
```

- **Bilgi bankası:** Kargo, iade, değişim, fatura, KVKK ve iletişim kurallarınızı başlıklar hâlinde yazın
  (bkz. `ornek_veri/bilgi_bankasi.md`). Cevapların kalitesi bu dosyanın güncelliğine bağlıdır.
- **Kategoriler:** `{"kategoriler": [...], "ekipler": [...]}` biçiminde bir JSON dosyasıyla kendi listenizi verin.

## Çıktı

`Özet` (kategori, aciliyet ve ekip dağılımı) · `Talepler` (aciliyete göre sıralı, cevap taslakları ve yapılacak işlemler)

## Testler

Testler gerçek API çağırmaz. Sahte model bilerek hatalı sınıflandırma yapar ve kod kurallarının bunu
düzelttiği doğrulanır.

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
