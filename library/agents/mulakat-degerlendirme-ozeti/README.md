# Mülakat Değerlendirme Özeti · AI Agent

> İnsan Kaynakları › İşe Alım · Gerisi Bende

Birden fazla mülakatçının **puan ve notlarını** aday başına **tek sayfalık karşılaştırmalı özete** çevirir.
Kararı işe alım yöneticisi verir; bu araç kararı kanıta dayalı hâle getirir.

## Nasıl çalışır?

1. **Hesap (kod):**
   - Aday × yetkinlik ortalaması ve yetkinlik ağırlıklarıyla **toplam puan** (5 üzerinden), sıralama
   - Mülakatçılar arası **puan farkı ≥ 2** olan yetkinlikler: kalibrasyon toplantısında konuşulmalı
   - Eksik puanlar
2. **Önyargı kontrolü (kod):** Notlar ortak ayrımcılık tarayıcısından geçer. Korunan özelliğe (yaş, medeni hâl,
   aile, sağlık, din, köken...) dayalı yorum içeren notlar **"değerlendirme dışı"** işaretlenir. Örnek: "evli ve
   iki çocuklu olduğu için sahada zorlanabilir". Bu puanlar hariç tutularak **ikinci bir toplam** hesaplanır;
   sıralama bu toplama göre yapılır.
3. **Gizlilik:** Aday ve mülakatçı adları modele gitmeden **takma adla** (A1, M1) değiştirilir. Rapor şirket
   içinde kaldığı için raporda gerçek adlar gösterilir.
4. **Özet (model):** Her aday için şunlar yazılır:
   - Güçlü yönler ve gelişim alanları (notlardan somut kanıtla)
   - Görüş ayrılıkları ve kalibrasyon sorusu
   - Referansta doğrulanacak noktalar
   - Öneri taslağı (ilerlet / beklet / ilerletme) ve yalnız işe dayalı gerekçe

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                         # örnek: 2 aday, 2 mülakatçı, 5 yetkinlik
python agent.py --girdi mulakat_notlari.xlsx --yetkinlikler yetkinlikler.xlsx
```

**Girdi:** Her satır bir mülakatçının bir yetkinlik için verdiği puan (1–5) ve nottur. **Mülakat Soru Seti
Hazırlama** agent'ının değerlendirme formundaki puanlar bu biçime kolayca aktarılır.

## Çıktı

`Karşılaştırma` (uyumsuz puanlar kırmızı, "Karar" sütunu) · `Aday Özetleri` · `Notlar` (değerlendirme dışı
işaretleri) · `Bilgi`

## Testler

Testler gerçek API çağırmaz. Ağırlıklı toplamlar elle hesaplanmış değerlerle karşılaştırılır (3,80 / 3,325 /
3,45). Adların modele gitmediği ve önyargılı notun işaretlendiği doğrulanır.

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
