# Sözleşme Ön İnceleme · AI Agent

> Hukuk › Avukat · Workers / Workless

Bir sözleşmeyi **sizin tarafınızın bakış açısından** okur ve şunları çıkarır:
- Riskli maddeler ve her birinin kimin lehine olduğu
- Gerekçe, müzakere önerisi ve alternatif madde metni
- 20 konuluk standart kontrol listesine göre **eksik** veya **belirsiz** düzenlemeler (mücbir sebep, KVKK,
  uygulanacak hukuk vb.)

> **Bu bir ön incelemedir, hukuki görüş değildir.** Rapordaki her öneri avukat tarafından değerlendirilmelidir
> ("Avukat Onayı" sütunu).

## Nasıl çalışır?

1. **Kod (yapay zekâ yok):** Sözleşme `.docx`, `.pdf` veya `.txt` dosyasından okunur ve "MADDE 5 – FESİH" ya da
   "5. FESİH" başlıklarından maddelere bölünür. Her madde konu anahtar kelimeleriyle etiketlenir.
2. **Maskeleme:** Modele gönderilmeden önce IBAN, TCKN, telefon ve e-posta maskelenir. `--gizle` ile verdiğiniz
   kişi ve firma adları `[GİZLİ-1]` gibi etiketlerle değiştirilir. Veri gönderilmeden önce onayınız alınır.
3. **Model:** Her madde için risk düzeyi (düşük / orta / yüksek), kimin lehine olduğu, gerekçe, öneri ve
   alternatif metin döner. Kontrol listesinde her konu için var / eksik / belirsiz durumu belirlenir.
4. **Rapor:** Excel ve kısa bir Markdown özeti üretilir. Uzun sözleşmeler bölümlere ayrılarak gönderilir.

Modelin yönergesi (`prompt.md`) yalnız doğrulanmış genel kurallara dayanmasını ister ve kanun maddesi numarası
uydurmasını yasaklar. Bu kurallar: TBK 115 (ağır kusurdan sorumsuzluk), TBK 182 / TTK 22 (ceza koşulu), TBK
20–25 (genel işlem koşulları), HMK 17 (yetki sözleşmesi), TTK 5/A (dava şartı arabuluculuk), damga vergisi,
KVKK 9 ve 12.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

`WW_PROVIDER=ollama` ile yerel bir model kullanırsanız sözleşme bilgisayarınızdan çıkmaz.

## Kullanım

```bash
python agent.py                                              # örnek sözleşmeyle (Müşteri adına)
python agent.py --girdi sozlesme.docx --taraf "Müşteri"
python agent.py --girdi sozlesme.pdf --taraf "Alıcı" --gizle "Ahmet Yılmaz" "Örnek A.Ş."
```

Taranmış (görüntü) PDF'lerden metin çıkarılamaz; önce OCR uygulayın. Eski `.doc` dosyalarını `.docx` olarak kaydedin.

## Çıktı

`Özet` (öncelikli aksiyonlar, genel değerlendirme) · `Madde Analizi` (yüksek riskten düşüğe) · `Kontrol Listesi`
· `Sözleşme Metni` · ayrıca aynı adla `.md` özeti

## Testler

Testler gerçek API çağırmaz; model yanıtı sahte bir fonksiyonla verilir.

```bash
python -m unittest discover -s tests
```

## Kaynaklar ve dayanaklar

Kontrol tarihi: 15.10.2026. Mevzuat sık değişir; paketteki süre, oran ve kuralları kullanmadan önce güncel resmî metinle
karşılaştırın. "Karşılaştırılan maddeler" dışındaki atıflar yalnız yol göstericidir.

| Kaynak | Bağlantı | Durum |
|---|---|---|
| 6102 sayılı Türk Ticaret Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.6102.pdf) | Resmî metinle karşılaştırılan maddeler: 5/A, 22, 796, 1434 |
| 6098 sayılı Türk Borçlar Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.6098.pdf) | Resmî metinle karşılaştırılan maddeler: 100, 115, 182, 315 |
| 6100 sayılı Hukuk Muhakemeleri Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.6100.pdf) | Resmî metinle karşılaştırılan maddeler: 127, 136, 281, 345, 347, 361 |
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
