# Karar ve İçtihat Özeti · AI Agent

> Dava Takip › Avukat · Gerisi Bende

Mahkeme kararını dosyaya ve içtihat arşivine girilecek biçimde **başlıklarla özetler**: künye, taraflar, olay, iddia ve
savunma, yargılama süreci, hukuki mesele, gerekçe, hüküm, oy durumu, atıf yapılan mevzuat ve kararlar, anahtar kelimeler.
Her bölüm karar metnindeki satırlara dayanır; numaralar ve hüküm alıntısı kodla metne karşı denetlenir.

## Nasıl çalışır?

1. **Hazırlık (kod):** Karar metni satır (paragraf) satır numaralanır.
2. **Özet (model):** Bölümler yazılır; her bölüm dayandığı satırları gösterir. Hüküm fıkrasından aynen alıntı yapılır.
3. **Denetim (kod):** Aşağıdakilerden biri varsa özet modele geri gönderilir (varsayılan en fazla 2 tur):
   - Esas no, karar no ya da karar tarihi metinde geçmiyor
   - Mevzuat maddesinin numarası dayandığı satırda yazmıyor
   - Atıf yapılan kararın esas / karar numarası metinde yok
   - Hüküm alıntısı metinde yok
   - Oy durumu metindeki "oybirliği / oyçokluğu" ifadesiyle çelişiyor
   - Dayanak satırı yok ya da metinde bulunmuyor

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                    # örnek (kurgusal) karar metni
python agent.py --girdi karar.pdf
python agent.py --girdi karar.docx --tur 3
```

## Çıktı

- `cikti/karar_ozeti.xlsx`:
  - `Karar Özeti`: bölümler ve dayanak satırları.
  - `Mevzuat ve Atıflar`
  - `Metin`: numaralı karar metni.
- `cikti/karar_ozeti.md`: arşive eklenecek özet.

## Dikkat

- **Özet hukuki görüş değildir.** Kullanmadan önce kararın tam metniyle ve resmî kaynakla (UYAP, Yargıtay / Danıştay karar arama)
  karşılaştırın. Kararın kesinleşip kesinleşmediğini ve sonradan değişen mevzuatı ayrıca kontrol edin.
- **Denetimin sınırı.** Kod numaraların ve alıntıların metinde geçtiğini doğrular; gerekçenin doğru özetlendiğini doğrulamaz.
- **Kişisel veri.** Karar metinleri taraf adları ve sağlık, ceza gibi özel nitelikli veriler içerebilir. Telefon, e-posta ve TCKN
  maskelenir; adlar maskelenmez. Gerekirse yerel model (Ollama) kullanın ya da anonimleştirilmiş metin verin.
- **Örnek veri.** Örnek karar metni kurgusaldır; daire, esas / karar numaraları ve atıf yapılan karar gerçek değildir.
  Mevzuat atıfları (TBK 49, 52; HMK 266) gerçek maddelerdir.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk taslakta bilerek karar numarasını yanlış yazar, bir kanun maddesini başka bir
satıra dayandırır ve oybirliğiyle verilen kararı "oyçokluğu" diye işaretler. Kodun bunları yakalayıp düzelttirdiği doğrulanır.

```bash
python -m unittest discover -s tests
```

## Kaynaklar ve dayanaklar

Kontrol tarihi: 15.10.2026. Mevzuat sık değişir; paketteki süre, oran ve kuralları kullanmadan önce güncel resmî metinle
karşılaştırın. "Karşılaştırılan maddeler" dışındaki atıflar yalnız yol göstericidir.

| Kaynak | Bağlantı | Durum |
|---|---|---|
| 6098 sayılı Türk Borçlar Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.6098.pdf) | Resmî metinle karşılaştırılan maddeler: 100, 115, 182, 315 |
| 6100 sayılı Hukuk Muhakemeleri Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.6100.pdf) | Resmî metinle karşılaştırılan maddeler: 127, 136, 281, 345, 347, 361 |

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
