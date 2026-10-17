# Mülakat Soru Seti Hazırlama · AI Agent

> İnsan Kaynakları › İşe Alım · Gerisi Bende

Pozisyon bilgisinden ve yetkinlik listesinden **yapılandırılmış (yetkinlik bazlı) mülakat soru seti**
hazırlar. Mülakatçının dolduracağı **formüllü değerlendirme formunu** da üretir.

## Neler üretir?

- **Sorular:** Her yetkinlik için 2–3 soru: **davranışsal** (STAR ile yanıtlanan geçmiş deneyim), **durumsal**
  (senaryo) veya **teknik**. Her sorunun yanında şunlar yer alır:
  - Takip (sondaj) soruları
  - Dinlenecek **olumlu ve olumsuz göstergeler**
  - Davranışsal **1 – 3 – 5 puan tanımları**
  - Tahmini süre
- **Açılış ve kapanış metinleri:** Mülakatçının okuyacağı kısa metinler.
- **Değerlendirme formu (Excel):** Puan hücreleri sarıdır. Yetkinlik ortalaması ve **ağırlıklı toplam puan**
  formülle hesaplanır. Yalnız puanlanmış yetkinlikler toplama girer.

## Kod denetimi (model taslağının üstünde)

1. **Ayrımcı soru taraması:** Soru ve takip soruları ortak tarayıcıdan geçer (`ayrimcilik.py`). Taranan
   konular: yaş, medeni hâl, çocuk/aile planı, gebelik, sağlık, din, köken/memleket, siyasi görüş, sendika,
   askerlik, görünüş ve gereksiz kişisel veri.
2. **Kapsam:** Her yetkinliğin en az bir sorusu olmalıdır.
3. **Süre:** Toplam süre, açılış ve kapanış dahil (10 dk) mülakat süresini (`--sure`) aşmamalıdır.

Sorun varsa taslak modele geri gönderilir (varsayılan en fazla 2 tur). Kalanlar raporda işaretlenir.
**Dayanak:** 4857 sayılı İş Kanunu md. 5; 6701 sayılı Kanun md. 3 ve 6; 6356 sayılı Kanun md. 25.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                               # örnek: saha satış temsilcisi, 5 yetkinlik
python agent.py --pozisyon ilan.docx --yetkinlikler yetkinlikler.xlsx --sure 60
python agent.py --pozisyon talep_formu.txt --sure 45                           # yetkinlikleri model önersin
```

## Çıktı

`Değerlendirme Formu` · `Soru Seti` · `Açılış ve Kapanış` (sorulmaması gerekenler, son kontrol) · `.md` soru seti

## Testler

Testler gerçek API çağırmaz. Sahte model ilk taslakta bilerek medeni hâl sorusu sorar, bir yetkinliği atlar ve
süreyi aşar. Kodun bunları yakalayıp düzelttirdiği doğrulanır.

```bash
python -m unittest discover -s tests
```

## Kaynaklar ve dayanaklar

Kontrol tarihi: 15.10.2026. Mevzuat sık değişir; paketteki süre, oran ve kuralları kullanmadan önce güncel resmî metinle
karşılaştırın. "Karşılaştırılan maddeler" dışındaki atıflar yalnız yol göstericidir.

| Kaynak | Bağlantı | Durum |
|---|---|---|
| 4857 sayılı İş Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.4857.pdf) | Resmî metinle karşılaştırılan maddeler: 17, 41, 46, 47, 53, 56, 63, 68, 69 |
| 6701 sayılı Türkiye İnsan Hakları ve Eşitlik Kurumu Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.6701.pdf) | Resmî metin bağlantısı; paketteki madde atıfları ayrıca karşılaştırılmadı |
| 6356 sayılı Sendikalar ve Toplu İş Sözleşmesi Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.6356.pdf) | Resmî metin bağlantısı; paketteki madde atıfları ayrıca karşılaştırılmadı |

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
