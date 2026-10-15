# İş İlanı Metni Hazırlama · AI Agent

> İnsan Kaynakları › İşe Alım · Workers / Workless

Departmanın doldurduğu **pozisyon talep formundan** yayına hazır iş ilanı hazırlar. İlan **ayrımcı ifade
içermez**. Talep formundaki "25-35 yaş arası, askerliğini yapmış erkek aday, hoş görünümlü" gibi riskli şartlar
kodla bulunur ve ilandan çıkarılır. İşin gereği olan şartlar (ör. sahada araç kullanılacaksa sürücü belgesi)
gerekçesiyle yazılır.

## Nasıl çalışır?

1. **Tarama (kod):** Talep formu ortak tarayıcıyla (`ayrimcilik.py`) taranır. İki seviye bulgu vardır:
   - **Yüksek risk:** yaş, cinsiyet, medeni hâl ve aile, görünüş, sağlık ve engellilik, din ve inanç, etnik
     köken, siyasi görüş ve sendika
   - **Dikkat:** askerlik, ehliyet, uyruk, anadili, gereksiz kişisel veri (fotoğraf, TC kimlik no, kan grubu),
     adli sicil, sigara
2. **Yazım (model):** İlanda başlık, giriş, sorumluluklar, aranan nitelikler, tercih sebepleri, sunulanlar,
   çalışma bilgisi ve başvuru bölümü (KVKK aydınlatma bağlantısı yer tutucusuyla) yer alır. Ayrıca kısa sosyal
   medya versiyonu yazılır.
   - Ayrımcı şartlar çıkarılır ve listelenir.
   - Kişilik özellikleri gözlemlenebilir yetkinliğe çevrilir ("güler yüzlü" → "müşteri ilişkilerinde güçlü
     iletişim").
3. **Denetim (kod):** Üretilen ilan yeniden taranır. Şu durumlarda taslak modele geri gönderilir (varsayılan en
   fazla 2 tur):
   - Yüksek riskli ifade kalmışsa
   - "Dikkat" ifadesi gerekçesiz kullanılmışsa
   - Gereksiz kişisel veri isteniyorsa

**Dayanak:** 4857 sayılı İş Kanunu md. 5 (eşit davranma), 6701 sayılı TİHEK Kanunu md. 3, 6 ve 7, 6356 sayılı
Kanun md. 25 (sendika), KVKK md. 4 (ölçülülük). TİHEK iş ilanlarındaki ayrımcı şartlara idari para cezası
vermektedir.

> Tarama kural tabanlıdır ve **hukuki değerlendirme yerine geçmez**. Yayımlamadan önce İK ve gerekiyorsa hukuk
> onayı alın.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                            # örnek talep formuyla (saha satış temsilcisi)
python agent.py --girdi talep_formu.docx --tur 3
```

## Çıktı

`İlan` (ilan metni, kısa versiyon, son kontrol) · `Uygunluk` (talep formunda bulunanlar, ilandan çıkarılanlar,
gerekçeyle tutulanlar, kalan sorunlar, "Onay") · `.md` ilan

## Testler

Testler gerçek API çağırmaz. Sahte model ilk taslakta bilerek "hoş görünümlü" bırakır; kodun bunu yakalayıp
düzelttirdiği doğrulanır.

```bash
python -m unittest discover -s tests
```

## Kaynaklar ve dayanaklar

Kontrol tarihi: 15.10.2026. Mevzuat sık değişir; paketteki süre, oran ve kuralları kullanmadan önce güncel resmî metinle
karşılaştırın. "Karşılaştırılan maddeler" dışındaki atıflar yalnız yol göstericidir.

| Kaynak | Bağlantı | Durum |
|---|---|---|
| 4857 sayılı İş Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.4857.pdf) | Resmî metinle karşılaştırılan maddeler: 17, 41, 46, 47, 53, 56, 63, 68, 69 |
| 6698 sayılı Kişisel Verilerin Korunması Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.6698.pdf) | Resmî metinle karşılaştırılan maddeler: 13 |
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
