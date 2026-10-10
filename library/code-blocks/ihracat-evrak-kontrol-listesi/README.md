# İhracat Evrak Kontrol Listesi · Kod Bloğu

> Dış Ticaret › Dış Ticaret Uzmanı · Workers / Workless

Sevkiyatın ülkesine, teslim şekline, taşıma türüne ve ödeme yöntemine göre **gereken ihracat evraklarını** listeler.
Hazırlanan evrakların bilgileri verilirse eksik evrakları ve evraklar arasındaki **tutarsızlıkları** (alıcı adı, koli
sayısı, ağırlık, tutar...) bulur. İnternete bağlanmaz.

## Nasıl çalışır?

### Gereken evraklar (varsayılan kural tablosu)

| Evrak | Ne zaman |
|---|---|
| Ticari fatura, çeki listesi, ihracat gümrük beyannamesi, ihracatçı birliği kaydı | Her ihracatta |
| Konşimento / CMR / hava yük senedi (AWB) / CIM | Taşıma: deniz / karayolu / havayolu / demiryolu |
| A.TR Dolaşım Belgesi | AB ülkesi; sanayi veya işlenmiş tarım ürünü (Gümrük Birliği) |
| EUR.1 / EUR-MED veya fatura beyanı | AB ülkesi ve tarım veya kömür-çelik ürünü; serbest ticaret anlaşması (STA) ülkesi |
| Menşe Şahadetnamesi | Diğer ülkeler (alıcı veya alıcı ülke istiyorsa zorunlu; listede "Önerilir") |
| Sigorta poliçesi / sertifikası | Teslim: CIF veya CIP |
| Akreditifte istenen belgeler | Ödeme: akreditif |
| Banka tahsil talimatı; poliçe (önerilir) | Ödeme: vesaik mukabili (poliçe akreditifte de) |
| Bitki sağlık / veteriner sağlık sertifikası | Bitkisel / hayvansal ürün |
| ISPM 15 ısıl işlem işareti | Ahşap ambalaj |
| Güvenlik bilgi formu (SDS) ve tehlikeli madde beyanı | Tehlikeli madde (IMDG / ADR / IATA DGR) |
| Dahilde İşleme İzin Belgesi taahhüt kaydı | Dahilde işleme |
| Analiz / sağlık sertifikası (önerilir) | Gıda |

- **Ülke grubu:** AB üyesi 27 ülke kodda tanımlıdır (Türkçe ve İngilizce adlarıyla). Serbest ticaret anlaşması olan
  ülkeler için sevkiyat bilgisine `Ülke Grubu;STA` yazın. Hiçbiri değilse "Diğer" kabul edilir ve uyarı verilir.
- **Kendi kurallarınız:** `--kurallar` dosyasıyla kural eklenir. Aynı adlı varsayılan kural değiştirilir; Koşul
  sütununa `kaldır` yazılırsa çıkarılır.
  - Koşul biçimi: `alan=değer1|değer2`, birden çok koşul `&` ile.
  - Alanlar: `ülke grubu`, `teslim`, `taşıma`, `ödeme`, `ürün türü`, `ahşap ambalaj`, `tehlikeli madde`... ve
    sevkiyat bilgisindeki diğer satırlar (ör. `ülke=suudi arabistan`).

### Tutarlılık kontrolü

Hazırlanan evrakların bilgileri `Evrak; Alan; Değer` biçiminde girilir. Aynı alan birden çok evrakta varsa
karşılaştırılır:
- Sayısal alanlar (tutar, koli, net / brüt ağırlık) 0,01 toleransla karşılaştırılır; "85.512,00" ile "85512" aynıdır.
- GTİP'te 6 haneli HS kodu ile 12 haneli GTİP uyumlu sayılır.
- Teslim şeklinde yalnız kural kodu karşılaştırılır ("CIF Hamburg" = "CIF").
- Metin alanlarında büyük-küçük harf ve noktalama farkı yok sayılır; harf farkı ("Textile" / "Textiles") tutarsızlıktır.

| Uyarı | Önem |
|---|---|
| Zorunlu evrak eksik | Yüksek |
| Alan evraklar arasında tutarsız; net ağırlık brütten büyük | Yüksek |
| Sevkiyat bilgisi eksik (teslim, taşıma, ödeme) | Yüksek |
| Ülke grubu varsayıldı; deniz kuralı (FOB, CIF...) ile deniz dışı taşıma; evrak sevkiyat bilgisinden farklı | Orta |
| Önerilen evrak eksik; listede olmayan evrak | Bilgi |

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek: Almanya, CIF, deniz, vesaik mukabili
python main.py --sevkiyat sevkiyat.csv                 # yalnız gereken evrak listesi
python main.py --sevkiyat sevkiyat.csv --evraklar evraklar.xlsx --kurallar ek_kurallar.csv
```

| Dosya | İçerik |
|---|---|
| Sevkiyat | İki sütun (Alan;Değer): Sevkiyat No, Alıcı, Ülke, Ülke Grubu, Teslim Şekli, Taşıma Şekli, Ödeme Şekli, Ürün Türü (Sanayi / İşlenmiş tarım / Tarım / Kömür çelik), Bitkisel Ürün, Hayvansal Ürün, Ahşap Ambalaj, Tehlikeli Madde, Dahilde İşleme, Gıda (Evet / Hayır) |
| Evraklar (isteğe bağlı) | Üç sütun: Evrak, Alan, Değer. Alanlar: Alıcı, Toplam Tutar, Döviz, Koli Sayısı, Net Ağırlık, Brüt Ağırlık, GTİP, Menşe, Teslim Şekli, Yükleme / Varış Limanı... |
| Kurallar (isteğe bağlı) | Evrak, Koşul, Zorunluluk, Düzenleyen, Açıklama |

## Çıktı

`ihracat_evrak_kontrol.xlsx`:
- `Kontrol Listesi`: evrak, zorunluluk, neden gerekli, düzenleyen, açıklama, durum (Hazır / Eksik); boş "Nüsha / Not"
  sütunu.
- `Tutarlılık`: alan × evrak matrisi; tutarsız alanlar üstte.
- `Uyarılar`.

## Dikkat

- **Kural tablosu genel bir başlangıç listesidir.** Gereken evraklar ürüne (GTİP), alıcı ülke mevzuatına ve alıcının
  taleplerine göre değişir. Listeyi gümrük müşavirinizle doğrulayın ve `--kurallar` ile kendi ürünlerinize uyarlayın.
- **Menşe ve dolaşım belgeleri:** A.TR menşe değil serbest dolaşım belgesidir. EUR.1 için ürünün ilgili anlaşmanın
  menşe kurallarını karşılaması gerekir. STA ülke listesi ve kurallar değişebilir; Ticaret Bakanlığı kaynaklarını
  kontrol edin.
- **Akreditifli işlemlerde** akreditif metni esastır. Belge uygunluğu için ayrıca "Akreditif Evrak Uygunluk Kontrolü"
  görevine bakın.
- **Örnek veri:** Örnek sevkiyat, firma ve evrak bilgileri kurgusaldır.

## Testler

Şunlar test edilir:
- AB + sanayi → A.TR; AB + tarım → EUR.1; STA → EUR.1; diğer → Menşe Şahadetnamesi.
- Taşıma türüne göre taşıma senedi; CIF / CIP'te sigorta; akreditif, vesaik mukabili, ahşap ambalaj, tehlikeli madde.
- Hazır / eksik evrak, tutarlılık (alıcı adı, koli sayısı; GTİP ve teslim şekli uyumu).
- Ek kural, kural kaldırma, varsayılan ülke grubu uyarısı, Excel ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Gümrük müşavirliği
hizmetinin yerine geçmez. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
