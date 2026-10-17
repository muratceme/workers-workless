# Resmî Yazı (Haciz/Müzekkere) Takibi · Kod Bloğu

> Bankacılık › Şube Bankacılığı › Operasyon Yetkilisi · Gerisi Bende

İcra dairesi, vergi dairesi, SGK, mahkeme ve savcılıktan gelen **haciz ihbarnamesi, haciz bildirisi, haciz
kaldırma ve bilgi talebi** yazılarını tür, cevap süresi ve müşteri bazında takip eder. Süresi geçen ve
yaklaşanları uyarır. İnternete bağlanmaz.

## Ne yapar?

- **Tür tanıma:** Yazı türü serbest metinden tanınır:
  - "89/1", "birinci haciz ihbarnamesi" → İİK 89/1.
  - Benzer şekilde 89/2, 89/3, 6183 haciz bildirisi, haciz kaldırma (fekki), tedbir, bilgi/belge talebi.
- **Son cevap günü:**

  | Durum | Son gün |
  |---|---|
  | Yazıda süre yazıyorsa | Tebliğ + yazıdaki gün |
  | İİK 89/1 haciz ihbarnamesi | Tebliğ + 7 gün |
  | İİK 89/2 haciz ihbarnamesi (itiraz) | Tebliğ + 7 gün |
  | İİK 89/3 haciz ihbarnamesi (ödeme / teslim ya da menfi tespit davası) | Tebliğ + 15 gün |
  | Haciz kaldırma | Tebliğ + `--fekk-gun` iş günü (banka içi hedef, varsayılan 1) |
  | Diğer türler (6183, bilgi talebi…) | Yazıdaki süre girilmeli; girilmemişse "Süre girilmeli" |

  - Süreler takvim günüdür ve tebliği izleyen günden başlar.
  - Son gün hafta sonuna veya resmî tatile denk gelirse izleyen ilk iş gününe uzar.
  - Yazıdaki süre yasal süreden farklıysa yazıdaki esas alınır ve not düşülür.
- **Durum:** Süresi geçti (kritik), bugün son gün, yaklaşan (`--uyari-gun` iş günü), süresi var, süre girilmeli,
  cevaplandı, geç cevaplandı.
- **Zincir kontrolü:**
  - Aynı dosyada 89/1 süresinde cevaplanmadan 89/2 gelmişse kritik uyarı: 89/2'ye 7 gün içinde itiraz
    edilmezse borç zimmette sayılır; hukuk birimine iletilmeli.
  - 89/3 her zaman kritik.
  - Haciz kaldırma yazısı gelen dosyalardaki açık haciz yazılarına not düşülür.
- **Veri kontrolleri:** TCKN/VKN kontrol hanesi, gelecekteki tebliğ tarihi, tebliğden önceki cevap tarihi,
  mükerrer kayıt.
- **Müşteri özeti:** Açık yazı sayısı ve toplam haciz tutarı (aynı dosyada 89/1 varsa 89/2 ve 89/3 tekrar sayılmaz), kurum
  sayısı ve en yakın son gün. Birden fazla kurumdan haciz gelen müşteriler işaretlenir; kredi izleme için de
  faydalıdır.

## Kullanım

```bash
pip install -r requirements.txt
python main.py                                                      # örnek: 13 yazı, rapor tarihi 08.10.2026
python main.py --girdi gelen_yazilar.xlsx --tarih 09.10.2026
python main.py --girdi gelen_yazilar.xlsx --tatiller tatiller.txt --uyari-gun 3 --fekk-gun 2
```

| Sütun | Açıklama |
|---|---|
| **Tebliğ Tarihi**, **Yazı Türü** | Zorunlu |
| Kayıt No, Kurum, Dosya No, Müşteri No, Müşteri Adı, TCKN/VKN, Tutar | İsteğe bağlı |
| Yazıdaki Süre (Gün) | Yazıda süre belirtilmişse |
| Cevap Tarihi | Cevap/işlem yapıldıysa |
| Sorumlu | İsteğe bağlı |

**Tatiller:** Sabit genel tatiller (2429 sayılı Ulusal Bayram ve Genel Tatiller Hakkında Kanun) ve 2025-2027 dini bayramları kodda tanımlıdır. Dini
bayram tarihlerini ve idari izin günlerini her yıl kontrol edin. 28 Ekim ve arife günlerinin yarım günü tatil
sayılmaz; eksik günleri `--tatiller` ile verin (satır
başına `GG.AA.YYYY`).

## Çıktı

`resmi_yazi_takibi.xlsx`:
- `Özet`.
- `Takip Listesi`: duruma ve son güne göre sıralı, "Yapılan İşlem" sütunuyla.
- `Müşteri Özeti`.
- `Süre Kuralları`: kullanılan süreler ve dayanakları.

## Dikkat

- **Hukuki görüş değildir:** Bu bir takip aracıdır. Cevabın içeriği (bakiye bildirimi, itiraz, bloke) ve
  sürelerin yorumu için bankanızın hukuk biriminin görüşü ve yazıdaki ifade esastır.
- **Kodda tanımlı süreler:** İİK 89/1 (7 gün), 89/2 (7 gün) ve 89/3 (15 gün). Diğer türlerde süreyi yazıdan girin.
- **Tebliğ tarihi:** Elektronik tebligat ve UYAP üzerinden gelen yazılarda tebliğ tarihinin nasıl belirlendiğini
  kontrol edin.
- **Kişisel veri ve bankacılık sırrı:** Liste TCKN/VKN gibi kişisel veriler ve bankacılık sırrı içerir; raporu
  yetkisiz kişilerle paylaşmayın.
- **Örnek veri:** Örnek kurumlar, dosyalar ve kişiler kurgusaldır.

## Testler

```bash
python -m unittest discover -s tests
```

## Kaynaklar ve dayanaklar

Kontrol tarihi: 15.10.2026. Mevzuat sık değişir; paketteki süre, oran ve kuralları kullanmadan önce güncel resmî metinle
karşılaştırın. "Karşılaştırılan maddeler" dışındaki atıflar yalnız yol göstericidir.

| Kaynak | Bağlantı | Durum |
|---|---|---|
| 2004 sayılı İcra ve İflas Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.3.2004.pdf) | Resmî metinle karşılaştırılan maddeler: 32, 33, 62, 67, 68, 78, 89, 146, 147, 149, 149/b, 150, 168 |
| 2429 sayılı Ulusal Bayram ve Genel Tatiller Hakkında Kanun | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.2429.pdf) | Resmî metinle karşılaştırılan maddeler: 1, 2 |
| 6183 sayılı Amme Alacaklarının Tahsil Usulü Hakkında Kanun | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.3.6183.pdf) | Resmî metin bağlantısı; paketteki madde atıfları ayrıca karşılaştırılmadı |

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

Ayrıntılar: `LICENSE` (MIT) ve `SORUMLULUK_REDDI.md`.
