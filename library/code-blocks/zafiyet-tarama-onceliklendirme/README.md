# Zafiyet Tarama Önceliklendirme · Kod Bloğu

> Bilgi Teknolojileri › Bilgi Güvenliği Uzmanı · Workers / Workless

Zafiyet tarayıcısı çıktısını yalnız CVSS'e göre değil, **aktif istismar bilgisi** ve **varlık bağlamıyla**
önceliklendirir. Taranabilen çıktılar Nessus, OpenVAS/Greenbone ve Qualys dökümleridir. Sonuçta yama ekibine
verilecek bir **düzeltme planı** çıkar. İnternete bağlanmaz.

## Neye bakar?

| Kaynak | Ne katar |
|---|---|
| CVSS (tarayıcıdan) | Teknik ciddiyet |
| **CISA KEV** kataloğu | Zafiyet gerçekte istismar ediliyor mu? Fidye yazılımında kullanılıyor mu? |
| **FIRST EPSS** skoru | Önümüzdeki 30 günde istismar edilme olasılığı |
| Varlık listesi | Kritiklik (1–5), internete açıklık, sorumlu ekip |

**Öncelik kuralları** (sırayla ilk uyan):

| Öncelik | Kural | Varsayılan süre |
|---|---|---|
| P1 | KEV + (internete açık veya kritik varlık veya fidye) · ya da EPSS ≥ %50, CVSS ≥ 7 ve internete açık/kritik | 7 gün |
| P2 | KEV · ya da CVSS ≥ 9 ve internete açık/kritik · ya da EPSS ≥ %10 ve CVSS ≥ 7 | 15 gün |
| P3 | CVSS ≥ 9 · ya da CVSS ≥ 7 ve internete açık/kritik | 30 gün |
| P4 | CVSS ≥ 7 · ya da CVSS ≥ 4 ve internete açık/kritik | 90 gün |
| P5 | diğerleri | 180 gün |

**Hedef tarih ve sıralama**
- **Hedef kapanış:** İlk görülme tarihi + düzeltme süresi. Süresi geçen bulgular gün sayısıyla işaretlenir.
  Süreleri `--sla sla.json` ile değiştirebilirsiniz.
- **Risk puanı:** Aynı öncelikteki bulguları sıralamak için kullanılır: CVSS × varlık ağırlığı × (1 + 4 × EPSS) ×
  (KEV ise 2) × (internete açıksa 1,5).

**Düzeltme paketleri:** Aynı bulgu ve aynı çözüm tek pakette toplanır. Böylece hangi yamanın kaç sunucuda ne kadar
riski kapattığı görülür ve yama ekibine sıralı bir iş listesi verilir.

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                            # kurgusal örnek (CVE-2099-…) verilerle
python main.py --tarama nessus.csv --varliklar varliklar.xlsx \
               --kev known_exploited_vulnerabilities.csv --epss epss_scores-current.csv.gz
python main.py --tarama openvas.csv --sla sla.json --bugun 08.10.2026
```

**KEV ve EPSS dosyaları**

Araç internete bağlanmaz; bu dosyaları kendiniz indirip verirsiniz.
- **KEV:** CISA "Known Exploited Vulnerabilities Catalog" sayfasından CSV veya JSON olarak indirilir.
- **EPSS:** FIRST EPSS günlük skor dosyası `.csv.gz` olarak indirilir. Dosyanın başındaki `#model_version`
  satırı otomatik atlanır.

Dosyalar verilmezse öncelik yalnız CVSS ve varlık bağlamıyla hesaplanır ve bu durum uyarı olarak raporlanır.

**Varlık listesi:** Host/IP, Hostname, Kritiklik (1–5 veya Kritik/Yüksek/Orta/Düşük), İnternete Açık (Evet/Hayır)
ve Sahibi sütunlarını içerir. Listede olmayan hostlar orta kritiklikte ve internete kapalı kabul edilir.

## Çıktı

`Özet` (öncelik dağılımı, KEV ve gecikme sayıları, kurallar) · `Bulgular` · `Düzeltme Planı` (sorumlu ve tarih
sütunları boş gelir) · `Varlıklar`

## Dikkat

- **Örnek veriler:** Örnekteki CVE kimlikleri, KEV ve EPSS kayıtları **kurgusaldır**. Gerçek analizde güncel
  KEV ve EPSS dosyalarını kullanın.
- **Kural ve süreler:** Öncelik kuralları ve süreler örnektir. Kurumunuzun zafiyet yönetimi politikasına göre
  ayarlayın. KEV'deki "dueDate" ABD kamu kurumları içindir ve burada yalnız bilgi olarak gösterilir.
- **Kapsam:** Tarayıcının bulamadığı zafiyetler, yapılandırma hataları ve tarama dışı varlıklar bu listede yer
  almaz.

## Testler

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

Ayrıntılar: `LICENSE` (MIT) ve `SORUMLULUK_REDDI.md`.
