# KKB ve Risk Merkezi Raporu Özeti · Kod Bloğu

> Krediler Tahsis › Kredi Tahsis Uzman Yardımcısı · Gerisi Bende

Risk Merkezi / KKB raporundan tabloya aktardığınız limit-risk satırlarını toplar: **toplam risk, banka sayısı, limit
kullanımı, gecikme ve dönemsel seyir**. İnternete bağlanmaz; rapor sorgulamaz.

## Nasıl hesaplar?

| Gösterge | Hesap |
|---|---|
| Toplam limit / risk | Son dönemdeki satırların toplamı; nakdi ve gayrinakdi ayrı |
| Banka sayısı | Limiti veya riski olan farklı banka; ayrıca riski olan banka sayısı |
| Limit kullanım oranı | risk ÷ limit × 100 (toplam, banka ve kredi türü bazında) |
| Banka payı | bankanın riski ÷ toplam risk × 100 |
| Dönemsel değişim | (bu dönem risk − önceki dönem risk) ÷ önceki dönem risk × 100 |

**Nakdi / gayrinakdi:** "Nakdi/Gayrinakdi" sütunu doluysa o kullanılır. Boşsa kredi türünde `teminat mektubu`,
`akreditif`, `kabul kredisi`, `aval`, `kefalet`, `garanti` geçen satırlar gayrinakdi, diğerleri nakdi sayılır.

| Bulgu | Kural | Düzey |
|---|---|---|
| Gecikmede kredi var | Son dönemde gecikme günü veya geciken tutar > 0 | Yüksek |
| Limit aşımı | Satırda risk > limit (limit > 0) | Yüksek |
| Olumsuz kayıt (açık) | Durumu `kapandı`, `ödendi`, `tasfiye`, `kaldırıldı` içermeyen kayıt | Yüksek |
| Olumsuz kayıt (kapanmış) | Durumu yukarıdakilerden birini içeren kayıt | Orta |
| Önceki dönemde gecikme | Son dönem dışındaki bir dönemde gecikme | Orta |
| Limit kullanımı yüksek | Toplam kullanım ≥ `--kullanim` (varsayılan %90) | Orta |
| Toplam risk hızlı arttı | Son dönem değişimi ≥ `--artis` (varsayılan %25) | Orta |
| Risk tek bankada yoğunlaşmış | En büyük banka payı ≥ `--yogunlasma` (varsayılan %50), riski olan en az iki banka varsa | Bilgi |
| Riskli banka sayısı arttı | İlk döneme göre riski olan banka sayısı en az 2 arttı | Bilgi |
| Limitsiz risk | Limit boş veya 0, risk > 0 | Bilgi |

Eşikler yasal veya sektörel bir ölçüt değildir; bankanızın politikasına göre parametrelerle değiştirin.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri
python main.py --rapor risk_raporu.xlsx --olumsuz olumsuz_kayitlar.xlsx --kullanim 85
```

| Dosya | Sütunlar |
|---|---|
| Rapor satırları | Dönem (AA.YYYY veya YYYY-AA), Banka, Kredi Türü, Nakdi/Gayrinakdi, Limit, Risk, Gecikme Gün, Geciken Tutar |
| Olumsuz kayıtlar (isteğe bağlı) | Tür, Tarih, Tutar, Durum, Açıklama |

Zorunlu sütunlar `Banka` ve `Risk`. Dönem sütunu yoksa tüm satırlar tek dönem sayılır. Raporda banka adı
görünmüyorsa "Banka 1", "Banka 2" gibi yazın; aynı banka her dönemde aynı adla geçmelidir.

## Çıktı

`cikti/risk_raporu_ozeti.xlsx`:

- `Özet`: son dönemin toplamları ve bulgu sayıları.
- `Banka Bazında`: limit, risk, kullanım, pay, aşım ve gecikme.
- `Tür Bazında`: kredi türüne göre dağılım.
- `Dönemsel Seyir`: dönem dönem toplamlar ve değişim.
- `Bulgular`: düzeyiyle birlikte bulgular; boş "Değerlendirme" sütunu.
- `Uyarılar`: okunamayan satırlar.

## Dikkat

- **Raporun yerine geçmez.** Kod yalnız tabloya aktardığınız satırları toplar. Aktarım hatası sonuca aynen yansır;
  toplamları raporun kendi toplam satırıyla karşılaştırın.
- **Rapor biçimleri farklıdır.** Risk Merkezi ve KKB raporlarının düzeni rapor türüne göre değişir; bu paket belirli
  bir rapor şablonunu okumaz, sütun adları yukarıdaki gibi olan bir tablo bekler.
- **Kredi notu yorumlamaz, kredi kararı vermez.** Bulgular girilen rakamların mekanik karşılaştırmasıdır.
- **Kişisel ve ticari veri:** Risk raporları gizlidir. Dosyaları bankanızın bilgi güvenliği kurallarına göre saklayın.
- **Dönem karşılaştırması** aynı kapsamda hazırlanmış dönemler için anlamlıdır (aynı bankalar, aynı kredi türleri).
- **Örnek veri:** Bankalar, tutarlar ve olumsuz kayıtlar kurgusaldır.

## Testler

- Toplamlar, banka payı ve kullanım oranı; nakdi / gayrinakdi ayrımı.
- Gecikme, limit aşımı, limitsiz risk, sınır değerler ve eşik parametreleri.
- Dönemsel değişim, risk artışı ve banka sayısı bulguları; olumsuz kayıtlar.
- Okuma hataları ve örnek veri uçtan uca.

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
