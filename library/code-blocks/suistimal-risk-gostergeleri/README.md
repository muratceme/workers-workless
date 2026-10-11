# Suistimal Risk Göstergeleri · Kod Bloğu

> Hasar Yönetimi › Hasar Uzmanı · Workers / Workless

Hasar dosyalarını dokuz basit göstergeyle tarar, tetiklenen göstergelerin puanlarını toplar ve **hangi dosyanın önce
ayrıntılı inceleneceğine** dair bir öncelik listesi çıkarır. İnternete bağlanmaz.

> **Gösterge kanıt değildir.** Yüksek puan suistimal olduğunu göstermez; dürüst bir sigortalının dosyası da birkaç
> göstergeyi tetikleyebilir. Bu çıktı tek başına ret, ödeme kesintisi veya ihbar gerekçesi olamaz.

## Nasıl hesaplar?

| Kod | Gösterge | Varsayılan eşik | Puan |
|---|---|---|---|
| `YAKIN_BASLANGIC` | Hasar, poliçe başlangıcından sonraki ilk N gün içinde | 30 gün | 20 |
| `YAKIN_BITIS` | Hasar, poliçe bitişinden önceki son N gün içinde | 15 gün | 10 |
| `GEC_IHBAR` | İhbar, hasardan N günden fazla sonra | 10 gün | 10 |
| `SIK_HASAR` | Aynı sigortalının (bu hasar dahil) son 365 günde N ve üzeri hasarı | 3 adet | 20 |
| `ZEYIL_SONRASI` | Hasar, teminat artırım zeylinden sonraki ilk N gün içinde | 30 gün | 20 |
| `BEDELE_YAKIN` | Talep tutarı ≥ sigorta bedelinin %N'i | %90 | 10 |
| `EKSPER_FARKI` | Talep, eksper tespitinin %N'inden fazla üzerinde | %30 | 15 |
| `BEYAN_TUTARSIZ` | Beyan edilen önceki hasar sayısı < dosyadaki kayıtlı önceki hasar sayısı | — | 15 |
| `TUTANAK_YOK` | Resmi tutanak "Yok" | — | 10 |

- **Puan** = tetiklenen göstergelerin puanları toplamı, en çok 100.
- **Düzey:** puan ≥ `--yuksek` (50) → Yüksek; ≥ `--orta` (25) → Orta; altı → Düşük.
- **Veri yoksa gösterge değerlendirilmez** ve dosyanın "Değerlendirilemeyen" sütununda listelenir; puana etkisi yoktur.
- **Sık hasar ve beyan tutarsızlığı** yalnız bu dosyadaki hasarlara bakar. Sigortalı, `Sigortalı No` varsa onunla,
  yoksa adıyla eşleştirilir.
- **Ayar dosyası** (`--gostergeler`): `Kod, Puan, Eşik`. Puanı 0 yaparsanız gösterge kapanır.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek veri, varsayılan puanlar
python main.py --hasarlar hasarlar.xlsx --gostergeler gostergeler.xlsx --yuksek 60 --orta 30
```

Hasar dosyası sütunları: Hasar No, Hasar Tarihi (zorunlu); Poliçe No, Sigortalı No, Sigortalı, Hasar Türü, Poliçe
Başlangıç, Poliçe Bitiş, Zeyil Tarihi, İhbar Tarihi, Sigorta Bedeli, Talep Tutarı, Eksper Tutarı, Beyan Edilen Önceki
Hasar, Resmi Tutanak (Var / Yok).

## Çıktı

`cikti/suistimal_gostergeleri.xlsx`:

- `İnceleme Önceliği`: puana göre sıralı dosyalar, tetiklenen ve değerlendirilemeyen göstergeler; boş "İnceleme Sonucu" sütunu.
- `Gösterge Ayrıntısı`: her tetiklenen gösterge için açıklama (kaç gün, yüzde kaç).
- `Gösterge Özeti`: kullanılan puan ve eşikler; göstergenin kaç dosyada tetiklendiği.
- `Uyarılar`: okunamayan alanlar, poliçe süresi dışındaki hasarlar.

## Dikkat

- **Karar aracı değildir.** Çıktı yalnız inceleme sırasını önerir. Hasarın reddi, eksik ödenmesi veya yetkili
  mercilere bildirim; somut delile, eksper ve araştırma raporuna ve şirketinizin usullerine dayanmalıdır.
- **Puanlar ve eşikler örnek varsayılanlardır.** İstatistiksel olarak doğrulanmış bir model veya sektör standardı
  değildir. Branşa göre anlamı değişir (ör. geç ihbar süresi); kendi verinizle sınayıp ayarlayın.
- **Yanlış alarm olağandır.** Yeni poliçede erken hasar veya bedele yakın talep tek başına olağan durumlardır.
- **Ayrımcılık yapmayın.** Göstergelere kişinin kimliği, mesleği, yaşadığı yer gibi özellikleri eklemeyin; yalnız
  dosyanın olgularına dayanın.
- **Kişisel veri:** Hasar dosyaları kişisel veri içerir. KVKK kapsamında yalnızca otomatik işleme dayanan ve kişi
  aleyhine sonuç doğuran karar verilmemeli; çıktı yetkili kişilerce değerlendirilmeli ve sınırlı erişimle saklanmalıdır.
- **Kapsam:** Kod yalnız verdiğiniz dosyaya bakar; sektör veri tabanlarını (ör. SBM kayıtları) sorgulamaz.
- **Örnek veri:** Sigortalılar, poliçeler ve hasarlar kurgusaldır.

## Testler

- Her göstergenin sınır değerleri (30. gün dahil, tam %30 tetiklemez vb.).
- Sık hasar penceresi (365 gün), beyan tutarsızlığı, sigortalı eşleşmesi.
- Puan tavanı, düzey eşikleri, kapatılan gösterge ve değiştirilen eşik.
- Okuma hataları ve örnek veri uçtan uca.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
