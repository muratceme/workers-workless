# Harcama Analizi · Kod Bloğu

> Satın Alma › Satın Alma Müdürü · Workers / Workless

Satın alma verisini kategori ve tedarikçi bazında **ABC sınıflandırmasıyla** analiz eder. Fiyat farklarını, tek kaynak
bağımlılığını, dağınık harcamayı ve olağandışı fiyat artışlarını bulur; bunlardan **pazarlık öncelikleri** listesi
çıkarır. İnternete bağlanmaz.

## Nasıl hesaplar?

- **TL karşılığı:** tutar × kur. Döviz satırında kur yoksa satır analize alınmaz. Tutar yoksa miktar × birim fiyat
  kullanılır.
- **Tedarikçi adları:** Büyük-küçük harf, Türkçe karakter ve şirket türü ekleri (A.Ş., Anonim Şirketi, Ltd. Şti.,
  San. Tic.) yok sayılarak birleştirilir. Rapordaki ad en sık kullanılan yazılıştır.
- **ABC:** Harcamaya göre azalan sırada, kendinden önceki kümülatif pay %80'in altındaysa A, %95'in altındaysa B, değilse C.
  Örnek: 50 / 30 / 15 / 5 → A, A, B, C.
- **Fiyat farkı:** Aynı malzeme ve birim **aynı ay içinde** farklı fiyatlarla alındıysa:
  - Ortalamanın üstü fazla = Σ (fiyat − ağırlıklı ortalama) × miktar, yalnız ortalamanın üstündeki alımlar.
    Bu tutar gerçekçi potansiyel olarak raporlanır.
  - En düşüğe göre fark = Σ (fiyat − en düşük fiyat) × miktar. Bu tutar üst sınırdır.
  - Karşılaştırma ay içinde yapılır, çünkü enflasyon dönemler arası fiyat farkını büyütür. `--fiyat-donem ceyrek` ile
    çeyrek içinde de yapılabilir.
- **Fiyat artışı:** Malzemenin ilk ve son ay ağırlıklı ortalama fiyatı karşılaştırılır. Artış, tüm malzemelerin
  medyan artışını `--artis` (10) puan aşarsa işaretlenir. Mutlak eşik kullanılmaz, çünkü genel enflasyon tüm
  malzemeleri eşiğin üstüne taşır.

### Pazarlık öncelikleri

| Tür | Koşul | Öncelik |
|---|---|---|
| Fiyat birliği / toplu pazarlık | Kategoride ortalamanın üstü fazla > 0 ve en az 2 tedarikçi | A → 1, B → 2, C → 3 |
| Tek kaynak bağımlılığı | En büyük tedarikçinin payı ≥ %80 (`--tek-kaynak`), kategori A veya B | A → 1, B → 2 |
| Tedarikçi konsolidasyonu | Kategoride ≥ 5 tedarikçi (`--dagitik`) ve en az 3'ü genel sıralamada C sınıfı | A → 2, diğer → 3 |
| Fiyat artışı incelemesi | Medyan artışın 10 puan üstü, kategori A veya B | 2 |

Aynı öncelikteki satırlar potansiyel tutara, sonra kategori harcamasına göre sıralanır.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                       # örnek: Ocak–Eylül 2026, 236 satır
python main.py --veri satinalma.xlsx
python main.py --veri satinalma.xlsx --a-esik 70 --b-esik 90 --artis 15 --fiyat-donem ceyrek
```

| Sütun | Açıklama |
|---|---|
| Tarih, Tedarikçi | Zorunlu |
| Kategori | Boşsa "Sınıflandırılmamış" |
| Malzeme, Miktar, Birim, Birim Fiyat | Fiyat analizleri için |
| Tutar | KDV hariç; yoksa miktar × fiyat |
| Döviz, Kur | TRY dışı satırlarda kur gerekir |
| Departman | Bilgi amaçlı |

## Çıktı

`harcama_analizi.xlsx`:
- `Özet`.
- `Pazarlık Öncelikleri`: gerekçe ve potansiyel tutar; boş "Sorumlu", "Hedef Tarih" ve "Sonuç" sütunları.
- `Kategori ABC` (grafikli), `Tedarikçi ABC`: sınıf özeti tablosuyla.
- `Kategori × Tedarikçi`: tedarikçi sayısı, en büyük pay, C sınıfı tedarikçiler ve dağılım.
- `Fiyat Farkları`, `Fiyat Artışları`.
- `Aylık Trend`: kategori × ay.
- `Veri`: TL karşılıklı tüm satırlar.
- `Uyarılar`.

## Dikkat

- **Fiyat farkı potansiyeli kesin tasarruf değildir.** Farklı fiyatlar kalite, teslim süresi, ödeme vadesi veya sipariş
  büyüklüğünden kaynaklanabilir.
- **Malzeme eşleşmesi** ada ve birime göre yapılır. Aynı malzeme farklı adlarla yazılmışsa malzeme kodunu "Malzeme"
  sütununa yazın.
- **Sınırlar** (ABC %80 / %95, tek kaynak %80, dağınıklık 5 tedarikçi) yaygın uygulamadır ama şirkete göre
  ayarlanmalıdır.
- **KDV:** Tutarlar KDV hariç olmalıdır; KDV dahil ve hariç tutarları karıştırmayın.
- **Örnek veri:** Örnek satın almalar, fiyatlar ve tedarikçi adları kurgusaldır.

## Testler

Şunlar test edilir:
- Toplamın CSV'den bağımsız hesapla doğrulanması ve kursuz döviz satırı.
- ABC sınırları, tedarikçi adı birleştirme, kategorisiz satır.
- Fiyat farkı tutarları (elle hesaplanmış örnek) ve farklı aydaki alımların karşılaştırılmaması.
- Medyana göre fiyat artışı, pazarlık öncelikleri, Excel ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
