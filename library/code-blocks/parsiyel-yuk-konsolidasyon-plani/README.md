# Parsiyel Yük Konsolidasyon Planı · Kod Bloğu

> Lojistik ve Taşımacılık › Parsiyel Operasyon › Parsiyel Operasyon Uzmanı · Workers / Workless

Bekleyen parsiyel yükleri varış bölgesine göre gruplar ve araçlara yerleştirir. Her aracın ağırlık, hacim ve
yükleme metresi (LDM) doluluğunu hesaplar. Sığmayan, bekleyen ve düşük doluluklu araçları listeler. İnternete
bağlanmaz.

## Nasıl hesaplar?

| Adım | Kural |
|---|---|
| LDM | ⌈palet adedi ÷ istif katı⌉ × palet eni × palet boyu ÷ 2,4. Dorse iç genişliği 2,4 m'dir; istiflenmeyen yükte istif katı 1 girilir. |
| Ödenebilir ağırlık | brüt kg, m³ × 333 ve LDM × 1.750 içinden en büyüğü (`--hacim-kg`, `--ldm-kg`) |
| Bölge sırası | En erken son yükleme tarihi olan bölge önce. |
| Bölge içi sıra | ADR'li yükler önce, sonra son yükleme tarihi, sonra büyük yük. |
| Yerleştirme | Yük, bölgede açılmış araçlardan sığdığı ilkine konur. Hiçbirine sığmazsa kalan en büyük uygun araç açılır. ADR donanımlı araçlar ADR'siz yükler için en sona bırakılır. |
| Küçültme | Bölge bitince her araç, yükünü taşıyabilen en küçük müsait araç tipine indirilir. |
| ADR | ADR'li yük yalnız ADR donanımlı araca konur. |

**Uyarılar:**

| Uyarı | Önem |
|---|---|
| Hiçbir araç tipine sığmayan yük (komple / özel araç gerekir) | Yüksek |
| Araç yetmediği için bekleyen yük | Yüksek |
| Son yükleme tarihi plan tarihinden önce | Yüksek |
| Doluluğu %60'ın altında araç (`--min-doluluk`). Yüklerin hepsi en az 2 gün bekleyebiliyorsa "bekletip birleştir" önerilir. | Orta |
| Plan tarihinde hazır olmayan yük | Bilgi |

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek: 25 yük, 4 bölge, 6 araç, plan tarihi 09.10.2026
python main.py --yukler yukler.xlsx --araclar araclar.csv --tarih 09.10.2026
python main.py --yukler y.xlsx --araclar a.csv --min-doluluk 70 --hacim-kg 300 --ldm-kg 1850
```

| Dosya | Sütunlar |
|---|---|
| Yükler | Yük No, Müşteri, Varış Bölgesi, Varış Şehri, Palet Adedi, Palet En (m), Palet Boy (m), İstif Katı, Brüt Ağırlık (kg), Hacim (m³), Hazır Tarihi, Son Yükleme Tarihi, ADR (Evet / Hayır), Navlun (EUR) |
| Araç tipleri | Araç Tipi, Adet, Yük Kapasitesi (kg), Hacim (m³), LDM, ADR |

Palet en / boy boşsa Euro palet (0,8 × 1,2 m) kabul edilir.

## Çıktı

`konsolidasyon_plani.xlsx`:
- `Araç Planı`: araç bazında şunları ve boş "Plaka / Not" sütununu gösterir:
  - bölge ve şehirler
  - yük ve palet sayısı
  - ağırlık / hacim / LDM ve ayrı ayrı doluluk oranları
  - belirleyici kısıt
  - ödenebilir ağırlık ve navlun toplamı
- `Yükleme Listesi`: araç → yük. Şehir sırasıyla dizilir, boşaltma sırasını planlamaya yardımcı olur.
- `Bekleyen Yükler`: hazır değil, sığmıyor veya araç yok; "Karar" sütunuyla.
- `Bölge Özeti`: bölge bazında toplamlar, araç sayısı ve ortalama doluluk; araç tipi tablosu.
- `Uyarılar`.

## Dikkat

- **Plan bir ön plandır:** Yerleştirme basit bir yaklaşımla yapılır; en iyi çözümü garanti etmez. Bazı durumlar
  hesaba katılmaz:
  - paletlerin araç içindeki geometrik yerleşimi
  - aks yükü dağılımı
  - yük emniyeti
  - ADR ayrım (segregasyon) kuralları
  - gıda ile kimyasalın aynı araçta taşınması

  Yükleme öncesinde operasyon tarafından kontrol edilmelidir.
- **Örnek değerler:** Araç kapasiteleri ile 333 kg/m³ ve 1.750 kg/LDM oranları yaygın örneklerdir. Kendi
  araçlarınızın ve tarifenizin değerlerini girin.
- **Örnek veri:** Örnek yükler ve müşteriler kurgusaldır.

## Testler

Şunlar test edilir:
- LDM ve ödenebilir ağırlık.
- Bölge ve ADR yerleştirmesi; araç küçültme.
- Sığmayan, hazır olmayan ve araç bekleyen yükler; düşük doluluk.
- Excel sayfaları ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
