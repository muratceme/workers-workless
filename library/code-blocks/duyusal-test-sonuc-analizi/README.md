# Duyusal Test Sonuç Analizi · Kod Bloğu

> Ar-Ge ve Ürün Geliştirme › Ar-Ge Uzmanı · Workers / Workless

Duyusal panel puanlarını **her özellik için** numuneler arasında **varyans analiziyle (ANOVA)** karşılaştırır; anlamlı
fark varsa **Fisher LSD** ile hangi numunelerin ayrıştığını harf gruplarıyla gösterir; panelistlerin ölçek kullanımını
kontrol eder. İstatistik hesapları kodun içindedir; ek kütüphane ve internet gerekmez.

## Nasıl hesaplar?

| Adım | Yöntem |
|---|---|
| Veri | Her satır bir panelistin bir numuneye verdiği puanlar; Panelist ve Numune dışındaki sayısal sütunlar özelliktir |
| ANOVA | Her panelist her numuneyi bir kez puanladıysa iki yönlü ANOVA (numune + panelist, etkileşimsiz); değilse tek yönlü ANOVA |
| Anlamlılık | Numune etkisinin p değeri < `--alfa` (varsayılan 0,05) |
| Çoklu karşılaştırma | Fisher LSD: \|ortalama farkı\| > t(1 − α/2; hata sd) × √(KO_hata × (1/nᵢ + 1/nⱼ)) |
| Harf grupları | Dengeli veride aynı harfi paylaşan numuneler arasında anlamlı fark yoktur |
| Panelist notu | Panelistin özellik ortalaması panel ortalamasından `--panelist` (varsayılan 1,5) puandan fazla sapıyor |

İki yönlü modelde panelistler arası ölçek farkı (bazı panelistlerin genel olarak düşük ya da yüksek puan vermesi) hata
teriminden ayrılır; bu, numune farklarını daha hassas gösterir.

**Örnek:** 12 panelist, 3 numune, 9 puanlık hedonik ölçek. Doku özelliğinde p = 0,0027 → numuneler arasında fark var;
mevcut reçete 6,83 (a) ve az şekerli 6,42 (a) birbirinden farksız, rakip ürün 5,58 (b) ikisinden de düşük.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                 # örnek veri
python main.py --puanlar panel_puanlari.xlsx
python main.py --puanlar panel_puanlari.xlsx --alfa 0,01 --panelist 2
```

Zorunlu sütunlar: Panelist ve Numune; her özellik ayrı sütun (ör. Görünüş, Koku, Tat, Doku, Genel Beğeni). Oturum,
Tarih, Tekrar, Sıra ve Açıklama sütunları özellik sayılmaz.

## Çıktı

`cikti/duyusal_test_analizi.xlsx`:

- `Özet`: özellik × numune ortalamaları ve harf grupları, F, p, sonuç, LSD, kullanılan model.
- `Çift Karşılaştırma`: her numune çifti için fark, LSD ve sonuç.
- `Tanımlayıcı`: puan sayısı, ortalama, standart sapma.
- `Uyarılar`: dengesiz veri, analiz yapılamayan özellik, sapan panelist.

## Dikkat

- **Test tasarımı sonuçtan önemlidir.** Numunelerin rastgele üç haneli kodlarla, dengeli sırayla ve aynı koşullarda
  sunulması; panelistlerin eğitimi ve sayısı sonucu belirler. Kod tasarım hatasını düzeltemez.
- **Yöntem sınırları:** Fisher LSD çok sayıda numunede yanlış pozitif riskini artırır; çok numuneli çalışmalarda Tukey
  gibi daha tutucu yöntemler ya da istatistik yazılımıyla doğrulama önerilir. Ölçek aralıklı veri gibi ele alınır.
- **Tekrarlı ölçüm** (aynı panelistin aynı numuneyi birden çok kez puanlaması) dengesiz veri sayılır ve tek yönlü
  analiz yapılır; bu durumda oturum etkisini ayrı analiz edin.
- **Örnek veri:** Panelistler, numune kodları ve puanlar kurgusaldır.

## Testler

- F ve t dağılımının bilinen kritik değerleri (F₀,₀₅(2;10) = 4,10, t₀,₀₂₅(10) = 2,228).
- İki ve tek yönlü ANOVA, LSD, harf grupları, çift karşılaştırma, sabit puanlı özellik, sapan panelist, okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
