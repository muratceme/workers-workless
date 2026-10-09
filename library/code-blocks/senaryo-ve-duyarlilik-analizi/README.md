# Senaryo ve Duyarlılık Analizi · Kod Bloğu

> Finans › Finansal Analist · Workers / Workless

Kur, faiz, satış ve maliyet varsayımlarından yıllık kârlılık ve nakit senaryoları üretir. Hangi varsayımın sonucu en
çok etkilediğini (tornado) ve başa baş satış hacmini / kuru hesaplar. İnternete bağlanmaz.

## Model

Sürücü tabanlı, tek yıllık bir modeldir:

| Kalem | Formül |
|---|---|
| Gelir | yurt içi hacim × yurt içi fiyat + ihracat hacmi × ihracat fiyatı (EUR) × EUR/TRY |
| Değişken maliyet | hacim × (ithal hammadde (USD) × USD/TRY + yerli değişken maliyet) |
| FAVÖK | gelir − değişken maliyet − personel − diğer sabit giderler |
| Faiz gideri | TL kredi × TL faiz + EUR kredi × EUR faiz × EUR/TRY |
| Kur farkı | net EUR pozisyonu × (EUR/TRY − dönem başı EUR/TRY) |
| Vergi öncesi kâr | FAVÖK − amortisman − faiz + kur farkı |
| Net kâr | vergi öncesi kâr − max(0, vergi öncesi kâr) × vergi oranı |
| Serbest nakit (basit) | FAVÖK − faiz − vergi − yatırım harcaması |

Ayrıca FAVÖK marjı, net kâr marjı ve faiz karşılama oranı (FAVÖK ÷ faiz) hesaplanır.

**Senaryolar:** Her satır bir parametreyi değiştirir:
- `%`: göreli değişim; örneğin EUR/TRY %15 artar.
- `puan`: oranlara eklenir; örneğin TL faiz +8 puan.
- `=`: değer atanır.

Belirtilmeyen parametreler temel değerde kalır.

**Tornado:** Her parametre tek tek değiştirilir (tutarlarda ±%10, oranlarda ±2 puan; `--adim`, `--oran-adim`). Net
kâra etkisi büyükten küçüğe sıralanır.

**Başa baş:** Diğer varsayımlar sabitken net kârı sıfırlayan satış hacmi ve EUR/TRY kuru.

| Kontrol | Önem |
|---|---|
| Temel senaryoda zarar | Yüksek |
| Faiz karşılama oranı 1,5'in altında | Orta |
| Negatif serbest nakit | Orta |
| Başa baş hacim mevcut hacmin %90'ının üstünde | Orta |
| Tanınmayan parametre, okunamayan senaryo satırı | Orta |
| Kısa döviz pozisyonu | Bilgi |

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                     # örnek: ihracatçı üretici, 4 senaryo
python main.py --varsayimlar v.csv --senaryolar s.csv --adim 10 --oran-adim 2
```

Varsayım dosyası `Parametre;Değer` listesidir; parametre adları `ornek_veri/varsayimlar.csv` ile aynı olmalıdır.
Oranlar `%40` veya `40` olarak yazılabilir. Kullanılmayan parametreyi 0 yazın veya silin; örneğin ihracat yoksa
ihracat payı 0.

## Çıktı

`senaryo_analizi.xlsx`:
- `Senaryolar`: her senaryo için tüm kalemler, temele göre farklar, net kâr / nakit grafiği, senaryo tanımları.
- `Tornado`: parametre bazında düşüş / artış etkisi ve grafik.
- `Başa Baş`: hacim ve kur, güvenlik payı.
- `Varsayımlar`: senaryolarda değişen değerler renkli.
- `Uyarılar`.

## Dikkat

- **Basit modeldir:** Tek yıllık ve doğrusaldır. Kapsam dışı kalanlar:
  - Stok ve alacak devir hızları, işletme sermayesi
  - Fiyatların kura tepkisi (kur geçişkenliği), kademeli faiz ve vergi istisnaları
- **Varsayımları birbiriyle tutarlı seçin:** Örneğin kur şoku senaryosunda faiz ve iç talep de değişebilir.
- **Kur farkı yalnız net EUR pozisyonu için hesaplanır.** USD borç / alacak varsa EUR karşılığı olarak ekleyin.
- **Örnek veri:** Örnek şirket ve varsayımlar kurgusaldır.

## Testler

Şunlar test edilir:
- Temel model kalemleri elle hesaplanmış değerlerle.
- Senaryo değişim türleri; kur şokunda kur farkı.
- Tornado sıralaması ve etkisi; başa baş hacimde net kâr = 0.
- Tanınmayan / eksik parametre uyarıları, Excel ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
