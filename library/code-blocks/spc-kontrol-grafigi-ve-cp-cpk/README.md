# SPC Kontrol Grafiği ve Cp/Cpk · Kod Bloğu

> Üretim › Kalite Mühendisi · Workers / Workless

Alt grup ölçümlerinden **X̄-R kontrol grafiğini** çizer, kontrol dışı noktaları **Nelson kurallarıyla**
işaretler ve şartname sınırları verilirse **Cp, Cpk, Pp, Ppk** proses yeterlilik indekslerini hesaplar.
Grafikler Excel'in kendi grafik nesneleridir. İnternete bağlanmaz.

## Hesaplar

| | Formül |
|---|---|
| X̄ grafiği | CL = X̄̄ · UCL/LCL = X̄̄ ± A2·R̄ |
| R grafiği | CL = R̄ · UCL = D4·R̄ · LCL = D3·R̄ |
| σ (alt grup içi) | R̄ / d2 → Cp, Cpk |
| σ (toplam) | tüm ölçümlerin örneklem standart sapması → Pp, Ppk |
| Cp / Pp | (USL − LSL) / 6σ |
| Cpk / Ppk | min(USL − X̄̄, X̄̄ − LSL) / 3σ (tek taraflı şartnamede yalnızca o taraf) |

A2, D3, D4, d2 sabitleri alt grup büyüklüğü 2–10 için standart SPC tablolarından alınmıştır.

**Nelson kuralları (X̄ grafiği):** K1 3σ dışında nokta · K2 merkezin aynı tarafında 9 nokta ·
K3 6 nokta sürekli artış/azalış · K5 3 noktanın 2'si 2σ dışında · K6 5 noktanın 4'ü 1σ dışında.
R grafiğinde sınır dışı aralıklar ayrıca işaretlenir.

**Cpk yorumu:** ≥ 1,67 çok iyi · ≥ 1,33 yeterli · 1,00–1,33 sınırda · < 1,00 yetersiz.

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                                    # örnek: Ø25 ±0,05 mil çapı
python main.py --girdi olcumler.xlsx --alt-sinir 24.95 --ust-sinir 25.05
python main.py --girdi olcumler.csv --ozellik "Çap" --ust-sinir 10.2
```

**Girdi biçimleri:** geniş (her satır bir alt grup, `Ölçüm 1 … Ölçüm n` veya `X1 … Xn` sütunları; isteğe bağlı
`Alt Grup` adı) ya da uzun (`Alt Grup` + `Ölçüm`/`Değer`). Birden çok özellik tek dosyadaysa `Özellik` sütunu ekleyin.

## Uyarılar

- Güvenilir sınırlar için en az 20–25 alt grup önerilir; daha azında uyarı verilir.
- Ölçüm sayısı diğerlerinden farklı alt gruplar hesaba katılmaz ve raporda belirtilir.
- Proses istatistiksel kontrolde değilse Cp/Cpk yorumlanmamalıdır; rapor bunu da belirtir.

## Testler

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
