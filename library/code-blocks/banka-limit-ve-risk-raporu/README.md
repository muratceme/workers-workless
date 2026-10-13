# Banka Limit ve Risk Raporu · Kod Bloğu

> Finans › Finans Uzmanı · Workers / Workless

Bankalardaki kredi limitlerinizi, kullanımları, teminatları ve faiz oranlarını **tek tabloda** toplar: hangi
bankada ne kadar boş limit var, hangi limit dolmak üzere veya süresi bitiyor, risk hangi bankada yoğunlaşmış,
ortalama borçlanma faizi ne. İnternete bağlanmaz.

## Nasıl hesaplar?

| Kavram | Kural |
|---|---|
| Boş limit, kullanım oranı | Limit − kullanım; kullanım ÷ limit |
| TL karşılığı | Tutar × kur (`--kurlar`). Kuru verilmeyen para birimi TL toplamlarına girmez ve uyarılır |
| Toplam risk | Nakdi kullanım + gayrinakdi kullanım (TL) |
| Banka payı | Bankanın toplam riski ÷ tüm bankaların toplam riski |
| Ağırlıklı ortalama faiz | Σ(kullanım × faiz) ÷ Σ kullanım; yalnız nakdi krediler, her para birimi için ayrı |
| Teminat / risk | Teminat tutarı ÷ kullanım (ikisi de yazılmışsa) |

**Örnek:** 4.600.000 TL %52, 2.000.000 TL %48 ve 2.100.000 TL %50 faizli üç kredi →
(4,6 × 52 + 2 × 48 + 2,1 × 50) ÷ 8,7 = **%50,60**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                          # örnek veri, durum tarihi 12.10.2026
python main.py --limitler limitler.xlsx --kurlar kurlar.xlsx
python main.py --limitler limitler.xlsx --kurlar kurlar.xlsx --esik 85 --gun 45 --yogunlasma 50
```

Zorunlu sütunlar: Banka, Limit, Kullanım. Diğerleri: Ürün, Sınıf (Nakdi / Gayrinakdi; boşsa nakdi sayılır),
Para Birimi (boşsa TL), Faiz Oranı, Limit Bitiş, Teminat, Teminat Tutarı.

| Seçenek | Anlamı | Varsayılan |
|---|---|---|
| `--esik` | Kullanım oranı bu yüzdeye ulaşınca "limit dolmak üzere" uyarısı | 90 |
| `--gun` | Limit bitişine bu kadar gün kala uyarı | 30 |
| `--yogunlasma` | Toplam riskteki payı bu yüzdeye ulaşan banka için uyarı | 40 |

## Çıktı

`cikti/banka_limit_risk.xlsx`:

- `Limit ve Risk`: satır bazında limit, kullanım, boş limit, kullanım oranı, faiz, kalan gün, teminat, TL karşılıkları.
- `Banka Özeti`: banka bazında nakdi ve gayrinakdi limit ile risk, boş limit, pay; altında ağırlıklı ortalama faiz.
- `Uyarılar`: limit aşımı, dolmak üzere olan limit, süresi dolmuş veya dolacak limit, banka yoğunlaşması, eksik kur.

## Dikkat

- **Rapor sizin kayıtlarınıza dayanır.** Bankanın sistemindeki limit, risk ve teminat bilgisiyle mutabakat yerine
  geçmez; Risk Merkezi raporuyla ayrıca karşılaştırın.
- **Kurları siz verirsiniz.** Kod kur çekmez; hangi günün ve hangi tür kurun (alış, satış) kullanılacağına siz
  karar verirsiniz. Örnek dosyadaki kurlar kurgusaldır.
- **Faiz oranları olduğu gibi alınır.** Yıllık basit, bileşik, komisyon ve vergi dahil gibi farklar ayırt edilmez;
  ağırlıklı ortalamanın anlamlı olması için oranları aynı esasla yazın.
- **Eşikler yönetim tercihidir.** %90, 30 gün ve %40 genel başlangıç değerleridir; yasal bir sınır değildir.
- **Kapsam dışı:** çek, senet ve DBS limitlerinin ayrı izlenmesi, grup limitleri ve alt limit ilişkisi, kredi
  sözleşmesi şartları (kovenant) ve teminat değerlemesi.
- **Örnek veri:** Bankalar, tutarlar, oranlar ve kurlar kurgusaldır.

## Testler

- Boş limit, kullanım oranı, TL karşılığı, banka özeti, pay ve ağırlıklı ortalama faiz (elle hesaplanmış).
- Limit aşımı, eşik, süre, yoğunlaşma ve eksik kur uyarıları.
- Okuma hataları ve komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
