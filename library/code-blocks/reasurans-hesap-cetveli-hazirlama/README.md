# Reasürans Hesap Cetveli Hazırlama · Kod Bloğu

> Sigorta › Reasürans › Reasürans Uzmanı · Workers / Workless

Orantılı bir trete (kotpar veya eksedan) için **dönem hesap cetvelini** prim ve hasar bordrolarından hazırlar.
Devredilen prim, komisyon, hasar, prim depo ve depo faizini hesaplar; bakiyeyi reasürör paylarına dağıtır.
İnternete bağlanmaz.

## Ne hesaplar?

| Kalem | Hesap |
|---|---|
| Devir oranı — kotpar | Tretedeki sabit oran |
| Devir oranı — eksedan | (sigorta bedeli − saklama payı) / sigorta bedeli. Devredilen bedel en fazla satır sayısı × saklama payıdır. Saklama payının altındaki poliçeler devredilmez. |
| İptal / zeyil / iade | İlk poliçe kaydının oranıyla |
| Devredilen prim | Brüt prim × devir oranı |
| Reasürans komisyonu | Devredilen prim × komisyon oranı |
| Devredilen ödenen hasar | Ödeme tarihi dönem içindeki hasarlar × poliçenin devir oranı |
| Devredilen muallak | Bilgi amaçlı; bakiyeye girmez |
| Prim depo tutulan | Devredilen prim × depo oranı |
| Prim depo iadesi ve faizi | Önceki dönem deposu; faiz = depo × yıllık faiz × gün / 365 (tutulma tarihinden dönem sonuna) |
| Bakiye | (prim + depo iadesi + faiz) − (komisyon + hasar + tutulan depo). Pozitifse şirket öder, negatifse reasürör öder. |
| Reasürör payları | Bakiye × pay; kuruş farkı son reasüröre yazılır |
| Kâr komisyonu ön hesabı | prim − komisyon − ödenen − muallak − yönetim gideri; pozitifse × oran. **Bilgi amaçlıdır.** |

## Kontroller

| Kontrol | Önem |
|---|---|
| Eksedan kapasite aşımı: sigorta bedeli > saklama + trete kapasitesi (fakültatif gerekir) | Yüksek |
| Hasarın poliçesi prim bordrosunda yok | Yüksek |
| Hasar tarihi poliçe süresi dışında (hesaba alınmaz) | Yüksek |
| Reasürör payları toplamı %100 değil | Yüksek |
| Devredilen hasar nakit hasar (cash call) limitini aşıyor | Orta |
| İptal / zeyil hareketinin ilk poliçe kaydı yok | Orta |
| Sıfır / negatif primli yeni poliçe | Orta |
| Dönem dışı prim hareketi veya hasar ödemesi (bu cetvele alınmaz) | Bilgi |

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                       # örnek eksedan trete, 2026 3. çeyrek
python main.py --trete trete.json --primler prim_bordrosu.xlsx --hasarlar hasar_bordrosu.xlsx
```

`trete.json` alanları (örnek: `ornek_veri/trete.json`):

| Alan | Açıklama |
|---|---|
| `tur` | `kotpar` veya `eksedan` |
| `devir_orani` | Kotpar için, ör. 0.40 |
| `saklama_payi`, `satir_sayisi` | Eksedan için |
| `komisyon_orani`, `prim_depo_orani`, `depo_faiz_orani` | Oranlar, ör. 0.30 |
| `nakit_hasar_limiti` | İsteğe bağlı |
| `kar_komisyonu_orani`, `yonetim_gideri_orani` | İsteğe bağlı |
| `donem_baslangic`, `donem_bitis` | GG.AA.YYYY |
| `onceki_depo` | `{"tutar", "tutulma_tarihi"}` — bu dönemde iade edilecek depo |
| `reasurorler` | `[{"ad", "pay"}]` |

**Bordrolar:**
- Prim hareketinin dönemi "Başlangıç" tarihine göre belirlenir. İptal ve zeyillerde bu sütuna hareket tarihini yazın.
- "Hareket" sütununda "İptal", "Zeyil", "İade" veya "Ek prim" yazan satırlar ilk poliçe kaydının oranını alır.

## Çıktı

`reasurans_hesap_cetveli.xlsx`:
- `Hesap Cetveli`: alacak / borç kalemleri; %100 ve reasürör sütunları; bakiye; muallak; kâr komisyonu ön hesabı.
- `Prim Bordrosu` ve `Hasar Bordrosu`: devir oranı, devredilen tutarlar, notlar.
- `Kontroller`: "İnceleme" sütunuyla.
- `Parametreler`.

## Dikkat

- **Trete şartları esas:** Hesaplar basitleştirilmiş, yaygın bir orantılı trete yapısına göredir. Prim depo
  (portföy) uygulaması, faiz esası, komisyon kademeleri (sliding scale), kâr komisyonu, zarar devri, kur ve
  portföy giriş/çıkışı tretenizin şartlarına göre değişir; cetveli imzalamadan önce sözleşmeyle karşılaştırın.
- **Hasar dönemi:** Hasarlar ödeme tarihine göre döneme alınır. Muallak yalnız bilgi olarak gösterilir.
- **Örnek veri:** Örnek trete, poliçeler, hasarlar ve reasürörler kurgusaldır.

## Testler

Şunlar test edilir:
- Örnek cetvelin tüm kalemleri ve bakiye (−1.749.185,00 TL).
- Eksedan oranları, kapasite sınırı, iptalin oranı; tüm kontroller.
- Elle hesaplanmış bir kotpar örneği; hatalı trete dosyası.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
