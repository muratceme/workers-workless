# Kumaş ve Aksesuar İhtiyaç Hesabı · Kod Bloğu

> Tekstil ve Konfeksiyon › Planlama › Planlama Uzmanı · Workers / Workless

Sipariş adetlerinden (model × renk × beden) **kumaş metrajını/kilosunu** ve **aksesuar adetlerini** hesaplar.
Stok düşülür, aksesuarlar ambalaj katına yuvarlanır, kumaşın kaç top tutacağı çıkarılır. İnternete bağlanmaz.

## Hesap

```
brüt ihtiyaç = sipariş adedi × (1 + fazla kesim %) × birim tüketim × (1 + fire %)
net ihtiyaç  = brüt ihtiyaç − stok
```

**Reçete satırının kapsamı**
- **Renge bağlı:** Kumaş, ribana, iplik ve fermuar gibi malzemeler ürün rengiyle ayrışır (Beyaz süprem, Lacivert
  süprem…).
- **Bedene bağlı:** Beden etiketi gibi malzemeler bedene göre ayrı listelenir.
- **Beden sütunu dolu:** Satır yalnız o beden için geçerlidir. Böylece **bedene göre farklı tüketim** girilebilir
  (S 0,50 kg, M 0,55 kg …).

**Malzeme kartı verilirse**
- **Metre ↔ kg:** Örme kumaş kg ile, dokuma metre ile planlanır. En (cm) ve gramaj (g/m²) varsa diğer birim de
  hesaplanır: `kg/metre = en/100 × gramaj/1000`.
- **Top/rulo:** Net ihtiyaç top boyuna bölünür ve yukarı yuvarlanır.
- **Yuvarlama:** Aksesuarlar tam sayıya ve ambalaj miktarının katına yuvarlanır (500'lük etiket, 5.000 m bobin,
  144'lük düğme). Asgari sipariş miktarının altında kalırsa asgariye çıkarılır.

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                         # örnek: 2 model, 3 renk, 3.400 adet
python main.py --siparis siparis.xlsx --recete recete.xlsx
python main.py --siparis siparis.xlsx --recete recete.xlsx --malzeme malzeme_karti.xlsx --stok stok.xlsx --fazla-kesim 3
```

| Dosya | Sütunlar |
|---|---|
| Sipariş | Model, Renk, Beden, Adet — veya Model, Renk, S, M, L, XL … (müşteri sipariş föyündeki gibi) |
| Reçete | Model, Malzeme Kodu, [Malzeme Adı], Tür (Kumaş/Aksesuar), Birim, Tüketim, [Fire %, Renge Bağlı, Bedene Bağlı, Beden, Malzeme Rengi] |
| Malzeme kartı | Malzeme Kodu, [Ambalaj Miktarı, Asgari Sipariş, Top Boyu, En (cm), Gramaj (g/m²), Tedarikçi, Tedarik Süresi] |
| Stok | Malzeme Kodu, [Renk, Beden], Stok |

**Malzeme Rengi** sütunu, ürün renginden farklı renkte malzeme kullanılan durumlar içindir. Örneğin beyaz
tişörtte lacivert ribana kullanılıyorsa bu sütuna "Lacivert" yazılır.

## Çıktı

`Özet` · `İhtiyaç Listesi` (kumaşlar üstte; brüt, stok, net, top, yuvarlanmış sipariş, tedarikçi) ·
`Hesap Detayı` (sipariş satırı × malzeme) · `Bilgi`

## İpuçları

- **Birim tüketim:** Kalıp ve pastal verimi değiştikçe güncel tutun. Pastal planı netleşmeden önceki tüketim
  tahminidir.
- **Fire oranı:** Kesim firesinin yanında boyahane/yıkama çekmesi ve kumaş hatası payını da kapsamalıdır.
- **Fazla kesim payı:** Numune, kalite kontrol ve ikinci kalite için müşteri ile anlaşılan orana göre verin.

## Testler

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
