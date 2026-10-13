# Acente Portföy ve Komisyon Raporu · Kod Bloğu

> Teknik ve Operasyon › Acente Yöneticisi · Workers / Workless

Acentenin üretimini **sigorta şirketi ve branş bazında** toplar; komisyon gelirini, ortalama komisyon oranını ve
tahsilat durumunu raporlar. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Komisyon | Dosyadaki komisyon tutarı; boşsa net prim × komisyon oranı ÷ 100 |
| Komisyon oranı % | Σ komisyon ÷ Σ net prim × 100 |
| Açık alacak | Brüt prim − tahsil edilen |
| Tahsilat oranı % | Σ tahsil edilen ÷ Σ brüt prim × 100 |
| Portföy payı % | Grubun brüt primi ÷ toplam brüt prim × 100 |

Net prim sütunu yoksa ya da boşsa brüt prim esas alınır. Eksi tutarlı satırlar (iptal, iade zeyli) toplamlardan
düşer ve "İptal / İade" sütununda sayılır; poliçe adedine girmez.

Uyarılar: bir şirketin portföy payı `--yogunluk` (50) yüzdesini aşıyorsa, şirket bazında tahsilat oranı
`--tahsilat` (70) yüzdesinin altındaysa, bir poliçede komisyon tutarı da oranı da yazılmamışsa.

**Örnek:** Net prim 12.000, komisyon oranı %10, komisyon tutarı boş → komisyon **1.200**. Bir şirkette brüt prim
157.500, tahsil edilen 105.000 → tahsilat oranı **%66,67**, açık alacak **52.500**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                        # örnek veri
python main.py --policeler policeler.xlsx
python main.py --policeler policeler.xlsx --yogunluk 40 --tahsilat 85
```

Zorunlu sütunlar: Sigorta Şirketi, Brüt Prim.

## Çıktı

`cikti/acente_portfoy_komisyon_raporu.xlsx`:

- `Şirketler`, `Branşlar`: poliçe adedi, prim, portföy payı, komisyon, komisyon oranı, açık alacak, tahsilat oranı.
- `Şirket × Branş`: aynı göstergeler ayrıntılı.
- `Aylar`: tanzim ayına göre poliçe, prim ve komisyon.
- `Poliçeler`, `Uyarılar`.

## Dikkat

- **Komisyon oranlarını kod varsaymaz.** Oranlar şirketle yaptığınız sözleşmeye, branşa ve ürüne göre değişir;
  tutarı ya da oranı dosyaya siz yazarsınız. Kod, şirketin hesapladığı komisyonla sizinkini karşılaştırmaz.
- **Komisyonun matrahı sizin verinizdir.** Komisyon genellikle vergiler hariç net prim üzerinden hesaplanır; net
  primi vermezseniz oran brüt prime göre çıkar ve düşük görünür.
- **Hak ediş ile tahsilat ayrıdır.** Rapor üretilen (tanzim edilen) poliçelerin komisyonunu gösterir; komisyonun
  şirketten fiilen ne zaman ve ne kadar ödendiğini, ek komisyon ve teşvikleri, iptallerde komisyon iadesini ayrıca
  izlemeniz gerekir.
- **Vergi hesabı yoktur.** Komisyon faturası, KDV ve BSMV durumu, stopaj ve benzeri kesintiler kapsam dışıdır.
- **"Tahsil edilen" müşteriden tahsilatı gösterir.** Şirkete aktarılan tutarlar ve cari hesap mutabakatı bu raporun
  konusu değildir.
- **Örnek veri:** Şirketler, poliçeler ve komisyon oranları kurgusaldır.

## Testler

- Komisyonun orandan hesaplanması, net primi olmayan poliçe, iptal satırlarının toplamlardan düşmesi.
- Komisyon oranı, tahsilat oranı, portföy payı, şirket × branş ve ay kırılımı.
- Yoğunlaşma ve tahsilat eşikleri, komisyonu yazılmamış poliçe uyarısı, okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
