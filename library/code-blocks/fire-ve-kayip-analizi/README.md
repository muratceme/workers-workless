# Fire ve Kayıp Analizi · Kod Bloğu

> Mağaza Operasyonları › Mağaza Müdürü · Workers / Workless

Fire kayıtlarını **ürün grubu ve neden bazında** analiz eder; sayım farklarını ekleyerek kayıtlı fire ile
**bilinmeyen kaybı** ayırır ve satışa oranla kayıp oranını hesaplar. İnternete bağlanmaz.

## Nasıl hesaplar?

| Kavram | Formül |
|---|---|
| Kayıtlı fire | Fire kayıtlarındaki tutarların toplamı (bozulma, son kullanma tarihi, hasar, imha …) |
| Bilinmeyen kayıp | −sayım farkı; sayım farkı eksi ise kayıp, artı ise sayım fazlası (kaybı azaltır) |
| Toplam kayıp | Kayıtlı fire + bilinmeyen kayıp |
| Fire % | Kayıtlı fire ÷ satış × 100 |
| Kayıp % | Toplam kayıp ÷ satış × 100 |
| Kayıt dışı payı % | Bilinmeyen kayıp ÷ toplam kayıp × 100 |
| Neden Pareto | Fire tutarına göre nedenlerin payı ve kümülatif payı; önceki kümülatif pay %80'in altındaysa A sınıfı |

Uyarılar: grubun kayıp oranı `--esik` (2) yüzdesini aşıyorsa (iki katını aşarsa yüksek); bilinmeyen kayıp toplam
kaybın yarısından fazlaysa; nedeni yazılmamış fire varsa. `--gruplar` verilmezse yalnız neden ve grup dökümü
yapılır, oranlar hesaplanmaz.

**Örnek:** Satış 350.000, kayıtlı fire 2.100, sayım farkı −9.800 → bilinmeyen kayıp 9.800, toplam kayıp 11.900,
kayıp oranı **%3,40**. Kaybın **%82,35**'i kayıt dışı.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                          # örnek veri
python main.py --fire fire_kayitlari.xlsx
python main.py --fire fire_kayitlari.xlsx --gruplar urun_gruplari.xlsx --esik 1,5
```

Zorunlu sütunlar: fire kayıtlarında Ürün Grubu ve Tutar; ürün gruplarında Ürün Grubu ve Satış.

## Çıktı

`cikti/fire_ve_kayip_analizi.xlsx`:

- `Ürün Grupları`: kayıp oranı en yüksekten başlayarak; satış, kayıtlı fire, sayım farkı, bilinmeyen ve toplam kayıp.
- `Neden Pareto`: fire nedenleri tutar, pay ve kümülatif payla.
- `Grup × Neden`: fire tutarı matrisi.
- `Fire Kayıtları`, `Uyarılar`.

## Dikkat

- **Üç tutar aynı esasta olmalı.** Fire ve sayım farkını maliyet bedeliyle, satışı satış bedeliyle verirseniz oran
  olduğundan düşük çıkar. Hepsini maliyetle ya da hepsini satış fiyatıyla verin; kod birini diğerine çevirmez.
- **Dönemler örtüşmeli.** Satış ve fire aynı dönemi kapsamalı; sayım farkı da iki sayım arasındaki o döneme ait
  olmalıdır. Yılda bir sayım yapıp aylık satışa bölerseniz oran anlamsızlaşır.
- **Sayım farkı kayıtlı fireyi içermemelidir.** Fire stoktan düşüldükten sonra kalan fark "bilinmeyen kayıp"tır.
  Fire sisteme işlenmeden sayım yapıldıysa aynı kayıp iki kez sayılır.
- **Sayım fazlası gerçek kazanç değildir.** Artı fark çoğunlukla sayım, teslim alma ya da ürün kodu hatasıdır;
  başka bir gruptaki eksikle birlikte incelenmelidir.
- **%2 eşiği politikanızdır.** Kabul edilebilir kayıp ürün grubuna göre çok değişir: taze ürünlerde yüksek, ambalajlı
  ürünlerde düşüktür. Grupları kendi hedefleriyle karşılaştırın.
- **Kod kaybın nedenini bilemez.** "Bilinmeyen kayıp" hırsızlık demek değildir; kayıtsız imha, teslim alma hatası,
  yanlış barkod ve kasa hataları da buradadır.
- **Örnek veri:** Ürün grupları ve tutarlar kurgusaldır.

## Testler

- Kayıtlı fire, bilinmeyen kayıp, toplam kayıp ve oranlar; sayım fazlası olan grup; satışı olmayan grup.
- Neden Pareto, kümülatif pay ve A sınıfı, grup × neden matrisi.
- Eşik ve iki kat eşiği, kayıt dışı payı uyarısı, grup dosyası olmadan çalışma, okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
