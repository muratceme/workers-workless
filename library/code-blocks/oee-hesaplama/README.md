# OEE Hesaplama · Kod Bloğu

> Üretim › Üretim Mühendisi · Workers / Workless

Vardiya ve makine bazında **Kullanılabilirlik × Performans × Kalite = OEE** hesaplar; kayıpları dakika
olarak ayrıştırır ve makine bazında özetler. İnternete bağlanmaz.

## Hesap

| | |
|---|---|
| Planlı üretim süresi | Vardiya süresi − planlı duruş (mola, planlı bakım) |
| Çalışma süresi | Planlı üretim süresi − plansız duruş (arıza, ayar, malzeme bekleme) |
| Kullanılabilirlik | Çalışma süresi ÷ planlı üretim süresi |
| Performans | İdeal çevrim süresi × toplam adet ÷ çalışma süresi |
| Kalite | Sağlam adet ÷ toplam adet |

Makine ve genel değerler **süre ve adet toplamlarından** hesaplanır; vardiya oranlarının basit ortalaması
alınmaz (uzun ve kısa vardiyaları yanlış ağırlıklandırır). Test, OEE literatüründeki klasik örnekle
(K %88,81 · P %86,11 · Q %97,80 · OEE %74,79) birebir doğrulanır.

Kayıplar: kullanılabilirlik kaybı (duruş dakikası), performans kaybı (yavaş çalışma ve küçük duruşlar),
kalite kaybı (hatalı ürünlere harcanan ideal süre), değer katan süre.

**Kontroller:** performans %100'ü aşarsa (ideal çevrim süresi gerçekçi değil veya veri hatalı), duruşlar
vardiya süresini aşarsa, hatalı adet toplamdan büyükse uyarı verilir.

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                   # örnek vardiya kayıtlarıyla dener
python main.py --girdi vardiya_kayitlari.xlsx
```

Sütunlar: **Makine, Vardiya Süresi (dk), Plansız Duruş (dk), İdeal Çevrim Süresi (sn), Toplam Üretim, Hatalı**
(zorunlu); Tarih, Vardiya, Ürün, Planlı Duruş (dk) (isteğe bağlı).

## Testler

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
