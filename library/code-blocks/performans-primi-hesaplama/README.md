# Performans Primi Hesaplama · Kod Bloğu

> İnsan Kaynakları › Ücretlendirme ve Yan Haklar Uzmanı · Workers / Workless

Hedef gerçekleşmelerinden ve prim skalasından çalışan bazında brüt performans primini hesaplar. Dönem içinde
giren ve ayrılan çalışanları, ağırlık hatalarını ve bütçe aşımını raporlar. İnternete bağlanmaz.

## Nasıl hesaplar?

1. **Hedef gerçekleşmesi:**
   - Artan hedef (ciro, adet): gerçekleşen ÷ hedef.
   - Azalan hedef (fire, süre, şikâyet): hedef ÷ gerçekleşen.
   - Hedefi 0 olan azalan hedef (ör. iş kazası sayısı): gerçekleşen 0 ise %100, değilse %0.
   - Gerçekleşme doğrudan "Gerçekleşme (%)" sütunuyla da verilebilir.
   - Her hedef `--hedef-tavan` ile sınırlanır (varsayılan %150). Böylece tek bir hedefteki aşırı başarı diğerlerini
     telafi etmez.
2. **Ağırlıklı skor:** Σ ağırlık × gerçekleşme. Ağırlıklar toplamı 100 değilse 100'e orantılanır ve uyarı verilir.
3. **Prim çarpanı:** Skaladan okunur.
   - İlk satırın altı 0'dır (eşik).
   - Satırlar arası doğrusaldır; `--kademeli` ile basamaklı hale gelir.
   - Son satırın üstü son çarpandır (tavan).
4. **Şirket çarpanı (isteğe bağlı):** çarpan = bireysel ağırlık × bireysel çarpan + (1 − bireysel ağırlık) × şirket
   çarpanı.
5. **Prim:** hedef prim × çarpan × kıst.
   - Hedef prim = brüt ücret × maaş katı veya doğrudan TL.
   - Kıst = dönemde çalışılan gün ÷ dönem günü.
   - Dönem bitmeden ayrılana varsayılan olarak ödenmez (`--ayrilana-ode`). Dönemde en az çalışılması gereken gün
     sayısı `--min-gun` ile verilir.

Örnek skala (`ornek_veri/skala.csv`): %80 → %50, %100 → %100, %120 → %150. %94 skor → %85 çarpan.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                     # örnek: 8 çalışan, 2026 1. yarıyıl
python main.py --calisanlar c.xlsx --hedefler h.xlsx --skala skala.csv --donem 01.01.2026 31.12.2026
python main.py --calisanlar c.xlsx --hedefler h.xlsx --donem 01.01.2026 31.12.2026 --sirket-carpani 110 --bireysel-agirlik 70 --butce 900000
```

| Dosya | Sütunlar |
|---|---|
| Çalışanlar | Sicil No, Ad Soyad, Departman, Pozisyon, Brüt Ücret (TL), Hedef Prim (Maaş) veya Hedef Prim (TL), İşe Giriş, Ayrılış Tarihi |
| Hedefler | Sicil No, Hedef, Ağırlık (%), Hedef Değer, Gerçekleşen, Yön (Artan / Azalan) veya Gerçekleşme (%) |
| Skala | Gerçekleşme (%), Prim Çarpanı (%) |

## Çıktı

`performans_primi.xlsx`:
- `Prim Listesi`:
  - hedef prim, ağırlıklı skor, bireysel ve toplam çarpan
  - çalışılan gün, kıst oranı, brüt prim
  - işaretler ve boş "Onay" sütunu
- `Hedef Detayı`: hedef bazında gerçekleşme, tavanlı değer ve skora katkı.
- `Departman Özeti`:
  - ortalama skor ve çarpan, toplam prim
  - çarpan dağılımı
  - kullanılan skala ve ayarlar
- `Uyarılar`: hedefi olmayan, ağırlık toplamı hatalı, verisi eksik, dönem içinde ayrılan ve listede olmayan sicil;
  bütçe aşımı (orantı katsayısıyla); ortalaması yüksek departman (kalibrasyon).

## Dikkat

- **Politika sizindir:** Skala, tavanlar, ayrılanlara ödeme ve asgari çalışma süresi kurum prim yönetmeliğinize
  göre belirlenmelidir. Örnek değerler kurgusaldır. Prim yönetmeliği veya iş sözleşmesi primi hak olarak
  tanımlıyorsa, ayrılan çalışana ödememe kararı hukuki değerlendirme gerektirir.
- **Vergi ve SGK:** Hesaplanan tutar brüttür. Prim ücret niteliğindedir; SGK primi ve gelir vergisi bordroda
  hesaplanır. Kümülatif matrah nedeniyle vergi dilimi değişebilir. Uygulamayı mali müşavirinizle doğrulayın.
- **Veri kalitesi:** Hedef ve gerçekleşen değerler onaylı kaynaktan (satış, üretim, finans raporları) alınmalıdır.
- **Kişisel veri:** Prim ve performans verisi gizlidir. Raporu yetkili kişilerle sınırlı paylaşın.

## Testler

Şunlar test edilir:
- Artan, azalan ve hedefi 0 olan hedefler; skala (eşik, doğrusal, kademeli, tavan).
- Kıst, ayrılan çalışan ve asgari gün.
- Şirket çarpanı, hedef tavanı ve bütçe aşımı.
- Ağırlık orantılama, uyarılar, Excel sayfaları ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
