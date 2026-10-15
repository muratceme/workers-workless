# Piyasa Günlüğü Özeti · AI Agent

> Hazine › Hazine Uzmanı · Workers / Workless

Gün içi piyasa verilerinden (kur, faiz, CDS, endeks, emtia) **kurum içi kısa piyasa günlüğü** hazırlar. Değişimleri
**kod hesaplar**; model özetler ve haberlerle bağlam kurar. Metindeki her rakam kodun hesapladığı değerlerle
karşılaştırılır.

## Nasıl çalışır?

1. **Hesap (kod):**
   - Kur, endeks ve emtiada tutar ve yüzde değişim.
   - Faizde (gösterge tahvil, politika faizi) ve CDS'te baz puan değişimi.
   - Günün en çok hareket eden fiyat göstergesi ve faiz / CDS göstergesi.
   - Tür sütunu yoksa gösterge adından tahmin edilir (USD/TRY → Kur, CDS → CDS, Tahvil → Faiz).
2. **Günlük (model):** Başlık, öne çıkanlar ve paragraflar yazılır. Her madde dayandığı gösterge (P…) ve haber (H…)
   numarasını gösterir.
3. **Denetim (kod):** Aşağıdakilerden biri varsa taslak modele geri gönderilir (varsayılan en fazla 2 tur):
   - Sayı, dayandığı göstergenin hazır rakamlarından ya da haber satırından değil
   - Faiz / CDS değişimi baz puan yerine yüzde ile yazılmış ("%43 geriledi")
   - "Yükseldi / geriledi" yönü değişimin işaretiyle çelişiyor
   - Neden-sonuç ("ardından", "nedeniyle") ya da beklenti ("yükselecek") ifadesi bir habere dayanmıyor
   - En çok hareket eden göstergeler ele alınmamış
   - Yatırım tavsiyesi sayılabilecek ifade var ("alım fırsatı", "öneririz", "hedef fiyat")
4. **Uyarı (kod):** Metnin sonuna "yatırım danışmanlığı kapsamında değildir" uyarısı eklenir.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py --tarih 14.10.2026                                         # örnek gün verisi
python agent.py --girdi piyasa.xlsx --haberler haberler.txt --tarih 15.10.2026
```

Veri tablosu örneği (Tür isteğe bağlı):

| Gösterge | Tür | Önceki Kapanış | Son |
|---|---|---|---|
| USD/TRY | Kur | 41,8520 | 41,9875 |
| TR 2Y Gösterge Tahvil | Faiz | 39,85 | 39,42 |
| TR 5Y CDS | CDS | 268 | 259 |

Verileri kurumunuzun lisanslı veri terminalinden ya da TCMB / Borsa İstanbul gibi resmî kaynaklardan alın. Agent veri
çekmez; yalnız verdiğiniz tabloyu kullanır.

## Çıktı

- `cikti/piyasa_gunlugu.xlsx`:
  - `Özet`: başlık, öne çıkanlar, paragraflar ve dayanakları.
  - `Veriler`: değişim, %, baz puan; en çok hareket edenler işaretli.
  - `Haberler`.
- `cikti/piyasa_gunlugu.txt`: dağıtıma hazır metin, veri tablosu ve uyarı.

## Dikkat

- **Kurum içi bilgilendirme içindir.** Müşterilere yatırım tavsiyesi niteliğinde dağıtım, SPK mevzuatı kapsamında yatırım
  danışmanlığı yetkisi gerektirir. Kod belirgin tavsiye ifadelerini yakalar; son kontrol sizdedir.
- **Neden-sonuç.** Model, haber notlarınızda olmayan bir nedeni ya da beklentiyi yazmaz; notlarınız ne kadar açıksa
  özet o kadar iyi olur.
- **Örnek veri:** Rakamlar ve haberler kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk taslakta bilerek:
- tahvil faizindeki 43 baz puanlık düşüşü "%43" yazar,
- BIST 100 değişimini "%1,6" diye yuvarlayıp "düştü" der,
- "USD/TRY yükselecek; dolar alım fırsatı sunuyor" diye tahmin ve tavsiye verir,
- günün en çok hareket eden göstergesi Brent'i atlar.

Kodun bunları yakalayıp düzelttirdiği doğrulanır.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Yatırım danışmanlığı değildir. API kullanım ücretleri ve metnin kullanımı kullanıcının
sorumluluğundadır. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
