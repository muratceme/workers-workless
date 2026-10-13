# Analiz Sonucu Spesifikasyon Kontrolü · Kod Bloğu

> Kalite Kontrol ve Laboratuvar › Kalite Kontrol Teknisyeni · Workers / Workless

Laboratuvar analiz sonuçlarını **sizin spesifikasyon limitlerinizle** karşılaştırır; uygunsuz partileri, sınıra yakın
değerleri ve eksik analizleri işaretler. İnternete bağlanmaz.

## Nasıl değerlendirir?

| Durum | Kural |
|---|---|
| Uygunsuz | Sonuç alt limitin altında veya üst limitin üstünde; nitel parametrede sonuç beklenenden farklı |
| Sınıra yakın | İki limit varsa limite uzaklık ≤ (üst − alt) × `--yakin` %; tek limit varsa sonuç limitin `--yakin` %'si kadar yakınında |
| Değerlendirilemedi | Sonuç sayıya çevrilemiyor veya "<101" gibi bir sonuç limitle karşılaştırmaya yetmiyor |
| Eksik analiz | Ürünün spesifikasyonunda olan parametrenin partide sonucu yok |
| Uygun | Diğer durumlar; limite eşit sonuç uygundur |

- `<10` gibi sonuçlarda gerçek değer 10'un altında kabul edilir: üst limit 100 ise uygundur. `>150` üst limit 100 ise
  uygunsuzdur.
- Parti kararı en ağır durumdur: Uygunsuz > Eksik analiz > Değerlendirilemedi > Sınıra yakın > Uygun.
- Spesifikasyonda **Ürün** sütunu boş bırakılan satır bütün ürünlere uygulanır; ürüne özel satır varsa o önceliklidir.

**Örnek:** pH limiti 4,3 – 4,8, yakınlık payı %10 → pay 0,05. Sonuç 4,78 üst limite 0,02 uzaklıkta → "Sınıra yakın".
Kuru madde alt limiti 40; sonuç 38,6 → "Uygunsuz, alt limitin 1,4 altında".

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek veri
python main.py --sonuclar analiz_sonuclari.xlsx --spesifikasyon spesifikasyon.xlsx
python main.py --sonuclar analiz_sonuclari.xlsx --spesifikasyon spesifikasyon.xlsx --yakin 5
```

- Sonuçlar: Parti No, Ürün, Parametre, Sonuç, Analiz Tarihi (her sonuç bir satır).
- Spesifikasyon: Ürün, Parametre, Alt Limit, Üst Limit, Beklenen (ör. Negatif), Birim.

## Çıktı

`cikti/spesifikasyon_kontrol.xlsx`:

- `Parti Özeti`: parti bazında karar önerisi, uygunsuz ve sınıra yakın sayısı, eksik parametreler.
- `Sonuçlar`: her sonuç, limiti, durumu ve limitten sapması.
- `Parametre Özeti`: parametre bazında uygunsuz sayısı, en düşük, en yüksek, ortalama.
- `Uyarılar`: uygunsuz, eksik, değerlendirilemeyen ve sınıra yakın sonuçlar; okunamayan satırlar.

## Dikkat

- **Limitler sizindir.** Kod mevzuat, kodeks veya müşteri limiti bilmez; spesifikasyon dosyanıza ne yazdıysanız onu
  uygular. Limitlerin güncel mevzuata ve müşteri şartnamesine uygunluğunu siz doğrularsınız.
- **Ölçüm belirsizliği dikkate alınmaz.** Limite çok yakın sonuçlarda uygunluk kararı, laboratuvarınızın karar
  kuralına göre verilir.
- **Mikrobiyolojik numune planları kapsam dışıdır.** n, c, m, M değerleriyle çok numuneli değerlendirme yapılmaz; her
  sonuç tek başına limitle karşılaştırılır.
- **"Karar önerisi" karar değildir.** Partiyi serbest bırakma, bloke etme veya yeniden analiz kararı yetkili kişinindir.
- **Tek limitte yakınlık payı limit değerine göre hesaplanır.** Alt limiti 40 olan parametrede %10 pay 44'e kadar
  "yakın" demektir; parametrelerinize göre `--yakin` değerini ayarlayın.
- **Örnek veri:** Ürün, limitler ve sonuçlar kurgusaldır; gerçek bir ürünün spesifikasyonu değildir.

## Testler

- İki limit, tek limit, nitel sonuç ve `<`, `>` işaretli sonuçların her sınır değeri.
- Parti kararı önceliği, eksik analiz, tekrarlı analiz, spesifikasyonu olmayan parametre, parametre özeti.
- Okunamayan spesifikasyon ve sonuç satırları, komut satırı hataları; örnek veri uçtan uca.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
