# Ödeme Önceliklendirme · Kod Bloğu

> Finans › Finans Uzmanı · Workers / Workless

Nakit dar olduğunda "bu hafta kime ödeyeceğiz?" sorusunu **açık ve her seferinde aynı sonucu veren bir kuralla**
yanıtlar: ödeme listesini sıralar, eldeki nakit bitene kadar "Öde", sonrası için "Ertele" der ve her kararın
gerekçesini yazar. İnternete bağlanmaz.

## Nasıl hesaplar?

Önce dönem belirlenir: `--tarih` gününden başlayan `--gun` (7) gün. Vadesi dönem sonundan sonra olan kalemler
"Vadesi gelmedi" sayılır ve sıraya girmez. Kalan kalemler şu sırayla dizilir:

1. **Zorunlu** sütunu "Evet" olanlar,
2. **Önem**: Kritik → Yüksek → Normal → Düşük,
3. vadesi geçmiş olanlar, sonra dönem içinde vadesi gelenler,
4. **günlük gecikme maliyeti** yüksek olan (tutar × aylık gecikme faizi % ÷ 30),
5. vadesi erken olan,
6. tutarı küçük olan.

Sonra kullanılabilir nakit (`--nakit` − `--asgari`) bu sırayla dağıtılır. Kalem kalan nakde sığıyorsa "Öde",
sığmıyorsa "Ertele" denir ve sıradaki kaleme geçilir; kısmi ödeme önerilmez.

**Örnek:** 150.000 TL borç, aylık %3 gecikme faizi → günlük gecikme maliyeti 150.000 × 3 ÷ 100 ÷ 30 = **150 TL**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                            # örnek veri: 900.000 nakit, dönem 12–18.10.2026
python main.py --odemeler odemeler.xlsx --nakit 900000
python main.py --odemeler odemeler.xlsx --nakit 1.250.000 --asgari 100.000 --tarih 12.10.2026 --gun 10
```

Zorunlu sütunlar: Alacaklı, Vade, Tutar. Diğerleri: Belge No, Tür, Önem (boşsa Normal), Gecikme Faizi (aylık %),
Zorunlu (Evet). Kendi dosyanızla çalışırken `--nakit` vermek zorunludur.

## Çıktı

`cikti/odeme_onceliklendirme.xlsx`:

- `Ödeme Sırası`: sıra, karar, alacaklı, vade, tutar, gerekçe, günlük gecikme maliyeti, ödeme sonrası kalan nakit ve
  sizin dolduracağınız onay sütunu.
- `Özet`: nakit, dönem ihtiyacı, ödenecek ve ertelenen toplamlar, ertelenenlerin günlük gecikme maliyeti.
- `Uyarılar`: karşılanamayan zorunlu ödemeler, ertelenen kritik ödemeler, nakit açığı, ertelenen kalemler.

## Dikkat

- **"Zorunlu" ve "Önem" sizin kararınızdır.** Kod, hangi ödemenin yasal olarak ertelenemeyeceğini (vergi, SGK
  primi, ücret gibi) veya hangi tedarikçinin vazgeçilmez olduğunu bilmez; bu sütunları doldurmazsanız sıra yalnız
  vadeye ve gecikme maliyetine göre oluşur.
- **Öneridir, ödeme talimatı değildir.** Sırayı yetkili kişi onaylamalıdır; kod bankaya bağlanmaz ve ödeme yapmaz.
- **Gecikme maliyeti yaklaşıktır.** Yalnız sizin yazdığınız aylık oranla basit hesaplanır; sözleşme cezaları,
  yasal gecikme zammı, vade farkı faturaları ve itibar kaybı hesaba girmez.
- **Büyük kalem atlanabilir.** Nakde sığmayan bir kalem ertelenirken sırası daha geride olan küçük kalemler
  ödenebilir. Kısmi ödeme veya vade görüşmesi daha doğruysa listeyi buna göre düzenleyip yeniden çalıştırın.
- **Tek para birimi.** Dövizli borçları kendi kurunuzla çevirip girin. Dönem içindeki tahsilatlar hesaba katılmaz.
- **Örnek veri:** Alacaklılar, tutarlar ve oranlar kurgusaldır.

## Testler

- Sıralama kuralının her adımı, nakit dağıtımı, sığmayan kalemin atlanması, asgari kasa.
- İleri vadeli kalemler, toplamlar, uyarılar (zorunlu, kritik, nakit açığı).
- Okuma hataları ve komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
