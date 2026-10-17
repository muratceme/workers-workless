# Servis Randevu Kapasite Planı · Kod Bloğu

> Satış Sonrası Hizmetler (Servis) › Servis Danışmanı · Gerisi Bende

Teknisyenlerin çalışma saati, verimliliği, çalışma günleri ve izinleriyle **günlük atölye kapasitesini** hesaplar; randevuları
iş türlerinin standart süreleriyle yükleyip **doluluk, kalan saat ve o güne daha kaç iş eklenebileceğini** gösterir. İnternete
ve DMS'e bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Çalışan teknisyen | Servis açık + teknisyenin çalışma günü + izinli değil |
| Mevcut saat | Çalışan teknisyenlerin günlük saatleri toplamı |
| Kapasite (saat) | Σ günlük saat × verimlilik % × (1 − `--rezerv` %) |
| Planlanan | Randevu süreleri: satırda Süre yoksa iş türünün standart süresi ("Bakım + Fren" gibi birden fazla iş toplanır) |
| Doluluk % | Planlanan ÷ kapasite × 100 |
| Eklenebilir iş | ⌊kalan saat ÷ iş türünün standart süresi⌋, her iş türü için |

Çalışma günleri `Pzt-Cmt`, `Pazartesi, Çarşamba, Cuma` ya da `Cmt-Pzt` gibi yazılabilir (boş = her gün). İzin tek tarih ya da
`19.10.2026-21.10.2026` aralığı olarak, virgülle ayrılmış liste hâlinde yazılır. `--kapali` (varsayılan Paz) servisin kapalı
olduğu günlerdir.

**Uyarılar:** kapasite aşıldı (doluluk > %100) · randevusu olan günde çalışan teknisyen yok · iş türü tanımsız ve süre yazılmamış
(Yüksek) · doluluk `--hedef` (varsayılan %85) üstünde · kapalı güne randevu (Orta) · doluluk %50'nin altında (boş kapasite) ·
plan dönemi dışındaki randevular (Bilgi).

**Örnek:** 7,5 saat × %120, 7,5 saat × %80 ve 9 saat × %100 çalışan üç teknisyen, %10 rezervle: (9 + 6 + 9) × 0,9 = **21,6 saat**
kapasite; 22 saatlik randevu → doluluk **%101,85**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                           # örnek veri (19–24.10.2026)
python main.py --teknisyenler teknisyenler.xlsx --isler is_turleri.xlsx --randevular randevular.xlsx
python main.py --teknisyenler t.xlsx --isler i.xlsx --randevular r.xlsx --baslangic 26.10.2026 --gun 12 --rezerv 15 --hedef 90
python main.py --teknisyenler t.xlsx --isler i.xlsx --randevular r.xlsx --kapali "Cmt, Paz"
```

`--baslangic` verilmezse ilk randevu günü alınır.

## Çıktı

`cikti/servis_kapasite_plani.xlsx`: `Günlük Kapasite` (doluluk > %100 kırmızı, hedef üstü sarı, %50 altı mavi), `Eklenebilir İş`,
`Randevular`, `Uyarılar`.

## Dikkat

- **Standart süreler ve verimlilik sizin verinizdir.** Marka iş süresi (flat rate) tablolarını ve teknisyenlerin gerçekleşen
  verimliliğini kullanın; örnek değerler yalnız örnektir.
- **Çalışma saatleri yasal sınırlar içinde olmalı.** Günlük saatleri girerken 4857 sayılı İş Kanunu'ndaki haftalık 45 saat ve
  fazla çalışma kurallarını gözetin (örnekte Pzt-Cmt çalışanlar 7,5, Pzt-Cum çalışanlar 9 saattir). Güncel kuralı mevzuattan
  doğrulayın.
- **Kapasite yalnız işçilik saatidir.** Lift / ekipman, parça bulunabilirliği, yıkama ve kalite kontrol süreleri hesaba girmez.
- **Örnek veri:** Teknisyenler, plakalar ve randevular kurgusaldır.

## Testler

- Çalışma günü ve izin, verimlilik ve rezervle kapasite, doluluk, kalan saat, eklenebilir iş, kapalı gün, toplam satırı.
- Kapasite / hedef / boş kapasite / tanımsız iş / dönem dışı uyarıları; gün adı ve tarih aralığı çözümleme, okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

## Kaynaklar ve dayanaklar

Kontrol tarihi: 15.10.2026. Mevzuat sık değişir; paketteki süre, oran ve kuralları kullanmadan önce güncel resmî metinle
karşılaştırın. "Karşılaştırılan maddeler" dışındaki atıflar yalnız yol göstericidir.

| Kaynak | Bağlantı | Durum |
|---|---|---|
| 4857 sayılı İş Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.4857.pdf) | Resmî metinle karşılaştırılan maddeler: 17, 41, 46, 47, 53, 56, 63, 68, 69 |

## Sorumluluk

Bu paketi indiren ve çalıştıran kişi aşağıdakileri kabul etmiş sayılır:

- **Tüm risk kullanıcıdadır.** Paket ücretsizdir ve "olduğu gibi" sunulur; doğruluk, eksiksizlik veya belirli bir amaca
  uygunluk konusunda hiçbir garanti verilmez. Paketi kendi bilgisayarınızda, kendi verinizle ve kendi kararınızla
  çalıştırırsınız. Hatalı veya eksik sonuç, kaçırılan süre, vergi / SGK cezası, veri kaybı, mali kayıp ve diğer her türlü
  doğrudan ya da dolaylı zarardan yazarlar ve katkıda bulunanlar sorumlu tutulamaz.
- **Sonuçları kontrol etmek sizin işinizdir.** Çıktılar karar desteği içindir. Beyan vermeden, ödeme yapmadan, sözleşme
  imzalamadan, personel / kredi / hasar / tedavi kararı almadan önce sonuçları yetkin bir kişi kontrol etmelidir.
- **Danışmanlık değildir.** Hukuki, mali, vergisel, tıbbi veya yatırım danışmanlığı yerine geçmez.
- **Mevzuat değişir.** Koddaki süre, oran, eşik ve kurallar yazıldığı tarihteki bilgiye dayanır ve eskiyebilir; özel
  durumları kapsamayabilir. Güncelliğini kontrol etmek kullanıcının sorumluluğundadır.
- **Örnek veri kurgusaldır.** Kişi, kurum, tutar ve tarihler gerçek değildir; kendi verinizle sonuçlar farklı olur.
- **Kişisel veri.** Verinizin KVKK, şirket politikaları ve sektör düzenlemelerine uygun işlenmesinden, saklanmasından ve
  paylaşılmasından siz sorumlusunuz.

Ayrıntılar: `LICENSE` (MIT) ve `SORUMLULUK_REDDI.md`.
