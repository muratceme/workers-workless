# Periyodik Bakım İş Emri Listesi · Kod Bloğu

> Bakım Onarım › Bakım Teknisyeni · Gerisi Bende

Bakım planından **seçilen haftanın periyodik bakım iş emirlerini** çıkarır: o hafta günü gelen ve gecikmiş bakımları
numaralar, gereken malzemeleri toplar, gün ve sorumlu bazında iş yükünü gösterir. İnternete bağlanmaz.

## Nasıl hesaplar?

| Kavram | Kural |
|---|---|
| Sonraki bakım tarihi | Son bakım tarihi + periyot |
| Periyot yazımı | `30` ya da `30 gün`, `2 hafta`, `3 ay`, `1 yıl`; ay ve yıl takvim ayıyla eklenir (31 Ocak + 1 ay = 28/29 Şubat) |
| Hafta | `--hafta` ile verilen tarihin pazartesisinden pazarına |
| Bu hafta | Sonraki bakım tarihi hafta içinde; planlanan gün o tarih |
| Gecikmiş | Sonraki bakım tarihi haftadan önce; gecikme = hafta başı − o tarih; planlanan gün hafta başı |
| İlk bakım | Son bakım tarihi yazılmamış; planlanan gün hafta başı |
| Malzeme ihtiyacı | İş emirlerinin `Malzemeler` hücrelerindeki kalemlerin toplamı |

`Malzemeler` hücresine kalemler virgülle ayrılarak, miktar `x` ile yazılır: `Gres (kg) x2, Temizlik bezi`. Miktarı
yazılmayan kalem 1 sayılır. İş emri numarası `PB-YYYYHH-sıra` biçimindedir (yıl ve ISO hafta numarası).

**Örnek:** Periyot 3 ay, son bakım 14.07.2026 → sonraki bakım **14.10.2026**; 12–18.10.2026 haftasında "Bu hafta"
olarak çarşamba gününe planlanır.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                              # örnek veri, 12.10.2026 haftası
python main.py --plan bakim_plani.xlsx
python main.py --plan bakim_plani.xlsx --hafta 19.10.2026 --kritik 14 --kapasite 30
```

Zorunlu sütunlar: Ekipman, Bakım Adı, Periyot.

## Çıktı

`cikti/periyodik_bakim_is_emirleri.xlsx`:

- `İş Emirleri`: numara, durum, planlanan gün, ekipman, bakım, gecikme, süre, sorumlu, malzemeler; doldurmanız için
  "Yapıldı" ve "Not" sütunları.
- `Malzeme İhtiyacı`: kalem bazında toplam miktar ve hangi ekipmanlar için gerektiği.
- `Haftalık Yük`: gün ve sorumlu bazında iş emri sayısı ve tahmini süre.
- `Sonraki Hafta`: bir sonraki hafta günü gelecek bakımlar.
- `Uyarılar`.

## Dikkat

- **Yalnız takvim bazlı bakımlar.** Çalışma saati, çevrim sayısı ya da sayaç değerine bağlı bakımlar hesaplanmaz;
  bunları tahmini tarihe çevirip plana siz yazmalısınız.
- **Sonraki tarih, son yapılan bakımdan sayılır.** Bakım geç yapıldıysa izleyen bakım da kayar. Sabit takvime
  (her ayın ilk haftası gibi) bağlı kalmak istiyorsanız son bakım tarihine planlanan tarihi yazın.
- **Plan dosyasını siz güncellersiniz.** Kod yapılan bakımı kaydetmez; iş emri tamamlanınca plandaki son bakım
  tarihini değiştirmezseniz aynı bakım sonraki hafta "gecikmiş" olarak yeniden çıkar.
- **Planlanan gün bir öneridir.** Üretim planını, duruş pencerelerini, hafta sonu ve tatilleri kod bilmez; günleri
  üretimle birlikte kesinleştirin. Kapasite uyarısı yalnız tahmini sürelerin toplamına bakar.
- **Malzeme stoğu kontrol edilmez.** Liste ihtiyacı gösterir; depoda olup olmadığına bakmaz.
- **Yasal periyodik kontroller ayrıdır.** Basınçlı kap, kaldırma ekipmanı ve elektrik tesisatı gibi mevzuatla
  zorunlu periyodik kontroller yetkili kişilerce yapılır; bu liste onların yerine geçmez.
- **Örnek veri:** Ekipmanlar, periyotlar ve malzemeler kurgusaldır.

## Testler

- Periyot yazımları (gün, hafta, ay, yıl), ay sonu taşması, hafta sınırları (pazartesi ve pazar dahil).
- Gecikmiş, bu hafta ve ilk bakım ayrımı, iş emri sıralaması ve numaralandırma, sonraki hafta listesi.
- Malzeme ayrıştırma ve toplama, sorumlu ve gün yükü, kapasite ve kritik gecikme uyarıları, okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

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
