# Yedek Parça Kritik Stok Önerisi · Kod Bloğu

> Bakım Onarım › Bakım Mühendisi · Workers / Workless

Arıza ve bakımda kullanılan parça kayıtları ile tedarik süresinden, yedek parçalar için **minimum stok seviyesi**
önerir; mevcut stokla karşılaştırıp eksik miktarı ve sipariş tutarını verir. İnternete bağlanmaz.

## Nasıl hesaplar?

| Kavram | Formül |
|---|---|
| Günlük tüketim | Dönemde kullanılan miktar ÷ `--gun` (365) |
| Tedarik süresi talebi | Günlük tüketim × tedarik süresi (gün) |
| Emniyet payı | Tedarik süresi talebi × kritiklik yüzdesi (`--emniyet 100;50;25`: Yüksek, Orta, Düşük) |
| Minimum stok | Yukarı yuvarla(tedarik süresi talebi + emniyet payı) |
| Taban (Yüksek ve Orta) | En az "tek seferde kullanılan en büyük miktar" |
| Taban (Yüksek) | Ayrıca en az 1 adet, dönemde hiç kullanılmamış olsa bile |
| Eksik miktar | Minimum stok − mevcut stok (sıfırdan büyükse) |
| Sipariş tutarı | Eksik miktar × birim fiyat |

Durum: mevcut stok sıfır ve minimum sıfırdan büyükse "Stok yok"; minimumun altındaysa "Minimumun altında"; dönemde
hiç kullanılmamış ve stoğu minimumun üstündeyse "Hareketsiz"; diğerleri "Yeterli". Rapor, her parçada minimumu
hangi kuralın belirlediğini ayrı sütunda yazar.

**Örnek:** Orta kritiklikte bir kayış yılda 6 adet kullanılmış, tedarik süresi 14 gün → tedarik süresi talebi 0,23,
emniyet payıyla 0,35 → yukarı yuvarlanınca 1. Ancak bir işte en çok 2 adet birlikte kullanılmış → minimum stok
**2**. Mevcut stok 1 → eksik **1**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                              # örnek veri, dönem 365 gün
python main.py --parcalar parcalar.xlsx --kullanim parca_kullanimi.xlsx
python main.py --parcalar parcalar.xlsx --kullanim parca_kullanimi.xlsx --gun 730 --emniyet 150;75;0
```

Zorunlu sütunlar: parçalarda Parça Kodu ve Tedarik Süresi; kullanımda Parça Kodu ve Miktar. `--gun`, kullanım
dosyasının kapsadığı dönemin uzunluğudur; yanlış verilirse günlük tüketim yanlış çıkar.

## Çıktı

`cikti/yedek_parca_kritik_stok.xlsx`:

- `Stok Önerisi`: önce stokta olmayan ve minimumun altındaki parçalar; tüketim, önerilen minimum, minimumu belirleyen
  kural, eksik miktar, sipariş tutarı, son kullanım ve kullanıldığı ekipmanlar.
- `Özet`: durum bazında parça sayısı ve toplam sipariş tutarı.
- `Uyarılar`.

## Dikkat

- **Bu bir kural hesabıdır, istatistiksel model değildir.** Yedek parça tüketimi seyrek ve düzensizdir; ortalama
  tüketime dayalı hesap, nadir ama pahalı arızaları iyi temsil etmez. Kritik ve pahalı parçalarda sonucu
  mühendislik değerlendirmesiyle (arıza olasılığı, duruş maliyeti, parçanın onarılabilirliği) birlikte kullanın.
- **Emniyet yüzdeleri varsayımdır.** %100 / %50 / %25, kritikliği yansıtmak için seçilmiş başlangıç değerleridir;
  hedef servis düzeyinden türetilmemiştir. Kendi politikanıza göre değiştirin.
- **Kritikliği siz belirlersiniz.** Kod, parçanın yokluğunda hattın durup durmayacağını bilemez; "Yüksek" dediğiniz
  her parça için en az 1 adet önerir.
- **Tedarik süresi tek değerdir.** Gümrük, minimum sipariş miktarı, tedarikçinin stok durumu ve sürenin değişkenliği
  hesaba girmez.
- **Yalnız minimum seviyeyi verir.** Sipariş miktarı, azami stok ve ekonomik sipariş büyüklüğü hesaplanmaz.
- **Geçmiş, geleceği göstermeyebilir.** Makine parkı yaşlandıkça ya da yenilendikçe tüketim değişir; kayıtlara
  girmeyen (depodan fişsiz çıkan) kullanımlar görünmez.
- **Örnek veri:** Parçalar, fiyatlar ve kullanım kayıtları kurgusaldır.

## Testler

- Günlük tüketim, tedarik süresi talebi, emniyet payı ve yukarı yuvarlama; kritikliğe göre taban kuralları.
- Durum ataması (stok yok, minimumun altında, yeterli, hareketsiz), eksik miktar ve sipariş tutarı.
- Listede olmayan parçanın kullanımı, fiyatı olmayan parça, emniyet yüzdelerinin değiştirilmesi, okuma ve komut satırı.

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
