# Bakım Maliyet ve Performans Raporu · Kod Bloğu

> Bakım Onarım › Bakım Müdürü · Gerisi Bende

Bakım iş emirlerinden **işçilik, malzeme ve dış hizmet maliyetlerini** toplar; planlı bakımın payını, arıza sayısını
ve ortalama onarım süresini ekipman, bakım türü ve ay bazında raporlar. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Toplam maliyet | İşçilik + malzeme + dış hizmet |
| Planlı bakım | Türü Planlı, Periyodik, Önleyici, Koruyucu ya da Kestirimci olan iş emirleri |
| Plansız bakım | Türü Arıza, Plansız, Acil ya da Düzeltici olan iş emirleri |
| Planlı bakım payı % | Planlı bakım maliyeti ÷ toplam maliyet × 100 |
| Arıza sayısı | Plansız iş emri adedi |
| Ortalama onarım süresi (MTTR) | Plansız duruş saati ÷ arıza sayısı |
| Arızalar arası süre (MTBF) | (`--calisma-saat` − plansız duruş saati) ÷ arıza sayısı |

Planlı bakım payı ayrıca iş emri adedi ve işçilik saati üzerinden de verilir. `--calisma-saat`, dönemde **ekipman
başına** planlı çalışma saatidir; verilmezse MTBF hesaplanmaz. Toplam satırında çalışma saati ekipman sayısıyla
çarpılır.

Uyarılar: planlı bakım maliyet payı `--planli` (70) yüzdesinin altındaysa; bir ekipmanda arıza sayısı `--ariza` (3)
ve üstündeyse; bakım türü tanınmadıysa.

**Örnek:** Bir ekipmanda dönemde 3 arıza, toplam 22 saat plansız duruş → ortalama onarım süresi **7,33 saat**.
Dönemde 1.500 saat çalışması planlanmışsa arızalar arası süre (1.500 − 22) ÷ 3 = **492,67 saat**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                              # örnek veri
python main.py --emirler bakim_is_emirleri.xlsx
python main.py --emirler bakim_is_emirleri.xlsx --planli 80 --ariza 2 --calisma-saat 1500
```

Zorunlu sütunlar: Ekipman, Bakım Türü.

## Çıktı

`cikti/bakim_maliyet_performans_raporu.xlsx`:

- `Özet`: maliyet kalemleri, planlı bakım payı (üç ölçüye göre), arıza, duruş, MTTR ve MTBF.
- `Ekipmanlar`: en maliyetliden başlayarak; maliyet payı ve aynı göstergeler.
- `Bakım Türleri`, `Aylar`, `İş Emirleri`, `Uyarılar`.

## Dikkat

- **Duruşun üretim kaybı maliyete dahil değildir.** Rapor yalnız bakım harcamasını gösterir; arızanın asıl maliyeti
  çoğu zaman kaybedilen üretimdir ve burada yoktur.
- **Planlı bakım payı için %70 bir hedef varsayımıdır.** Uygun oran ekipman yaşına, sektöre ve bakım stratejisine
  göre değişir; pay maliyete göre hesaplandığında tek pahalı arıza oranı ciddi biçimde düşürür, bu yüzden adet ve
  saat paylarıyla birlikte okuyun.
- **MTTR burada duruş süresidir.** Yalnız fiili onarım süresini değil, parça bekleme ve devreye alma dahil
  yazdığınız duruşu ölçer. MTBF için verdiğiniz çalışma saati bütün ekipmanlar için aynı kabul edilir; vardiya
  düzenleri farklıysa ekipmanları ayrı çalıştırın.
- **Her plansız iş emri bir arıza sayılır.** Aynı arızanın devamı için açılan ikinci iş emri arıza sayısını
  şişirir; küçük ayarlar için açılan iş emirleri de öyle.
- **İşçilik maliyetini siz verirsiniz.** Kod saat ücretinden maliyet hesaplamaz, fazla mesaiyi ayırmaz.
- **Kapsam dışı:** bütçe karşılaştırması, ekipman yenileme değeri oranı, stok maliyeti, garanti kapsamındaki işler.
- **Örnek veri:** Ekipmanlar, iş emirleri ve tutarlar kurgusaldır.

## Testler

- Maliyet toplamları, planlı bakım payı (maliyet, adet ve saat), arıza sayısı, MTTR ve MTBF.
- Tür eş anlamlıları ve tanınmayan tür, ekipman payı ve sıralaması, ay kırılımı.
- Planlı pay hedefi ve tekrarlayan arıza eşiği, okuma ve komut satırı.

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
