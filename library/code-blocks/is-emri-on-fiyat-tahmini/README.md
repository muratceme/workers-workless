# İş Emri Ön Fiyat Tahmini · Kod Bloğu

> Satış Sonrası Hizmetler (Servis) › Servis Danışmanı · Workers / Workless

Müşterinin şikâyeti ya da bakım türüne karşılık gelen **iş tanımlarını** (işçilik saati + parça listesi) parça fiyat listesi ve
saat ücretiyle fiyatlandırıp iş emri bazında **ön fiyat tahmini** ve tahmin aralığı çıkarır. İnternete ve DMS'e bağlanmaz.

## Nasıl hesaplar?

| Kalem | Formül |
|---|---|
| İşçilik | İş tanımındaki saat × adet × `--saat-ucreti` |
| Parça | İş tanımındaki miktar × adet × parça birim fiyatı (KDV hariç) |
| İskonto | İşçilik × `--iscilik-iskonto` % + parça × `--parca-iskonto` % |
| Net / KDV / toplam | Net = işçilik + parça − iskonto · KDV = net × `--kdv` % (varsayılan 20) · toplam = net + KDV |
| Tahmin aralığı | Toplam × (1 ± `--pay` %), varsayılan ±%10 |

Bir iş emrinde birden fazla iş kodu olabilir (her biri ayrı satır). İş tanımlarında aynı iş kodunun her satırı bir işçilik ya da
parça kalemidir.

**Uyarılar:** iş kodu tanımsız · parça fiyat listesinde yok (tutar 0 yazılır, iş emri "Eksik fiyat" işaretlenir) · okunamayan
satırlar (Yüksek) · iş tanımında işçilik yok (Bilgi).

**Örnek:** 1,5 saat × 2.000 TL işçilik + 4 lt × 500 TL yağ + 400 TL filtre; işçilikte %10, parçada %5 iskonto → net **4.980 TL**,
KDV %20 ile **5.976 TL** (aralık 5.378,40 – 6.573,60 TL).

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                          # örnek veri (saat ücreti 2.400 TL)
python main.py --talepler talepler.xlsx --isler is_tanimlari.xlsx --parcalar parcalar.xlsx --saat-ucreti 2400
python main.py --talepler t.xlsx --isler i.xlsx --parcalar p.xlsx --saat-ucreti 2400 --iscilik-iskonto 10 --parca-iskonto 5 --pay 15
```

Kendi verinizde `--saat-ucreti` zorunludur.

## Çıktı

`cikti/is_emri_on_fiyat.xlsx`: `Ön Fiyat Özeti` (eksik fiyatlı iş emirleri kırmızı), `Detay`, `Uyarılar`.

## Dikkat

- **Bu bir ön tahmindir.** Kesin fiyat araç incelendikten sonra belli olur; ek iş ya da parça gerekirse işe başlamadan müşteriye
  bildirip onayını alın ve iş emrine yazın. Tüketici mevzuatı ve satış sonrası hizmet yönetmeliklerindeki bilgilendirme
  yükümlülüklerini hukuk biriminizle doğrulayın.
- **Fiyatlar sizin listenizdir.** Saat ücreti, parça fiyatları, iskonto ve KDV oranı girdidir; örnek değerler yalnız örnektir.
  KDV oranını güncel mevzuattan doğrulayın.
- **İşçilik süreleri** marka iş süresi tablolarından alınmalıdır; aynı iş emrindeki örtüşen işlemler (ör. aynı sökme işlemi)
  için süre indirimi yapılmaz.
- **Örnek veri:** Müşteriler, plakalar, parça kodları ve fiyatlar kurgusaldır.

## Testler

- İşçilik ve parça tutarları, adet çarpanı, iskontolar, KDV, toplam ve aralık; birden fazla iş kodlu iş emri.
- Tanımsız iş kodu, fiyatı olmayan parça ve işçiliksiz iş uyarıları; okuma ve komut satırı.

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
