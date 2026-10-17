# Tahsilat ve Ödeme Listesi · Kod Bloğu

> Muhasebe › Ön Muhasebe Elemanı · Gerisi Bende

Açık alacak ve borç kalemlerinden **haftalık tahsilat ve ödeme listesini** ve günlük nakit etkisini çıkarır.
İnternete bağlanmaz.

## Nasıl çalışır?

| Kavram | Kural |
|---|---|
| Dönem | `--tarih` gününden başlayan `--gun` (7) gün |
| Vadesi geçmiş | Vadesi dönem başından önce; gecikme = dönem başı − vade |
| Bu dönem | Vadesi dönem içinde |
| İleri vadeli | Vadesi dönem sonundan sonra; listeye girmez, toplamı gösterilir |
| Günlük net | O gün vadesi gelen tahsilat − o gün vadesi gelen ödeme |
| Kümülatif nakit | Başlangıç kasası − vadesi geçmiş ödemeler + günlük netler |

Kümülatif nakit ihtiyatlı hesaplanır: gecikmiş borç hemen ödenecek, gecikmiş alacak dönem içinde tahsil
edilemeyecek sayılır.

**Örnek:** Kasa 10.000, gecikmiş ödeme 15.000 → başlangıç −5.000. Pazartesi 20.000 tahsilat → 15.000. Salı 30.000
ödeme → −15.000: nakit açığı uyarısı.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek veri, dönem 12–18.10.2026, kasa 10.000
python main.py --alacaklar alacaklar.xlsx --borclar borclar.xlsx
python main.py --alacaklar alacaklar.xlsx --borclar borclar.xlsx --tarih 12.10.2026 --gun 14 --kasa 250000
```

İki dosya da aynı biçimdedir: Cari, Vade, Tutar (zorunlu); Belge No, Açıklama. Yalnız birini de verebilirsiniz.
Tutar, kalemin **kalan (açık)** tutarıdır.

## Çıktı

`cikti/tahsilat_odeme_listesi.xlsx`:

- `Tahsilat Listesi`, `Ödeme Listesi`: vade sırasıyla kalemler, gecikme günü, sonuç için boş sütun; altta toplamlar.
- `Günlük Özet`: gün bazında tahsilat, ödeme, net ve kümülatif nakit.
- `Cari Özeti`: cari bazında vadesi geçmiş ve bu dönem tutarları.
- `Uyarılar`: nakit açığı, gecikmiş ödemeler, 30 günü aşan alacaklar, hafta sonuna denk gelen vadeler.

## Dikkat

- **Vade günü tahsilat günü değildir.** Liste ne zaman tahsil etmeniz gerektiğini gösterir; müşterinin o gün
  ödeyeceğini değil. Nakit planını tahsilat alışkanlıklarına göre ayrıca değerlendirin.
- **Tek para birimi.** Dövizli kalemleri kendi kurunuzla çevirip girin; kod kur bilmez.
- **Çek ve senetler.** Portföydeki çek ve senetleri vadesiyle alacak listesine, verdiğiniz çekleri borç listesine
  ekleyebilirsiniz; karşılıksız çıkma riski hesaba girmez.
- **Düzenli giderler listede yoksa nakit fazla görünür.** Maaş, kira, vergi ve SGK ödemelerini borç listesine ekleyin.
- **Hafta sonu ve resmi tatil.** Banka işlemleri izleyen iş gününe kayar; kod vadeyi kaydırmaz, yalnız uyarır.
- **Örnek veri:** Cariler, belgeler ve tutarlar kurgusaldır.

## Testler

- Vadesi geçmiş, bu dönem ve ileri vadeli ayrımı; gecikme günü; cari özetleri.
- Günlük net ve kümülatif nakit; nakit açığı uyarısının her iki biçimi; kasa verilmeden çalışma.
- Dönem uzunluğu, 30 günü aşan alacak, hafta sonu vadesi; okuma ve komut satırı hataları; örnek veri uçtan uca.

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
