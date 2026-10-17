# Faiz Riski Vade Uyumsuzluk (Gap) Analizi · Kod Bloğu

> Hazine ve Aktif-Pasif Yönetimi › Aktif-Pasif Yönetimi Uzmanı · Gerisi Bende

Faize duyarlı varlık ve yükümlülükleri **yeniden fiyatlama vadelerine** göre dilimlere dağıtır; gap tablosunu ve paralel
faiz şokunun net faiz gelirine ve ekonomik değere yaklaşık etkisini hesaplar. İnternete bağlanmaz.

## Nasıl hesaplar?

| Adım | Kural |
|---|---|
| Yeniden fiyatlama günü | Sabit faizli kalemde vade tarihi; değişken faizlide bir sonraki faiz belirleme tarihi |
| Süre t (yıl) | (Yeniden fiyatlama günü − rapor tarihi) ÷ 365; geçmiş tarih 0 sayılır |
| Dilimler | 0-1 ay (≤ 30 gün), 1-3 ay (≤ 91), 3-6 ay (≤ 182), 6-12 ay (≤ 365), 1-2 yıl (≤ 730), 2-5 yıl (≤ 1826), 5 yıl+ |
| Gap | Dilimdeki faize duyarlı varlıklar − faize duyarlı yükümlülükler |
| Kümülatif gap | Dilim gap'lerinin birikimli toplamı |
| Net faiz geliri etkisi (12 ay) | ± tutar × şok × (1 − t), yalnız t < 1 yıl için (varlık +, yükümlülük −) |
| Ekonomik değer etkisi | ∓ tutar × t ÷ (1 + faiz oranı) × şok (varlık −, yükümlülük +) |

- **Net faiz geliri:** Kalem t anında yeni orandan fiyatlanır ve yılın kalanında şok kadar fazla / eksik faiz işler.
  Bilançonun sabit kaldığı, şokun anlık ve tüm vadelerde aynı (paralel) olduğu varsayılır.
- **Ekonomik değer:** t, düzeltilmiş süre yerine kullanılan yaklaşık değerdir. Ara kupon ve taksitler yok sayıldığı
  için uzun vadeli kalemlerde etki olduğundan büyük çıkar. Faiz oranı girilmemişse iskonto yapılmaz.
- **Faize duyarsız kalemler** (nakit, faizsiz mevduat, özkaynak, sabit kıymet) gap'e girmez; ayrı satırda gösterilir.
- 1 yıla kadar kümülatif gap artıysa portföy "varlığa duyarlı", eksiyse "yükümlülüğe duyarlı" olarak yorumlanır.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek veri
python main.py --kalemler bilanco.xlsx --tarih 30.09.2026 --sok 100,-100,300 --ozkaynak 2.400.000.000
python main.py --kalemler bilanco.xlsx --para-birimi USD --sok 200,-200
```

| Sütun | Açıklama |
|---|---|
| Kalem, Tür, Tutar | Zorunlu. Tür: Varlık / Yükümlülük (Aktif / Pasif / Özkaynak da kabul edilir) |
| Faiz Tipi | Sabit, Değişken veya Faizsiz. Boşsa sabit sayılır |
| Vade Tarihi | Sabit faizli kalemin yeniden fiyatlama günü |
| Yeniden Fiyatlama Tarihi | Değişken faizli kalemde bir sonraki faiz belirleme tarihi |
| Faiz Oranı | Ekonomik değer hesabında iskonto için (isteğe bağlı) |
| Para Birimi | Boşsa TL |

## Çıktı

`cikti/faiz_gap_analizi.xlsx`:

- `Gap Tablosu`: dilim bazında varlık, yükümlülük, gap, kümülatif gap ve toplam varlığa oranları.
- `Şok Etkisi`: her şok için net faiz geliri ve ekonomik değer etkisi (özkaynak girildiyse oranı); dilim kırılımı.
- `Kalemler`: her kalemin yeniden fiyatlama günü ve dilimi.
- `Uyarılar`: tarihi olmayan veya geçmiş kalemler, okunamayan satırlar.

## Dikkat

- **Yasal ölçüm değildir.** Bankacılık hesaplarından kaynaklanan faiz oranı riskinin ölçümü ve raporlaması BDDK
  düzenlemelerindeki vade dilimleri, şok büyüklükleri ve varsayımlara göre yapılır. Bu paket iç analiz ve kontrol
  amaçlı basitleştirilmiş bir yaklaşımdır; şok büyüklüklerini `--sok` ile kendiniz belirlersiniz.
- **Vadesiz kalemler:** Vadesiz mevduat ve kredi kartı gibi sözleşme vadesi olmayan kalemlerin yeniden fiyatlama
  davranışı bir varsayımdır. Tarih girilmezse kalem faize duyarsız sayılır; çekirdek mevduat varsayımınızı tarih
  girerek yansıtın.
- **Taksitli krediler** tek satırda tek vadeyle girilirse nakit akışları doğru dağılmaz. Daha doğru sonuç için
  krediyi yeniden fiyatlama dilimlerine bölerek girin.
- **Kapsam dışı:** baz riski, verim eğrisi şeklindeki değişim (paralel olmayan şok), erken ödeme ve erken çekim
  opsiyonları, faiz tavan / tabanları.
- **Çok para birimi:** Tutarlar toplanırken çevrim yapılmaz. Para birimlerini `--para-birimi` ile ayrı çalıştırın.
- **Örnek veri:** Bilanço kalemleri ve oranlar kurgusaldır.

## Testler

- Dilim sınırları (30 / 31, 365 / 366 gün...), gap ve kümülatif gap.
- Net faiz geliri ve ekonomik değer formülleri elle hesaplanan değerlerle.
- Okuma kuralları: değişken faiz, tarihsiz kalem, geçmiş tarih, özkaynak, para birimi filtresi.
- Örnek veri uçtan uca.

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
