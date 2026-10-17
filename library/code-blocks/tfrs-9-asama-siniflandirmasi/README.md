# TFRS 9 Aşama Sınıflandırması · Kod Bloğu

> Krediler İzleme ve Takip › Kredi İzleme Uzmanı · Gerisi Bende

Kredileri gecikme günü ve risk göstergelerine göre **Aşama 1-2-3** olarak sınıflandırır, her kredinin gerekçesini yazar
ve önceki döneme göre aşaması değişenleri listeler. Eşikler parametredir. İnternete bağlanmaz.

## Nasıl hesaplar?

| Aşama | Ölçüt (biri yeterli) | Parametre |
|---|---|---|
| 3 — temerrüt | Gecikme > 90 gün | `--temerrut-gun` |
| | "Temerrüt" sütunu Evet | |
| 2 — kredi riskinde önemli artış | Gecikme > 30 gün | `--asama2-gun` |
| | Yeniden yapılandırma veya yakın izleme sütunu Evet | |
| | Not, ilk nota göre 3 kademe veya daha fazla kötüleşmiş | `--not-esik`, `--not-olcegi` |
| | Güncel temerrüt olasılığı (TO) ÷ ilk TO ≥ 2 | `--to-kat` |
| 1 | Diğerleri | |

- Gecikme sınırları dışlayıcıdır: 30. gün Aşama 1, 31. gün Aşama 2; 90. gün Aşama 2, 91. gün Aşama 3.
- **Gerekçeler birikir.** Bir kredi birden çok ölçütle üst aşamaya girebilir; hepsi yazılır.
- **Müşteri bazı (`--bulasma`):** Verilen segmentlerde bir kredisi Aşama 3 olan müşterinin diğer kredileri de Aşama 3
  olur. Diğer segmentlerde aşama değişmez, ancak "Müşterinin kredileri farklı aşamalarda" uyarısı verilir.
- **Not ölçeği** en iyiden en kötüye verilir (`"AAA,AA,A,BBB,BB,B,CCC"`). Verilmezse notlar sayısal kabul edilir ve
  küçük sayı iyidir.
- **Önceki dönem:** `--onceki` dosyasında "Aşama" sütunu varsa o kullanılır; yoksa önceki dosya aynı kurallarla
  sınıflandırılır. Aşaması iyileşen her kredi için izleme koşullarını hatırlatan bir uyarı yazılır.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek veri
python main.py --krediler krediler.xlsx --onceki gecen_ay.xlsx --not-olcegi "AAA,AA,A,BBB,BB,B,CCC,CC,C" --bulasma "Ticari,KOBİ"
python main.py --krediler krediler.xlsx --not-esik 2 --to-kat 3
```

| Sütun | Açıklama |
|---|---|
| Kredi No, Risk Bakiyesi | Zorunlu |
| Gecikme Gün | Boşsa 0 |
| Yeniden Yapılandırma, Yakın İzleme, Temerrüt | Evet / boş |
| İlk Not, Güncel Not | Kullandırım tarihindeki ve güncel derecelendirme notu |
| İlk TO, Güncel TO | Temerrüt olasılıkları (aynı birimde) |
| Müşteri No, Müşteri, Segment, Ürün | Kırılım ve müşteri bazı için |
| Aşama | Önceki dönemin aşaması (isteğe bağlı) |

## Çıktı

`cikti/tfrs9_asama.xlsx`:

- `Özet`: aşama bazında adet, bakiye ve pay; segment kırılımı; gerekçe dağılımı.
- `Sınıflandırma`: her kredi için aşama, önceki aşama ve gerekçe; boş "Değerlendirme Notu" sütunu.
- `Aşama Değişenler`, `Geçiş Matrisi`: önceki dönem bilgisi varsa.
- `Uyarılar`: veri sorunları, karışık aşamalı müşteriler, iyileşen krediler.

## Dikkat

- **Sınıflandırma önerisidir.** TFRS 9'da 30 günü aşan gecikme kredi riskinde önemli artışın, 90 günü aşan gecikme
  temerrüdün aksi ispat edilebilir karinesidir. Önemli artışın nasıl belirleneceği (not kademesi, TO katı, nitel
  göstergeler) bankanızın onaylı politikasına bağlıdır; varsayılan değerler yalnız başlangıç noktasıdır.
- **Karşılık hesaplanmaz.** Beklenen kredi zararı (TO × THK × TT, ileriye dönük senaryolar) bu paketin dışındadır.
- **BDDK sınıflandırması:** Kredilerin Sınıflandırılması ve Bunlar İçin Ayrılacak Karşılıklara İlişkin Usul ve
  Esaslar Hakkında Yönetmelik'teki grup sınıflandırması ile TFRS 9 aşamaları ilişkilidir ancak aynı şey değildir.
- **İyileşme:** Üst aşamadan alt aşamaya geçiş için izleme süresi ve ödeme performansı koşulları kodda kontrol
  edilmez; kod yalnız uyarır.
- **Müşteri bazı** kuralının hangi segmentlerde uygulanacağı politikanıza bağlıdır.
- **Kişisel veri:** Müşteri bilgisi içeren çıktıyı yalnız yetkili kişilerle paylaşın.
- **Örnek veri:** Krediler, notlar ve olasılıklar kurgusaldır.

## Testler

- Gecikme eşikleri (30 / 31, 90 / 91), nitel ölçütler, not kademesi ve TO katı.
- Müşteri bazı kuralı ve karışık aşama uyarısı; özet ve segment payları.
- Önceki dönem: geçiş matrisi, kötüleşen / iyileşen, "Aşama" sütunu.
- Okuma hataları ve örnek veri uçtan uca.

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
