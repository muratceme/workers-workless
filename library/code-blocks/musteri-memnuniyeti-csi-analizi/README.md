# Müşteri Memnuniyeti (CSI) Analizi · Kod Bloğu

> Satış Sonrası Hizmetler (Servis) › Satış Sonrası Hizmetler Müdürü · Gerisi Bende

Servis memnuniyet anketlerini **danışman, iş türü ve ay** bazında analiz eder: ortalama puan, en üst kutu oranı, NPS ve soru
bazında ortalamalar. Düşük puanlı anketleri (önce geri arama isteyenler) **geri arama listesine** alır. İnternete, marka
anket portalına ve CRM'e bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Ortalama | Genel memnuniyet puanlarının ortalaması (`--olcek` 10 ya da 5) |
| En üst kutu % | 10'luk ölçekte 9–10, 5'lik ölçekte 5 veren anket ÷ anket × 100 |
| Düşük puan | Genel memnuniyet ≤ `--dusuk` (varsayılan 10'lukta 6, 5'likte 3) |
| NPS | Tavsiye sorusunda (0–10) 9–10 verenlerin % − 0–6 verenlerin %; yalnız tavsiye yanıtı olan anketlerden |
| Soru ortalamaları | Başlığı "Soru" ile başlayan her sütun (ör. `Soru: Karşılama`, `Soru 2 - Teslim`) ayrı ortalanır |

**Uyarılar:** düşük puanlı anket (müşteri geri arama istiyorsa Yüksek, istemiyorsa Orta) · ölçek dışı ya da okunamayan puan
(Yüksek) · aynı anket iki kez (Orta) · `--min` (5) anketten az olan danışman (Bilgi; oranları karşılaştırmaya uygun değil).

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                          # örnek veri (temmuz–eylül 2026)
python main.py --anket anketler.xlsx
python main.py --anket anketler.xlsx --olcek 5 --dusuk 2 --min 10
```

Zorunlu sütun: Genel Memnuniyet. Tavsiye sütunu yoksa NPS hesaplanmaz.

## Çıktı

`cikti/csi_analizi.xlsx`: `Danışman`, `İş Türü`, `Ay` (her biri toplam satırıyla; az anketli satırlar mavi), `Geri Arama`
(sonuç notu sütunuyla), `Uyarılar`.

## Dikkat

- **Marka CSI puanı farklı hesaplanabilir.** Markaların kendi anket metodolojisi, ağırlıkları ve hedefleri vardır; bu rapor
  servis içi takip içindir, resmî marka puanının yerine geçmez.
- **NPS** Net Promoter Score'un yaygın tanımıyla hesaplanır (öneren % − kötüleyen %); tavsiye sorusu 0–10 olmalıdır.
- **Az anketle karşılaştırma yanıltır.** Danışman sıralamasını tek başına performans ya da prim kararı için kullanmayın.
- **Kişisel veri içerir.** Anket yanıtları ve yorumlar KVKK kapsamındadır; geri arama listesini yalnız yetkili personelle paylaşın.
- **Örnek veri:** Anketler, danışmanlar ve yorumlar kurgusaldır.

## Testler

- Ortalama, en üst kutu (10'luk ve 5'lik ölçek), düşük puan eşiği, NPS, soru ortalamaları, danışman / iş türü / ay kırılımları.
- Geri arama sıralaması ve uyarılar; soru sütunu tanıma, ölçek dışı puan, okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

## Kaynaklar ve dayanaklar

Kontrol tarihi: 15.10.2026. Mevzuat sık değişir; paketteki süre, oran ve kuralları kullanmadan önce güncel resmî metinle
karşılaştırın. "Karşılaştırılan maddeler" dışındaki atıflar yalnız yol göstericidir.

| Kaynak | Bağlantı | Durum |
|---|---|---|
| 6698 sayılı Kişisel Verilerin Korunması Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.6698.pdf) | Resmî metinle karşılaştırılan maddeler: 13 |

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
