# Temsilci İhtiyacı Hesaplama (Erlang C) · Kod Bloğu

> Müşteri Hizmetleri ve Çağrı Merkezi › Çağrı Merkezi Yöneticisi · Gerisi Bende

Aralık bazında (15/30/60 dk) çağrı tahmininden **gereken temsilci sayısını** Erlang C modeliyle hesaplar.
Hedef servis seviyesi (ör. çağrıların %80'i 20 saniyede) ve en yüksek doluluk oranını birlikte sağlayan en
küçük temsilci sayısını bulur; izin, eğitim, mola ve devamsızlık kayıpları (shrinkage) için planlanacak kişi
sayısını verir. İnternete bağlanmaz.

## Hesap

| | |
|---|---|
| Trafik (Erlang) | Çağrı × AHT ÷ aralık süresi |
| Bekleme olasılığı | Erlang C (Erlang B özyinelemesiyle, büyük trafiklerde de kararlı) |
| Servis seviyesi | 1 − P(bekleme) × e^(−(N − trafik) × hedef süre ÷ AHT) |
| Ortalama cevap süresi (ASA) | P(bekleme) × AHT ÷ (N − trafik) |
| Doluluk | Trafik ÷ N |
| Planlanan kişi | ⌈N ÷ (1 − kayıp oranı)⌉ |

Test, yayımlanmış çözümlü Erlang C örneğiyle birebir doğrulanır: 100 çağrı / 30 dk, AHT 180 sn, 14 temsilci →
P(bekleme) 0,1741 · servis seviyesi %88,8 · ASA 7,84 sn · doluluk %71,4.

**Model varsayımları:** müşteriler kapatmaz (terk yok), çağrılar Poisson dağılımlı gelir, görüşme süreleri
üsteldir. Terk oranı yüksek çağrı merkezlerinde gereken temsilci sayısı bir miktar fazla tahmin edilir.

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py --cagri 100 --aralik 30 --aht 180                    # tek aralık
python main.py --girdi cagri_tahmini.xlsx --aht 180 --sl 80 --hedef-sn 20 --maks-doluluk 85 --kayip 30
python main.py                                                      # örnek günlük tahminle dener
```

Girdi sütunları: **Aralık** (veya Saat) ve **Çağrı Sayısı** (zorunlu); Gün, AHT (sn) (isteğe bağlı; satır
bazında AHT verilirse o kullanılır).

## Testler

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
