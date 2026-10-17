# Teminat Takip Listesi · Kod Bloğu

> Şube Bankacılığı › Ticari Portföy Yöneticisi · Gerisi Bende

Kredi teminatlarının **geçerlilik, ekspertiz yenileme ve sigorta tarihlerini** izler; müşteri bazında teminatın
krediyi ne ölçüde karşıladığını ve teminat açığını hesaplar. İnternete bağlanmaz.

## Nasıl hesaplar?

| Kavram | Kural |
|---|---|
| Dikkate alınan değer | Teminat değeri × dikkate alma oranı (oran dosyasından; dosya yoksa %100) |
| Süresi dolmuş teminat | Geçerlilik tarihi geçmişse dikkate alınan değer sıfırdır |
| Karşılama % | Müşterinin dikkate alınan teminat toplamı ÷ kredi riski × 100 |
| Teminat açığı | Kredi riski − dikkate alınan teminat toplamı (eksiye düşmez) |
| Ekspertiz yenileme | Ekspertiz tarihi + `--ekspertiz-ay` (12) ay; geçmişse "yenilenmeli" |
| Geçerlilik ve sigorta | Bitiş geçmişse "dolmuş"; `--gun` (30) gün veya az kaldıysa "yaklaşıyor" |

**Örnek:** 5.000.000 değerinde ipotek × %75 = 3.750.000; 400.000 müşteri çeki × %80 = 320.000 → dikkate alınan
**4.070.000**. Kredi riski 3.500.000 → karşılama **%116,29**, açık yok.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                            # örnek veri, durum tarihi 12.10.2026
python main.py --teminatlar teminatlar.xlsx
python main.py --teminatlar teminatlar.xlsx --riskler riskler.xlsx --oranlar oranlar.xlsx --ekspertiz-ay 24 --gun 45
```

Zorunlu sütunlar: Müşteri No, Teminat Türü, Teminat Değeri. Risk dosyası verilmezse karşılama oranı hesaplanmaz;
oran dosyası verilmezse bütün teminatlar tam değeriyle alınır.

## Çıktı

`cikti/teminat_takip_listesi.xlsx`:

- `Teminatlar`: durum, dikkate alınan değer, geçerlilik, ekspertiz yenileme, sigorta bitişi ve yapılacak iş.
- `Müşteriler`: teminat değeri, dikkate alınan, kredi riski, karşılama oranı ve teminat açığı (açığı büyük olan üstte).
- `Teminat Türleri`: tür bazında adet, değer ve dikkate alınan tutar.
- `Uyarılar`.

## Dikkat

- **Dikkate alma oranları bankanızın kredi politikasıdır.** Örnek dosyadaki oranlar yalnız gösterim içindir; bir
  öneri ya da düzenleyici oran değildir. Yasal teminat sınıflandırması ve karşılık hesabında kullanılan dikkate
  alınma oranları bu paketin konusu değildir.
- **Ekspertiz yenileme süresi de politikadır.** 12 ay varsayılandır; teminat türüne ve kredi tutarına göre farklı
  süre uyguluyorsanız türleri ayrı dosyalarda çalıştırın.
- **Teminat değeri sizin girdinizdir.** Kod ipoteğin derecesini, üzerindeki diğer takyidatları, çekin keşidecisini
  ya da kefilin mali gücünü değerlendirmez. Aynı teminat birden çok krediyi karşılıyorsa bir kez yazın.
- **Süresi dolan teminat sıfırlanır, sigortası biten sıfırlanmaz.** Sigortası bitmiş teminat uyarılır ama değeri
  korunur; nasıl değerlendirileceğine siz karar verirsiniz.
- **Tutarlar tek para biriminde olmalı.** Dövizli teminat ve riskleri TL karşılığıyla girin; kur farkından doğan
  teminat açığı ayrıca izlenmelidir.
- **Müşteri sırrı:** Rapor müşteri risk ve teminat bilgisi içerir; yalnız yetkili kişilerle paylaşın.
- **Örnek veri:** Müşteriler, teminatlar, oranlar ve tutarlar kurgusaldır.

## Testler

- Geçerlilik sınırları (dün, bugün, 30 ve 31 gün), süresi dolan teminatın sıfırlanması, sigorta sınırları.
- Ekspertiz yenileme (ay ekleme, geçmiş, yaklaşan, tarihi olmayan ipotek), oranlı ve oransız değer, tanımsız tür.
- Müşteri bazında karşılama ve açık, teminatsız risk, risk dosyasız çalışma, tür özeti, okuma ve komut satırı.

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
