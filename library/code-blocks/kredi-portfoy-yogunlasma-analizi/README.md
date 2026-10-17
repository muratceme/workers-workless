# Kredi Portföy Yoğunlaşma Analizi · Kod Bloğu

> Krediler Tahsis › Kredi Tahsis Müdürü · Gerisi Bende

Kredi portföyünü **sektör, risk grubu ve teminat türüne** göre dağıtır; her boyut için yoğunlaşma endeksini (HHI)
hesaplar ve sizin belirlediğiniz limitleri aşan ya da limite yaklaşan kalemleri işaretler. İnternete bağlanmaz.

## Nasıl hesaplar?

| Kavram | Formül |
|---|---|
| Pay % | Değerin riski ÷ toplam risk × 100 |
| HHI | Payların (yüzde olarak) karelerinin toplamı; 0 – 10.000 arası |
| Eşdeğer sayı | 10.000 ÷ HHI (portföy eşit büyüklükte kaç parçaya bölünmüş gibi davranıyor) |
| İlk N payı | En büyük `--ilk` (5) değerin paylarının toplamı |
| Limit durumu | Pay limiti aşıyorsa "Aşım"; limitin %90'ı ve üstündeyse "Yaklaştı" |
| Risk grubu ÷ özkaynak | Grubun riski ÷ `--ozkaynak` × 100 (özkaynak verilirse) |

Risk grubu boş bırakılan müşteri kendi başına bir grup sayılır.

**Örnek:** Sektör payları %36,67 · %16,67 · %16,67 · %12 · %10 · %8 → HHI = 1.344,44 + 277,78 + 277,78 + 144 + 100 +
64 = **2.208**; eşdeğer sayı 10.000 ÷ 2.208 = **4,53** (altı sektör var ama portföy dört buçuk sektöre yayılmış gibi).

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                            # örnek veri ve örnek limitler
python main.py --krediler krediler.xlsx
python main.py --krediler krediler.xlsx --limitler limitler.xlsx --ilk 10 --ozkaynak 50000000
```

Zorunlu sütunlar: Müşteri, Risk. Limit dosyasında "Değer" sütununa `*` yazarsanız limit o boyuttaki her değere
uygulanır (örneğin hiçbir risk grubu portföyün %25'ini geçmesin).

## Çıktı

`cikti/kredi_portfoy_yogunlasma.xlsx`:

- `Özet`: üç boyut için değer sayısı, HHI, eşdeğer sayı, en büyük pay ve ilk N payı.
- `Sektör`, `Risk Grubu`, `Teminat Türü`: kredi ve müşteri sayısı, risk, pay, kümülatif pay, limit ve durum.
- `Uyarılar`: limit aşımı, limite yaklaşma, özkaynak sınırı aşımı, büyük krediler.

## Dikkat

- **Limitler sizin politikanızdır.** Örnek dosyadaki limitler yalnız gösterim içindir. HHI için "şu değerin üstü
  yoğundur" gibi bir eşik bu pakette yoktur; endeksi dönemden döneme ve kendi risk iştahınıza göre izleyin.
- **Özkaynak karşılaştırması yasal sınır hesabı değildir.** `--ozkaynak-limit` (25) ve `--buyuk-kredi` (10) varsayılanları
  bankacılık mevzuatındaki genel oranlar düşünülerek konmuştur; ancak yasal hesapta kredi türüne göre dikkate alınma
  oranları, istisnalar, dolaylı krediler ve risk grubu tanımı vardır ve kod bunların hiçbirini uygulamaz. Yasal
  raporlama için kullanmayın; oranları güncel mevzuattan doğrulayın.
- **Risk grubu tanımını siz yaparsınız.** Aynı gruba ait müşterileri aynı grup adıyla yazmazsanız yoğunlaşma düşük
  görünür. Sektör kodlarında da aynı yazımı kullanın.
- **Risk tanımı tutarlı olmalı.** Nakdi ve gayri nakdi riski, limit ya da kullanım tutarını karıştırmayın; tutarlar
  tek para biriminde olmalıdır.
- **Kapsam dışı:** vade, para birimi ve coğrafi bölge yoğunlaşması, sektörler arası korelasyon, ekonomik sermaye.
- **Müşteri sırrı:** Rapor müşteri ve risk bilgisi içerir; yalnız yetkili kişilerle paylaşın.
- **Örnek veri:** Müşteriler, gruplar ve tutarlar kurgusaldır.

## Testler

- Pay, kümülatif pay, HHI ve eşdeğer sayı (tek değer, eşit paylar, karışık); ilk N payı.
- Limit durumu sınırları (tam limitte, limitin %90'ında), `*` ile genel limit, özel limitin önceliği.
- Boş risk grubu, özkaynak sınırı ve büyük kredi, portföyde olmayan limit, okuma ve komut satırı.

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
