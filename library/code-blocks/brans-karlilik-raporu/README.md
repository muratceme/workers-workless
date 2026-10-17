# Branş Kârlılık Raporu · Kod Bloğu

> Teknik (Risk Kabul / Underwriting) › Teknik Müdür · Gerisi Bende

Branş bazında yazılan prim, karşılık değişimleri, hasar ve giderlerden **kazanılmış prim, gerçekleşen hasar, hasar
prim oranı, bileşik oran ve teknik sonucu** hesaplar. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Kazanılmış prim | Yazılan prim + devreden kazanılmamış primler karşılığı (KPK) − dönem sonu KPK |
| Gerçekleşen hasar | Ödenen hasar + dönem sonu muallak − devreden muallak |
| Hasar prim oranı % | Gerçekleşen hasar ÷ kazanılmış prim × 100 |
| Komisyon oranı % | Komisyon gideri ÷ kazanılmış prim × 100 |
| Gider oranı % | Faaliyet gideri ÷ kazanılmış prim × 100 |
| Bileşik oran % | (Gerçekleşen hasar + komisyon + faaliyet gideri) ÷ kazanılmış prim × 100 |
| Teknik sonuç | Kazanılmış prim − gerçekleşen hasar − komisyon − faaliyet gideri |
| Teknik marj % | Teknik sonuç ÷ kazanılmış prim × 100 |

Durum: bileşik oran `--bilesik` (100) üstündeyse "Teknik zarar"; değilse ve hasar prim oranı `--hp` (80) üstündeyse
"Hasar oranı yüksek". Aynı branş birden çok satırda geçiyorsa (ör. her ay bir satır) tutarlar toplanır; `Dönem`
sütunu doluysa dönem kırılımı ayrı sayfada verilir.

**Örnek:** Yazılan prim 48.000.000, devreden KPK 20.000.000, dönem sonu KPK 23.000.000 → kazanılmış prim
**45.000.000**. Ödenen hasar 30.000.000, muallak 6.000.000'dan 7.500.000'a çıktı → gerçekleşen hasar **31.500.000**,
hasar prim oranı **%70**. Komisyon %15 ve gider %10 ile bileşik oran **%95**, teknik sonuç **2.250.000**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek veri
python main.py --veri brans_verisi.xlsx
python main.py --veri brans_verisi.xlsx --hp 75 --bilesik 98
```

Zorunlu sütunlar: Branş, Yazılan Prim. Boş bırakılan tutar sütunları sıfır sayılır.

## Çıktı

`cikti/brans_karlilik_raporu.xlsx`:

- `Branşlar`: prim payı, kazanılmış prim, gerçekleşen hasar, giderler, teknik sonuç ve oranlar; son satır toplam.
- `Hesap Ayrıntısı`: kazanılmış prim ve gerçekleşen hasarın hangi tutarlardan geldiği.
- `Dönemler` (Dönem sütunu doluysa) ve `Uyarılar`.

## Dikkat

- **Bütün tutarlar aynı esasta olmalı.** Hepsini brüt ya da hepsini reasürans payı düşülmüş (net) verin; kod
  reasürans payını kendisi ayırmaz. Brüt prim ile net hasarı karıştırırsanız oranlar anlamsız olur.
- **Oranların paydası kazanılmış primdir.** Bazı raporlarda komisyon ve gider oranı yazılan prime bölünür; kendi
  raporlarınızla karşılaştırırken bu farkı göz önünde tutun.
- **Mali tablodaki teknik kâr değildir.** Yatırım gelirlerinin teknik bölüme aktarılan kısmı, devam eden riskler
  karşılığı, dengeleme karşılığı, rücu ve sovtaj gelirleri (ödenen hasardan düşmediyseniz) ve diğer teknik gelir-gider
  kalemleri hesapta yoktur.
- **Muallak tutarını siz verirsiniz.** Gerçekleşmiş ama rapor edilmemiş hasar (IBNR) dahil mi değil mi, kod bilmez;
  her iki dönem sonu için aynı tanımı kullanın. Karşılığın yeterliliğini kod değerlendirmez.
- **Dönem tutarlılığı sizdedir.** Bir satırdaki devreden karşılık, önceki satırın dönem sonu karşılığına eşit mi
  diye kontrol edilmez.
- **Eşikler politikanızdır.** %80 ve %100 varsayılanı genel kıyas değerleridir; branşa göre kabul edilebilir hasar
  oranı çok değişir.
- **Örnek veri:** Branşlar ve tutarlar kurgusaldır.

## Testler

- Kazanılmış prim, gerçekleşen hasar, oranlar, teknik sonuç ve toplam satırı (toplam oranları tutarlardan yeniden hesaplanır).
- Eşik sınırları, kazanılmış primin sıfır olması, eksi gerçekleşen hasar.
- Aynı branşın birden çok satırı, dönem kırılımı, okuma ve komut satırı.

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
