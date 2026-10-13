# Ödeme Planı Hesaplama · Kod Bloğu

> Bankacılık Operasyonları › Kredi Operasyon Uzmanı · Workers / Workless

Kredi tutarı, aylık faiz, vade ve ödeme tipinden **taksit taksit ödeme planı** çıkarır: her taksitte anapara, faiz,
KKDF, BSMV ve kalan anapara. Eşit taksitli, eşit anaparalı ve balonlu planı destekler. İnternete bağlanmaz.

## Nasıl hesaplar?

| Kavram | Formül |
|---|---|
| Dönem faizi | Kalan anapara × aylık faiz oranı |
| KKDF, BSMV | Dönem faizi × KKDF oranı; dönem faizi × BSMV oranı |
| Vergili oran (rv) | Aylık faiz oranı × (1 + KKDF oranı + BSMV oranı) |
| Eşit taksit | Anapara × rv ÷ (1 − (1 + rv)^−vade); taksitteki anapara = taksit − faiz − KKDF − BSMV |
| Eşit anapara | Her dönem (anapara ÷ vade) + o dönemin faizi ve vergileri; taksit giderek azalır |
| Balonlu | (Anapara − balon ÷ (1 + rv)^vade) × rv ÷ (1 − (1 + rv)^−vade) eşit taksit; son taksitte ayrıca balon |
| Son taksit | Kalan anaparanın tamamı + faizi ve vergileri (kuruş farkı son taksitte kapanır) |

**Örnek:** 1.000 kredi, aylık %10 faiz, 2 ay, KKDF %10, BSMV %10 → rv = %12; taksit = 1.000 × 0,12 ÷ (1 − 1 ÷ 1,2544) =
**591,70**. İlk taksit: faiz 100, KKDF 10, BSMV 10, anapara 471,70; kalan 528,30. İkinci taksit: faiz 52,83, KKDF
5,28, BSMV 5,28, anapara 528,30 → 591,69.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                            # örnek dosyadaki üç kredi
python main.py --tutar 100000 --faiz 3,5 --vade 12 --bsmv 15 --kkdf 15 --ilk-taksit 15.11.2026
python main.py --tutar 500000 --faiz 3 --vade 24 --tip esit-anapara --bsmv 5
python main.py --tutar 800000 --faiz 2,8 --vade 36 --tip balon --balon 300000 --bsmv 15 --kkdf 15
python main.py --krediler krediler.xlsx
```

Faiz **aylık** ve yüzde olarak girilir; virgül ondalık ayırıcıdır. Dosyada zorunlu sütunlar: Tutar, Aylık Faiz, Vade.
Ödeme tipi boşsa eşit taksit, BSMV ve KKDF boşsa sıfır alınır.

## Çıktı

`cikti/odeme_plani.xlsx`:

- `Özet`: her kredi için ilk ve son taksit, toplam faiz, KKDF, BSMV, masraf, toplam maliyet ve toplam geri ödeme.
- `Plan 1`, `Plan 2`, …: taksit no, tarih, taksit tutarı, anapara, faiz, KKDF, BSMV, kalan anapara ve toplam satırı.
- `Uyarılar`.

## Dikkat

- **BSMV ve KKDF oranlarını siz girersiniz.** Oranlar kredi türüne (tüketici, ticari, konut), kullanıcıya ve döneme
  göre değişir, bazı kredilerde sıfırdır. Kod bir oran varsaymaz; örnek dosyadaki oranlar yalnız gösterim içindir.
  Güncel oranları mevzuattan ya da bankanızın ürün parametrelerinden alın.
- **Dönemler eşit ve aylık kabul edilir.** Gerçek planlarda ilk dönem gün sayısına göre faiz, taksit tarihinin
  tatile gelmesi hâlinde kaydırma, ödemesiz dönem ve ara ödeme olabilir; bankacılık sisteminizin planıyla kuruş ya da
  ilk taksit farkı çıkması bu yüzden normaldir. Müşteriye verilecek plan bankacılık sisteminden alınmalıdır.
- **Yıllık maliyet oranı hesaplanmaz.** Tüketici kredilerinde sözleşme öncesi bilgilendirmede yer alması gereken
  efektif yıllık faiz ve yıllık maliyet oranı bu paketin kapsamında değildir.
- **Faiz sabittir.** Değişken faizli ve endeksli krediler, erken ödeme, gecikme faizi ve yapılandırma hesaplanmaz.
- **Masraf yalnız toplama eklenir**; taksitlere dağıtılmaz ve üzerinden vergi hesaplanmaz.
- **Örnek veri:** Krediler, faiz ve vergi oranları kurgusaldır; piyasa ya da yasal oran değildir.

## Testler

- Eşit taksit: elle doğrulanmış iki taksitli plan (vergisiz ve vergili), sıfır faiz, son taksitte kuruş kapanışı.
- Eşit anapara ve balonlu plan; anapara toplamının kredi tutarına eşitliği, son kalan anaparanın sıfır olması.
- Taksit tarihleri (ay sonu), geçersiz girdiler, toplu dosya, tek kredi komut satırı, okuma.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
