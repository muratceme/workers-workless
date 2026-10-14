# Vergi Takvimi Takibi · Kod Bloğu

> Muhasebe › Muhasebe Müdürü · Workers / Workless

Şirketlerin beyan ve bildirim yükümlülüklerinden **yaklaşan son günleri** çıkarır; verilmiş beyanları işaretler, geciken ve
yaklaşanları hatırlatır. Hafta sonu ve resmî tatillerde son günü izleyen iş gününe kaydırır, GİB'in süre uzatmalarını parametreyle
alır. İnternete ve GİB sistemlerine bağlanmaz.

## Nasıl hesaplar?

| Yükümlülük | Son gün (genel kural) |
|---|---|
| KDV | Dönemi izleyen ayın 28'i (beyan ve ödeme) |
| Muhtasar ve prim hizmet (aylık) | Dönemi izleyen ayın 26'sı |
| Damga vergisi · Konaklama vergisi | Dönemi izleyen ayın 26'sı |
| Ba-Bs formları | Dönemi izleyen ayın son günü |
| Geçici vergi | 1., 2. ve 3. çeyreği izleyen 2. ayın 17'si (17 Mayıs, 17 Ağustos, 17 Kasım); 4. dönem için 2022'den beri beyanname verilmez (7338 sayılı Kanun) |
| Kurumlar vergisi | Takvim yılını izleyen 30 Nisan |
| Gelir vergisi | İzleyen yılın 31 Mart'ı (beyan ve 1. taksit); 2. taksit 31 Temmuz |
| e-Defter berat (aylık) | Dönemi izleyen 4. ayın 10'u (gelir) / 14'ü (kurumlar); aralık için GV 10 Nisan, KV 14 Mayıs |

Son gün hafta sonuna ya da resmî tatile denk gelirse izleyen ilk iş günü alınır. Sabit ulusal tatiller ve 2026 dini bayramları
kodda; diğer yılların bayramlarını `--tatiller` ile ekleyin. GİB'in uzatmalarını `--uzatma ESKİ=YENİ` ile girin (ör. 2026 mayısında
26 Mayıs'taki beyanlar 3 Haziran'a uzatılmıştı). Pencere: bugünden `--geri` (60) gün öncesi ile `--ileri` (45) gün sonrası.

Durum: Gecikti · Bugün · Yaklaşıyor (`--uyari` 5 gün) · Planlı · Verildi · Geç verildi. Verilenler dosyasında dönem `2026/09`,
`2026/2. dönem`, `2025` ya da `2025 · 2. taksit` biçiminde yazılır.

**Uyarılar:** son günü geçmiş ve verildiği kaydı olmayan beyan · tanınmayan yükümlülük (Yüksek) · bugün ya da yakında biten beyan
(Orta) · geç verilmiş beyan (Bilgi).

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                          # örnek veri (14.10.2026)
python main.py --yukumlulukler yukumlulukler.xlsx --verilenler verilenler.xlsx
python main.py --yukumlulukler y.xlsx --verilenler v.xlsx --ileri 90 --uyari 7 --tatiller tatiller.csv
python main.py --yukumlulukler y.xlsx --uzatma 26.05.2026=03.06.2026 --uzatma 28.05.2026=05.06.2026
```

## Çıktı

`cikti/vergi_takvimi.xlsx`: `Vergi Takvimi` (son güne göre sıralı; tatil ya da uzatmayla kayan günlerde asıl gün notu), `Uyarılar`.

## Dikkat

- **Süreler kanun, tebliğ ve GİB duyurularıyla değişir.** Kod genel kuralları uygular; özel durumları (üç aylık muhtasar, özel
  hesap dönemi, tasfiye, işi bırakma, ÖTV, yerel asgari tamamlayıcı kurumlar vergisi gibi yeni yükümlülükler) kapsamaz.
  Takvimi her ay GİB vergi takvimi ve sirkülerleriyle karşılaştırın.
- **Ödeme ile beyan günleri** her yükümlülükte aynı olmayabilir; ödeme takibini ayrıca yapın.
- **Ba-Bs** genel kuralına (izleyen ayın son günü) ve bildirim haddine güncel VUK tebliğinden bakın.
- **Örnek veri:** Şirketler, sorumlular ve tarihler kurgusaldır.

## Testler

- KDV, muhtasar, Ba-Bs, geçici vergi (hafta sonu kaydırma dahil), kurumlar ve gelir vergisi, e-defter berat son günleri; tatil
  zinciri (bayram + hafta sonu).
- Verildi / geç verildi / gecikti / bugün / yaklaşıyor durumları, uzatma, dönem anahtarı, okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
