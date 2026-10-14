# Kurum Bazlı Gelir ve Kesinti Raporu · Kod Bloğu

> Anlaşmalı Kurumlar ve Faturalama › Anlaşmalı Kurumlar Müdürü · Workers / Workless

Kurum faturalarından **SGK, özel sigorta şirketleri ve diğer anlaşmalı kurumlar** bazında fatura, kesinti, tahsilat,
açık bakiye, **kesinti oranı**, **ortalama tahsil süresi** ve alacak yaşlandırması çıkarır. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Kesinti oranı % | Kesinti tutarı ÷ fatura tutarı × 100 |
| Net alacak | Fatura tutarı − kesinti |
| Açık bakiye | Net alacak − tahsil edilen |
| Tahsil oranı % | Tahsil edilen ÷ net alacak × 100 |
| Ortalama tahsil süresi | Σ(tahsil edilen × (tahsil tarihi − fatura tarihi)) ÷ Σ tahsil edilen (tutar ağırlıklı, gün) |
| Yaşlandırma | Açık bakiye, fatura tarihinden rapor tarihine geçen güne göre 0-30 / 31-60 / 61-90 / 91-180 / 180+ |

Kurum türü sütunu yoksa kurum adından tanınır: "SGK" → SGK, "sigorta" / ÖSS / TSS → Özel sigorta, diğerleri Diğer.

**Uyarılar:** kesinti oranı `--kesinti` (varsayılan %5) üstünde · ortalama tahsil süresi `--sure` (varsayılan 90) günü
aşan · 180 günden eski açık bakiye · kesinti faturadan büyük · tahsilat net alacaktan fazla.

**Örnek:** Sigorta Şirketi C'ye 1.636.900 fatura, 139.782,35 kesinti → kesinti oranı **%8,54** (eşik %5 → uyarı);
tahsil edilen faturalar ortalama **112,75 günde** ödenmiş.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                         # örnek veri (30.09.2026)
python main.py --faturalar kurum_faturalari.xlsx --tarih 30.09.2026
python main.py --faturalar kurum_faturalari.xlsx --kesinti 3 --sure 60
```

Zorunlu sütunlar: Kurum, Fatura Tarihi, Fatura Tutarı. Kesinti ve tahsilat boşsa sıfır sayılır. Kısmi tahsilatta
tahsil edilen toplam tutarı ve son tahsil tarihini yazın.

## Çıktı

`cikti/kurum_gelir_kesinti_raporu.xlsx`:

- `Kurumlar`: fatura tutarına göre sıralı; göstergeler, yaşlandırma ve toplam; eşiği aşanlar kırmızı.
- `Kurum Türleri`: SGK / Özel sigorta / diğer toplamları.
- `Açık Faturalar`: en eskiden yeniye; 90 ve 180 günü geçenler renkli.
- `Uyarılar`.

## Dikkat

- **Kesinti nedenlerini göstermez.** Kesinti oranı yüksek kurumda nedenleri (SUT uyumu, provizyon, eksik belge,
  anlaşma fiyatı) ayrıca inceleyin; itiraz edilip geri alınan kesintileri kesinti sütunundan düşerek yazın.
- **Tahsil süresi fatura tarihine göredir.** SGK'da fatura dönem sonu, teslim ve inceleme takvimine, özel sigortada
  sözleşmedeki ödeme vadesine bağlıdır; karşılaştırmayı aynı tür kurumlar arasında yapın.
- **Muhasebe kayıtlarıyla mutabakat** ayrıca yapılmalıdır; rapor muhasebe kaydı ya da cari hesap ekstresi yerine geçmez.
- **Eşikler politika varsayılanıdır;** sözleşmelerinize ve geçmiş ortalamalarınıza göre değiştirin.
- **Örnek veri:** Kurum adları, fatura numaraları ve tutarlar kurgusaldır.

## Testler

- Kesinti ve tahsil oranı, net alacak, açık bakiye, tutar ağırlıklı tahsil süresi, yaşlandırma dilim sınırları.
- Kurum türü tanıma, kurum türü toplamları, açık fatura listesi, uyarılar, hatalı satırlar, okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
