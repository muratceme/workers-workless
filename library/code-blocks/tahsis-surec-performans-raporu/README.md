# Tahsis Süreç Performans Raporu · Kod Bloğu

> Krediler Tahsis › Kredi Tahsis Müdürü · Workers / Workless

Kredi tekliflerinin **ne kadar sürede sonuçlandığını, onay / ret / iade oranlarını ve iade nedenlerini** şube,
segment ve analist bazında raporlar; süre hedefini aşmış bekleyen teklifleri listeler. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Süre (gün) | Karar tarihi − giriş tarihi; bekleyen teklifte durum tarihi − giriş tarihi (takvim günü) |
| Sonuçlanan | Onay + Ret + İade (Bekliyor hariç) |
| Onay %, Ret %, İade % | İlgili adet ÷ sonuçlanan × 100 |
| Onay % (tutar) | Onaylanan tutar ÷ sonuçlanan tutar × 100 |
| Ortalama ve medyan süre | Sonuçlanan tekliflerin süreleri üzerinden |
| Süre hedefi içinde % | Süresi `--sla` (5) gün ve altında olan sonuçlanan teklif ÷ sonuçlanan × 100 |
| İade oranı uyarısı | En az 3 sonuçlanan teklifi olan şubede iade oranı `--iade-esik` (%30) ve üstünde |

**Örnek:** Sonuçlanan 10 teklifin süreleri 1, 2, 2, 2, 3, 4, 4, 7, 10, 12 gün → ortalama **4,70**, medyan
**3,50** gün; 5 gün ve altında 7 teklif → hedef içinde **%70**. Ortalamayı iki uzun teklif yukarı çekiyor; medyan
tipik teklifin süresini daha iyi gösterir.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                            # örnek veri, durum tarihi 12.10.2026
python main.py --teklifler teklifler.xlsx
python main.py --teklifler teklifler.xlsx --sla 3 --iade-esik 25 --tarih 31.10.2026
```

Zorunlu sütunlar: Teklif No, Giriş Tarihi. Karar boşsa teklif bekliyor sayılır. Sonuçlanmış teklifte karar tarihi
zorunludur.

## Çıktı

`cikti/tahsis_surec_performans.xlsx`:

- `Özet`: adetler, oranlar, ortalama ve medyan süre, hedef içinde oranı, bekleyenler.
- `Şubeler`, `Segmentler`, `Analistler`: aynı göstergeler kırılımlı.
- `İade Nedenleri`: iade edilen tekliflerin neden dağılımı.
- `Bekleyenler`: en uzun bekleyenden başlayarak açık teklifler.
- `Teklifler` ve `Uyarılar`.

## Dikkat

- **Süre takvim günüdür.** Hafta sonu ve tatiller düşülmez; cuma girip pazartesi sonuçlanan teklif 3 gün görünür.
  Hedefinizi iş günü olarak tanımladıysanız `--sla` değerini buna göre büyütün.
- **Süre, tahsisin elinde geçen süre değildir.** Şubeden ek bilgi beklenen günler ayrılmaz. İade edilip yeniden
  gelen teklifi yeni numarayla girerseniz ayrı teklif sayılır; aynı numarayla girerseniz ilk giriş tarihi korunmalıdır.
- **Onay oranı bir kalite ölçüsü değildir.** Yüksek onay oranı iyi ön eleme de, gevşek değerlendirme de olabilir;
  analistleri bu oranla karşılaştırırken baktıkları tekliflerin segmentini ve büyüklüğünü hesaba katın.
- **Az teklifte oran yanıltır.** Kırılımları teklif sayısıyla birlikte okuyun.
- **Süre hedefi ve iade eşiği yönetim tercihidir**; bir sektör standardı değildir.
- **Kapsam dışı:** onay koşulları ve limit değişikliği, yetki seviyesine göre süre, onaylanan kredilerin sonraki
  performansı.
- **Örnek veri:** Şubeler, kişiler ve teklifler kurgusaldır.

## Testler

- Süre hesabı, ortalama ve medyan (tek ve çift sayıda teklif), hedef sınırı (tam hedefte ve bir gün üstü).
- Oranlar (adet ve tutar), bekleyen tekliflerin dışarıda kalması, bekleyenlerde iki kademeli uyarı.
- Şube iade oranı eşiği ve asgari teklif koşulu, iade nedeni dağılımı, okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
