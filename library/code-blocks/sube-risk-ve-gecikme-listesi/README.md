# Şube Risk ve Gecikme Listesi · Kod Bloğu

> Şube Bankacılığı › Şube Müdürü · Workers / Workless

Şube kredi portföyünden **gecikmedeki ve erken uyarı veren müşterileri** çıkarır; gecikme gününe ve tutara göre
sıralar, takip eşiğine kaç gün kaldığını gösterir. Haftalık portföy toplantısının gündem listesidir. İnternete
bağlanmaz.

## Nasıl hesaplar?

| Kavram | Kural |
|---|---|
| Gecikme dilimi | Gecikmesiz · 1–30 · 31–60 · 61–90 · 91+ gün (`--dilimler`) |
| Müşteri özeti | Müşterinin kredilerinin risk ve geciken tutar toplamı; en uzun gecikme günü |
| 1 – Kritik | En uzun gecikme ikinci dilim sınırını (60 gün) aşmış |
| 2 – Yüksek | En uzun gecikme ilk dilim sınırını (30 gün) aşmış **ya da** geciken tutar `--buyuk` (250.000) ve üstünde |
| 3 – Orta | Gecikme var |
| 4 – Erken uyarı | Gecikme yok ama "Erken Uyarı" sütununda not var |
| Sıra | Öncelik → gecikme günü (uzundan kısaya) → geciken tutar (büyükten küçüğe) |
| Takip eşiğine kalan | `--takip` (90) − en uzun gecikme günü |
| Gecikme oranı % | Gecikmedeki kredilerin risk bakiyesi ÷ toplam risk bakiyesi × 100 |

**Örnek:** Müşterinin iki kredisinden biri 72 gün gecikmede → en uzun gecikme 72 gün: **1 – Kritik**; 90 günlük
eşiğe **18 gün** kaldı. Risk toplamı iki kredinin toplamıdır (1.600.000).

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                            # örnek veri
python main.py --krediler krediler.xlsx
python main.py --krediler krediler.xlsx --dilimler "15;30;60;90" --takip 90 --buyuk 500000
```

Zorunlu sütunlar: Müşteri No, Risk Bakiyesi. Gecikmesi ve erken uyarı notu olmayan müşteriler öncelik listesine
girmez, yalnız toplamlarda yer alır.

## Çıktı

`cikti/sube_risk_gecikme_listesi.xlsx`:

- `Öncelik Listesi`: müşteri bazında öncelik, risk, geciken tutar, en uzun gecikme, eşiğe kalan gün, aksiyon sütunu.
- `Özet`: gecikme dilimlerine göre kredi sayısı ve risk, gecikme oranı, portföy yöneticisi kırılımı.
- `Krediler`: bütün krediler, en uzun gecikmeden başlayarak.
- `Uyarılar`: takip eşiğini aşan ve yaklaşan müşteriler, erken uyarılar, tutarsız veri.

## Dikkat

- **Öncelik kuralı bir çalışma varsayımıdır.** Dilimler, takip eşiği ve "büyük tutar" sınırı varsayılandır; bankanızın
  izleme politikasına göre ayarlayın. Bu rapor kredi sınıflandırması, karşılık hesabı ya da yasal takip kararı
  değildir; bunlar ilgili düzenlemelere ve kredi izleme biriminizin değerlendirmesine göre yapılır.
- **Gecikme günü sizin verinizdir.** Kod günü tarihten hesaplamaz; sistemden aldığınız gün sayısını kullanır. Rapor
  tarihi ile veri tarihi aynı olmalıdır.
- **Erken uyarı serbest metindir.** Kod notun içeriğini yorumlamaz; hücre doluysa müşteriyi listeye alır. Hangi
  sinyallerin izleneceğini (çek iadesi, limit aşımı, teminat değer kaybı, olumsuz istihbarat) siz belirlersiniz.
- **Müşteri toplamı tek para biriminde olmalı.** Dövizli kredileri TL karşılığıyla girin.
- **Kapsam dışı:** grup (risk grubu) bazında birleştirme, teminat durumu, tahsilat tahmini, yapılandırma önerisi.
- **Müşteri sırrı:** Rapor müşteri risk bilgisi içerir; yalnız yetkili kişilerle paylaşın.
- **Örnek veri:** Müşteriler, krediler ve tutarlar kurgusaldır.

## Testler

- Dilim sınırları (0, 1, 30, 31, 60, 61, 90, 91 gün), öncelik sınırları, büyük tutar kuralı, erken uyarı.
- Çok kredili müşteride toplama ve en uzun gecikme, takip eşiği (tam eşikte ve bir gün sonrası), özel dilimler.
- Gecikme oranı, portföy yöneticisi özeti, tutarsız veri uyarıları, okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
