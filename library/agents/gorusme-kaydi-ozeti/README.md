# Görüşme Kaydı Özeti · AI Agent

> Müşteri Hizmetleri ve Çağrı Merkezi › Müşteri Temsilcisi · Workers / Workless

Temsilcilerin serbest yazdığı **görüşme notlarını** CRM'e girilecek **standart kayda** çevirir: kategori,
talep, yapılan işlem, durum, bekleyen aksiyonlar, **müşteriye verilen sözler**, duygu ve risk. Ayrıca
**son tarihe göre sıralı takip listesi** üretir; böylece "yarın ararız" denip unutulan müşteri kalmaz.

## Nasıl çalışır?

1. **Kod:** Aynı müşterinin tekrar eden görüşmeleri bağlanır. Müşteri adları takma adla (K1, K2) değiştirilir;
   telefon, e-posta, IBAN ve TCKN maskelenir.
2. **Model:** Her not standart kayda çevrilir. Süre ifadeleri notta geçtiği gibi alınır ("yarın", "3 iş günü
   içinde", "bu hafta içinde").
3. **Kod, süre ifadelerini görüşme tarihine göre gerçek tarihe çevirir:**
   - **N iş günü:** Hafta sonu ve sabit tarihli resmî tatiller (1 Ocak, 23 Nisan, 1 Mayıs, 19 Mayıs, 15 Temmuz,
     30 Ağustos, 29 Ekim) atlanır. Dini bayramlar her yıl değiştiği için `--tatiller` dosyasıyla verilir.
   - **Bu hafta:** O haftanın son iş günü. **Haftaya:** Sonraki haftanın son iş günü.
   - **Diğerleri:** "N gün", "N hafta", "ay sonu" ve açık tarih.
4. **Kontroller:** Takip listesinde gecikmiş ve tarihsiz işler işaretlenir. Bekleyen ya da eskalasyondaki bir
   kayıtta aksiyon, sorumlu veya tarih yoksa uyarı verilir.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                  # örnek 5 görüşme
python agent.py --girdi gorusmeler.xlsx --tatiller dini_bayramlar_2026.txt
```

**Tatil dosyası:** Bir satıra bir tarih (GG.AA.YYYY) yazılır, örneğin dini bayram günleri ve arife yarım günleri.

## Çıktı

`Takip Listesi` (son tarih, durum, iş/söz, sorumlu, "Tamamlandı") · `Görüşme Kayıtları` · `Bilgi`

## Testler

Testler gerçek API çağırmaz. Tarih hesabı 29 Ekim tatiliyle test edilir: 27.10.2026 + 3 iş günü = 02.11.2026.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. API kullanım ücretleri ve modele gönderilen verinin hukuka uygunluğu
kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
