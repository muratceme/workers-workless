# Misafir Mesajlarına Cevap Taslağı · AI Agent

> Ön Büro › Resepsiyonist · Workers / Workless

Misafirlerin e-posta ve mesajlarına, **otelin kendi bilgilerine dayanarak** ve misafirin dilinde cevap taslağı hazırlar.
Bilgi dosyasında cevabı olmayan soruları ve iptal, iade, şikâyet gibi insan onayı gereken mesajları ayrıca işaretler.
Taslaklar gönderilmez.

## Nasıl çalışır?

1. **Hazırlık (kod):** Otel bilgileri numaralanır (B1, B2 …). Misafir adı ve rezervasyon numarası mesajdan çıkarılır;
   telefon ve e-posta maskelenir.
2. **Taslak (model):** Her mesaj için misafirin dilinde cevap yazılır. Hitapta `{AD}` yer tutucusu kullanılır, kod gerçek
   adı doldurur. Kullanılan bilgi numaraları, cevaplanamayan sorular ve insan onayı gerekip gerekmediği belirtilir.
3. **Denetim (kod):** Aşağıdakilerden biri varsa yalnız o mesaj modele geri gönderilir (varsayılan en fazla 2 tur):
   - Bilgi dosyasında olmayan bilgi numarası
   - Cevapta, kullanılan bilgilerde ve misafirin mesajında olmayan sayı, saat ya da ücret
   - Bilgiye dayanmayan "ücretsiz / garanti" ifadesi
   - "Onaylandı / confirmed" gibi kesin onay
   - İptal / iade / şikâyet içeren mesaj ya da cevaplanamayan soru insan onayına işaretlenmemiş
   - `{AD}` yer tutucusu yok

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                    # örnek mesajlar + otel bilgileri
python agent.py --girdi mesajlar.xlsx --bilgi otel_bilgileri.txt
```

Otel bilgileri dosyasının her satırına bir bilgi yazın: giriş / çıkış saatleri, kahvaltı, transfer ücreti, otopark, evcil
hayvan, havuz, iptal kuralları, iletişim. Satırları "B1:", "B2:" diye numaralayabilirsiniz; numarasız satırlar sırayla
numaralanır.

## Çıktı

- `cikti/misafir_cevaplari.xlsx`:
  - `Cevap Taslakları`: insan onayı gerekenler en üstte; "Onay" sütunuyla.
  - `Otel Bilgileri`: her bilginin hangi mesajlarda kullanıldığı.

## Dikkat

- **Taslaklar gönderilmez.** Gönderilmeden önce okunup onaylanmalıdır. İnsan onayı gerekenler ilgili yetkiliye iletilmelidir.
- **Bilgi dosyasını güncel tutun.** Ücret ve saatler yalnız bu dosyadan alınır; dosya eskiyse cevaplar da eski olur.
- **Çeviri kontrolü.** Bilmediğiniz dillerdeki cevapları gönderiminden önce bir çeviri aracıyla kontrol etmeniz önerilir.
- **Veri gönderimi.** Mesaj metni ve otel bilgileri seçtiğiniz yapay zekâ sağlayıcısına gönderilir; misafir adı ve rezervasyon
  numarası çıkarılır, telefon ve e-posta maskelenir.
- **Örnek veri:** Otel bilgileri ve misafirler kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk taslakta bilerek transfer ücretini yanlış yazar, iade edilemez rezervasyonun iade
talebini insan onayına işaretlemez ve geç çıkış talebini "onaylandı" diye cevaplar. Kodun bunları yakalayıp yalnız bu
mesajları yeniden yazdırdığı ve misafir adlarını modele göndermediği doğrulanır.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. API kullanım ücretleri ve gönderilen cevapların doğruluğu kullanıcının sorumluluğundadır. Bkz.
`LICENSE` ve `SORUMLULUK_REDDI.md`.
