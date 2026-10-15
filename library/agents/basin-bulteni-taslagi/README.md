# Basın Bülteni Taslağı · AI Agent

> Pazarlama › Pazarlama Müdürü · Workers / Workless

Onaylı bilgi notundan **kurumsal dilde basın bülteni taslağı** hazırlar. Rakamlar, alıntılar ve iddialar bilgi notuyla
karşılaştırılır: **alıntı uydurulamaz, abartılı iddia eklenemez.**

## Nasıl çalışır?

1. **Hazırlık (kod):** Bilgi notundan şirket, şehir, tarih, basın iletişimi ve "şirket hakkında" metni okunur.
   Bilgiler B1, B2 …, alıntılar konuşmacı adıyla A1, A2 … diye numaralanır.
2. **Taslak (model):** Başlık, spot ve paragraflar yazılır; her paragraf dayandığı bilgi ve alıntıları gösterir.
3. **Denetim (kod):** Aşağıdakilerden biri varsa taslak modele geri gönderilir (varsayılan en fazla 2 tur):
   - Bilgi notunda (paragrafta: dayandığı bilgilerde) olmayan bir sayı var
   - Tırnak içindeki söz verilen alıntılarda aynen geçmiyor, alıntı dayanakta yok ya da konuşmacının adı yazılmamış
   - "İlk", "lider", "en büyük", "rekor", "benzersiz" gibi iddia bilgi notunda geçmiyor
   - Başlık ya da spot karakter sınırını aşıyor (varsayılan 100 / 300)
4. **Bülten (kod):** Yer-tarih satırı ("İzmir, 15.10.2026 –"), "Şirket Hakkında" ve "Basın İletişimi" bölümleri aynen
   eklenir.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                       # örnek (kurgusal) santral açılışı
python agent.py --girdi bilgi_notu.docx --baslik-azami 90
```

Bilgi notu örneği:

```
Şirket: Örnek Enerji A.Ş.
Şehir: İzmir
Tarih: 15.10.2026
Basın İletişimi: Kurumsal İletişim · basin@… · 0232 …
Şirket Hakkında: …

Bilgiler:
- Kurulu gücü 45 MW olan güneş enerjisi santrali devreye alındı.
- Yatırım tutarı 38 milyon ABD dolarıdır.

Alıntılar:
Ayşe Kara, Genel Müdür: "…"
```

## Çıktı

- `cikti/basin_bulteni.txt`: yayına hazırlanacak bülten metni.
- `cikti/basin_bulteni.xlsx`: başlık, spot ve her paragrafın dayandığı bilgi ve alıntılar, kontrol sonucu.

## Dikkat

- **Onay sizde.** Rakamları ilgili birimlere, alıntıları konuşmacılara onaylatmadan yayımlamayın. Halka açık şirketlerde
  kamuyu aydınlatma yükümlülüğü (KAP) bülten öncesinde değerlendirilmelidir.
- **İddialar.** "İlk", "en büyük" gibi ifadeleri bilgi notuna yazmadan önce dayanağını kontrol edin; yanıltıcı iddialar
  haksız rekabet ve reklam mevzuatı açısından risk oluşturur.
- **Örnek veri:** Şirket, santral ve kişiler kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk taslakta bilerek başlığa "Türkiye'nin en büyük" iddiasını ekler, kurulu gücü
50 MW yazar ve genel müdüre söylemediği bir söz atfeder. Kodun bunları yakalayıp düzelttirdiği ve basın iletişimi
bilgilerinin modele gönderilmediği doğrulanır.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. API kullanım ücretleri ve yayımlanan metnin doğruluğu kullanıcının sorumluluğundadır. Bkz.
`LICENSE` ve `SORUMLULUK_REDDI.md`.
