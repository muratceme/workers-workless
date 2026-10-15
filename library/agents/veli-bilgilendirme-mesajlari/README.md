# Veli Bilgilendirme Mesajları · AI Agent

> Kayıt Kabul ve Öğrenci İşleri › Öğrenci İşleri Sorumlusu · Workers / Workless

Duyuru, devamsızlık ve eksik ödev bilgilerinden **kişiselleştirilmiş veli mesajları** hazırlar. **Öğrenci ve veli bilgisi
yapay zekâ modeline gönderilmez.** Model yalnız mesaj türü ve duyuru metni için yer tutuculu şablon yazar; adları, sınıfı ve
sayıları kod doldurur. Mesaj göndermez; SMS sistemine aktarılacak CSV üretir.

## Nasıl çalışır?

1. **Hazırlık (kod):** Listedeki mesaj türlerinden (Devamsızlık, Ödev, her duyuru) şablon kombinasyonları çıkarılır.
2. **Şablon (model):** Her kombinasyon için yer tutuculu kısa bir şablon yazılır: `{VELI}`, `{OGRENCI}`, `{SINIF}`,
   `{OZURSUZ}`, `{OZURLU}`, `{DERS}`, `{ODEV}`, `{OKUL}`, `{ILETISIM}`.
3. **Denetim (kod):** Aşağıdakilerden biri varsa şablonlar modele geri gönderilir (varsayılan en fazla 2 tur):
   - Zorunlu yer tutucu eksik ya da bilinmeyen yer tutucu var
   - Yer tutucuya ek getirilmiş ("{OGRENCI}'in"); ada göre yanlış ek çıkabilir
   - Duyurudaki tarih ya da saat şablonda aynen yok, ya da duyuruda olmayan sayı eklenmiş
   - Devamsızlık / ödev şablonunda sabit sayı yazılmış
   - Yargılayıcı ya da tehdit içeren ifade ("tembel", "başarısız", "sınıfta kalacak", "disiplin", "ceza")
   - En uzun değerlerle doldurulan mesaj `--azami` karakteri aşıyor (varsayılan 480)

   Kontrolden geçmeyen şablonla mesaj üretilmez.
4. **Doldurma (kod):** Telefonu olan veliler için mesajlar oluşturulur; telefonu olmayanlar "aranmalı" diye listelenir.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                          # örnek liste + duyurular
python agent.py --girdi liste.xlsx --duyurular duyurular.txt --okul "Okul Adı" --iletisim "0212 000 00 00"
python agent.py --girdi liste.xlsx --duyurular duyurular.txt --okul "…" --iletisim "…" --azami 300
```

Devamsızlık gün sayılarını e-Okul'dan aldığınız listeden doldurun. Duyuru dosyasında her duyuruyu "D1: …" biçiminde bir
satıra yazın ve listedeki "Duyuru" sütununa numarasını girin.

## Çıktı

- `cikti/veli_mesajlari.xlsx`:
  - `Şablonlar`: kontrol sonucu ve "Onay" sütunu.
  - `Mesajlar`: doldurulmuş mesajlar ve karakter sayısı.
  - `Gönderilemeyen`: telefonu olmayanlar ve kontrolden geçmeyen şablonlar.
- `cikti/veli_mesajlari_sms.csv`: Telefon; Mesaj.

## Dikkat

- **Şablonları okul idaresi onaylamalıdır.** Mesajlar gönderilmez; gönderimi okulunuzun SMS sistemi ya da e-Okul VBS üzerinden
  yaparsınız.
- **Devamsızlık bildirimleri.** Mevzuatın öngördüğü resmî veli bildirimleri (ör. ortaöğretimde özürsüz devamsızlıkla ilgili
  yazılı bildirim) e-Okul ve okul idaresi üzerinden yapılır; bu mesajlar bilgilendirme amaçlıdır, resmî bildirimin yerine
  geçmez.
- **Kişisel veri.** Öğrenci ve veli bilgisi modele gönderilmez; yalnız mesaj türleri ve duyuru metinleri gönderilir. CSV ve
  Excel çıktısını yalnız yetkili personel kullanmalıdır.
- **Örnek veri:** Öğrenciler, veliler ve telefonlar kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk taslakta bilerek devamsızlık şablonuna "10 günü aşarsa sınıfta kalacaktır" yazar,
ödev şablonunda yer tutucuya ek getirir ve iletişim telefonunu unutur, duyurudaki saati değiştirir. Kodun bunları yakalayıp
düzelttirdiği ve öğrenci / veli bilgisini modele göndermediği doğrulanır.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. API kullanım ücretleri ve gönderilen mesajların doğruluğu kullanıcının sorumluluğundadır. Bkz.
`LICENSE` ve `SORUMLULUK_REDDI.md`.
