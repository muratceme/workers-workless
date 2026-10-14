# Geri Çağırma Kampanyası Araç Eşleştirme · Kod Bloğu

> Satış Sonrası Hizmetler (Servis) › Garanti Uzmanı · Workers / Workless

Markanın gönderdiği **geri çağırma / servis kampanyası şasi listesini** servisin müşteri kayıtlarıyla eşleştirir; kampanyası
yapılmamış araçları iletişim bilgileriyle listeler ve kampanya başladıktan sonra servise gelip işlemi **yapılmadan giden**
araçları işaretler. İnternete, marka portalına ve DMS'e bağlanmaz.

## Nasıl hesaplar?

| Adım | Kural |
|---|---|
| Şasi normalizasyonu | Büyük harfe çevrilir; boşluk ve tire atılır ("nm0 abc-123" = "NM0ABC123") |
| Eşleşen (bizde kayıtlı) | Kampanya şasisi servis müşteri kayıtlarında var (aynı şasi birden fazla ise son servis tarihi en yeni kayıt) |
| Yapılan | Şasi + kampanya kodu yapılan işlemler listesinde var |
| Bekleyen | Eşleşen ama yapılmamış araç |
| Kaçırılan ziyaret | Bekleyen aracın son servis tarihi kampanya başlangıcına eşit ya da sonra |
| Tamamlanma % | Yapılan ÷ kampanya şasi sayısı × 100; bizdeki tamamlanma % = bizde kayıtlı olup yapılan ÷ eşleşen × 100 |

**Uyarılar:** kaçırılan ziyaret (Orta) · okunamayan satırlar (Yüksek) · şasi numarası 17 karakter değil / I, O, Q içeriyor ·
kampanyada iki kez yazılmış şasi · kampanya listesinde olmayan şasiye işlem yazılmış · bekleyen araçta telefon yok (Bilgi).

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                          # örnek veri (3 kampanya)
python main.py --kampanya kampanya_sasileri.xlsx --musteriler musteriler.xlsx --yapilanlar yapilanlar.xlsx
python main.py --kampanya kampanya_sasileri.xlsx --musteriler musteriler.xlsx    # hiçbir işlem yapılmamış sayılır
```

## Çıktı

- `cikti/geri_cagirma_eslestirme.xlsx`: `Kampanya Özeti`, `Bekleyen Araçlar` (kaçırılan ziyaretler üstte, sarı), `Uyarılar`.
- `cikti/geri_cagirma_eslestirme_bekleyen.csv`: arama / bilgilendirme listesi.

## Dikkat

- **Resmî liste markanın sistemidir.** Kampanya kapsamı, yapılmış işlem ve hak ediş durumu marka garanti portalında
  kontrol edilmelidir; bu araç yalnız servis içi takip içindir.
- **Güvenlikle ilgili geri çağırmalarda** araç sahibine ulaşmak önceliklidir; marka talimatlarına ve ürün güvenliği mevzuatının
  bildirim kurallarına uyun. Güvenlik amaçlı geri çağırma bildirimi ile ticari ileti farklıdır; iletişim kanalını hukuk / uyum
  biriminizle doğrulayın.
- **Kişisel veri içerir.** Şasi, plaka, ad ve telefon KVKK kapsamındadır; listeyi yalnız yetkili personelle paylaşın.
- **"Kaydımızda yok"** araçlar başka yetkili serviste işlem görüyor ya da sahibi değişmiş olabilir.
- **Örnek veri:** Kampanyalar, şasi numaraları, müşteriler ve telefonlar kurgusaldır.

## Testler

- Şasi normalizasyonu ve biçim kontrolü, eşleşme, yapılan / bekleyen / kayıtsız sayıları, tamamlanma oranları, kaçırılan ziyaret
  ve sıralama.
- Uyarılar, aynı şasinin en yeni kaydı, CSV çıktısı, okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
