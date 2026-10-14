# Mağaza Ziyaret Raporu · AI Agent

> Mağaza Operasyonları › Bölge Müdürü · Workers / Workless

Mağaza ziyaretinin **kontrol listesi** ve **ziyaret notlarından**, mağaza müdürüyle paylaşılacak aksiyon odaklı raporu hazırlar:
özet, bölüm uyum puanları, güçlü yönler, terminli ve sorumlu atanmış aksiyon planı, görüşülen konular ve takip ziyaretinde
bakılacaklar. Uygun olmayan hiçbir madde aksiyonsuz kalamaz.

## Nasıl çalışır?

1. **Puanlama (kod):**
   - Kontrol listesindeki maddelere kimlik verilir (K1, K2…).
   - Bölüm ve genel uyum puanı hesaplanır: Uygun 1, Kısmen 0,5, Uygun değil 0; DD (değerlendirme dışı) hesaba girmez.
   - Sonuç sütununa 2 / 1 / 0 da yazılabilir.
2. **Yazım (model):** Rapor ve aksiyon planı yazılır. Her aksiyon ilgili madde kimliklerini, sorumlu rolünü (notlarda isim
   yoksa), termini ve önceliği içerir.
3. **Denetim (kod):** Aşağıdaki durumlardan biri varsa taslak modele geri gönderilir (varsayılan en fazla 2 tur):
   - Aksiyonu olmayan "Uygun değil" ya da "Kısmen" madde
   - Geçersiz ya da ziyaret tarihi ile `--azami` (90) gün sonrası dışında kalan termin
   - Sorumlusu yazılmamış aksiyon
   - Önceliği Yüksek olmayan güvenlik maddesi (acil çıkış, yangın, İSG, elektrik)
   - Uygun olmayan maddenin güçlü yön gösterilmesi
   - Geçersiz madde kimliği

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                            # örnek ziyaret (14.10.2026)
python agent.py --kontrol kontrol_listesi.xlsx --notlar ziyaret_notlari.docx --tarih 14.10.2026
python agent.py --kontrol kontrol_listesi.xlsx --azami 30 --tur 3
```

`--tarih` verilmezse notlardaki ilk GG.AA.YYYY tarihi ziyaret günü sayılır.

## Çıktı

- `Rapor`: metin rapor ve son kontrol.
- `Aksiyon Planı`: öncelik sırasıyla; takip için "Durum" ve "Kapanış Notu" sütunları.
- `Bölüm Puanları`: %90 ve üstü yeşil, %70–90 sarı, altı kırmızı.
- `Kontrol Listesi`.
- `.md` rapor.

## Dikkat

- **Taslaktır.** Bölge müdürü kontrol edip mağaza müdürüyle paylaşmalıdır. Aksiyon terminleri ve sorumluları mağazayla
  mutabakatla kesinleştirilmelidir.
- **Güvenlik bulguları** (acil çıkış önünün kapalı olması gibi) raporu beklemeden yerinde giderilmeli ve işyeri İSG
  prosedürüne göre kaydedilmelidir.
- **Veri gönderimi.** Kontrol listesi ve notlar, seçtiğiniz yapay zekâ sağlayıcısına gönderilir. Telefon ve e-posta maskelenir;
  personel adları maskelenmez. Notlara personel hakkında gereksiz kişisel bilgi yazmayın.
- **Örnek veri:** Mağaza, kişiler ve bulgular kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk taslakta bilerek iade maddesine aksiyon yazmaz, acil çıkış aksiyonunu "Orta"
öncelikle verir, ziyaretten önceki bir termin yazar ve kısmen uygun bir maddeyi güçlü yön gösterir. Kodun bunları yakalayıp
düzelttirdiği doğrulanır.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. API kullanım ücretleri ve raporun doğruluğu kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
