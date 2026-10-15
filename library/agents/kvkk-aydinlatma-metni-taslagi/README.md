# KVKK Aydınlatma Metni Taslağı · AI Agent

> Hukuk › Avukat · Workers / Workless

Veri işleme envanterinden **6698 sayılı KVKK md. 10'a uygun aydınlatma metni taslağı** hazırlar. Amaç, hukuki sebep,
aktarım ve toplama yöntemi bölümlerini model yazar. **Veri sorumlusunun kimliği, md. 11 hakları ve başvuru yöntemleri
kod tarafından aynen eklenir.** Envanterde olmayan bir hukuki sebep ya da "vb." gibi genel ifade yazılırsa taslak
modele geri gönderilir.

## Nasıl çalışır?

1. **Envanter kontrolü (kod):** Her satırın hukuki sebebi KVKK bendine çevrilir ("5/2-c", "md. 5/2 (ç)", "Açık rıza" →
   özel nitelikli veride 6/3-a). Aşağıdakiler "Envanter Kontrol" sayfasında işaretlenir:
   - Hukuki sebep tanınmıyor (md. 5/1, 5/2-a … f, 6/3-a … g dışında)
   - Özel nitelikli veri (sağlık, biyometrik, din, ceza mahkûmiyeti, sendika üyeliği …) md. 5'e dayandırılmış
     (7499 sayılı Kanun'la değişik md. 6/3 şartlarından biri gerekir)
   - Yurt dışı aktarımda md. 9 mekanizması yok (yeterlilik kararı, standart sözleşme, bağlayıcı şirket kuralları,
     taahhütname, arızi aktarım)
2. **Taslak (model):** Giriş, amaç bölümleri (her biri dayandığı E satırlarıyla), aktarım ve toplama yöntemi.
3. **Denetim (kod):** Aşağıdakilerden biri varsa taslak modele geri gönderilir (varsayılan en fazla 2 tur):
   - Bir envanter satırı hiçbir bölümde yok; bölüm, dayandığı satırın amacını ya da hukuki sebebini yazmamış
   - Metinde envanterde olmayan bir hukuki sebep geçiyor (ör. meşru menfaate dayanan veri için "açık rızanıza dayanılarak")
   - Bir alıcı grubu, yurt dışı aktarım ve mekanizması ya da bir toplama yöntemi yazılmamış
   - "vb.", "ve benzeri", "sınırlı olmamak üzere" gibi genel ifadeler; "okudum, onaylıyorum" gibi açık rıza beyanı
4. **Metin (kod):** Veri sorumlusu (unvan, MERSİS, adres), md. 11'deki haklar ve Veri Sorumlusuna Başvuru Usul ve
   Esasları Hakkında Tebliğ'deki başvuru yolları (yazılı, KEP, güvenli elektronik imza / mobil imza, kayıtlı e-posta;
   en geç 30 gün, kural olarak ücretsiz) eklenir.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                # örnek (kurgusal) e-ticaret sitesi
python agent.py --envanter envanter.xlsx --sorumlu sirket.txt
```

Envanter örneği (`;` ile ayrılmış):

| Veri Kategorisi | Veriler | İşleme Amacı | Hukuki Sebep | Toplama Yöntemi | Alıcı Grubu | Aktarım Amacı | Yurt Dışı Aktarım |
|---|---|---|---|---|---|---|---|
| Kimlik | Ad, soyad | Üyelik ve sipariş süreçleri | 5/2-c | Üyelik formu | Kargo şirketleri | Teslimat | |
| İşlem Güvenliği | IP adresi | Bilgi güvenliği | 5/2-f | Otomatik | Bulut barındırma hizmeti sağlayıcısı | Sunucu hizmeti | Standart sözleşme |

## Çıktı

- `cikti/aydinlatma_metni.txt`: yayına hazırlanacak metin taslağı.
- `cikti/aydinlatma_metni.xlsx`: Metin (bölüm ve dayanaklar), Envanter Kontrol.

## Dikkat

- **Hukuki değerlendirme sizde.** Hukuki sebebin doğru seçilmesi (ör. meşru menfaat için denge testi) envanterin
  sorumluluğundadır; agent yalnız envantere sadık metin yazar. Yayından önce avukat / KVKK sorumlusu incelemelidir.
- **Açık rıza ayrı alınır.** Aydınlatma metni rıza beyanı içermez; açık rıza gereken işlemler için ayrı metin hazırlayın.
- **Kapsam:** Kurum karar ve rehberleri ile sektörel mevzuat (ör. sağlık, bankacılık, ticari elektronik ileti)
  ayrıca dikkate alınmalıdır. Kanunun 7499 sayılı Kanun'la (01.06.2024) değişen md. 6 ve 9 hükümleri esas alınmıştır;
  mevzuatın güncel hâlini kontrol edin.
- **Örnek veri:** Şirket kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk taslakta meşru menfaate dayanan güvenlik kayıtlarını "açık rızanıza
dayanılarak" diye yazar, "vb." kullanır, hukuki süreç satırını atlar, yurt dışı aktarımı yazmaz ve sona "okudum,
onaylıyorum" ekler. Kodun beş sorunu da yakaladığı, envanter kontrolünün özel nitelikli veriyi ve tanınmayan hukuki
sebebi işaretlediği doğrulanır.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. API kullanım ücretleri ve metnin hukuka uygunluğu kullanıcının sorumluluğundadır. Bkz.
`LICENSE` ve `SORUMLULUK_REDDI.md`.
