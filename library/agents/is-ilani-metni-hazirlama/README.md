# İş İlanı Metni Hazırlama · AI Agent

> İnsan Kaynakları › İşe Alım · Workers / Workless

Departmanın doldurduğu **pozisyon talep formundan** yayına hazır iş ilanı hazırlar. İlan **ayrımcı ifade
içermez**. Talep formundaki "25-35 yaş arası, askerliğini yapmış erkek aday, hoş görünümlü" gibi riskli şartlar
kodla bulunur ve ilandan çıkarılır. İşin gereği olan şartlar (ör. sahada araç kullanılacaksa sürücü belgesi)
gerekçesiyle yazılır.

## Nasıl çalışır?

1. **Tarama (kod):** Talep formu ortak tarayıcıyla (`ayrimcilik.py`) taranır. İki seviye bulgu vardır:
   - **Yüksek risk:** yaş, cinsiyet, medeni hâl ve aile, görünüş, sağlık ve engellilik, din ve inanç, etnik
     köken, siyasi görüş ve sendika
   - **Dikkat:** askerlik, ehliyet, uyruk, anadili, gereksiz kişisel veri (fotoğraf, TC kimlik no, kan grubu),
     adli sicil, sigara
2. **Yazım (model):** İlanda başlık, giriş, sorumluluklar, aranan nitelikler, tercih sebepleri, sunulanlar,
   çalışma bilgisi ve başvuru bölümü (KVKK aydınlatma bağlantısı yer tutucusuyla) yer alır. Ayrıca kısa sosyal
   medya versiyonu yazılır.
   - Ayrımcı şartlar çıkarılır ve listelenir.
   - Kişilik özellikleri gözlemlenebilir yetkinliğe çevrilir ("güler yüzlü" → "müşteri ilişkilerinde güçlü
     iletişim").
3. **Denetim (kod):** Üretilen ilan yeniden taranır. Şu durumlarda taslak modele geri gönderilir (varsayılan en
   fazla 2 tur):
   - Yüksek riskli ifade kalmışsa
   - "Dikkat" ifadesi gerekçesiz kullanılmışsa
   - Gereksiz kişisel veri isteniyorsa

**Dayanak:** 4857 sayılı İş Kanunu md. 5 (eşit davranma), 6701 sayılı TİHEK Kanunu md. 3, 6 ve 7, 6356 sayılı
Kanun md. 25 (sendika), KVKK md. 4 (ölçülülük). TİHEK iş ilanlarındaki ayrımcı şartlara idari para cezası
vermektedir.

> Tarama kural tabanlıdır ve **hukuki değerlendirme yerine geçmez**. Yayımlamadan önce İK ve gerekiyorsa hukuk
> onayı alın.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                            # örnek talep formuyla (saha satış temsilcisi)
python agent.py --girdi talep_formu.docx --tur 3
```

## Çıktı

`İlan` (ilan metni, kısa versiyon, son kontrol) · `Uygunluk` (talep formunda bulunanlar, ilandan çıkarılanlar,
gerekçeyle tutulanlar, kalan sorunlar, "Onay") · `.md` ilan

## Testler

Testler gerçek API çağırmaz. Sahte model ilk taslakta bilerek "hoş görünümlü" bırakır; kodun bunu yakalayıp
düzelttirdiği doğrulanır.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. API kullanım ücretleri ve yayımlanan ilanın hukuka uygunluğu kullanıcının
sorumluluğundadır. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
