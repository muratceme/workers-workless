# Provizyon Talebi Değerlendirme · Kod Bloğu

> Sigortacılık › Sağlık Sigortaları › Provizyon Uzmanı · Workers / Workless

Anlaşmalı sağlık kurumlarından gelen provizyon taleplerini ön değerlendirir. Talebi poliçe teminatı, yıllık limit,
katılım payı, bekleme süreleri, ön mevcut durum istisnaları ve genel şart kurallarıyla karşılaştırır; her talep
için ön karar, gerekçe ve ödenecek tutarı çıkarır. İnternete bağlanmaz.

## Karar sırası

Talepler tarih sırasıyla işlenir ve onaylanan tutar kalan limitten düşülür. İlk uyan kural kararı verir:

| Sıra | Kontrol | Ön karar |
|---|---|---|
| 1 | Kurum anlaşmalı değil: provizyon verilmez, sigortalı geri ödeme yoluyla başvurabilir | Ret |
| 2 | Poliçe yok veya talep tarihi poliçe süresi dışında | Ret |
| 3 | Teminat sigortalının planında yok | Ret |
| 4 | Genel şart istisnası (ICD-10 öneki, ör. Z41.1 estetik) | Ret |
| 5 | Poliçede yazılı ön mevcut durum istisnası (poliçedeki ICD-10 kodları) | Ret |
| 6 | Teminatın (ör. doğum) veya tanının (ör. katarakt) bekleme süresi, ilk giriş tarihinden bu yana dolmamış | Ret |
| 7 | ICD-10 kodu yok veya geçersiz; prim gecikmede; ilk girişten sonraki 90 gün içinde kronik tanı (`--kronik-gun`) | İnceleme |
| 8 | Ödenecek = talep × (1 − katılım payı); kalan yıllık limitle sınırlıdır | Onay / Kısmi Onay / Ret (limit doldu) |

**Ek kurallar:**
- **Bekleme süresi:** Yenilemelerde kazanılmış hakkı yansıtmak için bekleme süresi "İlk Giriş Tarihi"nden sayılır.
  Bu tarih boşsa poliçe başlangıcı kullanılır.
- **Kalan limit:** Önceki ödemeler (`--odemeler`) ve aynı çalışmada onaylanan talepler kalan limitten düşülür.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                     # örnek: 13 talep, 7 poliçe
python main.py --talepler talepler.xlsx --policeler policeler.xlsx --plan plan_teminatlari.csv --kurallar genel_sartlar.csv
python main.py --talepler t.xlsx --policeler p.xlsx --plan plan.csv --kurallar kurallar.csv --odemeler odemeler.xlsx --kronik-gun 180
```

| Dosya | Sütunlar |
|---|---|
| Provizyon talepleri | Talep No, Talep Tarihi, Kurum, Anlaşmalı, Sigortalı No, Poliçe No, Teminat, ICD-10, Tanı, İşlem, Talep Tutarı (TL) |
| Poliçeler | Poliçe No, Sigortalı No, Ad Soyad, Plan, Başlangıç, Bitiş, İlk Giriş Tarihi, Prim Durumu, Beyan / Ön Mevcut Durum İstisnası (ICD-10 kodlarıyla) |
| Plan teminatları | Plan, Teminat, Yıllık Limit (TL), Katılım Payı %, Bekleme Süresi (gün) |
| Genel şart kuralları | ICD-10 / Anahtar (önek), Kapsam (İstisna / Bekleme / Kronik), Bekleme Süresi (gün), Açıklama |
| Önceki ödemeler (isteğe bağlı) | Sigortalı No, Teminat, Ödenen Tutar (TL) |

## Çıktı

`provizyon_degerlendirme.xlsx`:
- `Değerlendirme`: her talep için şunları içerir:
  - ön karar ve gerekçeler
  - katılım payı, ödenecek tutar
  - kalan limit (önce / sonra)
  - kuruma gönderilecek kısa cevap taslağı
  - boş "Uzman Kararı" sütunu

  Altta karar özeti yer alır.
- `Limit Durumu`: sigortalı ve teminat bazında limit, kullanılan, kalan ve kullanım oranı.
- `Uyarılar`: sigortalı / poliçe uyuşmazlığı, yüksek tutarlı talep (100.000 TL ve üstü).

## Dikkat

- **Kurallar örnektir:** `genel_sartlar.csv`, plan limitleri, katılım payları ve bekleme süreleri yalnızca
  gösterim içindir. Örnek değerler şunlardır:
  - estetik, infertilite, obezite istisnaları
  - doğum için 300–365 gün, bazı ameliyatlar için 180 gün bekleme

  Kendi ürününüzün genel ve özel şartlarını ve poliçe zeyillerini girin.
- **Ön değerlendirmedir:** Tıbbi gereklilik, fiyat uygunluğu, kaza / hastalık ayrımı ve ön mevcut durumun
  araştırılması provizyon uzmanı ve şirket hekiminin işidir. "İnceleme" sonuçlarında ödenebilir tutar yalnızca
  bilgi olarak gösterilir.
- **ICD-10 eşleştirmesi:** Önek ile yapılır. "E11" kuralı E11.0–E11.9 kodlarının hepsini kapsar. Tanı kodu yanlış
  girilmişse kural uygulanamaz.
- **Örnek veri:** Örnek sigortalılar, poliçeler ve talepler kurgusaldır.

## Testler

Şunlar test edilir:
- Her karar yolu: anlaşmasız, süre dışı, teminat yok, istisna, ön mevcut durum, bekleme, inceleme, kısmi onay, limit
  doldu.
- Katılım payı ve kalan limit (önceki ödemeler ve aynı çalışmadaki onaylar).
- Excel sayfaları ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
