# Yönetim Kurulu Finans Sunumu Özeti · AI Agent

> Finans › Finans Müdürü · Workers / Workless

Dönem sonuç tablosundan **yönetim kurulu için kısa, yorumlu finans özeti** hazırlar. Değişimleri, bütçe sapmalarını ve
oranları **kod hesaplar**; model yorumlar ve önceliklendirir. Metindeki her rakam kodun hesapladığı değerlerle
karşılaştırılır.

## Nasıl çalışır?

1. **Hesap (kod):**
   - Her kalemin önceki döneme göre değişimi (tutar, %) ve bütçe sapması (tutar, %). Eksi ya da parantezli yazılan gider
     kalemleri mutlak değerle değerlendirilir ("gider arttı").
   - Tanınan kalemlerden brüt kâr, FAVÖK ve net kâr marjı (puan farkıyla) ve cari oran.
   - Bütçeden en çok sapan 3 kalem.
   - Yönetim notları N1, N2 … diye numaralanır.
2. **Özet (model):** Öne çıkanlar, bölümler (gelir, kârlılık, bütçe, nakit ve borç, riskler) ve karar gerektiren konular
   yazılır. Her madde dayandığı kalem / oran / not numarasını gösterir.
3. **Denetim (kod):** Aşağıdakilerden biri varsa taslak modele geri gönderilir (varsayılan en fazla 2 tur):
   - Maddedeki bir sayı, dayandığı kalem / oran / notun hazır rakamlarından biri değil (yuvarlama ya da yeni hesap)
   - "Arttı / azaldı" yönü değişimin işaretiyle çelişiyor
   - Neden-sonuç ifadesi ("nedeniyle", "kaynaklı") bir yönetim notuna dayanmıyor
   - Bütçeden en çok sapan kalemlerden biri ele alınmamış
   - "Kesinlikle", "garanti" gibi ifadeler var

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                         # örnek çeyrek sonuçları
python agent.py --girdi sonuclar.xlsx --notlar notlar.txt --donem "2026 3. Çeyrek" --birim "milyon TL"
```

Tablo örneği (Bütçe ve Önceki Dönem isteğe bağlı):

| Kalem | Cari Dönem | Önceki Dönem | Bütçe |
|---|---|---|---|
| Hasılat | 1.250,0 | 1.080,0 | 1.300,0 |
| Satışların Maliyeti | (905,0) | (760,0) | (930,0) |
| Net Dönem Kârı | 71,0 | 86,0 | 105,0 |

Oranlar için kalem adları: Hasılat / Net Satışlar, Brüt Kâr, FAVÖK / EBITDA, Net Dönem Kârı, Dönen Varlıklar, Kısa Vadeli
Yükümlülükler.

## Çıktı

- `cikti/yk_finans_ozeti.xlsx`:
  - `Özet`: maddeler ve dayanakları.
  - `Rakamlar`: kodun hesapladığı değişimler, bütçe sapmaları ve oranlar.
  - `Yönetim Notları`.
- `cikti/yk_finans_ozeti.txt`: sunuma ya da e-postaya eklenecek metin.

## Dikkat

- **Taslaktır.** Finans müdürü onaylamadan yönetim kuruluna sunulmamalıdır. Rakamlar girdiğiniz tabloya dayanır;
  tablonun doğruluğu (konsolidasyon, TFRS / VUK farkları) sizin sorumluluğunuzdadır.
- **Neden-sonuç.** Model, yönetim notunda yazmayan bir nedeni yorum olarak eklemez; notlarınız ne kadar açıksa özet o
  kadar iyi olur.
- **Gizlilik.** Dönem rakamları ve notlar modele gönderilir. Kamuya açıklanmamış finansal bilgi için şirketinizin
  içeriden öğrenilen bilgi politikasına (halka açık şirketlerde SPK düzenlemeleri) uyun.
- **Örnek veri:** Rakamlar kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk taslakta bilerek:
- hasılat artışını "%16" diye yuvarlar,
- azalan net kâra "yükseldi" der,
- finansman giderindeki artışın nedenini nota dayandırmadan yazar,
- bütçeden çok sapan esas faaliyet kârını atlar,
- "kesinlikle" der.

Kodun bunları yakalayıp düzelttirdiği doğrulanır.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. API kullanım ücretleri ve özetin kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
