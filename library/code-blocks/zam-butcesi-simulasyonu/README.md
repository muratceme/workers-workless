# Zam Bütçesi Simülasyonu · Kod Bloğu

> İnsan Kaynakları › Ücretlendirme ve Yan Haklar Uzmanı · Workers / Workless

Farklı zam senaryolarında kişi bazında yeni brüt ücretleri ve toplam personel maliyetini hesaplar. Senaryoları
yan yana karşılaştırır; istenirse her senaryoyu verilen bütçeye ölçekler. İnternete bağlanmaz.

## Nasıl hesaplar?

- **Senaryo dosyası:** Her satır bir kuraldır: Performans × Konum → Oran (%) + Sabit Tutar (TL). `*` her değere
  uyar; bir çalışana birden fazla satır uyarsa en özel olanı seçilir. Aynı dosyada örneğin şunlar tanımlanabilir:

| Senaryo | Performans | Konum | Oran (%) | Sabit Tutar (TL) |
|---|---|---|---|---|
| Genel %25 | * | * | 25 | |
| Seyyanen 3.000 TL + %20 | * | * | 20 | 3000 |
| Performans matrisi | A | Alt | 36 | |
| Performans matrisi | * | * | 20 | |

- **Konum:** Mevcut ücretin bant orta noktasına oranıdır (karşılaştırma oranı). 0,90'ın altı "Alt", 0,90–1,10
  arası "Orta", 1,10'un üstü "Üst" sayılır. Bant dosyası verilmezse herkes "Orta" sayılır.
- **Performans:** A / B / C / D veya 1–5 olabilir. 1–5 ölçeğinde 5 ve 4 "A", 3 "B", 2 "C", 1 "D" sayılır.
- **Yeni ücret:** mevcut × (1 + oran) + sabit tutar. Kısmi süreli çalışanda sabit tutar çalışma oranıyla çarpılır.
- **Asgari ücret tabanı:** Yeni ücret, geçerlilik dönemindeki brüt asgari ücretin altında kalırsa asgari ücrete
  çekilir; fark ayrıca gösterilir. Dönemin asgari ücreti `tr_parametreler.json` içinde yoksa önceki yılın değeri
  kullanılır ve uyarı verilir. Açıklandığında `--yeni-asgari` ile yeniden çalıştırın.
- **İşveren maliyeti:** brüt + SGK işveren payı + işsizlik sigortası işveren payı. Prime esas kazanç SGK tavanıyla
  sınırlıdır; oranlar `tr_parametreler.json` dosyasından gelir. İsteğe bağlı olarak 5 veya 2 puanlık teşvik
  düşülebilir (`--tesvik`).

**İsteğe bağlı seçenekler:**

| Seçenek | Ne yapar? |
|---|---|
| `--butce 25` | Toplam aylık brüt artışın bütçeyle farkını gösterir. Bütçeyi tutturmak için her senaryonun oran kısmı orantılı ölçeklenmiş bir "(bütçeye ölçekli)" kopyası da üretilir. |
| `--kist` | Son 12 ayda işe girenlere oran kısmı çalışılan ay / 12 oranında uygulanır. |
| `--bant-ustu-sinirla` | Yeni ücret bant üstünü aşarsa ücret üst sınırda tutulur. Aşan kısım × 12, tek seferlik ödeme olarak maliyete eklenir. Bant üstü `--bant-artis` ile yeni döneme taşınabilir. |
| `--yuvarla 100` | Yeni ücret 100 TL'nin katına yukarı yuvarlanır. |

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                     # örnek: 28 çalışan, 3 senaryo, geçerlilik 01.01.2027
python main.py --calisanlar calisanlar.xlsx --senaryolar senaryolar.csv --bantlar bantlar.xlsx --butce 25
python main.py --calisanlar c.xlsx --senaryolar s.csv --yeni-asgari 40000 --kist --gecerlilik 01.01.2027 --yuvarla 100
```

| Dosya | Sütunlar |
|---|---|
| Çalışanlar | Sicil No, Ad Soyad, Departman, Pozisyon, Kademe, Brüt Ücret (TL), İşe Giriş, Performans, Çalışma Oranı |
| Senaryolar | Senaryo, Performans, Konum, Oran (%), Sabit Tutar (TL) |
| Bantlar (isteğe bağlı) | Kademe, Pozisyon (boş = kademe bandı), Alt, Orta, Üst |

Çalışan ve bant dosyaları **Ücret Bandı Karşılaştırması** paketiyle aynı biçimdedir.

## Çıktı

`zam_butcesi_simulasyonu.xlsx`:
- `Senaryo Karşılaştırması`:
  - mevcut ve yeni aylık brüt, toplam ve kişi bazında artış oranları
  - asgari ücrete çekilen ve bant üstünde sınırlanan kişi sayısı
  - yıllık işveren maliyeti artışı, bütçe farkı
  - boş "Karar" sütunu
- `Kişi Bazında`: her senaryo için yeni brüt, artış oranı ve not.
- `Departman`: senaryo bazında aylık artış.
- `Senaryo Tanımları`: her kuralın uygulanan oranı ve kaç kişiye uyduğu.
- `Uyarılar`.

## Dikkat

- **Senaryolar örnektir:** Örnek oranlar ve matris kurgusaldır. Zam politikası, enflasyon beklentisi ve toplu iş
  sözleşmesi hükümleri kurumunuza aittir. Toplu iş sözleşmesi kapsamındaki çalışanlarda sözleşmedeki zam hükmü
  uygulanır.
- **Parametreler:** Yeni dönemin SGK oranları, tavanı ve asgari ücreti açıklandığında `tr_parametreler.json`
  kaynağıyla birlikte güncellenmelidir.
- **Maliyet kapsamı:** Yıllık maliyet aylık × 12 + tek seferlik ödemedir. İkramiye, yan haklar, fazla mesai ve
  kıdem tazminatı karşılığı dahil değildir.
- **Kişisel veri:** Ücret verisi gizlidir. Raporu yetkili kişilerle sınırlı paylaşın.

## Testler

Şunlar test edilir:
- Genel oran, seyyanen + oran (kısmi süreli dahil) ve matris konumu (pozisyon bandı dahil).
- SGK tavanlı işveren maliyeti; bütçeye ölçekleme.
- Kıst zam, bant üstü sınırı, asgari ücret tabanı ve yuvarlama.
- Kuralı olmayan çalışan; bant dosyası verilmeden konum kullanımı.
- Excel sayfaları ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
