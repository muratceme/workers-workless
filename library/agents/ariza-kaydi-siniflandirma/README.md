# Arıza Kaydı Sınıflandırma · AI Agent

> Bakım Onarım › Bakım Teknisyeni · Workers / Workless

Teknisyenlerin serbest metinle yazdığı arıza kayıtlarını **makine, arıza tipi ve müdahale türüne** göre etiketler. Etiketli
kayıtlardan kod; Pareto, makine bazında duruş ve MTTR / MTBF ile tekrarlayan arıza listesini çıkarır.

## Nasıl çalışır?

1. **Eşleştirme (kod):** Kayıttaki makine adı makine listesindeki kod, ad ve "Diğer Adlar" ile karşılaştırılır
   ("enjeksiyon 1" → ENJ-01).
2. **Etiketleme (model):**
   - Kod eşleştiremediyse makine
   - Arıza tipi: mekanik, elektrik, hidrolik, pnömatik, otomasyon, kalıp / takım, ayar / operatör, diğer
   - Müdahale türü ve değişen parça
   - Olası kök neden ve önleme önerisi
3. **Denetim (kod):** Aşağıdakilerden biri varsa kayıt modele geri gönderilir (varsayılan en fazla 2 tur):
   - Makine kodu listede yok
   - Makine, kodun kesin eşleştirdiği makineden farklı
   - Metinde "değiştirildi / yenisi takıldı" var ama parça ya da "Parça değişimi" müdahalesi eksik
   - Değişen parça metinde geçmiyor
   - Alıntı metinde yok
4. **Analiz (kod):**
   - Arıza tipine göre Pareto (duruş dakikası ve kümülatif %)
   - Makine bazında arıza sayısı, toplam ve ortalama duruş
   - Tekrarlayan arıza: aynı makinede aynı tip arıza, `--pencere` gün içinde `--esik` kez (varsayılan 30 gün, 3 kez)

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                    # örnek kayıtlar + makine listesi
python agent.py --girdi arizalar.xlsx --makineler makineler.xlsx
python agent.py --girdi arizalar.xlsx --makineler makineler.xlsx --calisma-saat 600 --pencere 60 --esik 4
```

## Çıktı

- `cikti/ariza_siniflandirma.xlsx`:
  - `Tekrarlayan Arızalar`
  - `Kayıtlar`: etiketler, eşleşme şekli (kesin / model önerisi / listede yok) ve "Onay" sütunu.
  - `Pareto`: tabloyla birlikte grafik.
  - `Makine Özeti`

## Dikkat

- **Etiketler öneridir.** Bakım sorumlusu "Onay" sütunuyla onaylamalıdır. Kök neden önerileri doğrulanmamış hipotezdir;
  kalıcı çözüm için yerinde analiz (5 Neden, balık kılçığı) gerekir.
- **MTTR / MTBF yaklaşıktır.** Ortalama duruş, onarım süresinin yaklaşığıdır (bekleme ve parça tedarik süresi dahil
  olabilir). MTBF yalnız `--calisma-saat` verilirse ve tüm makineler için aynı planlı çalışma süresi varsayılarak hesaplanır.
- **Veri gönderimi.** Arıza açıklamaları ve makine listesi seçtiğiniz yapay zekâ sağlayıcısına gönderilir. Telefon ve
  e-posta maskelenir. Gerekirse yerel model (Ollama) kullanın.
- **Örnek veri** kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk taslakta bilerek bir kaydı yanlış makineye bağlar, parça değişimini atlar ve
listede olmayan bir makine kodu uydurur. Kodun bunları yakalayıp düzelttirdiği, tekrarlayan hidrolik arızasını bulduğu ve
Pareto / makine özetini doğru hesapladığı doğrulanır.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. API kullanım ücretleri ve etiketlerin doğruluğu kullanıcının sorumluluğundadır. Bkz.
`LICENSE` ve `SORUMLULUK_REDDI.md`.
