# Müvekkil Dosya Durum Raporu · AI Agent

> Dava Takip › Avukat (Büro) · Workers / Workless

Dava ve icra dosyalarının son durumunu **müvekkile gönderilecek sade bir rapor** olarak hazırlar. **Müvekkil ve karşı
taraf adları yapay zekâ modeline gönderilmez;** `{MUVEKKIL}`, `{KT1}`, `{KT2}` … kodlarıyla gönderilir, rapora kod geri
yazar. Raporlar gönderilmez.

## Nasıl çalışır?

1. **Hazırlık (kod):** Dosya listesi müvekkile göre gruplanır ve her dosya için şunlar hesaplanır:
   - Kalan alacak (alacak − tahsil edilen)
   - Son işlemden bu yana geçen gün; `--hareketsiz` günü aşan dosya hareketsiz sayılır (varsayılan 90)
   - Sonraki işleme kalan gün
2. **Rapor (model):** Her müvekkil için genel özet, dosya bazında sade dille durum ve müvekkilden beklenenler yazılır.
3. **Denetim (kod):** Aşağıdakilerden biri varsa yalnız o müvekkilin raporu modele geri gönderilir (varsayılan en fazla
   2 tur):
   - Müvekkilin bir dosyası raporda yok ya da başka müvekkilin dosyası eklenmiş
   - Başka müvekkilin dosya ya da esas numarası metne karışmış
   - Başka müvekkilin karşı taraf kodu kullanılmış
   - Dosya metninde o dosyanın kaydında olmayan tarih ya da tutar var
   - Sonraki işlem / duruşma tarihi aynen yazılmamış
   - "Kazanacağız", "kesinlikle", "garanti" gibi sonuç vaadi var
4. **Raporlar (kod):** Adlar geri yazılır; kontrolden geçen her müvekkil için ayrı .txt rapor üretilir.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py --bugun 15.10.2026                                         # örnek dosya listesi
python agent.py --girdi dosyalar.xlsx --buro "Av. Ad Soyad Hukuk Bürosu" --hareketsiz 60
```

Dosya listesini büro yazılımınızdan ya da UYAP'tan aldığınız bilgilerle Excel'de tutabilirsiniz. Sütunlar:
Dosya No, Müvekkil, Tür, Mahkeme / Daire, Esas No, Karşı Taraf, Son İşlem Tarihi, Son İşlem, Sonraki İşlem Tarihi,
Sonraki İşlem, Alacak Tutarı, Tahsil Edilen.

## Çıktı

- `cikti/dosya_durum_raporu_<muvekkil>.txt`: müvekkil başına rapor.
- `cikti/dosya_durum_raporu.xlsx` → `Dosyalar` sayfası:
  - hareketsiz gün (sınırı aşanlar kırmızı), sonraki işleme kalan gün (14 gün ve altı sarı),
  - kalan alacak ve rapordaki açıklama.

## Dikkat

- **Avukat onayı.** Raporlar avukat okuyup onaylamadan müvekkile iletilmemelidir. Dosya bilgilerini UYAP kayıtlarıyla
  karşılaştırın.
- **Meslek sırrı.** Mahkeme, esas numarası, işlem ve tutar bilgileri modele gönderilir; adlar gönderilmez. Avukatlık
  meslek kuralları ve müvekkil sırrı açısından uygunluğunu değerlendirin. Gerekirse yerel model (Ollama) kullanın.
- **Sonuç taahhüdü yok.** Rapor bilgilendirme amaçlıdır; sonuna bunu belirten bir not eklenir.
- **Örnek veri:** Müvekkiller, dosyalar ve esas numaraları kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk raporda bilerek:
- duruşma tarihini 26.11.2026 yazar (kayıtta 19.11.2026),
- "davayı kazanacağımız kesindir" der,
- hareketsiz dosyayı rapordan çıkarır.

Kodun bunları yakaladığı, yalnız sorunlu müvekkilin raporunu yeniden yazdırdığı, adların modele gönderilmediği ve başka
müvekkilin dosyasının rapora karışmasını tespit ettiği doğrulanır.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Hukuki danışmanlık değildir. API kullanım ücretleri ve müvekkile iletilen bilgilerin
doğruluğu kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
