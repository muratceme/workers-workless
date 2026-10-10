# Anket Sonuç Analizi · Kod Bloğu

> Pazarlama › Pazarlama Uzmanı · Workers / Workless

Pazar araştırması anketlerinin **kapalı ve açık uçlu** cevaplarını analiz eder. Frekans ve ölçek tablolarını hata
payıyla verir, demografik kırılımlarda çapraz tablo ve **ki-kare testi** yapar, açık uçlu yanıtları bir kod
çerçevesiyle temalara ayırır ve veri kalitesi sorunlarını işaretler. İnternete bağlanmaz.

## Nasıl hesaplar?

| Soru türü | Çıktı |
|---|---|
| Tek seçim | Seçenek frekansı, yüzde, %95 hata payı: ±1,96 × √(p(1−p)/n) |
| Çoklu seçim | Her seçeneği işaretleyenlerin soruyu yanıtlayanlara oranı (toplam %100'ü aşabilir) |
| Ölçek (ör. 1–5) | Ortalama, medyan, standart sapma, üst iki kutu (en yüksek iki puan), alt iki kutu, puan dağılımı |
| Sayı | Ortalama, medyan, standart sapma, en düşük, en yüksek |
| Açık uçlu | Kod çerçevesine göre tema frekansları, örnek yanıtlar, kodlanmayan yanıtlar |
| Demografi | Frekans; çapraz tablolarda kırılım olarak kullanılır |

- **Çapraz tablolar:** `--kirilim` ile verilen sütunlara (varsayılan: türü Demografi olan sorular) göre sütun
  yüzdeleri.
  - Tek seçim sorularında **ki-kare bağımsızlık testi** yapılır; p < 0,05 olanlar "Anlamlı Farklar" sayfasına
    yazılır ve en büyük yüzde farkı özetlenir.
  - Beklenen frekansı 5'in altında olan hücreler tablonun %20'sini aşarsa test "güvenilmez" olarak işaretlenir.
  - Cramér V, ilişkinin gücünü gösterir (0–1).
  - Ölçek sorularında grup ortalamaları ve n verilir.
- **Açık uçlu kodlama:** Yanıtta kod çerçevesindeki anahtar kelimelerden biri **kelime başında** geçiyorsa yanıt o
  kodu alır. Örneğin "fiyat" kelimesi "fiyatı" ve "fiyatlar" ile eşleşir. Bir yanıt birden çok kod alabilir.
  - Hiçbir koda girmeyen yanıtlar "Kodlanmayan Yanıtlar" sayfasına, elle kodlanmak üzere yazılır.
  - Örnek yanıtlarda e-posta adresleri ve cep telefonu numaraları maskelenir.
- **Veri kalitesi:**
  - Mükerrer yanıt numarası (ikinci satır alınmaz).
  - Seçenek listesinde veya ölçek aralığında olmayan değer (sayılmaz).
  - Boş oranı %20'yi aşan soru.
  - En az 4 ölçek sorusunun tamamına aynı puanı veren yanıtlar ("düz çizgi").

| Uyarı | Önem |
|---|---|
| Soru kodu yanıt dosyasında sütun olarak yok | Yüksek |
| Geçersiz değer; mükerrer yanıt no; kodlanmayan yanıt oranı %25'in üstünde; kırılım bulunamadı | Orta |
| Düz çizgi yanıt; küçük kırılım grubu (n < 30); boş oranı yüksek; küçük örneklem (n < 100) | Bilgi |

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek: 240 yanıt, 10 soru + 3 demografi
python main.py --yanitlar yanitlar.xlsx --sorular sorular.xlsx --kodlar kod_cercevesi.xlsx
python main.py --yanitlar y.xlsx --sorular s.xlsx --kirilim "Yaş Grubu" Cinsiyet --min-grup 30
```

| Dosya | Sütunlar |
|---|---|
| Yanıtlar | Her satır bir yanıt. Yanıt No ve soru listesindeki her kod için bir sütun. Çoklu seçimde seçenekler hücre içinde virgülle ayrılır. |
| Sorular | Soru Kodu (yanıt dosyasındaki sütun adı), Soru, Tür, Seçenekler (`|` ile ayrılmış; ölçekte `1|2|3|4|5`). Seçenek sırası tablolarda korunur. |
| Kod çerçevesi (isteğe bağlı) | Soru Kodu (tüm açık uçlular için `*`), Kod, Anahtar Kelimeler (virgülle) |

Form araçlarının (Google Forms, Microsoft Forms) çıktısında sütun başlıklarını soru kodlarıyla değiştirin.

## Çıktı

`anket_analizi.xlsx`:
- `Özet`: yanıt sayısı, genel hata payı, soru listesi ve boş oranları.
- `Frekanslar`: tek / çoklu seçim ve demografi soruları.
- `Ölçek Soruları`: ortalamaya göre sıralı; sayı soruları altta.
- `Çapraz Tablolar`: soru × kırılım; ki-kare, p, Cramér V.
- `Anlamlı Farklar`: p < 0,05 olan soru-kırılım çiftleri.
- `Açık Uçlu Kodlar`, `Kodlanmayan Yanıtlar` (boş "Elle Kod" sütunu), `Uyarılar`.

## Dikkat

- **Hata payı ve testler basit rastgele örneklem varsayar.** Kolayda örneklem, çevrimiçi panel veya kota
  örnekleminde sonuçlar yalnız yön göstericidir. Ağırlıklandırma yapılmaz.
- **Çoklu karşılaştırma:** Çok sayıda çapraz tablo test edildiğinde bazıları yalnız şansla anlamlı çıkar (%5
  düzeyinde 20 testten biri). Anlamlı farkları iş bilgisiyle değerlendirin.
- **Anahtar kelimeyle kodlama** ince anlamı, olumsuzlamayı ve ironiyi yakalayamaz ("fiyatı gayet iyi" de Fiyat
  koduna girer). Kodlar konu başlığıdır, memnuniyet yönü değildir. Önemli kararlar öncesinde yanıtları okuyun.
- **Kişisel veri:** Açık uçlu yanıtlar kişisel veri içerebilir. Maskeleme yardımcıdır, garanti değildir.
- **Örnek veri:** Örnek anket ve yanıtlar kurgusaldır.

## Testler

Şunlar test edilir:
- Ki-kare p değerinin kritik değer tablosuyla karşılaştırılması (p = 0,05 ve 0,01).
- 2×2 tabloda ki-kare, serbestlik derecesi, Cramér V (elle hesap); küçük beklenen frekans.
- Hata payı, kelime başı kodlama, soruya özel kod, maskeleme.
- Örnek veride frekans toplamları, çoklu seçim, geçersiz değer, mükerrer yanıt, düz çizgi, yaşa göre anlamlı fark.
- Excel ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
