# Teftiş Raporu Bulgu Taslağı · AI Agent

> Bankacılık › Teftiş Kurulu › Müfettiş · Workers / Workless

Çalışma kâğıdındaki tespitlerden **teftiş raporu bulgu metinlerinin taslağını** hazırlar. Her bulgu şu
bölümlerden oluşur: durum (tespit), kriter (mevzuat dayanağı), neden, etki/risk, risk düzeyi ve öneriler.

**Dayanak yalnız sizin verdiğiniz metinlerden gelir.** Model ezberden mevzuat yazmaz. Her alıntının gösterilen
maddede birebir geçtiği kodla doğrulanır; doğrulanamayan dayanak işaretlenir.

> **Önemli — gizlilik:** Teftiş çalışma kâğıtları bankacılık sırrı ve kişisel veri içerir. Dış bir yapay zekâ
> servisine göndermeden önce bilgi güvenliği ve uyum birimlerinizin onayını alın. Veri bilgisayardan çıkmasın
> isteniyorsa `.env` içinde `WW_PROVIDER=ollama` ile yerel model kullanın.

## Nasıl çalışır?

1. **Okur (kod):**
   - Çalışma kâğıdı: `Tespit` sütunu olan `.csv` / `.xlsx`.
   - Kriter metinleri: `mevzuat/` klasöründeki veya adında "yönetmelik", "yönerge", "prosedür", "genelge",
     "tebliğ" geçen `.txt` / `.pdf` / `.docx` dosyaları. Bu metinler `Madde N` başlıklarından maddelere bölünür;
     madde başlığı yoksa metin tek parça kullanılır.
2. **Kontrol eder (kod):**

   | Kontrol | Önem |
   |---|---|
   | Tespit numarası tekrar ediyor; hatalı adet örneklemden büyük | yüksek |
   | Kanıt / çalışma kâğıdı referansı yok | orta |
   | Kriter metni verilmedi | orta |
   | Örneklem yazılmamış (hata oranı hesaplanamıyor) | bilgi |
   | Aynı birim ve konuda birden çok tespit (birleştirme adayı) | bilgi |

   - Hata oranı = hatalı / örneklem.
   - Madde sayısı 60'tan fazlaysa her tespit için ortak kelimesi en çok olan 4 madde seçilip gönderilir.
3. **Maskeler:**
   - "Personel" ve "Müşteri" sütunlarındaki adlar `[PERSONEL-n]` ve `[MÜŞTERİ-n]` olur. Birden çok ad `|` ile
     ayrılır.
   - Telefon, e-posta, TCKN ve IBAN maskelenir; ek adları `--gizle` ile verin.
   - Veri gönderilmeden önce onayınız alınır.
4. **Yazar (model):** resmî rapor dilinde bulgular yazar; benzer tespitleri birleştirebilir.
   - Neden: yalnız müfettiş notunda veya birim açıklamasında yazan neden aktarılır; yoksa "birimle görüşülerek
     belirlenmeli" yazılır.
   - Risk düzeyi (Yüksek / Orta / Düşük) gerekçesiyle önerilir; nihai karar müfettişindir.
5. **Denetler (kod):**

   | Durum | Sonuç |
   |---|---|
   | Alıntı gösterilen maddede birebir geçmiyor | "Alıntı doğrulanamadı" |
   | Madde verilen metinlerde yok | "Madde bulunamadı" |
   | Bulgunun doğrulanmış dayanağı yok | "DAYANAK MÜFETTİŞÇE EKLENMELİ" |
   | Bir tespit hiçbir bulguda kullanılmamış | uyarı |
   | Tespitlerde olmayan numara | yok sayılır |
   | Metinde, girdilerde olmayan sayı | işaretlenir |

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                    # örnek: kurgusal şube teftişi ve kurgusal yönerge
python agent.py --girdi ./teftis_klasoru
python agent.py --girdi ./teftis_klasoru --gizle "Ek Kişi Adı"
```

```
teftis_klasoru/
├── calisma_kagidi.xlsx
└── mevzuat/
    ├── sube_operasyon_yonergesi.pdf
    └── kart_proseduru.docx
```

| Çalışma kâğıdı sütunu | Açıklama |
|---|---|
| Tespit No, Birim, Konu, Tespit | Tespit zorunlu; diğerleri önerilir |
| Örneklem, Hatalı, Tutar | Hata oranı ve önem için |
| Kanıt | Çalışma kâğıdı / belge referansı |
| Personel, Müşteri | Maskelenir; birden çoksa `\|` ile |
| Kök Neden, Birim Açıklaması | Neden bölümüne yalnız buradan aktarılır |

## Çıktı

- **Taslak:** `bulgu_taslaklari.md`. Takma adlar geri açılmıştır; her bulguda boş bir "Birim cevabı" alanı
  bulunur.
- **Excel:** `bulgu_taslaklari.xlsx`:
  - `Özet`.
  - `Bulgular`: tüm bölümler; boş "Müfettiş Onayı" ve "Birim Cevabı" sütunları.
  - `Tespitler`: hata oranı ve hangi bulguda kullanıldığı.
  - `Dayanak Kontrolü`: her alıntı ve doğrulama sonucu.
  - `Kontroller`.

## Dikkat

- **Taslak niteliği:** Bulgu metinleri, risk düzeyleri ve öneriler taslaktır; müfettiş tarafından gözden
  geçirilmeli ve teftiş yönergenizdeki formata uyarlanmalıdır.
- **Güncel mevzuat:** Doğrulama, alıntının **verdiğiniz metinde** geçtiğini gösterir; metnin güncel ve konuyla
  ilgili olduğunu göstermez. Güncel mevzuatı resmî kaynaklardan (Resmî Gazete, mevzuat.gov.tr, BDDK) ve güncel iç
  düzenlemelerden alın.
- **Örnek veri:** Örnek şube, kişiler ve "Şube Operasyon Yönergesi" kurgusaldır; gerçek bir bankanın düzenlemesi
  değildir.

## Testler

Testler gerçek API çağırmaz. Şunlar test edilir:
- Çalışma kâğıdı ve madde ayrıştırma; madde seçimi; kod kontrolleri.
- Kişi adlarının modele gitmemesi.
- Doğru alıntı, değiştirilmiş alıntı ve uydurma madde ayrımı.
- Uydurma tespit numarası, kullanılmayan tespit ve uydurma sayı (190.000 TL).

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. API kullanım ücretleri ve modele gönderilen verinin hukuka ve kurum politikasına
uygunluğu kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
