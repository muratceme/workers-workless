# Performans Değerlendirme Özeti · AI Agent

> İnsan Kaynakları › İnsan Kaynakları Müdürü · Workers / Workless

Yöneticilerin performans değerlendirme formlarını (hedef ve yetkinlik kalemleri, puanlar, yorumlar) **çalışan başına
kısa bir özete** ve **dağılım raporuna** çevirir. Rapor hem geri bildirim görüşmesine hem de İK kalibrasyon
toplantısına hazırlık içindir.

## Nasıl çalışır?

1. **Kod, puanları hesaplar (yapay zekâ yok):**
   - **Puanlar:** Hedef puanı ve yetkinlik puanı (kalem ağırlıklarıyla), genel puan (varsayılan %60 hedef, %40
     yetkinlik) ve performans kategorisi.
   - **Dağılım:** Şirket ve departman bazında hesaplanır.
   - **Değerlendirici eğilimi:** Yönetici ortalaması şirket ortalamasından ±0,5 farklıysa cömert veya katı
     değerlendirme sinyali verilir. 4 ve üzeri çalışanda puanlar dar bir aralığa yığılmışsa ortada yığılma
     işaretlenir.
   - **Öz değerlendirme farkı:** Aynı kalemde yönetici puanıyla 1,5 ve üzeri fark varsa görüşmede ele alınması
     önerilir.
   - **Veri hataları:** Ağırlık toplamının 100 olmaması ve puanlanmamış kalemler raporlanır.
2. **Kod, yorumları tarar:** Yorumlar ortak ayrımcılık tarayıcısından geçer. **Korunan bir özelliği** gerekçe
   gösteren yorumlar "değerlendirme dışı" olarak işaretlenir. Bu özellikler yaş ("yaşı gereği"), gebelik ve doğum
   izni, aile durumu ("çocuklu olduğu için"), sağlık ("sağlık sorunları nedeniyle") ve dindir (İş K. md. 5, 6701
   s. Kanun).
3. **Model, özeti yazar:** Çalışan ve yönetici adları modele takma adla (P1, Y1) gider. Model yorumlara dayanarak
   şunları yazar:
   - güçlü yönler,
   - gelişim alanları,
   - uygulanabilir gelişim önerileri,
   - görüşmede konuşulacaklar.

   **Puan ve kategori modelden değil koddan gelir.**
4. **Kod, model metnini denetler:** Model metninde korunan özelliğe dayalı ifade kalırsa uyarı verilir. Rapor
   şirket içi olduğu için takma adlar raporda gerçek adlara çevrilir.

| Kategori | Genel puan (1–5) |
|---|---|
| Üstün | ≥ 4,5 |
| Beklentinin üzerinde | ≥ 3,8 |
| Beklentiyi karşılıyor | ≥ 3,0 |
| Gelişime açık | ≥ 2,0 |
| Beklentinin altında | < 2,0 |

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                   # örnek: 7 çalışan, 3 yönetici
python agent.py --girdi degerlendirmeler.xlsx --hedef-agirligi 70
```

Girdi, performans sisteminden alınan uzun biçimli bir dökümdür. Sütunlar: Sicil, Ad Soyad, Departman, Yönetici,
Tür (Hedef/Yetkinlik), Kalem, Ağırlık, Puan (1–5), Öz Puan (isteğe bağlı) ve Yorum.

## Çıktı

- **Çalışan Özeti:** Puanlar, kategori, güçlü yönler, gelişim alanları, öneriler, görüşme notları, uyarılar ve
  "İK Onayı" sütunu.
- **Dağılım:** Grafik ve departman kırılımı.
- **Değerlendiriciler:** Değerlendirici eğilimleri.
- **Kalem Detayı:** Kalem bazında puanlar ve işaretlenen yorumlar.
- **Özet Kartları:** Yazdırılabilir, çalışan başına bir blok.
- **Bilgi:** Hesap kuralları ve uyarılar.

## Dikkat

- **Taslak niteliği:** Özetler taslaktır. İK ve yönetici onayı olmadan çalışana iletilmemelidir. Ücret, terfi veya
  işten çıkarma kararlarında tek başına kullanılmamalıdır.
- **Değerlendirici eğilimi:** Bu tablo istatistiksel bir sinyaldir; bir yöneticinin hatalı olduğunu göstermez.
- **Kişisel veri:** Performans verisi kişisel veridir. Modele gönderim için KVKK kapsamında aydınlatma ve hukuki
  dayanağı değerlendirin. Ollama ile yerel model kullanırsanız veri bilgisayarınızdan çıkmaz.

## Testler

Testler gerçek API çağırmaz. Kontrol edilenler:
- ağırlıklı puanların elle hesapla tutması,
- ağırlık hatası,
- önyargılı yorumların yakalanması,
- öz değerlendirme farkı ve değerlendirici eğilimi,
- adların modele gitmemesi,
- sahte modelin yazdığı korunan özellik ifadesinin kod tarafından yakalanması.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. API kullanım ücretleri ve modele gönderilen verinin hukuka uygunluğu
kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
