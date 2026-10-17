# Performans Değerlendirme Özeti · AI Agent

> İnsan Kaynakları › İnsan Kaynakları Müdürü · Gerisi Bende

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

## Kaynaklar ve dayanaklar

Kontrol tarihi: 15.10.2026. Mevzuat sık değişir; paketteki süre, oran ve kuralları kullanmadan önce güncel resmî metinle
karşılaştırın. "Karşılaştırılan maddeler" dışındaki atıflar yalnız yol göstericidir.

| Kaynak | Bağlantı | Durum |
|---|---|---|
| 4857 sayılı İş Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.4857.pdf) | Resmî metinle karşılaştırılan maddeler: 17, 41, 46, 47, 53, 56, 63, 68, 69 |
| 6698 sayılı Kişisel Verilerin Korunması Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.6698.pdf) | Resmî metinle karşılaştırılan maddeler: 13 |
| 6701 sayılı Türkiye İnsan Hakları ve Eşitlik Kurumu Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.6701.pdf) | Resmî metin bağlantısı; paketteki madde atıfları ayrıca karşılaştırılmadı |

## Sorumluluk

Bu paketi indiren ve çalıştıran kişi aşağıdakileri kabul etmiş sayılır:

- **Tüm risk kullanıcıdadır.** Paket ücretsizdir ve "olduğu gibi" sunulur; doğruluk, eksiksizlik veya belirli bir amaca
  uygunluk konusunda hiçbir garanti verilmez. Paketi kendi bilgisayarınızda, kendi verinizle ve kendi kararınızla
  çalıştırırsınız. Hatalı veya eksik sonuç, kaçırılan süre, vergi / SGK cezası, veri kaybı, mali kayıp ve diğer her türlü
  doğrudan ya da dolaylı zarardan yazarlar ve katkıda bulunanlar sorumlu tutulamaz.
- **Sonuçları kontrol etmek sizin işinizdir.** Çıktılar karar desteği içindir. Beyan vermeden, ödeme yapmadan, sözleşme
  imzalamadan, personel / kredi / hasar / tedavi kararı almadan önce sonuçları yetkin bir kişi kontrol etmelidir.
- **Danışmanlık değildir.** Hukuki, mali, vergisel, tıbbi veya yatırım danışmanlığı yerine geçmez.
- **Mevzuat değişir.** Koddaki süre, oran, eşik ve kurallar yazıldığı tarihteki bilgiye dayanır ve eskiyebilir; özel
  durumları kapsamayabilir. Güncelliğini kontrol etmek kullanıcının sorumluluğundadır.
- **Örnek veri kurgusaldır.** Kişi, kurum, tutar ve tarihler gerçek değildir; kendi verinizle sonuçlar farklı olur.
- **Kişisel veri.** Verinizin KVKK, şirket politikaları ve sektör düzenlemelerine uygun işlenmesinden, saklanmasından ve
  paylaşılmasından siz sorumlusunuz.
- **Yapay zekâ hata yapabilir.** Model yanlış, eksik veya uydurma bilgi üretebilir. Kod kontrolleri bazı hataları
  yakalar, hepsini yakalayamaz. Her taslağı okuyup onaylamadan kullanmayın, göndermeyin, imzalamayın.
- **Maskeleme sınırlıdır.** Yalnız telefon, e-posta, TCKN, IBAN ve LinkedIn adresi kalıpları maskelenir. Ad soyad,
  adres, doğum tarihi, sağlık bilgisi ve serbest metindeki diğer kişisel veriler modele gidebilir. Hangi verinin bulut
  sağlayıcısına gönderildiğinden siz sorumlusunuz; gizli veride yerel model (Ollama) kullanın.
- **API ücretleri, kotalar ve API anahtarının güvenliği** size aittir.

Ayrıntılar: `LICENSE` (MIT) ve `SORUMLULUK_REDDI.md`.
