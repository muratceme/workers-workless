# Hasta Şikâyet Sınıflandırma · AI Agent

> Hasta Hizmetleri › Hasta İlişkileri ve Hasta Hakları Yöneticisi · Workers / Workless

Hasta ve yakınlarından gelen şikâyet, öneri ve teşekkürleri **birim, konu, tür ve önem** derecesine göre sınıflandırır.
Her başvuru için ilgili birime iletilecek inceleme sorusunu yazar ve cevap süresini takip eder.

## Nasıl çalışır?

1. **Hazırlık (kod):**
   - Satırdaki hasta adı metinden çıkarılır ([HASTA]); telefon, e-posta ve TCKN maskelenir.
   - Önem kuralları uygulanır. Şu belirtilerden biri olan başvuru "Yüksek" önemdedir:
     - Şiddet, tehdit ya da hakaret
     - Hasta güvenliği olayı: düşme, yanlış ilaç, yanlış taraf, enfeksiyon, yanık
     - Mahremiyet ihlali
2. **Sınıflandırma (model):** Birim (verilen listeden), konu, tür, önem, tarafsız özet, alıntı ve birime inceleme sorusu.
3. **Denetim (kod):** Aşağıdakilerden biri varsa yalnız o başvuru modele geri gönderilir (varsayılan en fazla 2 tur):
   - Birim listede yok
   - Alıntı metinde yok
   - Kural gereği "Yüksek" olması gereken başvuruya daha düşük önem verilmiş
4. **Süre takibi (kod):**
   - Son cevap tarihi = başvuru tarihi + `--sure` gün.
   - Açık başvurularda kalan gün; `--uyari` gün içinde dolanlar ve gecikenler.
   - Kapalı başvurularda zamanında cevap verilip verilmediği.
   - Kod bulgusuna göre yapılacaklar: çalışan güvenliği (beyaz kod) prosedürü, hasta güvenliği olay bildirimi, mahremiyet
     ihlali için KVKK sorumlusuna iletme.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                   # örnek başvurular + birim listesi
python agent.py --girdi basvurular.xlsx --birimler birimler.txt
python agent.py --girdi basvurular.xlsx --birimler birimler.txt --sure 15 --uyari 3 --bugun 15.10.2026
```

## Çıktı

- `cikti/hasta_basvurulari.xlsx`:
  - `Başvurular`: önem ve kalan güne göre sıralı; "Sorumlu" sütunu boş bırakılır.
  - `Birim × Konu`: şikâyet matrisi ve başvuru türü dağılımı.

## Dikkat

- **Cevap süresi.** Hastane başvuruları için tüm kurumlarda geçerli tek bir süre yoktur; kurumlar kendi hasta hakları
  prosedürlerinde süre belirler. Varsayılan 30 gün, 3071 sayılı Dilekçe Hakkının Kullanılmasına Dair Kanun'daki genel cevap
  süresidir. `--sure` değerini kurumunuzun prosedürüne ve başvurunun geldiği kanala (CİMER, SABİM vb.) göre ayarlayın.
- **Sınıflandırma öneridir.** Hasta hakları birimi her başvuruyu okuyarak karar verir. Bu araç tıbbi değerlendirme ya da
  kusur tespiti yapmaz, cevap metni yazmaz.
- **Sağlık verisi.** Başvurular KVKK kapsamında özel nitelikli kişisel veri içerebilir. Hasta adı metinden çıkarılır;
  metindeki diğer kişi adları (yakın, personel) ve sağlık bilgileri gönderilir. Kurumunuzun izni olmadan bulut tabanlı bir
  modele göndermeyin; gerekirse yerel model (Ollama) kullanın.
- **Örnek veri** kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk taslakta bilerek yataktan düşme olayına "Orta" önem verir, listede olmayan bir
birim yazar ve metinde olmayan bir alıntı kullanır. Kodun bunları yakalayıp yalnız bu başvuruları yeniden sorduğu, hasta
adlarını modele göndermediği ve süreleri doğru hesapladığı doğrulanır.

```bash
python -m unittest discover -s tests
```

## Kaynaklar ve dayanaklar

Kontrol tarihi: 15.10.2026. Mevzuat sık değişir; paketteki süre, oran ve kuralları kullanmadan önce güncel resmî metinle
karşılaştırın. "Karşılaştırılan maddeler" dışındaki atıflar yalnız yol göstericidir.

| Kaynak | Bağlantı | Durum |
|---|---|---|
| 6698 sayılı Kişisel Verilerin Korunması Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.6698.pdf) | Resmî metinle karşılaştırılan maddeler: 13 |
| 3071 sayılı Dilekçe Hakkının Kullanılmasına Dair Kanun | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.3071.pdf) | Resmî metin bağlantısı; paketteki madde atıfları ayrıca karşılaştırılmadı |

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
