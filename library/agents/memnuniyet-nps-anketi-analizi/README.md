# Memnuniyet (NPS) Anketi Analizi · AI Agent

> Müşteri Hizmetleri ve Çağrı Merkezi › Müşteri Deneyimi Uzmanı · Workers / Workless

NPS ve memnuniyet anketlerinin puanlarını ve yorumlarını analiz eder. NPS'yi **hata payıyla** verir, segment ve ay
bazında karşılaştırır, memnuniyet ölçütlerinden **öncelikli iyileştirme alanlarını** bulur. Yorumları temalara
ayırır ve yorumlara dayanan **iyileştirme önerileri** çıkarır.

## Nasıl çalışır?

1. **Kod (puanlar):**
   - **NPS** = destekçi (9–10) yüzdesi − kötüleyen (0–6) yüzdesi. 7–8 puan verenler pasiftir.
   - **Hata payı (%95):** 1,96 × √((p_d + p_k − (p_d − p_k)²) / n) × 100. Örnek: 100 yanıtta %50 destekçi ve %20
     kötüleyen → NPS +30 ± 15.
   - Segment, kanal, bölge ve ürün sütunları varsa her biri için ayrı NPS; ay bazında eğilim. Yanıtı `--min-yanit`
     (30) altındaki gruplar "Yanıt yetersiz" olarak işaretlenir.
   - **Ölçütler** (1–5 puanlı sütunlar): ortalama, üst iki kutu, alt iki kutu ve tavsiye puanıyla korelasyon.
     - Korelasyonu medyanın üstünde, ortalaması medyanın altında olan ölçüt **öncelikli iyileştirme** alanıdır.
     - Korelasyonu yüksek, ortalaması yüksek olan ölçüt güçlü yöndür.
2. **Model (yorumlar):**
   - Yorumlar maskelenerek gönderilir: e-posta, telefon, TCKN, IBAN. Müşteri adı ve segment gönderilmez; yalnız
     yorum metni ve puan gider.
   - Model her yorumu sabit tema listesine göre etiketler, duyguyu belirler ve yorumdan birebir kısa bir alıntı
     verir. Kod alıntının yorumda geçtiğini doğrular; geçmiyorsa alıntıyı çıkarır. Listede olmayan tema atılır.
   - Puanla duygu çelişen yorumlar işaretlenir (ör. 10 puan + olumsuz yorum; ölçek ters anlaşılmış olabilir).
3. **Kod + model (öneriler):**
   - Kod tema başına yorum, olumsuz yorum, kötüleyen ve pasif sayılarını hesaplar. Model bu özetten 3–6 öneri yazar.
   - Her öneri yorum kimliklerine dayanmalıdır. Geçerli kimliğe dayanmayan öneri rapora alınmaz. Rapordaki sayılar
     ve alıntılar koddan gelir.

`--yorum-yok` ile API kullanılmaz; yalnız puan analizi yapılır.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                              # örnek: 180 yanıt, 3 segment, 52 yorum
python agent.py --anket nps.xlsx --min-yanit 30
python agent.py --anket nps.xlsx --temalar "Ürün kalitesi" Fiyat Teslimat "Müşteri hizmetleri" "Mobil uygulama"
python agent.py --anket nps.xlsx --yorum-yok                 # yalnız puan analizi, API gerekmez
```

| Sütun | Açıklama |
|---|---|
| NPS / Tavsiye | 0–10 tam sayı (zorunlu) |
| Yanıt No, Tarih | İsteğe bağlı; tarih varsa aylık eğilim çıkar |
| Segment, Kanal, Bölge, Ürün | İsteğe bağlı kırılımlar |
| Diğer sayısal sütunlar | 1–5 puanlı memnuniyet ölçütleri; sütun adı ölçüt adıdır |
| Yorum / Görüş | Açık uçlu yorum |

## Çıktı

`nps_analizi.xlsx`:
- `Özet`: genel, kırılım ve ay bazında NPS, hata payı, destekçi / pasif / kötüleyen oranları.
- `Ölçütler`: ortalama, kutu oranları, korelasyon, konum (öncelikli iyileştirme, güçlü yön...).
- `Öneriler` (model): sorun, öneri, dayanak yorum sayısı, kötüleyen sayısı, doğrulanmış alıntılar; boş "Sorumlu /
  Karar" sütunu.
- `Temalar`: tema başına yorum, duygu ve müşteri grubu sayıları.
- `Yorumlar`: maskeli metin, temalar, duygu, alıntı, kontrol notları.
- `Uyarılar`.

Mor hücreler yapay zekâ tarafından yazılmıştır.

## Dikkat

- **Hata payını dikkate alın.** Küçük gruplarda NPS birkaç yanıtla çok değişir. İki grubun veya iki ayın aralıkları
  kesişiyorsa fark gerçek olmayabilir.
- **Korelasyon nedensellik değildir.** Öncelikli alanlar bir başlangıç noktasıdır; yorumlar ve operasyon verisiyle
  doğrulayın.
- **Model çıktısı taslaktır.** Temalar ve öneriler insan kontrolünden geçmelidir; aksiyon kararı ekibindir.
- **Kişisel veri:** Maskeleme yardımcıdır, garanti değildir. Yorumların API'ye gönderilmesinin KVKK açısından
  uygunluğu (aydınlatma, yurt dışına aktarım) sizin sorumluluğunuzdadır. Yerel model için `WW_PROVIDER=ollama`
  kullanabilirsiniz.
- **Örnek veri:** Örnek anket yanıtları kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Şunlar test edilir:
- NPS ve hata payı formülü (elle hesap), Pearson korelasyonu.
- Geçersiz puan, segment toplamları, öncelikli iyileştirme alanı, aylık eğilim.
- Maskeleme ve segmentin gönderilmemesi; uydurma tema, metinde olmayan alıntı, geçersiz kimlik ve dayanaksız öneri.
- Puan–yorum çelişkisi, Excel ve CLI.

```bash
python -m unittest discover -s tests
```

## Kaynaklar ve dayanaklar

Kontrol tarihi: 15.10.2026. Mevzuat sık değişir; paketteki süre, oran ve kuralları kullanmadan önce güncel resmî metinle
karşılaştırın. "Karşılaştırılan maddeler" dışındaki atıflar yalnız yol göstericidir.

| Kaynak | Bağlantı | Durum |
|---|---|---|
| 6698 sayılı Kişisel Verilerin Korunması Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.6698.pdf) | Resmî metinle karşılaştırılan maddeler: 13 |

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
