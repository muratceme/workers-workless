# Rapor Yorum ve Aksiyon Notu · AI Agent

> Yönetim raporları · Gerisi Bende

Bir Excel raporundan yönetici için **kısa bir yorum notu** yazar: özet, önem dereceli bulgular ve bulgulara bağlı aksiyon
önerileri. Rapor sitedeki bir kod bloğunun çıktısı olabilir (satış hedef-gerçekleşme, gecikme yaşlandırma, muallak
hasar, bütçe sapma, OEE…) ya da kendi raporunuz. **Nottaki her rakam raporla kod tarafından karşılaştırılır.**

## Nasıl çalışır?

1. **Okuma (kod):**
   - Raporun bütün sayfaları okunur. Hücreler Excel'de göründükleri gibi yazılır: %55,6 · 967.029,08 · 15.10.2026.
   - 40 satırı geçen tablolarda ilk 25 satır gönderilir. Her sayısal sütun için bütün satırların toplamı, ortalaması, en
     büyüğü ve en küçüğü (hangi satırda olduğuyla) kodla hesaplanıp eklenir.
   - `--takma-ad` ile verilen sütunlardaki değerler (ör. temsilci, müşteri adları) `Kişi-1`, `Kişi-2` … olarak gönderilir
     ve notta gerçek adlarla geri açılır.
   - Telefon, e-posta, TCKN ve IBAN maskelenir.
2. **Yorum (model):** Okuyucuya ve odak konusuna göre şunlar yazılır:
   - en fazla 5 maddelik özet,
   - en fazla 8 bulgu (Kritik / Önemli / Bilgi; her biri bir sayfaya dayanır),
   - aksiyon önerileri (sorumlu görev / birim, ilgili bulgu),
   - yorumu sınırlayan veri notları.
3. **Denetim (kod):** Aşağıdakilerden biri varsa not modele geri gönderilir (varsayılan en fazla 2 tur):
   - Notta raporda olmayan bir sayı var. Yuvarlama kabul edilir: 967.029,08 → 967.029; %55,6 → %56. Raporda olmayan
     fark, oran veya toplam kabul edilmez.
   - Bulgunun sayfası raporda yok; aksiyon geçerli bir bulguya bağlı değil.
   - Takma ada ek getirilmiş ("Kişi-4'ün"): gerçek ad geri konunca ek yanlış olur.
   - Maskeli bir bilgi kullanılmış; özet maddesi veya bulgu sayısı sınırın dışında.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                     # örnek satış hedef-gerçekleşme raporu
python agent.py --rapor cikti/hedef.xlsx --okuyucu "Bölge Satış Müdürü" --odak "hedefin gerisinde kalanlar" --takma-ad Temsilci
python agent.py --rapor ornek_veri/kredi_gecikme_yaslandirma_raporu.xlsx --okuyucu "Şube Müdürü" --takma-ad Müşteri
```

Örnek raporlar sitedeki **Satış Hedef-Gerçekleşme Raporu** ve **Gecikme Yaşlandırma Raporu** kod bloklarının kendi örnek
verileriyle üretilmiştir.

## Çıktı

- `cikti/rapor_notu.xlsx`:
  - `Not`: gönderime hazır not.
  - `Bulgular`: önem ve dayandığı sayfa.
  - `Aksiyonlar`: sorumlu, ilgili bulgu, Durum ve Termin sütunlarıyla; takip listesi olarak kullanılabilir.
  - `Kontrol`: düzeltme turu ve son denetim sonucu.
- `cikti/rapor_notu.txt`: aynı not, e-postaya yapıştırmak için.

## Dikkat

- **Rakamlar doğrulanır, yorum doğrulanmaz.** Kod, nottaki sayıların raporda olduğunu kontrol eder; o sayının doğru
  bağlamda kullanıldığını (ör. doğru temsilciye ait olduğunu) her zaman anlayamaz. Notu göndermeden önce okuyun.
- **Raporun kendisi kontrol edilmez.** Hatalı rapor hatalı not üretir.
- **Uzun tablolar** için model ilk 25 satırı ve kodun sütun özetlerini görür; bütün satırlara dair yorum yapamaz.
  Çok sayfalı büyük raporlarda yalnız ilgili sayfaları içeren bir kopya verin (sınır 60.000 karakter).
- **Kişisel veri.** Rapordaki bütün metin modele gider. Kişi ve müşteri adı sütunları için `--takma-ad` kullanın.
  Takma ad yalnız verilen sütunlara uygulanır; başlıklarda veya açıklamalarda geçen adları maskelemez.
- **Performans değerlendirmesi değildir.** Not, kişiler hakkında yargı yerine rakam ve hedefe göre durum yazar.
- **Örnek veri:** Raporlardaki kişiler, şubeler ve müşteriler kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk notta bilerek:
- raporda olmayan bir fark hesaplar (ay hedefine kalan 772.971 TL),
- takma ada ek getirir ("Kişi-4'ün").

Kodun bu 4 sorunu yakalayıp modele geri gönderdiği, yuvarlanmış rakamları ise kabul ettiği doğrulanır. Ayrıca şunlar
doğrulanır:
- hücre biçimleme (yüzde, binlik, tarih),
- uzun tablolarda kodun sütun özetleri,
- takma adların modele gitmemesi ve notta geri açılması.

```bash
python -m unittest discover -s tests
```

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
