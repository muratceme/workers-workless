# İş Tanımından İş Değerleme Puanlaması · AI Agent

> İnsan Kaynakları › Ücretlendirme ve Yan Haklar Uzmanı · Gerisi Bende

Serbest metin iş tanımlarını şirketinizin **puan faktör planına** göre değerlendirir:
- model her faktör için plandaki düzeylerden birini seçer ve iş tanımındaki kanıt cümlesini aynen yazar,
- **puanı, kademeyi ve mevcut kademeyle karşılaştırmayı kod yapar**.

Önerilen kademe, **Ücret Bandı Karşılaştırması** kod bloğunun çalışan dosyasındaki Kademe sütununa işlenebilir.

## Nasıl çalışır?

1. **Hazırlık (kod):** İş tanımları, değerleme planı (Faktör; Düzey; Tanım; Alt; Üst; Puan) ve kademe tablosu okunur.
   Alt / Üst aralığı olan faktör sayısal sayılır (deneyim yılı, bağlı çalışan, bütçe).
2. **Düzey seçimi (model):** Her faktör için düzey, kanıt ve gerekçe. Dayanak yoksa düzey boş bırakılır ve açık
   nokta olarak yazılır.
3. **Denetim (kod):** Aşağıdakilerden biri varsa yalnız o pozisyon modele geri gönderilir (varsayılan en fazla 2 tur):
   - Değerlendirilmeyen, planda olmayan ya da tekrar eden faktör; planda olmayan düzey
   - Kanıtın iş tanımında aynen geçmemesi
   - Sayısal faktörde kanıttaki en büyük değerin seçilen düzeyin aralığına düşmemesi ("12 milyon", "500 bin" çarpanları
     dahil). Kod doğru düzeyi de söyler.
   - Sayı gerektiren düzeyde kanıtta sayı olmaması ("rehberlik eder" bağlı çalışan sayılmaz)
   - Kanıtın yalnız "tercihen / tercih sebebi" yazan bir niteliğe dayanması
4. **Puan ve kademe (kod):**
   - Puanlar toplanır, kademe tablosundan kademe bulunur.
   - Mevcut kademeyle farklıysa "İnceleme" denir; kademe sınırına `--sinir` (varsayılan 10) puandan yakınsa not düşülür.
   - Düzeyi belirlenemeyen faktör varsa kademe hesaplanmaz.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                              # örnek 3 pozisyon, örnek plan
python agent.py --tanimlar is_tanimlari.docx --plan degerleme_plani.xlsx --kademeler kademeler.xlsx
```

İş tanımı örneği:

```
Pozisyon: Satış Müdürü | Departman: Satış | Mevcut Kademe: K5
Bölge satış ekibini (4 bölge müdürü ve 22 satış temsilcisi) yönetir. Yıllık 12 milyon TL satış ve pazarlama bütçesinden sorumludur.
Aranan nitelikler: Lisans mezunu, MBA tercih sebebi. Satışta en az 8 yıl ... deneyim.
```

## Çıktı

`cikti/is_degerleme.xlsx`:
- `Pozisyonlar`: faktör puanları, toplam, önerilen ve mevcut kademe, durum (Uyumlu / İnceleme / Eksik faktör / Yeni), not.
- `Faktörler`: seçilen düzey, iş tanımındaki kanıt ve gerekçe (mor: model).
- `Açık Noktalar`: iş tanımında eksik ya da belirsiz olan, yöneticiye sorulacak noktalar.
- `Değerleme Planı`: kullanılan plan ve kademe tablosu.
- `Kontrol`

## Dikkat

- **Örnek plan bir standart değildir.** `ornek_veri/degerleme_plani.csv` ve `kademeler.csv` yalnız gösterim içindir;
  faktörler, ağırlıklar ve kademe aralıkları şirketinizin iş değerleme sistemine göre belirlenmelidir.
- **Sonuç komite önerisidir.** Kademe değişikliği ücret kararıdır; iş değerleme komitesi ve yöneticiyle görüşülmeden
  uygulanmamalıdır.
- **İş tanımı kalitesi sonucu belirler.** Eksik ya da abartılı yazılmış iş tanımı yanlış puan verir; açık noktaları
  yöneticiyle tamamlayın.
- **Pozisyon değerlendirilir, kişi değil.** Çalışanın kişisel nitelikleri ve performansı bu puana girmez.
- **Örnek veri:** Pozisyonlar ve iş tanımları kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk turda Satış Müdürü'nün "en az 8 yıl" deneyimine "10+ yıl" düzeyini verir.
Kodun bunu yakalayıp doğru düzeyi (5–9 yıl) söylediği doğrulanır. Ayrıca şunlar doğrulanır:
- eksik / fazla faktör, planda olmayan düzey,
- aynen geçmeyen kanıt, sayısız kanıt ve "tercih" kuralı,
- "milyon / bin" çarpanları,
- puan, kademe, sınır notu ve eksik faktör durumu.

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
