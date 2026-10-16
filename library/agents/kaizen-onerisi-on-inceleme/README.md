# Kaizen Önerisi Ön İnceleme · AI Agent

> Endüstri Mühendisliği › Endüstri Mühendisi · Workers / Workless

Çalışanların serbest metinle yazdığı iyileştirme önerilerini **değerlendirmeye hazır kayda** çevirir. Örnek öneri:
"Kalıp değişimi 25 dakika kısalır, günde 4 kez, işi 2 kişi yapıyor". Bu metinden konu başlığı, kategori, kazanç
bileşenleri, uygulama maliyeti, uygulanabilirlik önerisi, benzer öneriler ve önerene sorulacak eksik bilgiler çıkar.
**Yıllık kazancı model değil kod hesaplar.** Çıktı **Kaizen Önerisi Değerlendirme** kod bloğuna doğrudan girer.

## Nasıl çalışır?

1. **Hazırlık (kod):** Öneri listesi okunur. Öneren adları sağlayıcıya gönderilmez (`[ÖNEREN]`); telefon, e-posta ve
   kimlik numarası maskelenir.
2. **İnceleme (model):** Her öneri için şunlar yazılır:
   - konu başlığı ve kategori (İSG, Kalite, Verimlilik, Maliyet, Enerji, 5S / Ergonomi),
   - metinde yazan kazanç bileşenleri:
     - zaman: dakika / saat × tekrar × kişi × periyot,
     - malzeme: miktar × birim fiyat × periyot,
     - doğrudan TL,
   - uygulama maliyeti,
   - uygulanabilirlik (1-5) ve gerekçesi,
   - aynı sorunu öneren diğer öneriler,
   - eksik bilgiler.
3. **Denetim (kod):** Aşağıdakilerden biri varsa yalnız o öneri modele geri gönderilir (varsayılan en fazla 2 tur):
   - Bir sayı öneri metninde yazmıyor. Rakamsız bir öneriye "5 dakika" gibi tahmin yazmak kabul edilmez.
   - Birim, periyot ya da tür tutarsız (ör. kg cinsinden zaman); malzeme kazancında birim fiyat yok
   - Benzer öneri numarası listede yok; uygulanabilirlik 1-5 dışında; gerekçe ya da konu boş
4. **Kazanç (kod):**

| Bileşen | Hesap |
|---|---|
| Zaman | dakika × tekrar × kişi × yıllık periyot × saat maliyeti ÷ 60 |
| Malzeme | miktar × tekrar × birim fiyat × yıllık periyot |
| Doğrudan TL | tutar × tekrar × yıllık periyot |

Yıllık periyot katsayıları: gün = `--gun` (varsayılan 250); hafta = gün ÷ 5; ay = 12; yıl = 1; vardiya = gün ×
`--vardiya`. Kazanç hesaplanamayan öneriler "bilgi isteyin" olarak işaretlenir. İSG önerileri kazançtan bağımsız
değerlendirilmek üzere ayrılır. Uygulama maliyeti yazılmamışsa uyarı verilir: değerlendirme bloğu boş maliyeti 0
sayar, geri ödeme puanı olduğundan yüksek çıkar.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                    # örnek 6 öneri (saat maliyeti 300 TL varsayılan)
python agent.py --oneriler oneriler.xlsx --saat-maliyeti 420 --gun 255 --vardiya 3
```

Ardından puanlama için:

```bash
python main.py --oneriler cikti/kaizen_on_inceleme.xlsx     # Kaizen Önerisi Değerlendirme kod bloğu klasöründe
```

**Saat maliyetini kendi şirketinize göre verin.** Varsayılan 300 TL yalnız örnek içindir; brüt ücret, işveren payları
ve yan haklar dahil saatlik maliyetinizi kullanın.

## Çıktı

`cikti/kaizen_on_inceleme.xlsx`:
- `Öneriler`:
  - İlk sütunlar **Kaizen Önerisi Değerlendirme** girdi sütunlarıyla aynı: Öneri No, Öneren, Bölüm, Konu, Yıllık Kazanç,
    Uygulama Maliyeti, Uygulanabilirlik.
  - Diğer sütunlar: kategori, kazanç hesabı (formül satırları), uygulanabilirlik gerekçesi, benzer öneriler, eksik
    bilgiler, durum, öneri metni ve son kontrol.
- `Kontrol`: kullanılan parametreler, düzeltme turu, son denetim.

## Dikkat

- **Kazanç önerenin beyanına dayanır.** "25 dakika kısalır" bir tahmindir; uygulamadan önce zaman etüdü veya deneme ile
  doğrulayın (**Zaman Etüdü ve Standart Süre** kod bloğu).
- **Uygulanabilirlik ve kategori yapay zekâ önerisidir.** Komite karar vermeden önce gözden geçirmelidir.
- **İSG önerileri** kazanç puanıyla sıralanmamalıdır. Ramak kala içeren önerileri İSG kaydına da işleyin.
- **Örnek veri:** Öneriler ve kişiler kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model, rakam içermeyen vida rafı önerisine ilk turda "5 dakika" kazanç uydurur.
Kodun bunu yakalayıp yalnız o öneriyi yeniden yazdırdığı doğrulanır. Ayrıca şunlar doğrulanır:
- zaman / malzeme / TL ve periyot (gün, hafta, vardiya) hesapları,
- İSG, benzer öneri ve maliyet yok uyarıları,
- öneren adlarının hiçbir mesajda geçmemesi.

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
