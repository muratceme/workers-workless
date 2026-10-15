# Müşteri Şikâyeti Kök Neden Analizi · AI Agent

> Kalite Güvence ve Kalite Kontrol › Kalite Mühendisi · Workers / Workless

Müşteri şikâyetlerini **ürün, parti ve hata türüne** göre kümeler; her küme için 6M kategorisinde olası kök neden
hipotezleri, doğrulama adımları ve geçici önlemler önerir. Çıktı, 8D raporunun D3 (geçici önlem) ve D4 (kök neden)
adımları için başlangıç noktasıdır.

## Nasıl çalışır?

1. **Kümeleme (kod):**
   - Parti kümesi: aynı ürün + partide `--esik` ve üstü şikâyet (varsayılan 2).
   - Ürün kümesi: aynı üründe `--urun-esik` ve üstü farklı partide şikâyet (varsayılan 3). Bu, tek partiye özgü değil
     sistematik (tasarım, kalıp, tedarikçi, süreç) bir sorun olabileceğini gösterir.
2. **Analiz (model):** Her şikâyete hata türü verilir. Her küme için ortak bulgu ve 1–4 hipotez yazılır; her hipotezde 6M
   kategorisi, dayanak şikâyetler, doğrulama adımı ve geçici önlem bulunur.
3. **Denetim (kod):** Aşağıdakilerden biri varsa taslak modele geri gönderilir (varsayılan en fazla 2 tur):
   - Sınıflandırılmamış şikâyet ya da açıklamada olmayan alıntı
   - Hipotezi olmayan küme
   - Hipotezin dayanağı o kümede olmayan şikâyet
   - Doğrulama adımı yazılmamış hipotez
   - Kesin sonuç gibi yazılmış hipotez ("kesinlikle", "kanıtlanmıştır")
4. **Rapor (kod):** Hata türü Pareto'su (şikâyet ve etkilenen adet) ve ürün × parti matrisi.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                         # örnek şikâyetler
python agent.py --girdi sikayetler.xlsx
python agent.py --girdi sikayetler.xlsx --esik 3 --urun-esik 4 --tur 3
```

## Çıktı

- `cikti/sikayet_kok_neden.xlsx`:
  - `Kümeler ve Hipotezler`: hipotezlerin yanında "Sonuç (doğrulandı / elendi)" sütunu.
  - `Şikâyetler`: hata türü, alıntı ve küme.
  - `Pareto`: tabloyla birlikte grafik.
  - `Ürün × Parti`: aynı partide 2 ve üstü şikâyet kırmızı.

## Dikkat

- **Hipotezler doğrulanmamıştır.** Kök neden; üretim kayıtları, numune ölçümü ve yerinde incelemeyle doğrulanmadan
  düzeltici faaliyet başlatılmamalı ve müşteriye kesin kök neden bildirilmemelidir.
- **Kümeleme ürün ve parti alanlarına dayanır.** Parti bilgisi eksik ya da tutarsız yazılmışsa kümeler eksik çıkar.
- **Veri gönderimi.** Ürün, parti, hat, tarih, adet ve şikâyet açıklaması seçtiğiniz yapay zekâ sağlayıcısına gönderilir;
  müşteri adı gönderilmez. Telefon ve e-posta maskelenir. Gerekirse yerel model (Ollama) kullanın.
- **Örnek veri** kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk taslakta bilerek bir kümeyi hipotezsiz bırakır, başka kümeden bir şikâyeti
dayanak gösterir ve bir hipotezi "kesinlikle" diye yazar. Kodun bunları yakalayıp düzelttirdiği ve kümeleri doğru
oluşturduğu doğrulanır.

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
