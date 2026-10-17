# Uygunsuzluk Kaydı Yazımı · AI Agent

> Üretim › Kalite Kontrol Elemanı · Gerisi Bende

Kalite kontrolcünün sahada aldığı kısa notları **standart uygunsuzluk (NCR) kaydına** çevirir. Notlar üretim hattından,
şantiyeden, laboratuvardan ya da final kontrolden gelebilir. Kayıt şunları içerir: nesnel hata tanımı, ürün ve parti,
miktarlar, hata kodu, sınıf (Kritik / Majör / Minör), ilk müdahale, karar önerisi ve eksik bilgiler. DÖF gerekip
gerekmediğine ve termine kod karar verir; DÖF listesi **DÖF Takibi** kod bloğuna doğrudan verilir.

## Nasıl çalışır?

1. **Hazırlık (kod):** Not dosyası kayıtlara ayrılır (`Kayıt: GG.AA.YYYY SS:DD | Yer`) ve satırlar numaralanır. Telefon,
   e-posta ve kimlik numaraları maskelenir. Hata kodu listesi verilmişse okunur.
2. **Kayıt (model):** Her not için başlık, hata tanımı, ürün, parti, kontrol / hatalı / etkilenen miktar, hata kodları,
   sınıf, yapılmış ilk müdahaleler, karar önerisi ve eksik bilgiler yazılır. Model hesap yapmaz, neden tahmin etmez.
3. **Denetim (kod):** Aşağıdakilerden biri varsa yalnız o kayıt modele geri gönderilir (varsayılan en fazla 2 tur):
   - Bir sayı, ürün ya da parti notta geçmiyor (ör. 3 + 1 hatayı toplayıp "4" yazmak da kabul edilmez)
   - Hata kodu listede yok; hatalı miktar kontrol miktarından büyük; dayanak satır geçersiz
   - Notta güvenlik işareti var ama sınıf Kritik değil. İşaretler: metal parça / metal dedektör, cam kırığı, yabancı
     madde, alerjen, kontaminasyon, yaralanma, iş kazası, taşıyıcı / yapısal, çökme, gıda güvenliği…
4. **Hesap (kod):**

| Alan | Kural |
|---|---|
| Hata oranı | hatalı ÷ kontrol × 100 (ikisi de notta yazıyorsa) |
| DÖF gerekli | Sınıf Kritik veya Majör, ya da aynı hata kodu girdide birden fazla kayıtta |
| DÖF termini | Kayıt tarihi + Kritik 7, Majör 15, Minör 30 gün (`--termin` ile değişir) |
| Eksik bilgi | Parti / lot / irsaliye no yoksa (izlenebilirlik), miktar yoksa eklenir |

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                     # örnek 4 kayıt (üretim, gıda, şantiye, tekstil)
python agent.py --girdi notlar.txt --kodlar hata_kodlari.csv --sorumlu "Kalite Mühendisi"
python agent.py --girdi notlar.txt --termin "kritik=3,majör=10,minör=30" --onek NCR
```

Not dosyası örneği:

```
Kayıt: 14.10.2026 09:40 | Hat 2 - Kapak montaj
Saatlik kontrolde 500 adetlik numuneden 12 kapakta sızdırmazlık testi geçmedi.
Ürün: 5 lt bidon kapağı KP-50, parti L2026-1012.
Hat durduruldu, son 2 saatin üretimi karantinaya alındı.
```

Ardından DÖF takibi için:

```bash
python main.py --doflar cikti/uygunsuzluklar_dof.xlsx      # DÖF Takibi kod bloğu klasöründe
```

## Çıktı

- `cikti/uygunsuzluklar.xlsx`:
  - `Uygunsuzluklar`: kayıt no (UGS-261014-01), tarih, yer, başlık, hata tanımı, ürün, parti, miktarlar, hata oranı, hata
    kodları, sınıf, ilk müdahale, karar önerisi, eksik bilgiler, DÖF ve termini, dayanak not satırları, son kontrol.
  - `DÖF Listesi`
  - `Kontrol`
- `cikti/uygunsuzluklar_dof.xlsx`: DÖF listesi tek başına; **DÖF Takibi** kod bloğunun girdi sütunlarıyla (DÖF No, Tür,
  Kaynak, Konu, Sorumlu, Açılış Tarihi, Termin …).

## Dikkat

- **Sınıf ve karar öneridir.** Ürünün kullanımı, ıskartası veya şartlı kabulü kaliteden sorumlu yetkilinin kararıdır.
  Gıda güvenliği ve iş güvenliği olaylarında işletmenizin acil durum / geri çağırma prosedürünü uygulayın; bu agent onun
  yerine geçmez.
- **Kritik işaret listesi sınırlıdır.** Kod yalnız listedeki ifadeleri tanır; listede olmayan bir güvenlik riskini
  model doğru sınıflandırmazsa kod fark etmez.
- **Kök neden yazılmaz.** Kayıt nesnel tespittir; kök neden analizi için 8D Rapor Taslağı veya Müşteri Şikâyeti Kök
  Neden Analizi agent'larını kullanın.
- **Kişisel veri.** Notlara personel adı yazmamanız önerilir (görev yeterlidir).
- **Örnek veri:** Ürünler, partiler ve olaylar kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk turda bilerek:
- metal dedektör olayını "Majör" sınıflar,
- notta olmayan bir ürün adı ("Hazır beton") yazar,
- 3 leke + 1 dikiş hatasını toplayıp hatalı miktarı "4" yazar.

Kodun üç sorunu yakalayıp yalnız o kayıtları yeniden yazdırdığı doğrulanır. Ayrıca şunlar doğrulanır:
- hata oranı, DÖF termini ve tekrarlayan hata kodu kuralı,
- DÖF dosyasının DÖF Takibi sütunlarıyla aynı olması.

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
