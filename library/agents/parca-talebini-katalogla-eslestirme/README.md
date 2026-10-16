# Parça Talebini Katalogla Eşleştirme · AI Agent

> Otomotiv › Satış Sonrası › Yedek Parça Danışmanı · Workers / Workless

Atölyeden ve müşterilerden gelen serbest metin parça taleplerini ("2020 Model A ön balata 1 takım, araç lifte bekliyor")
parça kataloğuyla eşleştirir:
- parça no, adet, model, yıl, aciliyet,
- bilgi eksikse **tahmin etmez**, talep sahibine sorulacak soruyu yazar.

**Model / yıl uyumunu, ön-arka / sol-sağ tutarlılığını ve stok durumunu kod denetler.** Talepler **Yedek Parça Sipariş
Önerisi** kod bloğunun parça dosyasına bekleyen iş emri olarak eklenir.

## Nasıl çalışır?

1. **Hazırlık (kod):** Katalog ve talepler okunur. Her talep için ad benzerliğine göre en fazla 8 aday parça seçilir.
   Stok ve fiyat modele gönderilmez; telefon, e-posta, TCKN ve `--gizle` adları maskelenir.
2. **Eşleştirme (model):** Talepteki her parça adaylardan biriyle eşleştirilir ya da soru yazılır.
3. **Denetim (kod):** Aşağıdakilerden biri varsa yalnız o talep modele geri gönderilir (varsayılan en fazla 2 tur):
   - Parça no adaylarda yok; istek talepte aynen geçmiyor
   - Adet, model ya da yıl talepte yazmıyor
   - **Uyumsuz parça:** katalogdaki uyumlu modeller ya da yıl aralığı dışında
   - **Taraf çelişkisi:** talep "sağ" derken sol parça. Talepte taraf yoksa ve katalogda aynı parçanın iki tarafı varsa
     soru gerekir.
   - Model / yıl yazmayan talepte tahminle eşleştirme
   - Parça no boşken soru yok; dayanaksız "Acil"
4. **Stok (kod):**
   - Talepler sırayla stoktan düşülür, acil talepler önce. Sonuç: karşılanır / kısmi / sipariş gerekir.
   - Adedi ya da parçası belirsiz satır "sorulacak" kalır ve stoktan düşülmez.
   - Denetimden geçmeyen talep de stoktan düşülmez.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                              # örnek 5 talep
python agent.py --talepler talepler.txt --katalog parcalar.xlsx --gizle "Ad Soyad"
```

Talep örneği:

```
[16.10.2026 09:10] Atölye (İE-2026-1040): 2020 Model A ön balata 1 takım, araç lifte bekliyor.
[16.10.2026 12:00] Atölye: fren balatası lazım, model yazmadı.
```

## Çıktı

- `cikti/parca_talepleri.xlsx`:
  - `Talepler`: eşleşen parça, adet, stok durumu, eksik adet, sorulacak soru (mor: model).
  - `Parça Özeti`: parça bazında talep, stok sonrası, eksik.
  - `Kontrol`
- `cikti/parca_talepleri_parcalar.xlsx`: talepler "Bekleyen İş Emri" sütununa eklenmiş katalog.

```bash
python ../../code-blocks/yedek-parca-siparis-onerisi/main.py --parcalar cikti/parca_talepleri_parcalar.xlsx
```

## Dikkat

- **Parçayı vermeden önce** şasi numarasıyla üreticinin elektronik parça kataloğunda (EPC) doğrulayın. Kod yalnız
  kataloğunuzdaki "Uyumlu Modeller" ve yıl aralığına bakar; motor tipi, donanım paketi ve ara model değişikliklerini
  bilemez.
- **Katalog kalitesi sonucu belirler.** Uyumlu model / yıl sütunları boşsa uyumluluk denetlenemez.
- **Örnek veri:** Parçalar, modeller ve talepler kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk turda "sağ taraf" far talebini sol fara eşleştirir. Kodun bunu yakalayıp
düzelttirdiği doğrulanır. Ayrıca şunlar doğrulanır:
- yıl / model uyumsuzluğu, model yokken tahmin, taraf sorusu,
- istek / adet / aday / aciliyet denetimleri,
- stok dağıtımı (kısmi, sorulacak, denetimden geçmeyen),
- kod bloğu dosyası,
- ad ve fiyatların gönderilmemesi.

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
