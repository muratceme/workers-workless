# Fire ve Verim Analizi · Kod Bloğu

> Üretim › Üretim Mühendisi · Workers / Workless

Hammadde girişi ile mamul çıkışından **ürün ve hat bazında fire oranını, verimi ve verim kaybının hammadde
maliyetini** hesaplar. İnternete bağlanmaz.

## Nasıl hesaplar?

```
Fire (kg)            = hammadde girişi − mamul çıkışı − geri kazanılan
Fire %               = fire ÷ hammadde girişi × 100
Verim %              = mamul çıkışı ÷ hammadde girişi × 100
Fire maliyeti        = fire × hammadde birim maliyeti
Standart üstü fire   = fire − hammadde girişi × standart fire %      (eksiyse 0)
Verim kaybı maliyeti = standart üstü fire × hammadde birim maliyeti
```

- **Örnek:** giriş 1.000 kg, çıkış 940 kg, geri kazanılan 20 kg → fire 40 kg (%4,00), verim %94,00. Birim maliyet
  50 TL/kg → fire maliyeti 2.000 TL. Standart fire %3 (30 kg) → standart üstü 10 kg → 500 TL.
- **Gruplarda oranlar toplamlardan** hesaplanır (kayıt oranlarının ortalaması alınmaz).
- **Standart üstü fire kayıt kayıt toplanır:** standardın altında kalan bir gün, standardı aşan başka bir günü silmez.
- Standart fire dosyada yoksa `--standart` ile tek bir oran verebilirsiniz; ikisi de yoksa yalnız fire ve verim hesaplanır.
- Çıkış + geri kazanılan girişten büyükse kayıt analize alınmaz ve uyarı verilir.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri
python main.py --kayitlar uretim_kayitlari.xlsx --standart 3
```

Sütunlar: Ürün, Hammadde Girişi (kg), Mamul Çıkışı (kg) (zorunlu); Tarih, Hat, Geri Kazanılan (kg), Birim Maliyet
(TL/kg), Standart Fire %.

## Çıktı

`cikti/fire_verim_analizi.xlsx`:

- `Özet`: toplam giriş, çıkış, fire, verim ve maliyetler.
- `Ürün Bazında`, `Hat Bazında`, `Ürün-Hat`: aynı göstergeler; verim kaybı maliyetine göre sıralı.
- `Kayıtlar`: kayıt bazında hesap; boş "Neden" sütunu.
- `Uyarılar`: okunamayan veya tutarsız kayıtlar.

## Dikkat

- **Yalnız hammadde maliyeti.** İşçilik, enerji, makine zamanı, bertaraf gideri ve hurda satış geliri hesaba girmez;
  gerçek kayıp genellikle daha yüksektir.
- **Birimler aynı olmalı.** Giriş ve çıkış aynı birimde (kg) girilmelidir; adet bazlı çıkışı ağırlığa çevirmeden yazmayın.
- **Dönem kayması:** Yarı mamul ve hat içi stok değişimi dikkate alınmaz. Vardiya sonunda hatta kalan malzeme fire
  gibi, önceki vardiyadan kalan malzeme verim fazlası gibi görünür.
- **Geri kazanılan malzeme** fire sayılmaz; kırma malzemenin değer kaybı varsa ayrıca hesaplayın.
- **Standart fire sizin hedefinizdir**; reçete veya proses standardından gelmelidir.
- **Örnek veri:** Hatlar, ürünler, miktarlar ve maliyetler kurgusaldır.

## Testler

- Kayıt bazında fire, verim, maliyet ve standart üstü fire (tam standart sınırı dahil).
- Ürün, hat ve ürün-hat grupları; ağırlıklı oranlar.
- Maliyeti veya standardı olmayan kayıtlar; okuma hataları; örnek veri uçtan uca.

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

Ayrıntılar: `LICENSE` (MIT) ve `SORUMLULUK_REDDI.md`.
