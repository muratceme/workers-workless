# Parti Verim ve Fire Raporu · Kod Bloğu

> Üretim › Üretim Şefi · Workers / Workless

Parti bazında hammadde girişi, mamul çıkışı ve fireden **verimi** hesaplar; hedef dışı partileri, ürün ve hat özetini
çıkarır. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Verim % | Mamul çıkışı ÷ hammadde girişi × 100 |
| Fire % | Fire ÷ hammadde girişi × 100 |
| Kayıt dışı fark | Giriş − mamul − fire − yeniden işleme |
| Hedeften sapma (puan) | Verim − hedef verim |
| Kayıp mamul | (Hedef − verim) ÷ 100 × giriş (yalnız hedefin altındaki partilerde) |
| Ürün ve hat verimi | Σ mamul ÷ Σ giriş × 100 (ağırlıklı) |

Durum: sapma −tolerans puanından kötüyse **Hedefin altında**, +toleranstan iyiyse **Hedefin üzerinde**, arada ise
**Hedefte**. Verimi %100'ü aşan, çıkış toplamı girişi geçen ve kayıt dışı farkı `--fark` (%5) üzerinde olan partiler
ayrıca uyarılır.

**Örnek:** Giriş 1.000 kg, mamul 880 kg, fire 70 kg; hedef %92. Verim %88, fire %7, kayıt dışı fark 50 kg (%5).
Sapma −4 puan → "Hedefin altında"; kayıp mamul 4 ÷ 100 × 1.000 = 40 kg.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek veri
python main.py --partiler partiler.xlsx --hedefler hedefler.xlsx
python main.py --partiler partiler.xlsx --hedef 90 --tolerans 1,5 --fark 3
```

- Partiler: Parti No, Hammadde Girişi, Mamul Çıkışı (zorunlu); Ürün, Tarih, Hat, Fire, Yeniden İşleme.
- Hedefler: Ürün, Hedef Verim %. Dosyada olmayan ürünler için `--hedef` kullanılır; o da yoksa "Hedef yok".

Miktarlarda binlik ayırıcı kullanmayın (`1250,5`).

## Çıktı

`cikti/parti_verim_raporu.xlsx`:

- `Partiler`: verim, fire, kayıt dışı fark, hedeften sapma, durum, kayıp mamul; neden için boş sütun.
- `Ürün Özeti`, `Hat Özeti`: ağırlıklı verim ve fire, en düşük ve en yüksek parti verimi, hedef altı parti sayısı.
- `Uyarılar`: hedefin altındaki partiler, şüpheli kayıtlar, okunamayan satırlar.

## Dikkat

- **Giriş ve çıkış aynı birimde olmalı.** Kilogram girip adet çıkış yazarsanız verim anlamsızdır; adedi birim ağırlıkla
  kilograma çevirin.
- **Nem kaybı fire değildir ama verimi düşürür.** Pişirme, kurutma, kavurma gibi proseslerde verim doğal olarak
  %100'ün altındadır; su eklenen proseslerde üzerine çıkabilir. Hedefi prosesinize göre siz belirlersiniz.
- **"Kayıt dışı fark" bir sinyaldir.** Tartılmayan fire, numune, nem kaybı veya eksik kayıttan kaynaklanır; suistimal
  kanıtı değildir.
- **Yeniden işleme mamul sayılmaz.** Sonraki partiye girdi olarak yazılıyorsa o partinin girişinde görünür.
- **Hat sütununda vardiya da kullanılabilir.** Kod bu sütuna göre ikinci bir özet çıkarır.
- **Örnek veri:** Ürünler, partiler ve hedefler kurgusaldır.

## Testler

- Verim, fire, kayıt dışı fark, sapma, kayıp mamul; tolerans sınırları ve hedefsiz ürün.
- Verimi %100'ü aşan, çıkışı girişten fazla, kayıt dışı farkı yüksek partiler.
- Ürün, hat ve genel özetlerin ağırlıklı hesabı; okunamayan satırlar, komut satırı hataları; örnek veri uçtan uca.

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
