# Sefer Kârlılık Hesabı · Kod Bloğu

> Karayolu Operasyon › Karayolu Operasyon Uzmanı · Gerisi Bende

Navlun geliri, alt nakliyeci bedeli, otoyol, gümrük ve diğer masraflardan **sefer bazında kârlılığı** hesaplar.
İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Toplam masraf | Alt nakliyeci + yakıt + otoyol / geçiş + gümrük + harcırah + diğer |
| Kâr, marj % | Navlun geliri − toplam masraf; kâr ÷ navlun geliri × 100 |
| Km başına gelir, masraf, kâr | İlgili tutar ÷ km (km yazılmışsa) |
| İşaret | Zarar: kâr < 0. Düşük marj: 0 ≤ marj < `--marj-esik` (varsayılan %10) |

Müşteri ve güzergah özetlerinde oranlar toplamlardan hesaplanır. Seferler en düşük kârdan başlayarak sıralanır.

**Örnek:** gelir 3.000; yakıt 1.100, otoyol 300, gümrük 150, harcırah 250, diğer 100 → masraf 1.900, kâr 1.100, marj
%36,67. 2.000 km için km başına gelir 1,50, masraf 0,95.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri
python main.py --seferler seferler.xlsx --marj-esik 12
```

Sütunlar: Sefer No, Navlun Geliri (zorunlu); Tarih, Müşteri, Güzergah, Km, Alt Nakliyeci, Yakıt, Otoyol, Gümrük,
Harcırah, Diğer.

## Çıktı

`cikti/sefer_karlilik.xlsx`:

- `Sefer Bazında`: gelir, masraf kalemleri, kâr, marj, km başına değerler; zarar eden ve düşük marjlı seferler işaretli.
- `Müşteri Bazında`, `Güzergah Bazında`: toplamlar, zarar eden sefer sayısı, marj.
- `Uyarılar`: okunamayan değerler, geliri veya masrafı olmayan seferler.

## Dikkat

- **Yalnız sefere yazılan masraflar hesaplanır.** Araç amortismanı veya kirası, sigorta, bakım, lastik, sürücü maaşı
  ve genel giderler girilmediyse kâr olduğundan yüksek görünür. Bu kalemleri sefer başına paylaştırıp "Diğer"
  sütununa ekleyebilirsiniz.
- **Tek para birimi, tek KDV esası.** Kur çevirisi yapılmaz; gelir ve masrafları aynı para birimine çevirip, hepsini
  KDV hariç girin.
- **Boş dönüş ayrı sefer yazıldıysa** geliri 0 görünür ve zarar çıkar; gidiş-dönüşü tek sefer olarak değerlendirmek
  istiyorsanız satırları birleştirin.
- **Okunamayan masraf 0 sayılır** ve uyarılır; `Uyarılar` sayfasına bakmadan sonuca güvenmeyin.
- **Örnek veri:** Müşteriler, güzergahlar ve tutarlar kurgusaldır.

## Testler

- Kâr, marj ve km başına değerler; geliri olmayan sefer; işaret sınırları (tam %10, kâr 0).
- Müşteri ve güzergah toplamları, sıralama.
- Okuma hataları ve örnek veri uçtan uca.

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
