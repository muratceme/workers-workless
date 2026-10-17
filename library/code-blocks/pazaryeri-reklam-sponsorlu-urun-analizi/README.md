# Pazaryeri Reklam (Sponsorlu Ürün) Analizi · Kod Bloğu

> Performans Pazarlama › Performans Pazarlama Uzmanı · Gerisi Bende

Pazaryeri reklam raporlarından **ürün bazında reklam maliyeti/satış oranını (ACOS)** hesaplar ve kârsız reklamları
işaretler. İnternete bağlanmaz; reklam panelinden indirdiğiniz raporla çalışır.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Tıklama oranı (CTR) % | Tıklama ÷ gösterim × 100 |
| Tıklama başı maliyet (CPC) | Harcama ÷ tıklama |
| Dönüşüm % | Sipariş ÷ tıklama × 100 |
| ACOS % | Harcama ÷ reklam satışı × 100 |
| ROAS | Reklam satışı ÷ harcama |
| Başabaş ACOS % | Ürünün reklam öncesi kâr marjı (`Marj %` sütunu; yoksa `--hedef-acos`) |
| Reklam sonrası kâr | Reklam satışı × marj % − harcama |

| Sınıf | Koşul |
|---|---|
| Harcama yok | Harcama 0 |
| Veri az | Reklam satışı yok ve tıklama < `--min-tiklama` (varsayılan 20) |
| Satışsız harcama | Reklam satışı yok ve tıklama ≥ `--min-tiklama` |
| Kârsız | ACOS > başabaş ACOS |
| Kârlı | ACOS ≤ başabaş ACOS |
| Marj bilinmiyor | Ürünün marjı yazılmamış ve `--hedef-acos` verilmemiş |

Aynı ürün birden çok kampanyada veya satırda geçiyorsa toplanır; oranlar toplamlardan hesaplanır.

**Örnek:** harcama 1.800, reklam satışı 6.000, marj %20 → ACOS %30 > %20 → Kârsız; reklam sonrası kâr
6.000 × 0,20 − 1.800 = −600.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri
python main.py --reklam reklam_raporu.xlsx --hedef-acos 20 --min-tiklama 30
```

Sütunlar: Ürün, Harcama, Reklam Satışı (zorunlu); Kampanya, Gösterim, Tıklama, Sipariş, Marj %.

## Çıktı

`cikti/reklam_analizi.xlsx`:

- `Ürün Bazında`: sınıf, göstergeler, başabaş ACOS, reklam sonrası kâr; toplam satırı; boş "Aksiyon" sütunu.
- `Kampanya Bazında`: kampanya başına aynı göstergeler.
- `Uyarılar`: okunamayan değerler, marjı olmayan veya satırlarda farklı marjı olan ürünler.

## Dikkat

- **Marj sizin verinizdir ve sonucu belirler.** Buradaki marj; ürün maliyeti, pazaryeri komisyonu, kargo ve iade payı
  düşüldükten sonra, reklamdan önce kalan orandır. Brüt marjı yazarsanız kârsız reklamlar kârlı görünür.
- **KDV esası aynı olmalı.** Reklam satışı KDV dahil, harcama KDV hariç raporlanıyorsa ACOS olduğundan düşük çıkar;
  ikisini aynı esasa getirin.
- **Reklam satışı, reklamın getirdiği ek satış değildir.** Panel, reklama tıklayıp satın alanları sayar; bu
  müşterilerin bir kısmı reklam olmasa da alacaktı. İlişkilendirme (atıf) süresi ve kuralı pazaryerine göre değişir.
- **Organik etki yoktur.** Reklamın sıralamaya ve organik satışa katkısı ölçülmez; yeni ürün tanıtımı gibi bilinçli
  zarar edilen reklamlar da "Kârsız" çıkar.
- **Kısa dönem yanıltır.** Birkaç günlük veride sipariş sayısı az olduğu için ACOS oynaktır; en az birkaç haftalık
  raporla çalışın.
- **Örnek veri:** Ürünler, kampanyalar ve tutarlar kurgusaldır.

## Testler

- CTR, CPC, dönüşüm, ACOS, ROAS ve reklam sonrası kâr; paydası sıfır olan durumlar.
- Sınıf sınırları (ACOS tam marja eşit, tam 20 tıklama), hedef ACOS, marjı olmayan ürün.
- Aynı ürünün satırlarının toplanması, kampanya toplamları, kârsız harcama payı.
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
