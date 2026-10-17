# Üretim Hedef Takibi · Kod Bloğu

> Satış ve Acente Yönetimi › Bölge Satış Yöneticisi · Gerisi Bende

Branş ve acente bazında **prim üretimini aylık ve yıl başından bugüne hedefle** karşılaştırır; yıllık hedefe ulaşmak
için kalan aylarda gereken aylık üretimi hesaplar. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Ay gerçekleşme % | Rapor ayındaki üretim ÷ rapor ayının hedefi × 100 |
| Kümülatif gerçekleşme % | Ocak'tan rapor ayına kadar üretim ÷ aynı dönemin hedefi × 100 |
| Fark | Kümülatif üretim − kümülatif hedef |
| Yıllık hedefe göre % | Kümülatif üretim ÷ yıllık hedef × 100 |
| Kalan hedef | Yıllık hedef − kümülatif üretim (eksiye düşmez) |
| Gereken aylık üretim | Kalan hedef ÷ kalan ay sayısı |

Durum kümülatif gerçekleşmeye göre verilir: %100 ve üstü "Hedefin üstünde", `--esik` (80) ve üstü "Hedefe yakın",
altı "Hedefin gerisinde"; hedefi olmayan satır "Hedefsiz".

Hedef dosyasında `Ay` sütununa ay (09.2026, 2026-09, Eylül 2026 ya da bir tarih) ya da yalnız yıl (2026)
yazabilirsiniz. **Yalnız yıl yazılan hedef on iki aya eşit bölünür.** Üretim dosyasında her satır bir poliçe ya da
aylık toplam olabilir; tarih hangi aya düşüyorsa o aya eklenir.

**Örnek:** Yıllık hedef 1.200.000 → aylık 100.000. Eylül sonu kümülatif hedef 900.000, üretim 880.000 →
**%97,78**. Kalan hedef 320.000, kalan 3 ay → ayda **106.666,67** gerekiyor.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                              # örnek veri, rapor ayı Eylül 2026
python main.py --hedefler hedefler.xlsx --uretim uretim.xlsx
python main.py --hedefler hedefler.xlsx --uretim uretim.xlsx --ay 09.2026 --esik 90
```

`--ay` verilmezse üretim dosyasındaki son ay rapor ayı olur. Rapor yılı dışındaki satırlar hesaba girmez.

## Çıktı

`cikti/uretim_hedef_takibi.xlsx`:

- `Acenteler`, `Branşlar`: kümülatif gerçekleşmeye göre en geriden başlayarak sıralı; son satır toplam.
- `Acente × Branş`: aynı göstergeler ayrıntılı.
- `Aylar`: yılın her ayı için toplam hedef, üretim ve gerçekleşme.
- `Uyarılar`.

## Dikkat

- **Yıllık hedefin eşit bölünmesi bir varsayımdır.** Üretiminiz mevsimselse (yenileme dönemleri, yıl sonu
  kampanyaları) kümülatif gerçekleşme yanıltır; hedefi aylık yazın.
- **"Üretim" tanımı verinize göredir.** Brüt prim mi net prim mi, iptal ve iade zeyilleri düşülmüş mü, tanzim
  tarihi mi başlangıç tarihi mi; kod ayırmaz. Hedefle aynı tanımı kullanın. Eksi tutarlı satırlar (iptal, iade)
  üretimden düşer.
- **Rapor ayı tamamlanmamışsa oran düşük çıkar.** Ayın ortasında çalıştırırsanız ay hedefinin tamamıyla yarım
  ayın üretimi karşılaştırılır.
- **Acente ve branş adları iki dosyada aynı yazılmalı.** Büyük-küçük harf ve Türkçe karakter farkı yok sayılır;
  başka yazım farkı ayrı satır üretir ve bir yanda "hedefsiz üretim", öbür yanda "üretim yok" uyarısı çıkar.
- **Kapsam dışı:** poliçe adedi ve yenileme oranı hedefleri, komisyon ve prim teşvik hesabı, hasar prim oranına
  göre kârlı üretim ayrımı.
- **Örnek veri:** Acenteler, hedefler ve üretim tutarları kurgusaldır.

## Testler

- Ay ve kümülatif gerçekleşme, yıllık hedefin aylara bölünmesi, kalan hedef ve gereken aylık üretim.
- Durum eşikleri, hedefsiz üretim, hedefi olup üretimi olmayan branş, rapor yılı dışındaki satırlar.
- Ay yazım biçimleri, rapor ayının üretimden bulunması, okuma ve komut satırı.

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
