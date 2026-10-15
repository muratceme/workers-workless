# Operasyonel Risk Kayıp Veri Analizi · Kod Bloğu

> İç Kontrol ve Risk Yönetimi › Risk Yönetimi Uzmanı · Workers / Workless

Operasyonel kayıp olaylarını **olay türü ve iş kolu bazında** toplar; kaç olay yaşandığını (sıklık) ve olay başına ne
kadar kaybedildiğini (şiddet) yan yana gösterir. İnternete bağlanmaz.

## Nasıl hesaplar?

| Kavram | Kural |
|---|---|
| Net kayıp | Brüt kayıp − geri kazanım (sigorta tazminatı dahil) |
| Sıklık | Olay adedi; sıklık payı % = grubun adedi ÷ toplam adet × 100 |
| Şiddet | Olay başına ortalama net kayıp; ortanca ve en büyük olay ayrıca verilir |
| Kayıp payı % | Grubun net kaybı ÷ toplam net kayıp × 100 |
| Geri kazanım % | Geri kazanım ÷ brüt kayıp × 100 |
| Tutar dilimi | Brüt kayba göre; `--dilimler` (10000;100000;1000000), alt sınır dahil, üst sınır hariç |
| Büyük olay | Net kaybı `--buyuk` (100.000) ve üstünde olan olay |
| Kapsam | Brüt kaybı `--esik` (0) ve üstünde olan olaylar; altındakiler yalnız sayılır |
| Dönem | Olay tarihinin yılı ve çeyreği |

**Örnek:** Brüt kayıp 95.000, sigortadan 80.000 tahsil edildi → net kayıp **15.000**, geri kazanım **%84,21**.
Olay, brüt tutarıyla 10.000 – 100.000 diliminde sayılır.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri
python main.py --kayiplar kayiplar.xlsx
python main.py --kayiplar kayiplar.xlsx --esik 5000 --buyuk 250000 --dilimler 25000;250000;2500000
```

Zorunlu sütunlar: Olay Tarihi, Brüt Kayıp. Olay türü ya da iş kolu boşsa olay "(belirtilmemiş)" grubuna düşer ve
uyarı üretilir.

## Çıktı

`cikti/operasyonel_kayip_analizi.xlsx`:

- `Özet`: adet, brüt, geri kazanım, net, ortalama, ortanca, en büyük olay ve payı.
- `Olay Türleri`, `İş Kolları`, `Dönemler`: sıklık ve şiddet göstergeleri kırılımlı.
- `Tür × İş Kolu`: net kayıp matrisi.
- `Tutar Dilimleri`, `Büyük Olaylar`, `Tüm Olaylar`, `Uyarılar`.

## Dikkat

- **Kod olayı sınıflandırmaz.** Olay türünü ve iş kolunu dosyada siz yazarsınız; kod yalnız aynı yazılanları
  toplar (büyük-küçük harf ve Türkçe karakter farkını yok sayar). Bankacılıkta yaygın olan Basel sınıflandırması
  yedi olay türü ve sekiz iş kolu kullanır; kurumunuzun kendi sınıflandırması farklıysa onu kullanın.
- **Sermaye hesabı değildir.** Kod kayıp dağılımı modeli kurmaz, operasyonel risk sermaye yükümlülüğü ya da
  yasal raporlama tutarı üretmez; yalnız geçmiş olayları özetler.
- **Ortalama, büyük olaylardan çok etkilenir.** Kayıp verisi genellikle az sayıda çok büyük olay içerir; bu yüzden
  ortalamayı ortanca ve en büyük olayla birlikte okuyun.
- **Tarih olarak hangisini yazdığınız sonucu değiştirir.** Olayın gerçekleştiği, tespit edildiği ve muhasebeleştiği
  tarihler farklı olabilir; dönem tablosu yazdığınız tarihe göredir. Bütün dosyada aynı tarihi kullanın.
- **Kapsam dışı:** ramak kala olaylar ve dolaylı kayıplar (yazmadıysanız), aynı kök nedene bağlı olayların
  birleştirilmesi, kredi riskiyle sınır olayların ayrıştırılması, dış veri ve senaryo analizi.
- **Örnek veri:** Olaylar, tutarlar ve birimler kurgusaldır.

## Testler

- Net kayıp, ortalama, ortanca, en büyük olay, paylar ve geri kazanım oranı.
- Dilim sınırları (sınır değerin kendisi üst dilime girer), raporlama eşiği, büyük olay eşiği.
- Tür × iş kolu matrisi, dönem (çeyrek) ataması, yazım farkı olan grup adlarının birleşmesi, okuma ve komut satırı.

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
