# Bölge Mağaza Karşılaştırma Raporu · Kod Bloğu

> Mağaza Operasyonları › Bölge Müdürü · Gerisi Bende

Bölgedeki mağazaları **metrekare verimliliği, LFL büyüme, dönüşüm ve stok devir hızına** göre karşılaştırır; her
mağazayı bölge ortalamasına göre endeksler ve sıralar. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| m² başına ciro | Dönem cirosu ÷ satış alanı (m²) |
| LFL büyüme % | (Ciro − geçen yıl ciro) ÷ geçen yıl ciro × 100 |
| Dönüşüm % | Fiş adedi ÷ ziyaretçi × 100 |
| Sepet ortalaması | Ciro ÷ fiş adedi |
| Stok devir hızı | Satılan malın maliyeti ÷ ortalama stok (dönem içinde) |
| Stok günü | Ortalama stok ÷ satılan malın maliyeti × dönem gün sayısı (`--gun`, varsayılan 30) |
| Bölgeye göre endeks | Mağazanın değeri ÷ bölgenin değeri × 100 |
| Ortalama sıra | m², LFL, dönüşüm ve stok devrindeki sıraların ortalaması (1 en iyi) |

Bölge değerleri mağaza oranlarının ortalaması değil, **toplamların oranıdır** (ör. bölge m² verimliliği = toplam ciro ÷
toplam alan). Bir gösterge yalnız gerekli iki sayıyı yazılmış mağazalardan hesaplanır.

**Uyarılar:** endeksi 100 − `--sapma` (varsayılan 20, yani 80) altında kalan m² verimliliği, dönüşüm ve stok devri;
LFL büyümesi `--lfl` (varsayılan −%5) altında kalan mağaza.

**Örnek:** Bornova 700 m²'de 1.050.000 ciro → m² başına 1.500; Ege bölgesi 1.700 m²'de 3.550.000 → 2.088,24. Endeks
1.500 ÷ 2.088,24 × 100 = **71,83** → 80'in altında, uyarı.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                              # örnek veri
python main.py --magazalar magaza_donem_verileri.xlsx --gun 31
python main.py --magazalar magaza_donem_verileri.xlsx --sapma 15 --lfl 0
```

Her satır bir mağazanın dönem toplamıdır. Zorunlu sütunlar: Mağaza ve Ciro. Bölge sütunu yoksa tüm mağazalar tek
bölge sayılır; varsa her mağaza kendi bölgesiyle karşılaştırılır.

## Çıktı

`cikti/bolge_magaza_karsilastirma.xlsx`:

- `Mağazalar`: bölge içinde ortalama sıraya göre; göstergeler, bölgeye göre endeksler (renkli) ve gösterge sıraları.
- `Bölgeler`: bölge toplamları; birden çok bölge varsa genel toplam.
- `Uyarılar`.

## Dikkat

- **Dönem aynı olmalı.** Tüm mağazaların ciro, ziyaretçi ve maliyet verileri aynı tarih aralığına ait olmalıdır; dönem
  içinde açılan, kapanan ya da tadilatta kalan mağazalar yanıltıcı sonuç verir.
- **LFL karşılaştırılabilirliği sizin kararınızdır.** Geçen yıl cirosunu yalnız karşılaştırılabilir (ör. en az 12 aydır
  açık, alanı değişmemiş) mağazalar için yazın; boş bırakılan mağaza LFL'ye girmez.
- **Stok maliyet bedeliyle olmalı.** Satılan malın maliyeti ve ortalama stok aynı değerleme esasında (maliyet) olmalıdır;
  satış fiyatlı stok devir hızını düşük gösterir. Ortalama stok için dönem başı ve sonu stoğun ortalaması
  kullanılabilir.
- **Eşikler politika varsayılanıdır.** %20 sapma ve −%5 LFL sınırı genel kabul görmüş oranlar değildir; konum, format
  ve mağaza büyüklüğü farkları gözetilmez. Uyarılar inceleme önerisidir.
- **Örnek veri:** Mağaza adları ve sayılar kurgusaldır.

## Testler

- m², LFL, dönüşüm, sepet, stok devri ve stok günü; bölge toplamları, endeksler, sıralar ve ciro payı.
- Eşikler, birden çok bölge ve genel toplam, aynı mağazanın iki kez yazılması, okuma ve komut satırı.

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
