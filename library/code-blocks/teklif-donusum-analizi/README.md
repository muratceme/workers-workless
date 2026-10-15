# Teklif Dönüşüm Analizi · Kod Bloğu

> Satış ve Fiyatlandırma › Lojistik Satış Uzmanı · Workers / Workless

Verilen navlun tekliflerinin **kazanılma oranını** hat, müşteri, ay, taşıma modu ve temsilci bazında hesaplar;
kaybedilen teklifleri **rakip fiyatıyla farkına** ve kayıp nedenine göre inceler, sonucu girilmemiş eski teklifleri
listeler. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Kazanma oranı % | Kazanılan ÷ (kazanılan + kaybedilen) × 100; beklemede ve iptal teklifler oranın dışında |
| Fiyat farkı % | (Teklif fiyatımız − rakip fiyatı) ÷ rakip fiyatı × 100; yalnız rakip fiyatı yazılmış sonuçlanan teklifler |
| Fiyat farkı bandı | ≤ 0 · 0–5 · 5–10 · 10 üstü; her bantta kazanma oranı |
| Kayıp nedeni payı | Nedene göre kaybedilen teklif ÷ toplam kaybedilen × 100 |

Hat, Çıkış ve Varış sütunlarından ("Mersin – Hamburg") ya da Hat sütunundan alınır. Durum sütunu boşsa teklif
beklemede sayılır; Kazanıldı / Kaybedildi (Reddedildi) / İptal / Beklemede (Açık, Değerlendirmede) yazımları tanınır.

**Uyarılar:** en az `--min` (varsayılan 5) sonuçlanmış teklifi olan hat ya da müşteride kazanma oranı `--esik`
(varsayılan %20) altında · `--bekleme` (varsayılan 14) günden uzun süredir sonuç bekleyen teklifler.

**Örnek:** Gebze – Paris hattında 8 teklif sonuçlanmış, 1'i kazanılmış → kazanma oranı **%12,50** (eşik %20 → uyarı).
Rakip fiyatı bilinen tekliflerde rakipten ucuz ya da eşit olduğumuz 11 tekliften 7'si kazanılmış (**%63,64**); %0'dan
pahalı olduğumuz tekliflerin hiçbiri kazanılmamış.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                              # örnek veri (30.09.2026)
python main.py --teklifler navlun_teklifleri.xlsx --tarih 30.09.2026
python main.py --teklifler navlun_teklifleri.xlsx --min 3 --esik 25 --bekleme 7
```

Zorunlu sütunlar: Tarih ve Teklif Fiyatı (ayrıca Teklif No ya da Müşteri). Revize teklifleri yeni numarayla yazın; aynı
numara ikinci kez alınmaz.

## Çıktı

`cikti/teklif_donusum_analizi.xlsx`:

- `Özet`: teklif sayıları, kazanma oranı, para birimine göre kazanılan tutar, kaybedilenlerde ortalama fiyat farkı.
- `Kırılımlar`: hat, müşteri, ay, taşıma modu ve temsilci bazında kazanma oranı.
- `Fiyat Farkı`: bantlara göre kazanma oranı ve rakip fiyatlı teklif listesi.
- `Kayıp Nedenleri`, `Bekleyen Teklifler`, `Uyarılar`.

## Dikkat

- **Rakip fiyatı ve kayıp nedeni müşteri beyanıdır;** doğrulanamayabilir. Fiyat farkı yalnız aynı para biriminde,
  aynı kapsamda (ör. navlun + yerel masraflar) verilmiş fiyatlar için anlamlıdır; kod para birimi çevirisi yapmaz ve
  kapsam farkını bilmez.
- **Az teklifte oran yanıltıcıdır.** `--min` altında sonuçlanmış teklifi olan hat ve müşteriler için uyarı verilmez;
  tabloda oranı yorumlarken teklif sayısına bakın.
- **Rakiplerle fiyat bilgisi paylaşmayın.** Fiyat bilgisi yalnız müşteri geri bildiriminden ve kamuya açık
  kaynaklardan toplanmalıdır; rakiplerle doğrudan fiyat bilgisi paylaşmak rekabet hukuku açısından risklidir.
- **Eşikler politika varsayılanıdır;** hattın ve müşterinin yapısına göre değiştirin.
- **Örnek veri:** Müşteri adları, hatlar, fiyatlar ve sonuçlar kurgusaldır.

## Testler

- Kazanma oranı (beklemede ve iptal hariç), kırılımlar, para birimine göre kazanılan tutar.
- Fiyat farkı ve bantlar, kayıp nedenleri, bekleyen eski teklifler, durum tanıma, uyarı eşikleri, okuma ve komut satırı.

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
