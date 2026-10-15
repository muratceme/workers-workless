# OEE Hesaplama · Kod Bloğu

> Üretim › Üretim Mühendisi · Workers / Workless

Vardiya ve makine bazında **Kullanılabilirlik × Performans × Kalite = OEE** hesaplar; kayıpları dakika
olarak ayrıştırır ve makine bazında özetler. İnternete bağlanmaz.

## Hesap

| | |
|---|---|
| Planlı üretim süresi | Vardiya süresi − planlı duruş (mola, planlı bakım) |
| Çalışma süresi | Planlı üretim süresi − plansız duruş (arıza, ayar, malzeme bekleme) |
| Kullanılabilirlik | Çalışma süresi ÷ planlı üretim süresi |
| Performans | İdeal çevrim süresi × toplam adet ÷ çalışma süresi |
| Kalite | Sağlam adet ÷ toplam adet |

Makine ve genel değerler **süre ve adet toplamlarından** hesaplanır; vardiya oranlarının basit ortalaması
alınmaz (uzun ve kısa vardiyaları yanlış ağırlıklandırır). Test, OEE literatüründeki klasik örnekle
(K %88,81 · P %86,11 · Q %97,80 · OEE %74,79) birebir doğrulanır.

Kayıplar: kullanılabilirlik kaybı (duruş dakikası), performans kaybı (yavaş çalışma ve küçük duruşlar),
kalite kaybı (hatalı ürünlere harcanan ideal süre), değer katan süre.

**Kontroller:** performans %100'ü aşarsa (ideal çevrim süresi gerçekçi değil veya veri hatalı), duruşlar
vardiya süresini aşarsa, hatalı adet toplamdan büyükse uyarı verilir.

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                   # örnek vardiya kayıtlarıyla dener
python main.py --girdi vardiya_kayitlari.xlsx
```

Sütunlar: **Makine, Vardiya Süresi (dk), Plansız Duruş (dk), İdeal Çevrim Süresi (sn), Toplam Üretim, Hatalı**
(zorunlu); Tarih, Vardiya, Ürün, Planlı Duruş (dk) (isteğe bağlı).

## Testler

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
