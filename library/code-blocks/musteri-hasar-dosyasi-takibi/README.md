# Müşteri Hasar Dosyası Takibi · Kod Bloğu

> Teknik ve Operasyon › Hasar Sorumlusu · Gerisi Bende

Acente veya broker tarafında izlediğiniz **açık hasar dosyalarını** öncelik sırasına koyar; sigorta şirketi ve eksper
bazında özetler, bekleyen evrakları listeler ve müşteriye gönderilecek evrak hatırlatma metnini hazırlar. İnternete
bağlanmaz.

## Nasıl hesaplar?

```
Dosya yaşı     = rapor tarihi − ihbar tarihi          (ihbar tarihi yoksa hasar tarihi)
Hareketsiz gün = rapor tarihi − son işlem tarihi      (son işlem tarihi yoksa dosya yaşı)
```

| Öncelik (sırayla denenir) | Koşul |
|---|---|
| Süre aşımı | Dosya yaşı > `--azami-gun` (varsayılan 45) |
| Acil | Hareketsiz gün > 2 × `--hareketsiz` (varsayılan 14) |
| Takip et | Hareketsiz gün > `--hareketsiz` (varsayılan 7) |
| Evrak bekleniyor | Bekleyen evrak var |
| Normal | Diğerleri |

- **Açık dosya:** aşaması `kapandı`, `ödendi`, `ret / reddedildi` veya `iptal` ile başlamayan dosya.
- **Önerilen aksiyon:** bekleyen evrak varsa "müşteriden evrak isteyin"; dosya hareketsizse aşamaya göre "eksperi
  arayın", "ödeme tarihini sorun" veya "sigorta şirketini arayın".
- Bekleyen evraklar `|` veya `;` ile ayrılır.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri (rapor tarihi 01.10.2026)
python main.py --dosyalar hasar_dosyalari.xlsx --hareketsiz 5 --azami-gun 30
```

Sütunlar: Dosya No, Aşama (zorunlu); Müşteri, Telefon, Branş, Sigorta Şirketi, Şirket Dosya No, Eksper, Hasar Tarihi,
İhbar Tarihi, Son İşlem Tarihi, Tahmini Tutar, Bekleyen Evraklar.

## Çıktı

`cikti/hasar_dosyasi_takibi.xlsx`:

- `Açık Dosyalar`: öncelik, dosya yaşı, hareketsiz gün, bekleyen evrak ve önerilen aksiyon; boş "Not" sütunu.
- `Şirket Bazında`, `Eksper Bazında`: açık dosya sayısı, ortalama ve en eski dosya yaşı, öncelik dağılımı.
- `Bekleyen Evraklar`: dosya × evrak; boş "İstendi" ve "Geldi" sütunları.
- `Uyarılar`: okunamayan tarihler, ihbar tarihi olmayan dosyalar, mükerrer kayıtlar.

`cikti/evrak_hatirlatma.txt`: müşteri bazında bekleyen evrak listesi (mesaja yapıştırılabilir).

## Dikkat

- **45 gün bir iç takip eşiğidir.** Tazminatın ödenme süresi; branşa, belgelerin eksiksiz teslim edildiği tarihe ve
  ilgili mevzuata (ör. Türk Ticaret Kanunu, zorunlu sigortalarda özel kanunlar) göre değişir. Kod sürenin hangi
  tarihten başladığını ve dolup dolmadığını hukuken değerlendirmez.
- **Liste sizin kaydınızdır.** Kod sigorta şirketinin veya eksperin sistemine bağlanmaz; "son işlem tarihi"
  güncellenmezse dosya hareketsiz görünür.
- **Mesaj göndermez.** Evrak hatırlatma metni taslaktır; gönderim sizin kanalınızdan yapılır.
- **Kişisel veri:** Dosyalar sağlık dahil hassas bilgi içerebilir; KVKK'ya uygun saklayın ve paylaşmayın.
- **Tahmini tutar** dosyaya girdiğiniz değerdir; ödenecek tazminatı göstermez.
- **Örnek veri:** Müşteriler, sigorta şirketleri, eksperler ve tutarlar kurgusaldır.

## Testler

- Öncelik sınırları (45 / 46 gün, 7 / 8 / 14 / 15 gün) ve öncelik sırası.
- Aşamaya göre aksiyon metinleri; ihbar veya son işlem tarihi olmayan dosyalar.
- Şirket ve eksper özetleri, bekleyen evrak listesi ve hatırlatma metni.
- Okuma hataları ve örnek veri uçtan uca.

```bash
python -m unittest discover -s tests
```

## Kaynaklar ve dayanaklar

Kontrol tarihi: 15.10.2026. Mevzuat sık değişir; paketteki süre, oran ve kuralları kullanmadan önce güncel resmî metinle
karşılaştırın. "Karşılaştırılan maddeler" dışındaki atıflar yalnız yol göstericidir.

| Kaynak | Bağlantı | Durum |
|---|---|---|
| 6102 sayılı Türk Ticaret Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.6102.pdf) | Resmî metinle karşılaştırılan maddeler: 5/A, 22, 796, 1434 |
| 6698 sayılı Kişisel Verilerin Korunması Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.6698.pdf) | Resmî metinle karşılaştırılan maddeler: 13 |

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
