# Muallak Hasar Takibi · Kod Bloğu

> Hasar Yönetimi › Hasar Uzmanı · Gerisi Bende

Açık hasar dosyalarını **yaş, muallak tutar ve bekleyen aksiyona göre** izler; uzun süredir işlem görmeyen dosyaları
öncelik sırasıyla listeler. İnternete bağlanmaz.

## Nasıl hesaplar?

| Kavram | Kural |
|---|---|
| Dosya yaşı (gün) | Durum tarihi − ihbar tarihi; dilimler 0–30, 31–90, 91–180, 181–365, 365+ |
| Hareketsiz gün | Durum tarihi − son işlem tarihi (son işlem yazılmamışsa ihbar tarihi) |
| Hareketsiz dosya | Hareketsiz gün `--hareketsiz` (30) ve üstünde |
| Kritik | Hareketsiz ve muallak tutarı `--buyuk` (250.000) ve üstünde |
| Yüksek | Hareketsiz, muallak tutarı eşiğin altında |
| Orta | Hareketli ama yaşı `--eski` (180) gün ve üstünde, ya da bekleyen aksiyon yazılmamış, ya da muallak sıfır |
| Olağan | Diğerleri |

Liste öncelik, sonra hareketsiz gün (uzundan kısaya), sonra tutar sırasıyla dizilir.

**Örnek:** İhbar 10.03.2026, son işlem 01.09.2026, durum tarihi 12.10.2026 → dosya **216 günlük**, **41 gündür**
işlem yok. Muallak 850.000 → **Kritik**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                              # örnek veri, durum tarihi 12.10.2026
python main.py --dosyalar acik_dosyalar.xlsx
python main.py --dosyalar acik_dosyalar.xlsx --hareketsiz 45 --eski 365 --buyuk 500000 --tarih 31.10.2026
```

Zorunlu sütunlar: Dosya No, İhbar Tarihi. Dosyaya yalnız açık (muallak) dosyaları koyun; kapanmış dosyaları kod
ayırt etmez.

## Çıktı

`cikti/muallak_hasar_takibi.xlsx`:

- `Takip Listesi`: öncelik sırasıyla; hareketsiz gün, yaş, muallak tutar, bekleyen aksiyon, sorumlu.
- `Özet`: dosya ve tutar toplamları, hareketsiz dosyalar, yaş dağılımı.
- `Bekleyen Aksiyonlar`, `Sorumlular`, `Branşlar`: dosya, tutar ve bekleme süresi kırılımlı.
- `Uyarılar`.

## Dikkat

- **"Hareketsiz", yazdığınız son işlem tarihine göredir.** Dosyada yapılan iş sisteme işlenmediyse dosya
  hareketsiz görünür; tersine, önemsiz bir kayıt güncellemesi dosyayı hareketli gösterir. Hangi işlemlerin "son
  işlem" sayılacağını siz belirlersiniz.
- **Kod rezervi değerlendirmez.** Muallak tutarın yeterli olup olmadığına, eksper raporuyla uyumuna ya da
  güncellenmesi gerekip gerekmediğine bakmaz; yalnız sıfır tutarı işaretler.
- **Bekleme nedeni dosyaya göre haklı olabilir.** Dava, tahkim ya da yurt dışı evrak bekleyen dosyada 30 gün
  işlem olmaması olağandır; eşikleri iş yükünüze göre ayarlayın ya da bu dosyaları ayrı çalıştırın.
- **Yasal süreler izlenmez.** Tazminat ödeme süreleri, zamanaşımı ve şikâyet cevap süreleri bu raporun konusu
  değildir.
- **Kapsam dışı:** hasar ödeme onayı, eksper ataması, rücu ve sovtaj takibi, rezerv gelişimi ve yeterlilik analizi.
- **Örnek veri:** Dosyalar, sigortalılar ve tutarlar kurgusaldır.

## Testler

- Hareketsiz eşiği sınırı (29 / 30 gün), büyük dosya eşiği, eski dosya eşiği, son işlem tarihi olmayan dosya.
- Öncelik sıralaması, yaş dilimleri, aksiyon ve sorumlu kırılımı, ortalamalar.
- Tutarsız tarih, sıfır muallak, yazılmamış aksiyon uyarıları, okuma ve komut satırı.

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
