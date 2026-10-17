# Oda Temizlik Atama Planı · Kod Bloğu

> Kat Hizmetleri › Kat Hizmetleri Müdürü (Housekeeper) · Gerisi Bende

Çıkış ve konaklama odalarını temizlik sürelerine göre kat görevlilerine **dengeli ve kat bütünlüğünü koruyarak**
atar; her görevli için temizlik sırasını çıkarır. İnternete bağlanmaz.

## Nasıl atar?

| Adım | Kural |
|---|---|
| Oda süresi | Çıkış odası `--cikis-dk` (30), konaklama odası `--konaklama-dk` (20), artı odaya özel ek süre |
| Atanmayan odalar | Durumu Boş, Temiz veya Arızalı olanlar |
| Görevli payı | Görevlinin kapasitesi ÷ toplam kapasite |
| Dağıtım | Odalar kat ve oda numarasına göre sıralanır; sıra bozulmadan paylarla orantılı dilimlere bölünür |
| Temizlik sırası | Önce aynı gün girişi olan çıkış odaları, sonra diğer çıkışlar, sonra konaklamalar |
| Yük % | Görevlinin toplam süresi ÷ kapasitesi × 100 |

Dağıtımda bir odanın orta noktası görevlinin kümülatif pay sınırını aşınca sonraki görevliye geçilir. Böylece her
görevli birbirine komşu odaları alır ve iş yükü kapasiteyle orantılı kalır.

**Örnek:** 14 oda, toplam 375 dk; kapasiteler 175, 175 ve 100 dk. Sınırlar 145,8 ve 291,7 dk. İlk görevli 1. kattaki
5 odayı (145 dk), ikinci görevli 2. kat ile 301'i (150 dk), üçüncü görevli 302–304'ü (80 dk) alır.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek veri
python main.py --odalar odalar.xlsx --gorevliler gorevliler.xlsx
python main.py --odalar odalar.xlsx --gorevliler gorevliler.xlsx --cikis-dk 35 --konaklama-dk 15 --kapasite 400
```

- Odalar: Oda No, Durum (zorunlu); Kat, Giriş Var (Evet), Ek Süre (dk), Oda Tipi.
- Görevliler: Görevli; Kapasite (dk) boşsa `--kapasite` (420) kullanılır.

## Çıktı

`cikti/temizlik_atama_plani.xlsx`:

- `Görevli Özeti`: oda sayısı, çıkış ve konaklama dağılımı, toplam süre, yük yüzdesi, katlar.
- `Atama Planı`: görevli bazında temizlik sırasıyla odalar; tamamlandı işareti için boş sütun.
- `Uyarılar`: kapasite aşımı, toplam iş yükünün ekibi aşması, okunamayan satırlar.

## Dikkat

- **Süreler varsayımdır.** 30 ve 20 dakika örnek değerlerdir; oda büyüklüğü, tesis standardı ve hizmet düzeyine göre
  değişir. Kendi ölçümlerinizi kullanın; büyük odalar için "Ek Süre" sütununu doldurun.
- **Kapasite net çalışma süresidir.** Mola, toplantı, malzeme hazırlığı ve kat arası yürüme süresini düşerek girin.
- **Yürüme yolu hesaplanmaz.** Kod kat ve oda sırasını korur; bloklar, kanatlar veya asansör konumu için oda
  numarası sıranız fiziksel sırayı yansıtmıyorsa atamayı elle düzeltin.
- **Plan gün içinde değişir.** Geç çıkış, erken giriş, uzatma ve "rahatsız etmeyin" odaları için listeyi güncelleyip
  yeniden çalıştırın.
- **Çalışma süreleri ve iş yükü sınırları** iş mevzuatına ve varsa toplu sözleşmenize göre belirlenir; kod bunu denetlemez.
- **Örnek veri:** Odalar ve görevliler kurgusaldır.

## Testler

- İki ve üç görevliyle dağıtım, kapasiteyle orantılı paylar, farklı süre ayarları, doğal oda sıralaması.
- Temizlik sırası (girişi olan çıkışlar önce), kümülatif süre, kapasite aşımı ve oda düşmeyen görevli uyarıları.
- Okunamayan oda ve görevli satırları, mükerrer oda, komut satırı hataları; örnek veri uçtan uca.

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
