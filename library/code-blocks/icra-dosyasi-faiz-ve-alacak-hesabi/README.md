# İcra Dosyası Faiz ve Alacak Hesabı · Kod Bloğu

> İcra Takip › İcra Takip Elemanı (Avukat Katibi) · Workers / Workless

Asıl alacak, faiz başlangıcı ve faiz türüne göre **dönemsel oranlarla işlemiş faizi**, kısmi ödemelerin mahsubunu ve
**güncel borcu** hesaplar. İnternete bağlanmaz.

## Nasıl hesaplar?

| Adım | Kural |
|---|---|
| Dönem faizi | Asıl alacak × yıllık oran ÷ 100 × gün ÷ `--gun-esasi` (365) |
| Gün sayısı | Dönem bitişi − dönem başlangıcı (başlangıç günü dahil, bitiş günü hariç) |
| Oran değişimi | Oran dosyasındaki her başlangıç tarihi yeni dönem açar |
| Kısmi ödeme | Önce masraf ve vekalet ücreti, sonra işlemiş faiz, kalan asıl alacaktan düşülür |
| Güncel borç | Kalan asıl + kalan faiz + kalan masraf |
| Günlük faiz | Kalan asıl × güncel oran ÷ 100 ÷ gün esası |

Basit faiz uygulanır; faize faiz yürütülmez. "Sabit Oran" dolu dosyalarda (akdi faiz) oran tablosuna bakılmaz.

**Örnek:** Asıl 100.000, faiz başlangıcı 01.01.2026, oran %24, 01.04.2026'dan itibaren %36; hesap tarihi 01.07.2026.
90 gün × %24 = 5.917,81; 91 gün × %36 = 8.975,34; işlemiş faiz 14.893,15; güncel borç 114.893,15.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek veri, hesap tarihi 13.10.2026
python main.py --dosyalar dosyalar.xlsx --oranlar faiz_oranlari.xlsx --tarih 13.10.2026
python main.py --dosyalar dosyalar.xlsx --oranlar faiz_oranlari.xlsx --odemeler odemeler.xlsx --gun-esasi 360
```

- Dosyalar: Dosya No, Asıl Alacak, Faiz Başlangıç (zorunlu); Borçlu, Faiz Türü veya Sabit Oran, Masraflar, Vekalet Ücreti.
- Faiz oranları: Faiz Türü, Başlangıç, Yıllık Oran. Tür adları dosyalar tablosundakiyle aynı yazılmalıdır.
- Ödemeler: Dosya No, Tarih, Tutar.

## Çıktı

`cikti/icra_faiz_hesabi.xlsx`:

- `Dosya Özeti`: işlemiş faiz, ödenen, kalan asıl, kalan faiz, güncel borç, günlük faiz.
- `Faiz Dönemleri`: her oran dönemi ve ödeme arası için gün, oran ve faiz.
- `Ödeme Mahsubu`: her ödemenin masrafa, faize ve asıl alacağa dağılımı.
- `Uyarılar`: oranı bulunamayan dosyalar, aralık dışı ve fazla ödemeler, okunamayan satırlar.

## Dikkat

- **Oranları siz verirsiniz.** Örnek dosyadaki oranlar örnektir. Yasal faiz, ticari avans faizi ve diğer oranların
  yürürlük tarihlerini Resmî Gazete ve TCMB duyurularından doğrulayıp oran dosyanızı güncel tutun.
- **Kapak hesabı değildir.** Tahsil harcı, cezaevi harcı, icra inkâr tazminatı, vekalet ücretine işleyen faiz ve
  icra dairesi yuvarlamaları dahil değildir. Bağlayıcı olan icra dairesinin hesabıdır.
- **Mahsup sırası varsayımdır.** Kod ödemeyi önce masrafa, sonra faize, sonra asıl alacağa mahsup eder. Sözleşme,
  ilam veya alacaklının beyanı farklı bir sıra gerektiriyorsa sonucu elle düzeltin.
- **Gün sayımı uygulamaya göre bir gün oynayabilir.** Kod başlangıç gününü sayar, bitiş gününü saymaz.
- **Takip öncesi işlemiş faiz, döviz alacağı ve değişken akdi faiz kapsam dışıdır.** Takip öncesi faizi "Masraflar"
  sütununa ekleyebilirsiniz; bu tutara faiz işletilmez.
- **Örnek veri:** Dosyalar, borçlular ve tutarlar kurgusaldır.

## Testler

- Oran değişimiyle dönemlere bölme, 360 ve 365 gün esası, sabit oran, hesap tarihi başlangıca eşitken sıfır faiz.
- Kısmi ödemenin masraf, faiz ve asıl alacağa mahsubu; aynı gün iki ödeme; fazla ödeme; aralık dışı ödeme.
- Oranı olmayan tür, başlangıçtan önceki hesap tarihi, okunamayan satırlar, komut satırı hataları; örnek veri uçtan uca.

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
