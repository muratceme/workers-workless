# Denetim Hazırlık Dosyası · Kod Bloğu

> Muhasebe › Muhasebe Müdürü · Gerisi Bende

Bağımsız denetçinin gönderdiği talep listesini izleme tablosuna çevirir: hangi belge gönderildi, hangisi hazır,
hangisi gecikti, kim neyi bekletiyor. İsterseniz belge klasörünüzü de tarar ve "hazır" denen dosyanın gerçekten
orada olup olmadığına bakar. İnternete bağlanmaz.

## Nasıl hesaplar?

| Kavram | Kural |
|---|---|
| Kapsam | Tüm talepler − "Uygulanmaz" olanlar |
| Hazırlık oranı | (Gönderildi + Hazır) ÷ kapsam |
| Gecikti | Gönderilmemiş ve istenen tarihi durum tarihinden önce (hazır olsa bile) |
| Yaklaşıyor | Henüz hazır değil ve istenen tarihe `--gun` (3) gün veya daha az kalmış |
| Gecikmeli teslim | Gönderim tarihi istenen tarihten sonra |
| Dosya kontrolü | `--klasor` verilirse "Dosya" sütunundaki adlar klasörde ve alt klasörlerinde aranır |

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                              # örnek veri, durum tarihi 12.10.2026
python main.py --talepler talepler.xlsx
python main.py --talepler talepler.xlsx --klasor denetim_belgeleri --tarih 12.10.2026 --gun 5
```

Zorunlu sütun: Talep. Diğerleri: No, Alan, Sorumlu, İstenen Tarih, Durum (Gönderildi / Hazır / Hazırlanıyor /
Başlamadı / Uygulanmaz), Gönderim Tarihi, Dosya, Not. Bir talebe birden çok dosya yazacaksanız adları virgülle
veya `|` ile ayırın.

## Çıktı

`cikti/denetim_hazirlik.xlsx`:

- `Talep Listesi`: her talebin izleme durumu, gecikme günü, yazılan dosyalar ve klasörde bulunamayanlar.
- `Alan Özeti`, `Sorumlu Özeti`: gönderilen, hazır, açık ve geciken talep sayıları, hazırlık yüzdesi.
- `Uyarılar`: geciken talepler, "hazır" yazıp klasörde olmayan dosyalar, sorumlusuz talepler, talebe bağlanmamış dosyalar.

## Dikkat

- **Kod belgelerin içeriğine bakmaz.** Dosyanın klasörde olması, doğru, güncel veya denetçi için yeterli olduğunu
  göstermez; mutabakatların tutup tutmadığı da kontrol edilmez.
- **Dosya eşleşmesi yalnız ada göredir.** Büyük-küçük harf ayrımı yapılmaz; aynı adlı iki dosya ayırt edilmez.
- **Talep listesini denetçiniz belirler.** Örnek liste yaygın başlıklardan oluşur; kapsam şirkete ve denetime
  göre değişir.
- **Gizlilik:** Kod dosya adlarını yalnız rapora yazar; hiçbir belgeyi açmaz, kopyalamaz veya dışarı göndermez.
- **Örnek veri:** Kişiler, tarihler ve dosya adları kurgusaldır.

## Testler

- İzleme durumları (gecikti, yaklaşıyor, hazır, gecikmeli teslim, terminsiz), hazırlık oranı, alan ve sorumlu özetleri.
- Klasör taraması: eksik dosya, adı yazılmamış dosya, alt klasör, talebe bağlanmamış dosya.
- Okuma hataları, tanınmayan durum, komut satırı.

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
