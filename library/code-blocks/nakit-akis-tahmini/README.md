# Nakit Akış Tahmini · Kod Bloğu

> Finans › Finans Uzmanı · Workers / Workless

Vadeli alacak ve borçlar, alınan/verilen çekler, kredi taksitleri ve tekrarlayan ödemelerden (maaş, kira,
vergi, SGK) **haftalık nakit akış tahmini** (varsayılan 13 hafta) üretir. Minimum nakit eşiğinin altına düşen
haftaları kırmızıyla gösterir. İnternete bağlanmaz.

## Varsayımlar (değiştirilebilir)

- `--tahsilat-gecikmesi`: "Müşteri" kategorisindeki tahsilatlar vadeden bu kadar gün sonra gerçekleşir.
- `--gecikmis-tahsil-orani`: başlangıçtan önce vadesi geçmiş alacakların bu yüzdesi 1. haftada tahsil edilir (varsayılan %50).
- Vadesi geçmiş borçlar 1. haftada ödenir.
- Tekrarlayan kalemlerin (Haftalık/Aylık) başlangıçtan önceki dönemleri tahmine alınmaz.

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                             # örnek kalemlerle dener
python main.py --girdi nakit_kalemleri.xlsx --acilis 1250000 --baslangic 05.10.2026 --minimum 250000 --tahsilat-gecikmesi 7
```

Sütunlar: **Tür** (Giriş/Çıkış), **Tarih**, **Tutar** zorunlu; Kategori (ör. Müşteri, Tedarikçi, Maaş, Kira,
Kredi, Vergi, SGK, Çek), Açıklama, Tekrar (Haftalık/Aylık), Bitiş isteğe bağlı.

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
