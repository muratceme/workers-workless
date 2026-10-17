# Medula Ret ve Kesinti Analizi · Kod Bloğu

> Anlaşmalı Kurumlar ve Faturalama › SGK Faturalama Uzmanı · Gerisi Bende

Medula ret/kesinti kayıtlarını **hata kodu, branş ve işlem bazında** analiz eder ve **tekrarlayan hataları** raporlar.
Medula'ya bağlanmaz; sizin dışa aktardığınız listeyi sayar. Hasta bilgisi gerekmez.

## Nasıl hesaplar?

| Çıktı | Kural |
|---|---|
| Hata kodu bazında | Adet, tutar, tutar payı %, birikimli pay % (tutara göre büyükten küçüğe), görüldüğü dönem sayısı |
| Tekrarlayan hata | Hata kodu en az `--min-donem` (2) farklı dönemde ve toplam en az `--min-adet` (3) kez görülmüş |
| Tekrarlayan hata-işlem | Aynı hata kodu + aynı işlem en az `--min-adet` kez görülmüş |
| Branş ve işlem bazında | Adet, tutar, pay, en sık hata kodu |
| Dönem eğilimi | Dönem bazında adet ve tutar; önceki döneme göre tutar değişimi % |

Dönem şu biçimlerde tanınır: `2026-08`, `08.2026`, `2026/8`, `Ağustos 2026` veya bir tarih. Tanınmayan dönem metni
olduğu gibi alınır, sıralamada sona konur ve uyarılır.

**Örnek:** bir hata kodu üç dönemde toplam 3 kayıtla 4.500 tutarında; toplam kesinti 6.800 → pay %66,18; üç dönemde
görüldüğü için "Tekrarlayan".

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri (3 dönem)
python main.py --kayitlar ret_kesinti_kayitlari.xlsx --min-donem 3 --min-adet 5
```

Sütunlar: Hata Kodu, Tutar (zorunlu); Dönem, Branş, İşlem Kodu, İşlem Adı, Hata Açıklaması. Her satır bir ret veya
kesinti kaydıdır.

## Çıktı

`cikti/ret_kesinti_analizi.xlsx`:

- `Hata Kodu Bazında`: Pareto sırasıyla hata kodları, tekrarlayanlar işaretli; boş "Düzeltici Aksiyon" sütunu.
- `Tekrarlayan Hata-İşlem`: aynı hatanın aynı işlemde tekrarlandığı çiftler.
- `Branş Bazında`, `İşlem Bazında`: adet, tutar, pay, en sık hata kodu.
- `Dönem Eğilimi`: dönemler arası değişim.
- `Uyarılar`: okunamayan tutarlar, tanınmayan dönem biçimleri.

## Dikkat

- **Kod, kuralları değerlendirmez.** Hata kodlarının anlamını, SUT hükümlerini ve Medula iş kurallarını bilmez; yalnız
  sayar ve gruplar. Hata açıklamaları sizin dosyanızdan gelir. Bir kesintinin haklı olup olmadığına ve itiraz
  edilip edilmeyeceğine bu rapor karar vermez.
- **Kesinti oranı yoktur.** Rapor yalnız ret/kesinti kayıtlarını içerir; toplam fatura tutarına oranı hesaplanmaz.
  Dönemler arası karşılaştırmada fatura hacminin değiştiğini göz önünde bulundurun.
- **Hata kodu yazımı tutarlı olmalı.** Aynı kodun farklı yazımları (boşluk, tire) ayrı kod sayılır; yalnız
  büyük-küçük harf farkı birleştirilir.
- **Tekrarlayan eşikleri tercihtir.** 2 dönem ve 3 kayıt başlangıç değeridir; kayıt hacminize göre ayarlayın.
- **Kişisel sağlık verisi:** T.C. kimlik no, hasta adı, takip numarası ve tanı bilgisi analiz için gerekmez; dışa
  aktarımdan çıkarın. Çıktıyı kurumunuzun veri güvenliği kurallarına göre saklayın.
- **Örnek veri:** Hata kodları (H-01 …), işlem kodları (I-100 …), açıklamalar ve tutarlar kurgusaldır; gerçek Medula
  veya SUT kodları değildir.

## Testler

- Dönem biçimlerinin tanınması.
- Hata kodu bazında adet, tutar, pay, birikimli pay; tekrarlayan hata sınırları.
- Tekrarlayan hata-işlem çiftleri, branş ve işlem toplamları, dönem değişimi.
- Okuma hataları ve örnek veri uçtan uca.

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
