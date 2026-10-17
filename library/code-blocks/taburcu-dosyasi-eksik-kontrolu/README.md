# Taburcu Dosyası Eksik Kontrolü · Kod Bloğu

> Tıbbi Sekreterlik › Yatan Hasta Tıbbi Sekreteri · Gerisi Bende

Taburcu olan hastaların dosyalarında **onam, epikriz, ameliyat notu ve konsültasyon belgelerinin eksiksizliğini**
kontrol eder. İnternete bağlanmaz. Hasta adı ve kimlik numarası gerekmez; protokol numarası yeterlidir.

## Nasıl hesaplar?

1. **Kural dosyası** hangi belgenin hangi dosyada aranacağını söyler:

   | Belge | Koşul | Diğer Adlar |
   |---|---|---|
   | Epikriz | Hepsi | Çıkış özeti |
   | Aydınlatılmış onam formu | Hepsi | Onam formu, Rıza formu |
   | Ameliyat notu | Ameliyat | Ameliyat raporu |
   | Konsültasyon formu | Konsültasyon | Konsültasyon notu |

   Koşul boş veya "Hepsi" ise belge her dosyada aranır. Bir sütun adıysa (ör. `Ameliyat`), taburcu listesinde o sütunu
   "Evet" olan dosyalarda aranır.
2. **Belge var** sayılması için belge listesinde o protokol numarasıyla, kuraldaki adla veya "Diğer Adlar"dan biriyle
   kayıt bulunmalıdır. Büyük-küçük harf ve Türkçe karakter ayrımı yapılmaz.
3. **Dosya durumu:** eksik yoksa Tam; eksik varsa Eksik; eksik var ve taburcudan beri geçen gün `--sure-gun`
   değerinden büyükse Süre aşıldı. Süre verilmezse bu durum kullanılmaz.
4. **Tamlık %** = tam dosya ÷ dosya sayısı × 100.

**Örnek:** ameliyat olmuş, konsültasyon istenmiş dosyada yalnız epikriz ve ameliyat notu var → eksik: aydınlatılmış
onam formu, konsültasyon formu. Taburcu 08.10.2026, rapor 12.10.2026, süre 3 gün → 4 gün geçmiş → Süre aşıldı.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri (rapor tarihi 12.10.2026, süre 3 gün)
python main.py --taburcu taburcu_listesi.xlsx --belgeler belge_listesi.xlsx --kurallar belge_kurallari.xlsx --sure-gun 3
```

| Dosya | Sütunlar |
|---|---|
| Taburcu listesi | Protokol No (zorunlu); Servis, Hekim, Taburcu Tarihi; kurallarda geçen koşul sütunları (Evet / Hayır) |
| Belge listesi | Protokol No, Belge. Dosyada bulunan her belge bir satır |
| Kurallar | Belge (zorunlu); Koşul, Diğer Adlar (virgülle ayrılır) |

## Çıktı

`cikti/taburcu_dosyasi_kontrolu.xlsx`:

- `Dosya Durumu`: süresi aşılanlardan başlayarak her dosya, geçen gün ve eksik belgeler; boş "Tamamlandı" sütunu.
- `Servis Özeti`: servis başına dosya, tam, eksik, tamlık %, en çok eksik kalan belge.
- `Belge Özeti`: belge başına gereken dosya sayısı ve eksik oranı.
- `Uyarılar`: belgesi olup taburcu listesinde olmayan protokoller, kuralda olmayan belge türleri, mükerrer protokoller.

## Dikkat

- **Zorunlu belgeleri kurumunuz belirler.** Örnek kural dosyası görev tanımındaki dört belgeyle bir başlangıçtır;
  mevzuatın veya kalite standartlarının tam listesi değildir. Kendi dosya içerik listenize göre düzenleyin.
- **Yalnız varlığa bakar.** Belgenin doldurulmuş, imzalanmış, tarihli ve doğru hastaya ait olduğunu kontrol etmez.
- **Belge adları tutmalı.** Bilgi sisteminizdeki belge adı kuraldaki adla veya "Diğer Adlar"la aynı olmalıdır.
  Eşleşmeyen belge türleri `Uyarılar` sayfasında listelenir; oradan bakıp kural dosyasını tamamlayın.
- **Süre kurum kuralıdır.** Dosya tamamlama süresi için kod bir değer varsaymaz; `--sure-gun` ile siz verirsiniz.
- **Kişisel sağlık verisi:** Dosyalara hasta adı, T.C. kimlik no, tanı gibi alanları koymayın; kontrol için gerekmez.
  Çıktıyı kurumunuzun veri güvenliği kurallarına göre saklayın.
- **Örnek veri:** Protokoller, servisler ve hekimler kurgusaldır.

## Testler

- Koşula göre gereken belgeler, eksik belgeler, diğer adla eşleşme, hiç belgesi olmayan dosya.
- Süre sınırı (tam 3. gün), süresiz çalışma, taburcu tarihi olmayan dosya.
- Servis ve belge özetleri, listede olmayan protokol, kuralda olmayan belge türü.
- Kural ve liste okuma hataları, eksik koşul sütunu ve örnek veri uçtan uca.

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
