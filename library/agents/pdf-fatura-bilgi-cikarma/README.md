# PDF Fatura Bilgi Çıkarma · AI Agent

> Muhasebe › Ön Muhasebe Elemanı · Workers / Workless

E-fatura XML'i olmayan faturaları kayıt tablosuna aktarır: PDF faturalar, e-postayla gelen fatura metinleri, Word
belgeleri. Satıcı, alıcı, VKN, tarih, vade, satırlar (miktar, birim fiyat, iskonto, tutar, KDV, tevkifat) ve toplamlar
çıkarılır; hesaplar kodla denetlenir. E-fatura XML'leriniz varsa yapay zekâya gerek yoktur: **E-Fatura Okuma ve
Listeleme** kod bloğunu kullanın.

## Nasıl çalışır?

1. **Hazırlık (kod):** Belge okunur.
   - IBAN'lar `[IBAN-1]`, 11 haneli kimlik numaraları (şahıs firmalarında TCKN) `[TCKN-1]` yer tutucusuna çevrilir. Bu
     numaralar sağlayıcıya **gönderilmez**.
   - Telefon, e-posta ve `--gizle` ile verilen adlar maskelenir.
   - 10 haneli VKN şirket bilgisi olduğu için maskelenmez.
2. **Çıkarım (model):** Fatura bilgileri belgede yazdığı gibi çıkarılır. Model hiçbir değeri hesaplamaz; belgede yazmayan
   alanı boş bırakır.
3. **Denetim (kod):** Aşağıdakilerden biri varsa çıkarım modele geri gönderilir (varsayılan en fazla 2 tur):
   - Fatura no, unvan, tarih veya mal/hizmet adı belgede aynen geçmiyor
   - Bir sayı belgede yazmıyor: model belgedeki yanlış tutarı "düzeltmiş" ya da vadeyi kendisi hesaplamış olabilir
   - VKN belgede yok; IBAN alanında yer tutucu yerine uydurma bir numara var
4. **Bulgular (kod):** Gerçek numaralar geri konur ve faturanın kendisi kontrol edilir:

| Önem | Kontrol |
|---|---|
| Hata | VKN (10 hane) ve TCKN (11 hane) kontrol haneleri |
| Hata | Satır tutarı = miktar × birim fiyat − iskonto |
| Hata | KDV = tutar × oran; oran fatura tarihinde geçerli mi (10.07.2023 ve sonrası %0, 1, 10, 20; öncesi %0, 1, 8, 18) |
| Hata | Tevkifat oranı tanımlı mı (2, 3, 4, 5, 7, 9, 10 / 10); tevkif edilen KDV = KDV × oran |
| Hata | Satırlar toplamı = mal/hizmet toplamı; satır KDV'leri = toplam KDV; vergiler dahil = mal/hizmet + KDV; ödenecek = vergiler dahil − tevkifat |
| Hata | Mükerrer fatura (aynı satıcı ve fatura no birden fazla dosyada) |
| Hata / Dikkat | Eksik fatura no, tarih, satıcı veya alıcı VKN, ödenecek tutar |
| Dikkat | Döviz faturası (kur kontrolü); e-posta metninden okunan fatura (e-Fatura / e-Arşiv aslını kontrol edin) |

Fark sınıfları: 0,01 TL'ye kadar fark tamam sayılır, 0,05 TL'ye kadar "yuvarlama farkı" (Bilgi), daha büyüğü hatadır.
Fatura toplamlarında sınır satır sayısı × 0,01 TL'ye kadar genişler.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                    # örnek 3 fatura (2 hatalı, 1 uygun)
python agent.py --girdi ./gelen_faturalar          # klasördeki tüm .pdf / .docx / .txt
python agent.py --girdi fatura.pdf --gizle "Ayşe Deneme"
```

Ardından ayrıntılı KDV kontrolü için:

```bash
python main.py --satirlar cikti/faturalar.xlsx     # Fatura KDV Tutarlılık Kontrolü kod bloğu klasöründe
```

## Çıktı

`cikti/faturalar.xlsx`:
- `Faturalar`: fatura başına no, tarih, vade, VKN'ler, unvanlar, toplamlar, satıcı IBAN'ı, durum (Hatalı / Kontrol / Uygun)
  ve bulgular.
- `Satırlar`: mal/hizmet satırları.
- `Bulgular`: bütün bulgular tek listede, model notlarıyla birlikte.
- `Kontrol`: dosya başına düzeltme turu ve son denetim sonucu.

`Faturalar` ve `Satırlar` sayfalarının sütunları **E-Fatura Okuma ve Listeleme** çıktısıyla aynıdır; dosya **Fatura KDV
Tutarlılık Kontrolü** kod bloğuna doğrudan verilebilir. Mor hücreler yapay zekâ ile çıkarılmıştır.

Ayrıca aynı bilgiler `cikti/faturalar.json` dosyasına yazılır.

## Dikkat

- **Kontrol edin.** Muhasebeleştirmeden ve ödemeden önce her faturayı aslıyla karşılaştırın. Model satır kaçırabilir,
  sütunları karıştırabilir; kod, sayıların belgede yazdığını ve hesapların tuttuğunu kontrol eder ama satırın doğru
  sütuna yazıldığını her zaman anlayamaz.
- **Taranmış PDF** metin içermez; önce OCR uygulayın (ör. tarayıcınızın "aranabilir PDF" seçeneği).
- **Bu bir vergi kontrolü değildir.** KDV oranının mal/hizmete uygunluğu, tevkifat kapsamı ve indirim hakkı muhasebe
  sorumlusunun değerlendirmesidir. Oran listesi 10.07.2023'te yürürlüğe giren Cumhurbaşkanı Kararı'na göredir; değişebilir.
- **Kişisel veri.** Unvanlar, adresler, kalemler ve tutarlar modele gönderilir. Şahıs firması adları için `--gizle`
  kullanabilirsiniz.
- **Örnek veri:** Firmalar, numaralar ve tutarlar kurgusaldır. Örnekteki VKN ve TCKN'ler yalnız kontrol hanesi algoritmasına
  uyacak şekilde seçilmiştir.

## Testler

Testler gerçek API çağırmaz. Sahte model, araç bakım faturasında ilk çıkarımda bilerek:
- belgedeki yanlış satır tutarını (3.850,00) "düzeltip" 3.750,00 yazar,
- "30 gün vadeli" ifadesinden vade tarihi hesaplar,
- satıcı unvanını kısaltır,
- uydurma bir IBAN yazar.

Kodun bunları yakalayıp modele geri gönderdiği ve faturanın kendi hatalarını bulduğu doğrulanır:
- bakım faturasında satır hesabı ve geçersiz %18 oranı,
- danışmanlık faturasında alıcı VKN kontrol hanesi.

IBAN ve TCKN'nin modele gönderilen hiçbir mesajda geçmediği de doğrulanır.

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
- **Yapay zekâ hata yapabilir.** Model yanlış, eksik veya uydurma bilgi üretebilir. Kod kontrolleri bazı hataları
  yakalar, hepsini yakalayamaz. Her taslağı okuyup onaylamadan kullanmayın, göndermeyin, imzalamayın.
- **Maskeleme sınırlıdır.** Yalnız telefon, e-posta, TCKN, IBAN ve LinkedIn adresi kalıpları maskelenir. Ad soyad,
  adres, doğum tarihi, sağlık bilgisi ve serbest metindeki diğer kişisel veriler modele gidebilir. Hangi verinin bulut
  sağlayıcısına gönderildiğinden siz sorumlusunuz; gizli veride yerel model (Ollama) kullanın.
- **API ücretleri, kotalar ve API anahtarının güvenliği** size aittir.

Ayrıntılar: `LICENSE` (MIT) ve `SORUMLULUK_REDDI.md`.
