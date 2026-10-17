# Toplu Ödeme Dosyası Kontrolü · Kod Bloğu

> Bankacılık Operasyonları › Ödeme Sistemleri Operasyon Uzmanı · Gerisi Bende

Müşterilerin yüklediği toplu ödeme (maaş, tedarikçi) dosyalarını **işleme almadan önce** kontrol eder: IBAN, tutar,
mükerrer satır, kontrol toplamı ve önceki dosyaya göre değişiklikler. İnternete bağlanmaz.

## Nasıl hesaplar?

| Kontrol | Kural | Sonuç |
|---|---|---|
| IBAN biçimi | Boşluk ve küçük harf temizlenir. TR IBAN 26 karakter, ülke kodundan sonra yalnız rakam, 11. karakter (rezerv) `0` | Hatalı |
| IBAN kontrol basamağı | ISO 13616 mod 97: ilk 4 karakter sona alınır, harfler A=10 … Z=35 yapılır, sayının 97'ye bölümünden kalan 1 olmalı | Hatalı |
| Yabancı IBAN | Mod 97 ile doğrulanır ve ayrıca işaretlenir | Kontrol edin |
| Tutar | Sıfırdan büyük, en çok 2 ondalık | Hatalı |
| Tek ödeme sınırı | `--ust-sinir` verilirse | Kontrol edin |
| Alıcı adı | Boş veya 3 karakterden kısa | Hatalı |
| TCKN / VKN | Verilmişse kontrol basamakları (TCKN 11 hane, VKN 10 hane) | Kontrol edin |
| Mükerrer ödeme | Aynı IBAN ve aynı tutar | Hatalı |
| Aynı IBAN, farklı alıcı · Aynı kimlik no, farklı IBAN | | Kontrol edin |
| Gönderen hesabına ödeme | `--gonderen-iban` ile aynı IBAN | Hatalı |
| Kontrol toplamı | `--beklenen-adet`, `--beklenen-toplam` ve dosyadaki `TOPLAM` satırı | İşleme alınamaz |

**Önceki dosyayla karşılaştırma** (`--onceki`): alıcılar önce kimlik numarasıyla, yoksa adla eşleştirilir.

| Bulgu | Anlamı |
|---|---|
| IBAN değişmiş | Aynı alıcının IBAN'ı önceki dosyadan farklı. Sahte IBAN değişikliği talebi yaygın bir dolandırıcılık yöntemidir; bağımsız bir kanaldan teyit edilmelidir |
| Tutar belirgin değişmiş | Değişim `--degisim-esik` (varsayılan %25) üzerinde |
| Yeni alıcı · Bu dosyada olmayan alıcı | Bilgi |
| Dosya önceki dosyanın birebir aynısı | Yanlış dosya yüklenmiş olabilir |

Bir satırda "Hatalı" bulgu veya kontrol toplamı farkı varsa sonuç **İşleme alınamaz** olur.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek veri
python main.py --dosya odemeler.xlsx --beklenen-adet 120 --beklenen-toplam 4.250.000,00
python main.py --dosya odemeler.xlsx --onceki gecen_ay.xlsx --gonderen-iban "TR.." --ust-sinir 500.000,00
```

Dosya sütunları: **IBAN** ve **Tutar** zorunlu; Alıcı Adı, TCKN / VKN, Açıklama isteğe bağlı. Son satır `TOPLAM` ile
başlıyorsa ödeme sayılmaz, tutarı dosya toplamıyla karşılaştırılır.

## Çıktı

`cikti/toplu_odeme_kontrolu.xlsx`:

- `Özet`: sonuç, satır ve tutar toplamları, hatalılar hariç toplam, bankalara göre dağılım.
- `Satır Kontrolü`: her satırın durumu (Hatalı / Kontrol edin / Temiz) ve bulguları.
- `Bulgular`: önem sırasıyla tüm bulgular; boş "Müşteriye Bildirildi" sütunu.

## Dikkat

- **Biçim kontrolüdür.** Kod IBAN'ın alıcıya ait olduğunu, hesabın açık olduğunu veya alıcı adının hesap sahibiyle
  uyuştuğunu bilemez. Bunlar ancak bankanın sistemlerinde doğrulanır.
- **Dosyayı değiştirmez.** Hatalı satırları düzeltmez veya çıkarmaz; düzeltme müşteriden yeni dosya veya yazılı talimat
  ile yapılmalıdır.
- **Banka listesi** (IBAN'ın 5-9. karakterleri) bilgi amaçlıdır ve eksik olabilir. Listede olmayan kodlar "Diğer /
  tanımsız" görünür; bu bir hata değildir.
- **VKN algoritması** yaygın kullanılan kontrol basamağı hesabıdır; şüpheli durumda vergi dairesi kaydı esastır.
- **Kişisel veri:** Dosya ad, IBAN ve kimlik numarası içerir. Çıktıyı yalnız yetkili kişilerle paylaşın; e-postayla
  göndermeden önce bankanızın veri güvenliği kurallarına bakın.
- **Örnek veri:** Kişiler, IBAN'lar ve kimlik numaraları rastgele üretilmiştir; gerçek kişilere ait değildir.

## Testler

- IBAN: mod 97 (TR, GB, DE örnekleri), uzunluk, rezerv alan, harf, boşluk temizliği.
- TCKN ve VKN kontrol basamakları.
- Satır hataları, mükerrerler, kontrol toplamı, önceki dosya karşılaştırması.
- Örnek veri uçtan uca (5 hatalı satır, toplam farkı) ve komut satırı.

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
