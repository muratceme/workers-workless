# Banka Mutabakatı · AI Agent

> Finans & Muhasebe › Muhasebe Elemanı · Gerisi Bende

Mutabakatın **eşleştirme kısmını yapay zekâya bırakmaz**: kod bloğunun deterministik motoru
çalışır, sonuç her seferinde aynıdır. Yapay zekâ yalnızca eşleşmeyen **açık kalemleri**
yorumlar:

- olası neden (masraf, faiz, zamanlama, tutar hatası, mükerrer kayıt…)
- karşı taraftaki ilişkili kalem
- düzeltme için **yevmiye kaydı taslağı** (ör. `780 Finansman Giderleri (B) / 102 Bankalar (A)`)
- güven düzeyi ve genel değerlendirme

Raporda yapay zekâ sütunları mor, **İnsan Onayı** sütunu sarıdır.

## Gizlilik

- Eşleşen kayıtlar **hiç gönderilmez**; yalnızca açık kalemlerin tarih, açıklama ve tutarı gider.
- Açıklamalardaki IBAN, TCKN, e-posta ve telefon maskelenir.
- Göndermeden önce onayınız istenir. Veri hiç çıkmasın istiyorsanız `WW_PROVIDER=ollama`.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env
```

`.env` içine kendi API anahtarınızı yazın (bkz. `.env.example`). API ücretleri size aittir.

## Kullanım

```bash
# Örnek veriyle deneyin
python agent.py

# Kendi dosyalarınız
python agent.py --banka ekstre.xlsx --defter muavin_102.xlsx --cikti cikti/mutabakat_ai.xlsx
```

Girdi standardı kod bloğuyla aynıdır: `tarih`, `aciklama` ve `tutar` (giriş +, çıkış −) ya da
`borc` + `alacak` sütunları; `.csv` veya `.xlsx`. Ayrıntı ve şablonlar: `ornek_veri/`.

| Parametre | Açıklama |
|---|---|
| `--banka`, `--defter` | Girdi dosyaları |
| `--cikti` | Excel çıktısı (varsayılan `cikti/mutabakat_ai.xlsx`) |
| `--gun-toleransi` | Tutar+tarih eşleşmesinde gün farkı (varsayılan 3) |
| `--evet` | Onay sorusunu atla |

## Testler

Testler gerçek API'yi çağırmaz:

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
