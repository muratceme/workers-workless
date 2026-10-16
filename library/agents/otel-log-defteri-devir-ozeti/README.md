# Otel Log Defteri Devir Özeti · AI Agent

> Otel Ön Büro › Night Auditor · Workers / Workless

Gece ya da vardiya log defteri kayıtlarını gelen vardiyanın ve departman sorumlularının 2 dakikada okuyacağı **devir
notuna** çevirir. Not şunları içerir:
- açık konular (departman, oda, sorumlu, takip saati, tutar),
- kapananlar,
- misafir şikâyetleri,
- kasa / ödeme konuları.

Sorumlusu yazılmamış açık konuları kod "SORUMLU YOK — atanmalı" diye işaretler; hiçbir log kaydı atlanamaz.

## Nasıl çalışır?

1. **Hazırlık (kod):** Log dosyası okunur: ilk satır `Vardiya: …`, kayıtlar `SS:DD metin`. Saatsiz satırlar önceki
   kaydın devamı sayılır. Telefon, e-posta, kimlik numarası ve `--gizle` ile verilen misafir adları maskelenir.
2. **Devir (model):** Kayıtlar maddelere ayrılır. Her madde için:
   - departman, oda, kısa özet;
   - durum (Açık / Kapandı / Bilgi);
   - logda yazıyorsa sorumlu, takip saati ve tutar;
   - öncelik ve misafir şikâyeti olup olmadığı;
   - dayandığı log satırları.
3. **Denetim (kod):** Aşağıdakilerden biri varsa model düzeltir (varsayılan en fazla 2 tur):
   - Bir log satırı hiçbir maddede kullanılmamış
   - Oda no, takip saati ya da tutar dayanak satırlarda yazmıyor; özette logda olmayan sayı var
   - Dayanak satırda acil işaret var ama açık konu "Acil" değil. Acil işaretler: VIP, transfer, teyit edilmedi, yangın,
     duman, su baskını, ambulans, sağlık, yaralı, kavga, kayıp eşya / çocuk, polis.
4. **Sıralama (kod):** Açık konular önce acil, sonra departman sırasıyla listelenir. Sorumlusu olmayanlar işaretlenir.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                     # örnek gece log defteri (8 kayıt)
python agent.py --girdi log_15_10.txt --gizle "Ayşe Deneme"
```

Log örneği:

```
Vardiya: 15.10.2026 23:00 - 16.10.2026 07:00 | Ön Büro - Gece
23:10 512 numaralı oda klimanın çalışmadığını bildirdi. Teknik nöbetçi filtreyi temizledi, çalıştı.
04:40 VIP misafir (oda 801) 08:00'de havalimanı transferi istedi; transfer firması teyit etmedi.
```

## Çıktı

- `cikti/devir_notu.xlsx`:
  - `Devir`: madde başına öncelik, durum, departman, oda, konu, sorumlu, takip saati, tutar, şikâyet, dayanak log
    satırları.
  - `Kontrol`
- `cikti/devir_notu.txt`: gelen vardiyanın okuyacağı devir notu.

## Dikkat

- **Log defteri asıl kayıttır.** Devir notu okumayı kolaylaştırır; log defterinin yerine geçmez.
- **Sorumlu tahmin edilmez.** Logda kimin yapacağı yazmıyorsa konu "sahipsiz" kalır; gelen vardiya atamalıdır.
- **Kişisel veri.** Oda numaraları ve log metni modele gider. Misafir adlarını `--gizle` ile verin ya da loga
  yazmayın (oda no yeterlidir).
- **Örnek veri:** Otel, misafirler, gruplar ve olaylar kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk turda bilerek:
- VIP transfer konusunu "Normal" önceliğe koyar,
- 06:10 halı kaydını atlar.

Kodun ikisini de yakaladığı doğrulanır. Ayrıca şunlar doğrulanır:
- oda, saat ve tutar denetimi,
- saatsiz satırın önceki kayda eklenmesi,
- misafir adının maskelenmesi,
- devir notunda acil ve sahipsiz konuların başta olması.

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
