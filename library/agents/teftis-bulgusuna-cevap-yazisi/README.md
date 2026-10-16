# Teftiş Bulgusuna Cevap Yazısı · AI Agent

> Banka Şube › Şube Operasyon Yöneticisi · Workers / Workless

Teftiş / iç kontrol bulgu listesini ve şubenin her bulgu için tuttuğu kısa notları okur. Her bulgu için şunları yazar:
- durum (Tamamlandı / Devam ediyor / Planlandı / İtiraz),
- kök neden,
- alınan ve planlanan aksiyonlar,
- termin ya da kapanış tarihi,
- teftiş birimine gidecek resmî cevap paragrafı.

Bulgu tablosundaki Termin, Kapanış Tarihi ve Aksiyon Sahibi sütunlarını doldurur. Çıktı doğrudan
**Kontrol / Bulgu Takibi** kod bloğuna verilebilir.

## Nasıl çalışır?

1. **Hazırlık (kod):** Bulgu listesi ve şube notları okunur. Notu olmayan bulgu modele gönderilmez, ayrıca listelenir.
   Telefon, e-posta, TCKN, IBAN ve `--gizle` ile verilen adlar maskelenir.
2. **Cevap (model):** Yalnız şube notunda yazan aksiyon, tarih ve sayılarla cevap yazılır.
3. **Denetim (kod):** Aşağıdakilerden biri varsa yalnız o bulgu modele geri gönderilir (varsayılan en fazla 2 tur):
   - Notta olmayan tarih ya da sayı (bulgudaki sayılar serbest)
   - Nottaki bir tarihin cevapta hiç geçmemesi (o tarihli iş atlanmış olabilir)
   - Notta "planlıyor / yapılacak / tamamlanacak" gibi henüz yapılmamış iş varken durumun "Tamamlandı" olması ya da
     planlanan aksiyonun boş kalması
   - "Tamamlandı" için notta tamamlandı ifadesi ya da kapanış tarihi olmaması
   - "Devam ediyor / Planlandı" için termin olmaması; terminin bulgu açılışından önce olması
4. **Yazı (kod):** Cevaplar tek bir yazıda birleştirilir; hitap, konu ve imza kod tarafından eklenir.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                    # örnek 3 bulgu (2 devam ediyor, 1 tamamlandı)
python agent.py --bulgular bulgular.xlsx --notlar sube_notlari.txt --sube "Kadıköy Şubesi"
python agent.py --gizle "Ayşe Yılmaz" "Mehmet Kaya"                 # notlarda geçen adlar gönderilmez
```

Şube notu örneği:

```
Bulgu: T-26-14
6 işlemin 4'ünde onay sözlü alınmış ama sisteme girilmemiş. Personele 30.09.2026'da eğitim verildi.
Onay zorunlu alan için BT'ye talep açıldı, BT 30.11.2026'ya kadar devreye almayı planlıyor.
```

## Çıktı

`cikti/bulgu_cevaplari.xlsx`:
- `Bulgular`: girdideki takip tablosu (Termin, Kapanış Tarihi, Aksiyon Sahibi güncellenmiş) ve yanında durum, kök neden,
  alınan / planlanan aksiyonlar, cevap. Model sütunları mor.
- `Cevap Yazısı`: teftiş birimine gidecek yazının taslağı.
- `Kontrol`: denetim turları, kalan sorunlar, notu olmayan bulgular.

Kod bloğuyla zincir:

```bash
python ../../code-blocks/kontrol-bulgu-takibi/main.py --bulgular cikti/bulgu_cevaplari.xlsx
```

## Dikkat

- **Cevabın doğruluğu şubenindir.** Agent yalnız nottaki bilgiyi düzenler; notta yazmayan bir aksiyonu "yapıldı"
  diye yazdırmayın. Gönderilmeden önce yetkili kişi okuyup onaylamalıdır.
- **Termin taahhüttür.** Yazıdaki her tarih teftiş biriminde takip edilir; tutturulamayacak tarihi notta yazmayın.
- **Banka içi yönetmelik** (cevap süresi, yazı biçimi, onay zinciri) bankadan bankaya değişir; kod bunları bilmez.
- **Örnek veri:** Şube, bulgular ve talep numaraları kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk turda T-26-14'ü "Tamamlandı" sayar ve BT'nin 30.11.2026'ya kadar
planladığı işi atlar. Kodun eksik tarihi ve "planlıyor" kuralını yakalayıp düzelttirdiği doğrulanır. Ayrıca şunlar
doğrulanır:
- notta olmayan sayı ve kapanış tarihi denetimleri,
- termin kuralları,
- notu olmayan bulgunun ayrılması,
- takip tablosunun kod bloğunun sütun sırasıyla yazılması.

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
