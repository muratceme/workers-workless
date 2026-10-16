# Serbest Tariften Standart Reçete Kartı · AI Agent

> Otel › Yiyecek-İçecek › Aşçıbaşı · Workers / Workless

Şeflerin serbest yazdığı tarifleri ("5 su bardağı mercimek, 3 yemek kaşığı tereyağı …") **standart reçete kartına**
dönüştürür:
- malzeme, fiyat listesindeki adla eşleştirilmiş,
- g / ml cinsinden standart miktar ve porsiyon başı miktar (kod hesabı),
- hazırlanış adımları,
- 14 alerjen grubuna göre alerjenler.

Çıktı **Porsiyon Maliyeti Hesaplama** kod bloğuna verilebilir.

## Nasıl çalışır?

1. **Hazırlık (kod):** Tarifler, fiyat listesi (yalnız malzeme adları gönderilir) ve ölçü dönüşüm tablosu okunur.
2. **Çıkarma (model):** Malzemeler tarifte yazan miktar ve ölçüyle çıkarılır, adımlar ve alerjenler yazılır. Model grama
   çevirmez.
3. **Denetim (kod):** Aşağıdakilerden biri varsa yalnız o tarif modele geri gönderilir (varsayılan en fazla 2 tur):
   - Tarifteki bir malzeme satırı alınmamış; kanıt tarifte aynen yok
   - Miktar ya da ölçü kanıtta yazmıyor; fire tarifte yazmıyor
   - Eşleştirilen ad tarifteki adla uyuşmuyor (tereyağı → ayçiçek yağı gibi)
   - Adımlarda tarifte olmayan süre / sıcaklık
   - **Kodun anahtar kelimelerle bulduğu alerjen** (un → gluten, tereyağı → süt, tahin → susam …) modelin listesinde yok
4. **Çevrim (kod):**
   - g, kg, ml, l doğrudan çevrilir.
   - Ev ölçüleri (su bardağı, yemek kaşığı, adet, tutam …) dönüşüm tablosundan, önce malzemeye özel satır, yoksa `*`
     satırıyla çevrilir.
   - Fiyat listesinde birimi "adet" olan malzeme adet kalır.
   - Çevrilemeyen satır işaretlenir ve kod bloğu dosyasına yazılmaz.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                              # örnek 2 tarif
python agent.py --tarifler tarifler.docx --fiyatlar fiyatlar.xlsx --donusum olcu_donusum.xlsx
```

Tarif örneği:

```
Tarif: Mercimek Çorbası | Porsiyon: 10
5 su bardağı kırmızı mercimek
3 yemek kaşığı tereyağı
Hazırlanışı: Soğanları doğrayıp tereyağında kavurun. ...
```

## Çıktı

- `cikti/recete_kartlari.xlsx`:
  - `Reçete Kartları`: tarifteki ve standart miktar, birim, porsiyon başı, fiyat listesinde olup olmadığı, çevrim notu
    (mor: model, mavi: kod).
  - `Hazırlanış ve Alerjen`: adımlar, modelin ve kodun bulduğu alerjenler, şefe sorulacaklar.
  - `Kontrol`
- `_receteler.xlsx` (Menü Kalemi; Malzeme; Miktar; Birim; Fire %) ve `_menu.xlsx` (porsiyon sayısıyla; satış fiyatını
  siz girin):

```bash
python ../../code-blocks/porsiyon-maliyeti-hesaplama/main.py --receteler cikti/recete_kartlari_receteler.xlsx --fiyatlar fiyatlar.xlsx --menu cikti/recete_kartlari_menu.xlsx
```

## Dikkat

- **Ölçü dönüşümleri yaklaşıktır.** `ornek_veri/olcu_donusum.csv` örnek değerler içerir. Bardak, kaşık ve "orta boy"
  malzemelerin ağırlığı mutfaktan mutfağa değişir. Standart reçeteyi tartarak doğrulayın ve tabloyu kendi
  değerlerinizle güncelleyin.
- **Alerjen bilgisi tarifteki adlara dayanır.** 14 alerjen grubu AB 1169/2011 Ek II ile Türk Gıda Kodeksi etiketleme
  yönetmeliğindeki gruplardır. Hazır ürünlerin (sos, baharat karışımı, ekmek, bulyon) etiketlerindeki alerjenleri ve
  çapraz bulaşmayı ayrıca kontrol edin. Misafire alerjen bilgisi vermeden önce sorumlu kişi onaylamalıdır.
- **Örnek veri:** Tarifler, fiyatlar ve dönüşüm değerleri kurgusal / yaklaşıktır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk turda mercimek çorbasındaki tereyağına rağmen süt alerjenini yazmaz.
Kodun bunu yakalayıp düzelttirdiği doğrulanır. Ayrıca şunlar doğrulanır:
- yanlış eşleştirme, miktar / ölçü, atlanan satır ve adımlarda uydurma süre denetimleri,
- alerjen anahtar kelimeleri,
- ölçü çevrimi (malzemeye özel, `*`, doğrudan, adet, eksik dönüşüm),
- kod bloğu dosyaları,
- fiyatların gönderilmemesi.

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
