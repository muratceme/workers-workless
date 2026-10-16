# Tahsis Kararından Kullandırım Şartları Çıkarma · AI Agent

> Bankacılık Operasyonları › Kredi Operasyon Uzmanı · Workers / Workless

Kredi komitesi / tahsis kararlarının serbest metnini okur. Kullandırım kontrolüne girecek bilgileri tabloya döker:
- onay bilgileri: limit, ürünler, azami vade, asgari faiz, geçerlilik, sözleşme tutarı,
- her teminat, ön koşul, belge ve taahhüt şartı: zamanlama, gereken tutar, termin, kararda geçen cümlesi.

Çıktı doğrudan **Kredi Kullandırım Ön Kontrolü** kod bloğuna verilebilir.

## Nasıl çalışır?

1. **Hazırlık (kod):** Kararlar okunur (dosya ya da klasör; bir dosyada birden çok karar olabilir, "Karar No:" satırı
   yeni kararı başlatır). Telefon, e-posta, TCKN, IBAN ve `--gizle` ile verilen kefil / ortak adları maskelenir.
2. **Çıkarma (model):** Onay bilgileri ve şartlar yalnız kararda yazandan alınır; her şart için kararın ilgili cümlesi
   kelimesi kelimesine yazılır. Zamanlaması yazmayan şart ihtiyatla "Kullandırım öncesi" sayılır ve açık nokta olarak
   listelenir.
3. **Denetim (kod):** Aşağıdakilerden biri varsa yalnız o karar modele geri gönderilir (varsayılan en fazla 2 tur):
   - Kararda olmayan tutar, oran, vade ya da tarih
   - Kararda bu adla geçmeyen ürün ya da müşteri numarası
   - Kaynak cümlenin kararda aynen geçmemesi
   - Teminat / şart içeren bir cümlenin (ipotek, rehin, kefalet, temlik, sigorta, ibraz, "alınacak", "tesis edilecek" …)
     hiçbir şarta alınmamış olması
   - Zamanlamanın cümleyle çelişmesi ("takiben 30 gün içinde" → kullandırım sonrası)
   - Gereken tutarın cümlede olmaması; cümlede tutar ya da tarih varken alanın boş kalması
4. **Hesap (kod):** "Limitin %20'si", "kredi tutarının %80'i" gibi yüzdeli teminat tutarlarını kod hesaplar. Sözleşme
   tutarı yoksa ya da limitin altındaysa, geçerlilik tarihi yoksa uyarı verir.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                              # örnek 2 karar (8 şart)
python agent.py --kararlar kararlar/ --gizle "Hasan Demirtaş" "Selin Demirtaş"
```

Karar örneği:

```
Karar No: ON-6101
Firmaya toplam 25.000.000 TL tutarında nakdi kredi limiti tahsis edilmiştir. ...
1. Bursa Nilüfer'deki fabrika binası üzerinde 1. derece 30.000.000 TL ipotek tesis edilmeden kullandırım yapılmayacaktır.
3. Fabrika binasının yangın sigortası poliçesi ... kullandırımı takiben 30 gün içinde ibraz edilecektir.
```

## Çıktı

`cikti/tahsis_sartlari.xlsx`:
- `Onaylar`: onay bilgileri (Kullandırım Ön Kontrolü sütun düzeninde), müşteri, kaynak dosya, son kontrol.
- `Şartlar`: şartlar, kararda geçen cümle ve kodun yüzde hesabı. Mor hücreler modelin, mavi hücreler kodun değeridir.
- `Açık Noktalar`: tahsis birimine sorulacaklar (modelin belirsizlikleri ve kodun uyarıları).
- `Kontrol`

Ayrıca `tahsis_sartlari_onaylar.xlsx` ve `tahsis_sartlari_sartlar.xlsx`. **Mevcut Risk, Sözleşme Durumu, Tesis Edilen
Tutar, Durum ve Tamamlanma Tarihi** sütunlarını kendi kayıtlarınızdan doldurduktan sonra kod bloğuna verin:

```bash
python ../../code-blocks/kredi-kullandirim-on-kontrolu/main.py --talepler talepler.xlsx --onaylar cikti/tahsis_sartlari_onaylar.xlsx --sartlar cikti/tahsis_sartlari_sartlar.xlsx
```

## Dikkat

- **Kararın aslı esastır.** Her satırı karar aslıyla karşılaştırın. Agent kararı yorumlamaz; kararda yazmayan bir şartı
  bilemez. Şartlar arası bağlar ("yalnız rotatif kullandırımlarda") tabloda tek sütuna sığmaz; açık noktalara bakın.
- **Yüzdeli tutarlar limit üzerinden hesaplanır.** Karar kullandırım tutarına oran veriyorsa tutar kullandırım bazında
  değişir.
- **Teminatın hukuken geçerli tesisini** (ipotek derecesi, rehin sicili, kefalet sözleşmesinin şekli) bu agent
  doğrulamaz.
- **Örnek veri:** Firmalar, kişiler ve kararlar kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk turda ON-6101'in 4. maddesini (müşteri çeki temliki) atlar. Kodun atlanan
cümleyi yakalayıp düzelttirdiği doğrulanır. Ayrıca şunlar doğrulanır:
- kararda olmayan sayı, ürün ve kaynak cümle,
- zamanlama ve tutar / termin denetimleri,
- yüzdeli tutar hesabı ve sözleşme uyarıları,
- kefil adlarının modele gitmemesi,
- kod bloğu dosyalarının sütun düzeni.

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
