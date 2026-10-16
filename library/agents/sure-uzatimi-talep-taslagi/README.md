# Süre Uzatımı Talep Taslağı · AI Agent

> İnşaat Teknik Ofis › Teknik Ofis Mühendisi · Workers / Workless

Günlük şantiye raporlarındaki gecikme olaylarını sınıflar, **gün hesabını kodla yapar** ve idareye **süre uzatımı talep
yazısı** taslağı hazırlar. Yazı yalnız kodun hesapladığı günlerle ve sizin verdiğiniz sözleşme maddeleriyle yazılır;
uydurma madde ya da fazladan gün kabul edilmez.

## Nasıl çalışır?

1. **Hazırlık (kod):** Günlük raporlar okunur (her biri `Tarih: GG.AA.YYYY` satırı içerir). Telefon, e-posta ve kimlik
   numarası maskelenir.
2. **Sınıflama (model):** Her rapordaki gecikme olayları sınıflanır:
   - neden türü: hava, idare kaynaklı, mücbir sebep, yer teslimi / engel, malzeme / tedarik, yüklenici kaynaklı, İSG,
     diğer;
   - durma derecesi (tam / kısmi / yok) ve etkilenen iş;
   - rapordan aynen alıntı (kod denetler).
3. **Gün hesabı (kod):**
   - Her gün bir kez sayılır: aynı gün iki talep nedeni varsa gün tektir.
   - Talep edilebilir nedenlerle (varsayılan: hava, idare kaynaklı, mücbir sebep, yer teslimi / engel; `--talep-edilebilir`
     ile değişir) iş durmuşsa gün talebe girer. O gün tam mı kısmi mi durduğu ayrıca gösterilir.
   - Yalnız yüklenici / tedarik kaynaklı olay olan günler talep dışı bırakılır.
   - Talep edilebilir bir olayla aynı gün yüklenici kaynaklı bir olay da varsa gün **eşzamanlı gecikme** olarak
     işaretlenir. İdare bu günlere itiraz edebilir; yazıda gizlenmez.
   - `--basvuru-gun` verilirse son başvuru tarihi = ilk olay tarihi + başvuru süresi.
4. **Yazı (model):** Kodun özet tablosu ve sözleşme maddeleriyle yazının gövdesi yazılır.
5. **Denetim (kod):** Aşağıdakilerden biri varsa model düzeltir (varsayılan en fazla 2 tur):
   - Alıntı raporda aynen geçmiyor
   - Yazıda özette ya da raporlarda olmayan bir sayı / tarih var
   - Sözleşme metninde olmayan bir madde var
   - Talep edilen toplam gün gövdede yok
   - Tehdit ya da kesinlik ifadesi var

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                     # örnek 5 günlük rapor (4 gün talep, 1 eşzamanlı gün)
python agent.py --raporlar ./gunluk_raporlar --sozlesme maddeler.txt --is "B Blok kaba inşaat" --idare "Kurgu Belediyesi Fen İşleri Müdürlüğü" --basvuru-gun 20
python agent.py --raporlar ./gunluk_raporlar --sozlesme maddeler.txt --talep-edilebilir "İdare kaynaklı" "Mücbir sebep"
```

Sözleşme metnine kendi sözleşmenizdeki ve ilgili genel şartnamedeki süre uzatımı maddelerini kopyalayın. Agent mevzuatı
kendisi bilmez; örnek maddeler kurgusaldır.

Günlük raporları **Günlük Şantiye Raporu** agent'ıyla hazırlıyorsanız her raporun başına `Tarih:` satırının
yazıldığından emin olun.

## Çıktı

- `cikti/sure_uzatimi.xlsx`:
  - `Olaylar`: tarih, rapor, neden, durma, etkilenen iş, açıklama, alıntı.
  - `Gün Hesabı`: gün başına talep edilebilir mi, duruş, nedenler, eşzamanlı gecikme, not; toplam ve son başvuru tarihi.
  - `Yazı`: gönderime hazır yazı (konu, gövde, ekler, imza).
  - `Kontrol`
- `cikti/sure_uzatimi.txt`: yazı.

## Dikkat

- **Süre uzatımı hakkı sözleşmenize bağlıdır.** Hangi nedenin süre uzatımı gerekçesi sayılacağını, başvuru süresini ve
  biçimini sözleşmeniz ve ilgili şartname belirler. Kodun gün hesabı bir ön hesaptır; idare farklı değerlendirebilir.
- **Kısmi duruş günleri** idare tarafından tam gün sayılmayabilir. Etkilenen işin kritik yolda olup olmadığını **İş
  Programı / Kritik Yol Analizi** ile gösterin.
- **Yazıyı göndermeden önce** teknik ofis şefi ve hukuk birimiyle kontrol edin; ekleri (raporlar, idare yazıları,
  fotoğraflar) mutlaka ekleyin.
- **Örnek veri:** İş, idare, sözleşme maddeleri ve olaylar kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model yazının ilk taslağında bilerek sözleşmede olmayan "Madde 30.3"e atıf yapar.
Kodun bunu yakaladığı doğrulanır. Ayrıca şunlar doğrulanır:
- gün hesabı: aynı gün bir kez, eşzamanlı gecikme, talep dışı günler, son başvuru tarihi,
- uydurma tarih / gün ve tehdit ifadesi yakalama,
- satır başı liste numaralarının sayı sayılmaması.

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
