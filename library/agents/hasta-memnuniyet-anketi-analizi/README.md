# Hasta Memnuniyet Anketi Analizi · AI Agent

> Hasta Hizmetleri › Hasta İlişkileri ve Hasta Hakları Yöneticisi · Workers / Workless

Hasta memnuniyet anketlerinin puanlarından **birim ve hekim bazında memnuniyet raporu** çıkarır. Açık uçlu yorumları
temalara ayırır ve birim bazında **iyileştirme başlıkları** önerir.

## Nasıl çalışır?

1. **Puanlar (kod):** 1–5 puanlı soru sütunları otomatik bulunur.
   - Birim ve hekim bazında yanıt sayısı, soru ortalamaları ve genel ortalama hesaplanır.
   - Memnuniyet oranı: 4 ya da 5 verilen yanıtların tüm yanıtlara oranı.
   - Ortalaması `--hedef`in altındaki birim × soru hücreleri "düşük" sayılır (varsayılan 4,0).
   - `--min-n`den az yanıtı olan hekimler değerlendirmeye alınmaz (varsayılan 10).
2. **Yorumlar (model):** Yalnız yorumlar ve birim puan ortalamaları gönderilir. Yorumdaki hekim adları [HEKİM] yapılır.
   - Her yorum temaya ve duyguya ayrılır, alıntıyla desteklenir.
   - Birim bazında iyileştirme başlıkları önerilir: gerekçe, dayanak yorumlar, ilgili düşük puanlı soru ve öneri.
3. **Denetim (kod):** Aşağıdakilerden biri varsa taslak modele geri gönderilir (varsayılan en fazla 2 tur):
   - Alıntı yorumda yok
   - Başlığın dayanağı olmayan ya da başka birime ait yorum
   - Tek yoruma dayanan ve o birimin düşük puanlı sorularıyla ilgisi olmayan başlık

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                        # örnek anketler
python agent.py --girdi anketler.xlsx
python agent.py --girdi anketler.xlsx --hedef 4.2 --min-n 20
```

Soru sütunlarının başlığı soru adı olmalı (ör. "Hekim İlgisi", "Bekleme Süresi") ve değerleri 1–5 arasında olmalıdır.
"Birim" sütunu zorunludur.

## Çıktı

- `cikti/memnuniyet_raporu.xlsx`:
  - `İyileştirme Başlıkları`: "Sorumlu" ve "Termin" sütunları boş bırakılır.
  - `Birim Özeti`: düşük puanlı soru hücreleri kırmızı.
  - `Hekim Özeti`: yetersiz örnekli hekimlerin puanları gösterilmez.
  - `Yorumlar`
  - `Tema Dağılımı`

## Dikkat

- **Küçük örneklem.** Az yanıtla hesaplanan ortalamalar yanıltıcıdır. Hekim bazında sonuçları performans değerlendirmesinde
  tek başına kullanmayın; `--min-n` değerini kurumunuzun ölçme yöntemine göre ayarlayın.
- **Anket yöntemi.** Bu araç anketi tasarlamaz, örnekleme yapmaz ve resmî bir ölçüm yöntemini uygulamaz. Kurumunuz bir
  ölçüm yöntemi kullanıyorsa (ör. Sağlıkta Kalite Standartları kapsamında) hesaplamayı onunla karşılaştırın.
- **Kişisel veri.** Modele yalnız yorum metni ve birim adı gider; hekim adları gizlenir, telefon / e-posta maskelenir. Yorumda
  geçen hasta ya da personel adları gizlenmez.
- **Örnek veri** kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk taslakta bilerek başka birime ait bir yorumu dayanak gösterir, tek yoruma
dayanan bir başlık ekler ve bir alıntıyı değiştirir. Kodun bunları yakalayıp düzelttirdiği, puanları doğru hesapladığı ve
hekim adlarını modele göndermediği doğrulanır.

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
