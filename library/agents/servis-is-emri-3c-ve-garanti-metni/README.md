# Servis İş Emri 3C ve Garanti Metni · AI Agent

> Otomotiv Servis › Servis Danışmanı · Workers / Workless

Müşterinin şikâyeti ve teknisyenin notlarından iş emrine ve garanti başvurusuna girecek **3C metnini** hazırlar:
- **Şikâyet** (Complaint),
- **Neden** (Cause),
- **Çözüm** (Correction).

Ayrıca değişen parçaları (parça no ile) ve işçilik süresini çıkarır. Garanti süresini ve km sınırını kod hesaplar;
darbe, kaza gibi garanti dışı işaretleri ayrıca gösterir.

## Nasıl çalışır?

1. **Hazırlık (kod):** İş emirleri okunur. Telefon, e-posta ve kimlik numarası maskelenir; plaka sağlayıcıya gönderilmez.
   Garanti koşulları tablosu okunur.
2. **3C (model):** Şikâyet yalnız müşterinin anlattığı belirtidir; neden ve çözüm teknisyen notuna dayanır. Ayrıca
   parçalar, işçilik ve garanti önerisi (Garanti başvurusu / Ücretli / Karar için bilgi gerekli) yazılır.
3. **Denetim (kod):** Aşağıdakilerden biri varsa yalnız o iş emri modele geri gönderilir (varsayılan en fazla 2 tur):
   - Parça numarası teknisyen notunda aynen geçmiyor ya da hiç yok
   - İşçilik süresi notta yazmıyor
   - Şikâyette müşterinin söylemediği bir sayı, neden / çözümde teknisyenin yazmadığı bir sayı var
   - Şikâyet bölümüne teknisyen bulgusu karışmış (ör. "zincir gergisi kaynaklı")
4. **Garanti kontrolü (kod):**
   - Trafiğe çıkış + garanti yılı < iş emri tarihi ya da km > garanti km'si → **Garanti dışı**.
   - Teknisyen notunda garanti dışı işaret varken (darbe, çarpma, kaza, yanlış yakıt, modifiye, sıvı / su girişi, bakım
     yapılmamış, aşırı yük, yangın, sel) garanti önerilmişse → **Kontrol**.
   - Model, trafiğe çıkış tarihi ya da iş emri tarihi eksikse süre hesaplanmaz ve uyarı verilir.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                    # örnek 2 iş emri (1 garanti, 1 darbe → ücretli)
python agent.py --girdi is_emirleri.txt --garanti garanti_kosullari.csv
```

İş emri örneği:

```
İş Emri: IE-2026-1023 | Plaka: 34 ABC 123 | Model: Kurgu X1 1.5 | Trafiğe çıkış: 12.03.2024 | KM: 48.200 | Tarih: 15.10.2026
Müşteri: Sabah soğukta ilk çalıştırınca motordan 2-3 saniye tıkırtı geliyor.
Teknisyen: Eksantrik zincir gergisi basınç tutmuyor. Zincir gergisi (parça no 13540-KX100) değiştirildi. İşçilik 2,5 saat.
```

## Çıktı

`cikti/is_emri_3c.xlsx`:
- `İş Emirleri`: iş emri başına şikâyet, neden, çözüm, değişen parçalar, işçilik, model önerisi ve gerekçesi, kodun
  garanti durumu (Garanti kapsamında / Kontrol / Garanti dışı / Ücretli), kod bulguları, eksik bilgiler.
- `Kontrol`

## Dikkat

- **Garanti kararı garanti uzmanınındır.** Üreticinin / distribütörün garanti politikası (parça bazlı garanti, kampanya,
  kullanım koşulları) bu kontrolden ayrıntılıdır. Garanti koşulları tablosunu güncel tutun.
- **Teşhis teknisyenindir.** Agent notta olmayan teşhis yazmaz; notu eksik olan iş emri eksik 3C üretir.
- **Örnek veri:** Model, parça numaraları ve araçlar kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk turda şikâyet bölümüne teknisyen bulgusunu ("zincir gergisi kaynaklı")
ekler. Kodun bunu yakaladığı doğrulanır. Ayrıca şunlar doğrulanır:
- parça no, işçilik ve sayı denetimleri,
- garanti süresi / km (29 Şubat dahil),
- darbe işaretiyle "Kontrol" durumu,
- plakanın mesajlara gitmemesi.

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
