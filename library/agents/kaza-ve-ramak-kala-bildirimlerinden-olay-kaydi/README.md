# Kaza ve Ramak Kala Bildirimlerinden Olay Kaydı · AI Agent

> İş Sağlığı ve Güvenliği › İşyeri Hekimi · Workers / Workless

Çalışanların ve amirlerin serbest metinle yazdığı kaza / ramak kala bildirimlerini olay kaydına dönüştürür:
- tür (İş Kazası / Ramak Kala / Belirsiz), kanıt cümlesiyle,
- neden kategorisi, kayıp gün, vardiya, kısa açıklama,
- iş kazalarında yaralanma bilgisi (vücut bölgesi, yaralanma türü, müdahale, sevk),
- iş kazalarında **SGK bildirim son günü** (kod hesabı).

Çıktı doğrudan **İş Kazası ve Ramak Kala Analizi** kod bloğuna verilebilir.

## Nasıl çalışır?

1. **Hazırlık (kod):** Bildirimler okunur. Bildiren ve `--gizle` ile verilen kazazede / tanık adları takma adla, telefon,
   e-posta ve TCKN maskelenir.
2. **Sınıflama (model):** Yalnız bildirimde yazana dayanır; tanı koymaz, yazmayan kusuru varsaymaz.
3. **Denetim (kod):** Aşağıdakilerden biri varsa yalnız o bildirim modele geri gönderilir (varsayılan en fazla 2 tur):
   - Kanıtın bildirimde aynen geçmemesi
   - Zarar ifadesi (yaralanma, sıkışma, incinme, revir, ilk yardım, sevk, rapor …) varken "Ramak Kala"; zarar
     yazmazken ya da "kimse yaralanmadı" yazarken "İş Kazası" denmesi
   - Bildirimde yazmayan kayıp gün, sayı ya da vardiya; ramak kalada kayıp gün ya da yaralanma bilgisi
   - Listede olmayan neden; açıklamada kişi adı; 120 karakteri aşan açıklama
4. **SGK son günü (kod):** Kaza tarihinden sonraki 3 iş günü. Kaza günü sayılmaz; cumartesi, pazar ve `--tatiller` ile
   verilen resmî tatiller iş günü değildir. `--tarih`e göre kalan gün ya da "Süresi geçti" yazılır.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                              # örnek 4 bildirim (2 iş kazası, 2 ramak kala)
python agent.py --bildirimler bildirimler.docx --tatiller 28.10.2026 29.10.2026 --gizle "Ad Soyad"
```

Bildirim örneği:

```
Bildirim: B-102 | Tarih: 15.10.2026 | Saat: 23:40 | Bölüm: Üretim | Bildiren: Serkan Polat
Gece vardiyasında pres hattında kalıp değişimi sırasında Ali Demir'in sol elinin işaret parmağı kalıp ile tabla
arasında sıkıştı. Revirde ilk yardım yapıldı, hastaneye sevk edildi, 2 dikiş atıldı ve 5 gün rapor verildi.
```

## Çıktı

- `cikti/kaza_bildirimleri.xlsx`:
  - `Bildirimler`: sınıflama, yaralanma, SGK bildirim son günü ve durumu, kanıt, eksik bilgiler (mor: model).
  - `Kontrol`
- `cikti/kaza_bildirimleri_olaylar.xlsx`: kod bloğunun olay dosyası (Tarih, Saat, Tür, Bölüm, Neden, Kayıp Gün,
  Vardiya, Açıklama). "Belirsiz" bildirimler bu dosyaya yazılmaz.

```bash
python ../../code-blocks/is-kazasi-ve-ramak-kala-analizi/main.py --olaylar cikti/kaza_bildirimleri_olaylar.xlsx
```

## Dikkat

- **SGK bildirim süresi:** 5510 sayılı Kanun md. 13 ve 6331 sayılı Kanun md. 14'e göre iş kazası en geç 3 iş günü içinde
  SGK'ya bildirilir. Kod kaza tarihini başlangıç alır (SGK genelgesindeki ihtiyatlı yaklaşım). Kaza işveren kontrolü
  dışında olduysa süre öğrenme tarihinden başlayabilir. Resmî tatilleri `--tatiller` ile vermezseniz hesaba girmez.
  Güncel mevzuatı kontrol edin.
- **İş kazası mı, değil mi** kararı (ör. işyeri dışı, yol kazası) hukuki değerlendirme gerektirir; kod yalnız metindeki
  zarar ifadesine bakar.
- **Sağlık bilgisi özel nitelikli kişisel veridir** (KVKK md. 6). Bildirimlerde tanı ve tedavi bilgisi varsa yerel model
  (Ollama) kullanın; çıktıyı yetkisiz kişilerle paylaşmayın. Yalnız tam ad soyad maskelenir.
- **Örnek veri:** Kişiler ve olaylar kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk turda B-103'teki bel tutulmasını "Ramak Kala" sayar. Kodun bunu yakalayıp
düzelttirdiği doğrulanır. Ayrıca şunlar doğrulanır:
- tür, kanıt, kayıp gün, vardiya, sayı, kişi adı ve neden denetimleri,
- iş günü hesabı (hafta sonu, resmî tatil),
- kod bloğu dosyasının sütun düzeni,
- adların mesajlara gitmemesi.

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
