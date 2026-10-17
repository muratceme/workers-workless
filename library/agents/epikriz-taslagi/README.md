# Epikriz Taslağı · AI Agent

> Tıbbi Sekreterlik › Tıbbi Sekreter · Gerisi Bende

Hekim notları ve tetkik sonuçlarından **epikriz taslağı** oluşturur. Öykü, muayene, işlemler, klinik seyir, çıkış
durumu ve öneriler bölümlerini model yazar. **Hasta kimliği, tanılar, tetkik tablosu ve çıkış ilaçları kod tarafından
aynen eklenir.** Notlarda olmayan değer, doz, tarih ya da tanı kodu yazılırsa, referans dışı bir tetkik atlanırsa
taslak modele geri gönderilir. **Hekim onayı zorunludur.**

## Nasıl çalışır?

1. **Hazırlık (kod):** Notlar N1, N2 …, tetkikler T1, T2 … diye numaralanır. Yatış süresi ve her tetkikin referans
   aralığına göre durumu (yüksek / düşük) hesaplanır. Hasta adı, TC kimlik no, protokol no ve doğum tarihi modele
   gönderilmez.
2. **Taslak (model):** Altı anlatı bölümü; her biri dayandığı N ve T numaralarıyla.
3. **Denetim (kod):** Aşağıdakilerden biri varsa taslak modele geri gönderilir (varsayılan en fazla 2 tur):
   - Notlarda ya da tetkiklerde olmayan sayı (değer, doz, tarih); geçersiz dayanak
   - "Tanılar" bölümünde olmayan ICD-10 kodu
   - Yatış-çıkış aralığı dışında tarih (öneriler hariç)
   - Referans dışı bir tetkik öykü, işlemler ya da seyirde anılmamış
   - Notlarda olmayan kesinlik ifadesi ("tamamen iyileşti", "sorunsuz iyileşme", "kesin tanı")
4. **Epikriz (kod):** Kimlik bilgileri, tanılar, tetkik tablosu (referans dışı işaretli), çıkış ilaçları (notlardaki
   gibi) ve hekim onay alanı eklenir.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                     # örnek (kurgusal) apendektomi yatışı
python agent.py --hasta hasta.txt --notlar notlar.docx --tetkik tetkikler.xlsx
```

Not dosyası örneği:

```
Tanılar:
- K35.8 Akut apandisit, diğer ve tanımlanmamış

Notlar:
- 11.10.2026: Laparoskopik apendektomi yapıldı.

Çıkış İlaçları:
- Parasetamol 500 mg tablet, ağrı olduğunda en fazla 3x1
```

## Çıktı

- `cikti/epikriz_taslagi.txt`: epikriz taslağı (hekim onay alanıyla).
- `cikti/epikriz_taslagi.xlsx`: Epikriz (bölüm ve dayandığı notlar), Tetkikler.

## Dikkat

- **Hekim onayı zorunlu.** Epikriz tıbbi bir belgedir; hekim okuyup düzeltmeden ve onaylamadan hastaya verilemez,
  hastane bilgi yönetim sistemine (HBYS) kaydedilemez.
- **Sağlık verisi.** Notlar ve tetkikler yapay zekâ sağlayıcısına gönderilir. Sağlık verisi KVKK'ya göre özel nitelikli
  kişisel veridir. Kurumunuzun yazılı izni ve veri işleme sözleşmesi yoksa çalıştırmayın ya da yerel model (Ollama)
  kullanın. Notlarda hasta adı dışında kimlik bilgisi (adres, yakın adı) bırakmayın.
- **Format:** Kurumunuzun epikriz şablonu farklıysa çıktıyı şablona aktarın.
- **Örnek veri:** Hasta ve klinik süreç kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk taslakta geçersiz bir tetkik numarasına dayanır, çıkıştan sonraki bir
tarihi yazar, "tamamen iyileşti / sorunsuz iyileşme" der, tanılarda olmayan K37 kodunu ve notta olmayan "1,5 g" dozunu
ekler, referans dışı dört tetkiki (lökosit, nötrofil, CRP ×2) anmaz. Kodun on sorunu da yakaladığı; hasta adı, TC kimlik
no, protokol no ve doğum tarihinin modele gönderilmediği doğrulanır.

```bash
python -m unittest discover -s tests
```

## Kaynaklar ve dayanaklar

Kontrol tarihi: 15.10.2026. Mevzuat sık değişir; paketteki süre, oran ve kuralları kullanmadan önce güncel resmî metinle
karşılaştırın. "Karşılaştırılan maddeler" dışındaki atıflar yalnız yol göstericidir.

| Kaynak | Bağlantı | Durum |
|---|---|---|
| 6698 sayılı Kişisel Verilerin Korunması Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.6698.pdf) | Resmî metinle karşılaştırılan maddeler: 13 |

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
