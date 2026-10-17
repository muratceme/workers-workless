# Randevu Hatırlatma Mesajları · AI Agent

> Hasta Hizmetleri › Hasta Kayıt Kabul Görevlisi · Gerisi Bende

Ertesi günün randevularından bölüm, hekim ve **kurumun onaylı hazırlık talimatlarını** içeren hatırlatma mesajları hazırlar.
**Hasta verisi yapay zekâ modeline gönderilmez.** Model yalnız bölüm, işlem ve dil kombinasyonları için yer tutuculu şablon yazar;
hasta adı, tarih, saat ve hekim bilgisini kod doldurur. Mesaj göndermez, SMS sistemine aktarılacak CSV üretir.

## Nasıl çalışır?

1. **Hazırlık (kod):**
   - Seçilen günün randevuları okunur (varsayılan yarın).
   - Bölüm + işlem + dil kombinasyonları çıkarılır.
   - Her kombinasyona onaylı talimatlar eşlenir: bölüm + işlem satırı, bölümün genel satırı (işlem boş) ve "Tümü" satırı.
2. **Şablon (model):** Her kombinasyon için kısa bir şablon yazılır. Yer tutucular: `{AD}`, `{TARIH}`, `{SAAT}`, `{BOLUM}`,
   `{HEKIM_IFADE}`, `{KURUM}`, `{ILETISIM}`. Model hasta adı, telefon ya da saat görmez.
3. **Denetim (kod):** Aşağıdaki durumlardan biri varsa şablon modele geri gönderilir (varsayılan en fazla 2 tur):
   - Zorunlu yer tutucu eksik ya da bilinmeyen yer tutucu var
   - Talimattaki bir sayı şablonda yok ("8 saat açlık" → "6 saat" olmuş)
   - Talimatta olmayan bir sayı eklenmiş
   - Talimatta olmayan ilaç, doz, insülin, kan sulandırıcı gibi bir tıbbi yönlendirme eklenmiş
   - En uzun değerlerle doldurulan mesaj `--azami` (480) karakteri aşıyor

   Kontrolden geçmeyen şablonla mesaj üretilmez.
4. **Doldurma (kod):** Telefonu olan hastalar için mesajlar oluşturulur; telefonu olmayanlar "aranmalı" diye listelenir.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                            # örnek randevular (15.10.2026)
python agent.py --randevular randevular.xlsx --talimatlar hazirlik.xlsx --kurum "Hastane Adı" --iletisim "0212 000 00 00"
python agent.py --randevular randevular.xlsx --talimatlar hazirlik.xlsx --tarih 20.10.2026 --kurum "…" --iletisim "…" --azami 300
```

Dil sütununda TR, EN, DE, AR ya da RU yazılabilir (boş = TR).

## Çıktı

- `cikti/randevu_hatirlatma.xlsx`:
  - `Şablonlar`: onaylı talimatlar, şablon metni ve "Onay" sütunu.
  - `Mesajlar`: doldurulmuş mesajlar ve karakter sayısı.
  - `Gönderilemeyen`: telefonu olmayanlar ve kontrolden geçmeyen şablonlar.
- `cikti/randevu_hatirlatma_sms.csv`: Telefon; Mesaj.

## Dikkat

- **Şablonları onaydan geçirin.** Hazırlık talimatları ve şablonlar sorumlu hekim / kalite birimi tarafından onaylanmalıdır.
  Örnek talimatlar yalnız örnektir; kurumunuzun onaylı talimatlarını kullanın.
- **Sağlık verisi.** Randevu bilgisi KVKK kapsamında özel nitelikli kişisel veridir. Bu araç hasta verisini modele göndermez;
  CSV ve Excel çıktısını da yalnız yetkili personel ve kurumun SMS sistemi kullanmalıdır.
- **Gönderim.** Hatırlatma mesajlarının gönderim izni, ileti türü ve İYS yükümlülüğü için kurumunuzun hukuk / uyum birimine danışın.
- **Denetim kural tabanlıdır.** Sayı ve anahtar kelime kontrolü anlam kaymasının hepsini yakalayamaz; şablonları okuyarak onaylayın.
- **Örnek veri:** Hastalar, telefonlar, hekimler ve talimatlar kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk taslakta bilerek gastroskopi açlık süresini 6 saat yazar, MR şablonuna talimatta
olmayan bir ilaç önerisi ekler ve bir şablonda `{ILETISIM}` yer tutucusunu unutur. Kodun bunları yakalayıp düzelttirdiği ve
modele hasta adı ya da telefonu gönderilmediği doğrulanır.

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
