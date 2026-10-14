# Randevu Hatırlatma Mesajları · AI Agent

> Hasta Hizmetleri › Hasta Kayıt Kabul Görevlisi · Workers / Workless

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

---
Bu paket "olduğu gibi" sunulur. API kullanım ücretleri ve gönderilen mesajların doğruluğu kullanıcının sorumluluğundadır. Bkz.
`LICENSE` ve `SORUMLULUK_REDDI.md`.
