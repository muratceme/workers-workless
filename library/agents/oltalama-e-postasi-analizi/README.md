# Oltalama E-postası Analizi · AI Agent

> Bilgi Teknolojileri › Bilgi Güvenliği Uzmanı · Workers / Workless

Kullanıcıların bildirdiği şüpheli e-postaları (.eml) inceler ve **risk değerlendirmesi** yazar. Teknik göstergeleri kod
çıkarır, model bunları yorumlar. **Bağlantılara gidilmez, ekler açılmaz.**

## Nasıl çalışır?

1. **Göstergeler (kod):** Her e-posta için:
   - **Kimlik doğrulama:** SPF / DKIM / DMARC sonucu (`Authentication-Results`).
   - **Başlık uyumsuzluğu:** Reply-To ya da Return-Path alan adı gönderenden farklı.
   - **Benzer alan adı:** güvenilir alan adınızı içeren ya da ona çok benzeyen adresler (`ornekbank-guvenlik.com`,
     `0rnekbank.com`) ve punycode (`xn--…`) adresler.
   - **Görünen ad:** gönderen adı kurumunuzu çağrıştırıyor ama adres başka.
   - **Bağlantı:** görünen metin ile gerçek adres farklı, doğrudan IP adresi ya da kısaltılmış bağlantı (bit.ly vb.).
   - **Ek:** çalıştırılabilir, betik, ISO / LNK / HTML ya da çift uzantılı ekler (`fatura.pdf.exe`), makrolu Office
     dosyaları ve arşivler.
   - **İçerik:** aciliyet, şifre / kart bilgisi ya da ödeme talebi; ücretsiz e-posta servisinden ödeme talebi.

   En yüksek gösterge seviyesi, e-postanın **en düşük risk seviyesini** belirler.
2. **Değerlendirme (model):** Risk seviyesi, sınıf (oltalama, zararlı ek, CEO / ödeme dolandırıcılığı, spam, meşru),
   gösterge numaralı gerekçeler, kullanıcıya tavsiye ve BT aksiyonları yazılır.
3. **Denetim (kod):** Aşağıdakilerden biri varsa yalnız o e-postanın analizi modele geri gönderilir (varsayılan en
   fazla 2 tur):
   - Risk, göstergelerin gerektirdiği seviyenin altında
   - Yüksek seviyeli bir gösterge gerekçede kullanılmamış ya da "Meşru" denmiş
   - Gösterge numarası ya da alıntı geçersiz
   - Şüpheli bağlantı veya alan adı etkisizleştirilmeden yazılmış (`hxxp`, `[.]` kullanılmalı)
   - Tavsiyede "bağlantıya tıklayın", "eki açın" gibi yönlendirme var

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                           # örnek 3 e-posta
python agent.py --girdi supheli.eml --guvenilir "firmam.com.tr,bankam.com.tr"
python agent.py --girdi bildirilenler/ --guvenilir "firmam.com.tr"
```

E-postayı Outlook'ta "Farklı Kaydet" ya da masaüstüne sürükleyerek, Gmail'de "İletiyi indir" ile .eml olarak
kaydedebilirsiniz. Ek olarak iletilmiş e-postada orijinal başlıklar kaybolabilir; mümkünse orijinal .eml'i kullanın.

`--guvenilir`: kurumunuzun alan adları ve sık taklit edilen iş ortaklarınızın (banka, kargo, tedarikçi) alan adları.

## Çıktı

`cikti/oltalama_analizi.xlsx`:
- `Özet`: e-posta başına kod riski, AI riski, sınıf, özet, kullanıcıya tavsiye, BT aksiyonları.
- `Göstergeler`: her gösterge, seviyesi ve AI gerekçesi.
- `Bağlantılar`: görünen metin ve etkisizleştirilmiş gerçek adres.

## Dikkat

- **Ön değerlendirmedir.** Kesin karar için kurumunuzun e-posta güvenlik ağ geçidi, sandbox ve olay müdahale
  prosedürünü kullanın. Bağlantıların gerçek içeriği ve eklerin davranışı incelenmez.
- **Gönderilen veri.** Konu, gövde (e-posta adresi, telefon ve IBAN maskeli), bağlantılar ve göstergeler modele
  gönderilir. Kişisel ya da gizli bilgi içeren iletilerde kurum politikanıza uyun.
- **Örnek veri:** E-postalar kurgusaldır. Alan adları ve IP adresi (198.51.100.0/24, belgeleme için ayrılmış blok)
  örnek amaçlıdır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk analizde bilerek:
- oltalama e-postasına "Orta" risk verir ve yüksek seviyeli göstergeleri atlar,
- sahte alan adını tıklanabilir yazar ve "bağlantıya tıklayın" der,
- ödeme dolandırıcılığında e-postada olmayan bir cümleyi alıntılar.

Kodun bunları yakaladığı ve yalnız sorunlu e-postaları yeniden analiz ettirdiği doğrulanır.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. API kullanım ücretleri ve güvenlik kararları kullanıcının sorumluluğundadır. Bkz. `LICENSE`
ve `SORUMLULUK_REDDI.md`.
