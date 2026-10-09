# KVKK Kişisel Veri İşleme Envanteri · Kod Bloğu

> Bilgi Teknolojileri › Bilgi Güvenliği Uzmanı · Workers / Workless

Departmanlardan toplanan veri işleme formlarını tek bir kişisel veri işleme envanterinde birleştirir. Eksik,
tutarsız veya riskli kayıtları KVKK maddelerine göre işaretler. İnternete bağlanmaz.

## Neleri kontrol eder?

- **Hukuki sebep:** Serbest yazılmış sebepler KVKK maddelerine eşlenir.
  - md. 5/1: açık rıza.
  - md. 5/2: (a) kanunlarda açıkça öngörülme, (b) fiili imkânsızlık, (c) sözleşmenin kurulması / ifası, (ç) hukuki
    yükümlülük, (d) alenileştirme, (e) hakkın tesisi, kullanılması, korunması, (f) meşru menfaat.

  Tanınmayan sebep işaretlenir.
- **Özel nitelikli veri (md. 6):** Sağlık, biyometrik, genetik, ceza mahkûmiyeti, din, sendika üyeliği, ırk / etnik
  köken, siyasi düşünce, felsefi inanç, kılık kıyafet ve cinsel hayat verileri anahtar kelimeyle tespit edilir.
  - Formdaki "Özel Nitelikli" işaretiyle çelişki işaretlenir.
  - Yalnız md. 6'da yer almayan bir sebebe (meşru menfaat, sözleşme) dayanma işaretlenir.
  - Çalışan sağlık verilerinde md. 6/3-(e) önerilir.
  - Biyometrik veride ölçülülük hatırlatılır.
- **Yurt dışı aktarım (md. 9):**
  - Ülke ve dayanak zorunludur: yeterlilik kararı, standart sözleşme, bağlayıcı şirket kuralları, taahhütname veya
    arızi aktarım.
  - 2024 değişikliğiyle açık rıza yalnız arızi aktarımda dayanak olabilir.
  - Standart sözleşmenin 5 iş günü içinde Kurula bildirimi hatırlatılır.
- **Diğer:**
  - Belirsiz saklama süresi ("süresiz", "gerektiği kadar") ve dayanaksız süre.
  - Boş teknik / idari tedbir.
  - Pazarlama sürecinde TCKN toplanması (ölçülülük).
  - Ödeme, fatura, bordro gibi sözleşme / yükümlülük süreçlerinde gereksiz açık rıza.
  - Mükerrer kayıt; formu gelmeyen departman.

| Bulgu | Önem |
|---|---|
| Özel nitelikli veride md. 6 dışı sebep; yurt dışı aktarımda dayanak / ülke yok; amaç veya sebep boş | Yüksek |
| Özel nitelik işareti çelişkisi, biyometrik veri, yurt dışında açık rıza, belirsiz süre, tedbir yok, ölçülülük, eksik alan, form gelmedi | Orta |
| Gereksiz açık rıza, süre dayanağı yok, standart sözleşme bildirimi, mükerrer kayıt | Bilgi |

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                     # örnek: 15 kayıt, 6 departman (biri form göndermemiş)
python main.py --formlar formlar.xlsx --departmanlar departmanlar.csv
```

Form sütunları için `ornek_veri/departman_formlari.csv` dosyasına bakın. Her satır bir süreç × veri kategorisidir.
Departman listesi `Departman;Sorumlu` biçimindedir.

## Çıktı

`kvkk_envanter.xlsx`:
- `Bulgular`: boş "Düzeltildi" sütunuyla.
- `Envanter`: normalleştirilmiş hukuki sebep ve KVKK maddesi, özel nitelik, yurt dışı aktarım özeti; boş "Durum"
  sütunu.
- `Özel Nitelikli Veriler`: uygunluk işaretiyle.
- `Yurt Dışı Aktarımlar`.
- `Saklama ve İmha`: kategori bazında süre ve dayanak; boş "İmha Yöntemi" sütunu. Saklama ve imha politikasına
  temel olur.
- `Kategori Özeti`: kişi grupları, amaçlar, sebepler ve departmanlar.
- `Departman Durumu`: Tamam / Gözden geçir / Düzeltme gerekli / Form gelmedi.

## Dikkat

- **Hukuki değerlendirme değildir:** Paket anahtar kelime ve kurallarla ön kontrol yapar. Hukuki sebep, saklama
  süresi ve aktarım dayanağı kararları KVKK uyum sorumlusu veya hukuk danışmanınızca verilmelidir.
- **Güncel mevzuat:** 7499 sayılı Kanunla (2024) md. 6 ve md. 9 değişmiştir. Kurul kararları ve rehberlerle güncel
  durumu doğrulayın. VERBİS kayıt yükümlülüğü ve istisnaları ayrıca değerlendirilmelidir.
- **Örnek veri:** Örnek formlar kurgusaldır.

## Testler

Şunlar test edilir:
- Özel nitelikli veri tespiti (yanlış pozitiflere karşı dahil) ve md. 6 sebebi kontrolü.
- Yurt dışı aktarım dayanakları; ölçülülük, belirsiz süre, gereksiz açık rıza, mükerrer kayıt, form gelmedi.
- Hukuki sebep çözümleme; Excel ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
