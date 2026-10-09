# Kullanıcı Yetki Gözden Geçirme · Kod Bloğu

> Bilgi Teknolojileri › Bilgi Güvenliği Uzmanı · Workers / Workless

Sistemlerdeki kullanıcı-yetki listelerini İK personel listesiyle karşılaştırır; buna periyodik erişim gözden
geçirmesi denir. Bulduklarını listeler: ayrılan personelin açık kalan hesapları, sahipsiz hesaplar, kullanılmayan
hesaplar, rol matrisi dışındaki yetkiler ve görevler ayrılığı (SoD) ihlalleri. Yöneticilere gönderilecek onay
listesini hazırlar. İnternete bağlanmaz.

## Nasıl çalışır?

- **Eşleştirme:** sicil no → e-posta → ad soyad sırasıyla yapılır; Türkçe karakter ve büyük/küçük harf duyarsızdır.
  Aynı adda birden çok personel varsa eşleşme "belirsiz" olarak işaretlenir.
- **Hesap türleri:**
  - Kişisel hesap İK'da bulunamazsa "sahipsiz" sayılır.
  - Ortak hesap işlem sorumluluğunu izlenemez kılar.
  - Servis hesabının sahibi ve amacı belgelenmelidir.
- **Rol matrisi:** Pozisyon × sistem → izinli roller. `*` pozisyonu herkese uygulanır.
- **Görevler ayrılığı:** Kurallar kişi bazında ve sistemler arasında uygulanır. Örneğin ERP'de ödeme hazırlayan biri
  banka portalında ödeme onaylayabiliyorsa bu bir ihlaldir. Rol adlarında `*` joker kullanılabilir
  (`Ödeme Onayı*`).

| Bulgu | Önem |
|---|---|
| Ayrılan personelin aktif hesabı | Yüksek |
| Ayrılış tarihinden sonra giriş yapılmış hesap (olay incelemesi) | Yüksek |
| İK'da karşılığı olmayan aktif kişisel hesap | Yüksek |
| Görevler ayrılığı ihlali | Yüksek |
| Ortak hesap; `--pasif-gun` (90) gündür kullanılmayan aktif hesap | Orta |
| Rol matrisi dışı yetki; belirsiz eşleşme | Orta |
| Servis hesabı; ayrılan personelin pasif hesabı; izindeki personel; hesabı olmayan personel | Bilgi |

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                     # örnek: 11 personel, 3 sistem, 21 hesap
python main.py --personel ik.xlsx --yetkiler yetkiler.xlsx --rol-matrisi matris.csv --gorev-ayriligi sod.csv --kritik-roller kritik.csv
python main.py --personel ik.xlsx --yetkiler y.xlsx --bugun 31.12.2026 --pasif-gun 60
```

| Dosya | Sütunlar |
|---|---|
| Personel | Sicil No, Ad Soyad, E-posta, Departman, Pozisyon, Durum (Aktif / Ayrıldı / İzinde), Ayrılış Tarihi, Yönetici |
| Yetkiler | Sistem, Kullanıcı Adı, Ad Soyad, E-posta, Sicil No, Rol, Son Giriş, Hesap Durumu (Aktif / Kilitli / Pasif), Hesap Türü (Kişisel / Ortak / Servis) |
| Rol matrisi | Pozisyon, Sistem, İzinli Roller (`\|` ile) |
| Görevler ayrılığı | Kural, Rol A, Rol B, Risk |
| Kritik roller | Rol |

Her rol ayrı bir satırdır. Birden çok sistemin dışa aktarımlarını tek dosyada "Sistem" sütunuyla birleştirin.

## Çıktı

`yetki_gozden_gecirme.xlsx`:
- `Bulgular`: boş "Aksiyon" ve "Tamamlandı" sütunlarıyla.
- `Yönetici Onay Listesi`: aktif hesapların her rolü, onaylayacak yöneticiye göre gruplanmış. Kritik rol ve bulgu
  işaretlidir; boş "Karar (Onay / İptal)" sütunu vardır. Sahipsiz hesaplar Bilgi İşlem'e düşer.
- `Görevler Ayrılığı`: kişi, kural, çakışan roller, risk; boş "Telafi Edici Kontrol" sütunu.
- `Hesap Eşleşmesi`: her hesabın eşleştiği personel ve eşleşme yöntemi.
- `Sistem Özeti`.

## Dikkat

- **Kurallar örnektir:** Rol matrisi, görevler ayrılığı kuralları ve kritik roller kurumunuzun yetki politikasına göre
  hazırlanmalıdır.
- **Ayrılıştan sonra giriş** yetkisiz erişim olabilir. Bu bulgu bilgi güvenliği olay yönetimi sürecine aktarılmalıdır.
- **Kişisel veri:** Personel ve erişim kayıtları kişisel veridir (KVKK). Raporu yetkili kişilerle sınırlı
  paylaşın ve saklama süresine uyun.
- **Örnek veri:** Örnek kişiler ve hesaplar kurgusaldır.

## Testler

Şunlar test edilir:
- Sicil / e-posta / ad eşleştirmesi; rollerin hesap bazında toplanması.
- Ayrılan personel, ayrılıştan sonra giriş, sahipsiz, ortak, servis ve kullanılmayan hesap bulguları.
- Rol matrisi (`*` dahil), sistemler arası görevler ayrılığı, kritik roller.
- Yönetici onay listesi, joker eşleşme, Excel ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
