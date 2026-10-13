# SGK Giriş-Çıkış Bildirge Kontrolü · Kod Bloğu

> İnsan Kaynakları › Bordro ve Özlük İşleri Uzmanı · Workers / Workless

İşe giriş ve işten çıkış listenizi SGK bildirge kayıtlarınızla karşılaştırır; **süresinde verilmemiş, eksik veya
tarihi uyuşmayan** bildirgeleri listeler. SGK sistemine veya internete bağlanmaz.

## Nasıl kontrol eder?

| Kontrol | Kural |
|---|---|
| İşe giriş bildirgesi son günü | İşe giriş tarihi − `--giris-once` gün (varsayılan 1: en geç bir gün önce) |
| İşten ayrılış bildirgesi son günü | İşten çıkış tarihi + `--cikis-gun` gün (varsayılan 10) |
| Süresinde / Geç | Bildirge tarihi son günden sonraysa "Geç" |
| Yok / Bekliyor | Kayıt yok: son gün geçtiyse "Yok", geçmediyse "Bekliyor" |
| Tarih uyuşmazlığı | Bildirgedeki işe giriş veya çıkış tarihi personel kaydından farklı |
| Karşılıksız bildirge | Bildirge var ama personel listesinde eşleşen kayıt yok |

Eşleştirme kimlik veya sicil numarasıyla, yoksa ad soyadla yapılır. Aynı kişinin birden çok çalışma dönemi varsa
bildirgeler tarihe göre eşlenir.

**Örnek:** İşe giriş 05.10.2026 → bildirge en geç 04.10.2026'da verilmiş olmalı; 05.10'da verildiyse 1 gün geç.
İşten çıkış 30.09.2026 → son gün 10.10.2026; 12.10'da verildiyse 2 gün geç.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek veri, rapor tarihi 13.10.2026
python main.py --personel personel.xlsx --bildirgeler bildirgeler.xlsx
python main.py --personel personel.xlsx --bildirgeler bildirgeler.xlsx --giris-once 0 --tarih 13.10.2026
```

- Personel: Ad Soyad, İşe Giriş Tarihi (zorunlu); TC Kimlik No veya Sicil No, İşten Çıkış Tarihi.
- Bildirgeler: Tür, Bildirge Tarihi (zorunlu); TC Kimlik No, Ad Soyad, Bildirilen Tarih. "Bildirge Tarihi" bildirgenin
  verildiği (onaylandığı) tarihtir.

## Çıktı

`cikti/sgk_bildirge_kontrol.xlsx`:

- `Bildirge Kontrolü`: personel bazında giriş ve çıkış bildirgesinin son günü, veriliş tarihi, durumu ve gecikmesi.
  Kimlik numaraları maskelidir (`123******50`).
- `Uyarılar`: eksik ve geç bildirgeler, tarih uyuşmazlıkları, karşılıksız bildirgeler, okunamayan satırlar.

## Dikkat

- **Süreler genel kuraldır; istisnaları siz yönetirsiniz.** İnşaat, balıkçılık ve tarım işyerlerinde, yabancı ülkeye
  gönderilen sigortalılarda ve ilk defa tescil edilecek işyerlerinde giriş bildirgesi için farklı süreler geçerlidir.
  Böyle işyerleri için `--giris-once 0` kullanın veya sonucu elle değerlendirin.
- **Tatil uzaması hesaplanmaz.** Son gün hafta sonu veya resmi tatile rastlıyorsa süre izleyen iş gününe uzayabilir;
  kod uzatmaz, bu yüzden bazı bildirgeleri gereğinden erken "Geç" veya "Yok" gösterebilir.
- **İdari para cezası hesaplanmaz.** Tutar asgari ücrete ve fiilin niteliğine göre belirlenir; kod yalnız riski gösterir.
- **Çıkış kodu, eksik gün nedeni ve meslek kodu denetlenmez.**
- **Kişisel veri.** Dosyalar kimlik numarası içerir; raporu yalnız yetkili kişilerle paylaşın.
- **Örnek veri:** Kişiler ve kimlik numaraları kurgusaldır.

## Testler

- Giriş ve çıkış bildirgesinde süresinde, geç, yok ve bekliyor durumlarının sınır günleri; parametre etkisi.
- Tarih uyuşmazlığı, aynı kişinin iki çalışma dönemi, karşılıksız bildirge, maskeleme.
- Okunamayan satırlar, komut satırı hataları; örnek veri uçtan uca ve çıktıda açık kimlik no bulunmaması.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
