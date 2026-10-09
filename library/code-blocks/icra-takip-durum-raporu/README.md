# İcra Takip Durum Raporu · Kod Bloğu

> Hukuk › Avukat · Workers / Workless

İcra dosyalarının aşama, tahsilat, masraf ve sürelerini tek raporda toplar. Her dosya için rapor tarihine kadar faiz
ve mahsup adımlarını gösteren bir **hesap dökümü** çıkarır. İtiraz, itirazın kaldırılması, itirazın iptali ve haciz
isteme sürelerini takip eder. İnternete bağlanmaz.

## Nasıl hesaplar?

### Dosya hesabı

- **Takip sonrası faiz:** kalan asıl alacak × yıllık faiz oranı × gün / 365 (basit faiz). Takip tarihinden başlar; her
  hareket tarihinde ve rapor tarihinde işletilir.
- **Masraflar** dosya borcuna eklenir.
- **Tahsilat mahsubu:** Her tahsilat önce masrafa, sonra faize, sonra asıl alacağa mahsup edilir (kısmi ödemenin önce
  faiz ve giderlere sayılması; TBK 100).
- **Fazla tahsilat:** Tahsilat bu hesaptaki borcu aşarsa aşan kısım uyarıyla gösterilir.

**Hesaplanmayanlar:** vekâlet ücreti, tahsil harcı, cezaevi harcı. Faiz oranı dönem içinde değişiyorsa (ör. avans faizi)
tek oranla yapılan hesap yaklaşıktır. Kesin dosya hesabı icra dairesinden alınmalıdır.

### Süreler

| Süre | Başlangıç | Süre | Dayanak |
|---|---|---|---|
| Ödeme emrine itiraz (ilamsız) | Ödeme emri tebliği | 7 gün | İİK 62 |
| Ödeme emrine itiraz (rehin) | Ödeme emri tebliği | 7 gün | İİK 149 |
| İtiraz / şikâyet (kambiyo) | Ödeme emri tebliği | 5 gün | İİK 168 |
| Ödeme süresi (kambiyo; bitmeden haciz istenemez) | Ödeme emri tebliği | 10 gün | İİK 168 |
| İcra emri: ödeme / itiraz (ilamlı) | İcra emri tebliği | 7 gün | İİK 32–33 |
| İtirazın kaldırılması (icra mahkemesi) | İtirazın alacaklıya tebliği | 6 ay | İİK 68 |
| İtirazın iptali davası | İtirazın alacaklıya tebliği | 1 yıl | İİK 67 |
| Haciz isteme (itiraz yoksa) | Ödeme / icra emri tebliği | 1 yıl | İİK 78 |

- Son gün hafta sonu veya ulusal bayrama rastlarsa izleyen iş gününe kayar. Dinî bayramlar dikkate alınmaz; son
  günü kontrol edin.
- Borçlu lehine süreler (itiraz, ödeme) yalnız bilgi olarak gösterilir.
- Haciz yapılmış dosyada haciz isteme süresi "Haciz yapıldı" olarak işaretlenir.
- İtiraz edilmiş dosyada haciz isteme süresi hesaplanmaz, çünkü itiraz ve dava süreleri bu süreye sayılmaz (İİK 78).

| Kontrol | Önem |
|---|---|
| Takip tarihi veya asıl alacak eksik (dosya alınmaz) | Yüksek |
| Alacaklı lehine süre geçmiş (İİK 67 / 68 / 78) | Yüksek |
| Alacaklı lehine süre sonuna 7 gün veya daha az kaldı | Yüksek |
| Süre sonuna 8–30 gün (`--uyari-gun`) kaldı | Orta |
| Fazla tahsilat; borç kapanmış ama dosya açık | Orta |
| İtiraz var, itirazın tebliğ tarihi yok; faiz oranı yok | Orta |
| Son işlemden bu yana 180 günden fazla geçmiş açık dosya (`--islemsiz-gun`) | Orta |
| Eşleşmeyen veya türü okunamayan hareket | Orta |
| Ödeme / icra emri tebliğ tarihi yok (30 günden eskiyse Orta) | Bilgi |

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                              # örnek: 12 dosya, rapor tarihi 09.10.2026
python main.py --dosyalar dosyalar.xlsx --hareketler hareketler.xlsx --bugun 09.10.2026 --islemsiz-gun 180
```

| Dosya | Sütunlar |
|---|---|
| Dosyalar | İcra Dairesi, Dosya No, Borçlu, Takip Türü, Takip Tarihi, Asıl Alacak, İşlemiş Faiz, Faiz Oranı (%24 veya 24), Ödeme Emri Tebliğ Tarihi, İtiraz (Var / Yok), İtiraz Tebliğ Tarihi, Haciz Tarihi, Aşama, Sorumlu, Son İşlem Tarihi, Durum (Açık / Kapandı...), Not |
| Hareketler | Dosya No, Tarih, Tür (Tahsilat / Masraf), Tutar, Açıklama |

Aynı esas numarası birden çok icra dairesinde varsa hareketlerde Dosya No'ya daire adını da yazın, örneğin
"İstanbul 12. İcra Dairesi 2026/1201 E.".

## Çıktı

`icra_takip_raporu.xlsx`:
- `Özet`: toplam alacak, tahsilat, tahsilat oranı, kalan; aşama ve takip türü dağılımı.
- `Dosyalar`: dosya bazında faiz, masraf, tahsilat ve kalan.
- `Hesap Dökümü`: her hareketteki dönem faizi ve mahsup (masrafa / faize / asla) adımları.
- `Süre Takibi`: boş "Yapılan İşlem" sütunuyla.
- `İşlemsiz Dosyalar`: boş "Planlanan İşlem" sütunuyla.
- `Aylık Tahsilat`, `Uyarılar`.

## Dikkat

- **Süreler ve faiz hesabı yardımcıdır.** Tebliğ tarihlerini UYAP kayıtlarıyla kontrol edin. Mevzuat değişikliklerini
  takip edin.
- **Faiz oranı:** Takip talebindeki faiz türünü (yasal, avans, sözleşme) ve oranını kullanın. Değişken oranlarda hesap
  yaklaşıktır.
- **Kişisel veri:** Borçlu bilgileri kişisel veridir. Raporu yalnız yetkili kişilerle paylaşın.
- **Örnek veri:** Örnek dosyalar, borçlular ve tutarlar kurgusaldır.

## Testler

Şunlar test edilir:
- Bir dosyanın faiz ve mahsup adımları elle hesaplanmış değerlerle.
- Fazla tahsilat ve kapanmış borç.
- İİK 62 / 67 / 68 / 78 / 149 / 168 süreleri ve hafta sonu kayması.
- Uyarılar (borçlu lehine sürelerin ve kapalı dosyanın uyarı üretmemesi dahil), Excel ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Hukuki görüş yerine
geçmez. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
