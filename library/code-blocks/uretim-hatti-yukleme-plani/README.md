# Üretim Hattı Yükleme Planı · Kod Bloğu

> Tekstil ve Konfeksiyon › Planlama › Planlama Uzmanı · Workers / Workless

Siparişlerin dakika değeri (SAM) ve hat kapasitesinden hat bazında yükleme planı çıkarır ve tahmini bitiş
tarihlerini hesaplar. Sevki gecikecek siparişleri ve hat doluluğunu gösterir. İnternete bağlanmaz.

## Nasıl hesaplar?

| Adım | Hesap |
|---|---|
| Hat günlük kapasitesi | operatör sayısı × günlük çalışma (dk) × verimlilik (dk) |
| Günlük adet | kapasite ÷ SAM. Yeni modelin ilk 3 gününde öğrenme eğrisi uygulanır: planlanan verimliliğin %50, %70 ve %85'i (`--ogrenme`). |
| Sıra | Siparişler sevk tarihine göre sıralanır, en erken sevk önce. |
| Hat seçimi | Sipariş, ürün grubuna uygun hatlar içinde en erken bitireceği hatta yüklenir. "Ürün Grupları" boş olan hat her ürünü diker. "Atanan Hat" yazılan sipariş o hatta kalır. |
| Başlangıç | Plan başlangıcı, kesim hazır tarihi ve hattın müsait olduğu tarihten en geç olanıdır. Önceki sipariş gün ortasında biterse kalan dakikalar yeni siparişe geçer. |
| Takvim | Pazar çalışılmaz (`--calisma-gunu 5` ile Cumartesi de). Resmî ve dini bayramlarda çalışılmaz. 28 Ekim ve bayram arifeleri yarım gün sayılır. |
| Sevk kontrolü | En geç bitiş = sevk tarihi − tampon (varsayılan 2 iş günü: final kontrol, yükleme). |

**Sevk durumu:**
- **GECİKECEK:** Tahmini bitiş, en geç bitişten sonra. Zamanında bitmesi için gereken günlük adet de gösterilir.
- **Riskli:** Bolluk 2 iş günü veya daha az.
- **Zamanında:** Bolluk 2 iş gününden fazla.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                     # örnek: 4 hat, 7 sipariş, başlangıç 12.10.2026
python main.py --hatlar hatlar.xlsx --siparisler siparisler.xlsx --baslangic 12.10.2026
python main.py --hatlar h.xlsx --siparisler s.xlsx --ogrenme 60,80 --tampon 3 --calisma-gunu 5 --ek-tatil fabrika_tatili.csv
```

| Dosya | Sütunlar |
|---|---|
| Hatlar | Hat, Operatör Sayısı, Günlük Çalışma (dk), Verimlilik %, Ürün Grupları (virgülle; boş = hepsi), Müsait Olduğu Tarih |
| Siparişler | Sipariş No, Model, Müşteri, Ürün Grubu, Adet, SAM (dk), Kesim Hazır Tarihi, Sevk Tarihi, Atanan Hat (isteğe bağlı) |
| Ek tatil (isteğe bağlı) | Tarih, Açıklama |

## Çıktı

`hat_yukleme_plani.xlsx`:
- `Yükleme Planı`: her siparişin şu bilgileri ve boş "Planlayıcı Notu" sütunu:
  - hat
  - başlangıç, tahmini bitiş ve en geç bitiş tarihi
  - bolluk (iş günü)
  - ortalama günlük adet
  - durum
- `Hat Yükü`: hat kapasitesi, yüklenen dakika, plan sonuna kadar doluluk oranı ve boşalma tarihi.
- `Günlük Plan`: hat × gün tablosu, her gün dikilecek adetlerle. Tatiller "—", yarım günler "½" ile gösterilir.
- `Uyarılar`: gecikme, az bolluk, uygun hat yok, uzmanlık dışı atama, hat bekleyen kesim ve boş hat; "Karar" sütunuyla.

## Dikkat

- **Plan bir öneridir:** Hat seçimi basit bir kuralla yapılır: en erken sevk önce, en erken bitiren hat. Siparişi
  bölme, hat birleştirme ve fazla mesai kararları planlayıcıya aittir. Gecikme uyarıları bu kararlar için gereken
  günlük adedi gösterir.
- **Örnek değerler:** Verimlilik, öğrenme eğrisi ve tampon değerleri örnektir. Kendi hat ve model geçmişinize göre
  girin. SAM değerleri iş etüdü veya GSD / benzeri sistemlerden alınmalıdır.
- **Tatil takvimi:** Dini bayram tarihleri 2025–2027 için tanımlıdır. Fabrika tatili ve köprü günlerini `--ek-tatil`
  ile ekleyin.
- **Örnek veri:** Örnek hatlar ve siparişler kurgusaldır.

## Testler

Şunlar test edilir:
- Kapasite ve öğrenme eğrisi hesabı; arife ve bayram günleri.
- Hat seçimi, uzmanlık ve sabit atama.
- Gecikme / bolluk hesabı ve gereken günlük adet.
- Hat doluluğu, Excel sayfaları ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
