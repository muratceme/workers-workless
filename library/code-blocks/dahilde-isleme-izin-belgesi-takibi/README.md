# Dahilde İşleme İzin Belgesi Takibi · Kod Bloğu

> Dış Ticaret › Dış Ticaret Uzmanı · Workers / Workless

Dahilde işleme izin belgelerinizin (DİİB) **ithalat izni ve ihracat taahhüdü gerçekleşmesini, kalan miktarları,
süre sonunu ve kapatma başvurusu için son günü** tek tabloda izler. Beyanname listesini belge satırlarıyla
eşleştirir. İnternete bağlanmaz.

## Nasıl hesaplar?

| Kavram | Kural |
|---|---|
| Gerçekleşen | Belge süresi içindeki hareketlerin toplamı; belge, tür ve madde adıyla eşleştirilir |
| Kalan taahhüt / kalan ithalat hakkı | Taahhüt − gerçekleşen ihracat; izin − gerçekleşen ithalat |
| Gerçekleşme % | Satırda miktar varsa miktar üzerinden, yoksa tutar üzerinden |
| Taahhüt tamam | Belgenin bütün ihracat satırlarında gerçekleşen ≥ taahhüt |
| Döviz kullanım oranı % | Gerçekleşen ithalat tutarı ÷ gerçekleşen ihracat tutarı × 100 (bilgi amaçlı) |
| Kapatma başvurusu son gün | Süre sonu + `--kapatma-ay` (3) ay |

| Durum | Koşul |
|---|---|
| Kapatma süresi geçmiş | Kapatma başvurusu için son gün geçti |
| Süresi doldu, taahhüt açık | Süre doldu, en az bir ihracat satırı tamamlanmadı |
| Süre yaklaşıyor | Süre sonuna `--gun` (60) gün veya az kaldı, taahhüt tamamlanmadı |
| Taahhüt tamamlandı, kapatma bekliyor | Bütün ihracat satırları tamam; kapatma başvurusu yapılacak |
| Devam ediyor | Diğerleri |

**Örnek:** Taahhüt 45.000 kg, süre içinde ihraç edilen 27.000 kg → gerçekleşme **%60**, kalan **18.000 kg**; süre
sonuna 34 gün var → **Süre yaklaşıyor**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                            # örnek veri, durum tarihi 12.10.2026
python main.py --belgeler belge_satirlari.xlsx --hareketler hareketler.xlsx
python main.py --belgeler belge_satirlari.xlsx --hareketler hareketler.xlsx --gun 90 --kapatma-ay 3
```

Belge dosyasında her satır bir belgenin bir maddesidir; belge tarihi ve süre sonu her satırda tekrar edilebilir (ilk
dolu değer alınır). Zorunlu sütunlar iki dosyada da Belge No ve Tür; miktar ya da tutardan en az biri dolu olmalıdır.

## Çıktı

`cikti/diib_takibi.xlsx`:

- `Belgeler`: durum, süre sonu, kalan gün, kapatma son günü, ihracat ve ithalat gerçekleşmesi, döviz kullanım oranı.
- `Satırlar`: madde bazında belge miktarı, gerçekleşen, kalan ve gerçekleşme yüzdesi.
- `Uyarılar`: süresi dolan ve yaklaşan belgeler, izin aşımı, süre sonrası hareket, eşleşmeyen belge ve madde.

## Dikkat

- **Bu bir takip listesidir, taahhüt hesabı değildir.** Belgenin kapatılması, taahhüt hesabına sayılacak
  beyannameler, kullanım oranları, fire ve ikincil işlem görmüş ürün, eşdeğer eşya ve telafi edici vergi gibi konular
  mevzuata ve belgenizin özel şartlarına göre değerlendirilir. Kapatma öncesi gümrük müşavirinizle çalışın.
- **Süreleri mevzuattan doğrulayın.** Kapatma başvurusu için varsayılan 3 ay bir çalışma varsayımıdır; yürürlükteki
  tebliğe ve belgenizin türüne göre farklıysa `--kapatma-ay` ile değiştirin. Ek süre, haklı sebep ve mücbir sebep
  süreleri otomatik eklenmez; aldıysanız süre sonunu güncelleyin.
- **Süre sonundan sonraki hareketler sayılmaz.** Kod bunları gerçekleşmeye katmaz ve uyarır.
- **Eşleştirme madde adıyladır.** Beyannamedeki madde adı belge satırındakiyle aynı yazılmalıdır (büyük-küçük harf
  ve Türkçe karakter farkı sorun olmaz). GTİP ile izliyorsanız Madde sütununa GTİP yazın.
- **Döviz kullanım oranı bilgi amaçlıdır.** Kod ithalatı CIF, ihracatı FOB değere çevirmez; yazdığınız tutarları
  oranlar. Tutarlar bir belgenin içinde aynı para biriminde olmalıdır.
- **Kapsam dışı:** dahilde işleme izni (gümrük idaresinden alınan izin), teminat ve indirimli teminat takibi,
  yurt içi alımlar, belge devri, taahhüt kapatma evrak listesi.
- **Örnek veri:** Belge numaraları, maddeler ve tutarlar kurgusaldır.

## Testler

- Durum sınırları: süre sonu bugün / dün, 60 ve 61 gün, kapatma son günü bugün / dün; ay ekleme (ay sonu).
- Miktar ve tutar bazlı tamamlanma, izin aşımı, süre sonrası hareket, tanımsız belge ve madde, tekrarlanan satır.
- Belge sıralaması, döviz kullanım oranı, süre sonu olmayan belge, okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
