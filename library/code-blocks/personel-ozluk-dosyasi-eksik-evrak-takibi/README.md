# Personel Özlük Dosyası Eksik Evrak Takibi · Kod Bloğu

> İnsan Kaynakları › İnsan Kaynakları Uzman Yardımcısı · Workers / Workless

Çalışan listesini zorunlu özlük evrakı listesiyle karşılaştırır. Eksik, süresi dolmuş ve dolmak üzere olan
belgeleri çalışan bazında raporlar ve çalışanlara gönderilecek hatırlatma metinlerini hazırlar. İnternete bağlanmaz.

## Nasıl çalışır?

- **Koşullu evrak listesi:** Her evrak satırına "Koşul" yazılabilir. Bir evrak farklı koşullarla birden fazla
  satırda yer alabilir; çalışana uyan ilk satır uygulanır. Örneğin periyodik muayenenin geçerliliği tehlike
  sınıfına göre değişebilir.

| Koşul örneği | Anlamı |
|---|---|
| `Tümü` (veya boş) | Herkes |
| `Cinsiyet=Erkek` | Alan değeri eşit |
| `Uyruk!=TC` | Alan değeri farklı (boş değilse) |
| `Pozisyon~şoför` | Alan değeri içerir |
| `Tehlike Sınıfı=Çok Tehlikeli`, `Gece Çalışması=Evet` | Çalışanlar dosyasındaki herhangi bir sütun |

- **Teslim süresi:** İşe girişten itibaren evrakın verilmesi gereken gün sayısıdır. 0 "işe başlamadan önce"
  demektir; boş bırakılırsa 30 gün sayılır. Süre dolmadıysa evrak "Bekleniyor", dolduysa "Eksik" olur. Teslim
  süresi 0 olan eksikler Yüksek önemdedir.
- **Geçerlilik (ay):** Teslim tarihinden itibaren sayılır. Kayıtta "Geçerlilik Bitiş" yazılıysa o tarih
  kullanılır. Süresi geçmişse "Süresi dolmuş" (Yüksek), 30 gün içinde dolacaksa "Yaklaşıyor" (`--uyari-gun`).
- **Ayrıca bildirilenler:** Evrak listesinde olmayan evrak adları ve çalışan listesinde olmayan sicil kayıtları
  (ör. ayrılmış personel).

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                     # örnek: 8 çalışan, 16 evrak türü, durum tarihi 08.10.2026
python main.py --calisanlar calisanlar.xlsx --kayitlar evrak_kayitlari.xlsx --liste evrak_listesi.csv
python main.py --calisanlar c.xlsx --kayitlar k.xlsx --liste l.csv --bugun 31.12.2026 --uyari-gun 60
```

| Dosya | Sütunlar |
|---|---|
| Çalışanlar | Sicil No, Ad Soyad, Departman, Pozisyon, İşe Giriş + koşullarda kullanılan sütunlar |
| Evrak kayıtları | Sicil No, Evrak, Teslim Tarihi, Geçerlilik Bitiş (isteğe bağlı) |
| Evrak listesi | Evrak, Koşul, Geçerlilik (ay), Teslim Süresi (gün) |

Evrak adları listeyle birebir aynı olmasa da kayıt eşleşir; adlardan biri diğerini içeriyorsa yeterlidir.

## Çıktı

`ozluk_eksik_evrak.xlsx`:
- `Evrak Matrisi`: çalışan × evrak; ✓ / EKSİK / SÜRE DOLDU / Yaklaşıyor / Bekleniyor / —, renkli.
- `Eksik Listesi`: önem, evrak, durum, açıklama; boş "Talep Edildi" ve "Teslim Alındı" sütunları.
- `Departman Özeti`: tamamlanma oranı ve tam dosyalı çalışan sayısı.
- `Hatırlatma Metinleri`: çalışana gönderilecek kişisel belge talebi taslakları.
- `Uyarılar`.

## Dikkat

- **Evrak listesi örnektir:** Özlük dosyasında bulunması gereken belgeleri kendi işyerinize göre belirleyin. Örnek
  süreler şunlardır:
  - periyodik muayene: çok tehlikeli 1, tehlikeli 3, az tehlikeli 5 yıl
  - temel İSG eğitimi: 1 / 2 / 3 yıl
  - gece çalışması raporu: 2 yıl

  Değerleri güncel mevzuatla (4857 sayılı İş Kanunu, 6331 sayılı İSG Kanunu ve yönetmelikleri), işyeri hekimi ve
  İSG uzmanıyla doğrulayın.
- **Kişisel veri:** Özlük belgeleri, özellikle sağlık raporları ve adli sicil kaydı, özel nitelikli kişisel veri
  içerir. Bu dosyaya erişimi sınırlayın; yalnız gerekli belgeleri isteyin (KVKK ölçülülük ilkesi).
- **Örnek veri:** Örnek çalışanlar ve kayıtlar kurgusaldır.

## Testler

Şunlar test edilir:
- Koşul türleri; teslim süresi (Bekleniyor / Eksik); geçerlilik ve yaklaşan belge.
- Ad eşleştirme; tanımsız evrak ve sicil.
- Departman özeti, Excel sayfaları ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
