# Müşteri Hasar Dosyası Takibi · Kod Bloğu

> Teknik ve Operasyon › Hasar Sorumlusu · Workers / Workless

Acente veya broker tarafında izlediğiniz **açık hasar dosyalarını** öncelik sırasına koyar; sigorta şirketi ve eksper
bazında özetler, bekleyen evrakları listeler ve müşteriye gönderilecek evrak hatırlatma metnini hazırlar. İnternete
bağlanmaz.

## Nasıl hesaplar?

```
Dosya yaşı     = rapor tarihi − ihbar tarihi          (ihbar tarihi yoksa hasar tarihi)
Hareketsiz gün = rapor tarihi − son işlem tarihi      (son işlem tarihi yoksa dosya yaşı)
```

| Öncelik (sırayla denenir) | Koşul |
|---|---|
| Süre aşımı | Dosya yaşı > `--azami-gun` (varsayılan 45) |
| Acil | Hareketsiz gün > 2 × `--hareketsiz` (varsayılan 14) |
| Takip et | Hareketsiz gün > `--hareketsiz` (varsayılan 7) |
| Evrak bekleniyor | Bekleyen evrak var |
| Normal | Diğerleri |

- **Açık dosya:** aşaması `kapandı`, `ödendi`, `ret / reddedildi` veya `iptal` ile başlamayan dosya.
- **Önerilen aksiyon:** bekleyen evrak varsa "müşteriden evrak isteyin"; dosya hareketsizse aşamaya göre "eksperi
  arayın", "ödeme tarihini sorun" veya "sigorta şirketini arayın".
- Bekleyen evraklar `|` veya `;` ile ayrılır.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri (rapor tarihi 01.10.2026)
python main.py --dosyalar hasar_dosyalari.xlsx --hareketsiz 5 --azami-gun 30
```

Sütunlar: Dosya No, Aşama (zorunlu); Müşteri, Telefon, Branş, Sigorta Şirketi, Şirket Dosya No, Eksper, Hasar Tarihi,
İhbar Tarihi, Son İşlem Tarihi, Tahmini Tutar, Bekleyen Evraklar.

## Çıktı

`cikti/hasar_dosyasi_takibi.xlsx`:

- `Açık Dosyalar`: öncelik, dosya yaşı, hareketsiz gün, bekleyen evrak ve önerilen aksiyon; boş "Not" sütunu.
- `Şirket Bazında`, `Eksper Bazında`: açık dosya sayısı, ortalama ve en eski dosya yaşı, öncelik dağılımı.
- `Bekleyen Evraklar`: dosya × evrak; boş "İstendi" ve "Geldi" sütunları.
- `Uyarılar`: okunamayan tarihler, ihbar tarihi olmayan dosyalar, mükerrer kayıtlar.

`cikti/evrak_hatirlatma.txt`: müşteri bazında bekleyen evrak listesi (mesaja yapıştırılabilir).

## Dikkat

- **45 gün bir iç takip eşiğidir.** Tazminatın ödenme süresi; branşa, belgelerin eksiksiz teslim edildiği tarihe ve
  ilgili mevzuata (ör. Türk Ticaret Kanunu, zorunlu sigortalarda özel kanunlar) göre değişir. Kod sürenin hangi
  tarihten başladığını ve dolup dolmadığını hukuken değerlendirmez.
- **Liste sizin kaydınızdır.** Kod sigorta şirketinin veya eksperin sistemine bağlanmaz; "son işlem tarihi"
  güncellenmezse dosya hareketsiz görünür.
- **Mesaj göndermez.** Evrak hatırlatma metni taslaktır; gönderim sizin kanalınızdan yapılır.
- **Kişisel veri:** Dosyalar sağlık dahil hassas bilgi içerebilir; KVKK'ya uygun saklayın ve paylaşmayın.
- **Tahmini tutar** dosyaya girdiğiniz değerdir; ödenecek tazminatı göstermez.
- **Örnek veri:** Müşteriler, sigorta şirketleri, eksperler ve tutarlar kurgusaldır.

## Testler

- Öncelik sınırları (45 / 46 gün, 7 / 8 / 14 / 15 gün) ve öncelik sırası.
- Aşamaya göre aksiyon metinleri; ihbar veya son işlem tarihi olmayan dosyalar.
- Şirket ve eksper özetleri, bekleyen evrak listesi ve hatırlatma metni.
- Okuma hataları ve örnek veri uçtan uca.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
