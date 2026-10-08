# Erken Uyarı Sinyalleri Listesi · Kod Bloğu

> Bankacılık › Krediler İzleme ve Takip › Kredi İzleme Uzmanı · Workers / Workless

Kredi müşterilerine ait dağınık sinyalleri tek listede birleştirir ve firma bazında **risk puanı** üretir. İzleme
biriminin hangi firmaya önce bakacağını gösterir. İnternete bağlanmaz.

> **Bu araç karar vermez.** Kredinin sınıflandırılması, aşaması ve karşılık kararı kurumunuzun yetkili birimine
> aittir. Puanlar yalnızca inceleme sırasını belirler.

## Sinyaller ve örnek puanlar

| Sinyal | Kaynak | Kural (varsayılan) | Puan |
|---|---|---|---|
| Gecikme | Portföy dökümü | ≥1 / ≥31 / ≥61 / ≥91 gün | 10 / 25 / 40 / 60 |
| Limit aşımı | Portföy (risk ÷ limit) | risk > limit | 15 |
| Limit doluluğu | Portföy | kullanım ≥ %95 | 5 |
| Karşılıksız çek | KKB / Risk Merkezi dökümü | son 12 ay, açık kayıt başına 15, ödenmiş 5 | en çok 30 |
| Protestolu senet | KKB / Risk Merkezi dökümü | son 12 ay, açık 10, ödenmiş 3 | en çok 25 |
| Haciz | Hacizler (İİK 89 / 6183) | kamu alacağı 15, özel alacaklı 12 | en çok 30 |
| Ciro düşüşü | Aylık ciro / hesap girişleri | son 3 ay, geçen yılın aynı aylarına göre ≥ %30 / ≥ %50 | 10 / 20 |

**Hesaplama**
- **Toplam puan:** Sinyal puanları toplanır, en çok 100 olur.
- **Sınıf:** ≥70 Kritik, ≥40 Yakın İzleme, ≥20 İzleme, diğerleri Normal.
- **Ciro karşılaştırması:** Mevsimselliği dışlamak için geçen yılın aynı 3 ayıyla yapılır. O veri yoksa önceki 3
  ayla karşılaştırılır.
- **Gecikme notları:** 30 günü aşan gecikmeye "yakın izleme (Aşama 2)", 90 günü aşana "donuk alacak (Aşama 3)"
  değerlendirme notu eklenir. Bu notlar BDDK'nın Kredilerin Sınıflandırılması ve Karşılıklar Yönetmeliği ile
  TFRS 9'daki 30/90 gün göstergelerine dayanır.
- **Portföyde olmayanlar:** Sinyali olup portföy dosyasında bulunmayan müşteriler "(portföyde yok)" olarak listelenir.

**Ayarlar:** Tüm puan ve eşikler `ayarlar.json` dosyasındadır. Bu değerler **örnektir**; kurumunuzun kredi
politikasına ve geçmiş temerrüt verisine göre kalibre edin.

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                                    # örnek: 7 firma, 30.09.2026
python main.py --portfoy portfoy.xlsx --cek-senet kkb.xlsx --haciz haciz.xlsx --ciro ciro.xlsx --tarih 30.09.2026
python main.py ... --ayar ayarlar.json
```

| Dosya | Sütunlar |
|---|---|
| Portföy | Müşteri No, Firma Adı, Şube, Segment, Limit, Risk, Gecikme Günü |
| Çek / senet | Müşteri No, Tarih, Tür (Karşılıksız Çek / Protestolu Senet), Tutar, Durum (Açık / Ödendi) |
| Haciz | Müşteri No, Tarih, Alacaklı, Tür, Tutar. Alacaklı ya da türde vergi, SGK, belediye, gümrük veya 6183 geçiyorsa kamu alacağı sayılır. |
| Ciro | Müşteri No, Dönem (2026-09 veya 09.2026), Ciro |

Müşteri No yerine VKN de kullanılabilir; tüm dosyalarda aynı anahtar olmalıdır.

## Çıktı

`Özet` (sınıf ve sinyal türü dağılımı) · `Firma Listesi` (puan sıralı; İzleme Kararı ve Açıklama sütunları boş) ·
`Sinyal Detayı` · `Puan Tablosu` (kullanılan kurallar)

## Dikkat

- **Gizlilik:** KKB ve Risk Merkezi verileri bankacılık sırrı ve kişisel veri niteliğindedir. Dosyaları yalnızca
  yetkili ortamlarda işleyin; çıktıyı kurum dışına göndermeyin.
- **Veri güncelliği:** Sinyaller dosyalardaki veriyle sınırlıdır. Gecikme günü ve ödenmiş/kapanmış durumunun güncel
  olduğundan emin olun.
- **Ek bilgiler:** Puan, firmanın sektörü, teminat yapısı ve grup riski gibi bilgileri içermez; değerlendirmede
  bunları ayrıca dikkate alın.

## Testler

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
