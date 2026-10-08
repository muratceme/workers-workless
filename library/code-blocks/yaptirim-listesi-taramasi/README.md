# Yaptırım Listesi Taraması · Kod Bloğu

> Bankacılık › Kurumsal Uyum › Uyum Uzmanı · Workers / Workless

Müşteri ve karşı taraf listesini yaptırım listeleriyle **bulanık eşleştirir**. Türkçe karakterleri, yazım ve
transliterasyon farklarını dikkate alır. Doğum tarihi ve uyrukla yanlış alarmları azaltır. İnternete bağlanmaz;
listeleri kendiniz indirip verirsiniz.

## Ne yapar?

**Liste kaynakları**
- **BM Güvenlik Konseyi Konsolide Listesi:** Resmî XML dosyası doğrudan okunur. Kişiler, kuruluşlar, takma adlar,
  doğum tarihleri, uyruk ve referans numaraları alınır.
- **Diğer listeler:** Resmî Gazete'de yayımlanan malvarlığı dondurma kararları, OFAC, AB ve kurum içi listeler
  için şu sütunları içeren bir CSV/Excel hazırlayın: `Ad, Diğer Adlar (; ile), Tür, Doğum Tarihi, Uyruk, Liste,
  Referans`.

**Ad normalleştirme**
- **Karakter farkları:** Türkçe ve aksanlı harfler, büyük/küçük harf ve noktalama farkları yok sayılır.
- **Unvanlar ve ekler:** Kişilerde unvanlar (Dr., Hacı, Al-, Bin, Abu…), kuruluşlarda şirket ekleri (A.Ş., Ltd.
  Şti., LLC, FZE, "ve"…) atılır.
- **Yaygın transliterasyonlar:** Mohammed/Muhammed/Mehmet, Hussein/Hüseyin, Yousef/Yusuf, Ahmed/Ahmet,
  Abdallah/Abdullah, Omar/Ömer, Othman/Osman… ve ph/f, w/v, q/k, kh/h, çift harf farkları eşitlenir.

**Eşleştirme**
- **Ad benzerliği:** Kelimeler sıradan bağımsız eşlenir. "Yılmaz Ahmet" ile "Ahmet Yılmaz" aynıdır. Baş harfler
  ("M. Kurgusali") de eşleşir. Takma adların her biri ayrıca denenir.
- **Tarih ve uyruk düzeltmesi:**

| Durum | Puan etkisi |
|---|---|
| Doğum tarihi tam eşleşiyor | +8 |
| Doğum yılı eşleşiyor | +5 |
| Doğum yılı farkı 2 yıldan fazla | −15 |
| Uyruk eşleşiyor | +3 |
| Uyruk farklı | −5 |
| Kişi/kuruluş türü farklı | −10 |

**Sonuç seviyeleri:** **Güçlü eşleşme** (≥ 95) · **Olası eşleşme** (≥ `--esik`, varsayılan 85) · **Zayıf
benzerlik** (≥ `--zayif`, varsayılan 75)

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                                           # kurgusal örnek listelerle
python main.py --musteriler musteriler.xlsx --liste consolidated.xml --liste ulusal_kararlar.xlsx
python main.py --musteriler karsi_taraflar.csv --liste liste.csv --esik 80 --zayif 70
```

## Çıktı

`Özet` · `Eşleşmeler` (puan, eşleşen ad biçimi, doğum ve uyruk karşılaştırması, analist kararı ve açıklama
sütunları) · `Bilgi`

## Dikkat

- **Ön eleme:** Bu tarama bir ön elemedir. Eşleşmeler kimlik bilgileriyle doğrulanmadan sonuç çıkarılmaz:
  doğum tarihi, uyruk, kimlik/vergi numarası ve adres.
- **Gerçek eşleşme:** İlgili mevzuat ve kurum prosedürü uygulanır. Örneğin 6415 ve 7262 sayılı Kanunlar
  kapsamındaki malvarlığı dondurma kararları.
- **Listelerin güncelliği:** Kullanıcının sorumluluğundadır. Listeler sık güncellenir; taramayı düzenli
  tekrarlayın.
- **Eşik ayarı:** Eşikleri düşürmek yanlış alarmı, yükseltmek kaçırma riskini artırır. Kurumunuzun risk
  iştahına göre ayarlayın ve sonuçları örneklemle test edin.
- **Örnek veriler:** Örnek listelerdeki kişi ve kuruluşlar **kurgusaldır**.

## Testler

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
