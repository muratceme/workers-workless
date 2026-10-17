# Tarife Etki Simülasyonu · Kod Bloğu

> Aktüerya › Aktüerya Uzmanı · Gerisi Bende

Tarife değişikliği senaryolarını mevcut portföye uygular; **prim hacmine ve hasar/prim (H/P) oranına** etkisini
senaryo, segment ve poliçe bazında gösterir. İnternete bağlanmaz.

## Nasıl hesaplar?

Senaryo dosyasındaki her satır bir kuraldır: `Alan` portföy dosyasındaki bir sütun adı, `Değer` o sütundaki değer,
`Değişim %` prim değişimi. `Alan` yerine `*` yazılırsa kural tüm poliçelere uygulanır.

```
Yeni prim      = prim × (1 + değişim₁) × (1 + değişim₂) × …        poliçeye uyan tüm kurallar zincirleme
Kalma oranı    = (1 − kayıp₁) × (1 − kayıp₂) × …                   zam nedeniyle ayrılacağı varsayılan pay
Beklenen prim  = yeni prim × kalma oranı
Beklenen hasar = hasar × kalma oranı
H/P %          = beklenen hasar ÷ beklenen prim × 100
```

- **Azami artış:** `--azami-artis 25` verilirse bir poliçenin primi en çok %25 artar; tavana takılan poliçe sayılır.
- **Kayıp %** sizin varsayımınızdır. Ayrılan poliçelerin primi de hasarı da aynı oranda düşer; yani ayrılanların
  portföy ortalamasıyla aynı riskte olduğu varsayılır.
- **Hedef H/P:** `--hedef-hp 60` verilirse gereken genel prim değişimi = mevcut H/P ÷ hedef H/P − 1 (kayıpsız).
- **Örnek:** prim 4.000, hasar 2.600 (H/P %65). C sınıfına (prim 3.000) %20 zam → prim 4.600, H/P %56,52.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri
python main.py --portfoy portfoy.xlsx --senaryolar senaryolar.xlsx --azami-artis 25 --hedef-hp 60
```

| Dosya | Sütunlar |
|---|---|
| Portföy | Poliçe No, Prim (zorunlu), Hasar + istediğiniz segment sütunları (Branş, Risk Sınıfı, Bölge ...) |
| Senaryolar | Senaryo, Alan (`*` = tümü), Değer, Değişim %, Kayıp %, Açıklama |

Poliçe No, Prim ve Hasar dışındaki tüm sütunlar segment sayılır ve kurallarda kullanılabilir. Aynı adı taşıyan
satırlar tek senaryonun kurallarıdır.

## Çıktı

`cikti/tarife_etki_simulasyonu.xlsx`:

- `Senaryo Özeti`: mevcut durum ve her senaryo için tarife primi, beklenen prim, değişim, H/P, H/P farkı, en yüksek artış.
- `Segment Etkisi`: her segment sütunu ve değeri için prim değişimi ve H/P (önce / sonra).
- `Poliçe Etkisi`: poliçe bazında her senaryodaki yeni prim ve değişim.
- `Uyarılar`: okunamayan satırlar, portföyde olmayan alanlar, hiçbir poliçeyi etkilemeyen senaryolar.

## Dikkat

- **Statik bir "olduğu gibi" hesaptır.** Geçmiş hasarların aynen tekrarlanacağı varsayılır. Hasar eğilimi, enflasyon,
  IBNR ve muallak gelişimi, hasar masrafları, komisyon, reasürans ve yeni iş hesaba girmez. Aktüeryal fiyatlama
  çalışmasının yerine geçmez.
- **Prim ve hasar aynı dönemi kapsamalı.** Prim sütununa kazanılmış primi, hasar sütununa aynı dönemin gerçekleşen
  hasarını (ödenen + muallak) yazın; yoksa H/P anlamsız olur.
- **Kayıp oranı varsayımdır.** Gerçekte zamdan sonra önce iyi riskler ayrılabilir (ters seçim); bu durumda H/P
  buradaki sonuçtan kötü çıkar.
- **Küçük segmentlerde H/P oynaktır.** Birkaç poliçelik segmentin H/P oranına dayanarak zam kararı vermeyin.
- **Karar değildir.** Tarife değişikliği şirketinizin fiyatlama ve onay süreçlerine tabidir.
- **Örnek veri:** Poliçeler, primler ve hasarlar kurgusaldır.

## Testler

- Tek kural, zincirleme kurallar, poliçe başına azami artış.
- Kayıp varsayımıyla beklenen prim, hasar ve H/P.
- Segment etkisi, hedef H/P için gerekli değişim, hiçbir poliçeye uymayan senaryo.
- Okuma hataları ve örnek veri uçtan uca.

```bash
python -m unittest discover -s tests
```

## Sorumluluk

Bu paketi indiren ve çalıştıran kişi aşağıdakileri kabul etmiş sayılır:

- **Tüm risk kullanıcıdadır.** Paket ücretsizdir ve "olduğu gibi" sunulur; doğruluk, eksiksizlik veya belirli bir amaca
  uygunluk konusunda hiçbir garanti verilmez. Paketi kendi bilgisayarınızda, kendi verinizle ve kendi kararınızla
  çalıştırırsınız. Hatalı veya eksik sonuç, kaçırılan süre, vergi / SGK cezası, veri kaybı, mali kayıp ve diğer her türlü
  doğrudan ya da dolaylı zarardan yazarlar ve katkıda bulunanlar sorumlu tutulamaz.
- **Sonuçları kontrol etmek sizin işinizdir.** Çıktılar karar desteği içindir. Beyan vermeden, ödeme yapmadan, sözleşme
  imzalamadan, personel / kredi / hasar / tedavi kararı almadan önce sonuçları yetkin bir kişi kontrol etmelidir.
- **Danışmanlık değildir.** Hukuki, mali, vergisel, tıbbi veya yatırım danışmanlığı yerine geçmez.
- **Mevzuat değişir.** Koddaki süre, oran, eşik ve kurallar yazıldığı tarihteki bilgiye dayanır ve eskiyebilir; özel
  durumları kapsamayabilir. Güncelliğini kontrol etmek kullanıcının sorumluluğundadır.
- **Örnek veri kurgusaldır.** Kişi, kurum, tutar ve tarihler gerçek değildir; kendi verinizle sonuçlar farklı olur.
- **Kişisel veri.** Verinizin KVKK, şirket politikaları ve sektör düzenlemelerine uygun işlenmesinden, saklanmasından ve
  paylaşılmasından siz sorumlusunuz.

Ayrıntılar: `LICENSE` (MIT) ve `SORUMLULUK_REDDI.md`.
