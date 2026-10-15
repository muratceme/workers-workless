# Gişe Gün Sonu Kasa Mutabakatı · Kod Bloğu

> Şube Bankacılığı › Gişe Yetkilisi · Workers / Workless

Gün sonu kasa sayımını (kupür bazında) sistem bakiyesiyle **gişe ve para birimi bazında** karşılaştırır. Fark varsa
günün işlemleri içinden araştırılacak adayları listeler. İnternete bağlanmaz.

## Nasıl hesaplar?

| Terim | Hesap |
|---|---|
| Sayım | Σ kupür × adet |
| Hesaplanan | Devir bakiyesi + günün girişleri − çıkışları (işlem dosyası verilirse) |
| Fark | Sayım − sistem bakiyesi (+ kasa fazlası, − kasa noksanı) |

Hesaplanan bakiye sistem bakiyesini tutmuyorsa ayrıca uyarı verilir (işlem dökümü eksik olabilir).

**Araştırma ipuçları** (her biri bir olasılıktır):

| İpucu | Mantık |
|---|---|
| Tutar farka eşit | Giriş kaydı var, para kasaya girmemiş; veya çıkış iki kez ödenmiş olabilir |
| Ters yön kaydı | İşlem tutarı farkın yarısı: giriş yerine çıkış (veya tersi) kaydedilmişse fark tutarın iki katı olur |
| Rakam yer değişimi | Fark 9'a bölünüyorsa: iki komşu rakamı yer değiştirince fark kadar değişen işlemler (5.400 ↔ 4.500) |
| Kupür sayımı | Fark tek bir kupürün en çok 5 katı: o kupür yeniden sayılmalı |
| Gişeler arası | Aynı para biriminde bir gişe fazla, diğeri aynı tutarda noksan |
| Döviz işlemi | Aynı gişede TL ve döviz farkı ters yönlü; fark oranı günün işlem kurlarıyla (±%2) karşılaştırılır |

Kasa limiti girilmişse limiti aşan tutar "ana kasaya devredilmeli" notuyla raporlanır.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek veri
python main.py --sayim sayim.xlsx --sistem sistem.xlsx --islemler islemler.xlsx
python main.py --sayim sayim.xlsx --sistem sistem.xlsx --tolerans 1
```

| Dosya | Sütunlar |
|---|---|
| Sayım | Gişe, Para Birimi (boşsa TL), Kupür, Adet |
| Sistem | Gişe, Para Birimi, Devir Bakiyesi, Sistem Bakiyesi, Kasa Limiti |
| İşlemler (isteğe bağlı) | İşlem No, Saat, Gişe, İşlem Türü, Para Birimi, Giriş, Çıkış, Kur, Açıklama |

İşlem dosyası verilmezse yalnız fark ve kupür ipucu üretilir.

## Çıktı

`cikti/kasa_mutabakati.xlsx`:

- `Mutabakat`: gişe ve para birimi bazında devir, giriş, çıkış, hesaplanan, sistem, sayım, fark, durum; boş "Sayan" ve
  "Kontrol Eden" sütunları.
- `Kupür Dökümü`: sayım tutanağına eklenecek kupür listesi.
- `Araştırma Listesi`: ipucu, ilgili işlem ve neye bakılacağı; boş "Sonuç" sütunu.
- `Uyarılar`: farklar, limit aşımları, okuma hataları.

## Dikkat

- **İpuçları kanıt değildir.** Fark ancak fiş, dekont, müşteri imzası ve gerektiğinde kamera kaydı incelenerek
  açıklanır. Liste, nereden başlanacağını gösterir.
- **Kasa farkı işlemleri** (geçici hesaba alma, tutanak, zimmet) bankanızın iç mevzuatına göre yürütülür; bu paket
  muhasebe kaydı önermez.
- **Döviz farkları** kendi para biriminde raporlanır; TL karşılığı hesaplanmaz.
- **Kupür listesi** tedavüldeki TL, USD, EUR, GBP ve CHF kupürlerini içerir. Tanınmayan kupür uyarı verir ama toplama
  katılır. Yıpranmış, sahte şüpheli veya tedavülden kalkmış paralar ayrıca ele alınmalıdır.
- **Kişisel veri:** İşlem açıklamalarında müşteri bilgisi varsa çıktıyı yetkisiz kişilerle paylaşmayın.
- **Örnek veri:** Gişeler, işlemler ve kurlar kurgusaldır.

## Testler

- Mutabık kasa, tolerans, fazla / noksan yönü, sistem bakiyesinin işlemlerle tutmaması, limit aşımı.
- İpuçları: tutar eşitliği (giriş / çıkış ayrımı), ters yön, rakam yer değişimi, kupür, gişeler arası, döviz oranı.
- Örnek veri uçtan uca (4 fark) ve okuma hataları.

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
