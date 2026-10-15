# Tarihsel Riske Maruz Değer (VaR) · Kod Bloğu

> İç Kontrol ve Risk Yönetimi › Risk Yönetimi Uzmanı · Workers / Workless

Portföy pozisyonları ve geçmiş fiyat serilerinden **tarihsel simülasyonla** günlük VaR ve beklenen kaybı (ES) hesaplar;
pozisyon katkılarını ve geriye dönük test sonucunu raporlar. İnternete bağlanmaz.

## Nasıl hesaplar?

| Adım | Formül |
|---|---|
| Günlük getiri | r(t) = P(t) ÷ P(t−1) − 1 |
| Senaryo kâr / zararı | K/Z(t) = Σ pozisyon değeri × r(t). Bugünkü pozisyonlar geçmişin her gününe uygulanır |
| VaR | k = ⌊N × (1 − güven)⌋ + 1; en kötü k'ıncı senaryonun zararı |
| Beklenen kayıp (ES) | En kötü k senaryonun ortalama zararı (VaR senaryosu dahil) |
| h günlük VaR | VaR(1 gün) × √h |
| Bireysel VaR | Pozisyon tek başına portföy olsaydı VaR'ı |
| Çeşitlendirme etkisi | Σ bireysel VaR − portföy VaR |
| ES katkısı | Pozisyonun, portföyün en kötü k senaryosundaki ortalama zararı |

Örnek: 250 senaryo ve %99 güvende k = 3; VaR en kötü üçüncü günün zararı, ES en kötü üç günün ortalamasıdır.

- **ES katkıları** toplanınca portföy ES'ini verir. Eksi katkı, pozisyonun en kötü günlerde kazandırdığını gösterir.
- **Geriye dönük test:** Her gün için önceki `pencere` günün senaryolarıyla VaR hesaplanır ve ertesi günün K/Z'si ile
  karşılaştırılır. Zarar VaR'ı aşarsa "aşım" sayılır. Pozisyonlar sabit varsayılır. 250 gün ve %99 güvende sonuç
  Basel geriye dönük test çerçevesindeki bölgelerle gösterilir: 0–4 aşım yeşil, 5–9 sarı, 10 ve üzeri kırmızı.
- **Veri kontrolleri:** eksik fiyatlı günler atlanır; günlük %20'yi aşan getiri ve 5 gün üst üste değişmeyen fiyat
  için uyarı verilir.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek veri
python main.py --pozisyonlar pozisyonlar.xlsx --fiyatlar fiyatlar.xlsx --guven 99 --pencere 250 --elde-tutma 10
```

| Dosya | Biçim |
|---|---|
| Pozisyonlar | Enstrüman, Risk Faktörü, Piyasa Değeri (kısa pozisyon eksi), Portföy |
| Fiyatlar | İlk sütun Tarih; her risk faktörü için bir sütun (günlük kapanış). Sütun adı "Risk Faktörü" ile aynı olmalı |

Tüm pozisyon değerleri aynı para biriminde (ör. TL) olmalıdır. Döviz pozisyonu için risk faktörü kurun kendisidir.

## Çıktı

`cikti/var_raporu.xlsx`:

- `Özet`: VaR, ES, brüt değere oranı, çeşitlendirme etkisi, h günlük değerler, geriye dönük test.
- `Pozisyonlar`: ağırlık, oynaklık, bireysel VaR / ES, VaR günündeki zarar, ES katkısı ve payı.
- `Senaryolar`: en kötüden en iyiye tüm senaryolar; VaR senaryosu işaretli.
- `Geriye Dönük Test`: aşım günleri.
- `Uyarılar`: veri sorunları, kısa geçmiş, yoğunlaşma.

## Dikkat

- **Yöntemin sınırı:** Tarihsel simülasyon, pencerede yaşanmamış bir şoku göremez. Sakin bir dönem düşük, çalkantılı
  bir dönem yüksek VaR üretir. Stres testiyle birlikte kullanılmalıdır.
- **Doğrusal pozisyonlar:** Kâr / zarar "değer × getiri" olarak hesaplanır. Opsiyon gibi doğrusal olmayan ürünler ve
  tahvillerde vade kısalması doğru yansımaz. Tahvil için fiyat serisi yoksa bu yöntem uygun değildir.
- **Karekök kuralı** getirilerin bağımsız ve aynı dağılımlı olduğunu varsayan bir yaklaşımdır.
- **Küçük örnek:** %99 güvende 250 senaryoyla VaR yalnız en kötü üç güne dayanır; tek bir gün sonucu belirgin
  değiştirebilir.
- **Yasal sermaye hesabı değildir.** Piyasa riski sermaye yükümlülüğü BDDK düzenlemelerine ve bankanızın onaylı
  modeline göre hesaplanır; bu paket iç izleme ve kontrol amaçlıdır.
- **Fiyat verisi:** Bölünme ve temettüye göre düzeltilmiş fiyat kullanın.
- **Örnek veri:** Fiyat serileri rastgele üretilmiştir; gerçek piyasa verisi değildir.

## Testler

- Sıra kuralı (N = 100, 250, 500), getiri, tek pozisyonda VaR / ES / karekök.
- Tam korunmada VaR = 0; ES katkılarının toplamı = ES; çeşitlendirme etkisi ≥ 0.
- Geriye dönük testte aşım sayımı; kısa geçmiş ve veri uyarıları.
- Dosya okuma (eksik fiyat, eksik risk faktörü) ve örnek veri uçtan uca.

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
