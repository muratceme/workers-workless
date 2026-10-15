# Uyum Programı Gösterge Raporu · Kod Bloğu

> Kurumsal Uyum › Uyum Müdürü · Workers / Workless

İnceleme, bildirim, eğitim ve yaptırım taraması gibi uyum göstergelerini **dönem dönem tek tabloda** toplar; her
göstergeyi hedefiyle ve önceki dönemle karşılaştırır, hedef dışı kalanları ve belirgin kötüleşmeleri işaretler.
Dönemsel uyum raporunun gösterge ekidir. İnternete bağlanmaz.

## Nasıl hesaplar?

| Kavram | Kural |
|---|---|
| Son dönem | Dosyadaki en büyük dönem etiketi (metin sırasıyla: `2026-Q1` < `2026-Q2`; `2026-09` < `2026-10`) |
| Değişim % | (Son dönem − önceki dönem) ÷ \|önceki dönem\| × 100 |
| Eğilim | "Yüksek iyi" göstergede artış İyileşti, azalış Kötüleşti; "Düşük iyi" göstergede tersi; yönsüzde Arttı / Azaldı |
| Hedef durumu | Yüksek iyi: değer ≥ hedef → Hedefte · Düşük iyi: değer ≤ hedef → Hedefte · değilse Hedef dışı |
| Hedefe sapma % | (Değer − hedef) ÷ \|hedef\| × 100 |
| Belirgin kötüleşme | Kötü yöndeki değişim `--degisim` (25) yüzde ve üstünde |

**Örnek:** "Ortalama uyarı inceleme süresi" (düşük iyi, hedef 5 gün): 4,8 → 6,5 gün. Değişim **+%35,42**, eğilim
**Kötüleşti**, hedef **dışı** (sapma +%30).

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                            # örnek veri
python main.py --gostergeler gostergeler.xlsx
python main.py --gostergeler gostergeler.xlsx --degisim 20 --donem 6
```

Her satır bir göstergenin bir dönemdeki değeridir. Zorunlu sütunlar: Dönem, Gösterge, Değer. Dönemi sıralanabilir
biçimde yazın (`2026-Q3`, `2026-09`). Yön boşsa gösterge yalnız izlenir, hedefle karşılaştırılmaz.

## Çıktı

`cikti/uyum_gosterge_raporu.xlsx`:

- `Gösterge Tablosu`: önce hedef dışı kalanlar; son dönemler yan yana, değişim, eğilim, hedef, sapma, durum ve sizin
  dolduracağınız açıklama sütunu.
- `Özet`: hedefte, hedef dışı, hedefsiz ve verisi eksik gösterge sayıları.
- `Uyarılar`: hedef dışı göstergeler, belirgin kötüleşme ve değişimler, son dönemde verisi olmayan göstergeler.

## Dikkat

- **Göstergeler ve hedefler sizindir.** Hangi göstergelerin izleneceği ve hedeflerin ne olacağı kurumunuzun uyum
  programına, risk değerlendirmesine ve yönetim kurulunun beklentisine bağlıdır. Örnek dosyadaki göstergeler ve
  hedefler yalnız gösterim içindir; bir kontrol listesi ya da mevzuat gereği değildir.
- **Sayı tek başına anlam taşımaz.** Şüpheli işlem bildirimi ya da incelenen uyarı sayısının artması iyi de kötü de
  olabilir; bu yüzden bu tür göstergeleri yönsüz bırakın ve açıklama sütununda yorumlayın. Bildirim sayısı için hedef
  koymak sakıncalıdır.
- **Rapor hüküm vermez.** Göstergelerin hedefte olması uyum programının yeterli olduğunu göstermez; nitel
  değerlendirme, bulgular ve düzeltici faaliyetler raporun asıl gövdesidir.
- **Gizlilik.** Şüpheli işlem bildirimlerine ilişkin bilgiler gizlidir; raporda toplu sayıların ötesine geçmeyin,
  müşteri ya da işlem düzeyinde bilgi yazmayın ve raporu yalnız yetkili kişilerle paylaşın.
- **Dönem etiketlerini tutarlı yazın.** `2026-Q3` ile `Q3 2026` ayrı sıralanır. Tarih hücreleri yıl-ay olarak okunur.
- **Örnek veri:** Göstergeler, değerler ve hedefler kurgusaldır.

## Testler

- Hedef durumu sınırları (tam hedefte, iki yön), hedefi ya da yönü olmayan gösterge, sıfır hedef.
- Değişim ve eğilim (iyileşme, kötüleşme, yönsüz, önceki dönem sıfır), belirgin değişim eşiği.
- Son dönemde verisi olmayan gösterge, aynı dönemin iki kez girilmesi, dönem sıralaması ve sayısı, okuma, komut satırı.

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
