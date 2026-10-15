# Kampanya Fiyat Simülasyonu · Kod Bloğu

> Pazaryeri Yönetimi › Pazaryeri Uzmanı · Workers / Workless

Planlanan **indirim ve kupon** kurgusunun her ürünün birim kârına ve marjına etkisini hesaplar; kampanyada zarar eden
ürünleri, **başabaş fiyatı** ve toplam kârı korumak için gereken **adet artışını** gösterir. İnternete bağlanmaz.

## Nasıl hesaplar?

| Kalem | Formül |
|---|---|
| Müşterinin ödediği (KDV dahil) | Satış fiyatı × (1 − indirim %) − kupon |
| Net satış (KDV hariç) | Müşterinin ödediği ÷ (1 + KDV %) |
| Komisyon | Komisyon % × müşterinin ödediği (`--matrah dahil`, varsayılan) ya da × net satış (`--matrah haric`) |
| Birim kâr | Net satış − komisyon − kargo − hizmet bedeli − birim maliyet |
| Marj % | Birim kâr ÷ net satış × 100 |
| Gerekli adet artışı % | (Mevcut birim kâr ÷ kampanya birim kârı − 1) × 100 |
| Başabaş fiyat | Birim kârın sıfır olduğu müşteri fiyatı (KDV dahil) |
| Min. marj fiyatı | Marjın `--min-marj` (varsayılan %10) olduğu müşteri fiyatı |
| Toplam kâr farkı | Kampanya birim kârı × beklenen adet − mevcut birim kâr × mevcut adet |

Kargo, hizmet bedeli ve maliyet **KDV hariç** yazılır (indirilebilir KDV maliyet sayılmaz). KDV sütunu boşsa `--kdv`
(varsayılan 20) kullanılır.

**Örnek:** Kahve makinesi 2.499 TL, %15 indirim + 100 TL kupon → müşteri 2.024,15 öder; net 1.686,79; komisyon
%15 = 303,62; kargo 65, hizmet 8, maliyet 1.150 → birim kâr **160,17** (marj %9,50). Mevcut birim kâr 484,65
olduğu için toplam kârı korumak üzere adedin **%202,59** artması gerekir.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                     # örnek veri
python main.py --urunler kampanya_urunleri.xlsx --min-marj 12
python main.py --urunler kampanya_urunleri.xlsx --matrah haric --kdv 10
```

Zorunlu sütunlar: Ürün Kodu, Satış Fiyatı (KDV dahil), Birim Maliyet. İndirim ve komisyon yüzde olarak yazılır.
Mevcut ve beklenen adet yazılırsa toplam kâr karşılaştırılır.

## Çıktı

`cikti/kampanya_fiyat_simulasyonu.xlsx`:

- `Simülasyon`: mevcut ve kampanya birim kârı, marjlar (renkli), gerekli adet artışı, başabaş ve min. marj fiyatı,
  toplam kâr farkı.
- `Birim Kırılım`: iki senaryo için fiyat → net → komisyon → giderler → kâr.
- `Uyarılar`: kampanyada ya da mevcut fiyatta zarar, marj eşik altında, beklenen adetin kârı korumaması.

## Dikkat

- **Pazaryeri kuralları farklıdır.** Komisyonun hangi tutar üzerinden hesaplandığı, kuponun kim tarafından
  karşılandığı, kargo anlaşması, hizmet / işlem bedelleri ve kampanya katılım koşulları pazaryerine ve kategoriye göre
  değişir. Kendi sözleşmenizdeki oranları ve matrahı yazın; kod bir pazaryerinin güncel kuralını bilmez.
- **Satıcı kuponu varsayılır.** Pazaryerinin karşıladığı kupon ya da indirim satıcının gelirini düşürmüyorsa kupon
  sütununa yazmayın.
- **İndirim ilanı kurallarına uyun.** İndirimli satış ilanlarında önceki fiyatın gösterilmesi ve indirimin gerçek
  olması ile ilgili tüketici mevzuatındaki kurallar kampanya kurgusundan ayrıca kontrol edilmelidir.
- **İade, iptal ve reklam maliyeti** hesaba girmez; gerekiyorsa birim başına tutarı hizmet bedeline ekleyin.
- **Örnek veri:** Ürünler, fiyatlar ve oranlar kurgusaldır.

## Testler

- Müşteri fiyatı, net satış, KDV dahil / hariç komisyon matrahı, birim kâr ve marj.
- Gerekli adet artışı, başabaş ve min. marj fiyatı, toplam kâr farkı, uyarılar, okuma ve komut satırı.

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
