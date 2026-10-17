# Kanal Maliyet Analizi · Kod Bloğu

> Rezervasyon ve Gelir Yönetimi › Gelir Müdürü (Revenue Manager) · Gerisi Bende

OTA, acente, GDS ve direkt kanallardan gelen oda gelirini **komisyon, işlem ücreti, rezervasyon başı ücret ve sabit kanal
maliyetleri düşüldükten sonra** karşılaştırır. Her kanal için net gelir, net ADR ve en kârlı kanala göre oda-gece başına farkı
gösterir. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Komisyon | Rezervasyon tutarı × komisyon %; satırda Komisyon % yazılıysa o, yoksa kanal tanımındaki oran |
| İşlem ücreti | Brüt gelir × işlem ücreti % (sanal POS, ödeme altyapısı) |
| Rezervasyon başı ücret | Rezervasyon sayısı × ücret (GDS, rezervasyon motoru vb.) |
| Sabit maliyet | Aylık sabit maliyet × ay sayısı (giriş tarihlerinin kapsadığı ay ya da `--ay`) |
| Maliyet % | Toplam kanal maliyeti ÷ brüt gelir × 100 |
| Net gelir / net ADR | Brüt − toplam maliyet · net gelir ÷ oda-gece |
| Net ADR farkı | Kanalın net ADR'si − en yüksek net ADR'li kanalınki |

İptal ve no-show durumundaki rezervasyonlar alınmaz. Gece sütunu yoksa çıkış − giriş kullanılır.

**Uyarılar:** kanal maliyet tablosunda olmayan kanal (maliyet sıfır sayılır) · sabit maliyeti olduğu hâlde satış getirmeyen
kanal (Orta) · satırdaki komisyonun kanal tanımından 1 puandan fazla farklı olması · maliyet yüzdesi `--esik` (varsayılan %25)
üstünde (Bilgi).

**Örnek:** OTA'dan 6.000 TL'lik rezervasyon %18 ve 3.000 TL'lik rezervasyon %20 komisyonla → komisyon 1.680 TL, maliyet
**%18,67**; 3 oda-gece için net ADR **2.440 TL**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                             # örnek veri (eylül–ekim 2026)
python main.py --rezervasyonlar rezervasyonlar.xlsx --kanallar kanallar.xlsx
python main.py --rezervasyonlar rezervasyonlar.xlsx --kanallar kanallar.xlsx --ay 3 --esik 20
```

Kanal tablosu verilmezse yalnız satırlarda yazan komisyonlar maliyet sayılır.

## Çıktı

`cikti/kanal_maliyet_analizi.xlsx`: `Kanal Karşılaştırma` (brüt gelire göre sıralı, toplam satırıyla), `Uyarılar`.

## Dikkat

- **Tüm tutarlar aynı esasta olmalı.** KDV dahil ve hariç tutarları karıştırmayın; komisyon matrahı sözleşmenize göre
  (KDV dahil / hariç, konaklama vergisi dahil / hariç) farklıysa tutarları ona göre hazırlayın.
- **Net fiyatlı (allotment / kontrat) acentelerde** tutar zaten net ise komisyonu 0 yazın; aksi hâlde maliyet iki kez düşülür.
- **Analiz yalnız doğrudan kanal maliyetini kapsar.** Rezervasyon ekibi, iptal oranı, ödeme vadesi ve tahsilat riski, misafir
  harcaması ve sadakat etkisi hesapta yoktur; kanal kararı vermeden önce bunları da değerlendirin.
- **Örnek veri:** Kanallar, oranlar ve tutarlar kurgusaldır.

## Testler

- Komisyon (satır ve kanal oranı), işlem ücreti, rezervasyon başı ücret, ay sayısına göre sabit maliyet, net gelir, net ADR,
  en iyi kanal farkı, toplam satırı.
- Tanımsız kanal, satışsız kanal, farklı komisyon uyarıları; iptallerin dışarıda kalması, okuma ve komut satırı.

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
