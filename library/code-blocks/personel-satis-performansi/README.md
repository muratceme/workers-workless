# Personel Satış Performansı · Kod Bloğu

> Mağaza Operasyonları › Mağaza Müdürü · Workers / Workless

Satış danışmanlarının **ciro, fiş, adet ve sepet ortalamalarını** karşılaştırır; hedefler ve prim kademeleri
verilirse **prim tablosunu** hazırlar. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Sepet ortalaması | Ciro ÷ fiş adedi |
| Fiş başına adet | Satılan adet ÷ fiş adedi |
| Ciro payı % | Personelin cirosu ÷ mağaza cirosu × 100 |
| Mağazaya göre % | Personelin sepet ortalaması (ya da fiş başına adedi) ÷ mağazanınki × 100 |
| Günlük ortalama ciro | Ciro ÷ satış kaydı olan gün sayısı |
| Hedef gerçekleşme % | Ciro ÷ hedef ciro × 100 |
| Prim oranı | Gerçekleşmenin ulaştığı en yüksek kademenin oranı |
| Prim tutarı | Ciro × prim oranı ÷ 100 |

Kademeler `--kademeler` ile `gerçekleşme:prim` biçiminde verilir: `90:1;100:2;120:3` → gerçekleşme %90 ve üstünde
cironun %1'i, %100 ve üstünde %2'si, %120 ve üstünde %3'ü. En düşük kademenin altında prim sıfırdır. Kademe
verilmezse prim hesaplanmaz; kod varsayılan oran kullanmaz (örnek veri bu kademelerle çalışır).

**Örnek:** Ciro 108.000, hedef 100.000 → gerçekleşme **%108**; %100 kademesine ulaştı → prim 108.000 × %2 =
**2.160**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                         # örnek veri ve örnek kademeler
python main.py --satislar personel_satislari.xlsx
python main.py --satislar personel_satislari.xlsx --hedefler hedefler.xlsx --kademeler 95:0,5;100:1,5;115:2,5 --esik 85
```

Zorunlu sütunlar: satışlarda Personel ve Ciro; hedeflerde Personel ve Hedef Ciro. Satış dosyasında her satır bir
personelin bir günü ya da tek bir satışı olabilir; personel bazında toplanır.

## Çıktı

`cikti/personel_satis_performansi.xlsx`:

- `Personel`: ciroya göre sıralı; pay, fiş, adet, sepet ortalaması, fiş başına adet ve mağazaya göre yüzdeleri.
- `Prim Tablosu` (hedef verilirse): gerçekleşme, prim oranı, prim tutarı ve onay sütunu.
- `Uyarılar`.

## Dikkat

- **Prim politikası sizindir.** Kademeler, hedefler ve primin cironun mu kârın mı yüzdesi olacağı şirket
  kararıdır. Kod yalnız verdiğiniz kademeleri uygular; ekip primi, ürün bazlı prim, iade düşümü ve tavan gibi
  kuralları uygulamaz.
- **Bordro hesabı değildir.** Prim tutarı brüttür; gelir vergisi, damga vergisi, SGK primi ve bordroya yansıtma
  ayrıca yapılmalıdır. Prim ödemesinin iş sözleşmesine ve şirket yönetmeliğine uygunluğunu insan kaynaklarıyla
  doğrulayın.
- **Satışın kime yazıldığı kayda bağlıdır.** Birden çok danışmanın ilgilendiği satışlar, kasada yanlış personel
  seçimi ve iade sonrası ciro düzeltmeleri sonucu değiştirir; ciro iadeler düşülmüş olarak verilmelidir.
- **Çalışma süresi farkı gözetilmez.** Yarı zamanlı ya da izinli personelin cirosu düşük görünür; günlük ortalama
  ciroya bakın ya da hedefleri çalışma süresine göre verin.
- **Kişisel veri içerir.** Rapor çalışan performansına ilişkindir; erişimi sınırlayın ve kişisel verilerin
  korunması kurallarına uyun.
- **Örnek veri:** Personel adları, satışlar ve hedefler kurgusaldır.

## Testler

- Ciro payı, sepet ortalaması, fiş başına adet ve mağazaya göre yüzdeler; çalışılan gün ve günlük ortalama.
- Hedef gerçekleşme, kademe sınırları (tam eşik üst kademeye girer), prim tutarı ve toplamı.
- Hedefi olmayan personel, satışı olmayan hedefli personel, kademe yazımı, okuma ve komut satırı.

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
