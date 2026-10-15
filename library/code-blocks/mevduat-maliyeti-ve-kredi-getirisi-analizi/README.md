# Mevduat Maliyeti ve Kredi Getirisi Analizi · Kod Bloğu

> Hazine ve Aktif-Pasif Yönetimi › Aktif-Pasif Yönetimi Uzmanı · Workers / Workless

Ürün bakiyeleri ve faiz oranlarından **ağırlıklı ortalama mevduat maliyetini, kredi getirisini ve aradaki marjı**
para birimi bazında hesaplar; hangi ürünün ortalamayı ne kadar yukarı ya da aşağı çektiğini gösterir. İnternete
bağlanmaz.

## Nasıl hesaplar?

| Kavram | Formül |
|---|---|
| Yıllık faiz tutarı | Bakiye × faiz oranı ÷ 100 |
| Ağırlıklı ortalama % | Σ (bakiye × faiz oranı) ÷ Σ bakiye (tür ve para birimi bazında) |
| Vadesiz hariç maliyet | Faiz oranı sıfırdan büyük mevduat ürünleri üzerinden aynı hesap |
| Ürünün katkısı (puan) | Ürünün gruptaki payı × faiz oranı; katkıların toplamı ağırlıklı ortalamayı verir |
| Marj (puan) | Kredi getirisi − mevduat maliyeti (aynı para biriminde) |
| Kredi / mevduat % | Kredi bakiyesi ÷ mevduat bakiyesi × 100 |
| Net faiz (basit) | Kredilerin yıllık faiz tutarı − mevduatın yıllık faiz tutarı |

**Örnek:** TL mevduat 2 milyar: vadesiz 400 milyon (%0), vadeli 600 milyon (%42), 800 milyon (%45), 200 milyon (%40)
→ yıllık faiz 692 milyon → maliyet **%34,60**; vadesiz hariç **%43,25**. Vadesiz mevduatın payı düşerse maliyetin
nereye gideceğini bu iki sayının farkı gösterir.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                            # örnek veri
python main.py --urunler urunler.xlsx
```

Zorunlu sütunlar: Ürün, Tür, Bakiye, Faiz Oranı. Para birimi boşsa TRY sayılır. Vadesiz mevduat için faiz oranını 0
yazın. Oranlarda virgül ondalık ayırıcıdır (`3,5` = %3,5).

## Çıktı

`cikti/mevduat_maliyeti_kredi_getirisi.xlsx`:

- `Marj`: para birimi bazında bakiyeler, kredi / mevduat oranı, maliyet, getiri, marj ve basit net faiz; altında tür
  bazında ağırlıklı ortalamalar.
- `Ürünler`: her ürünün gruptaki payı, faiz oranı, ortalamaya katkısı ve yıllık faiz tutarı.
- `Uyarılar`: eksi marj, ortalama mevduat maliyetinin altında getirili kredi, ortalama kredi getirisinin üstünde
  maliyetli mevduat.

## Dikkat

- **Bu basit bir marj görünümüdür; net faiz marjı raporu değildir.** Zorunlu karşılık maliyeti, mevduattaki vergi
  ve fon kesintileri, kredilerdeki kesintiler, bileşik faiz, vade ve yeniden fiyatlama farkı, operasyonel maliyet
  ve beklenen kredi zararı hesaba katılmaz. Ürün fiyatlamasında fon transfer fiyatlaması kullanıyorsanız karar için
  o esas alınmalıdır.
- **Oranlar aynı tanımda olmalı.** Basit yıllık oran mı, bileşik mi, brüt mü; bütün ürünlerde aynı olmalıdır. Aylık
  oranları yıllığa çevirip girin.
- **Bakiye an değeri mi ortalama mı?** Dönem sonu bakiyesi yerine ortalama bakiye kullanmak marjı daha doğru
  gösterir; hangisini kullandıysanız bütün satırlarda aynı olmalıdır.
- **Para birimleri birleştirilmez.** Her para birimi kendi içinde değerlendirilir; TL mevduatla fonlanan döviz kredisi
  gibi çapraz fonlama ve swap maliyeti bu analizde görünmez.
- **Stok oranları geçmişe bakar.** Portföydeki ortalama oran, bugün yeni açılan mevduat ve kredilerin oranından
  farklıdır; yeni işlem marjı için yalnız dönem içi açılışları ayrı çalıştırın.
- **Örnek veri:** Ürünler, bakiyeler ve oranlar kurgusaldır; piyasa oranı değildir.

## Testler

- Ağırlıklı ortalama, vadesiz hariç maliyet, pay ve katkı (katkıların toplamı = ortalama), yıllık faiz tutarı.
- Marj, kredi / mevduat oranı ve net faiz; eksi marj; yalnız mevduatı ya da yalnız kredisi olan para birimi.
- Ürün düzeyi uyarılar, para birimi ayrımı ve sırası, okuma (oran yazımları) ve komut satırı.

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
