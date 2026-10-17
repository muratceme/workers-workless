# Koli Listesi (Packing List) Hazırlama · Kod Bloğu

> Tekstil ve Konfeksiyon › Ütü-Paket ve Sevkiyat › Sevkiyat Sorumlusu · Gerisi Bende

Sevk edilecek adetlerden (sipariş × model × renk × beden) **koli planı** ve ihracata uygun, Türkçe-İngilizce
**packing list** hazırlar. Her koli için etiket listesi de üretir. İnternete bağlanmaz.

## Nasıl kolilenir?

| Yöntem | Kural |
|---|---|
| **Asorti** (`--yontem asorti --asorti S=1,M=2,L=2,XL=1`) | Her koliye sabit beden oranı; koli içi 24 ise set ×4 (S4 M8 L8 XL4) |
| **Solid** (varsayılan) | Her koliye tek renk-tek beden, koli içi adet kadar |
| **Kalanlar** | Asortiye girmeyenler solid; koli içini doldurmayan artıklar karışık koliye (veya eksik solid koli) |

**Hesaplananlar**
- **Koli numaraları:** Sürekli verilir (1–75, 76, 77–126 …).
- **Adetler:** Her satırda koli adedi, koli içi adet ve toplam adet bulunur.
- **Net ağırlık:** Adet ağırlığı × adet. Adet ağırlığı sevk dosyasındaki model ağırlığından ya da `--agirlik`
  değerinden alınır.
- **Brüt ağırlık:** Net ağırlık + koli darası (`--dara`).
- **Hacim (CBM):** Koli ölçüsünden (`--koli-olcu 60x40x40`) hesaplanır.

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                                              # örnek: 3 model-renk, asorti
python main.py --sevk sevk.xlsx --yontem asorti --asorti S=1,M=2,L=2,XL=1 --koli-ici 24
python main.py --sevk sevk.xlsx --yontem solid --koli-ici 30 --koli-olcu 60x40x30 --dara 0,7 \
               --gonderici "Örnek Tekstil A.Ş." --alici "Example Retail Ltd."
```

## Çıktı

`Packing List` (koli no, koli adedi, sipariş, model, renk, tür, beden dağılımı, koli içi, toplam, net/brüt kg,
CBM ve toplamlar) · `Özet` (sipariş-model-renk bazında) · `Koli Etiketleri` (her koli için bir satır)

## Dikkat

- **Ağırlıklar:** Gerçek ağırlıklar ilk kolilerin tartımıyla doğrulanmalıdır. Gümrük beyannamesi ve faturada
  aynı değerler kullanılmalıdır.
- **Müşteri kuralları:** Müşterinin paketleme talimatı (koli içi, asorti oranı, etiket bilgisi) bu ayarlardan
  önce gelir.

## Testler

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
