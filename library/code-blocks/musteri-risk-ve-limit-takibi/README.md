# Müşteri Risk ve Limit Takibi · Kod Bloğu

> Satış › Satış Destek Uzmanı · Workers / Workless

Sevkiyattan önce her müşterinin **açık bakiyesine, vadesi geçmiş alacağına ve kredi limitine** bakar; siparişi
"Sevk etme", "Onay gerekli" ya da "Uygun" olarak işaretler ve gerekçesini yazar. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Kural |
|---|---|
| Açık bakiye | Müşterinin açık kalemlerinin toplamı (eksi kalemler: avans, iade, alacak dekontu) |
| Vadesi geçmiş alacak | Vadesi durum tarihinden önce olan artı kalemlerin toplamı |
| En eski gecikme | Durum tarihi − en eski vadesi geçmiş kalemin vadesi |
| Sipariş sonrası bakiye | Açık bakiye + sevk bekleyen sipariş tutarı |
| Limit kullanımı % | Bakiye ÷ kredi limiti × 100 (sipariş öncesi ve sonrası ayrı) |
| Teminatsız risk | Açık bakiye − teminat (eksiye düşmez) |

| Karar | Koşul |
|---|---|
| Sevk etme | Sipariş sonrası bakiye limiti aşıyor **ya da** en eski gecikme `--gecikme` (30) günden uzun |
| Onay gerekli | Vadesi geçmiş alacak var, limit tanımlı değil, müşteri dosyada yok **ya da** sipariş sonrası kullanım `--esik` (%90) ve üstünde |
| Uygun | Hiçbiri yok |

**Örnek:** Limit 400.000, açık bakiye 340.000, bekleyen sipariş 80.000 → sipariş sonrası 420.000 (%105): limit
20.000 aşılıyor → **Sevk etme**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                            # örnek veri, durum tarihi 12.10.2026
python main.py --musteriler musteriler.xlsx --kalemler acik_kalemler.xlsx --siparisler siparisler.xlsx
python main.py --musteriler musteriler.xlsx --kalemler acik_kalemler.xlsx --esik 80 --gecikme 15
```

Zorunlu sütunlar: müşteri dosyasında Müşteri Kodu; açık kalemlerde ve siparişlerde Müşteri Kodu ve Tutar. Sipariş
dosyası verilmezse yalnız mevcut bakiyeye göre değerlendirme yapılır.

## Çıktı

`cikti/musteri_risk_limit_takibi.xlsx`:

- `Müşteriler`: önce sevk edilmeyecekler, sonra onay bekleyenler; limit, bakiye, kullanım, vadesi geçmiş, gerekçe.
- `Siparişler`: her sipariş için müşterinin kararı ve gerekçesi (sipariş dosyası verildiyse).
- `Uyarılar`: limit aşımı, uzun gecikme, vadesi geçmiş alacak, limite yaklaşan ve limiti tanımsız müşteriler.

## Dikkat

- **Karar kuralı bir çalışma varsayımıdır.** Hangi durumda sevkiyatın durdurulacağı şirketinizin kredi politikasına
  bağlıdır. Eşikleri politikanıza göre ayarlayın; son kararı yetkili kişi verir. Kod kimseyi "kötü müşteri" olarak
  sınıflamaz, yalnız sizin kurallarınıza göre işaret koyar.
- **Karar müşteri düzeyindedir.** Bir müşterinin birden çok siparişi varsa hepsi birlikte değerlendirilir; "ilk
  sipariş sığar, ikincisi sığmaz" ayrımı yapılmaz.
- **Bakiye açık kalem dosyanız kadar günceldir.** Henüz işlenmemiş tahsilat, portföydeki çek ve senetler ya da
  faturalanmamış irsaliyeler dosyada yoksa hesaba girmez. Vadesi gelmemiş çekleri risk sayıp saymamak sizin
  tercihinizdir; saymak istiyorsanız açık kalemlere ekleyin.
- **Teminat yalnız bilgi içindir.** Karara etki etmez; teminatın türü, geçerliliği ve paraya çevrilebilirliği
  değerlendirilmez.
- **Tutarlar tek para biriminde olmalı.** Kod kur çevirisi yapmaz.
- **Kapsam dışı:** müşteri kredi notu, ödeme alışkanlığı skoru, grup şirketleri için birleşik limit, alacak
  sigortası, şüpheli alacak karşılığı.
- **Ticari sır:** Müşteri bakiyeleri ve limitleri yer alır; raporu yalnız ilgili kişilerle paylaşın.
- **Örnek veri:** Müşteriler ve tutarlar kurgusaldır.

## Testler

- Karar sınırları: limit tam dolu ve 1 kuruş üstü, kullanım eşiği, gecikme 30 ve 31 gün, bugün vadeli kalem.
- Avans (eksi kalem), vadesiz kalem, limitsiz ve tanımsız müşteri, siparişsiz çalışma, teminatsız risk.
- Sıralama, sipariş kararları, toplamlar, mükerrer müşteri kodu, okuma ve komut satırı.

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
