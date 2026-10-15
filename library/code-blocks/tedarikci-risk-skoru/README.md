# Tedarikçi Risk Skoru · Kod Bloğu

> Satın Alma › Satın Alma Müdürü · Workers / Workless

Tedarikçilerinizi dört ölçüte göre **0–100 arası tek bir risk skoruna** indirger ve sıralar: tek kaynak
bağımlılığı, finansal durum, teslimat gecikmesi, kalite. Hangi tedarikçi için ikinci kaynak aranmalı, hangisiyle
iyileştirme planı yapılmalı sorusuna aynı ölçüyle bakmanızı sağlar. İnternete bağlanmaz.

## Nasıl hesaplar?

Her ölçüt 0 (risksiz) ile 100 (en riskli) arasında puanlanır:

| Ölçüt | Puan | Ağırlık |
|---|---|---|
| Tek kaynak | Tek kaynak kalem sayısı ÷ kalem sayısı × 100 | 30 |
| Finansal | İyi 0 · Orta 50 · Zayıf 100 | 25 |
| Gecikme | Gecikmeli teslimat % ÷ 20, en çok 100 | 25 |
| Kalite | Ret % ÷ 5, en çok 100 | 20 |

**Risk skoru** bu puanların ağırlıklı ortalamasıdır. Verisi olmayan ölçüt hesaptan çıkarılır ve kalan ağırlıklar
yeniden oranlanır; ikiden az ölçütü olan tedarikçi "Değerlendirilemedi" sayılır. Sınıf: 60 ve üstü **Yüksek**,
30 ve üstü **Orta**, altı **Düşük**.

**Örnek:** 8 kalemin 4'ünde tek kaynak (50), finansal durumu zayıf (100), gecikme %15 (75), ret %4 (80) →
(30 × 50 + 25 × 100 + 25 × 75 + 20 × 80) ÷ 100 = **74,75 → Yüksek**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                          # örnek veri
python main.py --tedarikciler tedarikciler.xlsx
python main.py --tedarikciler tedarikciler.xlsx --agirlik "tek=40;finans=20;gecikme=20;kalite=20" --gecikme-tavan 15 --yuksek 50
```

Zorunlu sütun: Tedarikçi. Diğerleri: Yıllık Alım Tutarı, Kalem Sayısı, Tek Kaynak Kalem, Gecikmeli Teslimat %,
Ret %, Finansal Durum (İyi / Orta / Zayıf).

| Seçenek | Anlamı | Varsayılan |
|---|---|---|
| `--agirlik` | Ölçüt ağırlıkları | tek=30;finans=25;gecikme=25;kalite=20 |
| `--gecikme-tavan`, `--ret-tavan` | Puanın 100 olduğu oranlar, % | 20, 5 |
| `--yuksek`, `--orta` | Sınıf eşikleri | 60, 30 |
| `--pay` | Yüksek riskli tedarikçi bu harcama payına ulaşıyorsa "kritik" | 10 |

## Çıktı

`cikti/tedarikci_risk_skoru.xlsx`:

- `Risk Skoru`: sınıf, skor, dört ölçüt puanı, yıllık alım, harcama payı ve ham veriler. En riskli en üsttedir.
- `Uyarılar`: kritik tedarikçiler (yüksek risk ve yüksek harcama payı), tek kaynak olup finansal durumu zayıf
  olanlar, harcama yoğunlaşması, eksik veya geçersiz veri.

## Dikkat

- **Skor bir önceliklendirme aracıdır.** Ağırlıklar, tavanlar ve eşikler yönetim tercihidir; nesnel bir kredi notu
  ya da denetim sonucu değildir. Sektörünüze göre değiştirin ve herkesin aynı ayarla çalıştığından emin olun.
- **Finansal durumu siz değerlendirirsiniz.** Kod bilanço okumaz, kredi veya sicil sorgusu yapmaz; "İyi / Orta /
  Zayıf" sizin kanaatinizdir. Bu kanaati belgeye (mali tablo, ödeme davranışı, bağımsız rapor) dayandırın.
- **Veri kalitesi sonucu belirler.** Gecikme ve ret oranlarını aynı dönem ve aynı tanımla (satır mı, sipariş mi,
  miktar mı) hesaplayın. Az teslimatı olan tedarikçide tek bir gecikme oranı çok oynatır.
- **Kapsam dışı:** ülke ve kur riski, sözleşme ve uyum riski, çevresel ve sosyal ölçütler, tedarikçinin
  tedarikçileri, stok ve ikame süresi.
- **Paylaşım:** Tedarikçi hakkındaki değerlendirmeler ticari sır niteliğindedir; raporu şirket dışına vermeyin.
- **Örnek veri:** Tedarikçiler ve değerler kurgusaldır.

## Testler

- Dört ölçüt puanı, tavan sınırı, ağırlıklı skor ve sınıf eşikleri (elle hesaplanmış).
- Eksik ölçütte ağırlıkların yeniden oranlanması, değerlendirilemeyen tedarikçi, özel ağırlıklar.
- Harcama payı, kritik tedarikçi ve yoğunlaşma uyarıları; okuma hataları ve komut satırı.

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
