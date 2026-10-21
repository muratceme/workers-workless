# Uygulama Yönetim Özeti · Kod Bloğu

> Sizden Gelenler › Abonelikli mobil uygulama · Gerisi Bende

Diğer Sizden Gelenler paketlerinin ürettiği özet dosyalarını birleştirip **günlük, haftalık ya da aylık** kısa bir
yönetim raporu çıkarır: gelir, kullanıcı, çökme, yapay zekâ maliyeti, destek ve mağaza yorumları tek sayfada, önceki
dönemle karşılaştırmalı. Rapor hem **metin** (e-postaya ya da Slack'e yapıştırılabilir) hem **Excel** olarak yazılır.
İnternete bağlanmaz.

## Hangi dosyaları okur?

Her paket çalıştığında çıktı klasörüne `<paket>.ozet.json` yazar. Bu paket verdiğiniz klasörleri (alt klasörler
dahil) tarar ve bulduklarını birleştirir; olmayan bölüm rapora girmez.

| Paket | Rapordaki bölüm |
|---|---|
| Uygulama Abonelik Gelir Raporu | Gelir: net gelir, Play brüt satış, satış ve iade adedi, komisyon, iptal |
| Uygulama Kullanıcı ve Tutma Analizi | Kullanıcı: günlük aktif (ortalama), yeni kullanıcı, oturum; D1 / D7 / D30 tutma |
| Uygulama Çökme Önceliklendirme | Çökme: FATAL, ANR, ölümcül olmayan hata; P1 sorunlar |
| Yapay Zekâ API Maliyet Raporu | Yapay zekâ: maliyet, istek, ay sonu tahmini |
| Uygulama Destek E-postası Sınıflandırma | Destek: e-posta, P1, insan onayı gereken |
| Mağaza Yorumu Sınıflandırma | Mağaza: yorum, ortalama puan, 1–2 yıldız, en çok istenen özellikler |

En kolayı, bütün paketlerin `--cikti` yolunu aynı klasöre vermektir (ör. `ozetler/gelir.xlsx`, `ozetler/cokme.xlsx`…).
Aynı paketin birden çok özeti bulunursa en yeni üretilen kullanılır.

## Nasıl hesaplar?

| Dönem | Bu dönem | Önceki dönem |
|---|---|---|
| `gunluk` | Bitiş günü | Bir önceki gün |
| `haftalik` | Son 7 gün | Ondan önceki 7 gün |
| `aylik` | Ay başından bitiş gününe | Önceki ayın aynı günleri |

- **Bitiş günü** `--tarih` ile verilir; verilmezse dosyalardaki en son gündür.
- Gösterge, dönemdeki günlerin **toplamıdır**; günlük aktif kullanıcı **günlük ortalama**, mağaza puanı yorum sayısıyla
  **ağırlıklı ortalama**dır.
- **Birleşik göstergeler:** yapay zekâ maliyeti ÷ aktif kullanıcı-gün (iki dosyada da olan günler), yapay zekâ
  maliyeti ÷ net gelir (para birimleri aynıysa ya da maliyet raporu `--kur` ile çalıştırıldıysa).
- **Kapsam:** Dönemin kaç gününde o kaynağın verisi olduğu yazılır. Eksik gün varsa toplamlar eksiktir.
- **Dikkat listesi:** önceki döneme göre kötü yönde en az %20 değişim (`--esik`; yalnız iki dönem de tam kapsanıyorsa),
  P1 çökme, ay sonu yapay zekâ tahmininin bütçeyi aşması, dönem sonuna kadar uzanmayan (eski) veri.

## Kullanım

```bash
python main.py                                                 # örnek özetlerle dener (kurgusal, Eylül 2026)
python main.py --kaynak ozetler --donem gunluk
python main.py --kaynak ozetler --donem haftalik --tarih 27.09.2026
python main.py --kaynak ozetler --donem aylik --esik 15
```

Terminal kullanmıyorsanız klasördeki `BASLAT.bat` (Windows) ya da `BASLAT.command` (macOS) dosyasına çift tıklayın.

## Her sabah 09.00'da çalıştırmak

Bilgisayarınızın zamanlayıcısıyla raporu her gün aynı saatte üretebilirsiniz. Paket zamanlamayı kendisi kurmaz;
aşağıdaki komutu yolları kendinize göre düzenleyip **siz** çalıştırırsınız.

Windows (Görev Zamanlayıcı):

```bat
schtasks /Create /SC DAILY /ST 09:00 /TN "Gerisi Bende - Gunluk Ozet" /TR "C:\Klasor\uygulama-yonetim-ozeti\gunluk.bat"
```

`gunluk.bat` içeriği:

```bat
cd /d C:\Klasor\uygulama-yonetim-ozeti
python main.py --kaynak C:\Klasor\ozetler --donem gunluk
```

macOS / Linux (`crontab -e`):

```
0 9 * * * cd /Users/ad/uygulama-yonetim-ozeti && python3 main.py --kaynak /Users/ad/ozetler --donem gunluk
```

Haftalık rapor için Pazartesi (`/SC WEEKLY /D MON`, cron'da `0 9 * * 1`), aylık rapor için ayın 1'i (`/SC MONTHLY /D 1`,
cron'da `0 9 1 * *`) ve `--tarih` olmadan çalıştırın. Ayın 1'inde önceki ayın raporunu almak için diğer paketlerin
verisi önceki ayın son gününde bitmelidir.

**Önemli:** Bu rapor, kaynak paketler o sabah yeni veriyle çalıştırıldıysa günceldir. Mağaza, Firebase, Crashlytics
ve e-posta verisinin her gün otomatik indirilmesi, paketlerin sırayla çalıştırılması ve raporun e-posta / Slack ile
gönderilmesi n8n / Make gibi bir otomasyon aracıyla kurulur; bu **kurulum hizmeti** kapsamındadır. Rapordaki
"eski veri" uyarısı, güncellenmemiş kaynağı gösterir.

## Çıktı

- `cikti/yonetim_ozeti.md`: kısa rapor metni (Dikkat, bölümler, son durum)
- `cikti/yonetim_ozeti.xlsx`: Özet (kötü yöndeki değişim kırmızı), Dikkat, Günlük (bu ve önceki dönemin günleri),
  Kaynaklar (hangi dosya, hangi tarih aralığı), Uyarılar

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
