# Yapılandırma Takip Listesi · Kod Bloğu

> Krediler İzleme ve Takip › Kredi İzleme Uzmanı · Gerisi Bende

Yeniden yapılandırılan kredilerin **yeni ödeme planına uyup uymadığını ve izleme süresinin ne zaman dolacağını**
tek listede izler; planı aksatan ve izleme süresi dolan kredileri öne çıkarır. İnternete bağlanmaz.

## Nasıl hesaplar?

| Kavram | Kural |
|---|---|
| Ödeme uyumu % | Ödenen tutar ÷ vadesi gelen tutar × 100 (vadesi gelen taksit yoksa hesaplanmaz) |
| Eksik ödeme | Vadesi gelen tutar − ödenen tutar (eksiye düşmez) |
| İzleme bitişi | Yapılandırma tarihi + izleme süresi (ay); satırda yoksa `--izleme-ay` (12) |
| Uyumsuz | Gecikme günü `--gecikme` (30) günü aşmış |
| Gecikmeli | Gecikme var ya da eksik ödeme var |
| İzleme süresi doldu | Gecikme ve eksik ödeme yok, izleme bitişi geçmiş |
| Uyumlu | Diğerleri |

**Örnek:** Vadesi gelen 400.000, ödenen 250.000 → uyum **%62,50**, eksik ödeme 150.000; gecikme 45 gün > 30 →
**Uyumsuz**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                            # örnek veri, durum tarihi 12.10.2026
python main.py --krediler krediler.xlsx
python main.py --krediler krediler.xlsx --gecikme 30 --izleme-ay 12 --tarih 31.10.2026
```

Zorunlu sütunlar: Müşteri No, Yapılandırma Tarihi. Boş bırakılan tutarlar sıfır, gecikme günü sıfır, yapılandırma
sayısı 1 sayılır.

## Çıktı

`cikti/yapilandirma_takip_listesi.xlsx`:

- `Yapılandırmalar`: önce uyumsuzlar; durum, kalan risk, vadesi gelen, ödenen, eksik ödeme, uyum, gecikme, izleme
  bitişi ve kalan gün.
- `Özet`: durum bazında kredi sayısı ve kalan risk; portföy geneli ödeme uyumu.
- `Uyarılar`: uyumsuz ve gecikmeli krediler, birden çok kez yapılandırılanlar, izleme süresi dolanlar.

## Dikkat

- **Bu bir izleme listesidir, sınıflandırma kararı değildir.** Yapılandırılan kredinin hangi grupta izleneceği,
  izleme süresinin uzunluğu, hangi gecikmenin yeniden donuk alacak sınıflamasını gerektirdiği ve karşılık tutarı
  kredilerin sınıflandırılmasına ilişkin düzenlemelere göre belirlenir. Varsayılan 30 gün ve 12 ay çalışma
  varsayımıdır; güncel düzenlemeyle ve iç politikanızla doğrulayın.
- **"İzleme süresi doldu" bir hatırlatmadır.** Sınıf değişikliği için süre dışında başka koşullar da aranır (belirli
  oranda anapara ödemesi, finansal durumun düzelmesi gibi); kod yalnız sürenin geçtiğini ve gecikme olmadığını söyler.
- **Vadesi gelen ve ödenen tutar sizin verinizdir.** Kod ödeme planını taksit taksit izlemez; yapılandırmadan bugüne
  toplam tahakkuk ve toplam tahsilat üzerinden çalışır. Fazla ödeme uyumu %100'ün üstüne çıkarır.
- **Tutarlar tek para biriminde olmalı.**
- **Kapsam dışı:** yapılandırma koşullarının (vade, faiz, ödemesiz dönem) değerlendirmesi, teminat durumu, nakit
  akışı analizi, karşılık hesabı.
- **Müşteri sırrı:** Rapor müşteri risk bilgisi içerir; yalnız yetkili kişilerle paylaşın.
- **Örnek veri:** Müşteriler, krediler ve tutarlar kurgusaldır.

## Testler

- Durum sınırları: gecikme 0 / 1 / 30 / 31 gün, eksik ödeme, izleme bitişi bugün / yarın, ay ekleme.
- Uyum yüzdesi (vadesi gelen sıfır, fazla ödeme), satırda ve komut satırında izleme süresi, çoklu yapılandırma.
- Durum özeti ve paylar, tutarsız veri uyarıları, sıralama, okuma ve komut satırı.

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
