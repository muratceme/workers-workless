# Kargo Teslim Performansı · Kod Bloğu

> E-ticaret Operasyon ve Kargo › E-ticaret Operasyon Uzmanı · Gerisi Bende

Gönderi listesinden **kargo firması (ve il) bazında** ortalama ve medyan teslim süresini, **zamanında teslim oranını**,
hasarlı, kayıp ve iade gönderi oranlarını hesaplar; geciken açık gönderileri listeler. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Teslim süresi | Teslim tarihi − kargoya veriliş tarihi (takvim günü) |
| Ortalama / medyan süre | Teslim edilen (hasarlılar dahil) gönderilerin süreleri üzerinden |
| Zamanında teslim % | Süresi `--hedef` (varsayılan 3) gün ve altı olan ÷ teslim edilen × 100 |
| Hasar / kayıp / iade % | İlgili gönderi ÷ toplam gönderi × 100 |
| Gecikmiş açık gönderi | Teslim edilmemiş, kayıp ya da iade olmamış ve kargoya verilişinden `--gecikme` (varsayılan 5) günden fazla geçmiş |

Durum sütunu boşsa teslim tarihi olan gönderi "Teslim edildi", olmayan "Yolda" sayılır. Tanınan durumlar: Teslim
edildi, Hasarlı, Kayıp, İade, Yolda / Dağıtımda / Transferde / Şubede / Teslim edilemedi.

**Uyarılar:** zamanında teslim `--zamaninda` (varsayılan %90) altında; hasar oranı `--hasar` (varsayılan %1) üstünde;
kayıp oranı `--kayip` (varsayılan %0,5) üstünde; her gecikmiş açık gönderi.

**Örnek:** Kargo B 49 gönderiyi teslim etmiş, bunların 31'i 3 gün içinde → zamanında teslim **%63,27**; 56 gönderide 1
kayıp → kayıp oranı **%1,79** (eşik %0,5 → uyarı).

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                                 # örnek veri (30.09.2026)
python main.py --gonderiler kargo_gonderileri.xlsx --tarih 30.09.2026
python main.py --gonderiler kargo_gonderileri.xlsx --hedef 2 --gecikme 4 --zamaninda 95 --hasar 0,5 --kayip 0,2
```

Zorunlu sütunlar: Kargo Firması ve Kargoya Veriliş Tarihi. İl sütunu varsa firma × il kırılımı da çıkar.

## Çıktı

`cikti/kargo_teslim_performansi.xlsx`:

- `Firmalar`: firma bazında göstergeler ve toplam; eşiği aşanlar renkli.
- `Firma ve İl`: il sütunu varsa firma × il göstergeleri.
- `Gecikmiş Gönderiler`: en eskiden yeniye.
- `Uyarılar`.

## Dikkat

- **Takvim günü kullanılır.** Hafta sonu, resmî tatil ve kargo firmasının çalışmadığı günler süreye dahildir; iş günü
  taahhüdüyle karşılaştırırken bunu gözetin. Saat bilgisi kullanılmaz.
- **Kargoya veriliş tarihi** firmanın gönderiyi teslim aldığı (ilk okutma) tarih olmalıdır; sipariş tarihi
  yazılırsa depo süresi de kargo süresine eklenir.
- **Durum kayıtları firmadan firmaya farklıdır.** Kendi durum adlarınızı tanınan adlara çevirin; tanınmayan durum
  satırı alınmaz ve uyarı verilir.
- **Eşikler politika varsayılanıdır;** kargo sözleşmenizdeki hizmet seviyesi taahhütlerine göre değiştirin. Hasar ve
  kayıp tazmini için sözleşme ve tutanak süreçleri ayrıca yürütülür.
- **Örnek veri:** Firma adları, gönderi numaraları ve tarihler kurgusaldır.

## Testler

- Teslim süresi, ortalama, medyan, en uzun süre, zamanında teslim, hasar, kayıp ve iade oranı, firma × il.
- Gecikmiş açık gönderi, eşikler, durum çevirisi, hatalı satırlar (teslim verilişten önce, yinelenen gönderi), okuma
  ve komut satırı.

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
