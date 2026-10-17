# İş Programı Kritik Yol Analizi · Kod Bloğu

> Planlama ve Proje Kontrol › Planlama Mühendisi · Gerisi Bende

Faaliyet süreleri ve bağımlılıklardan **kritik yolu, bolluk sürelerini ve proje bitiş tarihini** hesaplar (kritik yol
yöntemi, CPM). İnternete bağlanmaz.

## Nasıl hesaplar?

| Adım | Formül |
|---|---|
| Erken başlama (ES) | Öncüllerin (erken bitiş + gecikme) en büyüğü; öncülü yoksa 0 |
| Erken bitiş (EF) | ES + süre |
| Geç bitiş (LF) | Ardılların (geç başlama − gecikme) en küçüğü; ardılı yoksa proje süresi |
| Geç başlama (LS) | LF − süre |
| Toplam bolluk | LS − ES (projeyi geciktirmeden kayabileceği süre) |
| Serbest bolluk | Ardılların (ES − gecikme) en küçüğü − EF (ardılı geciktirmeden kayabileceği süre) |
| Kritik faaliyet | Toplam bolluğu 0 |
| Proje süresi | En büyük EF |

- Bağımlılık türü **bitiş-başlangıç (FS)**. Gecikme `B+2` (B bittikten 2 gün sonra), bindirme `B-1` yazılır.
- **Tarihler:** `--baslangic` verilirse hesaplanır. `--calisma-gunu` 7 (her gün, varsayılan), 6 (pazar hariç) veya 5
  (cumartesi-pazar hariç). Hafta sonuna denk gelen başlangıç ilk çalışma gününe kayar.
- Döngü (A → B → A), listede olmayan öncül ve mükerrer faaliyet kodu hata verir.

**Örnek:** A (5) → B (10) → C (20); C'den sonra D (7), E (12) ve G (6); F (10) D ve E'yi, H (5) F ve G'yi bekler → proje
62 gün; kritik yol A → B → C → E → F → H; D'nin bolluğu 5, G'nin 16 gün.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri (başlangıç 12.10.2026)
python main.py --faaliyetler faaliyetler.xlsx --baslangic 01.03.2027 --calisma-gunu 6
```

Sütunlar: Faaliyet Kodu, Süre (gün) (zorunlu); Faaliyet, Öncüller (virgülle ayrılmış faaliyet kodları).

## Çıktı

`cikti/kritik_yol.xlsx`:

- `Faaliyetler`: erken ve geç zamanlar, toplam ve serbest bolluk, kritik işareti, başlangıç ve bitiş tarihleri.
- `Kritik Yol`: kritik yollar (birden çoksa ilk 10).
- `Uyarılar`: birden çok kritik yol, öncülü veya ardılı olmayan birden çok faaliyet.

## Dikkat

- **Yalnız FS bağımlılığı.** Başlangıç-başlangıç (SS), bitiş-bitiş (FF) ve başlangıç-bitiş (SF) ilişkileri, zorunlu
  tarih kısıtları ve faaliyete özel takvimler desteklenmez.
- **Resmi tatiller ve hava koşulları yoktur.** Takvim yalnız haftalık çalışma gününü dikkate alır; bayram, resmi tatil
  ve çalışılamayan günleri süreye ekleyin veya bitiş tarihini kendiniz düzeltin.
- **Kaynak kısıtı yoktur.** Aynı ekibin veya makinenin iki faaliyette birden çalışamayacağı hesaba katılmaz; kaynak
  dengelemesi yapılmaz.
- **Süreler tahmindir.** Kritik yol, sürelerin doğruluğu kadar doğrudur; bolluğu az olan faaliyetler de küçük bir
  gecikmeyle kritik hale gelir.
- **Sözleşme programının yerini tutmaz.** İdareye veya işverene sunulan iş programı için planlama yazılımınızın
  çıktısı esastır; bu paket hızlı kontrol ve taslak içindir.
- **Örnek veri:** Faaliyetler ve süreler kurgusaldır.

## Testler

- İleri ve geri geçiş, toplam ve serbest bolluk, kritik yol.
- Gecikme ve bindirme, paralel kritik yollar, sıfır süreli kilometre taşı.
- Tarih hesabı (5, 6 ve 7 günlük hafta), hafta sonuna denk gelen başlangıç.
- Döngü, olmayan öncül, mükerrer kod ve okuma hataları; örnek veri uçtan uca.

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
