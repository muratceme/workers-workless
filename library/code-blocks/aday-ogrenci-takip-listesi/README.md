# Aday Öğrenci Takip Listesi · Kod Bloğu

> Kayıt Kabul ve Öğrenci İşleri › Kayıt Kabul Sorumlusu · Workers / Workless

Kayıt görüşmesi yapılan aday öğrencileri **aşama, kademe, kaynak ve takip tarihine** göre izler. Takibi gelmiş ya da unutulmuş
adayları listeler; aşama hunisini, kademe ve kaynak bazında **dönüşüm oranlarını** ve kayıp nedenlerini çıkarır. İnternete ve
okul yönetim yazılımına bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Kural |
|---|---|
| Aşamalar | Ön görüşme → Okul gezisi → Değerlendirme / tanışma → Ücret görüşmesi → Kayıt; Kayıp ayrı durum |
| Huni | Her aşamaya ulaşan aday; kayıp adaylar "Kayıp Aşaması" yazılıysa o aşamaya kadar, yazılı değilse yalnız ön görüşmede sayılır |
| Geçiş % | Aşamaya ulaşan ÷ önceki aşamaya ulaşan × 100 |
| Dönüşüm % | Kayıt ÷ aday × 100 |
| Kapanış % | Kayıt ÷ (kayıt + kayıp) × 100 (açık adaylar girmez) |
| Ortalama kayıt süresi | İlk temas → kayıt tarihi (gün) |
| Takip listesi | Açık adaylardan sonraki takip tarihi bugün ya da geçmiş olanlar ve takip tarihi yazılmamış olup son temastan `--takip` (7) gün geçenler |

Aşama metinden tanınır: Ön görüşme / Bilgi alma (boş = Ön görüşme) · Okul gezisi / ziyaret / tanıtım · Değerlendirme / tanışma /
sınav / mülakat / gözlem · Ücret / teklif / indirim / burs görüşmesi · Kayıt / Kesin kayıt · Kayıp / Vazgeçti / Başka okul.
Takip listesi ileri aşamadaki adaylar üstte olacak şekilde sıralanır.

**Uyarılar:** takip 3 günden fazla gecikmiş (Orta) · okunamayan satır (Yüksek) · aynı aday iki kez (Orta) · açık adayın sorumlusu
yok · kayıp nedeni yazılmamış (Bilgi).

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                      # örnek veri (14.10.2026)
python main.py --adaylar adaylar.xlsx
python main.py --adaylar adaylar.xlsx --tarih 20.10.2026 --takip 5
```

Zorunlu sütunlar: Öğrenci ve Aşama.

## Çıktı

`cikti/aday_ogrenci_takip.xlsx`: `Takip Listesi` (geciken takipler sarı), `Huni`, `Kademe`, `Kaynak`, `Sorumlu`, `Kayıp Nedenleri`,
`Uyarılar`.

## Dikkat

- **Çocuklara ve velilere ait kişisel veri içerir.** Öğrenci ve veli bilgileri KVKK kapsamındadır; aday velilerden aydınlatma ve
  gerekiyorsa açık rıza alınmalı, liste yalnız yetkili personelle paylaşılmalıdır. Kayıt olmayan adayların verilerini saklama
  süreniz dolunca silin. Ticari elektronik ileti (SMS / e-posta kampanyası) için ayrıca onay ve İYS kaydı gerekir.
- **Aşama adları okulunuza göre değişebilir.** Kendi süreç adlarınızı yukarıdaki aşamalardan birine karşılık gelecek şekilde yazın.
- **Kaynak ve dönüşüm karşılaştırmaları** az adayla yanıltıcı olabilir; dönem ve kademe bazında birlikte değerlendirin.
- **Örnek veri:** Adaylar, veliler ve telefonlar kurgusaldır.

## Testler

- Huni (kayıp aşaması dahil), geçiş %, dönüşüm ve kapanış oranları, ortalama kayıt süresi, kademe / kaynak / sorumlu özetleri.
- Takip listesi kuralları ve sıralama, uyarılar, aşama tanıma, okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

## Kaynaklar ve dayanaklar

Kontrol tarihi: 15.10.2026. Mevzuat sık değişir; paketteki süre, oran ve kuralları kullanmadan önce güncel resmî metinle
karşılaştırın. "Karşılaştırılan maddeler" dışındaki atıflar yalnız yol göstericidir.

| Kaynak | Bağlantı | Durum |
|---|---|---|
| 6698 sayılı Kişisel Verilerin Korunması Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.6698.pdf) | Resmî metinle karşılaştırılan maddeler: 13 |

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
