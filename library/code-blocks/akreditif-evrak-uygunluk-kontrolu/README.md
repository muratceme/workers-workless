# Akreditif Evrak Uygunluk Kontrolü · Kod Bloğu

> Dış Ticaret › Dış Ticaret Uzmanı · Workers / Workless

Akreditif şartlarını, ibraz edeceğiniz evraklardaki bilgilerle karşılaştırır ve **olası rezervleri ibrazdan önce**
listeler. Kontroller UCP 600 maddelerine dayanır; her bulguda ilgili madde yazar. İnternete bağlanmaz.

## Nasıl hesaplar?

| Kontrol | UCP 600 |
|---|---|
| İbraz, akreditif vadesinden sonra | md. 6(d) |
| İbraz, yüklemeden sonraki 21 günü (akreditif başka süre verdiyse onu) aşıyor | md. 14(c) |
| Vade veya ibraz için son gün hafta sonuna geliyorsa ilk iş gününe uzar | md. 29(a) |
| Yükleme, son yükleme tarihinden sonra (tatil nedeniyle uzamaz) | md. 29(c) |
| İstenen evrak yok; asıl / kopya sayısı eksik | md. 14, 17 |
| Evrak tarihi ibraz tarihinden sonra | md. 14(i) |
| Fatura: lehtar düzenlememiş, amir adına değil, para birimi farklı | md. 18(a) |
| Fatura tutarı akreditif tutarını (toleransla) aşıyor | md. 18(b), 30 |
| Faturadaki mal tanımı akreditiftekiyle uyuşmuyor | md. 18(c) |
| Miktar toleransı: "about / yaklaşık" ±%10; koli veya adet olarak belirtilmeyen miktarda ±%5; adet / kolide tolerans yok | md. 30 |
| Kısmi yükleme yasakken miktar eksik | md. 31 |
| Konşimento: yükleme (on board) kaydı yok, limanlar farklı, tam takım asıl ibraz edilmiyor | md. 20(a) |
| Taşıma belgesinde malın / ambalajın kusuruna ilişkin şerh var | md. 27 |
| Sigorta: para birimi farklı, tutar CIF / CIP değerinin %110'undan az, tarih yüklemeden sonra | md. 28 |
| Koli, ağırlık, miktar, alıcı gibi bilgiler evraklar arasında çelişiyor | md. 14(d) |

- Mal tanımı karşılaştırması büyük-küçük harf ve noktalama farkını yok sayar; kelime farkını rezerv riski sayar.
- Lehtar ve amir adlarında, akreditifteki adın evrakta geçmesi yeterlidir (adres eki sorun sayılmaz).
- Sigorta tarihi yüklemeden sonraysa, "Teminat Başlangıcı" alanı yükleme tarihinden önce girilmişse bulgu yazılmaz.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek veri
python main.py --akreditif akreditif.csv --istenen istenen_evraklar.csv --evraklar evraklar.csv --ibraz 02.11.2026
```

| Dosya | Biçim |
|---|---|
| Akreditif | `Alan;Değer` satırları: Akreditif No, Para Birimi, Tutar, Tutar Toleransı (% veya "about"), Miktar, Miktar Birimi, Miktar Toleransı, Son Yükleme Tarihi, Vade, İbraz Süresi (gün), Kısmi Yükleme (İzinli / Yasak), Lehtar, Amir, Mal Tanımı, Teslim Şekli, Yükleme Limanı, Boşaltma Limanı, Sigorta Oranı, Menşe |
| İstenen evraklar | Evrak, Asıl, Kopya, Şart (akreditifin 46A alanındaki özel şart; rapora elle kontrol için aktarılır) |
| Evraklar | Evrak, Alan, Değer. Alanlar: Tarih, Asıl Adet, Kopya Adet, Düzenleyen, Alıcı, Para Birimi, Tutar, Miktar, Mal Tanımı, Teslim Şekli, Yükleme Tarihi, Yükleme Kaydı, Düzenlenen Asıl, Yükleme Limanı, Boşaltma Limanı, Şerh, Sigorta Tutarı, Teminat Başlangıcı, Koli, Brüt Ağırlık, Net Ağırlık, Menşe Ülke |

Zorunlu akreditif alanları: Para Birimi, Tutar, Vade, Lehtar, Amir. Girilmeyen alanın kontrolü atlanır.

## Çıktı

`cikti/akreditif_kontrol.xlsx`:

- `Özet`: tarihler, en geç ibraz günü, kalan gün, rezerv riski sayısı.
- `Rezervler`: durum (Rezerv riski / Kontrol edin / Bilgi), UCP maddesi, evrak, akreditifteki ve evraktaki değer, öneri;
  boş "Düzeltildi" sütunu.
- `Evrak Listesi`: istenen ve ibraz edilen asıl / kopya sayıları; akreditifteki özel şartlar.
- `Çapraz Kontrol`: aynı bilginin evraklardaki değerleri yan yana.

## Dikkat

- **Kod evrakı okumaz**, yalnız sizin girdiğiniz alanları karşılaştırır. İmza, ciro, kaşe, "freight prepaid" kaydı,
  konşimentonun kimin emrine düzenlendiği gibi özel şartlar `Evrak Listesi` sayfasında elle kontrol için durur.
- **Rezerv kararını banka verir.** Bu rapor bir ön kontroldür; temiz çıkması evrakların kabul edileceği anlamına
  gelmez. Akreditif metni UCP 600'den farklı bir hüküm içeriyorsa akreditif geçerlidir.
- **Tatiller:** Kod yalnız Cumartesi ve Pazar'ı bilir. Resmî tatilleri ve ibraz yapılacak bankanın bulunduğu ülkenin
  tatillerini ayrıca kontrol edin.
- **Tolerans (md. 30(b)):** ±%5 miktar toleransı, çekiş toplamı akreditif tutarını aşmamak kaydıyla geçerlidir; tutar
  kontrolü ayrıca yapılır.
- ISBP 821 (Uluslararası Standart Bankacılık Uygulaması) ayrıntıları kapsam dışıdır.
- **Örnek veri:** Akreditif, firma adları ve evrak bilgileri kurgusaldır.

## Testler

- Süreler: 21 gün, akreditifin verdiği süre, vadenin ve son günün hafta sonuna gelmesi, geç yükleme.
- Tutar ve miktar toleransları (about, %5, adet), kısmi yükleme.
- Fatura, konşimento ve sigorta kuralları; istenen evrak ve adetler; evraklar arası çelişki.
- Örnek veri uçtan uca (6 rezerv riski) ve komut satırı.

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
