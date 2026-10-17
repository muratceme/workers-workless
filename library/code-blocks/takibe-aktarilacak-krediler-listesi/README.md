# Takibe Aktarılacak Krediler Listesi · Kod Bloğu

> Krediler İzleme ve Takip › Yasal Takip Uzmanı · Gerisi Bende

Gecikmesi eşiği aşan ya da hesabı kat edilen kredileri **yasal takibe aktarım listesine** dönüştürür; her kredi için
gereken evrakın tamam olup olmadığını gösterir, eksik evrakı adıyla yazar. İnternete bağlanmaz.

## Nasıl hesaplar?

| Kavram | Kural |
|---|---|
| Aday | Gecikme günü `--gun` (90) günü aşmış **ya da** hesap kat tarihi dolu |
| Evrak sütunları | Bilinen sütunların (Müşteri No, Müşteri, Kredi No, Risk, Gecikme Gün, Hesap Kat Tarihi, Not) dışındaki her sütun |
| Evrak değeri | Var · Yok · Gerekmez (boş hücre "Yok" sayılır) |
| Aktarıma hazır | Aday ve gereken evrakın tamamı "Var" |
| Evrak eksik | Aday ve en az bir evrak "Yok" |
| Aday değil | Eşiği aşmamış ve kat edilmemiş |
| Sıra | Evrak eksik → Aktarıma hazır → Aday değil; içinde risk büyükten küçüğe |

**Örnek:** Gecikme 140 gün, hesap 01.09.2026'da kat edilmiş, beş evraktan "Tebliğ Şerhi" yok → **Evrak eksik**;
rapor eksik evrakı adıyla yazar.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                            # örnek veri
python main.py --krediler krediler.xlsx
python main.py --krediler krediler.xlsx --gun 120
```

Zorunlu sütunlar: Müşteri No, Risk. Evrak listesini dosyanın sütunlarıyla siz tanımlarsınız; yeni bir evrak için
sütun eklemeniz yeterlidir. Bir kredi için aranmayan evrakı "Gerekmez" yazın (örneğin kefili olmayan kredide kefil
bilgileri).

## Çıktı

`cikti/takibe_aktarilacak_krediler.xlsx`:

- `Aktarım Listesi`: durum, risk, gecikme, kat tarihi, aday olma nedeni, her evrakın durumu ve eksik evrak.
- `Özet`: durum bazında kredi sayısı ve risk; adaylarda evrak bazında var / eksik / gerekmez sayıları.
- `Uyarılar`: evrakı eksik adaylar, eşiği aştığı hâlde hesabı kat edilmemiş krediler.

## Dikkat

- **Evrak listesi sizindir.** Örnek dosyadaki beş sütun yalnız gösterim içindir; takip dosyası için gereken evrak
  kredinin türüne, teminatına ve bankanızın hukuk birimi talimatına göre değişir. Kod bir evrakın hukuken yeterli
  ya da geçerli olduğunu değerlendirmez; yalnız işaretinize bakar.
- **Eşik ve aktarım kararı politikadır.** 90 gün varsayılandır. Bir kredinin ne zaman kat edileceği ve yasal takibe
  ne zaman başlanacağı ilgili düzenlemelere, sözleşme hükümlerine ve bankanızın yetkili birimlerinin kararına
  bağlıdır; liste karar vermez, hazırlık yapar.
- **Süreler izlenmez.** İhtarnamede verilen süre, tebliğ tarihi, zamanaşımı ve takip yolunun seçimi bu paketin
  kapsamında değildir; bunları avukatınız değerlendirir.
- **Müşteri bazında birleştirme yapılmaz.** Her satır bir kredidir; aynı müşterinin diğer kredilerinin temerrüt
  nedeniyle muaccel olup olmadığını ayrıca değerlendirin.
- **Müşteri sırrı:** Rapor müşteri risk bilgisi içerir; yalnız yetkili kişilerle ve avukatla paylaşın.
- **Örnek veri:** Müşteriler, krediler ve tutarlar kurgusaldır.

## Testler

- Adaylık sınırları (90 ve 91 gün, kat edilmiş ama gecikmesi düşük kredi), özel eşik.
- Evrak değerleri (Var, Yok, boş, Gerekmez), evrak sütunu olmayan dosya, sütun sırası farklı dosya.
- Sıralama, durum ve evrak özeti, hesap kat edilmemiş uyarısı, okuma ve komut satırı.

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
