# Periyodik Kontrol Takibi · Kod Bloğu

> İş Sağlığı ve Güvenliği › İş Güvenliği Uzmanı · Gerisi Bende

Kaldırma ekipmanı, basınçlı kap, elektrik tesisatı, yangın ekipmanı gibi iş ekipmanlarının **periyodik kontrol
tarihlerini** izler: süresi geçmiş, yaklaşan ve hiç kontrol edilmemiş ekipmanı listeler, aylık kontrol planını
çıkarır. İnternete bağlanmaz.

## Nasıl hesaplar?

| Kavram | Kural |
|---|---|
| Sonraki kontrol | Son kontrol tarihi + kontrol periyodu (ay); ay sonu taşarsa ayın son günü (31.08 + 6 ay = 28.02) |
| Kalan gün | Sonraki kontrol − durum tarihi |
| Süresi geçmiş | Kalan gün sıfırın altında |
| Yaklaşıyor | Kalan gün `--gun` (30) veya daha az |
| Hiç kontrol edilmemiş | Son kontrol tarihi yazılmamış |
| Periyot tanımsız | Kontrol periyodu yazılmamış; sonraki tarih hesaplanamaz |
| Uygun Değil | Son kontrolün sonucu uygun değilse tarih ne olursa olsun ayrıca uyarılır |

**Örnek:** Son kontrol 01.07.2026, periyot 3 ay → sonraki kontrol **01.10.2026**; durum tarihi 12.10.2026 →
**11 gün geçmiş**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                            # örnek veri, durum tarihi 12.10.2026
python main.py --ekipman ekipman.xlsx
python main.py --ekipman ekipman.xlsx --gun 45 --tarih 01.11.2026
```

Zorunlu sütun: Ekipman. Kontrol periyodu ay olarak tam sayı yazılır (`3`, `12`, `12 ay`).

## Çıktı

`cikti/periyodik_kontrol_takibi.xlsx`:

- `Ekipmanlar`: her ekipman için son ve sonraki kontrol, kalan gün, durum, son sonuç, kontrolü yapan, rapor no.
- `Yapılacak Kontroller`: hiç kontrol edilmemiş, süresi geçmiş ve yaklaşan ekipman tarih sırasıyla.
- `Tür Özeti`, `Konum Özeti`: durumlara göre ekipman sayıları.
- `Aylık Plan`: hangi ay kaç kontrol var (kontrol firmasıyla randevu planlamak için).
- `Uyarılar`.

## Dikkat

- **Kontrol periyodunu kod bilmez, siz yazarsınız.** Süreler ekipmanın türüne göre değişir; İş Ekipmanlarının
  Kullanımında Sağlık ve Güvenlik Şartları Yönetmeliği'ne, ilgili standarda ve imalatçı talimatına bakarak her
  ekipman için doğru süreyi girin. Örnek dosyadaki süreler yalnız gösterim içindir.
- **Kimin kontrol edebileceği ayrı bir konudur.** Periyodik kontrolü yapmaya yetkili kişinin niteliği mevzuatta
  ekipman türüne göre tanımlanır; kod raporu yazanın yetkisini doğrulamaz.
- **Periyodik kontrol tek tetikleyici değildir.** Yer değişikliği, tadilat, kaza ya da uzun süre kullanılmama gibi
  durumlarda süre dolmadan kontrol gerekebilir; bunlar bu listede görünmez.
- **"Uygun" durumu güvenli demek değildir.** Yalnız kontrol tarihinin geçmediğini gösterir; günlük ve kullanım
  öncesi kontroller, bakım ve risk değerlendirmesi ayrıca yürütülür.
- **Kapsam dışı:** bakım planı, kalibrasyon, kişisel koruyucu donanım, sağlık gözetimi, eğitim ve ortam ölçümü
  takibi.
- **Örnek veri:** Ekipmanlar, firmalar, tarihler ve rapor numaraları kurgusaldır.

## Testler

- Durum sınırları: dün / bugün / 30 ve 31 gün sonra dolan kontrol; ay sonu taşması ve artık yıl.
- Hiç kontrol edilmemiş, periyodu boş ekipman, "Uygun Değil" sonucu, ileri tarihli kontrol.
- İş listesi sırası, tür ve konum özeti, aylık plan, okuma (periyot yazımları) ve komut satırı.

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
