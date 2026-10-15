# Donanım Zimmet Envanteri · Kod Bloğu

> Bilgi Teknolojileri › BT Destek Uzmanı · Workers / Workless

Bilgisayar, telefon ve çevre birimi listenizi denetler: **kimde ne var, hangi cihazın garantisi bitiyor, hangisinin
yenileme zamanı gelmiş, hangi kayıtta tutarsızlık var.** Güncel personel listesini verirseniz, işten ayrılmış
kişilerin üzerinde kalan cihazları da bulur. İnternete bağlanmaz.

## Nasıl hesaplar?

| Kavram | Kural |
|---|---|
| Garanti | Bitiş tarihi geçmişse "Bitmiş", `--gun` (60) gün içindeyse "Bitiyor", değilse "Devam ediyor" |
| Planlanan yenileme | Alım tarihi + yenileme ömrü (satırdaki değer; boşsa `--omur`, varsayılan 4 yıl) |
| Yaş | (Durum tarihi − alım tarihi) ÷ 365,25 yıl |
| Durum boşsa | Zimmetlisi olan cihaz "Kullanımda", olmayan "Depoda" sayılır |

Zimmet denetimleri:

- Kullanımda görünen ama zimmetlisi yazılmamış cihaz,
- depoda veya hurda olduğu hâlde zimmeti düşülmemiş cihaz,
- aynı demirbaş numarası veya seri numarasının iki kayıtta bulunması,
- zimmetli kişinin `--personel` listesinde olmaması,
- bir kişide aynı türden birden çok kullanımda cihaz.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                              # örnek veri, durum tarihi 12.10.2026
python main.py --envanter envanter.xlsx
python main.py --envanter envanter.xlsx --personel personel.xlsx --gun 90 --omur 5
```

Zorunlu sütunlar: Demirbaş No, Tür. Diğerleri: Marka Model, Seri No, Zimmetli, Departman, Zimmet Tarihi, Alım
Tarihi, Garanti Bitiş, Durum (Kullanımda / Depoda / Arızalı / Hurda), Yenileme Ömrü (yıl). Personel dosyasında
"Ad Soyad" sütunu yeterlidir; eşleşme ada göre yapılır.

## Çıktı

`cikti/donanim_zimmet_envanteri.xlsx`:

- `Envanter`: her cihazın durumu, zimmetlisi, yaşı, garanti ve yenileme durumu.
- `Kişi Özeti`: kişi bazında üzerindeki cihazlar (zimmet tutanağı ve işten çıkış kontrolü için).
- `Tür Özeti`: tür bazında adet, durum dağılımı, garantisi biten ve yenileme zamanı gelen cihaz sayısı.
- `Uyarılar`: tutarsız kayıtlar, ayrılmış kişide kalan cihazlar, garanti ve yenileme hatırlatmaları.

## Dikkat

- **Envanter sizin kaydınızdır.** Kod cihazın fiilen kimde olduğunu bilemez; kaydı düzenli sayımla doğrulayın.
- **Kişi eşleşmesi ada göredir.** Aynı adı taşıyan iki çalışan ayırt edilemez; yazım farkı (ikinci ad, kısaltma)
  "listede yok" uyarısı üretebilir. Mümkünse sicil numarası olan tek bir yazım kullanın.
- **Yenileme ömrü şirket tercihidir.** 4 yıl genel bir başlangıç değeridir; muhasebedeki faydalı ömürle aynı şey
  değildir. Cihaz türüne göre "Yenileme Ömrü (yıl)" sütununu doldurun.
- **Kişisel veri:** Envanter ad, departman ve cihaz bilgisi içerir; raporu yalnız ilgili kişilerle paylaşın.
- **Kapsam dışı:** yazılım lisansları, cihazdaki verinin silinmesi, zimmet tutanağı basımı, amortisman.
- **Örnek veri:** Kişiler, cihazlar ve seri numaraları kurgusaldır.

## Testler

- Garanti ve yenileme durumları (sınır günleri, 29 Şubat, satır bazında ömür), yaş hesabı.
- Zimmet denetimleri: zimmetsiz, düşülmemiş zimmet, mükerrer numara, listede olmayan kişi, aynı türden çoklu cihaz.
- Kişi ve tür özetleri, okuma hataları, komut satırı.

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
