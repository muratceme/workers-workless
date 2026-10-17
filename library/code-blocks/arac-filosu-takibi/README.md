# Araç Filosu Takibi · Kod Bloğu

> İdari İşler › İdari İşler Sorumlusu · Gerisi Bende

Şirket araçlarının **muayene, trafik sigortası, kasko, periyodik bakım ve lastik** durumunu tek tabloda izler;
süresi geçmiş ve yaklaşan işleri tarih sırasıyla listeler. Muayenesi ya da sigortası bitmiş aracın yola çıkmasını
önlemek için haftada bir çalıştırmanız yeterlidir. İnternete bağlanmaz.

## Nasıl hesaplar?

| Kavram | Kural |
|---|---|
| Belge durumu | Muayene, trafik sigortası ve kasko için: bitiş geçmişse "Geçmiş", `--gun` (30) gün veya daha az kaldıysa "Yaklaşıyor", değilse "Uygun" |
| Sonraki bakım | Son bakım km + bakım aralığı (km) **ve** son bakım tarihi + bakım aralığı (ay); hangisi önce dolarsa |
| Bakım durumu | Güncel km sonraki bakım km'sini aşmışsa veya tarih geçmişse "Geçmiş"; kalan km `--km-esik` (1.000) ve altındaysa veya tarihe `--gun` gün kaldıysa "Yaklaşıyor" |
| Lastik | Kış lastiği dönemini verirseniz: dönem içinde yaz lastiği, dönem dışında kış lastiği olan araçlar işaretlenir |
| Araç durumu | Herhangi bir kalem geçmişse "Acil", yaklaşıyorsa "Yaklaşan", değilse "Uygun" |

**Örnek:** Son bakım 80.000 km'de ve 05.11.2025'te, aralık 15.000 km veya 12 ay, güncel km 95.400 → sonraki bakım
95.000 km: **400 km geçmiş**; tarihe göre 05.11.2026.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                            # örnek veri, durum tarihi 12.10.2026
python main.py --araclar araclar.xlsx
python main.py --araclar araclar.xlsx --gun 45 --km-esik 2000 --kis-baslangic 01.12 --kis-bitis 01.04
```

Zorunlu sütun: Plaka. Diğerleri: Marka Model, Kullanıcı, Muayene Bitiş, Trafik Sigortası Bitiş, Kasko Bitiş, Son
Bakım Tarihi, Son Bakım Km, Güncel Km, Bakım Aralığı (km), Bakım Aralığı (ay), Lastik (Yaz / Kış / Dört Mevsim).

## Çıktı

`cikti/arac_filosu_takibi.xlsx`:

- `Araçlar`: her araç için genel durum; muayene, trafik sigortası, kasko ve bakımın tarihi ve durumu; kalan km.
- `Yapılacak İşler`: geçmiş ve yaklaşan işler tarih sırasıyla; "Yapıldı" sütununu siz doldurursunuz.
- `Uyarılar`: süresi geçmiş ve yaklaşan belgeler, bakım, lastik, tutarsız km ve eksik bilgi.

## Dikkat

- **Tarihleri belgeden girin.** Kod muayene sıklığını, sigorta süresini veya aracınızın bakım aralığını bilmez;
  bitiş tarihlerini ruhsat, poliçe ve servis kaydından siz yazarsınız. Araç türüne göre muayene süresi farklıdır.
- **Kış lastiği dönemini siz verirsiniz.** Zorunluluğun kimleri kapsadığı ve hangi tarihler arasında geçerli
  olduğu mevzuatla belirlenir, illere göre uzatılabilir. Örnek çalıştırmadaki 01.12 – 01.04 yalnız gösterim
  içindir; güncel düzenlemeyi kontrol edip kendi tarihlerinizi girin.
- **Güncel km sizin beyanınızdır.** Km düzenli güncellenmezse bakım uyarısı geç gelir; tarihe göre bakım aralığını
  da doldurun.
- **Kapsam dışı:** motorlu taşıtlar vergisi, trafik cezaları, geçiş ücretleri, yakıt tüketimi, hasar kayıtları,
  sürücü belgesi ve yetki belgeleri, kiralık araçlarda sözleşme bitişi ve km sınırı.
- **Kişisel veri:** Kullanıcı adı yazıyorsanız raporu yalnız ilgili kişilerle paylaşın.
- **Örnek veri:** Plakalar, araçlar ve tarihler kurgusaldır.

## Testler

- Belge durumu sınırları (geçmiş, bugün, 30 ve 31 gün), eksik tarih, kasko için ayrı önem derecesi.
- Bakım: km ve tarih ölçütü, eşik sınırı, tutarsız km, yalnız km veya yalnız tarih bilgisi.
- Kış dönemi (yıl sonunu aşan dönem), lastik uyarıları, araç durumu, iş listesi sırası, okuma ve komut satırı.

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
