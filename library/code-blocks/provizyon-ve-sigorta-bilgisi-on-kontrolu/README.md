# Provizyon ve Sigorta Bilgisi Ön Kontrolü · Kod Bloğu

> Hasta Hizmetleri › Hasta Kayıt Kabul Görevlisi · Workers / Workless

Randevu listesindeki hastaların kimlik, ödeme tipi, sigorta ve provizyon bilgilerini **hasta gelmeden önce** kontrol
eder; eksikleri hasta bazında listeler. İnternete, Medula'ya veya sigorta şirketine bağlanmaz.

## Nasıl kontrol eder?

| Kontrol | Kural | Önem |
|---|---|---|
| Kimlik no | 11 hane, ilk hane 0 değil, son iki hane doğrulama algoritmasını tutuyor | Kritik |
| SGK müstahaklığı (SGK, TSS) | "Yok" yazılmışsa kritik; boşsa "sorgulanmamış" | Kritik / Eksik |
| Sigorta şirketi, poliçe no, poliçe bitişi (ÖSS, TSS) | Boş veya okunamıyor | Eksik |
| Poliçe süresi (ÖSS, TSS) | Poliçe bitişi randevu tarihinden önce | Kritik |
| Provizyon (ÖSS, TSS) | Reddedilmiş → kritik; beklemede veya numara yok → eksik | Kritik / Eksik |
| Ödeme tipi | SGK, ÖSS, TSS, Ücretli dışında veya boş | Eksik |
| Mükerrer randevu | Aynı hasta, aynı gün, aynı branş | Eksik |
| Telefon | Boş | Bilgi |

Kimlik doğrulaması: 10. hane = ((1, 3, 5, 7, 9. haneler toplamı × 7) − (2, 4, 6, 8. haneler toplamı)) mod 10;
11. hane = ilk 10 hanenin toplamı mod 10.

Durum: yüksek önemli eksik varsa **Kritik**, orta önemli eksik varsa **Eksik var**, yoksa **Hazır**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                  # örnek veri
python main.py --randevular randevular.xlsx
python main.py --randevular randevular.xlsx --baslangic 13.10.2026 --bitis 13.10.2026
```

Zorunlu sütunlar: Randevu Tarihi, Hasta, Ödeme Tipi. Diğerleri: Saat, TC Kimlik No, Branş, Hekim, Sigorta Şirketi,
Poliçe No, Poliçe Bitiş, Provizyon No, Provizyon Durumu, Müstahaklık (Var / Yok), Telefon.

## Çıktı

`cikti/provizyon_on_kontrol.xlsx`:

- `Kontrol Listesi`: randevu sırasıyla hasta, durum ve eksikler. Kimlik numarası maskelidir (`123******50`).
- `Özet`: ödeme tipi bazında hazır, eksik ve kritik randevu sayısı.
- `Uyarılar`: her eksik, ne yapılacağıyla birlikte.

## Dikkat

- **Kod sorgulama yapmaz.** Müstahaklık ve provizyon sonucunu Medula'dan veya sigorta şirketinden siz sorgulayıp
  listeye yazarsınız; kod yalnız listede yazanı denetler. "Hazır" görünen hastanın güvencesi o gün değişmiş olabilir.
- **Kimlik algoritması yalnız yazım hatasını yakalar.** Numaranın o kişiye ait olduğunu göstermez.
- **Provizyon her işlem için gerekmeyebilir.** Hangi işlemlerin ön onay istediği sigorta şirketine ve poliçeye göre
  değişir; "Provizyon alınmamış" bir hatırlatmadır.
- **Kişisel veri.** Liste kimlik ve sağlık verisi içerir. Girdi dosyasını ve raporu yetkisiz kişilerle paylaşmayın;
  KVKK ve kurumunuzun veri politikasına uyun.
- **Örnek veri:** Hasta adları, kimlik numaraları, şirketler ve poliçeler kurgusaldır.

## Testler

- Kimlik doğrulama, maskeleme; her kuralın tek tek sonucu (SGK, ÖSS, TSS, Ücretli).
- Mükerrer randevu, tarih filtresi, ödeme tipi özeti.
- Okunamayan satır, eksik sütun, komut satırı hataları; örnek veri uçtan uca ve çıktıda açık kimlik no bulunmaması.

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
