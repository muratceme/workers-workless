# Kredi Teklifi Eksik Belge Kontrolü · Kod Bloğu

> Krediler Tahsis › Kredi Tahsis Uzman Yardımcısı · Gerisi Bende

Şubeden gelen kredi tekliflerini **teklif tipine, müşteri tipine, teminata ve tutara göre** zorunlu belge ve bilgi
listesiyle karşılaştırır; eksikleri şubeye gönderilecek hazır metin olarak çıkarır. Zorunlu liste sizin kural
dosyanızdan okunur. İnternete bağlanmaz.

## Nasıl hesaplar?

1. **Kural eşleşmesi:** Her kural satırı bir belgeyi veya bilgiyi zorunlu kılar.

   | Sütun | Anlamı |
   |---|---|
   | Tür | `Belge` (belgeler dosyasında aranır) veya `Bilgi` (teklif dosyasında aynı adlı sütun dolu olmalı) |
   | Teklif Tipi, Müşteri Tipi | `*` veya boş = tümü |
   | Teminat | Teklifin teminatlarından biri eşleşirse kural uygulanır |
   | Asgari Tutar | Teklif tutarı bu tutara eşit veya büyükse kural uygulanır |
   | Geçerlilik (gün) | Belge tarihi kontrol tarihinden en çok bu kadar gün eski olabilir |

   Aynı belge birden çok kuralla isteniyorsa en kısa geçerlilik uygulanır.
2. **Belge durumu:** Eksik · Süresi geçmiş · Tarih yok (süreli belge, tarih girilmemiş) · Tamam. Aynı belgeden birden
   çok kayıt varsa en yeni tarihli olan esas alınır.
3. **Bilgi durumu:** Teklif satırında ilgili sütun boşsa Eksik.
4. **Sonuç:** Eksiği olmayan teklif "İncelemeye hazır", diğerleri "Şubeye iade".

Belge ve sütun adları büyük-küçük harf ve Türkçe karakter ayırmadan eşleştirilir.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek veri
python main.py --teklifler teklifler.xlsx --belgeler belgeler.xlsx --kurallar kurallar.xlsx --tarih 09.10.2026
```

| Dosya | Sütunlar |
|---|---|
| Teklifler | Teklif No, Şube, Müşteri, Teklif Tipi, Müşteri Tipi, Tutar, Teminat (virgülle), Gönderen + kurallardaki bilgi sütunları |
| Belgeler | Teklif No, Belge, Belge Tarihi |
| Kurallar | Tür, Ad, Teklif Tipi, Müşteri Tipi, Teminat, Asgari Tutar, Geçerlilik (gün), Açıklama |

Örnek kural satırları:

```text
Tür;Ad;Teklif Tipi;Müşteri Tipi;Teminat;Asgari Tutar;Geçerlilik (gün);Açıklama
Belge;Mali Tablolar;*;Tüzel;*;;150;Son dönem bilanço ve gelir tablosu
Belge;Ekspertiz Raporu;*;*;İpotek;;180;
Belge;Bağımsız Denetim Raporu;*;Tüzel;*;50.000.000;;
Bilgi;Ziyaret Tarihi;Yeni Limit;Tüzel;*;;;
```

## Çıktı

- `cikti/teklif_eksik_kontrolu.xlsx`:
  - `Teklif Özeti`: teklif bazında eksik belge / bilgi sayısı ve sonuç.
  - `Eksikler`: her eksik için durum, belge tarihi, yaşı, neden istendiği; boş "Şubeye Bildirildi" ve "Tamamlandı" sütunları.
  - `Şube Özeti`: iade oranı ve en sık eksikler.
  - `Uyarılar`: okunamayan kural, tanımsız teklif tipi, istenmeyen belge.
- `cikti/sube_iade_notlari.txt`: şube ve teklif bazında, e-postaya yapıştırılabilecek eksik listesi.

## Dikkat

- **Kural dosyası sizin sorumluluğunuzdadır.** Örnek dosyadaki belgeler, eşikler ve geçerlilik süreleri yalnız biçimi
  gösterir; bankanızın kredi politikasının veya mevzuatın istediği liste değildir.
- **İçerik kontrolü yapılmaz.** Kod belgenin gönderilip gönderilmediğine ve tarihine bakar; imza, onay, tutarlılık
  veya mali tabloların doğruluğunu değerlendirmez.
- **Kredi değerlendirmesi değildir.** "İncelemeye hazır" yalnız dosyanın tam olduğunu gösterir.
- **Bilgi sütunları:** Kural dosyasındaki bilgi adı teklif dosyasında sütun olarak yoksa tüm tekliflerde eksik sayılır
  ve uyarı verilir.
- **Kişisel veri ve ticari sır:** Teklif dosyaları müşteri bilgisi içerir; çıktıları yalnız yetkili kişilerle paylaşın.
- **Örnek veri:** Teklifler, firmalar ve kurallar kurgusaldır.

## Testler

- Kural koşulları: teklif tipi, müşteri tipi, teminat, tutar eşiği; en kısa geçerlilik.
- Belge durumları (eksik, süresi geçmiş, tarih yok, en yeni kayıt) ve bilgi kontrolü.
- Şube iade metni; okuma hataları; örnek veri uçtan uca.

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
