# Vadesi Dolan Mevduat Takibi · Kod Bloğu

> Şube Bankacılığı › Bireysel Portföy Yöneticisi · Gerisi Bende

Önümüzdeki günlerde vadesi dolacak mevduat ve yatırım ürünlerini **müşteri bazında toplar** ve kimi önce
arayacağınızı sıralar: vadesi en yakın olanlar, sonra tutarı büyük olanlar. İnternete bağlanmaz.

## Nasıl hesaplar?

| Kavram | Kural |
|---|---|
| Kalan gün | Vade tarihi − durum tarihi |
| Pencere | Kalan günü `--gun` (15) veya daha az olan hesaplar; vadesi geçmiş kayıtlar da listeye girer ve uyarılır |
| TL karşılığı | Tutar × kur (TRY için 1; diğerleri `--kurlar` dosyasından) |
| Müşteri toplamı | Müşterinin penceredeki hesaplarının TL karşılığı toplamı |
| 1 – Acil | En yakın vadeye `--acil` (3) gün veya az kalmış ya da vade geçmiş |
| 2 – Yüksek | Müşteri toplamı `--buyuk` (1.000.000 TL) ve üstünde |
| 3 – Normal | Diğerleri |
| Sıra | Öncelik → müşteri toplamı (büyükten küçüğe) → en yakın vade |

**Örnek:** Müşterinin 750.000 TL mevduatı 2 gün sonra, 20.000 USD mevduatı 8 gün sonra doluyor; kur 42,50 → toplam
750.000 + 850.000 = **1.600.000 TL**, en yakın vadeye 2 gün → **1 – Acil**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                            # örnek veri, durum tarihi 12.10.2026
python main.py --hesaplar hesaplar.xlsx --kurlar kurlar.xlsx
python main.py --hesaplar hesaplar.xlsx --gun 30 --acil 5 --buyuk 2500000
```

Zorunlu sütunlar: Müşteri No, Tutar, Vade Tarihi. Para birimi boşsa TRY sayılır. Kuru verilmeyen dövizli hesap
listede görünür ama müşteri toplamına katılmaz ve uyarılır.

## Çıktı

`cikti/vadesi_dolan_mevduat_takibi.xlsx`:

- `Görüşme Listesi`: öncelik sırasıyla müşteriler; toplam, en yakın vade, ürünler, son görüşme, gerekçe, not sütunu.
- `Hesaplar`: penceredeki hesaplar vade sırasıyla.
- `Özet`: öncelik dağılımı, para birimi ve portföy yöneticisi bazında toplamlar, vade takvimi.
- `Uyarılar`: kuru eksik hesap, vadesi geçmiş kayıt, hafta sonuna denk gelen vade.

## Dikkat

- **Öncelik kuralı bir çalışma varsayımıdır.** Eşikleri şubenizin portföyüne göre ayarlayın. Kod müşterinin başka
  bankadaki varlığını, ilişki derinliğini ya da ayrılma eğilimini bilmez.
- **Kod ürün ya da oran önermez.** Hangi ürünün, hangi oranla teklif edileceği bankanızın fiyatlamasına, müşterinin
  risk profiline ve uygunluk değerlendirmesine bağlıdır. Bu rapor yatırım tavsiyesi değildir.
- **Kurlar sizin girdinizdir.** Örnek dosyadaki kurlar kurgusaldır; kendi kurunuzu girin. TL karşılığı yalnız
  sıralama içindir, muhasebe değeri değildir.
- **Tatil günleri bilinmez.** Hafta sonuna denk gelen vadeler bildirilir; resmî ve dinî tatilleri ayrıca kontrol edin.
- **Müşteri sırrı ve kişisel veri.** Rapor müşteri adı ve varlık bilgisi içerir; bankacılık sırrı ve kişisel
  verilerin korunması kurallarına göre yalnız yetkili kişilerce, kurum içi sistemlerde kullanılmalıdır. Aramadan
  önce müşterinin iletişim izinlerini kontrol edin.
- **Kapsam dışı:** stopaj ve net getiri hesabı, faiz teklifi karşılaştırması, kredi ve sigorta ürünleri, kampanya
  uygunluğu.
- **Örnek veri:** Müşteriler, tutarlar, oranlar ve kurlar kurgusaldır.

## Testler

- Pencere ve öncelik sınırları (bugün, 3 ve 4 gün, 15 ve 16 gün, eşiğin tam üstü ve bir kuruş altı), vadesi geçmiş.
- Dövizli hesap ve eksik kur, müşteri bazında toplama, sıralama, otomatik yenilemesiz hesap sayısı, hafta sonu vade.
- Para birimi, portföy yöneticisi ve takvim özetleri, okuma (para birimi yazımları) ve komut satırı.

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
