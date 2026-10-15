# Gelen e-Fatura Kayıt Kontrolü · Kod Bloğu

> Muhasebe › Muhasebe Elemanı · Mali Müşavirlik Bürosu › Muhasebe Elemanı · Workers / Workless

GİB veya özel entegratör portalından alınan **gelen e-fatura / e-arşiv** listesini muhasebe **alış
kayıtlarıyla** karşılaştırır. Deftere işlenmemiş, mükerrer işlenmiş, tutarı farklı ve portalda karşılığı
olmayan kayıtları ayrı sayfalarda gösterir. İnternete bağlanmaz.

> Form Ba-Bs bildirimi, 565 Sıra No'lu VUK Genel Tebliği ile 01.10.2024'ten itibaren kaldırılmıştır.
> e-Belge verileri İdarece doğrudan izlenebildiğinden, gelen e-faturaların eksiksiz ve doğru kaydedildiğini
> dönem sonunda kontrol etmek daha da önemli hale gelmiştir.

## Eşleştirme

1. Fatura numarası + satıcı VKN (boşluk/tire ve büyük-küçük harf farkı yok sayılır)
2. Fatura numarası (kayıtta VKN yoksa)
3. Satıcı VKN + vergiler dahil tutar + tarih farkı ≤ `--gun-toleransi` → belge numarası **yanlış girilmiş**
   kayıtları bulur (Eşleşenler sayfasında sarı)

Dövizli e-faturalar faturadaki kurla TL'ye çevrilir. Aynı faturanın ZIP içindeki ikinci kopyası bir kez sayılır.

## Çıktı (`cikti/efatura_kayit_kontrolu.xlsx`)

| Sayfa | Anlamı |
|---|---|
| Deftere İşlenmemiş | Portalda var, muhasebede yok — KDV indirimi ve gider kaybı riski |
| Tutarı Farklı | Eşleşti ama kayıt tutarı farklı |
| Portalda Yok | Muhasebede var, e-fatura listesinde yok — kâğıt fatura, SMM, ithalat, yanlış dönem veya mükerrer kayıt olabilir |
| Mükerrer Kayıtlar | Aynı belge no + VKN birden çok kez kaydedilmiş |
| Eşleşenler | Tüm eşleşmeler ve yöntemi |

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                                         # örnek veriyle dener
python main.py --efatura "C:/Gelen/2026-09" --kayitlar muavin_alislar.xlsx
python main.py --efatura portal_listesi.xlsx --kayitlar alis_kayitlari.csv --gun-toleransi 5
```

Muhasebe kaydı sütunları (başlıklar esnek): **Belge No / Fatura No** ve **Tutar** zorunlu; Tarih, Cari VKN /
Vergi No, Cari Unvan, Açıklama isteğe bağlı. Portal listesi için de aynı başlık esnekliği geçerlidir
(ör. "Gönderici VKN/TCKN", "Vergiler Dahil Tutar").

## Testler

```bash
python -m unittest discover -s tests
```

## Kaynaklar ve dayanaklar

Kontrol tarihi: 15.10.2026. Mevzuat sık değişir; paketteki süre, oran ve kuralları kullanmadan önce güncel resmî metinle
karşılaştırın. "Karşılaştırılan maddeler" dışındaki atıflar yalnız yol göstericidir.

| Kaynak | Bağlantı | Durum |
|---|---|---|
| 213 sayılı Vergi Usul Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.4.213.pdf) | Resmî metinle karşılaştırılan maddeler: 274, 280, 315, mükerrer 315, 320, 323 |
| Form Ba-Bs'nin Eylül 2024 döneminden itibaren kaldırılması — 565 Sıra No'lu VUK Genel Tebliği (RG 25.09.2024 / 32673) | [Andersen](https://tr.andersen.com/tr/images/pdf/and_41-ba-bs-formlari-kaldirildi.pdf), [İSMMMO](https://ismmmo.org.tr/dosya/5020/Mevzuat-Dosya/25092024-vuk565-sn-gn-tb-babs-kaldirildi.pdf), [Alomaliye](https://www.alomaliye.com/2024/09/25/vergi-usul-kanunu-genel-tebligi-sira-no-565/amp/) | Teyit edildi (ikincil kaynaklar) |

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
