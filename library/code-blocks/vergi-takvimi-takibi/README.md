# Vergi Takvimi Takibi · Kod Bloğu

> Muhasebe › Muhasebe Müdürü · Gerisi Bende

Şirketlerin beyan ve bildirim yükümlülüklerinden **yaklaşan son günleri** çıkarır; verilmiş beyanları işaretler, geciken ve
yaklaşanları hatırlatır. Hafta sonu ve resmî tatillerde son günü izleyen iş gününe kaydırır, GİB'in süre uzatmalarını parametreyle
alır. İnternete ve GİB sistemlerine bağlanmaz.

## Nasıl hesaplar?

| Yükümlülük | Son gün (genel kural) |
|---|---|
| KDV | Dönemi izleyen ayın 28'i (beyan ve ödeme) |
| Muhtasar ve prim hizmet (aylık) | Dönemi izleyen ayın 26'sı |
| Damga vergisi · Konaklama vergisi | Dönemi izleyen ayın 26'sı |
| Ba-Bs formları | Yalnız Ağustos 2024 ve önceki dönemler; Eylül 2024'ten itibaren kaldırıldı (565 Sıra No'lu VUK Genel Tebliği) |
| Geçici vergi | Çeyreği izleyen 2. ayın 17'si (17 Mayıs, 17 Ağustos, 17 Kasım, 17 Şubat). 4. dönem 2022–2024'te kaldırılmıştı (7338 sayılı Kanun); 2025 hesap döneminden itibaren yeniden verilir (7566 sayılı Kanun, RG 19.12.2025) |
| Kurumlar vergisi | Takvim yılını izleyen 30 Nisan |
| Gelir vergisi | İzleyen yılın 31 Mart'ı (beyan ve 1. taksit); 2. taksit 31 Temmuz |
| e-Defter berat (aylık) | Dönemi izleyen 4. ayın 10'u (gelir) / 14'ü (kurumlar); aralık için GV 10 Nisan, KV 14 Mayıs |

Son gün hafta sonuna ya da resmî tatile denk gelirse izleyen ilk iş günü alınır. Sabit ulusal tatiller ve 2026 dini bayramları
kodda; diğer yılların bayramlarını `--tatiller` ile ekleyin. GİB'in uzatmalarını `--uzatma ESKİ=YENİ` ile girin (ör. 2026 mayısında
26 Mayıs'taki beyanlar 3 Haziran'a uzatılmıştı). Pencere: bugünden `--geri` (60) gün öncesi ile `--ileri` (45) gün sonrası.

Durum: Gecikti · Bugün · Yaklaşıyor (`--uyari` 5 gün) · Planlı · Verildi · Geç verildi. Verilenler dosyasında dönem `2026/09`,
`2026/2. dönem`, `2025` ya da `2025 · 2. taksit` biçiminde yazılır.

**Uyarılar:** son günü geçmiş ve verildiği kaydı olmayan beyan · tanınmayan yükümlülük (Yüksek) · bugün ya da yakında biten beyan
(Orta) · geç verilmiş beyan (Bilgi).

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                          # örnek veri (14.10.2026)
python main.py --yukumlulukler yukumlulukler.xlsx --verilenler verilenler.xlsx
python main.py --yukumlulukler y.xlsx --verilenler v.xlsx --ileri 90 --uyari 7 --tatiller tatiller.csv
python main.py --yukumlulukler y.xlsx --uzatma 26.05.2026=03.06.2026 --uzatma 28.05.2026=05.06.2026
```

## Çıktı

`cikti/vergi_takvimi.xlsx`: `Vergi Takvimi` (son güne göre sıralı; tatil ya da uzatmayla kayan günlerde asıl gün notu), `Uyarılar`.

## Dikkat

- **Süreler kanun, tebliğ ve GİB duyurularıyla değişir.** Kod genel kuralları uygular; özel durumları (üç aylık muhtasar, özel
  hesap dönemi, tasfiye, işi bırakma, ÖTV, yerel asgari tamamlayıcı kurumlar vergisi gibi yeni yükümlülükler) kapsamaz.
  Takvimi her ay GİB vergi takvimi ve sirkülerleriyle karşılaştırın.
- **Ödeme ile beyan günleri** her yükümlülükte aynı olmayabilir; ödeme takibini ayrıca yapın.
- **Ba-Bs** formları Eylül 2024 döneminden itibaren kaldırıldı; kod yalnız geçmiş dönemler için son gün üretir.
- **Örnek veri:** Şirketler, sorumlular ve tarihler kurgusaldır.

## Testler

- KDV, muhtasar, kaldırılan Ba-Bs, 4 dönemli geçici vergi (hafta sonu kaydırma dahil), kurumlar ve gelir vergisi, e-defter berat son günleri; tatil
  zinciri (bayram + hafta sonu).
- Verildi / geç verildi / gecikti / bugün / yaklaşıyor durumları, uzatma, dönem anahtarı, okuma ve komut satırı.

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
| 4. geçici vergi döneminin yeniden getirilmesi — 7566 sayılı Kanun (RG 19.12.2025) | [PwC](https://www.pwc.com.tr/tr/Hizmetlerimiz/vergi/vergi-bultenleri/2025/7566-sayili-kanunun-vergi-duzenlemeleri.pdf), [KPMG](https://kpmgvergi.com/yayinlar/mali-bultenler/vergi/7566-sayili-kanun-resmi-gazete-de-yayimlandi/3316) | Teyit edildi (ikincil kaynaklar) |
| e-Defter berat yükleme süreleri — Elektronik Defter Genel Tebliği (Sıra No: 1) değişikliği (RG 08.11.2024) | [İSMMMO](https://ismmmo.org.tr/dosya/5113/Mevzuat-Dosya/08112024-edefter-gn-tb.pdf), [Sovos](https://sovos.com/tr/blog/kdv/e-defter-berat-yukleme-sureleri/) | Teyit edildi (ikincil kaynaklar) |
| Konaklama vergisi beyanı (izleyen ayın 26'sı) — GİB beyanname düzenleme kılavuzu | [GİB](https://intvrg.gib.gov.tr/KONAKLAMA_VERGISI_BEYANNAMESI_DUZENLEME_KILAVUZU.pdf) | Teyit edildi (resmî kılavuz) |
| Kurumlar vergisi beyanı (izleyen yılın 30 Nisan'ı) | [Habertürk](https://www.haberturk.com/kurumlar-vergisi-mukellefleri-dikkat-2026-kurumlar-vergisi-beyan-suresi-ne-zaman-iste-son-gun-3880958-ekonomi) | Teyit edildi (ikincil kaynak) |
| KDV beyanı (izleyen ayın 28'i) | [GİB vergi takvimi](https://www.gib.gov.tr/) | Teyit edilemedi: aramayla ayrıca doğrulanmadı; GİB vergi takviminden kontrol edin |

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
