# Grup Rezervasyon Teklifi · Kod Bloğu

> Rezervasyon ve Gelir Yönetimi › Rezervasyon Sorumlusu · Gerisi Bende

Grup talebini (gün × kalem × adet) otelin **fiyat listesiyle** fiyatlandırır; oda, yemek, toplantı salonu ve diğer
hizmetleri gün gün hesaplayıp ücretsiz oda, grup indirimi, konaklama vergisi ve KDV ile **teklif dosyası** hazırlar.
İnternete bağlanmaz.

## Nasıl hesaplar?

| Adım | Formül |
|---|---|
| Satır tutarı | Adet × o gün geçerli KDV hariç fiyat (fiyat listesinde Başlangıç / Bitiş varsa sezon fiyatı seçilir) |
| Ücretsiz oda | Her gece ⌊o gecenin oda adedi ÷ (N + 1)⌋ oda, o gecenin en düşük oda fiyatından düşülür (`--ucretsiz N`; 20 → "20+1") |
| Grup indirimi | (Tutar − ücretsiz oda) × `--indirim` % |
| Konaklama vergisi | Konaklama vergisine tabi satırların indirimli tutarı × oran; **KDV matrahına girmez** |
| KDV | Her satırın indirimli tutarı × fiyat listesindeki KDV oranı |
| Genel toplam | Net + konaklama vergisi + KDV |
| Kapora | Genel toplam × `--kapora` % (varsayılan 30) |
| Opsiyon / bakiye | Opsiyon = teklif tarihi + `--opsiyon` gün (7) · bakiye son ödeme = ilk gün − `--bakiye-gun` (7) |

**Konaklama vergisi oranı:** `--kv` verilmezse geceye göre seçilir: 11263 sayılı Cumhurbaşkanı Kararıyla **01.05.2026 –
31.12.2026** arasındaki gecelere **%1**, diğer tarihlere kanuni oran **%2**. Fiyat listesinde "Konaklama Vergisi" sütunu boşsa
yalnız Tür = Oda satırları vergiye tabi sayılır. Yemek konaklamayla aynı bedel içinde satılıyorsa (ör. yarım pansiyon paketi)
o satırı "Evet" işaretleyin; ayrıca satılan yemek, salon ve diğer hizmetler vergiye tabi değildir.

**Uyarılar:** fiyat listesinde geçerli fiyatı olmayan kalem (Yüksek; teklife girmez) · aynı gün için birden fazla fiyat ·
bakiye tarihi teklif tarihinden önce · opsiyon girişten sonra (Orta) · kişi sayısı verilmedi (Bilgi).

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                                    # örnek veri (48 kişilik 3 günlük kongre)
python main.py --talep talep.xlsx --fiyatlar fiyatlar.xlsx --grup "ABC Derneği" --kisi 48
python main.py --talep talep.xlsx --fiyatlar fiyatlar.xlsx --ucretsiz 20 --indirim 5 --kapora 40 --opsiyon 10
python main.py --talep talep.xlsx --fiyatlar fiyatlar.xlsx --kv 2                  # tek konaklama vergisi oranı
```

Fiyatları **KDV hariç** yazın. Kalem adları talep ve fiyat listesinde aynı olmalıdır (büyük-küçük harf ve Türkçe karakter
farkı önemsenmez).

## Çıktı

`cikti/grup_teklifi.xlsx`: `Teklif` (müşteriye gidecek özet, tutarlar ve koşullar), `Detay` (gün × kalem; ücretsiz oda
satırı yeşil), `Kalem Özeti`, `Uyarılar`.

## Dikkat

- **Vergi oranlarını doğrulayın.** Konaklama vergisi oranı ve süresi Cumhurbaşkanı Kararıyla değişebilir; KDV oranları fiyat
  listenizden okunur, örnek listedeki oranlar (konaklama ve yemek %10, salon ve ekipman %20) yalnız örnektir. Güncel oranları
  Resmî Gazete ve GİB duyurularından ya da mali müşavirinizden teyit edin.
- **Teklif müsaitlik kontrolü yapmaz.** Oda ve salon müsaitliğini PMS ve satış takviminden kontrol edin.
- **Ücretsiz oda, indirim, kapora ve opsiyon** otel politikasıdır; varsayılanlar yalnız örnektir. İptal, düşüş (attrition)
  ve oda bloğu bırakma tarihleri sözleşmede ayrıca yazılmalıdır.
- **Örnek veri:** Fiyatlar ve grup kurgusaldır.

## Testler

- Satır tutarları, sezon fiyatı seçimi, ücretsiz oda (N+1), indirim, konaklama vergisi (KDV matrahı dışında), KDV, genel
  toplam, kişi başı, kapora, opsiyon ve bakiye tarihleri.
- Tarihe göre konaklama vergisi oranı, uyarılar, okuma hataları ve komut satırı.

```bash
python -m unittest discover -s tests
```

## Kaynaklar ve dayanaklar

Kontrol tarihi: 15.10.2026. Mevzuat sık değişir; paketteki süre, oran ve kuralları kullanmadan önce güncel resmî metinle
karşılaştırın. "Karşılaştırılan maddeler" dışındaki atıflar yalnız yol göstericidir.

| Kaynak | Bağlantı | Durum |
|---|---|---|
| Konaklama vergisi oranının 31.12.2026'ya kadar %1 olması — 11263 sayılı Cumhurbaşkanı Kararı (RG 01.05.2026 / 33240) | [KPMG](https://kpmgvergi.com/yayinlar/mali-bultenler/vergi/konaklama-vergisi-orani-11263-sayili-cumhurbaskani-karari-ile-yeniden-belirlendi/4483), [PwC](https://www.pwc.com.tr/tr/Hizmetlerimiz/vergi/vergi-bultenleri/2026/konaklama-vergisi-orani-1e-indirildi.pdf), [Alomaliye](https://www.alomaliye.com/2026/05/01/konaklama-vergisi-orani-2026-sonuna-kadar-yuzde-1-olarak-uygulanacak/) | Teyit edildi (ikincil kaynaklar) |

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
