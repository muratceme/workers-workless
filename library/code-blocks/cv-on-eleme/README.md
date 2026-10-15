# CV Ön Eleme · Kod Bloğu

> İnsan Kaynakları › İnsan Kaynakları Uzman Yardımcısı · Workers / Workless

Gelen CV'leri ilanın **zorunlu kriterlerine** göre ilk filtreden geçirir. Her aday için **GEÇTİ / ELENDİ /
MANUEL KONTROL** kararını kriter bazında gerekçesiyle listeler. İnternete bağlanmaz.

## Kriterler (`kriterler.json`)

| Alan | Örnek | Anlamı |
|---|---|---|
| `en_az_deneyim_yil` | `5` | CV'deki tarih aralıklarından (çakışanlar birleştirilerek) hesaplanan toplam yıl |
| `en_az_egitim` | `"Lisans"` | Lise < Ön Lisans < Lisans < Yüksek Lisans < Doktora |
| `zorunlu_diller` | `["İngilizce"]` | Hepsi gerekli |
| `zorunlu_beceriler` | `["Excel", "Bordro"]` | Hepsi gerekli (Türkçe karakter ve büyük/küçük harf duyarsız) |
| `beceri_gruplari` | `[{"liste": ["SAP","Logo"], "en_az": 1}]` | Listeden en az N tanesi |
| `zorunlu_ifadeler` | `[{"ifade": "SMMM", "gerekce": "..."}]` | Belge, sertifika vb. |

## Adil ve temkinli karar

- **Belirsiz bilgi elemez:** CV'den okunamayan bilgi (taranmış PDF, tarihsiz deneyim, eğitim yok) adayı
  **elemez**; aday **MANUEL KONTROL**e düşer.
- **Ayrımcılık denetimi:** Kriterler ortak tarayıcıdan geçer. Yaş, cinsiyet, medeni hâl, görünüş, sağlık, din,
  köken gibi kriterler verilirse paket **çalışmaz**. Ehliyet, askerlik, uyruk gibi "dikkat" kriterleri yalnız
  `gerekce` alanıyla kabul edilir (4857 sayılı Kanun md. 5; 6701 sayılı Kanun md. 3, 6, 7).
- **Kişisel bilgiler:** Ad, fotoğraf ve iletişim bilgisi karara girmez; rapordaki iletişim için gösterilir.

CV okuma çekirdeği **CV Raporlama** kod bloğuyla ortaktır (`cv_cekirdek.py`).

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                                  # örnek 4 CV + kriterler_ornek.json
python main.py --girdi ./cvler --kriterler kriterler.json
```

> Kelime ve tarih tabanlı okuma yapar. Eş anlamlı ifadeler ve tablo/görsel ağırlıklı CV'ler kaçabilir.
> Elenen adaylara dönüş yapmadan önce örneklem kontrolü yapın.

## Çıktı

`Ön Eleme` (karar, gerekçe, kriter sütunları, "İK Onayı") · `Kriterler` · `Bilgi`

## Testler

```bash
python -m unittest discover -s tests
```

## Kaynaklar ve dayanaklar

Kontrol tarihi: 15.10.2026. Mevzuat sık değişir; paketteki süre, oran ve kuralları kullanmadan önce güncel resmî metinle
karşılaştırın. "Karşılaştırılan maddeler" dışındaki atıflar yalnız yol göstericidir.

| Kaynak | Bağlantı | Durum |
|---|---|---|
| 4857 sayılı İş Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.4857.pdf) | Resmî metinle karşılaştırılan maddeler: 17, 41, 46, 47, 53, 56, 63, 68, 69 |
| 6701 sayılı Türkiye İnsan Hakları ve Eşitlik Kurumu Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.6701.pdf) | Resmî metin bağlantısı; paketteki madde atıfları ayrıca karşılaştırılmadı |

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
