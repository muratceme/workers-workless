# Çek-Senet Portföy Takibi · Kod Bloğu

> Finans › Finans Uzmanı · Workers / Workless

Alınan ve verilen çek-senetleri vade, banka, keşideci ve durum bazında izler. Haftalık vade listesi çıkarır, verilen
çeklerin banka karşılığını kontrol eder, riskli evrakı işaretler. İnternete bağlanmaz.

## Nasıl çalışır?

- **Durumlar:**

| Durum | Anlamı |
|---|---|
| Portföyde, Tahsile Verildi, Teminata Verildi | Alınan evrak açık |
| Verildi | Verilen evrak açık (ödenmemiş) |
| Tahsil Edildi, Ödendi, İade | Kapalı |
| Ciro Edildi | Portföyden çıkar; vadesine kadar müracaat riski izlenir |
| Karşılıksız, Protestolu, Takipte | Sorunlu |

- **Haftalık vade listesi:** Pazartesi başlangıçlı haftalar (`--hafta`, varsayılan 8). Her hafta için beklenen
  tahsilat, teminattaki çekler (ayrı), ödenecek tutar, net ve kümülatif net gösterilir. Vadesi geçmiş açık evrak ayrı
  satırdadır. Döviz bazında ayrı hesaplanır.
- **Banka karşılık kontrolü:**
  - Her bankadan verilen çekler vade sırasıyla toplanır.
  - Toplam, banka bakiyesi ile o bankaya tahsile verilmiş ve o tarihe kadar vadesi gelen alınan çeklerin toplamıyla
    karşılaştırılır.
  - Açık varsa ilk eksik vade ve tutar raporlanır.
- **Ağırlıklı ortalama vade:** Açık alınan portföy için Σ(tutar × kalan gün) ÷ Σ tutar.

| Kontrol | Önem |
|---|---|
| Vade + ibraz süresi geçmiş, hâlâ portföydeki alınan çek (`--ibraz-gun`, varsayılan 10; TTK md. 796) | Yüksek |
| Vadesi geçmiş, portföydeki alınan evrak; ödenmemiş görünen verilen evrak | Yüksek |
| Karşılıksız / protestolu evrak ve aynı keşidecinin diğer açık evrakı | Yüksek |
| Verilen çeklerde banka karşılık açığı | Yüksek |
| Tek keşidecinin açık alınan portföydeki payı %30'un üstünde (`--yogunluk`) | Orta |
| Tahsile / teminata verilmiş, vadesi geçmiş, sonucu girilmemiş evrak | Orta |
| Ciro edilmiş, vadesi gelmemiş çek; bakiyesi verilmemiş banka | Bilgi |

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                     # örnek: 18 evrak, rapor tarihi 09.10.2026
python main.py --portfoy portfoy.xlsx --banka-bakiyeleri bakiyeler.xlsx --hafta 12
python main.py --portfoy portfoy.xlsx --bugun 01.11.2026 --ibraz-gun 30
```

| Dosya | Sütunlar |
|---|---|
| Portföy | Tür (Alınan Çek / Alınan Senet / Verilen Çek / Verilen Senet), Seri No, Keşideci / Borçlu, Cari, Banka, Vade, Tutar, Döviz, Durum, Durum Tarihi |
| Banka bakiyeleri (isteğe bağlı) | Banka, Bakiye, Döviz |

## Çıktı

`cek_senet_portfoyu.xlsx`:
- `Özet`: döviz ve durum bazında adet / tutar, ağırlıklı ortalama vade.
- `Haftalık Vade`: grafikli.
- `Banka Karşılık`: verilen çek bazında kümülatif tutar, bakiye, tahsile verilen çekler, açık.
- `Portföy`: tüm evrak; kalan gün, vadesi geçmiş açıklar renkli; boş "Aksiyon" sütunu.
- `Keşideci Riski`: açık tutar, portföy payı, karşılıksız / protesto geçmişi.
- `Uyarılar`.

## Dikkat

- **İbraz süresi:** Türk Ticaret Kanunu md. 796'ya göre çek, düzenlendiği yerde ödenecekse 10 gün, başka yerde
  ödenecekse 1 ay içinde ibraz edilmelidir. Süre keşide tarihinden başlar; ileri tarihli çekte bu, üzerindeki
  tarihtir. Başka yerde ödenecek çekler için `--ibraz-gun 30` kullanın. Hukuki sonuçları avukatınızla değerlendirin.
- **Karşılık:** Karşılıksız çek, 5941 sayılı Çek Kanunu kapsamında idari ve cezai yaptırımlara yol açabilir. Karşılık
  kontrolü yalnız verilen banka bakiyesine ve tahsile verilen çeklere dayanır. Kredili hesap limitleri ve diğer
  ödemeler dahil değildir.
- **Durum güncelliği:** Sonuçlar durum kayıtlarının güncel olmasına bağlıdır.
- **Örnek veri:** Örnek evraklar ve bakiyeler kurgusaldır.

## Testler

Şunlar test edilir:
- Açık portföy toplamları ve ortalama vade.
- İbraz süresi, vadesi geçmiş, karşılıksız / riskli keşideci, yoğunlaşma ve ciro uyarıları.
- Haftalık tahsilat / teminat / ödeme; banka karşılık açığı.
- Farklı rapor tarihi ve ibraz süresi; Excel ve CLI.

```bash
python -m unittest discover -s tests
```

## Kaynaklar ve dayanaklar

Kontrol tarihi: 15.10.2026. Mevzuat sık değişir; paketteki süre, oran ve kuralları kullanmadan önce güncel resmî metinle
karşılaştırın. "Karşılaştırılan maddeler" dışındaki atıflar yalnız yol göstericidir.

| Kaynak | Bağlantı | Durum |
|---|---|---|
| 6102 sayılı Türk Ticaret Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.6102.pdf) | Resmî metinle karşılaştırılan maddeler: 5/A, 22, 796, 1434 |
| 5941 sayılı Çek Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.5941.pdf) | Resmî metin bağlantısı; paketteki madde atıfları ayrıca karşılaştırılmadı |

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
