# Test Sürüşü ve Teklif Takibi · Kod Bloğu

> Satış › Satış Danışmanı · Workers / Workless

Showroom'a gelen, test sürüşü yapan ve teklif alan müşterileri izler; **açık fırsatlar için geri dönüş önceliği** belirleyip
gerekçesiyle listeler. Danışman ve model bazında test sürüşü, teklif ve kapanış oranlarını, kayıp nedenlerini çıkarır.
İnternete ve CRM'e bağlanmaz.

## Nasıl hesaplar?

| Öncelik | Kural |
|---|---|
| 1 · Bugün ara | Planlanan arama tarihi bugün ya da geçmiş **veya** teklif verilmiş ve son temastan `--takip` (3) gün geçmiş |
| 2 | Test sürüşü yapılmış, teklif verilmemiş ve son temastan `--takip` gün geçmiş |
| 3 | Diğer açık fırsatlar (ileri tarihli planlı arama ya da yeni temas) |

Aynı öncelikte yüksek teklif tutarı ve uzun süredir temas edilmeyen müşteri üstte yer alır. Son temas, ilk temas / test
sürüşü / teklif / son temas tarihlerinin en yenisidir.

| Gösterge | Formül |
|---|---|
| Test sürüşü % / teklif % | Test sürüşü (teklif) yapılan fırsat ÷ fırsat × 100 |
| Kapanış % | Satış ÷ (satış + kayıp) × 100 (açık fırsatlar girmez) |
| Test sürüşü yapanlarda kapanış % | Test sürüşü yapıp sonuçlanan fırsatlarda satış oranı |
| Ortalama satış süresi | İlk temas → satış tarihi (gün) |

Durum metinden tanınır: Açık / Takipte / Sıcak / Soğuk (boş = Açık) · Satış / Satıldı / Kazanıldı · Kayıp / Vazgeçti / Rakip.

**Uyarılar:** planlanan arama `--takip` günden fazla gecikmiş (Orta) · son temastan `--soguma` (30) günden fazla geçmiş açık fırsat ·
satış tarihi ya da kayıp nedeni yazılmamış kayıt (Bilgi).

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                          # örnek veri (14.10.2026)
python main.py --firsatlar firsatlar.xlsx
python main.py --firsatlar firsatlar.xlsx --tarih 15.10.2026 --takip 2 --soguma 21
```

Zorunlu sütunlar: Müşteri ve Durum. Tarih sütunlarına tarih yazın ("Evet" değil).

## Çıktı

`cikti/test_surusu_teklif_takibi.xlsx`: `Geri Dönüş Listesi` (öncelik 1 kırmızı, 2 sarı), `Danışman`, `Model`, `Kayıp Nedenleri`,
`Uyarılar`.

## Dikkat

- **Öncelik kuralları basit ve şeffaftır;** satış olasılığı tahmini değildir. Takip süresini showroom politikanıza göre ayarlayın.
- **Kişisel veri içerir.** Müşteri adı ve telefon KVKK kapsamındadır; aramalar mevcut müşteri ilişkisi ve aydınlatma kapsamında
  yapılmalı, ticari elektronik ileti (SMS / e-posta) için ayrıca onay ve İYS kaydı gerekir.
- **Danışman karşılaştırması** fırsat sayısı az olduğunda yanıltıcı olabilir; kapanış oranını yalnız performans değerlendirmesi
  için kullanmayın.
- **Örnek veri:** Müşteriler, danışmanlar, modeller ve tutarlar kurgusaldır.

## Testler

- Öncelik kuralları ve sıralama, test sürüşü / teklif / kapanış oranları, test sürüşü yapanlarda kapanış, satış süresi, kayıp nedenleri.
- Kaçırılan arama, soğuyan fırsat ve eksik kayıt uyarıları; durum tanıma, okuma ve komut satırı.

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
