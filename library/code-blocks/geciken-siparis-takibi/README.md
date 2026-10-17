# Geciken Sipariş Takibi · Kod Bloğu

> E-ticaret Operasyon ve Kargo › E-ticaret Operasyon Uzmanı · Gerisi Bende

**Kargoya verilmesi gereken süreyi aşan siparişleri** kanal ve depo bazında listeler. İnternete bağlanmaz; sipariş
listesini siz verirsiniz.

## Nasıl hesaplar?

**Termin:** dosyadaki "Termin" tarihi doluysa o kullanılır. Boşsa sipariş tarihi + kanalın kargoya verme süresi
(kanal süre dosyası; kanal orada yoksa `--sure-gun`, varsayılan 2 gün). `--is-gunu` verilirse süre iş günü sayılır:
cumartesi ve pazar atlanır, resmi tatiller atlanmaz.

| Durum | Koşul |
|---|---|
| Gecikmiş | Kargoya verilmemiş ve rapor tarihi > termin. Gecikme = rapor tarihi − termin |
| Bugün son gün | Kargoya verilmemiş ve rapor tarihi = termin |
| Süresi var | Kargoya verilmemiş ve rapor tarihi < termin |
| Geç kargolandı | Kargoya veriliş tarihi > termin. Gecikme = kargo tarihi − termin |
| Zamanında kargolandı | Kargoya veriliş tarihi ≤ termin |

- **Zamanında kargolama %** = zamanında kargolanan ÷ kargolanan × 100.
- **Ortalama gecikme** yalnız hâlâ bekleyen gecikmiş siparişlerden hesaplanır. Gecikme takvim günüdür.
- Durum sütununda "iptal" geçen siparişler değerlendirilmez; sayısı rapora yazılır.

**Örnek:** sipariş 07.10.2026, kanal süresi 2 gün → termin 09.10.2026; rapor tarihi 12.10.2026, kargoya verilmemiş →
Gecikmiş, 3 gün.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri (rapor tarihi 12.10.2026)
python main.py --siparisler siparisler.xlsx --kanal-sure kanal_sureleri.xlsx
python main.py --siparisler siparisler.xlsx --sure-gun 1 --is-gunu --tarih 15.10.2026
```

| Dosya | Sütunlar |
|---|---|
| Siparişler | Sipariş No, Sipariş Tarihi (zorunlu); Kanal, Depo, Termin, Kargoya Veriliş Tarihi, Durum |
| Kanal süreleri (isteğe bağlı) | Kanal, Süre (gün) |

## Çıktı

`cikti/geciken_siparisler.xlsx`:

- `Bekleyen Siparişler`: en çok gecikenden başlayarak kargoya verilmemiş siparişler; kalan gün, gecikme; boş "Not" sütunu.
- `Kanal Özeti`, `Depo Özeti`: durum sayıları, zamanında kargolama oranı, ortalama ve en uzun gecikme.
- `Kargolananlar`: zamanında ve geç kargolanan siparişler.
- `Uyarılar`: okunamayan tarihler, mükerrer sipariş numaraları, süresi tanımlı olmayan kanallar.

## Dikkat

- **Süreleri siz verirsiniz.** Kargoya verme süresi her pazaryerinin kendi kuralıdır, kategoriye ve satıcıya göre
  değişir ve zaman içinde güncellenir. Kod hiçbir kanalın süresini bilmez; örnek dosyadaki süreler kurgusaldır.
  En doğrusu kanalın sipariş listesindeki son kargo tarihini "Termin" sütununa yazmaktır.
- **Gün düzeyinde çalışır.** Saat dikkate alınmaz: "gün içinde saat 14'e kadar" gibi kesim saatleri ve saat bazlı
  süreler hesaplanmaz. Sipariş tarihindeki saat bilgisi atılır.
- **Resmi tatiller atlanmaz.** `--is-gunu` yalnız hafta sonunu atlar; bayram ve resmi tatillerde termini elle girin.
- **Kargoya veriliş, teslim değildir.** Müşteriye teslim süresi ve kargo firması gecikmeleri bu pakette yoktur.
- **Kişisel veri:** Müşteri adı, adres ve telefon sütunlarına gerek yoktur; dosyaya koymayın.
- **Örnek veri:** Kanallar, depolar, siparişler ve süreler kurgusaldır.

## Testler

- Bekleyen sipariş durumları ve sınırları (termin günü, bir gün sonrası), dosyadaki termin, kanal süresi.
- Kargolanan siparişler, iş günü hesabı (cuma ve cumartesi siparişi).
- Kanal ve depo özeti, sıralama, iptal edilenlerin dışarıda bırakılması.
- Okuma hataları ve örnek veri uçtan uca.

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
