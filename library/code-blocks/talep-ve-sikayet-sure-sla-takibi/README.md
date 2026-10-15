# Talep ve Şikâyet Süre (SLA) Takibi · Kod Bloğu

> Müşteri Hizmetleri › Müşteri Temsilcisi · Workers / Workless

**Açık talep ve şikâyetleri hedef cevap ve çözüm sürelerine göre izler**; süresi geçen, bugün dolan ve yaklaşanları
sorumlu ve konu bazında listeler. İnternete bağlanmaz; talep listesini ve süre kurallarını siz verirsiniz.

## Nasıl hesaplar?

**Hedef süre:** SLA dosyasındaki kural. Arama sırası: konu + öncelik → yalnız konu → yalnız öncelik → genel satır
(konu ve öncelik boş). Hiçbiri yoksa komut satırı varsayılanı kullanılır (`--cevap-gun` 1, `--cozum-gun` 5).
Bir kuralda ilk cevap süresi boşsa ilk cevap için sıradaki kurala geçilir.

- **Cevap hedefi** = açılış tarihi + ilk cevap süresi; **çözüm hedefi** = açılış tarihi + çözüm süresi.
  `--is-gunu` verilirse süre iş günü sayılır: cumartesi ve pazar atlanır, resmi tatiller atlanmaz.
- **İlk cevap:** "İlk Cevap Tarihi"; boşsa ve talep kapanmışsa kapanış tarihi ilk cevap sayılır.

| Durum | Koşul |
|---|---|
| Gecikmiş | Kapanmamış ve rapor tarihi > çözüm hedefi. Gecikme = rapor tarihi − çözüm hedefi |
| Bugün son gün | Kapanmamış ve rapor tarihi = çözüm hedefi |
| Yaklaşıyor | Kapanmamış ve kalan gün ≤ `--uyari-gun` (varsayılan 1) |
| Süresi var | Kapanmamış, kalan gün daha fazla |
| Geç çözüldü | Kapanış tarihi > çözüm hedefi. Gecikme = kapanış − çözüm hedefi |
| Zamanında çözüldü | Kapanış tarihi ≤ çözüm hedefi |

- **İlk cevap durumu:** Zamanında / Geç (cevap verilmiş) ya da Bekliyor / Gecikmiş (cevap yok).
- **Çözüm süresi** = kapanış − açılış; **zamanında çözüm %** = zamanında çözülen ÷ kapanan × 100.
- Durum sütununda "iptal" geçen kayıtlar değerlendirilmez; sayısı rapora yazılır.

**Örnek:** açılış 08.10.2026, öncelik Normal, kural 1 / 5 gün → cevap hedefi 09.10, çözüm hedefi 13.10.2026; rapor
tarihi 15.10.2026, kapanmamış → Gecikmiş, 2 gün.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                          # örnek veri (rapor tarihi 15.10.2026)
python main.py --talepler talepler.xlsx --sla sla_kurallari.xlsx
python main.py --talepler talepler.xlsx --cevap-gun 1 --cozum-gun 3 --is-gunu --uyari-gun 2
```

| Dosya | Sütunlar |
|---|---|
| Talepler | Talep No, Açılış Tarihi (zorunlu); Kanal, Konu, Öncelik, Sorumlu, İlk Cevap Tarihi, Kapanış Tarihi, Durum |
| SLA kuralları (isteğe bağlı) | Konu, Öncelik, İlk Cevap (gün), Çözüm (gün). Konu ve öncelik boş bırakılabilir |

## Çıktı

`cikti/talep_sla_takibi.xlsx`:

- `Açık Talepler`: en çok gecikenden başlayarak kapanmamış talepler; yaş, cevap ve çözüm hedefi, kalan gün,
  uygulanan kural; boş "Not" sütunu.
- `Sorumlu Özeti`, `Konu Özeti`: durum sayıları, kapanan, zamanında çözüm oranı, ortalama çözüm süresi, ilk cevabı
  geç kalan talep sayısı, en uzun gecikme.
- `Kapananlar`: zamanında ve geç çözülen talepler, çözüm süresi.
- `Uyarılar`: okunamayan veya açılıştan önceki tarihler, kapalı görünüp kapanış tarihi olmayan kayıtlar, mükerrer
  talep numaraları ve kurallar, kuralı tanımlı olmayan konu/öncelikler.

## Dikkat

- **Süreleri siz verirsiniz.** Cevap ve çözüm süreleri kurumunuzun, sözleşmenin veya ilgili mevzuatın kuralıdır;
  kod hiçbir yasal süreyi bilmez ve varsaymaz. Mevzuatla belirlenmiş bir süre varsa (ör. kurum dışı başvurular,
  tüketici veya hasta başvuruları) güncel resmî metinden kontrol edip SLA dosyasına siz yazın. Örnek kurallar
  kurgusaldır.
- **Gün düzeyinde çalışır.** Saat dikkate alınmaz; "4 saat içinde cevap" gibi saatlik hedefler hesaplanmaz.
- **Resmi tatiller atlanmaz.** `--is-gunu` yalnız hafta sonunu atlar.
- **Kapanış tarihi esas alınır.** Durum "Kapandı" olsa da kapanış tarihi boşsa talep açık sayılır ve uyarı verilir.
- **Kişisel veri:** Müşteri, hasta veya misafir adı, iletişim bilgisi ve şikâyet metnine gerek yoktur; dosyaya
  koymayın. Hasta şikâyetlerinde sağlık bilgisi özel nitelikli kişisel veridir.
- **Örnek veri:** Talepler, kişiler ve süre kuralları kurgusaldır.

## Testler

- Açık talep sınırları (son gün, bir gün önce/sonra, uyarı günü), ilk cevap durumları, kapanan talepler.
- SLA kural arama sırası, boş cevap süresinde sıradaki kurala geçiş, iş günü hesabı.
- Sorumlu özeti, sıralama, iptal edilenlerin dışarıda bırakılması.
- Okuma hataları, uyarılar ve örnek veri uçtan uca.

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
