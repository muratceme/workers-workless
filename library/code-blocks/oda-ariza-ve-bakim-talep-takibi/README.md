# Oda Arıza ve Bakım Talep Takibi · Kod Bloğu

> Kat Hizmetleri › Kat Hizmetleri Müdürü (Housekeeper) · Workers / Workless

Odalardan bildirilen arıza ve bakım taleplerini **oda, kategori, öncelik ve bekleme süresine** göre takip eder. Hedef süresini
geçen açık talepleri, satışa kapalı odaları ve aynı odada tekrar eden arızaları listeler. İnternete ve teknik servis
yazılımına bağlanmaz; dışa aktarılmış talep listesini okur.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Süre (saat) | Tamamlanma − bildirim; açık talepte rapor zamanı − bildirim |
| Hedef süre | Önceliğe göre: Acil `--acil` (4) · Normal `--normal` (24) · Düşük `--dusuk` (72) saat |
| Hedefte % | Hedef süresi içinde tamamlanan ÷ tamamlanan × 100 |
| Kayıp oda-gece | Satışa kapalı işaretli taleplerde kapalı kalınan takvim günü (en az 1; açıksa rapor gününe kadar) |
| Tekrar eden arıza | Aynı oda + kategoride `--gun` (30) gün içinde `--tekrar` (3) ya da daha fazla talep |

Öncelik metinden tanınır: Acil / Yüksek / Kritik · Normal / Orta (boş = Normal) · Düşük / Planlı. Durum: Açık / Yeni (boş =
Açık) · Devam ediyor / Parça bekliyor · Tamamlandı / Kapandı / Çözüldü · İptal (iptaller alınmaz). Tamamlanma zamanı yazılıysa
talep tamamlanmış sayılır.

**Uyarılar:** hedef süresini geçmiş açık talep (Acil ise Yüksek, diğerleri Orta) · satışa kapalı açık oda · tekrar eden arıza
(Orta) · tamamlanma zamanı bildirimden önce (Yüksek) · tamamlanmış ama zamanı yazılmamış talep (Bilgi).

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                          # örnek veri (rapor zamanı 15.10.2026 09:00)
python main.py --talepler talepler.xlsx                                 # rapor zamanı: şimdi
python main.py --talepler talepler.xlsx --rapor-zamani "15.10.2026 09:00" --acil 2 --normal 12 --tekrar 2 --gun 60
```

Zaman `GG.AA.YYYY SS:DD` biçiminde yazılır; yalnız tarih yazılırsa saat 00:00 kabul edilir.

## Çıktı

`cikti/oda_ariza_takibi.xlsx`: `Açık Talepler` (önceliğe ve gecikmeye göre sıralı), `Kategori`, `Öncelik`, `Oda`,
`Tekrar Eden`, `Tüm Talepler`, `Uyarılar`.

## Dikkat

- **Hedef süreler otel politikasıdır.** Varsayılanlar (4 / 24 / 72 saat) yalnız başlangıç önerisidir; teknik servis
  sözleşmenize ve standartlarınıza göre değiştirin.
- **Misafir güvenliğini etkileyen arızalar** (elektrik kaçağı, gaz, su baskını, kilit) öncelik sütununda Acil yazılmadıkça
  Acil sayılmaz; bildirim sırasında doğru öncelik verildiğinden emin olun.
- **Kayıp oda-gece tahminidir.** Odanın gerçekten satılıp satılamayacağı doluluğa bağlıdır; PMS'teki OOO / OOS kayıtlarıyla
  karşılaştırın.
- **Örnek veri:** Odalar, talepler ve personel adları kurgusaldır.

## Testler

- Açık ve kapalı taleplerde süre, hedef aşımı, açık talep sıralaması, hedefte %, ortalama çözüm süresi, kayıp oda-gece,
  tekrar penceresi.
- Öncelik / durum / zaman tanıma, uyarılar, okuma hataları ve komut satırı.

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
