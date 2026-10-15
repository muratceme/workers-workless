# İç Tetkik Bulgu Takibi · Kod Bloğu

> Kalite Güvence ve Kalite Kontrol › Kalite Müdürü · Workers / Workless

ISO 9001 iç tetkik bulgularını **madde, birim ve kapanış durumuna** göre takip eder. Termini geçen uygunsuzlukları, kök nedeni ya
da düzeltici faaliyeti yazılmadan kapatılanları, etkinliği doğrulanmamış faaliyetleri ve **farklı tetkiklerde tekrar eden**
bulguları listeler. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Kural |
|---|---|
| Termin | Satırdaki termin; yoksa tetkik tarihi + hedef süre (majör `--major` 30, minör `--minor` 60 gün). Gözlem ve iyileştirme fırsatında termin aranmaz |
| Gecikme | Açık bulguda bugün − termin |
| Kapanış süresi | Kapanış tarihi − tetkik tarihi (gün) |
| Tekrar eden | Aynı birim + madde, en az iki farklı tetkikte uygunsuzluk |
| Özetler | Birim, madde (madde numarasına göre sıralı) ve tetkik bazında: bulgu, majör, minör, açık, geciken, ortalama kapanış, etkin olmayan faaliyet |

Tür metinden tanınır: Majör / Büyük · Minör / Küçük / Uygunsuzluk · Gözlem · İyileştirme fırsatı / Öneri. Etkinlik: Etkin · Etkin
değil · boş (bekliyor).

**Uyarılar:** termini geçen majör uygunsuzluk · okunamayan satır (Yüksek) · termini geçen minör · sorumlusu yazılmamış açık bulgu ·
kök neden ya da düzeltici faaliyet yazılmadan kapatılan · etkinliği olumsuz · tekrar eden bulgu (Orta) · etkinliği doğrulanmamış (Bilgi).

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                          # örnek veri (3 tetkik, 14.10.2026)
python main.py --bulgular bulgular.xlsx
python main.py --bulgular bulgular.xlsx --major 15 --minor 45
```

## Çıktı

`cikti/ic_tetkik_bulgulari.xlsx`: `Açık Bulgular`, `Birim`, `Madde`, `Tetkik`, `Tekrar Eden`, `Tüm Bulgular`, `Uyarılar`.

## Dikkat

- **Hedef süreler kalite el kitabınızdan gelmelidir.** Varsayılan 30 / 60 gün yalnız örnektir. Belgelendirme kuruluşu tetkiklerindeki
  bulgular için kuruluşun verdiği süreler geçerlidir.
- **Standart metni pakette yoktur.** Madde sütunu yalnız numara olarak kullanılır; madde başlıklarını kendi tetkik formunuzdan alın.
- **Tekrar ölçütü birim + madde eşleşmesidir;** aynı maddede farklı konular da tekrar sayılabilir. Bulgu açıklamalarını okuyarak
  değerlendirin.
- **Örnek veri:** Bulgular, birimler ve tarihler kurgusaldır.

## Testler

- Varsayılan termin, gecikme, kapanış süresi, açık bulgu sıralaması, tekrar eden bulgu, madde numarası sıralaması, toplamlar.
- Uyarılar, tür ve etkinlik tanıma, okuma ve komut satırı.

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
