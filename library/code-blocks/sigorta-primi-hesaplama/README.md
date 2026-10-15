# Sigorta Primi Hesaplama · Kod Bloğu

> Teknik (Risk Kabul / Underwriting) › Teknik Uzman (Underwriter) · Workers / Workless

**Sizin verdiğiniz** tarife fiyatları ve indirim-sürprim tablosuyla teminat bazında prim hesaplar; teklif bazında
toplar. İnternete bağlanmaz. Pakette gömülü tarife yoktur.

## Nasıl hesaplar?

```
Baz prim (yıllık) = bedel × fiyat (‰) ÷ 1000
Süre katsayısı    = gün ÷ 365            (başlangıçtan tam bir yıl sonra biten poliçede 1)
Çarpan            = zincirleme: (1 + sürprim₁) × … × (1 − indirim₁) × …
                    toplamsal : 1 + Σ sürprim − Σ indirim
Hesaplanan prim   = baz prim × süre katsayısı × çarpan      (kuruşa yuvarlanır)
Net prim          = hesaplanan prim; teminatın asgari priminin altındaysa asgari prim
Vergi             = teklifin net primi × --vergi %
```

- **Tarife eşleşmesi:** önce teminat + risk sınıfı aranır; yoksa teminatın `*` (tüm sınıflar) satırı kullanılır.
  İkisi de yoksa prim hesaplanmaz ve uyarı verilir.
- **İndirim-sürprim:** satırdaki kodlar tablodan okunur. Kod tabloda yoksa veya o teminata uygulanmıyorsa atlanır ve
  not düşülür. Aynı kod iki kez yazılmışsa bir kez uygulanır.
- **İndirim tavanı:** `--azami-indirim 30` verilirse çarpan 0,70'in altına inemez; tavan uygulanan satır işaretlenir.
- **Örnek:** 65.000.000 TL × 0,90 ‰ = 58.500 TL; %15 ve %10 indirim zincirleme → 58.500 × 0,85 × 0,90 = 44.752,50 TL.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek veri
python main.py --riskler riskler.xlsx --tarife tarife.xlsx --indirimler indirim_surprim.xlsx --vergi 5
python main.py --yontem toplamsal --azami-indirim 30
```

| Dosya | Sütunlar |
|---|---|
| Riskler | Teklif No, Sigortalı, Risk Sınıfı, Teminat, Bedel, Başlangıç, Bitiş (veya Süre gün), İndirim / Sürprim (`SPR\|HSZ`) |
| Tarife | Teminat, Risk Sınıfı (`*` = hepsi), Fiyat (‰), Asgari Prim |
| İndirim-sürprim | Kod, Ad, Tür (`İndirim` / `Sürprim`), Oran %, Teminat (`*` = hepsi; birden çok teminat virgülle) |

Tarih verilmezse süre 365 gün sayılır.

## Çıktı

`cikti/prim_hesabi.xlsx`:

- `Teklif Özeti`: teklif bazında net prim, vergi, brüt prim; hesaplanamayan teminat sayısı.
- `Teminat Primleri`: bedel, fiyat, baz prim, süre katsayısı, uygulanan indirim-sürprimler, çarpan, net prim ve notlar.
- `Uyarılar`: tarifesi bulunmayan teminatlar, tanınmayan kodlar, okunamayan satırlar.

## Dikkat

- **Tarife sizindir.** Fiyatlar, risk sınıfları, asgari primler ve indirim oranları örnek dosyalarda kurgusaldır;
  şirketinizin onaylı tarifesini ve yetki limitlerini kullanın.
- **Kısa süre tarifesi uygulanmaz.** Bir yıldan kısa poliçelerde prim gün esasıyla (gün ÷ 365) hesaplanır; şirketiniz
  kısa süre tablosu kullanıyorsa sonuç farklı olur. Asgari prim süreye göre azaltılmaz.
- **Vergi varsayılanı 0'dır.** Net prime eklenecek vergi (ör. BSMV) oranını siz verirsiniz; oran ve istisnalar branşa
  göre değişir. Kod branş bazında ayrı vergi oranı uygulamaz.
- **Zorunlu tarifeler** (ör. zorunlu deprem, trafik) kendi mevzuat tarifesiyle hesaplanır; bu paket onların yerine geçmez.
- **Risk kabul kararı değildir.** Kod reasürans, kümül, yetki limiti ve tarife dışı fiyat onayı kontrolü yapmaz.
- **Örnek veri:** Sigortalılar, bedeller ve fiyatlar kurgusaldır.

## Testler

- Baz prim, `*` sınıfı, tarifede olmayan sınıf.
- Zincirleme ve toplamsal indirim-sürprim, tekrar eden kod, uygulanmayan ve tanınmayan kod, indirim tavanı.
- Gün esaslı süre katsayısı, asgari prim.
- Teklif toplamı ve vergi; okuma hataları; örnek veri uçtan uca.

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
