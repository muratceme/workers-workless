# Rücu Potansiyeli Tespiti · Kod Bloğu

> Hasar Yönetimi › Rücu Uzmanı · Gerisi Bende

Ödenen hasar dosyalarını **karşı tarafın kusur oranı, kimliği ve sigortacısına** göre tarar; rücu edilebilir dosyaları,
kalan tutarları ve iç takip tarihlerini listeler. İnternete bağlanmaz.

## Nasıl hesaplar?

```
Rücu potansiyeli = ödenen tazminat × karşı taraf kusur oranı %
Kalan tutar      = rücu potansiyeli − tahsil edilen rücu        (eksiye düşmez)
Takip tarihi     = hasar tarihi + --takip-yil (varsayılan 2) yıl
```

Sigortacı, ödediği tazminat tutarıyla sınırlı olarak sigortalının yerine geçer (halefiyet); bu yüzden rücu tutarı
ödenen tazminatı aşamaz ve ödeme yapılmadan rücu doğmaz. Örnek: 80.000 TL ödeme, karşı taraf %75 kusurlu → 60.000 TL.

| Durum | Koşul (sırayla denenir) |
|---|---|
| Ödeme yapılmamış | Ödenen tazminat 0 |
| Kusur bilgisi eksik | Karşı taraf var, kusur oranı boş |
| Rücu yok | Kusur oranı 0, veya karşı taraf da kusur oranı da yok |
| Karşı taraf belirsiz | Kusur oranı > 0, karşı taraf ve sigortacısı boş |
| Tamamlandı | Kalan tutar 0 |
| Aynı şirket | Karşı tarafın sigortacısı `--sirket` ile verdiğiniz ad |
| Asgari tutarın altında | Kalan tutar < `--asgari` |
| Takip süresi geçmiş | Takip tarihi rapor tarihinden önce |
| Rücu edilebilir · acil | Takip tarihine `--acil-gun` (varsayılan 90) gün veya daha az kaldı |
| Rücu edilebilir | Diğerleri |

**Muhatap:** karşı tarafın sigortacısı girildiyse sigorta şirketi, yoksa karşı tarafın kendisi. "Muhatap Bazında"
sayfası rücu edilebilir dosyaları muhataba göre toplar (toplu talep ve mutabakat için).

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri (rapor tarihi 01.10.2026)
python main.py --hasarlar odenen_hasarlar.xlsx --tarih 31.12.2026 --asgari 5000 --sirket "Şirket Adı"
python main.py --baz odeme --takip-yil 2               # takip süresini ödeme tarihinden başlat
```

Sütunlar: Hasar No, Ödenen Tazminat (zorunlu); Poliçe No, Sigortalı, Branş, Hasar Tarihi, Ödeme Tarihi, Karşı Taraf,
Karşı Taraf Kusur %, Karşı Taraf Sigortacısı, Karşı Poliçe No, Tahsil Edilen Rücu.

## Çıktı

`cikti/rucu_potansiyeli.xlsx`:

- `Özet`: durum bazında dosya sayısı ve tutarlar; rücu edilebilir toplam.
- `Rücu Dosyaları`: potansiyel, tahsilat, kalan tutar, muhatap, takip tarihi, kalan gün, durum ve not; boş "Yapılan İşlem" sütunu.
- `Muhatap Bazında`: muhatap başına dosya sayısı, acil dosya ve toplam tutar.
- `Uyarılar`: okunamayan satırlar, potansiyeli aşan tahsilat.

## Dikkat

- **Takip tarihi zamanaşımı hesabı değildir.** Rücu alacağında zamanaşımının süresi ve başlangıcı; olayın hukuki
  niteliğine (haksız fiil, sözleşme, trafik kazası, taşıma vb.) ve somut duruma göre değişir. Buradaki süre yalnız
  dosyaların gözden kaçmaması için bir iç takip süresidir. Süre hesabını mutlaka hukuk biriminiz yapsın.
- **Kod kusur tespiti yapmaz.** Kusur oranı dosyaya ne girildiyse odur; kaynağını (tutanak, eksper raporu, bilirkişi)
  ayrıca belgeleyin.
- **Hukuki engelleri bilmez.** Rücudan feragat klozu, sigortalının yakınlarına veya çalışanlarına rücu sınırları,
  karşı poliçenin limiti ve muafiyeti, ibra belgeleri hesaba girmez.
- **Sovtaj ve faiz dahil değildir.** Yalnız ödenen tazminat üzerinden hesap yapılır.
- **Örnek veri:** Sigortalılar, karşı taraflar ve sigorta şirketleri kurgusaldır.

## Testler

- Tutar, tahsilat sonrası kalan, muhatap seçimi.
- Tüm durumlar ve öncelik sırası; asgari tutar sınırı; aynı şirket.
- Takip tarihi: acil sınırı (90. gün), süre geçmiş, ödeme bazlı hesap, 29 Şubat.
- Muhatap ve durum özetleri; okuma hataları; örnek veri uçtan uca.

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
