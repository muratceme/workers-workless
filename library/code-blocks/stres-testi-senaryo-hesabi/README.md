# Stres Testi Senaryo Hesabı · Kod Bloğu

> İç Kontrol ve Risk Yönetimi › Risk Yönetimi Uzmanı · Workers / Workless

Kur, faiz ve takip oranı şoklarının **kârlılık ve sermaye yeterliliği** üzerindeki etkisini senaryo bazında hesaplar.
Bankanın kendi duyarlılık rakamlarını senaryolara uygulayan basit ve izlenebilir bir modeldir. İnternete bağlanmaz.

## Nasıl hesaplar?

| Adım | Formül |
|---|---|
| Kur etkisi | Net yabancı para pozisyonu × kur şoku % |
| Faiz etkisi | (Net faiz geliri duyarlılığı + değerleme duyarlılığı) × şok (bp) ÷ 100 |
| Ek takip | Toplam krediler × takip oranı artışı (puan) ÷ 100 |
| Kredi zararı | Ek takip × karşılık oranı % |
| Toplam şok etkisi | Kur etkisi + faiz etkisi − kredi zararı |
| Stres sonrası kâr | Baz yıllık kâr + toplam şok etkisi |
| Sermaye etkisi | min(0, stres sonrası kâr) |
| Stres sonrası RAV | RAV + yabancı para RAV × kur şoku % |
| Sermaye yeterliliği | (Özkaynak + sermaye etkisi) ÷ stres sonrası RAV |
| Hedef için ek özkaynak | max(0, hedef oran × stres sonrası RAV − stres sonrası özkaynak) |

- **Kâr tamponu:** Zarar önce yılın beklenen kârından karşılanır; yalnız kârı aşan kısım sermayeden düşer.
  `--kar-tamponu-yok` ile şok etkisi doğrudan sermayeden düşülür.
- **İhtiyatlılık:** Olumlu etki (ör. uzun pozisyonda kur kazancı) sermayeye eklenmez. Karşılık ayrılan tutar RAV'dan
  düşülmez. Vergi etkisi hesaplanmaz.
- **Ters stres:** Diğer şoklar sıfırken, tek başına hangi büyüklükte kur / faiz / takip şokunun oranı asgariye
  indirdiği ikiye bölme yöntemiyle bulunur.
- **Durum:** oran asgarinin (`--asgari`, varsayılan %8) altındaysa "Asgarinin altında", hedefin (`--hedef`, varsayılan
  %12) altındaysa "Hedefin altında".

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek veri
python main.py --parametreler banka.csv --senaryolar senaryolar.csv --asgari 8 --hedef 12
python main.py --parametreler banka.csv --senaryolar senaryolar.csv --kar-tamponu-yok
```

**Parametreler** (`Alan;Değer` satırları):

| Alan | Açıklama |
|---|---|
| Özkaynak, Risk Ağırlıklı Varlıklar, Toplam Krediler | Zorunlu |
| Yabancı Para RAV | RAV içindeki yabancı para cinsi kısmın TL karşılığı |
| Net Yabancı Para Pozisyonu | TL karşılığı; uzun artı, kısa eksi |
| Net Faiz Geliri Duyarlılığı | Faizler 100 baz puan artarsa 12 aylık net faiz geliri değişimi |
| Değerleme Duyarlılığı | Faizler 100 baz puan artarsa menkul kıymet değerleme farkı |
| Takipteki Krediler, Karşılık Oranı, Baz Yıllık Kâr | Girilmezse 0 / %60 / 0 alınır ve uyarı verilir |

**Senaryolar:** Senaryo, Kur Şoku (%), Faiz Şoku (bp), Takip Oranı Artışı (puan), Karşılık Oranı (%; boşsa parametredeki),
Açıklama.

## Çıktı

`cikti/stres_testi.xlsx`:

- `Senaryo Sonuçları`: şok etkisi, stres sonrası kâr, özkaynak, RAV, sermaye yeterliliği, durum, ek özkaynak ihtiyacı.
- `Etki Dökümü`: kur, faiz (iki bileşen) ve kredi zararının ayrı ayrı etkisi; RAV değişimi.
- `Ters Stres`: oranı asgariye indiren tekil şok büyüklükleri.
- `Parametreler`, `Uyarılar`.

## Dikkat

- **Doğrusal duyarlılık modelidir.** Büyük şoklarda duyarlılık sabit kalmaz; şoklar arasındaki etkileşim, ikinci tur
  etkiler (mevduat kaçışı, teminat değer kaybı, kur artışının YP kredi kalitesine etkisi) modellenmez.
- **Yasal raporlamanın yerini tutmaz.** Stres testi ve içsel sermaye yeterliliği değerlendirmesi (İSEDES) BDDK
  düzenlemelerine ve bankanızın onaylı metodolojisine göre yapılır.
- **Asgari ve hedef oranlar** parametredir. Sermaye tamponları ve çekirdek / ana sermaye oranları kapsam dışıdır;
  güncel yasal oranları kontrol edin.
- **Girdiler sizin sorumluluğunuzdadır.** Duyarlılık rakamları aktif-pasif ve piyasa riski sistemlerinizden alınmalıdır.
- **Örnek veri:** Banka büyüklükleri ve senaryolar kurgusaldır.

## Testler

- Birleşik senaryoda her etki elle hesaplanan değerle; kâr tamponu açık ve kapalı.
- Olumlu etkinin sermayeye eklenmemesi; senaryoya özel karşılık oranı.
- Durum sınıfları ve ek özkaynak ihtiyacı; ters stres değerleri (kapalı çözümle karşılaştırılmış).
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
