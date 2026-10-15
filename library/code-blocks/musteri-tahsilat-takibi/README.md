# Müşteri Tahsilat Takibi · Kod Bloğu

> Teknik ve Operasyon › Acente Yöneticisi · Workers / Workless

Müşterilerden tahsil edilecek prim taksitlerini **vadeye göre**, sigorta şirketine **aktarılacak tutarlarla
birlikte** izler: kimden ne alacağınızı, hangi şirkete ne zaman ne ödeyeceğinizi ve ikisi arasındaki açığı gösterir.
İnternete bağlanmaz.

## Nasıl hesaplar?

| Kavram | Kural |
|---|---|
| Müşteriden kalan | Tutar − tahsil edilen |
| Gecikmiş | Kalan > 0 ve vade durum tarihinden önce; yaşlandırma 1–30, 31–60, 61–90, 90+ gün |
| Yaklaşıyor | Kalan > 0 ve vadeye `--gun` (7) gün veya az kaldı |
| Şirkete kalan borç | Şirkete aktarılacak tutar − şirkete ödenen (aktarılacak boşsa taksit tutarı) |
| Vadesi geçmiş şirket borcu | Şirkete ödeme vadesi durum tarihinden önce olan kalan borç |
| Tahsil edilip aktarılmayan | min(tahsil edilen, aktarılacak tutar) − şirkete ödenen (sıfırdan büyükse) |
| Acentenin finanse ettiği | Şirkete ödenen − tahsil edilen (sıfırdan büyükse) |

Uyarılar: müşteriden tahsil edilmiş prim, şirket vadesi geçtiği hâlde aktarılmamışsa (yüksek); müşteri ödemeden
şirket vadesi geçmişse; şirkete müşteriden tahsil edilenden fazlası ödenmişse; müşterinin gecikmiş taksiti varsa
(en eski gecikme `--gecikme`, 30 gün ve üstündeyse yüksek).

**Örnek:** Taksit 8.000, müşteriden tamamı tahsil edildi; şirkete aktarılacak 6.800, şirkete ödenen 0, şirkete
ödeme vadesi 05.10.2026. Durum tarihi 12.10.2026 → **6.800 tahsil edilip aktarılmamış**, vadesi geçmiş.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                          # örnek veri, durum tarihi 12.10.2026
python main.py --taksitler taksitler.xlsx
python main.py --taksitler taksitler.xlsx --gun 15 --gecikme 45 --tarih 31.10.2026
```

Zorunlu sütunlar: Müşteri, Vade, Tutar. Her satır bir taksittir.

## Çıktı

`cikti/musteri_tahsilat_takibi.xlsx`:

- `Taksitler`: gecikmişler üstte; müşteri ve şirket tarafı yan yana.
- `Özet`: müşteri alacağı ve şirket borcu toplamları, gecikmiş alacağın yaşlandırması.
- `Müşteriler`: kalan, gecikmiş tutar, en uzun gecikme, yaklaşan vade.
- `Şirketler`: kalan borç, vadesi geçmiş borç, yakında ödenecek, aktarılmayan tahsilat, finanse edilen tutar.
- `Uyarılar`.

## Dikkat

- **Aktarılacak tutarı siz yazarsınız.** Şirkete ödenecek tutar genellikle komisyon düşülmüş net tutardır, ancak
  ödeme yöntemi (şirketin doğrudan tahsilatı, kredi kartı, acente tahsilatı) ve sözleşmeniz bunu değiştirir. Kod
  komisyon hesaplamaz; sütun boşsa taksit tutarının tamamını borç sayar.
- **Müşteri doğrudan şirkete ödüyorsa bu satırlar acente borcu değildir.** Şirketin kendi tahsil ettiği
  (kredi kartı, otomatik ödeme) taksitleri dosyaya koymayın ya da şirket tarafı sütunlarını sıfır girin.
- **Tahsil edilen primin aktarımı mevzuat ve sözleşme konusudur.** Aktarım süreleri ve yetkiler acentelik
  sözleşmenizle ve ilgili düzenlemelerle belirlenir; kod yalnız yazdığınız ödeme vadesine göre gecikmeyi gösterir.
- **Cari hesap mutabakatı değildir.** Şirket ekstresiyle karşılaştırma, iptal ve iade zeyilleri, komisyon
  mahsubu ve kur farkları kapsam dışıdır.
- **Fazla tahsilat mahsup edilmez.** Bir taksitte tutardan fazla tahsilat yazılmışsa uyarı verilir, fazlası başka
  taksite aktarılmaz.
- **Örnek veri:** Müşteriler, şirketler ve tutarlar kurgusaldır.

## Testler

- Vade sınırları (bugün vadeli taksit gecikmiş sayılmaz), yaklaşan vade, yaşlandırma dilimleri.
- Şirket borcu, vadesi geçmiş borç, tahsil edilip aktarılmayan ve acentenin finanse ettiği tutarlar.
- Müşteri ve şirket kırılımı, gecikme eşiğine göre uyarı düzeyi, okuma ve komut satırı.

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
