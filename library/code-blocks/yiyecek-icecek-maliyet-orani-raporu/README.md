# Yiyecek-İçecek Maliyet Oranı Raporu · Kod Bloğu

> Yiyecek ve İçecek › Yiyecek İçecek Müdürü · Workers / Workless

Dönem stok hareketleri ve satışlardan **departman (outlet) bazında yiyecek ve içecek maliyet oranını** hesaplar; hedef ve
önceki dönemle karşılaştırır, hedefin üstündeki departmanlar için fazla maliyeti tahmin eder. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Tüketim maliyeti | Açılış stok + alım + transfer giriş − transfer çıkış − kapanış stok |
| Satılan maliyet | Tüketim − personel yemeği − ikram (bu iki kalem satış maliyetinden çıkarılıp ilgili gider hesabına aktarılır) |
| Maliyet oranı % | Satılan maliyet ÷ net satış × 100 |
| Fark | Oran − hedef (puan); fazla maliyet ≈ fark × net satış ÷ 100 |
| Stok günü | Ortalama stok ((açılış + kapanış) ÷ 2) ÷ (tüketim ÷ dönem günü `--gun`) |
| Toplamlar | Yiyecek, içecek ve genel toplam; hedef ve önceki dönem oranı satışla ağırlıklandırılır |

Mutfaktan bara (ör. meyve) ya da bardan mutfağa (ör. yemekte kullanılan şarap) geçen malzemeyi transfer sütunlarına yazın.
Tür metinden tanınır: Yiyecek / Food / Mutfak · İçecek / Beverage / Bar.

**Uyarılar:** satış olmadan maliyet · negatif kapanış stok ya da maliyet (Yüksek) · oran hedefin `--tolerans` (varsayılan 1)
puandan fazla üstünde (Orta) · oran hedefin 5 puandan fazla altında (sayım fazlası, eksik fatura ya da porsiyon sorunu
olabilir) · departmanlar arası transfer girişi ile çıkışının tutmaması (Bilgi).

**Örnek:** açılış 100 + alım 500 − transfer çıkış 30 − kapanış 120 = 450 tüketim; personel yemeği 40 ve ikram 10 düşülünce
satılan maliyet 400; 1.500 satışla oran **%26,67**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                 # örnek veri (bir aylık dönem, 6 departman)
python main.py --donem donem.xlsx
python main.py --donem donem.xlsx --tolerans 0.5 --gun 31
```

Zorunlu sütunlar: Departman, Tür, Kapanış Stok ve Net Satış. Boş tutar sütunları sıfır sayılır.

## Çıktı

`cikti/yiyecek_icecek_maliyet_orani.xlsx`: `Maliyet Oranı` (hedefin üstü kırmızı, içinde yeşil), `Uyarılar`.

## Dikkat

- **Stok sayımı doğru olmalı.** Oran, açılış ve kapanış sayımına doğrudan bağlıdır; sayım tarihleri ve değerleme yöntemi
  (ortalama maliyet, FIFO) dönemler arasında aynı olmalıdır.
- **Satışlar KDV ve servis ücreti hariç net satış** olmalıdır; alımlar da KDV hariç yazılmalıdır.
- **Hedef oranlar işletmeye özeldir** (konsept, fiyat seviyesi, her şey dahil sistemi). Örnekteki hedefler yalnız örnektir.
- **Her şey dahil (all inclusive) otellerde** satış geliri outlet bazında oluşmadığı için oran yerine kişi başı maliyet
  izlemek daha anlamlı olabilir.
- **Örnek veri:** Departmanlar ve tutarlar kurgusaldır.

## Testler

- Tüketim ve satılan maliyet, oran, hedef farkı, önceki döneme değişim, stok günü, satış ağırlıklı toplam hedefi.
- Hedef üstü / çok altı, satış yok, negatif stok, denk olmayan transfer uyarıları; tür tanıma, okuma ve komut satırı.

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
