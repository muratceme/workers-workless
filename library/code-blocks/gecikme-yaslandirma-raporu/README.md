# Gecikme Yaşlandırma Raporu · Kod Bloğu

> Krediler İzleme ve Takip › Kredi İzleme Uzmanı · Workers / Workless

Kredileri gecikme gün aralıklarına (**Güncel, 1-30, 31-60, 61-90, 90+**) dağıtır; şube, ürün ve müşteri bazında raporlar.
Önceki dönem dosyası verilirse aralıklar arası geçişi de gösterir. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Hesap |
|---|---|
| Gecikme günü | Rapor tarihi − ödenmemiş en eski taksitin vadesi ("Gecikme Gün" sütunu doluysa o kullanılır) |
| Aralık | 1-30, 31-60, 61-90, 90+ (sınırlar dahil: 30. gün 1-30 aralığındadır) |
| Pay % | Aralıktaki risk bakiyesi ÷ toplam risk bakiyesi |
| PAR30 / PAR90 | 30 (90) günden fazla gecikmiş kredilerin **tüm risk bakiyesi** ÷ toplam risk bakiyesi |
| Müşteri bazı | Müşterinin tüm bakiyesi, en çok gecikmiş kredisinin aralığında gösterilir |
| Geçiş matrisi | Önceki dönemde bir aralıkta olan bakiyenin bu dönem hangi aralığa geçtiği |

- **PAR** yalnız geciken taksiti değil kredinin tamamını riskli sayar. "Geciken Tutar" sütunu ayrıca toplanır.
- **Müşteri bazı:** Bir kredisi 90+ gün gecikmiş müşterinin güncel görünen diğer kredileri "Daha İyi Aralıktaki Bakiye"
  sütununda gösterilir.
- **Geçiş matrisi** önceki dönemin bakiyesiyle ölçülür. Köşegenin sağı kötüleşme, solu iyileşmedir. Bu dönem listede
  olmayan krediler "Kapandı" sütununa düşer; yeni krediler ayrı satırdadır.
- Şube ve ürün tablolarında PAR30'u portföy ortalamasının 1,5 katını aşanlar kırmızı işaretlenir.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek veri
python main.py --krediler krediler.xlsx --tarih 30.09.2026
python main.py --krediler krediler.xlsx --onceki agustos.xlsx --tarih 30.09.2026 --onceki-tarih 31.08.2026
python main.py --krediler krediler.xlsx --araliklar 30,60,90,180
```

| Sütun | Açıklama |
|---|---|
| Kredi No, Risk Bakiyesi | Zorunlu |
| Müşteri No, Müşteri, Şube, Ürün | Kırılımlar için |
| Geciken Tutar | Vadesi geçmiş ödenmemiş tutar |
| En Eski Vade Tarihi | Ödenmemiş en eski taksitin vadesi; boşsa kredi güncel sayılır |
| Gecikme Gün | Doluysa tarih yerine bu kullanılır |

## Çıktı

`cikti/gecikme_yaslandirma.xlsx`:

- `Yaşlandırma`: aralık bazında adet, bakiye, pay; müşteri bazında dağılım; PAR oranları (varsa önceki dönemle).
- `Şube Bazında`, `Ürün Bazında`: grup × aralık bakiye, PAR30 ve PAR90.
- `Gecikmeli Krediler`: gecikme gününe göre sıralı liste; boş "Aksiyon / Not" sütunu.
- `Müşteri Bazında`: gecikmesi olan müşteriler.
- `Geçiş Matrisi`, `Aralık Değiştirenler`: önceki dönem dosyası verildiyse.
- `Uyarılar`: veri tutarsızlıkları.

## Dikkat

- **Bu rapor bir sınıflandırma kararı değildir.** Kredilerin Sınıflandırılması ve Bunlar İçin Ayrılacak Karşılıklara
  İlişkin Usul ve Esaslar Hakkında Yönetmelik'te 30 günden fazla gecikme yakın izleme, 90 günden fazla gecikme donuk
  alacak göstergesidir; sınıflandırma ayrıca borçlunun ödeme gücü, yeniden yapılandırma gibi ölçütlere bakar.
  Karşılık hesabı bu paketin kapsamında değildir.
- **Gecikme günü tanımı:** Bankanızın gün sayma kuralı (tatil, kısmi ödeme, ödeme tahsis sırası) farklı olabilir. En
  doğrusu sistemden gelen "Gecikme Gün" sütununu kullanmaktır.
- **Risk bakiyesi tanımı** (anapara, faiz reeskontu, gayrinakdi) sizin dosyanıza bağlıdır; kod yorumlamaz.
- **Kişisel veri:** Müşteri bilgisi içeren çıktıyı yalnız yetkili kişilerle paylaşın.
- **Örnek veri:** Krediler ve müşteriler kurgusaldır; oranlar gerçek bir portföyü yansıtmaz.

## Testler

- Aralık sınırları (30 / 31, 90 / 91), özel aralıklar.
- PAR ve pay hesapları, şube ve ürün kırılımı, müşteri bazı.
- Geçiş matrisi: kötüleşen, iyileşen, kapanan, yeni kredi.
- Okuma hataları ve tutarsızlık uyarıları; örnek veride toplamların tutması.

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
