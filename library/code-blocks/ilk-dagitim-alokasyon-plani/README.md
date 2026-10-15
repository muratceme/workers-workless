# İlk Dağıtım (Alokasyon) Planı · Kod Bloğu

> Perakende › Ürün Planlama ve Alokasyon › Alokasyon Uzmanı · Workers / Workless

Yeni sezon ürünlerinin depo stoğunu mağaza potansiyeli, küme ve beden eğrisine göre mağazalara dağıtır. Depo
sistemine aktarılabilecek bir **sevk listesi** üretir. İnternete bağlanmaz.

## Nasıl dağıtır?

1. **Rezerv:** Her bedenin depo stoğunun `--rezerv` kadarı (varsayılan %30) tamamlama sevkiyatları için depoda
   kalır.
2. **Mağaza ağırlığı:** Potansiyel (geçmiş sezon satış adedi veya puan) × küme katsayısı (`--kume A=1,3 C=0,7`).
3. **Beden eğrisi:** Mağazanın kendi eğrisi kullanılır; yoksa kümesinin, o da yoksa genel (`*`) eğri. Eğri modelin
   bedenlerini içermiyorsa (ör. pantolon numaraları) depo stoğunun beden dağılımı kullanılır.
4. **Paylaştırma:** Her model-renk-beden için önce **asgari teşhir adedi** verilir (`--asgari`, varsayılan 1).
   Stok yetmiyorsa asgari adet en güçlü mağazalardan başlar. Kalan miktar ağırlık × beden payı oranında **en büyük
   kalan yöntemiyle** tam sayılara bölünür, böylece toplam her zaman dağıtılabilir stoka eşit olur.
5. **Asorti:** `--asorti S=1,M=2,L=2,XL=1` verilirse dağıtım paket sayısı üzerinden yapılır.
6. **Kapasite:** Mağaza kapasitesi aşılırsa gönderim oransal azaltılır ve fazla depoda kalır.

**Kontroller**
- **Kırık seri:** Bir model-rengi alan mağazaya ana bedenlerden biri gitmiyorsa işaretlenir. Varsayılan ana
  bedenler en küçük ve en büyük hariç olanlardır; `--ana-bedenler M,L` ile değiştirilebilir.
- **Kapasite kullanımı** ve **depoda kalan stok** raporlanır.
- **Pasif mağazalar:** `Aktif = Hayır` olan mağazalara gönderim yapılmaz.

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                                 # örnek: 3 model-renk, 9 mağaza (1'i pasif)
python main.py --stok depo_stok.xlsx --magazalar magazalar.xlsx --egri beden_egrisi.xlsx --rezerv 30
python main.py --stok depo_stok.xlsx --magazalar magazalar.xlsx --kume A=1,3 B=1 C=0,7 --asgari 2
python main.py --stok depo_stok.xlsx --magazalar magazalar.xlsx --asorti S=1,M=2,L=2,XL=1
```

| Dosya | Sütunlar |
|---|---|
| Depo stoğu | Model, Renk, [Kategori], Beden, Stok — veya Model, Renk, S, M, L, XL … |
| Mağazalar | Mağaza, Küme, Potansiyel, [Kapasite, Aktif] |
| Beden eğrisi | Küme (veya mağaza kodu ya da `*`), S, M, L … (pay veya yüzde) |

## Çıktı

`Model Özeti` · `Mağaza Özeti` · `Dağıtım Matrisi` (mağaza × model-renk-beden) · `Sevk Listesi` (depo sistemine
aktarım için) · `Kontroller`

## Dikkat

- **Potansiyel verisi:** Geçmiş sezondaki satış kayıpları (stoksuzluk) potansiyeli düşük gösterebilir. Mümkünse
  stoksuz günlerde düzeltilmiş satışı kullanın.
- **Kapasite ve asorti:** Kapasite azaltması asorti paket bütünlüğünü bozabilir. Paketle çalışıyorsanız
  kapasiteyi paket katı olarak verin.

## Testler

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
