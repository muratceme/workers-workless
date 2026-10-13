# Satış Primi Hesaplama · Kod Bloğu

> Satış › Satış Müdürü · Workers / Workless

Temsilcilerin hedef ve gerçekleşen satışından **gerçekleşme yüzdesini ve brüt primi** hesaplar. Prim dilimlerini
ve oranlarını sizin dosyanızdan okur; her temsilcinin priminin hangi dilimden geldiğini döker. İnternete bağlanmaz.

## Nasıl hesaplar?

| Kavram | Formül |
|---|---|
| Gerçekleşme % | Gerçekleşen satış ÷ hedef × 100 |
| Kademe yöntemi (varsayılan) | Gerçekleşme yüzdesinin düştüğü dilimin oranı **satışın tamamına** uygulanır |
| Dilimli yöntem | Her dilime düşen satış kendi oranıyla primlendirilir; dilime düşen satış = min(gerçekleşen, hedef × üst sınır) − hedef × alt sınır |
| Tahsilat esası | Prim × min(tahsilat ÷ gerçekleşen satış, 1) |
| Tavan | Prim `--tavan` tutarını aşamaz |

**Örnek** (dilimler: %80'den %1, %100'den %1,5, %120'den %2): hedef 1.000.000, satış 1.250.000 → gerçekleşme
**%125**.

- Kademe: 1.250.000 × %2 = **25.000,00**
- Dilimli: 200.000 × %1 + 200.000 × %1,5 + 50.000 × %2 = 2.000 + 3.000 + 1.000 = **6.000,00**

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                            # örnek veri, kademe yöntemi
python main.py --gerceklesme gerceklesme.xlsx --dilimler dilimler.xlsx
python main.py --gerceklesme gerceklesme.xlsx --dilimler dilimler.xlsx --yontem dilimli --esas tahsilat --tavan 20000
```

Zorunlu sütunlar: gerçekleşme dosyasında Temsilci ve Gerçekleşen; dilim dosyasında Alt Sınır ve Prim Oranı. En düşük
alt sınırın altında kalan gerçekleşme prim almaz. Oranlarda virgül ondalık ayırıcıdır (`1,5` = %1,5).

## Çıktı

`cikti/satis_primi.xlsx`:

- `Primler`: hedef, satış, gerçekleşme, kademe, hesaplanan prim, tahsilat oranı, brüt prim ve toplam.
- `Dilim Dökümü`: her temsilcinin primi hangi dilimden, hangi tutar ve oranla geldi.
- `Prim Sistemi`: hesapta kullanılan dilimler (denetim izi).
- `Uyarılar`: hedefi olmayan temsilci, tavan uygulanan prim, eksik tahsilat.

## Dikkat

- **Dilimler ve oranlar sizin prim sisteminizdir.** Örnek dosyadaki oranlar yalnız gösterim içindir; bir öneri ya
  da piyasa ortalaması değildir. Hesap, yazılı prim yönetmeliğinizle ya da sözleşmenizle aynı olmalıdır.
- **Sonuç brüt primdir.** Prim ücret niteliğindedir ve bordroya girer; sigorta primi, gelir vergisi ve damga
  vergisi kesintileri bu pakette hesaplanmaz. Net tutar için bordro uzmanınıza danışın.
- **Kademe yönteminde eşik etkisi vardır.** Gerçekleşme %99,9'dan %100'e çıkınca oran satışın tamamına uygulandığı
  için prim bir anda sıçrar. Dilimli yöntemde bu olmaz. Hangisinin uygulanacağı sisteminizde yazılı olmalıdır.
- **Tahsilat esası tek bir kural uygular:** prim, satışın tahsil edilen kısmı oranında ödenir. Vadesi gelmemiş
  alacak, iade, iskonto ve şüpheli alacak için sisteminizde başka kural varsa tutarları buna göre hazırlayın.
- **Kapsam dışı:** ürün ya da kârlılık bazlı farklı oranlar, ekip ve yönetici primi, yeni müşteri primi, dönemler
  arası mahsup ve geri alma (clawback), kıdeme göre farklı sistem.
- **Kişisel veri:** Kişi bazında prim tutarları yer alır; raporu yalnız yetkili kişilerle paylaşın.
- **Örnek veri:** Kişiler, hedefler ve satışlar kurgusaldır.

## Testler

- Kademe sınırları (dilimin tam alt sınırı ve bir kuruş altı), dilimli yöntemde dilimlere dağılım ve toplam.
- Tahsilat esası (kısmi, tam, satıştan büyük, boş), tavan, hedefi olmayan temsilci, sıfır satış.
- Dilim dosyası doğrulaması (sırasız, tekrarlı, eksi), toplamlar, okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
