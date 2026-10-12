# İthalat Vergi Hesaplama · Kod Bloğu

> Gümrük Operasyon › Gümrük Operasyon Uzmanı · Workers / Workless

CIF kıymet, gümrük vergisi, ilave gümrük vergisi ve KDV oranlarından **ithalatta ödenecek vergilerin ön hesabını**
yapar. İnternete bağlanmaz. Oranları siz verirsiniz.

## Nasıl hesaplar?

| Kalem | Formül |
|---|---|
| CIF (TL) | CIF kıymet × `--kur` |
| Gümrük vergisi (GV) | CIF × GV % |
| İlave gümrük vergisi (İGV) | CIF × İGV % |
| KKDF | CIF × `--kkdf` % (yalnız vadeli ödeme şekillerinde; varsayılan 0) |
| ÖTV | (CIF + GV + İGV + KKDF + diğer vergi + damga payı + yurt içi gider payı) × ÖTV % |
| KDV | (ÖTV matrahı + ÖTV) × KDV % |
| Toplam vergi | GV + İGV + KKDF + diğer vergi + damga + ÖTV + KDV |
| KDV hariç maliyet | CIF + yurt içi gider + (toplam vergi − KDV) |

`--damga` ve `--yurt-ici-gider` tutarları kalemlere CIF oranında dağıtılır; kuruş farkı son kaleme eklenir.

**Örnek:** CIF 10.000 × kur 40 = 400.000 TL; GV %5 = 20.000; İGV %20 = 80.000; yurt içi gider payı 4.000 → KDV matrahı
504.000; KDV %20 = 100.800; toplam vergi 200.800.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri (kur 40, yurt içi gider 6.000)
python main.py --kalemler ithalat_kalemleri.xlsx --kur 41,25 --yurt-ici-gider 8500 --damga 1200 --kkdf 6
```

Sütunlar: CIF Kıymet, KDV % (zorunlu); Kalem / GTİP, Açıklama, GV %, İGV %, ÖTV %, Diğer Vergi (tutar, TL; ör. dampinge
karşı vergi). Boş bırakılan GV, İGV ve ÖTV oranı 0 sayılır; KDV oranı boş bırakılamaz.

## Çıktı

`cikti/ithalat_vergileri.xlsx`:

- `Vergi Özeti`: CIF, vergi kalemleri, toplam vergi, KDV matrahı, KDV hariç maliyet.
- `Kalem Bazında`: her kalem için oranlar, matrahlar, vergiler ve vergi yükü %.
- `Uyarılar`: okunamayan kıymet ve oranlar.

## Dikkat

- **Ön hesaptır, beyan değildir.** Beyanname gümrük müşavirince hazırlanır; gümrük idaresinin tahakkuku esastır.
- **Bu pakette kaynak teyidi yapılmamıştır.** Matrah sıralaması genel uygulamaya göre kurulmuştur; güncel mevzuatı
  ve kendi eşyanıza özgü kuralları gümrük müşavirinize doğrulatın.
- **Oranları kod bilmez.** GTİP'e göre vergi oranı, menşe, tercihli rejim (A.TR, EUR.1, menşe şahadetnamesi), ek mali
  yükümlülük, gözetim, kota ve muafiyet kontrolü yapılmaz. Örnek dosyadaki oranlar kurgusaldır.
- **Gümrük kıymeti CIF'ten farklı olabilir.** Royalti, komisyon, kap ve ambalaj gibi kıymete eklenen unsurlar varsa
  kıymeti onlarla birlikte girin. Kur, beyannamenin tescil tarihindeki kura göre belirlenir.
- **Maktu ÖTV hesaplanmaz.** Litre, kilogram veya adet başına ÖTV'si olan eşyada tutarı "Diğer Vergi" sütununa
  yazmayın (KDV matrahı doğru çıksa da ÖTV satırı boş kalır); bu eşya için paket uygun değildir.
- **Yurt içi giderlerin kapsamı size aittir.** KDV matrahına hangi giderlerin girdiğine siz karar verirsiniz.
- **Örnek veri:** Kalemler, oranlar ve kur kurgusaldır.

## Testler

- GV, İGV, KKDF, damga, diğer vergi, ÖTV ve KDV matrah sıralaması; KDV'siz kalem; kuruş yuvarlama.
- Damga ve yurt içi giderin kalemlere dağıtımı, toplamlar.
- Okuma hataları ve örnek veri uçtan uca.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
