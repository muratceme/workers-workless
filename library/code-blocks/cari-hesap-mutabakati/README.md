# Cari Hesap Mutabakatı · Kod Bloğu

> Muhasebe › Muhasebe Elemanı · Workers / Workless

Kendi cari ekstreniz (ör. 120 Alıcılar / 320 Satıcılar muavini) ile **karşı tarafın gönderdiği ekstreyi**
karşılaştırır; bakiye farkını kalem kalem açıklar ve her fark için olası nedeni yazar. Ay sonu ve yıl sonu
mutabakatlarında, bağımsız denetim öncesinde ve borç/alacak teyidinde kullanılır. İnternete bağlanmaz.

## Nasıl eşleştirir?

Karşı tarafın kayıtları genelde **ters yönlüdür** (sizin borcunuz onun alacağıdır); yön, ortak belge
numaralarından otomatik algılanır. Tüm tutarlar sizin bakışınıza çevrilir ve sırayla:

| Adım | Kural | Sonuç |
|---|---|---|
| 0 | Devir / açılış satırları | Mutabık ya da devir farkı |
| 1 | Belge no + tutar | Mutabık |
| 2 | Belge no aynı, tutar farklı | **Tutar farkı** (KDV hariç/dahil veya KDV tevkifatı ihtimali denetlenir) |
| 3 | Tutar aynı, tarih ± `--tolerans` gün | Mutabık (belge nosu olmayan ödeme, havale, çek) |
| 4 | Tutar aynı, tarih uzak | **Tarih farkı** (bakiyeyi etkilemez, dönem kontrolü önerilir) |

Belge numaralarında yalnız harf ve rakam karşılaştırılır; 16 haneli e-Fatura numarası
(`ABC2026000000123`) ile kısa sıra numarası (`123`) eşleşebilir, farklı seriler eşleşmez.

## Farklar ve olası nedenler

Eşleşmeyen kalemler belge türüne göre (fatura, iade, ödeme/tahsilat, çek/senet, kur farkı, vade farkı)
yorumlanır. Dönem sonuna `--son-gun` günden yakın kalemler **yolda** (karşı tarafta sonraki döneme
kaydedilmiş) olabilir diye işaretlenir. Aynı belgenin aynı tutarla iki kez kaydı **mükerrer kayıt**
uyarısı verir.

Özet sayfasında *Fark (biz − karşı)* ile *Kalemlerle açıklanan fark* her zaman eşittir; açıklanamayan
fark sıfır değilse ekstrelerden biri eksik dönem içeriyordur.

> Nedenler kural tabanlı tahmindir; kesin neden belgeler karşılaştırılarak doğrulanmalıdır.

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                                  # örnek ekstrelerle dener
python main.py --biz bizim_ekstre.xlsx --karsi karsi_taraf_ekstresi.xlsx
python main.py --biz bizim.xlsx --karsi karsi.xlsx --baslangic 01.07.2026 --bitis 30.09.2026 \
               --tolerans 5 --biz-adi "Şirketimiz A.Ş." --karsi-adi "Tedarikçi Ltd. Şti."
```

**Girdi sütunları:** Tarih, Borç, Alacak (veya işaretli Tutar) zorunlu; Belge No (Fatura No, Evrak No),
Belge Türü, Açıklama isteğe bağlı. Başlıklar Türkçe muhasebe programlarındaki yaygın adlarla tanınır.

## Çıktı

`Özet` · `Farklar` · `Eşleşenler` · `Bizim Ekstre` · `Karşı Taraf Ekstresi` · `Mutabakat Mektubu`
(bakiyeyi içeren, imzaya hazır mektup taslağı) · `Bilgi`

## Testler

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
