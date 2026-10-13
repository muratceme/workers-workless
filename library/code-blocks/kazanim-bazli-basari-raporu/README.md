# Kazanım Bazlı Başarı Raporu · Kod Bloğu

> Ölçme Değerlendirme ve Rehberlik › Ölçme Değerlendirme Uzmanı · Workers / Workless

Çoktan seçmeli sınav cevaplarını cevap anahtarı ve soru-kazanım eşleşmesiyle değerlendirir; **sınıf ve öğrenci
bazında kazanım başarı yüzdelerini** ve soru analizini raporlar. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Kazanım başarısı % | Kazanımın sorularındaki doğru sayısı ÷ (öğrenci sayısı × soru sayısı) × 100 |
| Sınıf × kazanım | Aynı oran, yalnız o sınıfın öğrencileriyle |
| Öğrenci × kazanım | Öğrencinin o kazanımdaki doğru sayısı ÷ kazanımın soru sayısı × 100 |
| Soru analizi | Doğru %, boş %, en çok işaretlenen yanlış seçenek |
| Destek gerekli | Başarısı `--esik` (%50) altında kalan kazanım |

- Boş bırakılan cevap yanlış sayılmaz, ayrı gösterilir; ancak başarı yüzdesinde doğru sayılmadığı için oranı düşürür.
- Anahtarda doğru cevabı `İptal` veya `X` olan soru değerlendirmeye alınmaz.
- Soru başlıkları `1`, `S1`, `Soru 1` biçimlerinde yazılabilir.

**Örnek:** "Üçgende açılar" kazanımı 2 soruyla ölçülmüş, 8 öğrenci toplam 6 doğru yapmış: 6 ÷ (8 × 2) = %37,5 →
"Destek gerekli". Aynı kazanımda 9-A %50, 9-B %25.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek veri
python main.py --cevaplar cevaplar.xlsx --anahtar cevap_anahtari.xlsx
python main.py --cevaplar cevaplar.xlsx --anahtar cevap_anahtari.xlsx --esik 60
```

- Cevaplar: Öğrenci, Sınıf ve her soru için bir sütun; hücrede öğrencinin işaretlediği seçenek (boş = boş bırakmış).
- Cevap anahtarı: Soru, Doğru Cevap, Kazanım.

## Çıktı

`cikti/kazanim_basari_raporu.xlsx`:

- `Kazanım Özeti`: kazanım bazında doğru, yanlış, boş; genel ve sınıf bazında başarı; eşik altı öğrenci sayısı.
- `Öğrenci × Kazanım`: her öğrencinin kazanım yüzdeleri ve destek gereken kazanımları.
- `Soru Analizi`: soru bazında doğru ve boş oranı, en çok işaretlenen yanlış seçenek.
- `Uyarılar`: eşiğin altındaki kazanımlar, sınıf düzeyinde düşük kazanımlar, iptal ve eksik sorular.

## Dikkat

- **Net ve puan hesaplanmaz.** Yanlışın doğruyu götürdüğü sınavlarda net hesabı ayrıca yapılır; bu rapor kazanım
  düzeyinde doğru oranına bakar.
- **Az soruyla ölçülen kazanım kaba bir göstergedir.** Tek soruyla ölçülen kazanımda öğrenci ya %0 ya %100 alır;
  sorunun kendisi hatalı veya çeldiricisi güçlü olabilir. Soru analizine birlikte bakın.
- **Eşik pedagojik bir karardır.** %50 varsayılan değerdir; kurumunuzun ölçütüne göre değiştirin.
- **Açık uçlu ve kısmi puanlı sorular kapsam dışıdır.** Yalnız tek doğru cevaplı sorular değerlendirilir.
- **Öğrenci verisi kişisel veridir.** Raporu yalnız ilgili öğretmen ve yöneticilerle paylaşın; veli paylaşımında
  yalnız o öğrencinin satırını kullanın.
- **Örnek veri:** Öğrenciler, kazanımlar ve cevaplar kurgusaldır.

## Testler

- Kazanım, sınıf ve öğrenci düzeyinde başarı yüzdeleri; doğru, yanlış, boş sayıları; soru analizi ve çeldirici.
- Eşik değişimi, sınıf düzeyinde düşük kazanım, iptal soru, cevap sütunu olmayan soru.
- Farklı başlık yazımları, okunamayan anahtar satırları, komut satırı hataları; örnek veri uçtan uca.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
