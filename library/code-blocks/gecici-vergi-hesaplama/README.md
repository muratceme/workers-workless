# Geçici Vergi Hesaplama · Kod Bloğu

> Beyanname ve Vergi › Serbest Muhasebeci Mali Müşavir (SMMM) · Workers / Workless

Dönem sonu kümülatif mizandan ticari kârı bulur; verdiğiniz ilave, istisna, zarar ve indirimlerle **geçici vergi
matrahını ve ödenecek vergiyi** adım adım hesaplar. İnternete bağlanmaz.

## Nasıl hesaplar?

| Adım | Kural |
|---|---|
| Ticari bilanço kârı | 600–689 hesaplarının (alacak bakiye − borç bakiye) toplamı |
| + İlaveler | Kanunen kabul edilmeyen giderler (tür: KKEG) |
| − İstisnalar | Zarar olsa dahi indirilecekler (tür: İstisna); sonucu zarara çevirebilir |
| − Geçmiş yıl zararı | En çok kalan kâr kadar (tür: Geçmiş Yıl Zararı) |
| − Diğer indirimler | Kazanç varsa indirilecekler, en çok kalan kâr kadar (tür: İndirim) |
| = Matrah | Negatifse 0; dönem zararı ayrıca gösterilir |
| Hesaplanan vergi | Matrah × `--oran` ÷ 100 |
| Ödenecek | Hesaplanan − `--onceki` − `--stopaj`; negatifse 0 |

- Mizanda hem üç haneli ana hesaplar hem alt hesaplar varsa yalnız ana hesaplar toplanır.
- 690 ve sonrası kapanış hesapları alınmaz; **kapanış kaydı öncesi** mizanı kullanın.
- 7'li maliyet hesaplarında bakiye varsa (yansıtma yapılmamışsa) uyarı verilir.

**Örnek:** Ticari kâr 680.000 + KKEG 45.000 − istisna 30.000 − geçmiş yıl zararı 120.000 − indirim 10.000 = matrah
565.000. Oran %25 → 141.250. Önceki dönem 80.000 ve kesinti 5.000 düşülünce ödenecek 56.250.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek veri (oran %25, önceki 80.000, kesinti 5.000)
python main.py --mizan mizan.xlsx --oran 25
python main.py --mizan mizan.xlsx --duzeltmeler duzeltmeler.xlsx --oran 25 --onceki 80000 --stopaj 5000
python main.py --ticari-kar=680000 --duzeltmeler duzeltmeler.xlsx --oran 25
```

- Mizan: Hesap Kodu, Hesap Adı, Borç Bakiye, Alacak Bakiye. Bakiye sütunları yoksa Borç ve Alacak toplamlarının farkı alınır.
- Düzeltmeler: Kalem, Tür, Tutar.
- `--oran` **zorunludur**; kod bir oran varsaymaz. Negatif kâr için `--ticari-kar=-50000` biçiminde yazın.

## Çıktı

`cikti/gecici_vergi.xlsx`:

- `Hesap`: ticari kârdan ödenecek vergiye kadar her adım.
- `Gelir Tablosu Hesapları`: kâra giren hesaplar ve etkileri.
- `Düzeltmeler`: verdiğiniz ilave, istisna, zarar ve indirim kalemleri.
- `Uyarılar`: dengesiz mizan, yansıtılmamış maliyet hesapları, kullanılamayan zarar ve indirimler.

## Dikkat

- **Ön hesaptır, beyanname değildir.** Oranı, hangi giderin kabul edilmediğini, istisna ve indirimlerin kapsamını ve
  sırasını kod bilmez; mükellefin türüne (kurumlar, gelir) ve yürürlükteki mevzuata göre siz belirlersiniz.
- **Dönem sonu işlemleri yapılmış olmalı.** Amortisman, reeskont, kur değerlemesi, stok sayımı ve maliyet kayıtları
  eksikse mizandaki kâr yanlıştır; kod bunu göremez.
- **Enflasyon düzeltmesi, yatırım indirimi, indirimli oran, asgari vergi gibi özel hesaplar kapsam dışıdır.**
- **Geçmiş yıl zararının mahsup süresi ve koşulları kontrol edilmez.** Tutarı siz doğrularsınız.
- **Gelir vergisi mükelleflerinde** ticari kazanç dışındaki unsurlar ve tarife bu hesapta yoktur.
- **Örnek veri:** Mizan ve düzeltme kalemleri kurgusaldır.

## Testler

- Tam hesap; zarar ve indirimin kâr kadar sınırlanması; istisnanın zarar doğurması; mahsup fazlası.
- Ana hesap ve alt hesap ayrımı, kapanış ve 7'li hesap uyarıları, dengesiz mizan.
- Bakiye sütunu olmayan mizan, okunamayan satırlar, komut satırı hataları; örnek veri uçtan uca.

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
