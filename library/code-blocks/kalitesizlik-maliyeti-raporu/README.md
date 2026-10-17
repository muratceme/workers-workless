# Kalitesizlik Maliyeti Raporu · Kod Bloğu

> Kalite Güvence ve Kalite Kontrol › Kalite Müdürü · Gerisi Bende

Hurda, yeniden işleme, iade ve garanti maliyetlerini **ürün ve neden bazında** toplar; iç ve dış hata maliyetini
ayırır, Pareto sıralamasıyla maliyetin çoğunu yaratan ürün ve nedenleri gösterir. İnternete bağlanmaz.

## Nasıl hesaplar?

| Kavram | Kural |
|---|---|
| İç hata maliyeti | Ürün müşteriye ulaşmadan yakalanan hatalar: hurda, yeniden işleme, ayıklama, yeniden kontrol |
| Dış hata maliyeti | Müşteriye ulaştıktan sonra ortaya çıkan hatalar: iade, garanti, müşteri şikâyeti, ceza |
| Pay % | Kalemin tutarı ÷ toplam kalitesizlik maliyeti × 100 |
| Kümülatif % | Kalemler büyükten küçüğe sıralandığında payların birikimli toplamı |
| A sınıfı | Kendisinden önceki kümülatif pay %80'in altında olan kalem |
| Dış hata payı % | Dış hata maliyeti ÷ toplam × 100; `--dis` (50) üstündeyse uyarılır |
| Ciroya oran % | Toplam kalitesizlik maliyeti ÷ `--ciro` × 100 |

Tanınmayan maliyet türleri "Sınıflanmamış" altında gösterilir ve bildirilir.

**Örnek:** Toplam 195.000 kalitesizlik maliyetinin 103.000'i iade ve garanti → dış hata payı **%52,82**. Ciro
6.500.000 ise kalitesizlik maliyetinin ciroya oranı **%3**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                   # örnek veri, ciro 6.500.000
python main.py --kayitlar kalitesizlik_kayitlari.xlsx
python main.py --kayitlar kalitesizlik_kayitlari.xlsx --ciro 12500000 --dis 40
```

Zorunlu sütunlar: Maliyet Türü, Tutar.

## Çıktı

`cikti/kalitesizlik_maliyeti_raporu.xlsx`:

- `Özet`: toplam, iç ve dış hata, ciroya oran; maliyet türleri ve aylık seyir.
- `Ürün Pareto`, `Neden Pareto`: tutar, pay, kümülatif pay, A sınıfı, iç ve dış hata ayrımı.
- `Ürün × Tür`: maliyet matrisi.
- `Kayıtlar`, `Uyarılar`.

## Dikkat

- **Kod maliyeti hesaplamaz, toplar.** Hurdanın maliyeti (malzeme, işçilik, genel gider payı dahil mi), yeniden
  işlemenin işçilik bedeli ve iadenin nakliye dahil tutarı sizin maliyet muhasebenizden gelir.
- **Kalite maliyetinin yalnız "hata" kısmıdır.** Önleme (eğitim, süreç iyileştirme) ve ölçme-değerlendirme
  (kontrol, test, kalibrasyon) maliyetleri bu raporda yoktur; toplam kalite maliyeti için onları ayrıca eklemelisiniz.
- **Kayda girmeyen kayıplar görünmez.** Müşteri kaybı, itibar zararı, acil sevkiyat, fazla stok ve plan bozulması
  gibi dolaylı maliyetler genellikle kayıt altında değildir; gerçek maliyet raporlanandan büyüktür.
- **Dış hata payı için %50 bir uyarı eşiğidir, standart değildir.** Sektöre ve ürüne göre beklenen dağılım değişir.
- **Neden kodları tutarlı olmalı.** Aynı nedenin farklı yazımları ayrı satır üretir; kod yalnız büyük-küçük harf ve
  Türkçe karakter farkını yok sayar.
- **Örnek veri:** Ürünler, nedenler, tutarlar ve ciro kurgusaldır.

## Testler

- İç ve dış hata ayrımı, tür eş anlamlıları, tanınmayan tür, pay ve ciroya oran.
- Ürün ve neden Pareto sıralaması, kümülatif pay ve A sınıfı sınırı, ürün × tür matrisi, ay tablosu.
- Dış hata eşiği, nedeni yazılmamış kayıt uyarısı, okuma ve komut satırı.

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
