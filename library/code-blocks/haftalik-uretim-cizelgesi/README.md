# Haftalık Üretim Çizelgesi · Kod Bloğu

> Üretim › Üretim Planlama › Üretim Planlama Uzmanı · Workers / Workless

Açık siparişleri, ürün rotalarını ve makine kapasitelerini kullanarak **sonlu kapasiteli** haftalık üretim
çizelgesi üretir. Hangi işin hangi makinede, hangi gün ve saatte yapılacağını, hangi siparişin **gecikeceğini**
ve **darboğaz** makineyi gösterir. İnternete bağlanmaz.

## Yöntem

1. **Sıralama:** Siparişler **en erken termine göre** (EDD) sıralanır. Eşitlikte `Öncelik` değeri yüksek
   olan önce gelir.
2. **Yerleşim:** Her siparişin operasyonları rota sırasıyla, kendi iş merkezindeki makineler arasından **en erken
   bitirebilecek makineye** yerleştirilir. Bir operasyon, önceki operasyon tüm partiyi bitirmeden başlamaz
   (transfer partisi yok).
3. **Takvim:**
   - Net süre = (hazırlık + miktar × birim süre) ÷ verimlilik
   - Vardiyalar 08:00'de başlar ve 8 saattir; 2 vardiya 08:00–24:00, 3 vardiya 24 saat.
   - Vardiya bitince iş sonraki vardiyada devam eder. Hafta sonu (`--calisma-gunleri`) ve tatiller atlanır.

> Sezgisel bir çizelgedir, optimum değildir. Malzeme hazır olma, kalıp/aparat ve operatör kısıtları dikkate
> alınmaz. Çizelgeyi planlamacı gözden geçirmelidir.

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                                     # örnek: 6 iş emri, 8 makine, 02.11.2026 haftası
python main.py --siparisler is_emirleri.xlsx --rotalar rotalar.xlsx --makineler makineler.xlsx --baslangic 02.11.2026
python main.py ... --calisma-gunleri 6 --tatiller tatiller.txt --hafta 2
```

## Çıktı

`Sipariş Planı` · `Haftalık Çizelge` (makine × gün; %60 üzeri sarı, %95 üzeri kırmızı) · `Operasyon Planı` ·
`Makine Doluluğu` (darboğaz grafiği) · `Bilgi`

## Testler

Elle hesaplanmış senaryo test edilir:
- Termini önce olan iş önce yapılır.
- İkinci işin torna operasyonu boştaki CNC-2'ye gider ve Salı 10:30'da biter.
- Vardiyaya yayılma, hafta sonu, tatil, 3 vardiya ve verimlilik hesapları doğrulanır.
- Aynı makinede çakışma olmadığı kontrol edilir.

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
