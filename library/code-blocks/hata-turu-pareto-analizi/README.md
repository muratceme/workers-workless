# Hata Türü Pareto Analizi · Kod Bloğu

> Kalite Kontrol › Kalite Kontrol Şefi · Gerisi Bende

Ara ve final kontrol hata kayıtlarını **hata türü, operasyon ve hat bazında Pareto analizine** tabi tutar; hataların
çoğunu oluşturan az sayıdaki türü ve operasyonu gösterir, kontrol adetleri verilirse hat bazında hata oranını
hesaplar. İnternete bağlanmaz.

## Nasıl hesaplar?

| Kavram | Kural |
|---|---|
| Pay % | Kalemin hata adedi ÷ toplam hata adedi × 100 |
| Kümülatif % | Kalemler büyükten küçüğe sıralandığında payların birikimli toplamı |
| A sınıfı | Kendisinden önceki kümülatif pay `--esik` (80) yüzdesinin altında olan kalem |
| Hata oranı (DHU) | Hata adedi ÷ kontrol edilen adet × 100; yani 100 parçadaki hata sayısı |
| En sık hata | Hattın kendi içinde payı en büyük hata türü |

Hata adedi boş bırakılan satır 1 hata sayılır. Hata oranı yalnız `--kontrol` dosyası verilirse hesaplanır; dosyada
hat başına birden çok satır varsa adetler toplanır.

Uyarılar: hattın hata oranı `--dhu` (5) üstündeyse; "Diğer" grubunun payı %10'u aşıyorsa; türü yazılmamış hata varsa.

**Örnek:** 116 hatanın 34'ü dikiş atlaması → pay **%29,31**. Bir hatta iki günde 820 parça kontrol edilmiş, 46 hata
bulunmuş → hata oranı 100 parçada **5,61**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                              # örnek veri
python main.py --hatalar hata_kayitlari.xlsx
python main.py --hatalar hata_kayitlari.xlsx --kontrol kontrol_adetleri.xlsx --esik 70 --dhu 3
```

Zorunlu sütunlar: hatalarda Hat ve Hata Türü; kontrol adetlerinde Hat ve Kontrol Edilen Adet.

## Çıktı

`cikti/hata_turu_pareto.xlsx`:

- `Hata Türü Pareto`, `Operasyon Pareto`: sıra, sınıf, adet, pay, kümülatif pay.
- `Hat Pareto`: aynı tablo; ayrıca kontrol edilen adet, hata oranı ve hattın en sık hatası.
- `Hat × Hata Türü`: adet matrisi.
- `Özet`: toplamlar, A sınıfı tür sayısı, genel hata oranı, ara ve final kontrol dağılımı.
- `Uyarılar`.

## Dikkat

- **Hata oranı, hatalı parça yüzdesi değildir.** Bir parçada birden çok hata sayılabildiği için 100 parçadaki hata
  sayısı (DHU), hatalı parça oranından yüksek çıkar ve %100'ü aşabilir. Hatalı parça oranı için hatalı parça adedini
  ayrıca kaydetmeniz gerekir.
- **Ara ve final kontrolü birlikte oranlamak yanıltabilir.** Aynı parça hem ara hem final kontrolde sayıldıysa
  kontrol adetleri de buna göre verilmelidir. En sağlıklısı her kontrol noktasını ayrı çalıştırmaktır.
- **Hata türleri tutarlı kodlanmalı.** "Dikiş atlaması" ile "atlama" ayrı tür sayılır; kod yalnız büyük-küçük harf
  ve Türkçe karakter farkını yok sayar. Standart bir hata listesi kullanın.
- **Hataların ağırlığı eşit sayılır.** Kritik, majör ve minör hata ayrımı yapılmaz; iplik temizliği ile ölçü hatası
  aynı ağırlıktadır. Ağırlıklı analiz için ciddiyete göre ayrı dosyalarla çalıştırın.
- **Kod kök nedeni bulmaz.** Pareto nereye bakılacağını gösterir; neden (makine ayarı, operatör, kumaş, metot)
  sahada incelenmelidir.
- **Kabul örneklemesi (AQL) kararı vermez.** Parti kabul/ret değerlendirmesi bu paketin konusu değildir.
- **Örnek veri:** Hatlar, operasyonlar ve adetler kurgusaldır.

## Testler

- Pay, kümülatif pay ve A sınıfı sınırı; tür, operasyon ve hat sıralaması; eşit adette ada göre sıra.
- Hata oranı, hattın en sık hatası, hat × tür matrisi, ara ve final dağılımı.
- Kontrol adedi olmayan hat, "Diğer" ve türsüz hata uyarıları, okuma ve komut satırı.

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
