# Kaizen Önerisi Değerlendirme · Kod Bloğu

> Endüstri Mühendisliği ve Süreç Geliştirme › Endüstri Mühendisi · Gerisi Bende

Çalışan iyileştirme önerilerini **kazanç, geri ödeme süresi ve uygulanabilirliğe göre** puanlar, sıralar ve üç
sınıfa ayırır. Aynı ölçütleri her öneriye aynı biçimde uygular; kararı siz verirsiniz. İnternete bağlanmaz.

## Nasıl hesaplar?

| Kavram | Kural |
|---|---|
| Kazanç puanı (1–5) | Yıllık kazancın geçtiği eşik sayısı + 1; `--kazanc 10000;50000;100000;250000` |
| Geri ödeme süresi (ay) | Uygulama maliyeti ÷ (yıllık kazanç ÷ 12); maliyet sıfırsa 0 |
| Geri ödeme puanı (1–5) | 5 − sürenin aştığı eşik sayısı; `--geri-odeme 3;6;12;24`; kazanç sıfırsa 1 |
| Uygulanabilirlik (1–5) | Dosyada sizin verdiğiniz puan; boşsa 3 |
| Toplam puan (20–100) | Üç puanın ağırlıklı ortalaması × 20; `--agirlik 50;30;20` |
| Sınıf | Puan ≥ `--oncelikli` (70) "Öncelikli"; ≥ `--degerlendir` (50) "Değerlendir"; altı "Düşük öncelik" |
| Net ilk yıl kazancı | Yıllık kazanç − uygulama maliyeti |

Eşit puanlı önerilerde net ilk yıl kazancı yüksek olan önce gelir.

**Örnek:** Yıllık kazanç 180.000, maliyet 45.000, uygulanabilirlik 4. Kazanç puanı 4 (üç eşiği geçiyor), geri ödeme
45.000 ÷ 15.000 = 3 ay → puan 5. Toplam = (4×50 + 5×30 + 4×20) ÷ 100 × 20 = **86** → "Öncelikli".

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                              # örnek veri
python main.py --oneriler kaizen_onerileri.xlsx
python main.py --oneriler kaizen_onerileri.xlsx --kazanc 25000;100000;250000;1000000 --agirlik 40;20;40 --oncelikli 75
```

Zorunlu sütunlar: Konu, Yıllık Kazanç.

## Çıktı

`cikti/kaizen_degerlendirme.xlsx`:

- `Öneriler`: puan sırasıyla; sınıf, kazanç, maliyet, net ilk yıl, geri ödeme süresi ve üç alt puan.
- `Özet`: sınıf bazında sayılar, toplam ve öncelikli önerilerin kazanç ve maliyeti.
- `Bölümler`: bölüm bazında öneri sayısı, kazanç, maliyet ve ortalama puan.
- `Uyarılar`.

## Dikkat

- **Puan, girdiğiniz tahminler kadar iyidir.** Yıllık kazanç ve maliyet öneri sahibinin ya da sizin tahmininizdir;
  kod bunları doğrulamaz. Kazancı tahmin ederken aynı yöntemi (hangi işçilik ücreti, hangi üretim adedi) bütün
  önerilere uygulayın.
- **Parasal olmayan faydalar puanlanmaz.** İş güvenliği, ergonomi, kalite ve çalışan memnuniyeti gibi kazanımlar
  kazanç sütununa yazılmadıysa öneri düşük puan alır; bu öneriler ayrıca bildirilir ve ayrı değerlendirilmelidir.
  Güvenlik önerilerini parasal sıralamaya sokmamak genellikle daha doğrudur.
- **Eşikler ve ağırlıklar politikanızdır.** Varsayılan tutarlar orta ölçekli bir işletme için seçilmiş başlangıç
  değerleridir; kendi ölçeğinize ve öneri sisteminize göre değiştirin.
- **Ödül hesabı değildir.** Öneri sahibine verilecek ödül ya da prim bu paketin konusu değildir.
- **Geri ödeme basit hesaptır.** Paranın zaman değeri, bakım ve işletme giderleri, kazancın kaç yıl süreceği
  hesaba girmez.
- **Kapsam dışı:** önerinin teknik fizibilitesi, mükerrer öneri kontrolü, uygulama takibi ve gerçekleşen kazancın
  doğrulanması.
- **Örnek veri:** Öneriler, kişiler ve tutarlar kurgusaldır.

## Testler

- Kazanç ve geri ödeme puanlarının eşik sınırları, maliyeti ya da kazancı sıfır olan öneriler.
- Ağırlıklı toplam puan, sınıf sınırları, eşit puanda sıralama, ağırlıkların değiştirilmesi.
- Bölüm özeti, eksik uygulanabilirlik puanı, okuma ve komut satırı.

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
