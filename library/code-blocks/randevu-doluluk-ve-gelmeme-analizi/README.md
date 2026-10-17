# Randevu Doluluk ve Gelmeme Analizi · Kod Bloğu

> Hasta Hizmetleri › Randevu Planlama Sorumlusu · Gerisi Bende

Hekim ve bölüm bazında **randevu doluluk, gelmeme (no-show) ve iptal oranlarını** hesaplar; koşullar uygunsa
**fazla randevu (overbooking) önerisi** çıkarır. İnternete bağlanmaz. Hasta kimlik bilgisi gerekmez.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Gelmeme oranı % | Gelmeyen ÷ (gelen + gelmeyen) × 100 (iptaller paydaya girmez) |
| İptal oranı % | İptal ÷ (gelen + gelmeyen + iptal) × 100 |
| Kapasite | Günlük randevu kapasitesi × çalışılan gün (yazılmadıysa randevusu olan gün sayısı) |
| Doluluk % | (Gelen + gelmeyen) ÷ kapasite × 100 |
| Gerçekleşen doluluk % | Gelen ÷ kapasite × 100 |

**Fazla randevu önerisi (hekim başına, günlük):**

`ek randevu = günlük kapasite × p ÷ (1 − p) × emniyet katsayısı` (aşağı yuvarlanır; p = gelmeme oranı)

- `p ÷ (1 − p)`, gelmeyenleri ortalamada tam telafi eden orandır. `--emniyet` (varsayılan 0,5) bunun yarısını önerir;
  `--ust-sinir` (varsayılan %10) ek randevuyu günlük kapasitenin belirli bir oranıyla sınırlar.
- Öneri yalnız şu üç koşul birlikte sağlanırsa verilir: gerçekleşen randevu ≥ `--min-randevu` (30), gelmeme oranı ≥
  `--min-gelmeme` (%5), doluluk ≥ `--min-doluluk` (%90). Doluluk düşükse önce boş randevular doldurulmalıdır.

**Örnek:** 5 günde 75 gelen, 20 gelmeyen, 10 iptal; günlük kapasite 20 → gelmeme 20 ÷ 95 = %21,05; doluluk 95 ÷ 100 =
%95; tam telafi 20 × 0,2105 ÷ 0,7895 = 5,33; × 0,5 = 2,67 → 2; üst sınır 20 × %10 = 2 → günde 2 ek randevu.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri
python main.py --randevular randevular.xlsx --kapasite kapasite.xlsx
python main.py --randevular randevular.xlsx --kapasite kapasite.xlsx --emniyet 0,3 --ust-sinir 5 --min-doluluk 95
```

| Dosya | Sütunlar |
|---|---|
| Randevular | Tarih, Hekim, Durum (zorunlu); Bölüm, Adet. Durum: Geldi, Gelmedi, İptal. Her satır bir randevu; toplu veri için "Adet" sütunu kullanılabilir |
| Kapasite (isteğe bağlı) | Hekim, Günlük Kapasite (zorunlu); Çalışılan Gün |

## Çıktı

`cikti/randevu_analizi.xlsx`:

- `Hekim Bazında`: oranlar, kapasite, doluluk, önerilen ek randevu ve gerekçesi; boş "Karar" sütunu.
- `Bölüm Bazında`: bölüm toplamları ve oranları.
- `Haftanın Günü`: güne göre gelmeme ve iptal oranı (en yüksek gün vurgulanır).
- `Uyarılar`: tanınmayan durumlar, kapasitesi olmayan hekimler, %100'ü aşan doluluk.

## Dikkat

- **Öneri ortalamaya dayanır.** Gelmeme günden güne değişir; herkesin geldiği günlerde bekleme süresi uzar, hekim
  mesaisi aşılır. Karar hekim, bölüm sorumlusu ve yönetimle birlikte verilmelidir; kod yalnız hesap sunar.
- **Emniyet katsayısı ve üst sınır tercihtir.** 0,5 ve %10 temkinli başlangıç değerleridir, bir standart değildir.
  Muayene süresi uzun veya işlem ağırlıklı branşlarda fazla randevu uygun olmayabilir.
- **Kurum kuralları önce gelir.** Randevu sisteminizin ve bağlı olduğunuz mevzuatın ek randevuya izin verip
  vermediğini ayrıca kontrol edin; kod bunu bilmez.
- **Geçmiş dönem gelecek için garanti değildir.** Hatırlatma mesajı, mevsim veya hekim değişikliği gelmeme oranını
  değiştirir. Öneriyi uyguladıktan sonra aynı raporu yeniden çalıştırın.
- **Bekleyen randevular analize girmez.** Durumu Geldi, Gelmedi veya İptal olmayan satırlar sayılmaz ve uyarılır.
- **Kişisel veri:** Hasta adı, T.C. kimlik no ve telefon sütunlarına gerek yoktur; dosyaya koymayın.
- **Örnek veri:** Hekimler, bölümler ve sayılar kurgusaldır.

## Testler

- Gelmeme, iptal, doluluk ve gerçekleşen doluluk; kapasitesi verilmeyen hekim.
- Fazla randevu önerisi: emniyet katsayısı, üst sınır, üç koşulun sınır değerleri.
- Bölüm toplamları, haftanın günü kırılımı, genel toplam.
- Durum tanıma, okuma hataları ve örnek veri uçtan uca.

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
