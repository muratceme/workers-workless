# Rücu Tahsilat Takibi · Kod Bloğu

> Hasar Yönetimi › Rücu Uzmanı · Gerisi Bende

Rücu dosyalarını **talep, cevap, tahkim/dava ve tahsilat aşamalarında** izler; cevabı geciken, kabul edildiği hâlde
tahsil edilmeyen ve zamanaşımı tarihi yaklaşan dosyaları öne çıkarır. İnternete bağlanmaz.

## Nasıl hesaplar?

| Kavram | Kural |
|---|---|
| Kalan tutar | Rücu tutarı − tahsil edilen |
| Tahsilat oranı % | Tahsil edilen ÷ rücu tutarı × 100 |
| Kapalı dosya | Aşaması "Tahsil edildi" ya da "Kapatıldı" olan veya tamamı tahsil edilen dosya |
| Zamanaşımı geçmiş | Zamanaşımı tarihi durum tarihinden önce |
| Zamanaşımı yaklaşıyor | Zamanaşımı tarihine `--gun` (90) gün veya az kaldı |
| Talep gönderilmedi | Talep tarihi boş |
| Cevap gecikti | Talep gönderilmiş, cevap tarihi boş ve `--cevap` (30) günden fazla geçmiş |
| Ret, karar bekliyor | Aşama "Reddedildi"; tahkim, dava ya da kapatma kararı verilmemiş |
| Tahsilat gecikti | Aşama "Kabul edildi", tutar kalmış ve cevabın üzerinden `--tahsil` (30) günden fazla geçmiş |
| Cevap süresi (gün) | Cevap tarihi − talep tarihi |

Bir dosya birden çok koşulu sağlıyorsa tablodaki sırayla ilk uyan durum yazılır. Tanınan aşamalar: Hazırlık, Talep
gönderildi, Kabul edildi, Reddedildi, Tahkim, Dava / İcra, Tahsil edildi, Kapatıldı.

**Örnek:** Talep 01.08.2026, cevap yok, durum tarihi 12.10.2026 → **72 gündür** cevap bekleniyor, durum "Cevap
gecikti". Başka bir dosyada zamanaşımı tarihi 30.11.2026 → **49 gün** kaldı, durum "Zamanaşımı yaklaşıyor".

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                              # örnek veri, durum tarihi 12.10.2026
python main.py --dosyalar rucu_dosyalari.xlsx
python main.py --dosyalar rucu_dosyalari.xlsx --gun 180 --cevap 15 --tahsil 45 --tarih 31.10.2026
```

Zorunlu sütunlar: Rücu No, Rücu Tutarı.

## Çıktı

`cikti/rucu_tahsilat_takibi.xlsx`:

- `Açık Dosyalar`: öncelik sırasıyla; durum, aşama, kalan tutar, cevap bekleme süresi, zamanaşımına kalan gün.
- `Özet`: toplam talep, tahsilat, oran, açık dosyaların durum dağılımı.
- `Aşamalar`, `Karşı Taraflar`: dosya, tutar, tahsilat oranı ve ortalama cevap süresi.
- `Tüm Dosyalar`, `Uyarılar`.

## Dikkat

- **Zamanaşımı tarihini kod hesaplamaz.** Rücu hakkının dayanağına, sözleşmeye ve olayın niteliğine göre süre ve
  başlangıcı değişir; tarihi hukuk biriminizle belirleyip dosyaya siz yazarsınız. Kod yalnız o tarihe kalan günü
  sayar. Tarih yazılmamış açık dosyalar ayrıca bildirilir.
- **30 günlük cevap ve tahsilat eşikleri çalışma varsayımıdır.** Şirketler arası protokollerde ya da
  sözleşmelerinizde farklı süreler olabilir; `--cevap` ve `--tahsil` ile değiştirin.
- **Tahkim ve davadaki dosyalar "Süreçte" görünür.** Kod duruşma, bilirkişi ya da karar tarihlerini izlemez.
- **Faiz ve masraf ayrı izlenmelidir.** Tahsil edilen tutar rücu tutarını aşarsa uyarı üretilir; kalan tutar sıfır
  kabul edilir.
- **Kısmi kabul ayrı gösterilmez.** Karşı taraf tutarın bir kısmını kabul ettiyse rücu tutarını güncellemeniz ya da
  dosyayı ikiye bölmeniz gerekir.
- **Kapsam dışı:** rücu hakkının doğup doğmadığının ve kusur oranının değerlendirilmesi, sovtaj, reasürans payı.
- **Örnek veri:** Dosyalar, şirketler ve tutarlar kurgusaldır.

## Testler

- Durum önceliği ve sınırları: zamanaşımı (geçmiş / 90 ve 91 gün), cevap (30 / 31 gün), kabul sonrası tahsilat.
- Kapalı dosya ayrımı (tam tahsil, kısmi tahsil, kapatma), kalan ve tahsil edilemeyen tutarlar.
- Aşama ve karşı taraf kırılımı, ortalama cevap süresi, tanınmayan aşama, okuma ve komut satırı.

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
