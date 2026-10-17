# Lab Dip ve Onay Takibi · Kod Bloğu

> Müşteri Temsilciliği (Merchandising) › Production Merchandiser · Gerisi Bende

Renk (lab dip), baskı, nakış ve aksesuar onaylarının **gönderim ve müşteri cevabı durumunu sipariş bazında** izler:
cevabı geciken, reddedilip yeniden gönderilmesi gereken ve onay termini geçen ya da yaklaşan kalemleri öne çıkarır.
İnternete bağlanmaz.

## Nasıl hesaplar?

Bir **onay kalemi** sipariş + onay türü + renk/detay birleşimidir. Dosyada her satır bir gönderim turudur; kalemin
durumunu son tur belirler (tur numarası, yoksa gönderim tarihi en büyük olan).

| Durum | Kural |
|---|---|
| Gönderilmedi | Son turda gönderim tarihi yok |
| Cevap bekleniyor | Gönderilmiş, cevap tarihi yok |
| Cevap gecikti | Cevap bekleme süresi `--cevap` (7) günü aşmış |
| Onaylandı | Son cevabın sonucu Onay |
| Reddedildi / Yorumlu, yeniden gönderilecek | Son cevabın sonucu Ret ya da Yorum |
| Sonuç yazılmamış | Cevap tarihi var, sonuç boş |

| Gösterge | Formül |
|---|---|
| Termin geçti | Onaylanmamış kalemde onay termini durum tarihinden önce |
| Termin yaklaşıyor | Termine `--gun` (7) gün veya az kaldı |
| Onay oranı % | Onaylanan kalem ÷ toplam kalem × 100 |
| İlk turda onay % | Tek turda onaylanan kalem ÷ onaylanan kalem × 100 |
| Cevap süresi (gün) | Cevap tarihi − gönderim tarihi; bütün turların ortalaması |

**Örnek:** Baskı onayı 01.10.2026'da gönderildi, cevap yok, onay termini 10.10.2026; durum tarihi 12.10.2026 →
**11 gündür** cevap bekleniyor ("Cevap gecikti") ve termin **2 gün önce** geçti.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                            # örnek veri, durum tarihi 12.10.2026
python main.py --onaylar onay_kayitlari.xlsx
python main.py --onaylar onay_kayitlari.xlsx --cevap 5 --gun 10 --tur 2 --tarih 20.10.2026
```

Zorunlu sütunlar: Sipariş No, Onay Türü. Sonuç için Onay, Ret ya da Yorum yazın (Onaylandı, Approved, Reddedildi,
Rejected, Revize ve Şartlı da tanınır).

## Çıktı

`cikti/lab_dip_onay_takibi.xlsx`:

- `Onay Kalemleri`: açık kalemler önce, termini geçenler en üstte; tur sayısı, son gönderim ve cevap, bekleme süresi.
- `Özet`: onay oranı, ilk turda onay, bekleyen cevaplar, termin durumu, durum dağılımı.
- `Siparişler`, `Müşteriler`, `Onay Türleri`: aynı göstergeler kırılımlı; ortalama müşteri cevap süresi.
- `Uyarılar`.

## Dikkat

- **Onay terminini siz belirlersiniz.** Termin, onayın gelmemesi hâlinde üretim takviminin kayacağı tarihtir (kumaş
  boyama, baskı ya da aksesuar siparişinin verilmesi gereken gün). Kod kritik yol hesabı yapmaz, sevk tarihinden
  geriye saymaz.
- **7 günlük cevap eşiği çalışma varsayımıdır.** Müşterinizin el kitabındaki ya da sözleşmedeki cevap süresine göre
  `--cevap` ile değiştirin.
- **"Yorumlu" kalem onaylanmış sayılmaz.** Müşteri "yorumla onay" verdiyse ve yeniden gönderim gerekmiyorsa sonucu
  Onay olarak yazın.
- **Kalem adları tutarlı olmalı.** Aynı rengin iki turunda detay farklı yazılırsa ("Lacivert" ve "Navy") iki ayrı
  kalem oluşur. Büyük-küçük harf ve Türkçe karakter farkı yok sayılır.
- **Kapsam dışı:** renk ölçüm değerleri (ΔE) ve ışık kaynağı, numune (fit, PP, TOP) onayları, test raporları, kargo
  takip numaraları.
- **Örnek veri:** Siparişler, müşteriler ve tarihler kurgusaldır.

## Testler

- Durum ataması ve son turun belirlenmesi; cevap bekleme sınırı (7 / 8 gün), termin sınırları.
- Onay oranı, ilk turda onay, ortalama cevap süresi; sipariş, müşteri ve tür kırılımı.
- Çok turlu onay uyarısı, tanınmayan sonuç, okunamayan tarih, okuma ve komut satırı.

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
