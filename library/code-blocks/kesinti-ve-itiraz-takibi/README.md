# Kesinti ve İtiraz Takibi · Kod Bloğu

> Anlaşmalı Kurumlar ve Faturalama › Anlaşmalı Kurumlar Uzmanı · Workers / Workless

Sigorta şirketlerinin yaptığı **kesintileri neden bazında sınıflandırır, itiraz edilebilir olanları ve süreleri**
takip eder. İnternete bağlanmaz. Hasta bilgisi gerekmez.

## Nasıl hesaplar?

1. **Neden kategorisi:** kesinti nedeni metni, kural dosyasındaki anahtar kelimelerle karşılaştırılır; dosyadaki
   sıraya göre **ilk eşleşen** kategori atanır. Eşleşme yoksa "Diğer" olur.
2. **İtiraz edilebilir:** kural dosyasındaki Evet / Hayır. "Diğer" kategorisi "İncelenecek" sayılır.
3. **Son itiraz günü** = kesinti tarihi + şirketin itiraz süresi. Süre önce süre dosyasından, şirket orada yoksa
   `--itiraz-gun` değerinden alınır; ikisi de yoksa o şirkette süre takibi yapılmaz ve uyarı verilir.
4. **Durum:**

   | Durum | Koşul |
   |---|---|
   | Kabul edildi / Kısmi kabul / Reddedildi | İtiraz sonucu yazılmış |
   | Sonuç bekleniyor | İtiraz tarihi var, sonuç yok |
   | İtiraz edilmeyecek | İtiraz yok ve kategori "Hayır" |
   | Süre geçti | İtiraz yok ve rapor tarihi > son itiraz günü |
   | Süre yaklaşıyor | İtiraz yok ve kalan gün ≤ `--yaklasan-gun` (varsayılan 5) |
   | İtiraz bekliyor | İtiraz yok, süresi var veya süre tanımlı değil |

5. **Geri kazanım %** = geri alınan tutar ÷ sonuçlanan itirazların kesinti tutarı × 100. "Kabul" sonucunda geri alınan
   tutar boşsa kesintinin tamamı alınmış sayılır.

**Örnek:** kesinti 30.09.2026, şirketin süresi 15 gün → son itiraz günü 15.10.2026; rapor 12.10.2026 → 3 gün kaldı →
Süre yaklaşıyor.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri (rapor tarihi 12.10.2026)
python main.py --kesintiler kesintiler.xlsx --kurallar neden_kurallari.xlsx --sureler itiraz_sureleri.xlsx
python main.py --kesintiler kesintiler.xlsx --kurallar neden_kurallari.xlsx --itiraz-gun 30 --yaklasan-gun 7
```

| Dosya | Sütunlar |
|---|---|
| Kesintiler | Sigorta Şirketi, Kesinti Tarihi, Kesinti Tutarı (zorunlu); Fatura No, Kesinti Nedeni, İtiraz Tarihi, İtiraz Sonucu (Kabul / Kısmi kabul / Ret), Geri Alınan Tutar |
| Neden kuralları | Neden Kategorisi, Anahtar Kelimeler (virgülle); İtiraz Edilebilir (Evet / Hayır) |
| İtiraz süreleri (isteğe bağlı) | Sigorta Şirketi, İtiraz Süresi (gün) |

## Çıktı

`cikti/kesinti_itiraz_takibi.xlsx`:

- `Kesinti Takibi`: süresi geçen ve yaklaşanlardan başlayarak her kesinti; kategori, son itiraz günü, kalan gün; boş "Not" sütunu.
- `Şirket Özeti`: kesinti tutarı, itiraz edilen, geri alınan, geri kazanım %, itiraz bekleyen ve süresi geçen tutar.
- `Neden Özeti`: kategori başına adet, tutar ve pay.
- `Uyarılar`: süresi tanımlı olmayan şirketler, sınıflandırılamayan nedenler, süresinden sonra yapılmış itirazlar.

## Dikkat

- **Süreler ve itiraz hakkı sözleşmeye bağlıdır.** Her sigorta şirketiyle yapılan anlaşmada itiraz süresi, yöntemi
  ve itiraz edilebilecek kesinti türleri farklı olabilir. Kod bir süre varsaymaz; örnek dosyadaki 30 ve 15 gün
  kurgusaldır. Sürenin hangi tarihten başladığını (bildirim, ödeme, mutabakat) da sözleşmenizden kontrol edin.
- **"İtiraz edilebilir" sütunu bir ön ayrımdır.** Örnek kural dosyasındaki Evet / Hayır değerleri örnektir, hukuki
  değerlendirme değildir. "Hayır" kategorisindeki bir kesinti de somut olayda itiraza değer olabilir.
- **Sınıflandırma kelime eşleşmesidir.** Metni anlamaz; "provizyon limiti aşıldı" gibi iki kategoriye uyan nedenlerde
  dosyada önce gelen kural geçerlidir. Kural sırasını ve `Neden Özeti`ndeki "Diğer" payını gözden geçirin.
- **Takvim günü sayar.** İş günü, resmi tatil ve son günün tatile gelmesi hesaplanmaz.
- **Kişisel sağlık verisi:** Hasta adı, T.C. kimlik no ve tanı sütunlarına gerek yoktur; dosyaya koymayın.
- **Örnek veri:** Şirketler, faturalar, tutarlar ve süreler kurgusaldır.

## Testler

- Neden sınıflandırması (ilk eşleşen kural, eşleşmeyen neden), itiraz sonucu tanıma.
- Süre sınırları (son gün, bir gün sonrası, yaklaşan gün eşiği), şirkete özel süre, süresi olmayan şirket.
- Şirket ve neden özetleri, geri kazanım oranı.
- Okuma hataları ve örnek veri uçtan uca.

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
