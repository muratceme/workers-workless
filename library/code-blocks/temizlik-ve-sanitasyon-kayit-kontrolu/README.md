# Temizlik ve Sanitasyon Kayıt Kontrolü · Kod Bloğu

> Üretim › Üretim Şefi · Gerisi Bende

Temizlik ve sanitasyon **planını** dönemdeki **kayıtlarla** karşılaştırır; her görev için beklenen ve yapılan temizliği,
**uyum yüzdesini**, eksik gün / hafta / ayı ve aralığı aşılan görevleri listeler; uygunsuz doğrulama sonucunda düzeltici
faaliyet yazılmamış kayıtları işaretler. İnternete bağlanmaz.

## Nasıl kontrol eder?

| Sıklık | Beklenen | Eksik sayılan |
|---|---|---|
| Vardiya | Her çalışma günü `--vardiya` kez | Kayıt sayısı vardiya sayısından az olan gün |
| Günlük | Her çalışma günü 1 kez | Kaydı olmayan çalışma günü |
| Haftalık | Dönemdeki her takvim haftası (pazartesi-pazar) 1 kez | Kaydı olmayan hafta |
| Aylık | Dönemdeki her ay 1 kez | Kaydı olmayan ay |
| N gün (ör. "3 gün") | İki temizlik arasında en çok N gün | N günden uzun kayıtsız aralık (dönem başı ve sonu dahil) |

Uyum % = kaydı bulunan periyot ÷ beklenen periyot × 100. `--kapali` ile verilen günler (ör. "Paz") vardiya ve günlük
görevlerde çalışma günü sayılmaz. Plandaki hat / alan boşsa görev tüm alanlardaki aynı adlı kayıtlarla eşleşir.

**Diğer kontroller:** sonucu "Uygun değil" olup düzeltici faaliyeti yazılmamış kayıt (Yüksek) · planda olmayan görev ·
onaylanmamış kayıt (bilgi).

**Örnek:** Drenaj kanalları en çok 3 günde bir temizlenmeli; 04.10'dan sonraki kayıt 09.10'da → **05.10–08.10 arası 4 gün
kayıtsız**, uyum %80.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                                  # örnek veri (2 vardiya, pazar kapalı)
python main.py --plan temizlik_plani.xlsx --kayitlar temizlik_kayitlari.xlsx --baslangic 28.09.2026 --bitis 11.10.2026 --vardiya 2 --kapali Paz
python main.py --plan temizlik_plani.xlsx --kayitlar temizlik_kayitlari.xlsx --kapali "Cmt,Paz"
```

- **Plan:** Hat / Alan, Görev, Sıklık.
- **Kayıtlar:** Tarih, Vardiya, Hat / Alan, Görev, Yapan, Kontrol Eden, Sonuç (Uygun / Uygun değil), Düzeltici Faaliyet.

Dönem verilmezse kayıtların ilk ve son günü kullanılır.

## Çıktı

`cikti/sanitasyon_kayit_kontrolu.xlsx`:

- `Görev Uyumu`: uyuma göre (en düşük önce) beklenen, yapılan, uyum %, kayıt ve uygunsuz sayısı, eksik / geciken periyotlar.
- `Kayıtlar`: dönemdeki kayıtlar; uygunsuz sonuç ve onaysız kayıt renkli.
- `Uyarılar`.

## Dikkat

- **Kayıt temizliğin kanıtı değildir.** Temizliğin etkinliği görsel kontrol, ATP / mikrobiyolojik doğrulama ve
  denetimle değerlendirilir; kod yalnız kayıtların plana uygunluğunu kontrol eder.
- **Plan, gıda güvenliği sisteminizden gelmelidir.** Sıklık ve yöntem HACCP planı, ön gereksinim programları ve müşteri
  standartlarına (ör. BRCGS, IFS) göre belirlenir.
- **Üretim yapılmayan günler** `--kapali` ile tanımlanmazsa eksik sayılır; bayram ve bakım duruşlarını dönem dışında bırakın.
- **Örnek veri:** Hatlar, görevler, kişiler ve sonuçlar kurgusaldır.

## Testler

- Vardiya, günlük, haftalık, aylık ve N gün sıklıkları; kapalı günler; uyum yüzdesi ve genel uyum.
- Uygunsuz sonuç, planda olmayan görev, onaysız kayıt, genel (hat bağımsız) görev, okuma ve komut satırı.

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
