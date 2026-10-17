# Tıbbi Rapor Yazım Kontrolü · AI Agent

> Tıbbi Sekreterlik › Tıbbi Sekreter · Gerisi Bende

Hekimin dikte ettiği rapor metnini **yazım, tıbbi terim, taraf, birim ve kısaltma** açısından kontrol eder.
**Hasta kaydıyla tutarlılığı (yaş, cinsiyet, tarihler, TC kimlik no, protokol no, sağ / sol) kod denetler.** Model metni
yeniden yazmaz; yalnız alıntı ve öneri verir. Bir sayıyı değiştiren ya da anlamı değiştiren öneri "hekime sor"
işaretlenmeden kabul edilmez.

## Nasıl çalışır?

1. **Gizleme (kod):** Hasta kaydı modele gönderilmez. Metindeki hasta adı `{HASTA}` olur; TC kimlik no, protokol no,
   telefon ve e-posta maskelenir.
2. **Tutarlılık (kod):**
   - "45 yaşında" gibi ifadeler doğum tarihine göre rapor tarihindeki yaşla aynı mı
   - Cinsiyetle çelişen ifadeler (erkek hastada "kadın hasta", "gebe", "uterus"; kadın hastada "prostat" …)
   - Rapor tarihinden sonraki ya da doğum tarihinden önceki tarihler
   - Metindeki TC kimlik no ve protokol no kayıttan farklı mı
   - Rapor türü tek taraflıysa ("Sağ diz MR") karşı tarafın geçtiği cümleler
3. **Yazım kontrolü (model):** Yazım, tıbbi terim, taraf tutarlılığı, birim / ölçü, kısaltma, noktalama bulguları.
4. **Denetim (kod):** Aşağıdakilerden biri varsa sonuç modele geri gönderilir (varsayılan en fazla 2 tur):
   - Alıntı metinde aynen geçmiyor; öneri boş ya da alıntıyla aynı; aynı alıntı iki kez yazılmış
   - Öneri bir sayıyı değiştiriyor ama "hekime sor" işaretli değil (ör. "8 mm" → "18 mm")
   - Yazım önerisi alıntıdan çok farklı (ör. "Minimal efüzyon" → "Belirgin sıvı birikimi": anlam değişiyor)
5. **Önerilen metin (kod):** Yalnız "hekime sor" olmayan öneriler uygulanır ve ayrı bir dosyaya yazılır. Asıl rapor
   değişmez.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                           # örnek (kurgusal) diz MR raporu
python agent.py --hasta hasta.txt --rapor dikte.docx
```

## Çıktı

- `cikti/rapor_kontrol.xlsx`: Bulgular (kırmızı: hekime sorulacak, sarı: yazım önerisi).
- `cikti/rapor_kontrol.txt`: önerilen metin.

## Dikkat

- **Hekim onayı.** Rapor hekimin sorumluluğundadır. Öneriler hekim onayı olmadan rapora işlenmez. Tıbbi içerik,
  ölçüm, doz ve tarih düzeltmeleri yalnız hekim tarafından yapılır.
- **Sağlık verisi.** Rapor metni yapay zekâ sağlayıcısına gönderilir. Maskeleme yardımcıdır, garanti değildir. Sağlık
  verisi KVKK'ya göre özel nitelikli kişisel veridir. Kurumunuzun yazılı izni ve veri işleme sözleşmesi yoksa bu agent'ı
  çalıştırmayın ya da yerel model (Ollama) kullanın.
- **Örnek veri:** Hasta ve rapor kurgusaldır; TC kimlik no yaygın kullanılan test numarasıdır.

## Testler

Testler gerçek API çağırmaz. Kodun örnek rapordaki beş tutarsızlığı bulduğu doğrulanır: kayıtta 47 yaşında erkek hasta
için "45 yaşında kadın hasta", rapor tarihinden sonraki bir tarih, farklı protokol no ve sağ diz raporunda "Sol dizde".
Sahte model ilk çıktıda metinde olmayan bir alıntı yazar, "8 mm"yi "18 mm" yapar, "Minimal efüzyon"u "Belirgin sıvı
birikimi" yapar ve bir bulguyu iki kez yazar. Kodun dört sorunu da yakaladığı; hasta adı, TC kimlik no, protokol no ve
doğum tarihinin modele gönderilmediği doğrulanır.

```bash
python -m unittest discover -s tests
```

## Kaynaklar ve dayanaklar

Kontrol tarihi: 15.10.2026. Mevzuat sık değişir; paketteki süre, oran ve kuralları kullanmadan önce güncel resmî metinle
karşılaştırın. "Karşılaştırılan maddeler" dışındaki atıflar yalnız yol göstericidir.

| Kaynak | Bağlantı | Durum |
|---|---|---|
| 6698 sayılı Kişisel Verilerin Korunması Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.6698.pdf) | Resmî metinle karşılaştırılan maddeler: 13 |

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
- **Yapay zekâ hata yapabilir.** Model yanlış, eksik veya uydurma bilgi üretebilir. Kod kontrolleri bazı hataları
  yakalar, hepsini yakalayamaz. Her taslağı okuyup onaylamadan kullanmayın, göndermeyin, imzalamayın.
- **Maskeleme sınırlıdır.** Yalnız telefon, e-posta, TCKN, IBAN ve LinkedIn adresi kalıpları maskelenir. Ad soyad,
  adres, doğum tarihi, sağlık bilgisi ve serbest metindeki diğer kişisel veriler modele gidebilir. Hangi verinin bulut
  sağlayıcısına gönderildiğinden siz sorumlusunuz; gizli veride yerel model (Ollama) kullanın.
- **API ücretleri, kotalar ve API anahtarının güvenliği** size aittir.

Ayrıntılar: `LICENSE` (MIT) ve `SORUMLULUK_REDDI.md`.
