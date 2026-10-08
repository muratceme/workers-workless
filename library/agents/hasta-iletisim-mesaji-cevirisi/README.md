# Hasta İletişim Mesajı Çevirisi · AI Agent

> Sağlık (Özel Hastane) › Uluslararası Hasta Hizmetleri › Uluslararası Hasta Koordinatörü · Workers / Workless

Yurt dışından gelen hasta mesajlarını tıbbi terimleri, dozları ve değerleri koruyarak Türkçeye çevirir. Mesajları
aciliyete göre sıralar ve hastanın kendi dilinde, kültürüne uygun bir cevap taslağı hazırlar. Kod çeviriyi ve
cevabı denetler. **Taslaktır: koordinatör, tıbbi içerikte doktor kontrol etmeden gönderilmez.** Kendi API
anahtarınızla çalışır.

## Nasıl çalışır?

1. **Girdiler:** `mesajlar/` klasöründeki her dosya bir mesajdır (.txt, .docx, .pdf, .eml). Başta "Hasta:" ve
   "Tarih:" satırları olabilir. Hastanın adı farklı alfabelerde yazılmışsa ";" ile ayırın, ör. `Hasta: Ivan Primer;
   Иван Пример`. `kurum_bilgileri.txt` ve `terimce.csv` de okunur.
2. **Maskeleme:**
   - Hasta adları ve adın parçaları `[HASTA-n]` olarak gönderilir.
   - Uluslararası telefon, e-posta, pasaport numarası, TCKN ve IBAN maskelenir.
   - Sağlık verisi özel nitelikli kişisel veri olduğu için gönderimden önce açık onay istenir.
3. **Model her mesaj için yazar:**
   - dil
   - Türkçe çeviri
   - terimler ve hastanın talepleri
   - aciliyet: Acil / Öncelikli / Rutin
   - hastanın dilinde cevap taslağı ve cevabın Türkçe karşılığı
   - koordinatörün dolduracağı yer tutucular, ör. `[FİYAT]`, `[TARİH]`
   - doktora iletilecek tıbbi hususlar
4. **Modele konan sınırlar:**
   - **Tıbbi tavsiye vermez:** Tedavi uygunluğu, ilaç gibi konularda "doktorumuz değerlendirecek" der.
   - **Acil durumda:** Hastayı ilk cümlede bulunduğu yerdeki acil servise yönlendirir.
   - **Uydurmaz:** Kurum bilgilerinde olmayan fiyat, tarih ve süre yazmaz; yer tutucu koyar.
5. **Kod denetler:**

| Kontrol | Önem |
|---|---|
| Doz / ölçü değeri (500 mg, 38,9 °C, 5 мг…) çeviride aynı sayı ve birimle yok | Yüksek |
| Çeviride kırmızı bayrak ifadesi var ama aciliyet "Acil" değil. İfadeler: ateş, nefes darlığı, göğüs ağrısı, akıntı, kanama, bilinç, bayılma, nöbet, şiddetli ağrı… | Yüksek |
| Acil mesajın cevabı acil servise / acil numaraya yönlendirmiyor | Yüksek |
| Cevapta mesajda ve kurum bilgilerinde olmayan sayı var (uydurma fiyat / tarih / süre) | Yüksek |
| Terimcedeki terim mesajda geçiyor ama Türkçe karşılığı çeviride yok | Orta |
| Cevap ile Türkçe karşılığının sayıları farklı | Orta |
| Mesajın alfabesi (Kiril, Arap) ile modelin bildirdiği dil uyumsuz | Orta |
| Doldurulacak yer tutucular; doktora iletilecek husus | Bilgi |

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # kendi API anahtarınızı yazın; sağlık verisi için WW_PROVIDER=ollama önerilir
```

## Kullanım

```bash
python agent.py                                          # örnek: 4 kurgusal mesaj (İngilizce, Almanca, Rusça)
python agent.py --girdi ./gelen_kutusu --gizle "Refakatçi Adı" "Dr. Ad Soyad"
```

| Dosya | İçerik |
|---|---|
| `mesajlar/*.txt` (.docx, .pdf, .eml) | Hasta mesajı; isteğe bağlı "Hasta:" ve "Tarih:" başlık satırları |
| `kurum_bilgileri.txt` | Cevapta kullanılabilecek kurum bilgileri: transfer, tercümanlık, süreç, istenen tetkikler |
| `terimce.csv` (isteğe bağlı) | Kaynak Terim, Türkçe |

## Çıktı

- `hasta_mesajlari.xlsx`:
  - `Mesajlar`: aciliyete göre sıralı. Çeviri, talepler, cevap taslağı ve Türkçesi, doldurulacaklar, doktora
    iletilecek hususlar, kontroller ve boş "Koordinatör Onayı" sütunu.
  - `Terimler`.
  - `Kontroller`.
- `hasta_mesajlari_cevaplar/<mesaj>.md`: her mesaj için Türkçe çeviri, talepler, cevap taslağı, Türkçe karşılığı,
  koordinatörün dolduracakları ve kontroller.
- **Takma adlar:** Çıktılarda takma adlar gerçek adlarla değiştirilmiştir.

## Dikkat

- **Sağlık verisi:** Hasta mesajları özel nitelikli kişisel veridir (KVKK md. 6). Bir bulut API'sine göndermeden önce
  şu ikisini kontrol edin:
  - hastanın açık rızası
  - kurumunuzun yurt dışına aktarım politikası

  Mümkünse yerel model (Ollama) kullanın. Maskeleme yalnız tanınan kalıpları gizler; mesajı göndermeden önce
  gözden geçirin.
- **Tıbbi sorumluluk:**
  - **Kararlar:** Aciliyet sınıflaması ve çeviri yardımcıdır; tıbbi değerlendirme ve hastaya verilen her tıbbi
    bilgi doktorun sorumluluğundadır.
  - **Acil şikâyetler:** Kodun kırmızı bayrak listesi sınırlıdır. Acil şikâyetli her mesaj hemen nöbetçi doktora
    iletilmelidir.
- **Çeviri kalitesi:** Dozlar ve sayılar kodla denetlenir; ifade inceliklerini ve kültürel uygunluğu bir tıbbi
  tercüman kontrol etmelidir.
- **Örnek veri:** Örnek hastalar, mesajlar ve kurum bilgileri kurgusaldır.

## Testler

Testler gerçek API çağırmaz; model yanıtı sahte bir fonksiyonla verilir. Şunlar test edilir:
- Okuma; ölçü ve sayı çıkarma.
- Maskeleme: ad, uluslararası telefon, pasaport.
- Doz koruma; uydurma sayı; kırmızı bayrak / aciliyet uyuşmazlığı; terimce; alfabe-dil uyumu.
- Çıktılar ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
