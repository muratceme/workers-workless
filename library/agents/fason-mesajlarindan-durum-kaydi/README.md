# Fason Mesajlarından Durum Kaydı · AI Agent

> Tekstil › Fason Takip Sorumlusu · Gerisi Bende

Fason atölyelerden gelen kısa mesajları (WhatsApp, SMS, e-posta) iş takip kaydına dönüştürür:
- **Dönüş:** teslim edilen mal (tarih, adet, hatalı adet, kalite),
- **Söz:** atölyenin verdiği teslim tarihi,
- **Sorun:** arıza, malzeme eksiği, kalite sorunu.

Kalan adedi ve termin farkını kod hesaplar. Dönüşler **Fason Atölye İş Takibi** kod bloğunun dönüş dosyası biçiminde
yazılır.

## Nasıl çalışır?

1. **Hazırlık (kod):** Açık işler, önceki dönüşler ve mesajlar okunur. Telefon, e-posta, TCKN ve `--gizle` adları
   maskelenir. Kapanmış işler modele gönderilmez.
2. **Kayıt (model):** Her mesajdan bir ya da daha çok kayıt çıkarılır. Göreli tarihler mesaj tarihine göre çevrilir;
   hesap yapılmaz.
3. **Denetim (kod):** Aşağıdakilerden biri varsa yalnız o mesaj modele geri gönderilir (varsayılan en fazla 2 tur):
   - İş no listede ya da mesajda yok; mesajda geçen bir iş kayda alınmamış
   - Adet / hatalı adet mesajda yazmıyor; hatalı adet dönen adetten büyük
   - Tarih ne mesajda yazıyor ne de "bugün, yarın, cuma, pazartesi, haftaya, ayın 14'ü" ifadesinden çıkıyor (kod bu
     ifadeleri mesaj tarihine göre kendisi çözer)
   - Dönüş mesaj tarihinden sonra ya da kanıtta tamamlanmış teslim ifadesi yok ("göndereceğiz", "getiririz" dönüş
     değildir)
   - Kanıt mesajda aynen yok; söz / sorun kaydında kalite bilgisi
4. **Hesap (kod):**
   - kalan adet (çıkan − önceki dönüşler − bu mesajlardaki dönüşler),
   - söz tarihinin terminden farkı,
   - çıkan adedi aşan dönüş toplamı (mükerrer kayıt uyarısı),
   - söz verilen adet ile kalan adedin farkı.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                              # örnek 5 mesaj
python agent.py --mesajlar whatsapp.txt --isler fason_isler.xlsx --donusler donusler.xlsx
```

Mesaj örneği:

```
[07.10.2026 17:40] Atölye B: FS-07 dikim bitti, kalan 900'ü yarın sabah göndereceğiz.
[08.10.2026 11:05] Atölye B: FS-07 900 adet teslim edildi. Sizin kalite 45 tanesinde etek boyu ölçü dışı dedi.
```

## Çıktı

- `cikti/fason_mesaj_kayitlari.xlsx`:
  - `Kayıtlar`: mesaj başına kayıtlar (mor: model), kod notu, son kontrol.
  - `Sözler`: termin, söz tarihi, termine göre gün farkı, kalan adet.
  - `Kontrol`
- `cikti/fason_mesaj_kayitlari_donusler.xlsx`: kod bloğunun dönüş dosyasına eklenecek satırlar.

```bash
python ../../code-blocks/fason-atolye-is-takibi/main.py --isler fason_isler.xlsx --donusler donusler.xlsx
```

## Dikkat

- **Dönüş atölyenin beyanıdır.** Depo girişi ve kalite kontrol sonucuyla karşılaştırmadan hakedişe esas almayın.
- **Göreli tarihler** mesajın gönderildiği güne göre çözülür. Mesaj dökümünde tarih yoksa sonuç yanlış olur.
- **Örnek veri:** Atölyeler, işler ve mesajlar kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk turda "yarın sabah göndereceğiz" sözünü dönüş sayar. Kodun bunu yakalayıp
düzelttirdiği doğrulanır. Ayrıca şunlar doğrulanır:
- göreli tarih çözümü (yarın, cuma, aynı gün adı, ayın N'i, geçmiş gün → gelecek ay),
- iş no, adet, kalite ve kapsam denetimleri,
- kalan adet, termin farkı ve aşım uyarısı,
- kod bloğu dosyası,
- kapanmış işlerin gönderilmemesi.

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
- **Yapay zekâ hata yapabilir.** Model yanlış, eksik veya uydurma bilgi üretebilir. Kod kontrolleri bazı hataları
  yakalar, hepsini yakalayamaz. Her taslağı okuyup onaylamadan kullanmayın, göndermeyin, imzalamayın.
- **Maskeleme sınırlıdır.** Yalnız telefon, e-posta, TCKN, IBAN ve LinkedIn adresi kalıpları maskelenir. Ad soyad,
  adres, doğum tarihi, sağlık bilgisi ve serbest metindeki diğer kişisel veriler modele gidebilir. Hangi verinin bulut
  sağlayıcısına gönderildiğinden siz sorumlusunuz; gizli veride yerel model (Ollama) kullanın.
- **API ücretleri, kotalar ve API anahtarının güvenliği** size aittir.

Ayrıntılar: `LICENSE` (MIT) ve `SORUMLULUK_REDDI.md`.
