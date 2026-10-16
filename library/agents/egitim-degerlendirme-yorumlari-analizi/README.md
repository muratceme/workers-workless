# Eğitim Değerlendirme Yorumları Analizi · AI Agent

> İnsan Kaynakları › Eğitim ve Gelişim Uzmanı · Workers / Workless

Eğitim sonrası değerlendirme anketlerindeki **açık uçlu yorumları** eğitim bazında okur:
- yorumları temaya (içerik, uygulama, eğitmen, süre / tempo, materyal, ortam, organizasyon) ve yöne (Olumlu / Olumsuz /
  Öneri) ayırır,
- güçlü yönleri ve geliştirme alanlarını, dayandığı yorumlarla çıkarır,
- eğitmene yapıcı bir geri bildirim notu yazar.

Puan tarafını (katılım, ön / son test, memnuniyet) **Eğitim Değerlendirme Raporu** kod bloğu hesaplar. Aynı dosyaya yorum
sütunlarını eklemeniz yeterli.

## Nasıl çalışır?

1. **Hazırlık (kod):** Anket okunur. Katılımcı sütunu modele gönderilmez; yorumlarda geçen katılımcı ve eğitmen adları
   takma adla, telefon ve e-posta maskelenir. Yorum yazan yanıt sayısı `--min-yanit`'tan (varsayılan 3) az olan eğitim
   anonimlik için analiz edilmez.
2. **İstatistik (kod):** Eğitim bazında yanıt ve yorum sayısı, memnuniyet ortalaması, 4–5 verenlerin oranı.
3. **Etiket ve not (model):** Her yorum tema ve yöne ayrılır; benzer yorumlar başlık altında toplanır, öneriler ve
   eğitmen notu yazılır.
4. **Denetim (kod):** Aşağıdakilerden biri varsa yalnız o eğitim modele geri gönderilir (varsayılan en fazla 2 tur):
   - Etiketlenmemiş, birden çok kez etiketlenmiş ya da bilinmeyen yorum
   - Olumsuz / öneri içeren bir yorumun hiçbir geliştirme alanında olmaması
   - Güçlü yönün olumlu olmayan, geliştirme alanının olumsuz / öneri olmayan yoruma dayanması
   - 1–2 puan veren yanıtın yorumunun yalnız olumlu sayılması
   - Notta istatistikte ya da yorumlarda olmayan sayı; notta katılımcı adı
5. **Sayım (kod):** Tema tablosu, yön sayıları ve geliştirme alanlarının yorum sayısına göre sırası kodla yapılır. Tek
   yoruma dayanan alan "Tekil yorum" diye işaretlenir.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                              # örnek 3 eğitim (biri az yanıt nedeniyle atlanır)
python agent.py --anket degerlendirme.xlsx --min-yanit 5
```

| Sütun | Açıklama |
|---|---|
| Eğitim | Zorunlu |
| Eğitmen, Katılımcı, Memnuniyet | İsteğe bağlı (Katılımcı yalnız maskeleme için okunur) |
| Yorum sütunları | Başlığında "yorum", "öneri", "beğen", "geliştir", "görüş" geçen her sütun |

## Çıktı

`cikti/egitim_yorum_analizi.xlsx`:
- `Özet`: eğitim bazında yanıt, yorum, memnuniyet, olumlu / olumsuz / öneri içeren yorum sayısı, ilk 3 geliştirme alanı.
- `Geliştirme Alanları`: başlık, yorum sayısı, öneri ve dayandığı yorumlar.
- `Güçlü Yönler`, `Yorumlar` (tema / yön etiketleriyle), `Tema Tablosu`, `Eğitmen Notları`, `Kontrol`.

Raporda katılımcı adı yer almaz; yorumlarda geçen katılımcı adları takma adla kalır. Mor hücreler modelin, diğerleri
kodundur.

## Dikkat

- **Anonimlik:** Küçük gruplarda yorumun üslubu ya da içeriği kişiyi tanıtabilir. Raporu eğitmenle paylaşmadan önce
  okuyun; gerekirse `--min-yanit` değerini yükseltin. Yalnız tam ad soyad maskelenir; yorumda geçen tek bir ad
  maskelenmeyebilir.
- **Tekil yorum genelleme değildir.** Tek yoruma dayanan alanı karar verirken ayrıca tartın.
- **Puan ile yorum çelişebilir.** Etiketler yorumun içeriğine göredir; memnuniyet ortalaması ayrı okunmalıdır.
- **Örnek veri:** Kişiler, eğitimler ve yorumlar kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk turda İleri Excel'in salon / eski Excel yorumunu olumsuz etiketler ama
hiçbir geliştirme alanına koymaz. Kodun bunu yakaladığı doğrulanır. Ayrıca şunlar doğrulanır:
- etiket, güçlü yön ve geliştirme alanı denetimleri,
- düşük puan, sayı ve takma ad kuralları,
- yorum bazlı sayım,
- az yanıtlı eğitimin gönderilmemesi,
- katılımcı ve eğitmen adlarının mesajlara ve rapora gitmemesi.

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
