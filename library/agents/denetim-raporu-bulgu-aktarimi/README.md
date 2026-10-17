# Denetim Raporu Bulgu Aktarımı · AI Agent

> Tekstil Uygunluk › Sosyal Uygunluk Uzmanı · Gerisi Bende

Sosyal uygunluk, çevre ya da müşteri denetim raporlarındaki **bulguları düzeltici faaliyet takip tablosuna** aktarır.
Her bulgu için şunlar yazılır: konu, kısa bulgu metni, raporun derecesi, rapordan alıntı, düzeltici faaliyet önerisi ve
sorumlu. Raporun derecesi değiştirilmez. Termini ve sıfır tolerans işaretini kod belirler. Çıktı **Denetim Bulgu ve
Düzeltici Faaliyet Takibi** kod bloğuna doğrudan girer.

## Nasıl çalışır?

1. **Hazırlık (kod):** Rapor okunur. Telefon, e-posta ve kimlik numarası maskelenir; `--gizle` adları takma adla gider.
2. **Çıkarım (model):** denetim türü, tarihi, tesis; her bulgu için konu, bulgu, rapordaki derece ve karşılığı (Kritik /
   Majör / Minör / Gözlem), alıntı, düzeltici faaliyet önerisi, sorumlu görev.
3. **Denetim (kod):** Aşağıdakilerden biri varsa çıkarım modele geri gönderilir (varsayılan en fazla 2 tur):
   - Alıntı raporda aynen geçmiyor
   - Derece rapordaki dereceyle aynı değil: model kendi görüşüyle yükseltemez ya da düşüremez
   - Denetim tarihi ya da bulgudaki bir sayı raporda yok
   - Düzeltici faaliyet önerisi boş
4. **Değerlendirme (kod):**
   - **Termin** = denetim tarihi + dereceye göre gün. Varsayılan: Kritik 7, Majör 30, Minör 60, derecesi belirtilmemiş 30.
     Müşterinizin veya denetim kuruluşunun sürelerini `--termin` ile verin.
   - **Sıfır tolerans konuları** raporun derecesinden bağımsız işaretlenir ve termini Kritik süresiyle hesaplanır:
     çocuk / genç işçi, zorla çalıştırma, kimlik belgesinin alıkonması, ücretin ödenmemesi ya da asgari ücret altı,
     rüşvet, çift kayıt, kilitli çıkış, fiziksel şiddet, taciz.
   - **Gözlemler** (iyi uygulamalar) takip listesine alınmaz, Ayrıntı sayfasında gösterilir.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                     # örnek rapor (4 bulgu, 1 gözlem)
python agent.py --rapor denetim.pdf --onek BSCI-26 --termin "kritik=3,majör=30,minör=90"
```

Ardından takip için:

```bash
python main.py --bulgular cikti/denetim_bulgulari.xlsx     # Denetim Bulgu ve Düzeltici Faaliyet Takibi kod bloğu klasöründe
```

## Çıktı

`cikti/denetim_bulgulari.xlsx`:
- `Bulgular`: takip bloğunun sütunları. Bulgu No, Denetim, Denetim Tarihi, Tesis, Konu, Bulgu, Derece, Düzeltici
  Faaliyet, Sorumlu, Termin, Durum (Açık), Tamamlanma Tarihi, Kanıt.
- `Ayrıntı`: rapordaki derece, rapordan alıntı, sıfır tolerans işareti, uyarılar (gözlemler dahil).
- `Kontrol`

## Dikkat

- **Düzeltici faaliyetler öneridir.** Kök neden analizi yapılıp sorumlularla birlikte kesinleştirilmelidir. Müşteriye
  gönderilecek CAP (düzeltici faaliyet planı) için müşterinin formatını kullanın.
- **Sıfır tolerans listesi** sınırlıdır ve anahtar kelimeyle çalışır; müşterinizin davranış kurallarındaki listeyle
  karşılaştırın. Bu konularda müşteri bildirimi ve yasal yükümlülükler (ör. çocuk işçi) için hukuk desteği alın.
- **Kişisel veri.** Raporda çalışan adları geçiyorsa `--gizle` ile verin.
- **Örnek veri:** Tesis, denetim kuruluşu ve bulgular kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk turda genç işçi bulgusunun derecesini rapordaki "Majör"den "Kritik"e
yükseltir. Kodun bunu yakaladığı doğrulanır; dereceyi değiştirmeden sıfır tolerans işaretini ve kısa termini kod koyar.
Ayrıca şunlar doğrulanır:
- alıntı, tarih ve sayı denetimleri,
- termin hesabı, gözlemlerin takip dışı kalması,
- çıktının takip bloğu sütunlarıyla aynı olması,
- gizlenen adın mesaja gitmemesi.

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
