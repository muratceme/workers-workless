# Destek Talebi Sınıflandırma · AI Agent

> Bilgi Teknolojileri › BT Destek Uzmanı · Workers / Workless

BT hizmet masasına e-posta veya ticket sisteminden gelen talepleri **kategori, ekip, etki, aciliyet ve türe**
(olay / hizmet talebi / güvenlik olayı) göre sınıflandırır. **ITIL önceliğini** ve **SLA hedef zamanlarını**
hesaplar, kullanıcıya ilk yanıt taslağı ve destek uzmanına çözüm önerisi yazar.

## Nasıl çalışır?

1. **Kod (yapay zekâ yok):**
   - **Güvenlik olayı belirtileri:** oltalama bağlantısına tıklama, şifre girme, virüs, fidye yazılımı, şüpheli
     giriş
   - **Olası yaygın kesinti:** Aynı konuda (ağ/internet, e-posta, ERP) kısa sürede gelen çok sayıda talep;
     varsayılan 30 dakikada 4 ve üzeri. Bu talepler tek bir **ana olay** altında gruplanır.
   - E-posta, telefon, IBAN ve TCKN maskelenir; gönderen bilgisi modele gitmez.
2. **Model:** Her talebi sınıflandırır, etki ve aciliyeti belirler, ilk yanıtı ve çözüm önerisini yazar. Şifre
   sıfırlamada kullanıcıdan şifre istemez. Güvenlik olayında güvenli ilk adımları verir.
3. **Kod, öncelik ve SLA:**
   - **ITIL etki × aciliyet matrisi** (P1–P5) uygulanır ve `sla.json`'dan ilk yanıt ve çözüm hedef zamanları
     hesaplanır.
   - Kod kuralları modelin üstündedir: güvenlik olayı her zaman **Bilgi Güvenliği** ekibine gider ve en az
     **P2** olur. Yaygın kesinti kümesindeki talepler de en az P2 olur.

| Etki \ Aciliyet | Yüksek | Orta | Düşük |
|---|---|---|---|
| Yüksek | P1 | P2 | P3 |
| Orta | P2 | P3 | P4 |
| Düşük | P3 | P4 | P5 |

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                   # örnek 10 talep (5'i aynı internet kesintisi)
python agent.py --girdi talepler.xlsx --sla sla.json --pencere 20 --esik 3
```

## Çıktı

`Yaygın Kesinti` (ana olay önerisi) · `Talepler` (önceliğe göre sıralı, "Onay") · `Özet`

## Testler

Testler gerçek API çağırmaz. Sahte model bilerek düşük öncelik verir. Kodun güvenlik olayını ve yaygın kesintiyi
P2'ye yükselttiği ve SLA zamanlarını doğru hesapladığı doğrulanır.

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
