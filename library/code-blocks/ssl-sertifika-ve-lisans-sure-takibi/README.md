# SSL Sertifika ve Lisans Süre Takibi · Kod Bloğu

> Bilgi Teknolojileri › Sistem Uzmanı · Workers / Workless

SSL / TLS sertifikası, alan adı, yazılım lisansı, abonelik ve destek sözleşmelerinin **bitiş tarihlerini** izler. Süresi dolan ve
yaklaşanları, lisans adedini aşan ya da atıl kalan kullanımı listeler; önümüzdeki 12 ayın yenileme maliyetini çıkarır. İnternete
bağlanmaz, sertifikaları sunucudan okumaz; envanter listesini okur.

## Nasıl hesaplar?

| Gösterge | Kural |
|---|---|
| Kalan gün | Bitiş tarihi − rapor günü |
| Durum | Süresi doldu · Kritik (≤ `--kritik` 7 gün) · Yaklaşıyor (SSL ve alan adı için ≤ `--yakin` 30, lisans ve sözleşmeler için ≤ `--izle` 90 gün) · İzlemede (SSL / alan adı ≤ 90 gün) · Geçerli |
| Kullanım % | Kullanılan ÷ lisans adedi × 100; %100 üstü uyum riski, `--atil` (%60) altı atıl lisans |
| Yenileme takvimi | Önümüzdeki 12 ayda bitecek varlıkların yıllık maliyeti, ay ve para birimi bazında |

Tür metinden tanınır: SSL / TLS / sertifika · alan adı / domain · destek / bakım · abonelik / SaaS / bulut · lisans (boş = yazılım
lisansı).

**Uyarılar:** süresi doldu · bitişe 7 günden az · lisans adedi aşıldı · okunamayan satır (Yüksek) · bitiş yaklaşıyor (Orta) · atıl
lisans · yaklaşan varlıkta sorumlu yok (Bilgi). Otomatik yenilemesi açık varlıklarda ödeme aracının ve yenilemenin
gerçekleştiğinin doğrulanması hatırlatılır.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                          # örnek veri (14.10.2026)
python main.py --varliklar varliklar.xlsx
python main.py --varliklar varliklar.xlsx --kritik 3 --yakin 14 --izle 60 --atil 50
```

Zamanlanmış görevle (ör. her pazartesi) çalıştırıp çıktıyı e-postayla göndermek için işletim sisteminin zamanlayıcısını kullanın.

## Çıktı

`cikti/sertifika_lisans_takibi.xlsx`: `Varlıklar` (en acil üstte), `Yenileme Takvimi`, `Uyarılar`.

## Dikkat

- **Envanter güncel olmalı.** Rapor yalnız listede olanları izler; bitiş tarihlerini sertifika ve sağlayıcı panellerinden düzenli
  olarak aktarın.
- **Sertifika geçerlilik süreleri kısalıyor.** Tarayıcı ve sertifika otoritelerinin ortak kuralları sertifika ömrünü kademeli olarak
  kısaltıyor; elle yenilenen sertifikalarda otomatik yenilemeye (ACME vb.) geçmeyi değerlendirin. Güncel azami süreyi sertifika
  sağlayıcınızdan teyit edin.
- **Lisans uyumu** sözleşme koşullarına bağlıdır (kullanıcı, cihaz, çekirdek başına); kullanım sayısını lisans modeline uygun sayın.
- **Örnek veri:** Alan adları, sağlayıcılar ve tutarlar kurgusaldır.

## Testler

- Durum eşikleri (SSL / alan adı ve lisans için ayrı), sıralama, kullanım %, aşım ve atıl lisans, 12 aylık yenileme takvimi.
- Tür tanıma, uyarılar, okuma ve komut satırı.

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
