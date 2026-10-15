# Kümül Risk Kontrolü · Kod Bloğu

> Reasürans › Reasürans Uzmanı · Workers / Workless

Yürürlükteki poliçelerin deprem bedellerini **kümül bölgesi** ve **lokasyon (aynı adres)** bazında toplar; sizin
verdiğiniz kapasitelerle karşılaştırıp aşım ve eşiğe yakın durumları listeler. İnternete bağlanmaz.

## Nasıl hesaplar?

```
Şirket payı bedeli = sigorta bedeli × şirket payı %          (koasürans; boşsa %100)
Bölge kümülü       = bölgedeki yürürlükteki poliçelerin şirket payı bedelleri toplamı
PML tutarı         = bölge kümülü × PML %                    (boşsa %100)
Kullanım %         = PML tutarı ÷ bölge kapasitesi × 100
Lokasyon kullanımı = lokasyondaki toplam şirket payı bedeli ÷ lokasyon kapasitesi × 100
```

| Durum | Koşul |
|---|---|
| Aşım | Kullanım > %100 |
| Eşiğe yakın | Kullanım ≥ `--uyari` (varsayılan %80) |
| Uygun | Altı |
| Kapasite tanımsız | Bölge kapasite dosyasında yok |

- **Yürürlükte:** başlangıç ≤ rapor tarihi < bitiş. Süresi bitmiş ve henüz başlamamış poliçeler alınmaz; sayıları
  raporun başında yazılır.
- **Lokasyon:** `Adres Kodu` doluysa onunla, değilse il + ilçe + adres metniyle eşleştirilir (büyük-küçük harf,
  Türkçe karakter ve noktalama ayrımı yapılmaz). Birden çok poliçesi olan veya eşiğe ulaşan lokasyonlar listelenir.
- **Örnek:** bölge kümülü 2.300 milyon TL, PML %15 → 345 milyon TL; kapasite 400 milyon TL → kullanım %86,25.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri (rapor tarihi 01.10.2026)
python main.py --policeler policeler.xlsx --kapasiteler kapasiteler.xlsx --lokasyon-kapasite 1000000000 --tarih 31.12.2026
```

| Dosya | Sütunlar |
|---|---|
| Poliçeler | Poliçe No, Deprem Bedeli (zorunlu); Sigortalı, Kümül Bölgesi, İl, İlçe, Adres, Adres Kodu, Şirket Payı %, Başlangıç, Bitiş |
| Kapasiteler | Kümül Bölgesi, Kapasite, PML % |

`--lokasyon-kapasite` verilmezse lokasyon sayfasında yalnız birden çok poliçesi olan adresler, kapasite
karşılaştırması olmadan listelenir.

## Çıktı

`cikti/kumul_risk_kontrolu.xlsx`:

- `Bölge Kümülü`: bölge bazında toplam bedel, PML tutarı, kapasite, kullanım, boş kapasite ve durum.
- `Lokasyon Kümülü`: aynı adresteki poliçeler, toplam bedel, kullanım ve aşım tutarı.
- `Poliçeler`: kümüle giren poliçeler ve şirket payı bedelleri.
- `Uyarılar`: kapasitesi tanımsız bölgeler, okunamayan satırlar, adresi olmayan poliçeler.

## Dikkat

- **Bölge, PML ve kapasite sizindir.** Kümül bölgesi tanımları, PML oranları ve kapasite tutarları trete
  şartlarınıza ve modelleme çalışmanıza göre girilmelidir; örnek dosyadaki değerler kurgusaldır.
- **Deprem modellemesi değildir.** PML tutarı, girdiğiniz tek bir oranla çarpımdır; yapı tarzı, zemin, kat sayısı,
  muafiyet ve koasürans (deprem) şartları ayrıca hesaba katılmaz.
- **Adres eşleşmesi yazıma bağlıdır.** "4. Cadde" ile "4. Cad." farklı lokasyon sayılır. Sağlıklı sonuç için adres
  kodu (ör. UAVT) kullanın; komşu parsellerdeki riskler de aynı lokasyon olarak yakalanmaz.
- **Kapasite tanımı:** Karşılaştırdığınız tutarın (brüt bedel, şirket payı, konservasyon sonrası) kapasite
  tanımınızla aynı olduğundan emin olun. İhtiyari reasürans devirleri düşülmez.
- **Örnek veri:** Sigortalılar, adresler, bedeller ve kapasiteler kurgusaldır.

## Testler

- Bölge kümülü, PML, kullanım ve durum sınırları (tam kapasite aşım değildir; uyarı eşiği dahildir).
- Lokasyon kümülü: aynı adresin farklı yazımı, tek poliçeyle aşım, kapasitesiz liste.
- Yürürlük filtresi, şirket payı, okuma hataları; örnek veri uçtan uca.

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
