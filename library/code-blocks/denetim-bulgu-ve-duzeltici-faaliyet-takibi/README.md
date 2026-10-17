# Denetim Bulgu ve Düzeltici Faaliyet Takibi · Kod Bloğu

> Sürdürülebilirlik ve Sosyal Uygunluk › Sosyal Uygunluk Uzmanı · Gerisi Bende

BSCI, Sedex ve müşteri denetimlerinden gelen bulguları **düzeltici faaliyet planı (CAP), sorumlu ve termin** bazında
takip eder; gecikmiş, planı eksik ve kanıtı eksik bulguları öne çıkarır. İnternete bağlanmaz.

## Nasıl hesaplar?

**Açık bulgu:** durumu `tamamlandı`, `kapandı` veya `doğrulandı` ile başlamayan bulgu.

| Takip durumu | Koşul |
|---|---|
| Gecikmiş | Açık ve termin rapor tarihinden önce |
| Planı eksik | Açık; düzeltici faaliyet, sorumlu veya termin boş |
| Yaklaşıyor | Açık; termine `--yaklasan` (varsayılan 14) gün veya daha az kaldı |
| Zamanında | Açık; termin daha ileride |
| Kanıt eksik | Kapalı; kanıt sütunu boş veya `Yok` |
| Geç tamamlandı | Kapalı; tamamlanma tarihi terminden sonra |
| Tamamlandı | Kapalı; kanıtı var ve zamanında |

- Liste önce takip durumuna, sonra dereceye (kritik → majör → minör → gözlem), sonra gecikme gününe göre sıralanır.
- **Tamamlanma %** = kapalı bulgu ÷ toplam bulgu × 100; denetim, sorumlu, konu ve derece bazında ayrı ayrı verilir.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri (rapor tarihi 12.10.2026)
python main.py --bulgular bulgular.xlsx --tarih 31.10.2026 --yaklasan 30
```

Sütunlar: Bulgu, Durum (zorunlu); Bulgu No, Denetim, Denetim Tarihi, Tesis, Konu, Derece, Düzeltici Faaliyet,
Sorumlu, Termin, Tamamlanma Tarihi, Kanıt.

## Çıktı

`cikti/duzeltici_faaliyet_takibi.xlsx`:

- `Takip Listesi`: öncelik sırasıyla bulgular, kalan gün ve notlar; boş "Güncelleme" sütunu.
- `Denetim Bazında`, `Sorumlu Bazında`, `Konu Bazında`, `Derece Bazında`: toplam, açık, gecikmiş, planı eksik,
  kanıtı eksik bulgu sayıları ve tamamlanma yüzdesi.
- `Uyarılar`: okunamayan tarihler, mükerrer bulgu numaraları.

## Dikkat

- **Standardı yorumlamaz.** Kod BSCI, SMETA veya müşteri davranış kurallarının gereklerini bilmez; denetim notu,
  derece veya geçerlilik süresi hesaplamaz. Bulgu derecesi denetim raporunda ne yazıyorsa odur.
- **Kapanışı denetçi doğrular.** "Tamamlandı" yalnız sizin kaydınızdır; bulgunun kapanmış sayılması denetim
  kuruluşunun veya müşterinin doğrulamasına bağlıdır. Platformlardaki (ör. denetim portalı) terminler esastır.
- **Kanıt içeriği kontrol edilmez.** Kod yalnız kanıt sütununun dolu olup olmadığına bakar.
- **Kök neden:** Düzeltici faaliyetin bulguyu gerçekten giderip gidermediği ve tekrarını önleyip önlemediği ayrıca
  değerlendirilmelidir.
- **Kişisel veri:** Bulgularda çalışan adı veya ücret bilgisi varsa dosyayı sınırlı erişimle saklayın.
- **Örnek veri:** Tesisler, bulgular ve sorumlular kurgusaldır.

## Testler

- Açık bulgularda termin sınırları (dün, bugün, 14. ve 15. gün), planı eksik bulgu.
- Kapalı bulgularda kanıt ve geç tamamlanma; sıralama; denetim ve sorumlu özetleri.
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
