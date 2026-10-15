# Çek Takas Kontrol Listesi · Kod Bloğu

> Şube Bankacılığı › Operasyon Yetkilisi · Workers / Workless

Takasa verilecek ve takastan gelen çekler için **gün sonu kontrol listesi** üretir: ileri tarihli, ibraz süresi
geçmiş, tutarı tutarsız, karekodu, cirosu ya da imzası sorunlu ve listede iki kez geçen çekleri işaretler. İnternete
bağlanmaz.

## Nasıl hesaplar?

| Kontrol | Kural | Sonuç |
|---|---|---|
| İleri tarihli | Keşide tarihi işlem tarihinden sonra | İşleme alma |
| İbraz süresi | Son gün = keşide tarihi + `--ibraz-gun` (10) gün; "Farklı Yer" Evet ise + `--ibraz-ay` (1) ay | Geçmişse İncele |
| Tutar tutarlılığı | Rakamla ve yazıyla yazılan tutar aynı olmalı | Farklıysa İncele |
| Karekod | "Yok" ya da "Okunamadı" | İşleme alma |
| Ciro silsilesi | "Kopuk" ya da "Eksik" | İşleme alma |
| İmza | "Yok" | İşleme alma |
| Mükerrer | Aynı banka ve çek numarası listede birden çok kez | İşleme alma |
| İşaretlenmemiş kontrol | Karekod, ciro, imza ya da yazıyla tutar boş | İncele |

**Örnek:** Keşide tarihi 28.09.2026, aynı yerde ödenecek → ibraz için son gün **08.10.2026**; işlem tarihi 12.10.2026
→ 4 gün geçmiş: **İncele**. Aynı çek "Farklı Yer: Evet" olsaydı son gün 28.10.2026 olur, sorun çıkmazdı.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                            # örnek veri, işlem tarihi 12.10.2026
python main.py --cekler cekler.xlsx
python main.py --cekler cekler.xlsx --tarih 15.10.2026 --ibraz-gun 10 --ibraz-ay 1
```

Zorunlu sütunlar: Çek No, Keşide Tarihi, Tutar (rakam). Karekod, Ciro Silsilesi ve İmza sütunlarına fiziki
kontrolün sonucunu yazarsınız: Var / Tam / Okundu ya da Yok / Kopuk / Okunamadı.

## Çıktı

`cikti/cek_takas_kontrol_listesi.xlsx`:

- `Kontrol Listesi`: her çek için sonuç, ibraz son günü, kontrol işaretleri, bulgu ve paraf sütunu.
- `Özet`: yön bazında çek sayısı, tutar ve işleme alınmayacak tutar.
- `Bulgular`: her bulgu ayrı satırda, açıklamasıyla.
- `Uyarılar`.

## Dikkat

- **Kod çeki görmez.** Karekodun okunup okunmadığına, imzanın bulunup bulunmadığına ve ciro zincirinin düzgün
  olduğuna siz bakarsınız; kod yazdığınız işaretleri değerlendirir. İmzanın keşideciye ait olup olmadığı, sahtecilik
  ve tahrifat kontrolü bu paketin dışındadır.
- **Süreleri ve ileri tarihli çek kuralını mevzuattan doğrulayın.** İbraz süreleri (10 gün / 1 ay) varsayılan
  olarak Türk Ticaret Kanunu'ndaki genel süreler düşünülerek konmuştur; farklı kıtada düzenlenen çekler, tatile
  rastlayan son gün ve keşide tarihinden önce ibraz yasağına ilişkin dönemsel düzenlemeler hesaba katılmaz.
- **"İncele" ret demek değildir.** İbraz süresi geçmiş ya da tutarı tutarsız çekte ne yapılacağı mevzuata ve
  bankanızın uygulama talimatına göre belirlenir; kod yalnız işaret koyar.
- **Karşılık, ödeme yasağı ve hesap durumu sorgulanmaz.** Bunlar bankacılık sisteminde yapılır.
- **Kapsam dışı:** yabancı para çekler, hamiline çek bildirimi, karşılıksız çek işlemleri ve yasal bildirimler,
  takas odası saatleri ve iade kodları.
- **Gizlilik:** Çek ve banka bilgileri müşteri sırrıdır; raporu yalnız yetkili kişilerle paylaşın.
- **Örnek veri:** Çek numaraları, bankalar ve tutarlar kurgusaldır.

## Testler

- İleri tarihli, bugün tarihli, ibraz süresinin son günü ve bir gün sonrası; başka yerde ödenecek çekte ay ekleme.
- Tutar farkı, karekod / ciro / imza sorunları, boş bırakılan kontroller, mükerrer çek (banka ayrımıyla).
- Sonuç önceliği (engel, dikkat, eksik), yön bazında özet, okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

## Kaynaklar ve dayanaklar

Kontrol tarihi: 15.10.2026. Mevzuat sık değişir; paketteki süre, oran ve kuralları kullanmadan önce güncel resmî metinle
karşılaştırın. "Karşılaştırılan maddeler" dışındaki atıflar yalnız yol göstericidir.

| Kaynak | Bağlantı | Durum |
|---|---|---|
| 6102 sayılı Türk Ticaret Kanunu | [mevzuat.gov.tr](https://www.mevzuat.gov.tr/mevzuatmetin/1.5.6102.pdf) | Resmî metinle karşılaştırılan maddeler: 5/A, 22, 796, 1434 |

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
