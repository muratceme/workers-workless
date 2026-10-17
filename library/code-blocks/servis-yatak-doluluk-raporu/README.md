# Servis Yatak Doluluk Raporu · Kod Bloğu

> Tıbbi Sekreterlik › Yatan Hasta Tıbbi Sekreteri · Gerisi Bende

Yatış kayıtlarından **servis bazında yatak doluluk oranını, ortalama kalış süresini ve yatak devir hızını**
hesaplar; rapor tarihinde yatan hastaları, boş yatakları ve **beklenen taburcu listesini** çıkarır. İnternete ve
HBYS'ye bağlanmaz; dışa aktarılmış dosyaları okur.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Hasta günü | Dönem içinde hastanın yatakta geçirdiği gece sayısı; aynı gün yatış-taburcu 1 gün sayılır |
| Yatak doluluk oranı % | Hasta günü ÷ (yatak sayısı × dönemin gün sayısı) × 100 |
| Ortalama kalış süresi | Dönemin hasta günü ÷ dönemde taburcu edilen hasta sayısı |
| Taburcu olanların ort. yatışı | Dönemde taburcu edilenlerin yatıştan taburcuya gün sayısı ortalaması |
| Yatak devir hızı | Dönemde taburcu edilen hasta ÷ yatak sayısı |
| Şu an yatan / boş yatak | Rapor tarihinde yatışı başlamış, taburcu tarihi boş ya da rapor tarihinden sonra olan hastalar |

Dönem varsayılan olarak rapor tarihinin ayının ilk gününden rapor tarihine kadardır. Taburcu tarihi boş hastalar dönem
sonuna kadar yatıyor sayılır.

**Uyarılar:** doluluk `--yuksek` (varsayılan %90) üstünde · `--dusuk` (varsayılan %50) altında (bilgi) · yatan hasta
yatak sayısından fazla · planlanan taburcu tarihi geçmiş · `--uzun` (varsayılan 14) günden uzun yatış · yatak sayısı
tanımlı olmayan servis · yatışı olmayan servis.

**Örnek:** Dahiliye 12 yatak, eylülün 30 gününde 312 hasta günü → doluluk 312 ÷ (12 × 30) × 100 = **%86,67**; 52 taburcu
→ ortalama kalış 312 ÷ 52 = **6,00 gün**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                                      # örnek veri (30.09.2026)
python main.py --servisler servisler.xlsx --yatislar yatislar.xlsx --tarih 30.09.2026
python main.py --servisler servisler.xlsx --yatislar yatislar.xlsx --baslangic 01.07.2026 --bitis 30.09.2026 --tarih 30.09.2026 --ileri 2
```

Zorunlu sütunlar: servislerde Servis ve Yatak Sayısı; yatışlarda Servis ve Yatış Tarihi (Taburcu Tarihi yatan
hastalarda boş).

## Çıktı

`cikti/servis_yatak_doluluk.xlsx`:

- `Servisler`: servis bazında göstergeler ve toplam; doluluk renkli.
- `Beklenen Taburcular`: planlanan taburcusu rapor tarihi + `--ileri` gün içinde olanlar; tarihi geçenler kırmızı.
- `Yatan Hastalar`: rapor tarihinde yatan hastalar ve yattığı gün.
- `Uyarılar`.

## Dikkat

- **Tanımları kurumunuzla doğrulayın.** Hasta günü sayımı (gece yarısı sayımı, aynı gün yatış-çıkış), çıkan hasta
  tanımı (ölüm ve servisler arası nakil dahil mi) ve yatak sayısı (aktif / ruhsatlı / ek yatak) kurumun ve Sağlık
  Bakanlığı'nın güncel istatistik tanımlarına göre farklı olabilir. Kod, servisler arası nakli ayrı bir yatış satırı
  olarak yazmanızı bekler.
- **Yoğun bakım, günübirlik ve gözlem yatakları** ayrı servis olarak tanımlanmalıdır; karıştırılırsa oranlar bozulur.
- **Hasta verisi içerir.** Protokol numarası ve hasta adı özel nitelikli kişisel veriyle ilişkilidir; dosyayı yalnız
  yetkili kişilerle paylaşın, gerekmiyorsa hasta adı sütununu çıkarın.
- **Planlanan taburcu tahminidir;** liste hekim kararının yerine geçmez.
- **Örnek veri:** Servisler, protokol numaraları, baş harfler ve tarihler kurgusaldır.

## Testler

- Hasta günü (aynı gün, dönem dışı, yatmaya devam eden), doluluk, ortalama kalış, devir hızı, anlık durum ve toplam.
- Beklenen ve tarihi geçen taburcu, uzun yatış, yatak sayısını aşan hasta, tanımsız servis, okuma ve komut satırı.

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
