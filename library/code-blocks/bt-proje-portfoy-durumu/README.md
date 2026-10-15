# BT Proje Portföy Durumu · Kod Bloğu

> Bilgi Teknolojileri › BT Müdürü · Workers / Workless

Proje listenizi yönetim toplantısına götürülecek **tek sayfalık kırmızı-sarı-yeşil tabloya** çevirir. Her projenin
takvimde ve bütçede nerede olduğunu kazanılmış değer yöntemiyle hesaplar; rengi kişisel kanaat değil, sayı belirler.
İnternete bağlanmaz.

## Nasıl hesaplar?

| Kavram | Kural |
|---|---|
| Planlanan ilerleme % | (Durum tarihi − başlangıç) ÷ (planlanan bitiş − başlangıç), 0–100 arasında |
| Planlanan değer (PD) | Bütçe × planlanan ilerleme % |
| Kazanılan değer (KD) | Bütçe × gerçekleşen ilerleme % |
| Takvim endeksi (TPE) | KD ÷ PD. 1'in altı: plandan geride |
| Maliyet endeksi (MPE) | KD ÷ harcanan. 1'in altı: yapılan işe göre pahalı |
| Tahmini toplam maliyet | Bütçe ÷ MPE (bugünkü verim sürerse) |

Renk: endeks 0,95 ve üstü **Yeşil**, 0,85 ve üstü **Sarı**, altı **Kırmızı**. Ek kurallar:

- Planlanan bitişi geçmiş ve bitmemiş proje takvimde Kırmızı,
- tahmini bitişi planı aşan proje takvimde en az Sarı,
- harcaması bütçeyi aşmış proje bütçede Kırmızı.

**Genel renk**, takvim ve bütçe renginin kötü olanıdır. Risk sütunu sizin değerlendirmenizdir; ayrıca gösterilir,
genel rengi değiştirmez.

**Örnek:** Bütçe 4.000.000, harcanan 3.100.000, ilerleme %55 → KD 2.200.000, MPE 2.200.000 ÷ 3.100.000 = 0,71
(Kırmızı). Tahmini toplam maliyet 4.000.000 ÷ 0,7097 = **5.636.363,64**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                       # örnek veri, durum tarihi 12.10.2026
python main.py --projeler projeler.xlsx
python main.py --projeler projeler.xlsx --tarih 12.10.2026 --yesil 0,90 --sari 0,80
```

Zorunlu sütunlar: Proje, Başlangıç, Planlanan Bitiş, İlerleme %. Diğerleri: Sahip, Tahmini Bitiş, Bütçe,
Harcanan, Risk (Yüksek / Orta / Düşük), Durum (Devam / Tamamlandı / Beklemede / İptal; boşsa Devam), Öncelik.

## Çıktı

`cikti/bt_proje_portfoy.xlsx`:

- `Portföy`: her proje için genel renk, planlanan ve gerçekleşen ilerleme, TPE, MPE, takvim ve bütçe rengi, tahmini
  toplam maliyet ve aşım, risk. Kırmızı projeler en üsttedir.
- `Özet`: durum ve renk sayıları, portföy bütçesi, harcanan, kazanılan değer, tahmini toplam maliyet.
- `Uyarılar`: kırmızı projeler, bitişi geçmiş veya bütçesi aşılmış projeler, tutarsız kayıtlar.

## Dikkat

- **İlerleme yüzdesi sizin beyanınızdır.** Bütün hesap bu sayıya dayanır; "yüzde doksanda takılan" projeler tabloyu
  olduğundan iyi gösterir. Mümkünse ilerlemeyi teslim edilen iş paketlerine göre ölçün.
- **Planlanan ilerleme doğrusal varsayılır.** İşin takvime eşit dağıldığı kabul edilir; harcaması sonda yoğunlaşan
  (örneğin donanım alımı son ayda olan) projelerde endeksler yanıltıcı olabilir.
- **Tahmini toplam maliyet bir uzatmadır.** Bugünkü verimin aynen süreceğini varsayar; taahhüt edilmiş ama henüz
  faturalanmamış tutarlar ile sözleşme değişiklikleri hesaba girmez.
- **Eşikler tercihtir.** 0,95 ve 0,85 yaygın başlangıç değerleridir; kurumunuzun proje yönetim ilkelerine göre
  değiştirin.
- **Kapsam dışı:** kaynak (kişi) yükü, projeler arası bağımlılık, fayda takibi, kapsam değişiklik yönetimi.
- **Örnek veri:** Projeler, kişiler ve tutarlar kurgusaldır.

## Testler

- Planlanan ilerleme, PD, KD, TPE, MPE ve tahmini toplam maliyet (elle hesaplanmış); renk eşik sınırları.
- Bitişi geçmiş proje, planı aşan tahmin, bütçe aşımı, bütçesiz proje, devam etmeyen projeler.
- Portföy toplamları, sıralama, okuma hataları ve komut satırı.

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
