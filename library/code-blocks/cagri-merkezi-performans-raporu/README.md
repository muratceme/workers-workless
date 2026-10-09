# Çağrı Merkezi Performans Raporu · Kod Bloğu

> Müşteri Hizmetleri ve Çağrı Merkezi › Çağrı Merkezi Yöneticisi · Workers / Workless

Çağrı kayıtlarından **karşılama oranı, terk oranı, servis seviyesi, ASA, AHT ve ilk temasta çözüm** oranını
hesaplar. Sonuçları gün, kuyruk, 30 dakikalık aralık ve temsilci bazında gösterir. Her aralık için hedef servis
seviyesine ulaşmak için gereken temsilci sayısını **Erlang C** ile hesaplar. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Tanım |
|---|---|
| Gelen | Cevaplanan + terk. `--kisa-terk` (5 sn) dolmadan kapatılan çağrılar sayılmaz (yanlış arama, hemen kapatma). |
| Karşılama oranı | cevaplanan / gelen |
| Terk oranı | terk / gelen; hedef `--terk-hedef` (%5) |
| Servis seviyesi (SL) | T içinde cevaplanan / (cevaplanan + T'den sonra terk edilen). T = `--sl-sure` (20 sn), hedef `--sl-hedef` (%80). T içinde terk edilen çağrı paydaya girmez. |
| ASA | Cevaplanan çağrılarda ortalama bekleme |
| AHT | (görüşme + beklemeye alma + çağrı sonrası iş) / cevaplanan |
| FCR (tekrar arama) | Cevaplanan çağrıdan sonra aynı numara `--tekrar-gun` (7) gün içinde yeniden aramadıysa çözülmüş sayılır. Veri bitişine 7 günden yakın çağrılar paydaya girmez, çünkü tekrar arayıp aramadıkları henüz bilinmez. |
| FCR (temsilci beyanı) | "Çözüldü" sütunu varsa: Evet / değerlendirilen |

SL'nin farklı tanımları vardır; örneğin bazı sistemler T içinde cevaplananı tüm gelene böler. Karşılaştırma
yaparken santralinizin tanımını kontrol edin.

### Personel ihtiyacı (Erlang C)

- Her 30 dakikalık aralık için hafta içi günlerin ortalama gelen çağrısı ve AHT'si kullanılır.
- Trafik (Erlang) = çağrı × AHT / 1.800 sn. Hedef SL'ye ulaşan en küçük temsilci sayısı bulunur.
- Planlanacak temsilci = gereken / (1 − `--kayip`). Kayıp payı (%30) mola, eğitim, toplantı ve devamsızlığı
  karşılar.
- "Cevaplayan temsilci", o aralıkta en az bir çağrı cevaplayan temsilci sayısıdır. Vardiyadaki gerçek sayı için
  santralin oturum (login) raporunu kullanın.
- Erlang C, terk edilmeyen ve sürekli gelen bir çağrı akışı varsayar. Bu nedenle gerçek ihtiyacı bir miktar
  fazla gösterebilir.

### Uyarılar

| Uyarı | Önem |
|---|---|
| Dönem SL hedefin altında; dönem terk oranı hedefin üstünde | Yüksek |
| Günlük SL hedefin altında; kuyruk terk oranı yüksek | Orta |
| Aralıkta cevaplayan temsilci Erlang ihtiyacından az ve SL hedef altı | Orta |
| Temsilcide `--kisa-gorusme` (10 sn) altı görüşme oranı > %5 (erken kapatma / hat sorunu) | Orta |
| Temsilci AHT'si ekip medyanının 1,5 katından fazla | Bilgi |

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                              # örnek: 15 iş günü, 3 kuyruk, 12 temsilci
python main.py --cagrilar cagrilar.xlsx --sl-sure 20 --sl-hedef 80 --kayip 30
python main.py --cagrilar cagrilar.xlsx --sl-sure 30 --sl-hedef 90 --kisa-terk 10 --tekrar-gun 3
```

| Sütun | Açıklama |
|---|---|
| Başlangıç | Tarih ve saat (veya ayrı Tarih / Saat sütunları) |
| Kuyruk, Temsilci | Kuyruk / hat / skill; terk edilen çağrıda temsilci boş |
| Bekleme, Görüşme, Beklemeye Alma, Çağrı Sonrası İş | Saniye veya ss:dd:sn |
| Durum | Cevaplandı / Terk |
| Arayan | FCR ve tekrar arayanlar için; raporda maskelenir |
| Kategori, Çözüldü | İsteğe bağlı |

## Çıktı

`cagri_merkezi_raporu.xlsx`:
- `Özet`: genel göstergeler ve tanımları; kuyruk tablosu.
- `Günlük`: SL hedefe göre renkli.
- `Aralıklar`: 30 dakikalık ortalama çağrı, SL, ASA, AHT, Erlang C ihtiyacı, kayıp payıyla planlanacak sayı,
  cevaplayan temsilci, doluluk; iki grafik.
- `Temsilciler`: cevaplanan, AHT ve medyana oranı, beklemeye alma ve çağrı sonrası iş payı, kısa görüşme, FCR.
- `Tekrar Arayanlar`: 7 gün içinde 3 veya daha fazla arayan numaralar (maskeli).
- `Uyarılar`.

## Dikkat

- **Temsilci göstergeleri** tek başına performans değerlendirmesi için kullanılmamalıdır. Çağrı türü, kuyruk ve
  vardiya farklarını göz önüne alın; kalite değerlendirmesiyle birlikte yorumlayın.
- **Kişisel veri:** Arayan numaraları kişisel veridir. Rapor numaraları maskeler, ancak girdi dosyasını güvenli
  saklayın. Çağrı kayıtlarının işlenmesi için KVKK aydınlatma yükümlülüğünü kontrol edin.
- **Örnek veri:** Örnek çağrılar, temsilciler ve numaralar kurgusaldır.

## Testler

Şunlar test edilir:
- Erlang C: M/M/1 ve M/M/2 için bilinen değerler, en küçük temsilci sayısı.
- Gelen, cevaplanan, terk, kısa terk ve SL'nin CSV'den bağımsız hesapla doğrulanması.
- FCR tekrar arama mantığı ve veri sonu dışlaması; numara maskeleme.
- Temsilci uyarıları, Excel ve CLI.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
