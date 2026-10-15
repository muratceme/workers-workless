# Konteyner Demuraj ve Detention Takibi · Kod Bloğu

> Deniz ve Hava Yolu Operasyon › Denizyolu Operasyon Uzmanı · Workers / Workless

Konteynerlerin **limanda bekleme (demuraj) ve iade (detention) sürelerini serbest gün kurallarıyla** izler, oluşan ve
işlemekte olan masrafı hesaplar. İnternete bağlanmaz.

## Nasıl hesaplar?

| Kalem | Süre |
|---|---|
| Demuraj günü | Limandan çıkış − tahliye + 1 (çıkış yoksa rapor tarihine kadar) |
| Detention günü | Boş iade − limandan çıkış + 1 (iade yoksa rapor tarihine kadar) |

- `--gun-sayimi haric` verilirse başlangıç günü sayılmaz (+ 1 yapılmaz).
- **Ücret:** tarife dosyasındaki gün dilimlerine göre, dilime düşen gün × günlük ücret. Ücreti 0 olan ilk dilimler
  serbest süredir.
- **Tarife seçimi:** konteynerin hattına ve tipine en özel satırlar kullanılır: hat + tip > yalnız hat > yalnız tip > genel.
- **Açık konteynerlerde** kalan serbest gün, son serbest gün ve bugün işleyen günlük ücret gösterilir.

**Örnek:** tarife 1–7. gün 0, 8–14. gün 40, 15. günden sonra 80. Konteyner 18 gündür limanda → 7 × 40 + 4 × 80 = 600;
bugünkü günlük ücret 80.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri (rapor tarihi 12.10.2026)
python main.py --konteynerler konteynerler.xlsx --tarife tarife.xlsx --gun-sayimi haric --tarih 15.10.2026
```

| Dosya | Sütunlar |
|---|---|
| Konteynerler | Konteyner No, Tahliye Tarihi (zorunlu); Tip, Hat, Dosya No, Limandan Çıkış, Boş İade |
| Tarife | Ücret Türü (Demuraj / Detention), Gün Başlangıç, Günlük Ücret (zorunlu); Hat, Tip (boş = hepsi), Gün Bitiş (boş = sonrası), Para Birimi |

## Çıktı

`cikti/demuraj_detention.xlsx` → `Konteyner Takibi` (ücreti işleyenlerden başlayarak her konteyner; gün ve ücretler,
açık kalem, kalan serbest gün, son serbest gün, günlük ücret; toplam satırı) ve `Uyarılar`.

## Dikkat

- **Tarife hattındır, siz girersiniz.** Serbest süre, dilimler ve ücretler hatta, limana, konteyner tipine, döneme ve
  sözleşmenize göre değişir. Kod hiçbir hattın tarifesini bilmez; örnek tarifedeki rakamlar kurgusaldır.
- **Gün sayımı hatta göre değişir.** Başlangıç gününün sayılıp sayılmadığı, hafta sonu ve resmi tatillerin serbest
  süreye dahil olup olmadığı tarifeye bağlıdır. Kod takvim günü sayar; iş günü esaslı serbest süre desteklenmez.
- **Birleşik serbest süre desteklenmez.** Demuraj ve detention'ın tek bir serbest sürede toplandığı tarifelerde
  sonuç yanlış olur.
- **İthalat akışı esas alınmıştır** (tahliye → limandan çıkış → boş iade). İhracatta boş alım ve dolu teslim
  tarihlerini aynı sütunlara uyarlayarak kullanabilirsiniz; liman ardiyesi ayrı bir masraftır ve hesaplanmaz.
- **Hattın faturası esastır.** Sonuç ön hesaptır; açık konteynerlerde tutar her gün artar.
- **Örnek veri:** Konteyner numaraları, hatlar ve ücretler kurgusaldır.

## Testler

- Dilim ücretleri ve sınırları (7., 8., 14., 15. gün), serbest süresi olmayan tarife.
- Kapalı ve açık konteynerler, kalan serbest gün, son serbest gün, başlangıç günü hariç sayım.
- Tarife seçimi önceliği, tarifesi olmayan konteyner, toplamlar ve sıralama.
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
