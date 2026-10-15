# Günlük Likidite Raporu · Kod Bloğu

> Hazine ve Aktif-Pasif Yönetimi › Hazine Uzmanı · Workers / Workless

Vadesi gelen varlık ve yükümlülüklerin nakit akışlarından, **para birimi bazında günlük ve haftalık** likidite açığı /
fazlası tablosunu hazırlar. Basit bir stres senaryosunu yan yana gösterir. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Hesap |
|---|---|
| Kullanılabilir hazır değer | Tutar × (1 − iskonto %) |
| Günlük net akış | Girişler − çıkışlar |
| Gün sonu likidite | Kullanılabilir hazır değerler + kümülatif net akış |
| Likidite açığı | Gün sonu likiditenin eksiye düştüğü gün |
| Dayanma süresi | Likiditenin ilk kez eksiye düştüğü güne kadar geçen takvim günü |
| Büyük akış | Dönemin toplam girişinin (veya çıkışının) %20'si ve üzerindeki tek akış |

**Stres senaryosu:** "Beklenen" girişler `--giris-kesinti` (varsayılan %50) kadar azaltılır, "Beklenen" çıkışlar
`--cikis-artis` (varsayılan %20) kadar artırılır. "Kesin" akışlar değişmez.

**Tarih kuralları:**

- Hafta sonu ve tatil dosyasındaki günlere düşen akışlar bir sonraki iş gününe kaydırılır.
- Rapor tarihinden önceki akışlar ilk güne alınır ve uyarı verilir.
- Dönemden (`--gun`, varsayılan 30 takvim günü) sonraki akışlar rapora alınmaz.

Para birimleri ayrı raporlanır; çevrim yapılmaz.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek veri
python main.py --akislar akislar.xlsx --hazir hazir_degerler.xlsx --tarih 05.10.2026 --gun 30
python main.py --akislar akislar.xlsx --hazir hazir_degerler.xlsx --giris-kesinti 70 --cikis-artis 30 --tatiller tatiller.csv
```

| Dosya | Sütunlar |
|---|---|
| Akışlar | Tarih, Kalem, Yön (Giriş / Çıkış), Tutar, Para Birimi, Kesinlik (Kesin / Beklenen), Karşı Taraf |
| Hazır değerler | Kalem, Tutar, Para Birimi, İskonto (%) |
| Tatiller | Her satırda bir tarih |

Kesinlik boşsa akış "Kesin" sayılır.

## Çıktı

`cikti/likidite_raporu.xlsx`:

- `Özet`: para birimi bazında açılış, toplam akışlar, en düşük likidite, ilk açık günü, dayanma süresi (baz ve stres).
- `Günlük <PB>`: gün gün giriş, çıkış, net, kümülatif ve gün sonu likidite; stres sütunu.
- `Haftalık <PB>`: haftalık toplamlar ve hafta içi en düşük likidite.
- `Büyük Akışlar`, `Hazır Değerler`, `Uyarılar`.

## Dikkat

- **Yasal oran hesabı değildir.** Likidite karşılama oranı ve diğer yasal likidite ölçümleri BDDK düzenlemelerindeki
  tanım, dikkate alınma oranı ve vade dilimlerine göre yapılır. Bu paket günlük nakit akışı izleme aracıdır.
- **Gün içi likidite görünmez.** Tablo gün sonu bakiyeleriyle çalışır; ödeme saatlerinden kaynaklanan gün içi açıklar
  kapsam dışıdır.
- **Sonuç girdiye bağlıdır.** Mevduatın yenilenme oranı, kredi kullandırım tahmini ve iskonto oranları sizin
  varsayımınızdır. "Kesin / Beklenen" ayrımını tutarlı yapın.
- **Stres senaryosu basittir**; tek parametreli bir duyarlılıktır, bankanızın likidite stres testi çerçevesinin
  yerini tutmaz.
- **Yarım gün tatiller** tam iş günü sayılır. Yabancı para akışlarında ilgili ülkenin tatillerini ayrıca kontrol edin.
- **Örnek veri:** Akışlar ve tutarlar kurgusaldır.

## Testler

- Gün gün projeksiyon, stresin yalnız beklenen akışlara uygulanması, iskonto.
- Açık günü, dayanma süresi, haftalık toplamlar, para birimi ayrımı, büyük akışlar.
- Okuma: hafta sonu ve tatil kaydırma, geçmiş tarih, hatalı satırlar; örnek veri uçtan uca.

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
