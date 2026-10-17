# Mükellef Evrak Eksik Takibi · Kod Bloğu

> Muhasebe ve Defter İşleri › Mali Müşavir Stajyeri · Gerisi Bende

Her mükelleften dönem için beklenen belgeleri (fatura, banka ekstresi, puantaj, kasa) gelenlerle karşılaştırır;
**eksik listesini, tamamlanma oranını ve hazır hatırlatma metnini** çıkarır. İnternete bağlanmaz.

## Nasıl çalışır?

| Kavram | Kural |
|---|---|
| Beklenen | Beklenen belgeler dosyasındaki her mükellef – belge türü satırı |
| Geldi | Gelen kayıtlarda aynı mükellef, aynı belge türü ve seçilen dönem varsa |
| Son teslim tarihi | Dönemi izleyen ayın `--son-gun` günü (varsayılan 10) |
| Eksik | Son teslim geçmemişse "Bekleniyor", geçmişse gecikme günüyle "Gecikmiş" |
| Geç geldi | Geliş tarihi son teslim tarihinden sonra |
| Tamamlanma % | Gelen belge türü sayısı ÷ beklenen belge türü sayısı × 100 |

Mükellef durumu: **Tamam**, **Eksik var** veya **Hiç gelmedi**. Mükellef ve belge adları büyük-küçük harf ve Türkçe
karakter farkı gözetmeden eşleştirilir.

**Örnek:** Eylül 2026 dönemi, son teslim 10.10.2026, bugün 13.10.2026. Dört belge beklenen mükelleften satış
faturaları 05.10'da, alış faturaları 12.10'da (geç) gelmiş; banka ekstresi ve puantaj yok → "Eksik var", %50,
2 belge 3 gündür gecikmede.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek veri, Eylül 2026, bugün 13.10.2026
python main.py --beklenen beklenen_belgeler.xlsx --gelen gelen_belgeler.xlsx --donem 2026-09
python main.py --beklenen beklenen_belgeler.xlsx --gelen gelen_belgeler.xlsx --donem 09.2026 --son-gun 7 --tarih 13.10.2026
```

- Beklenen belgeler: Mükellef, Belge Türü (zorunlu); Sorumlu, İletişim. Liste bir kez hazırlanır, her ay kullanılır.
- Gelen belgeler: Mükellef, Belge Türü, Dönem (`2026-09`, `09.2026`, `Eylül 2026`), Geliş Tarihi, Not.

## Çıktı

`cikti/evrak_eksik_takibi.xlsx`:

- `Mükellef Özeti`: durum, tamamlanma, eksik belgeler ve kopyalanabilir hatırlatma metni; en eksik mükellef üstte.
- `Belge Matrisi`: mükellef × belge türü tablosu (Geldi, Geç geldi, Bekleniyor, Gecikmiş).
- `Eksik Belgeler`: eksik her belge, gecikme günü ve sorumlusuyla.
- `Sorumlu Özeti`: personel bazında eksik belge sayısı ve mükellefler.
- `Uyarılar`: gecikmiş belgeler, listede olmayan gelen belgeler, okunamayan satırlar.

## Dikkat

- **Son teslim günü büronuzun iç kuralıdır.** Yasal bir süre değildir; beyanname ve bildirim tarihlerine göre kendi
  takviminizi `--son-gun` ile belirleyin. Yasal süreleri ayrıca takip edin.
- **Kod belgenin içeriğine bakmaz.** "Geldi" yalnız kaydın var olduğunu gösterir; faturaların eksiksiz, ekstrenin
  tam ay olduğunu siz kontrol edersiniz.
- **E-belgeler ayrı izlenebilir.** E-fatura ve e-arşiv gibi portaldan çekilen belgeleri beklenen listesine yalnız
  mükelleften istemeniz gerekenler için ekleyin.
- **Adlar tutarlı yazılmalı.** "Banka Ekstresi" ile "Banka Ekstreleri" farklı tür sayılır; listede olmayan gelen
  belgeler Uyarılar sayfasında görünür.
- **Hatırlatma metni bir taslaktır.** Göndermeden önce okuyun; kod e-posta veya mesaj göndermez.
- **Örnek veri:** Mükellefler, personel ve iletişim bilgileri kurgusaldır.

## Testler

- Dönem yazımlarının çözümü; geldi, geç geldi, bekleniyor, gecikmiş durumları; mükellef durumu ve tamamlanma oranı.
- Son teslim günü ve rapor tarihinin etkisi, yıl sonu dönemi, aynı belgenin iki kez gelmesi, başka dönemin belgesi.
- Sorumlu özeti, hatırlatma metni, listede olmayan belge; okuma ve komut satırı hataları; örnek veri uçtan uca.

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
