# Malzeme Talep ve Sevk Takibi · Kod Bloğu

> Şantiye › Şantiye Şefi · Gerisi Bende

Şantiye malzeme taleplerini **sipariş ve teslim durumuyla** birlikte izler; ihtiyaç tarihi geçen, termini ihtiyaçtan
sonraya verilen ve henüz siparişi verilmemiş kalemleri **imalatı geciktirecek riskler** olarak öne çıkarır;
tedarikçilerin zamanında teslim oranını hesaplar. İnternete bağlanmaz.

## Nasıl hesaplar?

| Durum | Kural |
|---|---|
| Teslim edildi | Teslim edilen ≥ talep miktarı |
| Kısmi teslim | 0 < teslim edilen < talep |
| Siparişte | Sipariş no ya da sipariş tarihi var, teslim yok |
| Sipariş bekliyor | Sipariş yok |

| Risk | Kural |
|---|---|
| İhtiyaç tarihi geçti (Yüksek) | Malzeme tam gelmedi ve ihtiyaç tarihi rapor tarihinden önce |
| Termin ihtiyaçtan sonra (Orta) | Tam gelmemiş kalemin tedarikçi termini ihtiyaç tarihinden sonra |
| Sipariş verilmedi (Orta) | Sipariş yok ve ihtiyaca `--tedarik` (varsayılan 7) gün ya da daha az kalmış |

Ayrıca ihtiyaç tarihinden sonra teslim edilen ve talepten fazla teslim edilen kalemler için bilgi notu verilir.

**Tedarikçi:** zamanında teslim % = son teslim tarihi termin tarihinde ya da önce olan kalem ÷ termini yazılmış teslim
edilmiş kalem × 100; ortalama talep → teslim süresi (gün).

**Örnek:** Gazbeton blok 14.10.2026'da gerekiyordu; tedarikçinin termini 16.10.2026 ve malzeme gelmedi → "İhtiyaç
tarihi geçti" ve "Termin ihtiyaçtan sonra".

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                               # örnek veri (15.10.2026)
python main.py --talepler malzeme_talepleri.xlsx --tarih 15.10.2026
python main.py --talepler malzeme_talepleri.xlsx --tedarik 14
```

Zorunlu sütunlar: Malzeme, Miktar, İhtiyaç Tarihi. Kısmi teslimlerde toplam teslim edilen miktarı ve son teslim
tarihini yazın.

## Çıktı

`cikti/malzeme_talep_takibi.xlsx`:

- `Talepler`: riskli kalemler üstte; durum, kalan miktar, ihtiyaca kalan gün, sipariş ve teslim bilgisi.
- `Tedarikçiler`: teslim edilen kalem, zamanında teslim %, ortalama süre.
- `Uyarılar`.

## Dikkat

- **Tedarik süresi malzemeye göre değişir.** Hazır beton, demir ve özel imalat ürünlerinin teslim süreleri farklıdır;
  `--tedarik` tek bir genel değerdir. Uzun teslimli kalemleri ayrı dosyada daha büyük değerle çalıştırın.
- **Kabul ve kalite kontrolü ayrıdır.** Teslim edilen miktar irsaliyeyle uyumlu, kabul edilmiş miktar olmalıdır; reddedilen
  malzemeyi teslim sayıya katmayın.
- **İş programıyla bağlantı:** İhtiyaç tarihini iş programındaki imalat başlangıcından geriye doğru belirleyin; kod iş
  programını okumaz.
- **Örnek veri:** Talep numaraları, tedarikçiler ve miktarlar kurgusaldır.

## Testler

- Durum ayrımı, kalan miktar, ihtiyaç tarihi / termin / sipariş riskleri, geç ve fazla teslim.
- Tedarikçi zamanında teslim oranı ve ortalama süre, sıralama, okuma ve komut satırı.

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
