# Taşıma Talebi Bilgi Çıkarma · AI Agent

> Lojistik › Havayolu Operasyon Uzmanı · Gerisi Bende

Müşterilerden ya da iç birimlerden e-postayla gelen **taşıma taleplerini** teklif ve planlama tablolarına aktarır:
- hava / deniz **navlun teklif talepleri** → kap listesi (adet, ölçü, ağırlık),
- **yurt içi sevkiyat talepleri** → sipariş satırları (palet, ağırlık, hacim, teslim tarihi).

Toplam ağırlık, hacim ve hacimsel ağırlığı kod hesaplar. Çıktılar **Navlun Teklif Hesaplama** ve **Sevkiyat Araç
Planlama** kod bloklarına doğrudan girer.

## Nasıl çalışır?

1. **Hazırlık (kod):** Talepler okunur; telefon, e-posta, IBAN ve kimlik numaraları maskelenir.
2. **Çıkarım (model):**
   - talep türü, referans, çıkış / varış, Incoterm, mod, hazır tarihi, mal tanımı;
   - tehlikeli madde ve istiflenemezlik bilgisi;
   - kaplar: adet, ölçü, ağırlık ve ağırlığın kap başına mı toplam mı olduğu;
   - yurt içi sevkiyat satırları.
3. **Denetim (kod):** Bir sayı ya da tarih belgede yazmıyorsa çıkarım modele geri gönderilir (varsayılan en fazla 2 tur).
   Örnekler: belgedeki tutarsız toplamı "düzeltmek", koli satırına hesaplanmış toplam yazmak.
4. **Hesap ve bulgular (kod):**

| Kalem | Kural |
|---|---|
| Kap başına ağırlık | "toplam 1.800 kg, 6 sandık" → 1.800 ÷ 6 = 300 kg |
| Brüt ağırlık / hacim | Σ adet × kg; Σ adet × boy × en × yükseklik ÷ 1.000.000 m³ |
| Hacimsel ağırlık (hava) | Σ adet × boy × en × yükseklik ÷ `--bolen` (varsayılan 6000) |
| Hata | Belgedeki toplam brüt ağırlık kaplardan hesaplananı tutmuyor |
| Dikkat | Tehlikeli madde (havada IATA DGR, denizde IMDG, karayolunda ADR beyanı gerekir); istiflenemez yük; eksik ölçü / ağırlık; ağırlık türü belirsiz; yurt içi satırda eksik palet / ağırlık / hacim / tarih |

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                    # örnek: hava, deniz ve yurt içi talep
python agent.py --girdi ./talepler --bolen 6000
```

Ardından:

```bash
python main.py --kaplar cikti/km-2026-88_kaplar.xlsx --tarife tarife_hava.csv --mod hava   # Navlun Teklif Hesaplama
python main.py --siparisler cikti/sevkiyat_siparisleri.xlsx                                 # Sevkiyat Araç Planlama
```

## Çıktı

- `cikti/tasima_talepleri.xlsx`:
  - `Talepler`: talep başına tür, mod, müşteri / referans, çıkış-varış, teslim şekli, hazır tarihi, mal, kap adedi, brüt,
    hacim, hacimsel ağırlık, bulgular ve son kontrol.
  - `Sevkiyat Siparişleri`
  - `Bilgi`
- `cikti/<referans>_kaplar.xlsx`: her navlun talebi için kap listesi (Navlun Teklif Hesaplama girdisi).
- `cikti/sevkiyat_siparisleri.xlsx`: yurt içi siparişler (Sevkiyat Araç Planlama girdisi).

## Dikkat

- **Teklif vermeden önce** ölçü ve ağırlıkları müşteriye teyit ettirin. Kodun "belgedeki toplam tutmuyor" bulgusu,
  müşterinin yazdığı rakamlardan birinin yanlış olduğunu gösterir; hangisinin yanlış olduğunu söylemez.
- **Tehlikeli madde.** Agent yalnız belgede yazanı aktarır; UN numarası veya sınıf tahmin etmez. Kabul ve beyan için
  lisanslı DGR / IMDG personeline başvurun.
- **Hacimsel ağırlık böleni** taşıyıcıya ve hizmete göre değişebilir (ör. ekspres kargoda 5000); `--bolen` ile verin.
- **Örnek veri:** Firmalar, siparişler ve yükler kurgusaldır. Örnek hava talebinde müşterinin yazdığı toplam (254 kg)
  bilerek kaplarla tutmaz (264 kg).

## Testler

Testler gerçek API çağırmaz. Sahte model hava talebinin ilk çıkarımında bilerek:
- belgedeki toplamı 264'e "düzeltir",
- koli satırına hesaplanmış toplam ağırlık (54 kg) yazar.

Kodun ikisini de yakaladığı doğrulanır. Ayrıca şunlar doğrulanır:
- kap başına ağırlık, brüt, hacim ve hacimsel ağırlık hesapları,
- tehlikeli madde, istif ve eksik bilgi bulguları,
- kap ve sevkiyat dosyalarının kod bloklarının sütunlarıyla aynı olması.

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
- **Yapay zekâ hata yapabilir.** Model yanlış, eksik veya uydurma bilgi üretebilir. Kod kontrolleri bazı hataları
  yakalar, hepsini yakalayamaz. Her taslağı okuyup onaylamadan kullanmayın, göndermeyin, imzalamayın.
- **Maskeleme sınırlıdır.** Yalnız telefon, e-posta, TCKN, IBAN ve LinkedIn adresi kalıpları maskelenir. Ad soyad,
  adres, doğum tarihi, sağlık bilgisi ve serbest metindeki diğer kişisel veriler modele gidebilir. Hangi verinin bulut
  sağlayıcısına gönderildiğinden siz sorumlusunuz; gizli veride yerel model (Ollama) kullanın.
- **API ücretleri, kotalar ve API anahtarının güvenliği** size aittir.

Ayrıntılar: `LICENSE` (MIT) ve `SORUMLULUK_REDDI.md`.
