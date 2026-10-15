# Final Kontrol Raporu · Kod Bloğu

> Kalite Kontrol › Kalite Kontrol Şefi · Workers / Workless

Final denetim kayıtlarından **kritik, majör ve minör hata sayılarını AQL kabul sayılarıyla** karşılaştırır; denetim
bazında KABUL / RET sonucu, hata dağılımı ve atölye özeti içeren rapor üretir. İnternete bağlanmaz.

## Nasıl hesaplar?

Her denetim için hata kayıtları sınıfına göre toplanır ve denetim satırındaki kabul sayılarıyla karşılaştırılır.

| Sonuç | Koşul |
|---|---|
| RET | Herhangi bir sınıfta hata sayısı > kabul sayısı |
| KABUL | Tüm sınıflarda hata sayısı ≤ kabul sayısı (eşitlik kabuldür) |
| Değerlendirilemedi | Majör veya minör kabul sayısı girilmemiş ve girilen sınıflarda aşım yok |

- **Kritik kabul sayısı boşsa 0** alınır (tek kritik hata RET).
- **Hata oranı %** = hata sayısı ÷ numune adedi × 100.
- **Hata dağılımı:** hata türleri adet sırasına göre, pay ve birikimli payla (Pareto).
- **Atölye bazında:** denetim sayısı, ret oranı, majör hata oranı ve en sık hata.
- Hata satırında adet boşsa 1 sayılır.

**Örnek:** numune 50, kabul sayıları kritik 0 / majör 3 / minör 5. Bulunan: majör 5, minör 2 → majör 5 > 3 → RET.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                         # örnek veri
python main.py --denetimler denetimler.xlsx --hatalar hatalar.xlsx
```

| Dosya | Sütunlar |
|---|---|
| Denetimler | Denetim No, Numune Adedi (zorunlu); Sipariş No, Model, Atölye, Müşteri, Parti Adedi, Kabul Kritik, Kabul Majör, Kabul Minör, Tarih, Denetçi |
| Hatalar | Denetim No, Sınıf (zorunlu); Hata, Adet, Bölge |

## Çıktı

`cikti/final_kontrol_raporu.xlsx`:

- `Denetim Sonuçları`: sınıf bazında hata / kabul, majör hata oranı, sonuç ve gerekçe; boş "Aksiyon" sütunu.
- `Hata Dağılımı`: hata türü, sınıf, adet, pay ve birikimli pay.
- `Atölye Bazında`: denetim, kabul, ret, hata sayıları ve en sık hata.
- `Uyarılar`: denetimi olmayan hatalar, tanınmayan sınıflar, okunamayan değerler.

## Dikkat

- **Kabul sayılarını siz girersiniz.** Kod örneklem planı üretmez; numune adedi ve kabul sayıları müşterinizin
  kalite el kitabındaki muayene seviyesi ve AQL değerlerine göre belirlenmiş olmalıdır. Plan bulmak için bu
  kütüphanedeki "AQL Örneklem Planı" paketini kullanabilirsiniz.
- **Hata sınıflandırması denetçinindir.** Aynı hatanın kritik / majör / minör sayılması müşterinin hata
  sınıflandırma listesine göre değişir; kod sınıfı değiştirmez.
- **Hata mı, hatalı ürün mü?** Kabul sayıları çoğu planda hatalı ürün adedine göredir. Bir üründe birden çok hata
  varsa en ağır sınıftan tek hata olarak girmeniz gerekebilir; müşterinizin kuralını uygulayın.
- **Sonuç sevk kararı değildir.** RET sonrası yüzde yüz ayıklama, yeniden denetim veya müşteri onaylı sevk kararları
  kalite ve müşteri sürecine aittir.
- **Örnek veri:** Siparişler, atölyeler, müşteriler ve hatalar kurgusaldır.

## Testler

- Kabul / ret sınırları (kabul sayısına eşitlik), kritik hata, hatasız denetim.
- Kabul sayısı eksik denetim; Pareto ve atölye özeti; sınıf adlarının tanınması.
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
