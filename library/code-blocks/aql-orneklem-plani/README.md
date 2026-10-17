# AQL Örneklem Planı · Kod Bloğu

> Üretim › Kalite Kontrol Elemanı · Tekstil › Kalite Kontrol · Gerisi Bende

Parti büyüklüğü, muayene seviyesi ve AQL değerinden **tek örneklemeli normal muayene planını** bulur:
örneklem sayısı, kabul (Ac) ve ret (Re) sayısı. Muayenede bulunan hata sayıları verilirse partiyi
**KABUL / RET** olarak değerlendirir. Kritik, majör ve minör hatalar için ayrı AQL desteklenir. İnternete bağlanmaz.

## Tablolar ve kapsam

- Örneklem kod harfleri ve kabul/ret sayıları **MIL-STD-105E Tablo I ve II-A**'dan (kamu malı) alınmıştır;
  **ISO 2859-1 Tablo 1 ve 2-A** ile eşdeğerdir. Tablo, yayımlanmış ISO 2859-1 tablolarıyla satır satır karşılaştırılarak test edilmiştir.
- Normal muayene, tek örnekleme, **AQL 0,010 – 10** (yüzde kusurlu). Sıkılaştırılmış/azaltılmış muayene ve
  100 birimdeki kusur sayısı için 10 üstü AQL'ler kapsam dışıdır.
- Tablodaki **oklar** otomatik uygulanır: okun gösterdiği ilk plan kullanılır ve raporda belirtilir.
- Örneklem parti büyüklüğüne eşit veya büyükse **%100 muayene** önerilir.
- AQL olarak **0** girilirse (yaygın kritik hata uygulaması) tek hata partiyi reddeder.

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py --parti 1200 --aql 2.5                       # J → 80 adet, Ac 5, Re 6
python main.py --parti 3000 --kritik 0 --majör 2.5 --minör 4.0
python main.py --girdi muayeneler.xlsx                       # toplu değerlendirme
python main.py                                              # örnek listeyle dener (kritik 0, majör 2,5, minör 4,0)
```

Toplu girdi sütunları: **Parti Büyüklüğü** (zorunlu); Sipariş No, Seviye, Kritik AQL, Majör AQL, Minör AQL,
Kritik Hata, Majör Hata, Minör Hata (isteğe bağlı).

## Testler

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
