# Reasürans Slipi Özeti ve Kontrolü · AI Agent

> Reasürans › Reasürans Uzmanı · Gerisi Bende

Reasürans slipinden ya da anlaşma metninden (Türkçe / İngilizce) **ana şartları özet tabloya** çıkarır. Ardından:
- beklenen (onaylanan) şartlarla karşılaştırır,
- reasürör paylarının toplamını kontrol eder,
- zorunlu klozların (siber, yaptırım, nükleer…) metinde geçip geçmediğine bakar.

## Nasıl çalışır?

1. **Hazırlık (kod):** Slip metni okunur. Telefon, e-posta ve IBAN maskelenir. Beklenen şartlar ve kloz listesi yalnız
   kodda kullanılır, sağlayıcıya gönderilmez.
2. **Çıkarım (model):** Ana şartlar slipteki yazımla çıkarılır:
   - tür, dilim, dönem, kapsam;
   - limit, öncelik, para birimi;
   - yeniden tesis (adet ve koşul), MDP, prim oranı, saat klozu;
   - istisnalar, hasar bildirimi, uyuşmazlık;
   - reasürörler, payları ve lider.
3. **Denetim (kod):** Aşağıdakilerden biri varsa çıkarım modele geri gönderilir (varsayılan en fazla 2 tur):
   - Tarih slipte o biçimde yazmıyor
   - Sayı slipte yazmıyor (ör. limit + önceliği toplayıp yazmak)
   - Reasürör adı slipte aynen geçmiyor
4. **Kontroller (kod):**

| Önem | Kontrol |
|---|---|
| Hata | Reasürör payları toplamı %100 değil: açık pay ya da fazla yerleşim (signing down) |
| Hata | Beklenen şartla fark: sayısal kalemlerde rakam, tarihlerde tarih, metinde içerik karşılaştırılır |
| Dikkat | Zorunlu kloz bulunamadı (anahtar kelime araması); lider reasürör yok ya da birden fazla |
| Dikkat | Hasar fazlası anlaşmasında limit / öncelik, Cat XL'de saat klozu yok; beklenen kalem slipte bulunamadı |

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                     # örnek Cat XL slipi (4 bulgu)
python agent.py --slip slip_2027.pdf --beklenen beklenen.csv --klozlar zorunlu_klozlar.csv
```

Beklenen şartlar dosyası örneği:

```
Kalem;Beklenen
Limit;400.000.000
Öncelik;50.000.000
MDP;18.000.000
```

Kalem adları: Tür, Başlangıç, Bitiş, Limit, Öncelik, Para Birimi, Reinstatement Adedi, MDP, Prim Oranı %, Saat Klozu,
Kapsam, Hasar bildirimi, Uyuşmazlık.

## Çıktı

`cikti/slip_ozeti.xlsx`:
- `Özet`: şart başına slipteki değer, istisnalar, model notları.
- `Reasürörler`: ad, pay, lider, toplam.
- `Karşılaştırma`: kalem, slipteki değer, beklenen, durum (Uygun / Hata / Dikkat / Bilgi), not.
- `Bulgular`: kodun bütün bulguları.
- `Kontrol`

## Dikkat

- **Bağlayıcı olan slipin kendisidir.** Özet çalışma kolaylığı içindir; imzadan önce metni okuyun.
- **Kloz kontrolü anahtar kelime aramasıdır.** Klozun varlığını gösterir, içeriğini, kapsamını ya da güncel model kloz
  olup olmadığını değerlendirmez.
- **Kotapar / eksedan anlaşmalarında** komisyon, kâr komisyonu, portföy transferi gibi şartları beklenen dosyasına ekleyin;
  özet tablosunda olmayan kalemler "Bilgi" olarak elle kontrole bırakılır.
- **Örnek veri:** Şirketler, reasürörler ve şartlar kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk çıkarımda bilerek:
- limiti önceliğiyle toplar (450.000.000),
- lider reasürörün adını tamamlar ("Kurgu Reinsurance AG"),
- bitiş tarihini farklı biçimde yazar.

Kodun üçünü de yakaladığı doğrulanır. Ayrıca şunlar doğrulanır:
- pay toplamı (eksik ve fazla yerleşim),
- MDP farkı, tarih karşılaştırması,
- zorunlu kloz ve lider reasürör kontrolleri,
- beklenen şartların sağlayıcıya gönderilmemesi.

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
