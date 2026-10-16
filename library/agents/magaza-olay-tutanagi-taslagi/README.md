# Mağaza Olay Tutanağı Taslağı · AI Agent

> Perakende Mağaza › Kasa Sorumlusu · Workers / Workless

Kasa farkı, hırsızlık şüphesi, hasarlı / kayıp ürün ve iade usulsüzlüğü şüphesi gibi mağaza olaylarının kısa notlarından
**nesnel, imzaya hazır olay tutanağı taslağı** hazırlar. Tutanak kimseyi suçlamaz; suçlayıcı ifadeleri kod engeller.
Kasa farkını ve ürün tutarını model değil kod hesaplar.

## Nasıl çalışır?

1. **Hazırlık (kod):** Not dosyası olaylara ayrılır (`Olay: GG.AA.YYYY SS:DD | Yer`) ve satırlar numaralanır.
   `--gizle` ile verilen kişi adları `[KİŞİ-n]` olarak gönderilir ve tutanakta geri açılır. Telefon, e-posta ve kimlik
   numarası maskelenir.
2. **Taslak (model):** Her olay için şunlar yazılır:
   - olay türü, nesnel özet ve zaman çizelgesi,
   - etkilenen ürünler ve kasa sayım bilgisi,
   - kanıtlar, yapılan bildirimler, yapılacaklar ve eksik bilgiler.
3. **Denetim (kod):** Aşağıdakilerden biri varsa yalnız o olay modele geri gönderilir (varsayılan en fazla 2 tur):
   - Notta olmayan bir sayı ya da saat var. Farkı veya toplamı hesaplayıp yazmak da kabul edilmez.
   - Suçlayıcı ya da yargılayıcı ifade var: hırsız, çaldı / çalmış, suçlu, dolandırıcı, kasıtlı, kasten, bilerek, yalan,
     zimmet, aşırdı, yürüttü, hileli… ("alarm çaldı", "telefon çaldı" ve "hırsızlık şüphesi" ifadeleri sayılmaz.)
   - Dayanak not satırları geçersiz
   - Kasa farkı olayında sayım tutarları da eksik bilgi de yok
4. **Hesap ve tutanak (kod):**
   - Kasa farkı = sayılan − beklenen. Notta yazan farkla tutmazsa "sayımı tekrarlayın" hatası verilir.
   - Ürün tutarı = adet × birim fiyat; notta yazan tutarla karşılaştırılır.
   - Tutanak sabit bir şablonla, bölüm numaraları ve imza alanlarıyla (düzenleyen, tanık, mağaza yöneticisi) oluşturulur.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                     # örnek 3 olay (kasa farkı, alarm, depo hasarı)
python agent.py --girdi olaylar.txt --gizle "Ayşe Deneme" "Can Kurgu" --magaza "Örnek Market Kadıköy"
```

Not örneği:

```
Olay: 15.10.2026 22:10 | Örnek Market Kadıköy - Kasa 2
Gün sonu sayımında sistemde görünen nakit 18.450,00 TL, sayılan 18.200,00 TL. Fark 350 TL eksik.
```

## Çıktı

- `cikti/olay_tutanaklari.xlsx`:
  - `Olaylar`: tutanak no (OT-261015-01), tarih, yer, olay türü, konu, etkilenen tutar, kasa farkı, hesap bulguları,
    tutanak metni, son kontrol.
  - `Kontrol`
- `cikti/olay_tutanaklari.txt`: imzaya hazır tutanaklar.

## Dikkat

- **Tutanak bir taslaktır.** İmzalamadan önce olaya tanık olanlarla birlikte okuyun ve düzeltin. Tutanakta yer alan
  bilgilerin doğruluğundan imzalayanlar sorumludur.
- **Suç isnadı yapmayın.** Hırsızlık veya usulsüzlük şüphesinde kişileri suçlayan ifadeler hem hukuki risk hem de
  kişilik haklarının ihlali olabilir. Kamera kaydını ve kanıtları koruyun; adli süreç için güvenlik / hukuk biriminize
  başvurun.
- **Personel hakkında karar aracı değildir.** Kasa farkları için disiplin veya kesinti kararı bu taslağa dayanarak
  verilmemelidir.
- **Kişisel veri.** Notlarda geçen ve `--gizle` ile verilmeyen adlar modele gider. Personel ve müşteri adlarını
  `--gizle` ile verin ya da notlara görevleriyle yazın.
- **Örnek veri:** Mağaza, kişiler ve olaylar kurgusaldır. Örnek notta kasa farkı bilerek yanlış yazılmıştır (350 TL;
  doğrusu 250 TL).

## Testler

Testler gerçek API çağırmaz. Sahte model ilk turda bilerek:
- hırsızlık şüphesi özetine "ürünleri çaldı" yazar,
- depo olayında olmayan bir not satırına dayanır.

Kodun ikisini de yakaladığı doğrulanır. Ayrıca şunlar doğrulanır:
- "alarm çaldı" ifadesinin suçlama sayılmaması,
- kasa farkının kodla hesaplanıp nottaki yanlış farkın bulunması,
- adların modele gitmeyip tutanakta geri açılması,
- bölüm numaralarının sıralı olması.

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
