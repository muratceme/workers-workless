# Kredi Teklif Dosyası Hazırlama · AI Agent

> Bankacılık › Ticari Bankacılık › Ticari Portföy Yöneticisi · Workers / Workless

Firma bilgisi, mali tablolar, KKB/memzuç riski, istihbarat ve teminat bilgilerinden **kredi tahsis birimine
gönderilecek teklif özetini** hazırlar.
- **Kod:** Sayıları ve tabloları üretir: rasyolar, risk dağılımı, teminat karşılaması.
- **Model:** Değerlendirme metnini taslak olarak yazar.
- **Portföy yöneticisi:** Teklifi kontrol edip imzalar.

> **Önemli — bankacılık sırrı:** Müşteri bilgileri bankacılık sırrıdır ve KKB verileri kişisel/ticari sır
> niteliğindedir. Bu bilgileri dış bir yapay zekâ servisine göndermek bankanızın politikasına aykırı olabilir.
> Kullanmadan önce bilgi güvenliği ve uyum birimlerinizin onayını alın. Veri bilgisayardan çıkmasın isteniyorsa
> `.env` içinde `WW_PROVIDER=ollama` ile yerel bir model kullanın.

## Nasıl çalışır?

1. **Dosyaları tanır (kod):** Klasördeki dosyalar türlerine göre ayrılır:
   - Talep ve firma bilgisi: adında "talep" geçen `.txt`, `.docx` veya `.pdf`.
   - Mali tablolar: Kalem × yıl tablosu.
   - KKB/memzuç risk tablosu.
   - Teminat listesi.
   - İstihbarat ve diğer notlar.
2. **Hesaplar ve kontrol eder (kod):** Bu sonuçlar kesindir ve modele "değiştirme" talimatıyla verilir.
   - **Rasyolar:** cari oran, asit-test, kaldıraç, özkaynak/aktif, FAVÖK ve marjı, net finansal borç/FAVÖK, faiz
     karşılama, alacak ve stok devir süreleri, büyüme. FAVÖK verilmezse brüt kâr − faaliyet giderleri + amortisman
     olarak hesaplanır.
   - **Uyarılar:** Bilanço denkliği, negatif özkaynak, cari oran < 1,2, kaldıraç > %70, net finansal borç/FAVÖK > 3,
     faiz karşılama < 1,5, zarar, ciro düşüşü, finansman giderinin cirodan hızlı artışı, tahsil ve stok sürelerinin
     uzaması.
   - **KKB:** Sektördeki limit doluluğu, limit aşımı ve bilançodaki banka kredileriyle tutarlılık.
   - **Teminat karşılama oranı:** Teminat değerleri türüne göre katsayıyla çarpılır ve talebe bölünür.
   - **İstihbarat:** "karşılıksız", "protesto", "haciz", "icra", "gecikme", "yapılandırma" gibi ifadeler taranır.
     "bulunmamaktadır" gibi olumsuzlanmış cümleler sayılmaz.
3. **Maskeler:**
   - **Takma adlar:** Firma unvanı `[FİRMA]`, VKN `[VKN]` olur; `--gizle` ile verilen ortak ve kefil adları
     soyadlarıyla birlikte `[GİZLİ-1]` gibi takma adlar alır.
   - **İletişim ve kimlik bilgileri:** Telefon, e-posta, TCKN ve IBAN maskelenir.
   - **Onay:** Veri gönderilmeden önce onayınız alınır.
4. **Yazar (model):** Model şu bölümleri yazar: firma ve talep, mali değerlendirme, istihbarat, teminatlar, güçlü
   yönler, riskler, önerilen şartlar, eksik belgeler ve portföy yöneticisi görüşü (taslak). Model metnindeki sayılar
   girdilerle karşılaştırılır; girdilerde olmayan sayılar **işaretlenir**.
5. **Rapor:**
   - `.md` teklif özeti: tablolar koddan, takma adlar geri açılmış olarak.
   - Excel: Özet, Kontroller ('Portföy Yöneticisi Onayı' sütunuyla), Rasyolar, KKB Risk ve Teminatlar sayfaları.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                     # örnek: kurgusal ambalaj firması, 20 milyon TL talep
python agent.py --girdi ./teklif_klasoru --gizle "Ortak Adı" "Kefil Adı"
python agent.py --girdi ./teklif --katsayilar katsayilar.csv         # bankanızın teminat katsayıları
```

| Dosya | Biçim |
|---|---|
| `talep_*.txt` | `Firma Unvanı:`, `VKN:`, `Talep:`, `Talep Edilen Toplam Limit: … TL`, faaliyet, ortaklar, amaç… (satır başına bir alan) |
| `mali_tablolar.csv` | İlk sütun `Kalem`, sonraki sütunlar yıllar. Kalemler: Dönen Varlıklar, Hazır Değerler, Ticari Alacaklar, Stoklar, Toplam Aktif, Kısa/Uzun Vadeli Yabancı Kaynaklar, Kısa/Uzun Vadeli Banka Kredileri, Özkaynaklar, Net Satışlar, Satışların Maliyeti, Brüt Kâr, Faaliyet Giderleri, Amortisman, [FAVÖK], Finansman Giderleri, Net Kâr |
| `kkb_risk.csv` | Banka, Nakdi Limit, Nakdi Risk, Gayrinakdi Limit, Gayrinakdi Risk |
| `teminatlar.csv` | Teminat Türü, Açıklama, Değer |
| `istihbarat*.txt` ve diğer notlar | Serbest metin |

**Teminat katsayıları (örnektir)**

| Teminat | Katsayı |
|---|---|
| İpotek | 0,75 |
| Nakit / mevduat blokajı, KGF | 1 |
| Müşteri çeki | 0,5 |
| Araç rehni | 0,5 |
| Makine rehni | 0,4 |
| Alacak temliki | 0,4 |
| Ticari işletme rehni | 0,3 |
| Senet | 0,3 |
| Şahsi kefalet | 0 |

Kendi oranlarınız için `Tür;Katsayı` biçiminde bir CSV verin. Katsayı 0,6 veya 60 olarak yazılabilir.

## Çıktı

`kredi_teklif_ozeti.md` · `kredi_teklif_ozeti.xlsx`: `Özet` · `Kontroller` · `Rasyolar` · `KKB Risk` · `Teminatlar`

## Dikkat

- **Kararın sahibi:** Çıktı bir **taslaktır**. Kredi kararı, derecelendirme ve fiyatlama bankanızın yetkili organlarına
  ve iç modellerine aittir. Eşikler ve katsayılar örnektir.
- **Mali tablolar:** VUK bilançoları kullanılır. Bağımsız denetimli veya TFRS/BOBİ FRS tablolar varsa onları tercih
  edin; ara dönem verisini ayrıca ekleyin.
- **Okunamayan dosyalar:** Taranmış PDF'ler ve fotoğraflar okunmaz; atlanan dosyalar raporda listelenir.

## Testler

Testler gerçek API çağırmaz.
- **Rasyolar:** Örnek dosyadaki rasyolar elle hesaplanan değerlerle karşılaştırılır.
- **Kontroller:** Bilerek bırakılmış bulgular test edilir: finansman giderinin hızlı artışı, limit doluluğu, %90
  teminat karşılaması, karşılıksız çek, tedarikçi gecikmesi.
- **Güvenlik:** Maskeleme ve uydurma sayının yakalanması test edilir.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. API kullanım ücretleri ve modele gönderilen verinin hukuka ve banka politikasına
uygunluğu kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
