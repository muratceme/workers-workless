# Hasar Dosyası Evrak Eksik Kontrolü · Kod Bloğu

> Sigorta › Hasar Yönetimi › Hasar Uzman Yardımcısı · Workers / Workless

Branş ve hasar türüne göre **gerekli evrak listesini** dosyaya gelen evraklarla karşılaştırır. Sigortalıdan veya
hak sahibinden istenecekleri çıkarır ve her eksikli dosya için **evrak talep yazısı taslağı** hazırlar.
İnternete bağlanmaz.

## Ne yapar?

1. **Gerekli evrakları seçer:** `evrak_listesi.csv` içinden dosyanın branşına, hasar türüne ve koşullarına uyan
   satırlar alınır.
   - Koşullar: vekil, tüzel kişi, rehinli, resmi tutanak, yaralanma, ölüm, komşu.
   - Örnek: "vekil" koşulu varsa vekaletname, "rehinli" varsa rehin alacaklı muvafakatnamesi istenir.
2. **Gelen evrakları eşleştirir:** Evrak adları anahtar kelimelerle eşleştirilir; Türkçe ekler tolere edilir
   ("tutanak" → "tutanağı", "kaza" → "kazası").
   - Bir evrak birden çok gereği karşılayabilir: "Ehliyet ve ruhsat fotokopisi" hem sürücü belgesini hem ruhsatı
     karşılar.
   - Tanınmayan evraklar ayrıca listelenir.
3. **Eksikleri ayırır:**
   - Sigortalıdan / hak sahibinden istenecekler.
   - İç takip: ör. ekspertiz raporu (Kimden: "Şirket / eksper"). Bunlar talep yazısına girmez.
4. **Durum verir:**

   | Durum | Anlamı |
   |---|---|
   | Tamam | Eksik evrak yok |
   | Eksik evrak (N) | Sigortalıdan istenecek N evrak var |
   | Eksik evrak (N) · Hatırlatma | İhbardan bu yana `--hatirlatma-gun` (varsayılan 15) gün geçmiş |
   | İç evrak bekleniyor | Sigortalı evrakı tamam, iç takip bekliyor |
   | Evrak listesi tanımsız | Branş / hasar türü için listede satır yok |

5. **Uyarır:** Listede olmayan dosyaya gelen evrak; hasar tarihinden önce ihbar tarihi.

## Varsayılan evrak listesi

`evrak_listesi.csv` aşağıdaki branş ve hasar türleri için yaygın istenen belgeleri içerir:

| Branş | Hasar türleri |
|---|---|
| Kasko | çarpışma, hırsızlık, cam kırılması |
| Trafik | maddi, bedeni |
| Konut / İşyeri | yangın, hırsızlık, dahili su |

**Bu liste örnektir.** Şirketinizin hasar evrak listesine, ilgili genel şartlara ve dosyanın özelliğine göre
satır ekleyip çıkarın.

| Sütun | Açıklama |
|---|---|
| Branş, Hasar Türü | Birden çok değer `\|` ile; tümü için `*` |
| Evrak | Talep yazısında görünen ad |
| Anahtar Kelimeler | Gelen evrak adında aranacak ifadeler, `\|` ile |
| Koşul | Boşsa her dosyada; doluysa dosyanın "Koşullar" alanında bu ifadelerden biri varsa istenir |
| Kimden | "Sigortalı", "Hak sahibi", "Şirket / eksper"… "Şirket" veya "Eksper" ile başlayanlar iç takiptir |
| Not | Talep yazısında parantez içinde görünür |

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                          # örnek 9 dosya (rapor tarihi 08.10.2026)
python main.py --dosyalar dosyalar.xlsx --evraklar evraklar.xlsx
python main.py --dosyalar dosyalar.xlsx --klasor ./hasar_dosyalari      # alt klasör adı = dosya no, dosya adı = evrak
python main.py --dosyalar dosyalar.xlsx --evraklar evraklar.xlsx --evrak-listesi sirket_listesi.csv --hatirlatma-gun 10
```

| Dosya | Sütunlar |
|---|---|
| Dosyalar | Dosya No, Branş, Hasar Türü, Sigortalı (veya Hak Sahibi), Hasar Tarihi, İhbar Tarihi, Koşullar (`\|` veya virgülle) |
| Evraklar | Dosya No, Evrak, Geliş Tarihi |

## Çıktı

- **Excel:** `hasar_evrak_kontrolu.xlsx`:
  - `Özet`: durum dağılımı ve uyarılar.
  - `Dosyalar`: gerekli / gelen sayısı, eksikler, tanınmayan evraklar, durum; boş "Takip Notu" sütunu.
  - `Eksik Evraklar`: evrak başına satır; boş "İstendi mi? / Tarih" sütunu.
  - `Gelen Evraklar`: hangi gereği karşıladığı.
  - `Talep Yazıları`: yazı taslakları.
- **Talep yazıları:** `talep_yazilari/<dosya no>.txt`; şirket adını ve iletişim bilgisini ekleyin.

## Dikkat

- **Ad eşleştirmesi:** Evrakın içeriğini değil yalnız adını kontrol eder. Belgenin okunaklı, doğru kişiye ait ve
  geçerli olduğunu elle kontrol edin.
- **Listenin güncelliği:** Bir evrakın istenip istenmeyeceği genel şartlara, şirket uygulamasına ve dosyanın
  özelliğine bağlıdır. Talep yazısını göndermeden önce gözden geçirin.
- **Kişisel veri:** Hasar dosyaları kişisel veri (sağlık verisi dahil) içerebilir; çıktıyı yetkisiz kişilerle
  paylaşmayın.
- **Örnek veri:** Örnek dosyalar ve kişiler kurgusaldır.

## Testler

Şunlar test edilir:
- Örnek 9 dosyanın durumları; koşullu evraklar; ad eşleştirme ("fotokopisi" fotoğraf sayılmaz).
- Talep yazısında iç evrak olmaması; hatırlatma eşiği.
- Klasör modu ve uyarılar.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
