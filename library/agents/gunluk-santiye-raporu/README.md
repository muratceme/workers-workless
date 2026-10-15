# Günlük Şantiye Raporu · AI Agent

> Şantiye › Saha Mühendisi · Workers / Workless

Gün içinde tutulan dağınık saha notlarından **günlük şantiye raporu** hazırlar: hava durumu, ekip / taşeron bazında
personel, makine-ekipman, yapılan imalatlar, malzeme girişleri, sorunlar, İSG olayları, ziyaretler ve ertesi günün planı.
Her madde not satırına dayanır; sayılar notlardaki değerlerle aktarılır.

## Nasıl çalışır?

1. **Ön inceleme (kod):** Notlar satır satır numaralanır. İSG anahtar kelimesi geçen satırlar bulunur: kaza, yaralanma,
   burkulma, düşme, çarpma, korkuluk, iskele, emniyet kemeri, baret, yangın, elektrik çarpması, göçük, ramak kala.
2. **Rapor (model):** Notlar bölümlere ayrılır; her madde dayandığı satır numaralarını gösterir.
3. **Denetim (kod):** Aşağıdakilerden biri varsa rapor modele geri gönderilir (varsayılan en fazla 2 tur):
   - Dayanağı olmayan ya da notlarda bulunmayan satıra dayanan madde
   - Dayandığı satırda yazmayan kişi sayısı, makine adedi, miktar ya da sıcaklık
   - İSG bölümüne girmeyen İSG satırı
   - Hiçbir bölümde kullanılmayan not satırı
4. **Toplamlar (kod):** Toplam personel ve makine sayısını kod hesaplar.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                      # örnek saha notları
python agent.py --girdi saha_notlari.docx
python agent.py --girdi saha_notlari.txt --tur 3
```

## Çıktı

- `cikti/gunluk_santiye_raporu.xlsx`:
  - `Günlük Rapor`: başlık bilgileri, toplamlar ve bölüm bölüm maddeler (not satırıyla).
  - `Notlar`: her satırın kullanıldığı bölüm; kullanılmayan satırlar kırmızı, İSG satırları sarı.
- `cikti/gunluk_santiye_raporu.md`

## Dikkat

- **Taslaktır.** Şantiye şefi notlarla karşılaştırıp onaylamalıdır.
- **Resmî kayıtların yerine geçmez.** Şantiye defteri, yapı denetim kayıtları ve iş kazası bildirimi (SGK'ya yasal sürede)
  ilgili mevzuata göre ayrıca yapılır.
- **İSG taraması anahtar kelimeye dayanır;** farklı ifade edilmiş olayları kaçırabilir.
- **Veri gönderimi.** Notlar seçtiğiniz yapay zekâ sağlayıcısına gönderilir. Telefon ve e-posta maskelenir; kişi adları
  maskelenmez. Gerekirse yerel model (Ollama) kullanın.
- **Örnek veri** kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk taslakta bilerek eksik korkuluk satırını yalnız sorunlara yazar, kalıpçı sayısını
yanlış verir ve yapı denetim ziyaretini atlar. Kodun bunları yakalayıp düzelttirdiği ve toplamları doğru hesapladığı
doğrulanır.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. API kullanım ücretleri ve raporun doğruluğu kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
