# Sözleşme Ön İnceleme · AI Agent

> Hukuk › Avukat · Workers / Workless

Bir sözleşmeyi **sizin tarafınızın bakış açısından** okur ve şunları çıkarır:
- Riskli maddeler ve her birinin kimin lehine olduğu
- Gerekçe, müzakere önerisi ve alternatif madde metni
- 20 konuluk standart kontrol listesine göre **eksik** veya **belirsiz** düzenlemeler (mücbir sebep, KVKK,
  uygulanacak hukuk vb.)

> **Bu bir ön incelemedir, hukuki görüş değildir.** Rapordaki her öneri avukat tarafından değerlendirilmelidir
> ("Avukat Onayı" sütunu).

## Nasıl çalışır?

1. **Kod (yapay zekâ yok):** Sözleşme `.docx`, `.pdf` veya `.txt` dosyasından okunur ve "MADDE 5 – FESİH" ya da
   "5. FESİH" başlıklarından maddelere bölünür. Her madde konu anahtar kelimeleriyle etiketlenir.
2. **Maskeleme:** Modele gönderilmeden önce IBAN, TCKN, telefon ve e-posta maskelenir. `--gizle` ile verdiğiniz
   kişi ve firma adları `[GİZLİ-1]` gibi etiketlerle değiştirilir. Veri gönderilmeden önce onayınız alınır.
3. **Model:** Her madde için risk düzeyi (düşük / orta / yüksek), kimin lehine olduğu, gerekçe, öneri ve
   alternatif metin döner. Kontrol listesinde her konu için var / eksik / belirsiz durumu belirlenir.
4. **Rapor:** Excel ve kısa bir Markdown özeti üretilir. Uzun sözleşmeler bölümlere ayrılarak gönderilir.

Modelin yönergesi (`prompt.md`) yalnız doğrulanmış genel kurallara dayanmasını ister ve kanun maddesi numarası
uydurmasını yasaklar. Bu kurallar: TBK 115 (ağır kusurdan sorumsuzluk), TBK 182 / TTK 22 (ceza koşulu), TBK
20–25 (genel işlem koşulları), HMK 17 (yetki sözleşmesi), TTK 5/A (dava şartı arabuluculuk), damga vergisi,
KVKK 9 ve 12.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

`WW_PROVIDER=ollama` ile yerel bir model kullanırsanız sözleşme bilgisayarınızdan çıkmaz.

## Kullanım

```bash
python agent.py                                              # örnek sözleşmeyle (Müşteri adına)
python agent.py --girdi sozlesme.docx --taraf "Müşteri"
python agent.py --girdi sozlesme.pdf --taraf "Alıcı" --gizle "Ahmet Yılmaz" "Örnek A.Ş."
```

Taranmış (görüntü) PDF'lerden metin çıkarılamaz; önce OCR uygulayın. Eski `.doc` dosyalarını `.docx` olarak kaydedin.

## Çıktı

`Özet` (öncelikli aksiyonlar, genel değerlendirme) · `Madde Analizi` (yüksek riskten düşüğe) · `Kontrol Listesi`
· `Sözleşme Metni` · ayrıca aynı adla `.md` özeti

## Testler

Testler gerçek API çağırmaz; model yanıtı sahte bir fonksiyonla verilir.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. API kullanım ücretleri ve modele gönderilen verinin hukuka uygunluğu
kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
