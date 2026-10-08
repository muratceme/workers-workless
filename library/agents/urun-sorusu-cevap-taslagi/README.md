# Ürün Sorusu Cevap Taslağı · AI Agent

> E-Ticaret › Müşteri Deneyimi › Müşteri Hizmetleri Temsilcisi · Workers / Workless

Pazaryerlerinde (Trendyol, Hepsiburada, Amazon, N11…) ürün sayfasına gelen müşteri sorularına cevap taslağı
hazırlar. Taslaklar **ürün kartınıza** ve **şirket politikalarınıza** (kargo, iade, değişim) dayanır. Kod,
pazaryeri kurallarına aykırı veya doğrulanamayan cevapları yakalar.

## Nasıl çalışır?

1. **Kod, soruları hazırlar (yapay zekâ yok):**
   - **Ürün kartı:** Her soru ürün koduyla ürün kartına bağlanır.
   - **Politika eşleşmesi:** Politika metninden soruyla ilgili en fazla 2 bölüm BM25 aramasıyla bulunur. Arama
     çekirdeği "Sık Sorulan Sorulara Cevap Taslağı" paketinden gelir.
   - **Etiketler:** Ürün kartı bulunamayan, iletişim bilgisi isteyen ve sağlıkla ilgili sorular etiketlenir.
2. **Model, taslağı yazar:** Yalnız ürün kartı ve verilen politika bölümleriyle kısa bir cevap yazar ve hangi
   özelliğe dayandığını belirtir. Bilgi yoksa uydurmaz.
3. **Kod, taslağı denetler (modelin üstündedir):**

| Durum | Kodun yaptığı |
|---|---|
| Telefon, e-posta, web adresi, WhatsApp/Instagram, platform dışı yönlendirme | Reddedilir; pazaryeri kurallarına aykırı |
| Cevaptaki ölçü/ağırlık/sıcaklık/yüzde ürün kartında veya politikada yok | "Doğrulanamayan sayısal bilgi" olarak işaretlenir |
| Sağlık/tedavi iddiası | Kaldırılması istenir |
| "Yarın elinizde" gibi kesin teslim sözü | İşaretlenir |
| Verilmemiş bir kaynağa dayanma | Dayanaktan çıkarılır |
| Ürün kartı yok veya kapsam yok | İnsan gerekli |

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                           # örnek 7 soru
python agent.py --sorular sorular.xlsx --urunler urunler.xlsx --politika politikalar.md
python agent.py --sorular sorular.xlsx --urunler urunler.xlsx --politika ./politikalar --azami 300
```

| Dosya | Sütunlar |
|---|---|
| Sorular | Soru No, Tarih, Kanal, Ürün Kodu, Soru (pazaryeri panelinden alınan döküm) |
| Ürün kartları | İlk sütun ürün kodu (Stok Kodu/Barkod); diğer sütunlar özellikler (Malzeme, Ölçü, Yıkama…) |
| Politikalar | Başlıklı .md/.txt/.docx/.pdf (# Kargo, # İade …) veya bu dosyaların klasörü |

## Çıktı

`Özet` (en çok soru alan ürünler dahil) · `Cevaplar` ("Onay" sütunu dahil)

**İpucu:** Sık sorulan bir bilgi (ölçü, malzeme, yıkama) ürün kartında yoksa önce kartı tamamlayın. Hem cevaplar
güçlenir hem de soru sayısı azalır.

## Veri gizliliği

- **Gönderilen veri:** Modele soru metni (kişisel veriler maskeli), ilgili ürün kartı ve politika bölümleri gider.
- **Gönderilmeyen veri:** Müşteri adı gönderilmez.
- **Yerel model:** Ollama ile yerel model kullanırsanız veri bilgisayarınızdan çıkmaz.

## Testler

Testler gerçek API çağırmaz. Sahte model bilerek dört hata yapar:
- uydurma sayı yazar,
- kesin teslim sözü verir,
- WhatsApp numarası yazar,
- verilmeyen bir kaynağa dayanır.

Testler kodun bu dört hatayı yakaladığını doğrular.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. API kullanım ücretleri ve modele gönderilen verinin hukuka uygunluğu
kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
