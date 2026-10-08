# Sık Sorulan Sorulara Cevap Taslağı · AI Agent

> Bilgi Teknolojileri › BT Destek Uzmanı · Workers / Workless

Çalışanların BT hizmet masasına sorduğu sorulara (VPN, parola, yazıcı, e-posta, yazılım…) **şirketin kendi bilgi
bankasından** cevap taslağı hazırlar. Her cevap, dayandığı bölümü gösterir. Bilgi bankasının cevaplayamadığı
sorular ayrıca listelenir; bu liste yeni yazılacak maddeler için yol gösterir.

## Nasıl çalışır?

1. **Kod, bilgi bankasını bölümlere ayırır (yapay zekâ yok):**
   - **Bölümleme:** Tek dosya ya da .md, .txt, .docx ve .pdf dosyalarından oluşan bir klasör verilebilir. Metin
     başlıklarına (`#`, `##`) göre bölünür. Başlık yoksa paragraflar birleştirilerek bölümlenir.
   - **Arama:** Her soru için en ilgili bölümler **BM25** aramasıyla bulunur. Türkçede kelimenin ilk 5 harfi kök
     kabul edilir.
   - **Gönderilen bölümler:** Modele bilgi bankasının tamamı değil, yalnız soruyla eşleşen bölümler (varsayılan en
     fazla 4) gönderilir.
   - **Maskeleme:** Sorudaki e-posta, telefon, IBAN ve TCKN maskelenir. "şifrem: …" gibi yazılmış parolalar da
     `[GİZLİ]` olarak maskelenir.
   - **Kural etiketleri:** Güvenlik olayı belirtisi (oltalama, virüs, hesap ele geçirme), arıza ifadesi
     ("açılmıyor", "hata veriyor") ve soruda yazılmış parola etiketlenir.
2. **Model, cevap taslağını yazar:**
   - Yalnız verilen bölümlere dayanır ve kullandığı bilginin kaynağını `[K3]` biçiminde gösterir.
   - Kapsamı belirtir: tam, kısmi veya yok.
   - Gerekirse tek bir netleştirme sorusu ekler.
   - Kullanıcıdan asla parola istemez.
3. **Kod, cevabı denetler (modelin üstündedir):**

| Durum | Kodun yaptığı |
|---|---|
| Modele verilmemiş bir bölüm kaynak gösterildi | Kaynak çıkarılır; kaynak kalmazsa kapsam "yok" sayılır |
| Taslak kullanıcıdan parola/PIN istiyor | İnsan onayı zorunlu, not düşülür |
| Güvenlik olayı belirtisi | Her zaman insana ve Bilgi Güvenliği'ne yönlendirilir |
| Arıza olabilir | Destek talebi açılması önerilir |
| Kapsam yok | İnsan gerekli; soru "Bilgi Bankası Boşlukları"na eklenir |

Kod ayrıca her cevaba bir **güven** düzeyi verir. Kapsam tam, arama puanı yüksek ve insan gerekmiyorsa güven
**yüksek**, kapsam tam veya kısmi ise **orta**, diğer durumlarda **düşük** olur.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                # örnek 9 soru, 3 belgelik bilgi bankası
python agent.py --girdi sorular.xlsx --bilgi ./bilgi_bankasi
python agent.py --soru "Telefonuma şirket mailini nasıl kurarım?" --bilgi ./bilgi_bankasi
python agent.py --girdi sorular.xlsx --bilgi BT_SSS.docx --k 6
```

**Bilgi bankası ipuçları:**
- **Başlıklar:** Her konuyu ayrı bir başlık altında yazın. Başlıklar aramada iki kat ağırlık taşır.
- **Bağlam:** Bölümler modele tek başına gönderilir. Bu yüzden her bölüm kendi içinde anlaşılır olmalı ("yukarıda
  anlatıldığı gibi" demeyin).
- **Taranmış PDF'ler:** Önce OCR'dan geçirin.

## Çıktı

`Özet` · `Cevaplar` ("Onay" sütunu dahil) · `Kaynak Eşleşmeleri` (her sorunun bulduğu bölümler ve puanları) ·
`Bilgi Bankası Boşlukları`

## Veri gizliliği

- **Gönderilen veri:** Modele yalnız soru metni (maskeli) ve eşleşen bilgi bankası bölümleri gider. Kullanıcı adı
  ve departman gönderilmez.
- **Bilgi bankası içeriği:** İç IP adresleri, sunucu adları ve yönetici parolaları gibi gizli bilgiler bilgi
  bankasında yer almamalıdır. Ollama ile yerel model kullanırsanız veri bilgisayarınızdan çıkmaz.

## Testler

Testler gerçek API çağırmaz. Arama kalitesi örnek sorularla test edilir. Sahte model bilerek üç hata yapar:
uydurma kaynak gösterir, parola isteyen bir taslak yazar ve güvenlik olayını insana yönlendirmez. Testler kodun bu
üç hatayı yakaladığını doğrular.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. API kullanım ücretleri ve modele gönderilen verinin hukuka uygunluğu
kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
