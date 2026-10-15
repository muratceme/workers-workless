# Karar ve İçtihat Özeti · AI Agent

> Dava Takip › Avukat · Workers / Workless

Mahkeme kararını dosyaya ve içtihat arşivine girilecek biçimde **başlıklarla özetler**: künye, taraflar, olay, iddia ve
savunma, yargılama süreci, hukuki mesele, gerekçe, hüküm, oy durumu, atıf yapılan mevzuat ve kararlar, anahtar kelimeler.
Her bölüm karar metnindeki satırlara dayanır; numaralar ve hüküm alıntısı kodla metne karşı denetlenir.

## Nasıl çalışır?

1. **Hazırlık (kod):** Karar metni satır (paragraf) satır numaralanır.
2. **Özet (model):** Bölümler yazılır; her bölüm dayandığı satırları gösterir. Hüküm fıkrasından aynen alıntı yapılır.
3. **Denetim (kod):** Aşağıdakilerden biri varsa özet modele geri gönderilir (varsayılan en fazla 2 tur):
   - Esas no, karar no ya da karar tarihi metinde geçmiyor
   - Mevzuat maddesinin numarası dayandığı satırda yazmıyor
   - Atıf yapılan kararın esas / karar numarası metinde yok
   - Hüküm alıntısı metinde yok
   - Oy durumu metindeki "oybirliği / oyçokluğu" ifadesiyle çelişiyor
   - Dayanak satırı yok ya da metinde bulunmuyor

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                    # örnek (kurgusal) karar metni
python agent.py --girdi karar.pdf
python agent.py --girdi karar.docx --tur 3
```

## Çıktı

- `cikti/karar_ozeti.xlsx`:
  - `Karar Özeti`: bölümler ve dayanak satırları.
  - `Mevzuat ve Atıflar`
  - `Metin`: numaralı karar metni.
- `cikti/karar_ozeti.md`: arşive eklenecek özet.

## Dikkat

- **Özet hukuki görüş değildir.** Kullanmadan önce kararın tam metniyle ve resmî kaynakla (UYAP, Yargıtay / Danıştay karar arama)
  karşılaştırın. Kararın kesinleşip kesinleşmediğini ve sonradan değişen mevzuatı ayrıca kontrol edin.
- **Denetimin sınırı.** Kod numaraların ve alıntıların metinde geçtiğini doğrular; gerekçenin doğru özetlendiğini doğrulamaz.
- **Kişisel veri.** Karar metinleri taraf adları ve sağlık, ceza gibi özel nitelikli veriler içerebilir. Telefon, e-posta ve TCKN
  maskelenir; adlar maskelenmez. Gerekirse yerel model (Ollama) kullanın ya da anonimleştirilmiş metin verin.
- **Örnek veri.** Örnek karar metni kurgusaldır; daire, esas / karar numaraları ve atıf yapılan karar gerçek değildir.
  Mevzuat atıfları (TBK 49, 52; HMK 266) gerçek maddelerdir.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk taslakta bilerek karar numarasını yanlış yazar, bir kanun maddesini başka bir
satıra dayandırır ve oybirliğiyle verilen kararı "oyçokluğu" diye işaretler. Kodun bunları yakalayıp düzelttirdiği doğrulanır.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. API kullanım ücretleri ve özetin doğruluğu kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
