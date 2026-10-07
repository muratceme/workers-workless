# Müşteri Talebi Sınıflandırma ve Cevap Taslağı · AI Agent

> Müşteri Hizmetleri ve Çağrı Merkezi › Müşteri Temsilcisi · Workers / Workless

E-posta, form ve sosyal medyadan gelen müşteri taleplerini **kategori, aciliyet, duygu ve ilgili ekibe** göre
sınıflandırır. Her talep için **yalnızca şirketinizin bilgi bankasına dayanan** bir cevap taslağı hazırlar.
Hiçbir cevap otomatik gönderilmez; her taslak için "Temsilci Onayı" sütunu vardır.

## Nasıl çalışır?

1. **Kod (yapay zekâ yok):**
   - Sipariş numaralarını bulur.
   - Aynı göndericinin tekrar eden taleplerini bağlar.
   - Kural etiketleri koyar: **KVKK başvurusu**, **hukuki risk** (avukat, dava, tüketici hakem heyeti,
     şikâyet sitesi, CİMER), **acil**, **tekrar yazıyor**.
   - E-posta, telefon, IBAN ve TCKN'yi maskeler. Gönderici adresi modele hiç gönderilmez.
2. **Model:** Her talep sınıflandırılır ve bilgi bankasına göre cevap taslağı yazılır. Model sipariş durumu,
   tarih, tutar, tazminat gibi bilgi bankasında olmayan bilgileri uyduramaz. Bilgi yoksa veya bir işlem
   gerekiyorsa talep **insan gerekli** olarak işaretlenir ve yapılacak işlem yazılır.
3. **Kod kuralları modelin üstündedir:**
   - KVKK ve hukuki risk etiketli talepler her zaman insana yönlendirilir ve en az "yüksek" aciliyet alır.
   - Tekrar eden talepler "düşük" aciliyette kalmaz.
   - KVKK başvurularına **30 günlük yasal cevap süresi** (KVKK md. 13) yazılır.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                      # örnek taleplerle
python agent.py --girdi talepler.xlsx --bilgi bilgi_bankasi.md
python agent.py --girdi ./gelen_kutusu --bilgi bilgi_bankasi.md      # .eml / .txt klasörü
python agent.py --girdi talepler.xlsx --bilgi bb.md --kategoriler kategoriler.json
```

- **Bilgi bankası:** Kargo, iade, değişim, fatura, KVKK ve iletişim kurallarınızı başlıklar hâlinde yazın
  (bkz. `ornek_veri/bilgi_bankasi.md`). Cevapların kalitesi bu dosyanın güncelliğine bağlıdır.
- **Kategoriler:** `{"kategoriler": [...], "ekipler": [...]}` biçiminde bir JSON dosyasıyla kendi listenizi verin.

## Çıktı

`Özet` (kategori, aciliyet ve ekip dağılımı) · `Talepler` (aciliyete göre sıralı, cevap taslakları ve yapılacak işlemler)

## Testler

Testler gerçek API çağırmaz. Sahte model bilerek hatalı sınıflandırma yapar ve kod kurallarının bunu
düzelttiği doğrulanır.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. API kullanım ücretleri ve modele gönderilen verinin hukuka uygunluğu
kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
