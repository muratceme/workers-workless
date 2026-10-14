# Vardiya Devir Notu Özeti · AI Agent

> Üretim › Vardiya Amiri · Workers / Workless

Biten vardiyanın **defter notlarını**, gelen vardiyanın hızla okuyabileceği bir **devir notuna** çevirir: vardiya bilgisi, üretim
durumu, açık sorunlar (öncelik, sonraki vardiyaya aksiyon, sorumlu), arıza ve duruşlar, güvenlik, kalite ve diğer bilgiler. Her
madde defterdeki satır numarasına dayanır; güvenlik olayları ve duruşlar atlanamaz.

## Nasıl çalışır?

1. **Ön inceleme (kod):**
   - Defter satır satır numaralanır.
   - Güvenlik anahtar kelimesi geçen satırlar bulunur: kaza, yaralanma, ramak kala, yangın, duman, kıvılcım, kaçak, sızıntı,
     dökülme, çarpma, düşme, kimyasal, acil durdurma, LOTO.
   - "Duruş" geçen satırlardaki süreler dakikaya çevrilir ("Duruş 45 dk", "1,5 saat").
2. **Yazım (model):** Devir notu bölümlere ayrılır. Her madde `kaynak_satirlar` ile defter satırlarını gösterir. Defterde
   olmayan neden, süre, kişi ya da numara yazılmaz.
3. **Denetim (kod):** Aşağıdaki durumlardan biri varsa taslak modele geri gönderilir (varsayılan en fazla 2 tur):
   - Kaynağı olmayan ya da defterde bulunmayan satıra dayanan madde
   - Güvenlik bölümünde yer almayan güvenlik satırı
   - Arıza / duruş bölümüne girmeyen duruş satırı
   - Defterdekinden farklı ya da defterde olmayan duruş dakikası
   - Hiçbir bölümde kullanılmayan defter satırı

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                            # örnek gece vardiyası defteriyle
python agent.py --girdi vardiya_defteri.docx --tur 3
```

## Çıktı

- `Devir Notu`: bölümler, öncelik / durum, sorumlu, defter satırı ve "Devralan Onayı" sütunu.
- `Defter`: her satırın hangi bölümde kullanıldığı; kullanılmayan satırlar kırmızı.
- `.md` devir notu.

## Dikkat

- **Taslaktır.** Devralan vardiya amiri defterle karşılaştırıp onaylamalıdır.
- **Resmî bildirimin yerine geçmez.** İş kazası ve ramak kala olaylarının işyeri prosedürüne göre bildirimi ve kaydı ayrıca
  yapılır. Güvenlik taraması anahtar kelimeye dayanır; farklı ifade edilmiş olayları kaçırabilir.
- **Veri gönderimi.** Defter, seçtiğiniz yapay zekâ sağlayıcısına gönderilir. Telefon ve e-posta maskelenir; personel adları
  maskelenmez. Gerekirse yerel model (Ollama) kullanın.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk taslakta bilerek bir güvenlik satırını atlar, bir duruş süresini yanlış yazar ve
bir satırı kullanmaz. Kodun bunları yakalayıp düzelttirdiği doğrulanır.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. API kullanım ücretleri ve devir notunun doğruluğu kullanıcının sorumluluğundadır. Bkz. `LICENSE`
ve `SORUMLULUK_REDDI.md`.
