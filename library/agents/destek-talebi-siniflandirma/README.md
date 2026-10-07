# Destek Talebi Sınıflandırma · AI Agent

> Bilgi Teknolojileri › BT Destek Uzmanı · Workers / Workless

BT hizmet masasına e-posta veya ticket sisteminden gelen talepleri **kategori, ekip, etki, aciliyet ve türe**
(olay / hizmet talebi / güvenlik olayı) göre sınıflandırır. **ITIL önceliğini** ve **SLA hedef zamanlarını**
hesaplar, kullanıcıya ilk yanıt taslağı ve destek uzmanına çözüm önerisi yazar.

## Nasıl çalışır?

1. **Kod (yapay zekâ yok):**
   - **Güvenlik olayı belirtileri:** oltalama bağlantısına tıklama, şifre girme, virüs, fidye yazılımı, şüpheli
     giriş
   - **Olası yaygın kesinti:** Aynı konuda (ağ/internet, e-posta, ERP) kısa sürede gelen çok sayıda talep;
     varsayılan 30 dakikada 4 ve üzeri. Bu talepler tek bir **ana olay** altında gruplanır.
   - E-posta, telefon, IBAN ve TCKN maskelenir; gönderen bilgisi modele gitmez.
2. **Model:** Her talebi sınıflandırır, etki ve aciliyeti belirler, ilk yanıtı ve çözüm önerisini yazar. Şifre
   sıfırlamada kullanıcıdan şifre istemez. Güvenlik olayında güvenli ilk adımları verir.
3. **Kod, öncelik ve SLA:**
   - **ITIL etki × aciliyet matrisi** (P1–P5) uygulanır ve `sla.json`'dan ilk yanıt ve çözüm hedef zamanları
     hesaplanır.
   - Kod kuralları modelin üstündedir: güvenlik olayı her zaman **Bilgi Güvenliği** ekibine gider ve en az
     **P2** olur. Yaygın kesinti kümesindeki talepler de en az P2 olur.

| Etki \ Aciliyet | Yüksek | Orta | Düşük |
|---|---|---|---|
| Yüksek | P1 | P2 | P3 |
| Orta | P2 | P3 | P4 |
| Düşük | P3 | P4 | P5 |

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                   # örnek 10 talep (5'i aynı internet kesintisi)
python agent.py --girdi talepler.xlsx --sla sla.json --pencere 20 --esik 3
```

## Çıktı

`Yaygın Kesinti` (ana olay önerisi) · `Talepler` (önceliğe göre sıralı, "Onay") · `Özet`

## Testler

Testler gerçek API çağırmaz. Sahte model bilerek düşük öncelik verir. Kodun güvenlik olayını ve yaygın kesintiyi
P2'ye yükselttiği ve SLA zamanlarını doğru hesapladığı doğrulanır.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. API kullanım ücretleri ve modele gönderilen verinin hukuka uygunluğu
kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
