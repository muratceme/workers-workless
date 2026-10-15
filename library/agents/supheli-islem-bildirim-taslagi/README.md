# Şüpheli İşlem Bildirim Taslağı · AI Agent

> Kurumsal Uyum › Uyum Uzmanı · Workers / Workless

İşlem dökümü ve inceleme notlarından **MASAK şüpheli işlem bildirimi için olay özeti ve şüphe gerekçesi taslağı**
hazırlar. Şüphe göstergelerini ve son bildirim tarihini **kod hesaplar**. Model yalnız bu göstergelere ve işlemlere
dayanarak yazar; kesin hüküm dili ya da müşteriye bildirim yapılacağını ima eden ifade kullanılırsa taslak geri gönderilir.

## Nasıl çalışır?

1. **Kodlama (kod):** Müşteri adı `{MUSTERI}`, karşı taraflar `{KT1}`, `{KT2}` … olur. Müşteri numarası gönderilmez.
   Açıklama ve gözlemlerdeki ad, soyad, telefon, IBAN ve e-posta maskelenir. Çıktıda adlar geri yazılır.
2. **Göstergeler (kod):**
   - **Hızlı aktarım:** aynı gün giren tutarın %90'ından fazlasının çıkması (`--oran`)
   - **Profil uyumsuzluğu:** dönem girişlerinin beyan edilen aylık gelire oranı
   - **Çok sayıda farklı gönderen:** 5 ve daha fazla farklı kişiden EFT / havale
   - **Çıkışlarda yoğunlaşma:** çıkışların en az %70'inin tek karşı tarafa gitmesi
   - **Eşik altı (parçalama):** `--esik` verilirse eşiğin %90'ı ile eşik arasındaki işlemler
3. **Süre (kod):** Şüphenin oluştuğu tarihten itibaren **10 iş günü** (hafta sonu, sabit resmî tatiller ve `--tatil` ile
   verilen günler hariç). Gecikmesinde sakınca bulunan hâllerde bildirim derhal yapılır.
4. **Taslak (model):** Olay özeti ve şüphe gerekçeleri. Her biri dayandığı işlem, gözlem (G) ve gösterge (K)
   numaralarını gösterir.
5. **Denetim (kod):** Aşağıdakilerden biri varsa taslak modele geri gönderilir (varsayılan en fazla 2 tur):
   - Dökümde ya da hesaplarda olmayan tutar veya tarih
   - Hiçbir gerekçede kullanılmayan gösterge; yalnız gözleme dayanan gerekçe
   - Kesin hüküm dili ("aklamıştır", "dolandırıcıdır", "kesinlikle") — bildirim şüpheye dayanır
   - Müşteriye bildirim yapılacağını ima eden ifade (5549 sayılı Kanun md. 4/2)

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                 # örnek (kurgusal) hesap
python agent.py --islemler dokum.xlsx --not inceleme.txt --esik 50000 --tatil 20.03.2026 21.03.2026
```

## Çıktı

- `cikti/supheli_islem_taslagi.txt`: iç taslak (gizli).
- `cikti/supheli_islem_taslagi.xlsx`: Taslak; İşlemler (göstergeye konu işlemler sarı).

## Dikkat

- **Gizlilik.** Bildirimde bulunulduğu, yükümlülük denetimi elemanları ve yargılama sırasında mahkemeler dışında,
  işlemin tarafları dâhil hiç kimseye açıklanamaz (5549 md. 4/2). Çıktıları yalnız yetkili uyum personeliyle paylaşın.
- **Karar sizde.** Bildirim kararı uyum görevlisinindir; bildirim MASAK'ın elektronik sistemi üzerinden, MASAK'ın
  güncel rehber ve formlarına göre yapılır. Agent yalnız metin taslağı hazırlar.
- **Veri gönderimi.** İşlem bilgileri (kodlanmış hâlde) yapay zekâ sağlayıcısına gönderilir. Kurumunuzun gizlilik ve
  dış hizmet politikası buna izin vermiyorsa yerel model (Ollama) kullanın ya da agent'ı çalıştırmayın.
- **Süre.** 10 iş günü süresi Suç Gelirlerinin Aklanmasının ve Terörün Finansmanının Önlenmesine Dair Tedbirler Hakkında
  Yönetmelik'e dayanır. Dini bayramlar her yıl değiştiği için `--tatil` ile girin; mevzuatın güncel hâlini kontrol edin.
- **Örnek veri:** Müşteri, kişiler ve şirket kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk taslakta bir tutarı yanlış yazar (140.000 yerine 145.000), üç göstergeyi
kullanmaz, gerekçeyi yalnız gözleme dayandırır, "suç gelirini kesinlikle aklamıştır" der ve müşteriye bildirim
yapılmasını önerir. Kodun yedi sorunu da yakaladığı, müşteri ve karşı taraf adlarının ve müşteri numarasının modele
gönderilmediği, son bildirim tarihinin (29 Ekim dâhil tatiller atlanarak) doğru hesaplandığı doğrulanır.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. API kullanım ücretleri ve bildirim kararı kullanıcının sorumluluğundadır. Bkz. `LICENSE`
ve `SORUMLULUK_REDDI.md`.
