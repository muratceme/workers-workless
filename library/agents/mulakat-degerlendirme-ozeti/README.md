# Mülakat Değerlendirme Özeti · AI Agent

> İnsan Kaynakları › İşe Alım · Workers / Workless

Birden fazla mülakatçının **puan ve notlarını** aday başına **tek sayfalık karşılaştırmalı özete** çevirir.
Kararı işe alım yöneticisi verir; bu araç kararı kanıta dayalı hâle getirir.

## Nasıl çalışır?

1. **Hesap (kod):**
   - Aday × yetkinlik ortalaması ve yetkinlik ağırlıklarıyla **toplam puan** (5 üzerinden), sıralama
   - Mülakatçılar arası **puan farkı ≥ 2** olan yetkinlikler: kalibrasyon toplantısında konuşulmalı
   - Eksik puanlar
2. **Önyargı kontrolü (kod):** Notlar ortak ayrımcılık tarayıcısından geçer. Korunan özelliğe (yaş, medeni hâl,
   aile, sağlık, din, köken...) dayalı yorum içeren notlar **"değerlendirme dışı"** işaretlenir. Örnek: "evli ve
   iki çocuklu olduğu için sahada zorlanabilir". Bu puanlar hariç tutularak **ikinci bir toplam** hesaplanır;
   sıralama bu toplama göre yapılır.
3. **Gizlilik:** Aday ve mülakatçı adları modele gitmeden **takma adla** (A1, M1) değiştirilir. Rapor şirket
   içinde kaldığı için raporda gerçek adlar gösterilir.
4. **Özet (model):** Her aday için şunlar yazılır:
   - Güçlü yönler ve gelişim alanları (notlardan somut kanıtla)
   - Görüş ayrılıkları ve kalibrasyon sorusu
   - Referansta doğrulanacak noktalar
   - Öneri taslağı (ilerlet / beklet / ilerletme) ve yalnız işe dayalı gerekçe

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                         # örnek: 2 aday, 2 mülakatçı, 5 yetkinlik
python agent.py --girdi mulakat_notlari.xlsx --yetkinlikler yetkinlikler.xlsx
```

**Girdi:** Her satır bir mülakatçının bir yetkinlik için verdiği puan (1–5) ve nottur. **Mülakat Soru Seti
Hazırlama** agent'ının değerlendirme formundaki puanlar bu biçime kolayca aktarılır.

## Çıktı

`Karşılaştırma` (uyumsuz puanlar kırmızı, "Karar" sütunu) · `Aday Özetleri` · `Notlar` (değerlendirme dışı
işaretleri) · `Bilgi`

## Testler

Testler gerçek API çağırmaz. Ağırlıklı toplamlar elle hesaplanmış değerlerle karşılaştırılır (3,80 / 3,325 /
3,45). Adların modele gitmediği ve önyargılı notun işaretlendiği doğrulanır.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. API kullanım ücretleri ve işe alım kararının hukuka uygunluğu kullanıcının
sorumluluğundadır. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
