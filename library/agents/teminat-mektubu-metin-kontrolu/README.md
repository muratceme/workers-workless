# Teminat Mektubu Metin Kontrolü · AI Agent

> Bankacılık Operasyonları › Dış Ticaret Operasyon Uzmanı · Workers / Workless

Lehtarın ya da müşterinin istediği **teminat mektubu metnini bankanın standart metni ve talep formuyla karşılaştırır**.
Lehtar, borçlu, tutar, yazıyla tutar, para birimi ve vadeyi **kod kontrol eder**. Hükümleri model cümle cümle
sınıflandırır; faiz, kendiliğinden uzama, devir / temlik gibi riskli hükümler "aynı" ya da "düşük risk" sayılırsa
sonuç modele geri gönderilir.

## Nasıl çalışır?

1. **Hazırlık (kod):** Standart metindeki yer tutucular formdaki bilgilerle doldurulur; tutar yazıyla da yazılır
   (2.450.000,00 → "İki Milyon Dört Yüz Elli Bin Türk Lirası"). İki metin cümlelere bölünür (S1 …, T1 …). Lehtar ve borçlu
   unvanları `{LEHTAR}`, `{BORCLU}` olarak gönderilir.
2. **Alan kontrolü (kod):**
   - Formdaki lehtar, borçlu ve sözleşme talep metninde aynen geçiyor mu
   - Talep metnindeki rakamla tutarlar ve para birimi formla aynı mı; yazıyla tutar rakamla uyuşuyor mu
   - Geçerlilik tarihi formdaki vadeyle aynı mı (form "Süresiz" ise metinde tarih olmamalı); vade düzenleme tarihinden sonra mı
   - Standart metindeki hangi cümleler talep metninde yok (ör. "Bu tarihe kadar bildirilmeyen talepler hükümsüzdür")
3. **Karşılaştırma (model):** Her talep cümlesi: Aynı / İfade farkı / Esaslı sapma / Standartta yok, risk (Düşük / Orta /
   Yüksek) ve açıklama; genel değerlendirme.
4. **Denetim (kod):** Aşağıdakilerden biri varsa sonuç modele geri gönderilir (varsayılan en fazla 2 tur):
   - Sınıflandırılmamış ya da iki kez sınıflandırılmış talep cümlesi; geçersiz standart cümle numarası
   - "Aynı" denen cümle standarttan farklı (benzerlik %90'ın altında ya da sayı / tarih farklı)
   - Standartta olmayan riskli hüküm (faiz, kendiliğinden uzama, devir / temlik, yabancı hukuk / tahkim, süresizlik,
     kısmi ödeme, ödeme süresi / ek koşul) esaslı sapma ve en az orta risk olarak işaretlenmemiş
   - Açıklamada metinlerde olmayan sayı

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                         # örnek (kurgusal) kesin teminat talebi
python agent.py --form talep.txt --standart banka_standart.docx --metin lehtar_metni.pdf
```

## Çıktı

- `cikti/teminat_kontrol.xlsx`: Karşılaştırma (yeşil: aynı, sarı: ifade farkı, kırmızı: esaslı sapma / standartta yok /
  eksik standart hüküm), Alan Kontrol.
- `cikti/teminat_kontrol.txt`: alan sorunları ve esaslı sapmaların özeti.

## Dikkat

- **Ön kontroldür.** Standart dışı metinlerin kabulü bankanın hukuk ve kredi birimlerinin onayına tabidir. Kamu
  ihalelerinde idarenin standart formu, dış ticarette URDG gibi kurallara atıf ve muhabir banka metinleri ayrıca
  değerlendirilmelidir.
- **Standart metin sizin.** Örnekteki standart metin kurgusaldır; kendi bankanızın onaylı metnini kullanın.
- **Banka sırrı.** Talep metni yapay zekâ sağlayıcısına gönderilir (unvanlar kodlanır). Kurum politikanız izin vermiyorsa
  yerel model (Ollama) kullanın.
- **Örnek veri:** Şirketler ve sözleşme kurgusaldır.

## Testler

Testler gerçek API çağırmaz. Sahte model ilk sonuçta farklı geçerlilik tarihli cümleyi "aynı", kendiliğinden uzama
hükmünü "ifade farkı / düşük risk" sayar, devir-temlik cümlesini atlar ve açıklamaya olmayan bir süre yazar. Kodun dört
sorunu da yakaladığı; alan kontrolünün vade farkını, değiştirilmiş tutar ve para birimini, yazıyla tutar uyumsuzluğunu ve
eksik standart hükmü bulduğu doğrulanır.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. API kullanım ücretleri ve mektup metninin kabulü kullanıcının sorumluluğundadır. Bkz.
`LICENSE` ve `SORUMLULUK_REDDI.md`.
