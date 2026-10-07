# Hasar Dosyası Özeti · AI Agent

> Sigorta › Hasar Yönetimi › Hasar Uzmanı · Workers / Workless

Bir hasar dosyasındaki belgeleri okuyup **karar için tek sayfalık özet** hazırlar. Belgeler poliçe özeti, ihbar
ve beyan, kaza tespit tutanağı, ekspertiz raporu ve faturalar olabilir. Belgeler arası tutarsızlıkları
**kodla** bulur; değerlendirmeyi model yazar, kararı hasar uzmanı verir.

## Nasıl çalışır?

1. **Belgeleri ayırır (kod):** Klasördeki `.pdf`, `.docx` ve `.txt` dosyaları okunur ve türüne göre ayrılır:
   poliçe, ihbar/beyan, tutanak, ekspertiz, fatura, ehliyet, ruhsat.
2. **Kontrol eder (kod):** Bu bulgular kesindir ve modele "değiştirme" talimatıyla verilir.
   - Belgelerde farklı yazılmış **hasar tarihi**
   - Sigortalı plakasına benzeyen ama farklı **plaka** (ör. faturada `34 ABC 128`)
   - **Poliçe dönemi dışında** hasar ve poliçe başlangıcına yakın (30 gün) hasar
   - Fatura KDV hariç tutarının **eksper onaylı tutarı aşması**
   - İhbar gecikmesi ve **eksik evrak** (varsayılan liste veya `--evrak` dosyası)
3. **Maskeler:** Plakalar her belgede aynı takma adla (`[PLAKA-A]`) değiştirilir. `--gizle` ile verilen kişi
   adları `[GİZLİ-1]` olur. Telefon, e-posta, TCKN ve IBAN maskelenir. Veri gönderilmeden önce onayınız alınır.
4. **Özetler (model):** Olay, kusur ve rücu, teminat değerlendirmesi (taslak), tutarsızlıklar (kaynak belge
   adlarıyla), açık sorular, önerilen adımlar ve karar özeti yazılır. Modelin uydurduğu belge adları
   "kaynak doğrulanamadı" olarak işaretlenir.
5. **Rapor:** Tek sayfalık `.md` özet ve Excel üretilir. Rapor şirket içinde kaldığı için takma adlar
   gerçek değerlere geri çevrilir.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                          # örnek kasko dosyasıyla
python agent.py --girdi ./dosya_2026_0457 --gizle "Sigortalı Adı" "Karşı Sürücü"
python agent.py --girdi ./dosya --evrak evrak_listesi.txt                # branşınıza göre evrak listesi
```

**Evrak listesi:** Her satıra bir belge türü yazılır. Kullanılabilecek türler: Poliçe, İhbar / beyan, Kaza
tespit tutanağı, Ekspertiz raporu, Fatura, Ehliyet, Ruhsat.

Fotoğraflar ve taranmış PDF'ler okunmaz; bunlar "okunamayan" olarak listelenir.

## Çıktı

`hasar_ozeti.md` (tek sayfa) · `hasar_ozeti.xlsx`: `Özet` · `Kontroller` (önem, bulgu, kaynak, 'Uzman Onayı') · `Belgeler`

## Testler

Testler gerçek API çağırmaz. Örnek dosyada bilerek bırakılmış dört tutarsızlığın ve iki eksik evrakın kodla
bulunduğu doğrulanır.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. API kullanım ücretleri ve modele gönderilen verinin hukuka uygunluğu
kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
