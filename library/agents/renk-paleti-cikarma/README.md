# Renk Paleti Çıkarma · AI Agent

> Tasarım ve Koleksiyon › Moda Tasarımcısı · Workers / Workless

İlham görsellerinden **baskın renkleri çıkarır**, bunları **sizin renk referans tablonuzdaki** (ör. lisanslı Pantone
TCX listeniz) en yakın kodlarla eşleştirir ve bir **koleksiyon renk paleti** oluşturur. Renk çıkarma, eşleştirme ve
nitelikler (sıcak / soğuk, açık / koyu) **kodla** hesaplanır; görseller yapay zekâya gönderilmez. Model paletin adını,
hikâyesini ve renklerin rollerini yazar; uydurma kod, yanlış yüzde ya da yanlış niteleme yakalanır.

## Nasıl çalışır?

1. **Renk çıkarma (kod):** Her görsel küçültülür ve Pillow ile renk sayısı azaltılır (median cut). Tüm görsellerden
   gelen renkler eşit ağırlıkla birleştirilir; CIEDE2000 farkı `--birlestir` (varsayılan 8) altında kalan renkler tek renk
   sayılır. En büyük paylı `--renk` (varsayılan 6) renk alınır.
2. **Eşleştirme (kod):** Her renk referans tablosundaki en yakın renkle CIEDE2000 (ΔE00) ile eşleştirilir; ΔE00 >
   `--esik` (varsayılan 3) ise "yakın eşleşme yok" yazılır. Sıcaklık (CIELAB ton açısı ve doygunluk) ve açıklık (L*)
   hesaplanır.
3. **Palet (model):** Palet adı ve hikâyesi; her renge rol (Ana / Destek / Vurgu), tanım ve kullanım önerisi;
   kombinasyonlar.
4. **Denetim (kod):** Aşağıdakilerden biri varsa sonuç modele geri gönderilir (varsayılan en fazla 2 tur):
   - Rolü olmayan ya da iki rol alan renk; ana renk en büyük paylı iki renkten biri değil; vurgu rengi ana renkten büyük
   - Paletle eşleşmeyen referans kodu ya da tabloda olmayan "19-4052" gibi bir kod (kod uydurulamaz)
   - Kodun hesabında olmayan yüzde; kodun hesabıyla çelişen "sıcak / soğuk / açık / koyu" nitelemesi
5. **Çıktı (kod):** Excel (örnek hücreler gerçek renkle boyalı) ve HTML palet kartı.

## Kurulum

```bash
pip install -r requirements.txt        # Pillow dahil
cp .env.example .env                   # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                                        # örnek (kurgusal) ilham görselleri
python agent.py --gorseller ilham/ --referans pantone_tcx.csv --renk 8 --brif "İlkbahar-Yaz kadın keten"
```

## Çıktı

- `cikti/renk_paleti.xlsx`: Palet sayfası.
- `cikti/renk_paleti.html`: tarayıcıda açılan palet kartı.

## Dikkat

- **Referans tablosu sizin.** Pantone renk kodları ve değerleri Pantone LLC'nin tescilli verisidir; bu pakette yer
  almaz. `ornek_veri/referans_renkler.csv` kurgusal kodlardan oluşur. Kendi lisanslı listenizi "Kod; Ad; HEX" olarak
  kaydedin. HEX değerleri ekran için yaklaşık değerlerdir.
- **Ekran ≠ kumaş.** Fotoğraftaki renk ışığa, kameraya ve ekrana göre değişir. Renk onayını fiziksel renk kartı ve
  laboratuvar boyama örneği (lab dip) ile standart ışık kabininde yapın.
- **Örnek veri:** Görseller kodla üretilmiş renk bloklarıdır.

## Testler

Testler gerçek API çağırmaz. CIEDE2000 hesabı yayımlanmış test verisiyle (Sharma vd., 2005) doğrulanır. Sahte model ilk
çıktıda en büyük paylı rengi vurgu, küçük paylı rengi ana renk yapar, soğuk gece mavisine "sıcak" der, tabloda olmayan
bir kod ve hesapta olmayan "%40" yazar, bir rengi atlar. Kodun altı sorunu da yakaladığı doğrulanır.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. API kullanım ücretleri ve renk onayı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
