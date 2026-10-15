# Tıbbi Rapor Yazım Kontrolü · AI Agent

> Tıbbi Sekreterlik › Tıbbi Sekreter · Workers / Workless

Hekimin dikte ettiği rapor metnini **yazım, tıbbi terim, taraf, birim ve kısaltma** açısından kontrol eder.
**Hasta kaydıyla tutarlılığı (yaş, cinsiyet, tarihler, TC kimlik no, protokol no, sağ / sol) kod denetler.** Model metni
yeniden yazmaz; yalnız alıntı ve öneri verir. Bir sayıyı değiştiren ya da anlamı değiştiren öneri "hekime sor"
işaretlenmeden kabul edilmez.

## Nasıl çalışır?

1. **Gizleme (kod):** Hasta kaydı modele gönderilmez. Metindeki hasta adı `{HASTA}` olur; TC kimlik no, protokol no,
   telefon ve e-posta maskelenir.
2. **Tutarlılık (kod):**
   - "45 yaşında" gibi ifadeler doğum tarihine göre rapor tarihindeki yaşla aynı mı
   - Cinsiyetle çelişen ifadeler (erkek hastada "kadın hasta", "gebe", "uterus"; kadın hastada "prostat" …)
   - Rapor tarihinden sonraki ya da doğum tarihinden önceki tarihler
   - Metindeki TC kimlik no ve protokol no kayıttan farklı mı
   - Rapor türü tek taraflıysa ("Sağ diz MR") karşı tarafın geçtiği cümleler
3. **Yazım kontrolü (model):** Yazım, tıbbi terim, taraf tutarlılığı, birim / ölçü, kısaltma, noktalama bulguları.
4. **Denetim (kod):** Aşağıdakilerden biri varsa sonuç modele geri gönderilir (varsayılan en fazla 2 tur):
   - Alıntı metinde aynen geçmiyor; öneri boş ya da alıntıyla aynı; aynı alıntı iki kez yazılmış
   - Öneri bir sayıyı değiştiriyor ama "hekime sor" işaretli değil (ör. "8 mm" → "18 mm")
   - Yazım önerisi alıntıdan çok farklı (ör. "Minimal efüzyon" → "Belirgin sıvı birikimi": anlam değişiyor)
5. **Önerilen metin (kod):** Yalnız "hekime sor" olmayan öneriler uygulanır ve ayrı bir dosyaya yazılır. Asıl rapor
   değişmez.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env  → kendi API anahtarınızı yazın
```

## Kullanım

```bash
python agent.py                                           # örnek (kurgusal) diz MR raporu
python agent.py --hasta hasta.txt --rapor dikte.docx
```

## Çıktı

- `cikti/rapor_kontrol.xlsx`: Bulgular (kırmızı: hekime sorulacak, sarı: yazım önerisi).
- `cikti/rapor_kontrol.txt`: önerilen metin.

## Dikkat

- **Hekim onayı.** Rapor hekimin sorumluluğundadır. Öneriler hekim onayı olmadan rapora işlenmez. Tıbbi içerik,
  ölçüm, doz ve tarih düzeltmeleri yalnız hekim tarafından yapılır.
- **Sağlık verisi.** Rapor metni yapay zekâ sağlayıcısına gönderilir. Maskeleme yardımcıdır, garanti değildir. Sağlık
  verisi KVKK'ya göre özel nitelikli kişisel veridir. Kurumunuzun yazılı izni ve veri işleme sözleşmesi yoksa bu agent'ı
  çalıştırmayın ya da yerel model (Ollama) kullanın.
- **Örnek veri:** Hasta ve rapor kurgusaldır; TC kimlik no yaygın kullanılan test numarasıdır.

## Testler

Testler gerçek API çağırmaz. Kodun örnek rapordaki beş tutarsızlığı bulduğu doğrulanır: kayıtta 47 yaşında erkek hasta
için "45 yaşında kadın hasta", rapor tarihinden sonraki bir tarih, farklı protokol no ve sağ diz raporunda "Sol dizde".
Sahte model ilk çıktıda metinde olmayan bir alıntı yazar, "8 mm"yi "18 mm" yapar, "Minimal efüzyon"u "Belirgin sıvı
birikimi" yapar ve bir bulguyu iki kez yazar. Kodun dört sorunu da yakaladığı; hasta adı, TC kimlik no, protokol no ve
doğum tarihinin modele gönderilmediği doğrulanır.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. API kullanım ücretleri ve raporun doğruluğu kullanıcının sorumluluğundadır. Bkz. `LICENSE`
ve `SORUMLULUK_REDDI.md`.
