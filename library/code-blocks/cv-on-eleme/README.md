# CV Ön Eleme · Kod Bloğu

> İnsan Kaynakları › İnsan Kaynakları Uzman Yardımcısı · Workers / Workless

Gelen CV'leri ilanın **zorunlu kriterlerine** göre ilk filtreden geçirir. Her aday için **GEÇTİ / ELENDİ /
MANUEL KONTROL** kararını kriter bazında gerekçesiyle listeler. İnternete bağlanmaz.

## Kriterler (`kriterler.json`)

| Alan | Örnek | Anlamı |
|---|---|---|
| `en_az_deneyim_yil` | `5` | CV'deki tarih aralıklarından (çakışanlar birleştirilerek) hesaplanan toplam yıl |
| `en_az_egitim` | `"Lisans"` | Lise < Ön Lisans < Lisans < Yüksek Lisans < Doktora |
| `zorunlu_diller` | `["İngilizce"]` | Hepsi gerekli |
| `zorunlu_beceriler` | `["Excel", "Bordro"]` | Hepsi gerekli (Türkçe karakter ve büyük/küçük harf duyarsız) |
| `beceri_gruplari` | `[{"liste": ["SAP","Logo"], "en_az": 1}]` | Listeden en az N tanesi |
| `zorunlu_ifadeler` | `[{"ifade": "SMMM", "gerekce": "..."}]` | Belge, sertifika vb. |

## Adil ve temkinli karar

- **Belirsiz bilgi elemez:** CV'den okunamayan bilgi (taranmış PDF, tarihsiz deneyim, eğitim yok) adayı
  **elemez**; aday **MANUEL KONTROL**e düşer.
- **Ayrımcılık denetimi:** Kriterler ortak tarayıcıdan geçer. Yaş, cinsiyet, medeni hâl, görünüş, sağlık, din,
  köken gibi kriterler verilirse paket **çalışmaz**. Ehliyet, askerlik, uyruk gibi "dikkat" kriterleri yalnız
  `gerekce` alanıyla kabul edilir (4857 sayılı Kanun md. 5; 6701 sayılı Kanun md. 3, 6, 7).
- **Kişisel bilgiler:** Ad, fotoğraf ve iletişim bilgisi karara girmez; rapordaki iletişim için gösterilir.

CV okuma çekirdeği **CV Raporlama** kod bloğuyla ortaktır (`cv_cekirdek.py`).

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                                  # örnek 4 CV + kriterler_ornek.json
python main.py --girdi ./cvler --kriterler kriterler.json
```

> Kelime ve tarih tabanlı okuma yapar. Eş anlamlı ifadeler ve tablo/görsel ağırlıklı CV'ler kaçabilir.
> Elenen adaylara dönüş yapmadan önce örneklem kontrolü yapın.

## Çıktı

`Ön Eleme` (karar, gerekçe, kriter sütunları, "İK Onayı") · `Kriterler` · `Bilgi`

## Testler

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. CV'ler kişisel veridir; işleme ve saklama KVKK'ya uygun yapılmalıdır.
Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
