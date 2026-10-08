# Taşeron Hakediş Kontrolü · Kod Bloğu

> İnşaat › Teknik Ofis › Teknik Ofis Mühendisi · Workers / Workless

Taşeronun sunduğu hakedişi **sözleşme birim fiyat cetveli**, **saha (yeşil defter) metrajı** ve **önceki onaylı
hakedişle** karşılaştırır. Farkları tutar etkisiyle listeler ve onaylanabilir bu dönem tutarını hesaplar.
İnternete bağlanmaz.

## Kontroller

| Seviye | Kontrol |
|---|---|
| Hata | **Aritmetik:** önceki + bu dönem ≠ toplam |
| Hata | **Tutar hesabı:** miktar × birim fiyat ≠ tutar |
| Hata | **Birim:** birim sözleşmeden farklı |
| Hata | **Mükerrer poz** |
| Hata | **Önceki miktar:** taşeronun "önceki" miktarı onaylı önceki kümülatiften farklı |
| Yüksek | **Birim fiyat:** sözleşme fiyatından farklı |
| Yüksek | **Sözleşmede yok:** poz sözleşmede yok (ek iş / yeni birim fiyat onayı gerekir) |
| Yüksek | **Saha metrajı aşımı:** toplam miktar saha metrajını (`--tolerans` dahil) aşıyor |
| Dikkat | **Sözleşme miktarı aşımı** |
| Dikkat | **Negatif bu dönem miktarı** |
| Dikkat | **Saha metrajı yok:** poz için ölçüm yok |
| Bilgi | **Hakedişte yok:** sahada ölçülmüş ama talep edilmemiş poz |

Poz numaraları yazım farkına rağmen eşleşir: `15.150.1005`, `15 150 1005` ve `15/150/1005` aynı kabul edilir.

## Onaylanabilir tutar

```
onaylanan kümülatif = min(taşeron toplamı, saha metrajı, onaylı önceki + talep edilen bu dönem [, sözleşme miktarı])
onaylanan bu dönem  = onaylanan kümülatif − onaylı önceki
tutar               = onaylanan bu dönem × sözleşme birim fiyatı
```

- **Fiyat farkı:** Taşeron daha yüksek fiyat yazmışsa sözleşme fiyatı esas alınır.
- **Sözleşmede olmayan pozlar:** Yeni fiyat onaylanana kadar ödemeye alınmaz.
- **Saha metrajı olmayan pozlar:** Ölçülene kadar onaylanmaz.
- **Sözleşme miktarı sınırı:** `--sozlesme-siniri` verilirse kümülatif, sözleşme miktarını da aşamaz.

Hesap **vergi ve kesinti öncesidir**. KDV, KDV tevkifatı, stopaj, teminat ve avans kesintileri için
**Hakediş Hesaplama** paketini kullanın.

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                   # örnek: 5 no'lu hakediş, 6 poz
python main.py --sozlesme birim_fiyatlar.xlsx --hakedis taseron_hakedis_5.xlsx --saha saha_metraji.xlsx --onceki onayli_hakedis_4.xlsx
python main.py ... --tolerans 1 --sozlesme-siniri
```

| Dosya | Sütunlar |
|---|---|
| Sözleşme | Poz No, Tanım, Birim, Birim Fiyat, [Sözleşme Miktarı] |
| Taşeron hakedişi | Poz No, Tanım, Birim, Birim Fiyat, Önceki Miktar, Bu Dönem Miktar, Toplam Miktar, Toplam Tutar |
| Saha metrajı | Poz No, Kümülatif Miktar (aynı poz birden fazla satırdaysa toplanır) |
| Önceki onaylı | Poz No, Onaylanan Kümülatif Miktar |

Başlığın üstündeki firma/hakediş adı satırları ve alttaki "Toplam" satırı atlanır.

## Çıktı

- **Özet:** talep, onaylanabilir, fark ve seviyelere göre bulgu sayısı.
- **Bulgular:** taşeron açıklaması ve karar sütunlarıyla.
- **Poz Karşılaştırma:** talep ve onay, miktar ve tutar olarak yan yana.

## Dikkat

- **Ön kontrol:** Bu araç bir ön kontroldür. Kesin hakediş, sözleşme şartları (fiyat farkı, ek iş protokolleri,
  iş artışı sınırları) ve şantiye şefi / kontrol mühendisi onayıyla belirlenir.
- **Saha metrajı:** Saha metrajının güncel ve imzalı olduğundan emin olun.

## Testler

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
