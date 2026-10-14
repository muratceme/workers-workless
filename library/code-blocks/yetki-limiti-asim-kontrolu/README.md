# Yetki Limiti Aşım Kontrolü · Kod Bloğu

> Teknik (Risk Kabul / Underwriting) › Teknik Müdür · Workers / Workless

Onaylanan sigorta tekliflerini **teknik personel yetki matrisiyle** karşılaştırır: branş bazında azami sigorta bedeli ve azami
indirim aşılmış mı, onaylayanın o branşta yetkisi var mı, bir risk limitin altında kalmak için parçalara bölünmüş olabilir mi.
İnternete ve poliçe sistemine bağlanmaz.

## Nasıl kontrol eder?

| Kontrol | Kural |
|---|---|
| Yetki satırı | Onaylayan + branş; branşa özel satır yoksa onaylayanın "Tümü" (boş branş) satırı |
| Bedel aşımı | Sigorta bedeli > azami bedel |
| İndirim aşımı | Uygulanan indirim % > azami indirim % |
| Limite yakın | Bedel ya da indirim limitin `--yakin` (%90) ve üstünde |
| Yetki yok | Onaylayanın o branş için (ve "Tümü") satırı yok |
| Para birimi | Teklif ile yetki limitinin para birimi farklıysa "kontrol edilemedi" (çevrilmez) |
| Bölme şüphesi | Aynı sigortalı + branş, aynı onaylayan, `--gun` (30) gün içindeki tekliflerin bedel toplamı limiti aşıyor (tek tek aşmıyor) |

**Uyarılar:** bedel ya da indirim limiti aşıldı · yetki tanımı yok · okunamayan satır (Yüksek) · bölme şüphesi · para birimi farklı (Orta).

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                          # örnek veri (temmuz–ekim 2026)
python main.py --yetkiler yetkiler.xlsx --teklifler teklifler.xlsx
python main.py --yetkiler yetkiler.xlsx --teklifler teklifler.xlsx --yakin 80 --gun 60
```

## Çıktı

`cikti/yetki_limiti_kontrolu.xlsx`: `Teklif Kontrolü` (aşımlar üstte), `Personel` (onaylayan bazında aşım ve yetkisiz onay sayısı),
`Bölme Şüphesi`, `Uyarılar`.

## Dikkat

- **Yetki matrisi şirket içi düzenlemedir.** Matrisi yönetim kurulu / genel müdürlükçe onaylanan güncel yetki tablosundan
  aktarın. Limitler kişi yerine unvan bazındaysa her personel için satır açın.
- **Bedel tanımı branşa göre değişir** (araç bedeli, yeniden yapım bedeli, sefer başı limit, PML, EML). Yetki tablosundaki tanımla aynı
  bedeli kullanın; reasürans ve koasürans payları hesaba girmez.
- **Bölme şüphesi bir göstergedir;** farklı lokasyon ya da farklı riskler gerçekten ayrı olabilir. Dosya incelemesiyle doğrulayın.
- **Örnek veri:** Personel, sigortalılar, limitler ve tutarlar kurgusaldır.

## Testler

- Branşa özel ve "Tümü" yetkisi, bedel ve indirim aşımı, limite yakınlık, yetkisiz onay, farklı para birimi, bölme penceresi.
- Sıralama, personel özeti, okuma hataları ve komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
