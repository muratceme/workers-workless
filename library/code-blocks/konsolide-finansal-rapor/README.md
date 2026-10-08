# Konsolide Finansal Rapor · Kod Bloğu

> Finans › Finans Müdürü · Workers / Workless

Grup şirketlerinin Tekdüzen mizanlarını birleştirir, grup içi bakiye ve işlemleri eler ve **konsolide bilanço ile
gelir tablosu** üretir (tam konsolidasyon). Bilanço ve gelir tablosunda şirketler ve konsolide tutarlar yan yana
gösterilir. İnternete bağlanmaz.

## Eliminasyonlar

| Tür | Ne yapar | Kayıt |
|---|---|---|
| **Cari** | Bir şirketin grup içi alacağını (ör. 120.02) diğerinin borcuyla (ör. 320.02) karşılaştırır. | Mutabık tutar elenir. Fark **Hata** olarak raporlanır ve ilgili hesapta kalır. |
| **Gelir-Gider** | Grup içi mal satışı, hizmet, faiz ve kira gibi işlemleri eler. | Dr satıcının gelir hesabı (600.02) / Cr alıcının gider hesabı (621, 631…) |
| **Stoktaki Kâr** | Grup içi alınıp yıl sonunda stokta kalan malların içindeki gerçekleşmemiş kârı eler. | Dr 621 / Cr 153 |
| **Düzeltme** | Elle girilen konsolidasyon kaydıdır. | Dr Hesap / Cr Karşı Hesap |
| **Sermaye** (otomatik) | Ana ortaklıktaki yatırımı (245.xx), bağlı ortaklığın özkaynağıyla (50–58) eler. | Ayrıntılar aşağıda. |

**Sermaye eliminasyonu**
- **Şerefiye (261):** yatırım − pay × edinim tarihindeki özkaynak.
- **Edinim sonrası yedekler:** Grup payı 570 Geçmiş Yıllar Kârları hesabına yazılır.
- **Kontrol gücü olmayan paylar:** (1 − pay) × bağlı ortaklık özkaynağı. Özkaynakta ayrı grup olarak gösterilir.
  Dönem kârındaki payları 59 grubunda indirim satırıdır.
- **Negatif şerefiye:** Dikkat olarak raporlanır ve 570'e yazılır; gözden geçirilmelidir.
- **Pay %50 veya altı:** Kontrol yoksa tam konsolidasyon değil özkaynak yöntemi uygulanır; uyarı verilir.

**Ortak çekirdek:** Mali tablo düzeni (MSUGT ayrıntılı bilanço ve gelir tablosu) ve ters bakiye virmanları
"Mizandan Mali Tablo Hazırlama" paketiyle aynıdır.

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                        # örnek: Holding + Pazarlama (%100) + Lojistik (%80)
python main.py --sirketler sirketler.xlsx --eliminasyonlar eliminasyonlar.xlsx --unvan "Örnek Grup" --tolerans 1
```

**`sirketler`**

| Sütun | İçerik |
|---|---|
| Kod | Kısa şirket kodu (ANA, BP…) |
| Ünvan | Şirket ünvanı |
| Mizan | Mizan dosyasının yolu, `sirketler` dosyasına göre |
| Pay % | Ana ortaklığın payı |
| Yatırım Hesabı | Ana ortaklıktaki alt hesap (245.01) |
| Edinim Özkaynağı | Edinim tarihindeki özkaynak |
| Rol | Ana / Bağlı |

**`eliminasyonlar`**

| Sütun | İçerik |
|---|---|
| Tür | Cari / Gelir-Gider / Stoktaki Kâr / Düzeltme |
| Şirket, Hesap Kodu | Kaydın bir tarafı |
| Karşı Şirket, Karşı Hesap Kodu | Kaydın diğer tarafı |
| Tutar | Cari türünde boş bırakılırsa tutar mizandan okunur |
| Açıklama | Serbest metin |

**Mizanlar:** Kapanış öncesi mizan kullanın; gelir tablosu hesapları açık olmalıdır. Grup içi işlemleri ayrı alt
hesaplarda (120.02, 600.02 gibi) izlemek eliminasyonu kolaylaştırır.

## Çıktı

- `Bilgi`: kâr dağılımı, şerefiye, kontrol gücü olmayan paylar ve yöntem notu.
- `Bulgular`: mutabakat farkları ve uyarılar.
- `Bilanço Aktif`, `Bilanço Pasif` ve `Gelir Tablosu`: şirketler ve konsolide tutarlar yan yana.
- `Virmanlar`.
- `Eliminasyon Kayıtları`: borç/alacak yevmiye biçiminde.
- `Çalışma Tablosu`: hesap × şirket, toplam, eliminasyon borç/alacak ve konsolide tutar.

## Dikkat — kapsam

Bu araç yönetim raporlaması için **basitleştirilmiş** bir konsolidasyon yapar:

- **Hesaplanmayanlar:** Edinimdeki gerçeğe uygun değer düzeltmeleri, şerefiye değer düşüklüğü testi, eliminasyonların
  ertelenmiş vergi etkisi, yabancı para ile tutulan bağlı ortaklıkların çevrimi, iştirakler ve iş ortaklıkları için
  özkaynak yöntemi, dönem içinde edinim ve elden çıkarma, karşılıklı ve dolaylı ortaklık yapıları.
- **Kontrol gücü olmayan paylar:** Bağlı ortaklığın kendi mizanındaki kâr üzerinden hesaplanır. Grup içi kârın
  "yukarı akış" (bağlı ortaklıktan ana ortaklığa satış) etkisi dağıtılmaz.
- **Yasal raporlama:** TFRS veya BOBİ FRS'ye göre hazırlanan yasal konsolide finansal tabloların ve bağımsız
  denetimin yerine geçmez. Konsolidasyon yükümlülüğünü ve uygulanacak çerçeveyi mali müşavirinizle / denetçinizle
  belirleyin.
- **Doğrulama:** Örnek veriler kurgusaldır ve testlerde elle hesaplanan sonuçlarla doğrulanır.

## Testler

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
