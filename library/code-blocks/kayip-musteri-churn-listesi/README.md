# Kayıp Müşteri (Churn) Listesi · Kod Bloğu

> CRM ve Sadakat Programı › CRM Uzmanı · Workers / Workless

Sadakat programı üyelerinin alışveriş geçmişinden **alışveriş sıklığı düşen** üyeleri bulur; değer segmentine göre
önceliklendirir ve **ileti izni olan** üyelerle geri kazanım kampanyası listesi çıkarır. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Alışveriş sayısı | Farklı alışveriş günü sayısı (aynı gün birden çok fiş tek ziyaret) |
| Ortalama aralık | (Son alışveriş − ilk alışveriş) ÷ (alışveriş sayısı − 1) gün |
| Son alışverişten beri | Analiz tarihi − son alışveriş (gün) |

| Durum | Kural |
|---|---|
| Kayıp | `--kayip` (varsayılan 180) günden uzun süredir alışveriş yok |
| Risk altında | En az iki alışverişi var; son alışverişten beri geçen süre ortalama aralığın `--kat` (varsayılan 2) katından ve `--min` (varsayılan 30) günden uzun |
| Tek alışveriş | Tek alışverişi var, üzerinden `--tek` (varsayılan 90) günden fazla geçmiş |
| Aktif | Diğerleri |

Kural üyenin **kendi alışkanlığına** göre çalışır: 10 günde bir gelen üye 30 gündür gelmiyorsa risk altındadır; ayda
bir gelen üye için 30 gün normaldir.

**Değer segmenti:** toplam harcama sıralamasında ilk %20 Yüksek, sonraki %30 Orta, kalan Düşük.

**Kampanya listesi:** Risk altında, Tek alışveriş ve Kayıp üyelerden ileti izni "Evet" olanlar; sıra: önce risk
altındakiler (geri kazanması en kolay), sonra tek alışverişliler, sonra kayıplar; her grupta harcamaya göre.

**Örnek:** U1009 ortalama 10,1 günde bir alışveriş yapıyor; son alışverişinden beri 96 gün geçmiş → 96 > 20,2 →
**risk altında**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                                       # örnek veri (30.09.2026)
python main.py --alisverisler uye_alisverisleri.xlsx --uyeler uyeler.xlsx --tarih 30.09.2026
python main.py --alisverisler uye_alisverisleri.xlsx --uyeler uyeler.xlsx --kayip 120 --kat 1,5 --tek 60
```

Zorunlu sütunlar: alışverişlerde Üye No, Tarih, Tutar; üyelerde Üye No ve İleti İzni. Üye dosyası verilmezse
kampanya listesi boş kalır (izin bilinmeden ileti gönderilemez).

## Çıktı

`cikti/kayip_musteri_listesi.xlsx`:

- `Kampanya Listesi`: izinli hedef üyeler.
- `Tüm Hedef Üyeler`: izinsizler dahil; izni olmayanlar sarı.
- `Tüm Üyeler`: alışveriş sayısı, ortalama aralık, geçen gün, harcama, değer segmenti, durum.
- `Özet`: duruma göre üye sayısı ve harcama.
- `Uyarılar`.

## Dikkat

- **Ticari elektronik ileti izni zorunludur.** 6563 sayılı Kanun ve ilgili yönetmelik uyarınca SMS, e-posta ve arama
  ile ticari ileti için alıcının önceden onayı gerekir ve izinler İleti Yönetim Sistemi'ne (İYS) kayıtlı olmalıdır.
  Listeyi göndermeden önce izinleri İYS'den doğrulayın; kod İYS'ye bağlanmaz.
- **Kişisel veri işlenir.** Alışveriş geçmişiyle segmentasyon KVKK kapsamında kişisel veri işlemedir; amacın ve hukuki
  sebebin (gerekiyorsa açık rıza) aydınlatma metninizle uyumlu olduğunu doğrulayın, dosyaya erişimi sınırlayın.
- **Kural tabanlıdır, tahmin modeli değildir.** Mevsimlik alışveriş yapan, taşınan ya da kartını kullanmadan alışveriş
  yapan üyeler yanlış sınıflanabilir. Eşikler politika varsayılanıdır; sektörünüzün alışveriş sıklığına göre
  değiştirin.
- **Örnek veri:** Üye numaraları, adlar ve alışverişler kurgusaldır.

## Testler

- Farklı gün sayımı, ortalama aralık, geçen gün; Kayıp, Risk altında, Tek alışveriş ve Aktif ayrımı.
- Değer segmenti, izin süzgeci ve liste sırası, analiz tarihinden sonraki kayıt, eşikler, okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Sonuçların doğruluğu ve kullanımı kullanıcının sorumluluğundadır. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
