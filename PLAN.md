# Plan ve Sıradaki İşler (PLAN)

> Yapılanlar için [PROGRESS.md](PROGRESS.md). Bu dosya "ne yapacağız ve nasıl" sorusunun cevabıdır.

## Hedef

Sitedeki **Sektör → Departman → Rol → Görev** ağacını Türkiye'deki gerçek şirket yapılarına ve
gerçek unvanlara göre doldurmak; ardından görevleri **gerçekten işe yarayan** kod blokları ve
agent'larla hayata geçirmek. Hız değil doğruluk ve kullanılabilirlik önceliklidir.

## Araştırma protokolü (uydurma yok)

| Ne | Kaynak | Not |
|---|---|---|
| Departman adları | Türk şirketlerinin faaliyet raporları / "organizasyon yapısı" sayfaları (ör. İş Bankası, Ziraat, sigorta ve sanayi şirketleri) | Her sektör için en az 2 şirket |
| Rol (unvan) adları | kariyer.net pozisyon sayfaları ("Toplam Kullanıcı", "Benzer Meslekler") | Uygulama tarayıcısıyla okunur (WebFetch 403 alıyor) |
| Görev tanımları | Gerçek iş ilanlarındaki "İş Tanımı" + MYK Ulusal Meslek Standartları | Paket yazılan her görev için ilan/standart kontrolü |
| Yasal parametreler (bordro, kıdem, vergi) | Resmî kaynaklar (GİB, SGK, ÇSGB, Resmî Gazete) | `parametreler.json` içinde kaynak + tarih ile |

Kaynaklar `catalog/KAYNAKLAR.md` dosyasına işlenir.

## Aşamalar

### Aşama 1 — Katalog (ağaç) yeniden kurulumu
- [ ] Sektör listesi (Türkiye'de istihdamı/şirket sayısı yüksek sektörler)
- [ ] Her sektör için departmanlar (gerçek adlandırma)
- [ ] Her departman için roller (kariyer.net'te doğrulanmış unvanlar)
- [ ] Her rol için görevler (ilan/MYK'dan)
- [ ] Ortak departmanlar: İK, Muhasebe, Finans, BT, Satın Alma, Hukuk, İdari İşler, Pazarlama, Satış, Müşteri Hizmetleri, İSG, Dış Ticaret
- [ ] catalog.mjs formatını genişlet: görev başına `tur` önerisi (kod/agent/ikisi), kaynak notu

Planlanan sektörler (sıra = öncelik):
1. Bankacılık 2. Sigortacılık 3. Üretim / Sanayi 4. Tekstil ve Hazır Giyim 5. Perakende
6. E-ticaret 7. Sağlık (özel hastane) 8. Lojistik ve Kargo 9. İnşaat 10. Gıda Üretimi
11. Turizm ve Otelcilik 12. Mali Müşavirlik (SMMM bürosu) 13. Hukuk Bürosu 14. Otomotiv Bayi ve Servis
15. Enerji 16. Eğitim (özel okul) 17. Faktoring / Leasing

### Aşama 2 — Paketler (öncelik: Türkiye'ye özgü + yüksek fayda + doğrulanabilir)
Kod blokları (deterministik):
- [ ] e-Fatura / e-Arşiv UBL-TR XML okuyucu → Excel (KDV kırılımı, tevkifat)
- [ ] TCKN / VKN / IBAN toplu doğrulama (cari ve personel listesi temizliği)
- [ ] Brüt-net bordro hesaplama (2026 parametreleri, kaynaklı)
- [ ] Kıdem ve ihbar tazminatı hesaplama
- [ ] Yıllık izin hakedişi hesaplama (4857 s. Kanun md. 53)
- [ ] Cari hesap yaşlandırma (alacak/borç)
- [ ] Form Ba/Bs mutabakat kontrolü
- [ ] Nakit akış tahmini (vadeli alacak/borç listesinden 13 hafta)
- [ ] Bütçe–gerçekleşen sapma raporu
- [ ] Stok ABC/XYZ analizi, yeniden sipariş noktası, stok yaşlandırma
- [ ] Satın alma teklif karşılaştırma (ağırlıklı puanlama)
- [ ] OEE hesaplama, SPC (X̄-R, Cp/Cpk)
- [ ] Tekstil: ölçü tablosu grading, 4 puan kumaş kontrol raporu
- [ ] İnşaat: hakediş hesaplama (kesintiler: KDV, tevkifat, teminat, stopaj)
- [ ] E-ticaret: pazaryeri sipariş kârlılık hesabı
- [ ] Otel: doluluk / ADR / RevPAR raporu
- [ ] Mali tablo rasyo analizi (kredi analisti)
Agent'lar (muhakeme gereken):
- [ ] Müşteri talebi/e-posta sınıflandırma + cevap taslağı
- [ ] Sözleşme ön inceleme (risk maddeleri)
- [ ] Mali tablo yorum raporu (rasyo koduna dayanır)
- [ ] Toplantı/görüşme notu → aksiyon listesi
- [ ] Hasar dosyası özeti
- [ ] Ürün açıklaması üretici
- [ ] Müşteri yorum analizi
- [ ] İş ilanı + mülakat soru seti hazırlayıcı
- [ ] 8D rapor taslağı
- [ ] PDF fatura okuma (e-Fatura olmayan faturalar)

### Aşama 3 — Yayın
- [ ] GitHub deposu (API 500 hatası; tekrar dene veya kullanıcıdan boş depo iste)
- [ ] GitHub Pages
- [ ] Agent'ların gerçek API ile testi (kullanıcı anahtar verecek)

## Çalışma kuralları
- Her paket STANDART.md'ye uyar ve `scripts/kontrol.py --test` geçmeden işaretlenmez.
- Yasal hesaplamalarda parametreler kodun içine gömülmez; `parametreler.json` + kaynak + yürürlük tarihi.
- Her önemli adımdan sonra PROGRESS.md güncellenir.
