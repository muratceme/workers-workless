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
- [x] Sektör listesi (Türkiye'de istihdamı/şirket sayısı yüksek sektörler)
- [x] Her sektör için departmanlar (gerçek adlandırma)
- [x] Her departman için roller (kariyer.net'te doğrulanmış unvanlar)
- [x] Her rol için görevler (ilan/MYK'dan)
- [x] Ortak departmanlar: İK, Muhasebe, Finans, BT, Satın Alma, Hukuk, İdari İşler, Pazarlama, Satış, Müşteri Hizmetleri, İSG, Dış Ticaret
- [x] Görev kayıt defteri (aynı görev birden çok rolde)
- [ ] (Sonra) Hukuk bürosu, otomotiv, eğitim için daha çok rol; SOON sektörleri (Enerji, Telekom, Faktoring, Kimya-İlaç)

Planlanan sektörler (sıra = öncelik):
1. Bankacılık 2. Sigortacılık 3. Üretim / Sanayi 4. Tekstil ve Hazır Giyim 5. Perakende
6. E-ticaret 7. Sağlık (özel hastane) 8. Lojistik ve Kargo 9. İnşaat 10. Gıda Üretimi
11. Turizm ve Otelcilik 12. Mali Müşavirlik (SMMM bürosu) 13. Hukuk Bürosu 14. Otomotiv Bayi ve Servis
15. Enerji 16. Eğitim (özel okul) 17. Faktoring / Leasing

### Aşama 2 — Paketler
Aşama 2'nin ilk listesi TAMAMLANDI (2026-10-07). Sıradaki: stok ABC/XYZ + yeniden sipariş noktası → otel doluluk/ADR/RevPAR
→ tekstil 4 puan kumaş kontrolü → ölçü tablosu grading → agent'lar: toplantı notu→aksiyon, müşteri yorum analizi, iş ilanı+mülakat,
8D rapor taslağı, mali tablo yorum raporu, PDF fatura okuma.
Eski sıra: Fine-Kinney → AQL → OEE → SPC → Erlang C → alacak yaşlandırma → nakit akış →
cari mutabakat → Ba-Bs → bütçe sapma → hakediş (inşaat) → pazaryeri kârlılık → teklif karşılaştırma → agent'lar
 (öncelik: Türkiye'ye özgü + yüksek fayda + doğrulanabilir)
Kod blokları (deterministik):
- [x] e-Fatura / e-Arşiv UBL-TR XML okuyucu → Excel (KDV kırılımı, tevkifat)
- [x] TCKN / VKN / IBAN toplu doğrulama (cari ve personel listesi temizliği)
- [x] Brüt-net bordro hesaplama (2026 parametreleri, kaynaklı)
- [x] Kıdem ve ihbar tazminatı hesaplama
- [x] Yıllık izin hakedişi hesaplama (4857 s. Kanun md. 53)
- [x] Cari hesap yaşlandırma (alacak/borç)
- [~] ~~Form Ba/Bs~~ — 565 Sıra No'lu VUK Tebliği ile 01.10.2024'ten itibaren KALDIRILDI; yerine "Gelen e-Fatura Kayıt Kontrolü"
- [x] Nakit akış tahmini (vadeli alacak/borç listesinden 13 hafta)
- [x] Bütçe–gerçekleşen sapma raporu
- [x] Cari hesap mutabakatı (karşı taraf ekstresiyle)
- [ ] Stok ABC/XYZ analizi, yeniden sipariş noktası, stok yaşlandırma
- [x] Satın alma teklif karşılaştırma (ağırlıklı puanlama)
- [x] OEE hesaplama, SPC (X̄-R, Cp/Cpk)
- [ ] Tekstil: ölçü tablosu grading, 4 puan kumaş kontrol raporu
- [x] İnşaat: hakediş hesaplama (kesintiler: KDV, tevkifat, teminat, stopaj)
- [x] E-ticaret: pazaryeri sipariş kârlılık hesabı
- [ ] Otel: doluluk / ADR / RevPAR raporu
- [x] Mali tablo rasyo analizi (kredi analisti)
Agent'lar (muhakeme gereken):
- [x] Müşteri talebi/e-posta sınıflandırma + cevap taslağı
- [x] Sözleşme ön inceleme (risk maddeleri)
- [ ] Mali tablo yorum raporu (rasyo koduna dayanır)
- [ ] Toplantı/görüşme notu → aksiyon listesi
- [x] Hasar dosyası özeti
- [x] Ürün açıklaması üretici
- [ ] Müşteri yorum analizi
- [ ] İş ilanı + mülakat soru seti hazırlayıcı
- [ ] 8D rapor taslağı
- [ ] PDF fatura okuma (e-Fatura olmayan faturalar)

### Aşama 3 — Yayın
- [x] GitHub deposu
- [x] GitHub Pages
- [ ] Agent'ların gerçek API ile testi (kullanıcı anahtar verecek)

## Çalışma kuralları
- Her paket STANDART.md'ye uyar ve `scripts/kontrol.py --test` geçmeden işaretlenmez.
- Yasal hesaplamalarda parametreler kodun içine gömülmez; `parametreler.json` + kaynak + yürürlük tarihi.
- Her önemli adımdan sonra PROGRESS.md güncellenir.
