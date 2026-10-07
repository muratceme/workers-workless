/* İnşaat (müteahhitlik) — unvanlar Dafne Mühendislik, Akçadağ İnşaat, İnart Enerji, Aşçıoğlu,
   Ceyen İnşaat, Fırat Zirve ilanlarına dayanır (teknik ofis, hakediş, metraj, şantiye). */

export const INSAAT = {
  "insaat-teknik-ofis": {
    name: "Teknik Ofis", short: "TOF",
    desc: "Metraj, keşif, hakediş, maliyet ve taşeron sözleşmeleri.",
    roles: [
      { id: "teknik-ofis-muhendisi", name: "Teknik Ofis Mühendisi", level: "Uzman", tasks: [
        ["Hakediş Hesaplama", "İmalat metrajları ve birim fiyatlardan kümülatif ve dönem hakedişini; KDV, KDV tevkifatı, gelir/kurumlar vergisi stopajı, teminat ve avans kesintileriyle hesaplar.", 6],
        ["Taşeron Hakediş Kontrolü", "Taşeronun sunduğu hakedişi sözleşme birim fiyatları, saha metrajları ve önceki hakedişlerle karşılaştırıp farkları listeler.", 4],
        ["Metraj Cetveli Kontrolü", "Metraj cetvelindeki hesap satırlarını (boyut × adet) yeniden hesaplar, poz toplamlarını ve mükerrer satırları kontrol eder.", 3],
        ["Birim Fiyat Analizi", "Malzeme, işçilik ve makine rayiçlerinden poz bazında birim fiyat analizi ve kâr-genel gider dahil fiyat oluşturur.", 3],
      ]},
      { id: "teknik-ofis-sefi", name: "Teknik Ofis Şefi", level: "Yönetici", tasks: [
        ["Proje Maliyet Kontrol Raporu", "Bütçe, gerçekleşen maliyet ve hakediş gelirlerinden iş kalemi bazında maliyet sapması ve tamamlanma tahmini üretir.", 3],
        ["Fiyat Farkı Hesaplama", "Sözleşme katsayılarına (a, b1, b2, c...) ve TÜİK endekslerine göre dönem fiyat farkı katsayısını ve tutarını hesaplar.", 3],
      ]},
    ],
  },

  "insaat-santiye": {
    name: "Şantiye", short: "ŞNT",
    desc: "Saha yönetimi, imalat takibi ve günlük raporlama.",
    roles: [
      { id: "saha-muhendisi", name: "Saha Mühendisi", level: "Uzman", tasks: [
        ["Günlük Şantiye Raporu", "Saha notlarından hava durumu, personel-makine sayısı, yapılan imalatlar ve sorunları içeren günlük raporu hazırlar.", 3],
        ["Beton Numune ve Kırım Takibi", "Beton döküm kayıtları ve 7/28 günlük numune kırım sonuçlarını takip edip sınıf dayanımının altında kalanları işaretler.", 2],
        ["Şantiye Puantajı", "Taşeron ve kendi personelinin günlük puantajını ekip ve imalat bazında toplar.", 2],
      ]},
      { id: "santiye-sefi", name: "Şantiye Şefi", level: "Yönetici", tasks: [
        ["İş İlerleme Raporu", "İş programı ve gerçekleşen imalat yüzdelerinden S-eğrisi, gecikme ve kritik iş kalemleri raporunu üretir.", 3],
        ["Malzeme Talep ve Sevk Takibi", "Şantiye malzeme taleplerini satın alma ve sevkiyat durumuyla takip eder, imalatı geciktirecek eksikleri işaretler.", 2],
      ]},
    ],
  },

  "insaat-planlama": {
    name: "Planlama ve Proje Kontrol", short: "PLN",
    desc: "İş programı, kaynak planı ve ilerleme ölçümü.",
    roles: [
      { id: "planlama-muhendisi", name: "Planlama Mühendisi", level: "Uzman", tasks: [
        ["İş Programı Kritik Yol Analizi", "Faaliyet süreleri ve bağımlılıklardan kritik yolu, bolluk sürelerini ve proje bitiş tarihini hesaplar.", 3],
        ["Kazanılmış Değer Analizi", "Planlanan değer, kazanılmış değer ve gerçekleşen maliyetten SPI, CPI ve tamamlanma maliyet tahminini hesaplar.", 3],
      ]},
    ],
  },

  "insaat-ihale": {
    name: "İhale ve Teklif", short: "İHL",
    desc: "İhale dokümanı inceleme, keşif ve teklif hazırlığı.",
    roles: [
      { id: "ihale-muhendisi", name: "İhale Mühendisi", level: "Uzman", tasks: [
        ["İhale Dokümanı Özeti", "İdari ve teknik şartnamelerden yeterlik kriterleri, teminat, süre, ceza ve özel şartları özetler.", 4],
        ["Teklif Birim Fiyat Karşılaştırması", "Taşeron ve tedarikçi tekliflerini poz bazında karşılaştırıp en uygun fiyatlarla teklif cetveli oluşturur.", 3],
      ]},
    ],
  },

  "insaat-kalite": {
    name: "Kalite Kontrol", short: "KLT",
    desc: "İmalat kontrolleri, test ve uygunsuzluklar.",
    roles: [
      { id: "kalite-kontrol-muhendisi-insaat", name: "Kalite Kontrol Mühendisi", level: "Uzman", tasks: [
        ["İmalat Kontrol Formu Takibi", "İmalat kontrol talepleri (ITP) ve onay durumlarını kat, blok ve imalat türü bazında takip eder.", 2],
        "DÖF Takibi",
      ]},
    ],
  },
};
