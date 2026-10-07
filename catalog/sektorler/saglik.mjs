/* Sağlık (özel hastane) — unvanlar Memorial, Medicana, Medipol, Levent Hastanesi, Sentez Sağlık
   ilanlarına; birimler SKS (Sağlıkta Kalite Standartları) ve hastane görev tanımlarına dayanır. */

export const SAGLIK = {
  "saglik-hasta-hizmetleri": {
    name: "Hasta Hizmetleri", short: "HHZ",
    desc: "Hasta kabul, randevu, danışma ve hasta ilişkileri.",
    roles: [
      { id: "hasta-kayit-kabul-gorevlisi", name: "Hasta Kayıt Kabul Görevlisi", level: "Giriş", tasks: [
        ["Hasta Kabul Gün Sonu Kasa Kontrolü", "Gün içindeki tahsilat, fatura ve iadeleri HBYS kayıtlarıyla karşılaştırıp kasa devir raporunu hazırlar.", 2],
        ["Provizyon ve Sigorta Bilgisi Ön Kontrolü", "Randevulu hastaların SGK ve özel sağlık sigortası bilgilerini ön kontrol edip eksik provizyon bilgisini listeler.", 3],
        ["Randevu Hatırlatma Mesajları", "Ertesi günün randevularından bölüm, hekim ve hazırlık bilgisi içeren hatırlatma mesajları hazırlar.", 2],
      ]},
      { id: "hasta-iliskileri-yoneticisi", name: "Hasta İlişkileri ve Hasta Hakları Yöneticisi", level: "Yönetici", tasks: [
        ["Hasta Şikâyet Sınıflandırma", "Hasta şikâyet ve önerilerini birim, konu ve önem derecesine göre sınıflandırıp yasal cevap süresiyle takip eder.", 3],
        ["Hasta Memnuniyet Anketi Analizi", "Anket puan ve yorumlarından birim ve hekim bazında memnuniyet raporu ve iyileştirme başlıkları çıkarır.", 3],
      ]},
      { id: "randevu-planlama-sorumlusu", name: "Randevu Planlama Sorumlusu", level: "Uzman", tasks: [
        ["Randevu Doluluk ve Gelmeme Analizi", "Hekim ve bölüm bazında randevu doluluk, gelmeme (no-show) ve iptal oranlarını hesaplayıp fazla randevu (overbooking) önerisi çıkarır.", 3],
      ]},
    ],
  },

  "saglik-tibbi-sekreterlik": {
    name: "Tıbbi Sekreterlik", short: "TSK",
    desc: "Poliklinik ve servis sekreterliği, tıbbi dokümantasyon.",
    roles: [
      { id: "tibbi-sekreter", name: "Tıbbi Sekreter", level: "Giriş", tasks: [
        ["Epikriz Taslağı", "Hekim notları ve tetkik sonuçlarından epikriz taslağı oluşturur (hekim onayı zorunludur).", 5],
        ["Tıbbi Rapor Yazım Kontrolü", "Hekimin dikte ettiği rapor metnini yazım, tarih ve hasta bilgisi tutarlılığı açısından kontrol eder.", 2],
      ]},
      { id: "yatan-hasta-tibbi-sekreteri", name: "Yatan Hasta Tıbbi Sekreteri", level: "Giriş", tasks: [
        ["Taburcu Dosyası Eksik Kontrolü", "Taburcu olan hastaların dosyalarında onam, epikriz, ameliyat notu ve konsültasyon belgelerinin eksiksizliğini kontrol eder.", 3],
        ["Servis Yatak Doluluk Raporu", "Servis bazında yatak doluluk, ortalama yatış süresi ve beklenen taburcu listesini çıkarır.", 2],
      ]},
    ],
  },

  "saglik-anlasmali-kurumlar": {
    name: "Anlaşmalı Kurumlar ve Faturalama", short: "AKF",
    desc: "SGK (Medula) ve özel sağlık sigortası faturalama, kesinti ve itiraz yönetimi.",
    roles: [
      { id: "anlasmali-kurumlar-uzmani", name: "Anlaşmalı Kurumlar Uzmanı", level: "Uzman", tasks: [
        ["Özel Sağlık Sigortası Fatura Ön Kontrolü", "Fatura kalemlerini sigorta şirketi anlaşma fiyatları, provizyon onayı ve poliçe limitleriyle karşılaştırıp kesinti riskli kalemleri işaretler.", 5],
        ["Kesinti ve İtiraz Takibi", "Sigorta şirketlerinin yaptığı kesintileri neden bazında sınıflandırır, itiraz edilebilir olanları ve süreleri takip eder.", 3],
      ]},
      { id: "faturalama-uzmani-sgk", name: "SGK Faturalama Uzmanı", level: "Uzman", tasks: [
        ["SGK Fatura Ön Kontrolü", "Medula'ya gönderilecek hizmet kalemlerini yapılandırılabilir SUT kurallarına (tanı-işlem uyumu, adet sınırı, tekrar süresi) göre ön kontrolden geçirir.", 5],
        ["Medula Ret ve Kesinti Analizi", "Medula ret/kesinti kayıtlarını hata kodu, branş ve işlem bazında analiz edip tekrarlayan hataları raporlar.", 3],
      ]},
      { id: "anlasmali-kurumlar-muduru", name: "Anlaşmalı Kurumlar Müdürü", level: "Yönetici", tasks: [
        ["Kurum Bazlı Gelir ve Kesinti Raporu", "SGK ve özel sigorta şirketleri bazında fatura, tahsilat, kesinti oranı ve ortalama tahsil süresini raporlar.", 2],
      ]},
    ],
  },

  "saglik-kalite": {
    name: "Kalite Yönetimi", short: "KLT",
    desc: "SKS uyumu, kalite göstergeleri ve olay bildirimleri.",
    roles: [
      { id: "kalite-birim-sorumlusu", name: "Kalite Birim Sorumlusu", level: "Uzman", tasks: [
        ["Kalite Göstergeleri Hesaplama", "Kayıtlardan SKS kalite göstergelerini (düşme, bası yarası, enfeksiyon oranı vb.) formüllerine göre dönemsel hesaplar.", 3],
        ["Güvenlik Raporlama Sistemi Olay Analizi", "Hasta ve çalışan güvenliği olay bildirimlerini tür, birim ve önlenebilirlik açısından analiz eder.", 2],
        ["SKS Öz Değerlendirme Takibi", "SKS standartlarına göre öz değerlendirme puanlarını ve eksik dokümanları birim bazında takip eder.", 2],
      ]},
    ],
  },

  "saglik-uluslararasi": {
    name: "Uluslararası Hasta Hizmetleri", short: "UHH",
    desc: "Yabancı hasta koordinasyonu, teklif ve iletişim.",
    roles: [
      { id: "uluslararasi-hasta-koordinatoru", name: "Uluslararası Hasta Koordinatörü", level: "Uzman", tasks: [
        ["Tedavi Fiyat Teklifi Hazırlama", "Tedavi planı ve paket fiyat listesinden konaklama ve transfer dahil hastaya gönderilecek teklif dosyasını hazırlar.", 3],
        ["Hasta İletişim Mesajı Çevirisi", "Hasta ile yazışmaları tıbbi terimleri koruyarak çevirir ve kültüre uygun cevap taslağı hazırlar.", 4],
        ["Hasta Seyahat ve Tedavi Takvimi", "Uçuş, konaklama, muayene ve operasyon tarihlerinden hasta bazında takvim ve hatırlatmalar oluşturur.", 2],
      ]},
    ],
  },

  "saglik-eczane": {
    name: "Eczane ve Tıbbi Malzeme", short: "ECZ",
    desc: "İlaç ve sarf malzeme stokları, son kullanma tarihleri.",
    roles: [
      { id: "hastane-eczacisi", name: "Hastane Eczacısı", level: "Uzman", tasks: [
        ["Son Kullanma Tarihi Takibi", "İlaç ve sarf stoklarını son kullanma tarihine göre listeler, yaklaşanları ve tüketim hızına göre süresi içinde tükenmeyecekleri işaretler.", 2],
        ["Kritik Stok Seviyesi Raporu", "Günlük tüketim ve tedarik süresine göre kritik stok seviyesinin altına düşen ilaç ve malzemeleri listeler.", 2],
      ]},
    ],
  },

  "saglik-yonetim": {
    name: "Hastane Yönetimi", short: "YNT",
    desc: "İdari ve operasyonel performans.",
    roles: [
      { id: "hastane-muduru", name: "Hastane Müdürü", level: "Yönetici", tasks: [
        ["Hastane Operasyon Panosu", "Poliklinik, yatış, ameliyat, doluluk ve gelir göstergelerini bölüm bazında günlük panoda toplar.", 2],
        ["Hekim Performans Raporu", "Hekim bazında muayene, ameliyat, yatış ve gelir göstergelerini raporlar.", 2],
      ]},
    ],
  },
};
