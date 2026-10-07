/* Bankacılık — birim adları İş Bankası ve Ziraat Bankası organizasyon yapılarına,
   unvanlar Yapı Kredi, TEB, Burgan Bank ilanlarına dayanır (catalog/KAYNAKLAR.md). */

export const BANKA = {
  "banka-sube": {
    name: "Şube Bankacılığı", short: "ŞUB",
    desc: "Gişe, şube operasyonu, bireysel ve ticari portföy yönetimi.",
    roles: [
      { id: "gise-yetkilisi", name: "Gişe Yetkilisi", level: "Giriş", tasks: [
        ["Gişe Gün Sonu Kasa Mutabakatı", "Gün sonu kasa sayımını (kupür bazında) sistem bakiyesiyle karşılaştırır, döviz ve TL farklarını işlem bazında araştırma listesiyle raporlar.", 3],
        ["Müşteri Talimat Kontrolü", "EFT, havale ve virman talimatlarında IBAN, unvan, tutar (rakam-yazı) ve imza yetkisi alanlarını kontrol edip eksik veya hatalı talimatları işaretler.", 4],
        "VKN / TCKN / IBAN Doğrulama",
      ]},
      { id: "operasyon-yetkilisi", name: "Operasyon Yetkilisi", level: "Giriş", tasks: [
        ["Resmî Yazı (Haciz/Müzekkere) Takibi", "İcra dairesi, mahkeme ve kurumlardan gelen haciz, bilgi talebi ve müzekkereleri tür, yasal cevap süresi ve müşteri bazında listeler, süresi yaklaşanları uyarır.", 4],
        ["Müşteri Evrak (KYC) Eksik Kontrolü", "Müşteri dosyalarını ürün ve müşteri tipine göre zorunlu belge listesiyle karşılaştırıp eksik ve süresi dolan belgeleri çıkarır.", 3],
        ["Çek Takas Kontrol Listesi", "Takasa verilecek ve takastan gelen çeklerde keşide tarihi, karekod, ciro silsilesi ve tutar tutarlılığı kontrol listesini üretir.", 2],
      ]},
      { id: "sube-operasyon-yoneticisi", name: "Şube Operasyon Yöneticisi", level: "Yönetici", tasks: [
        ["Şube Operasyonel Kontrol Listesi", "Günlük, haftalık ve aylık periyodik kontrolleri (kasa, kıymetli evrak, anahtar, yetki) sorumlu ve tamamlanma durumuyla takip eder.", 2],
        ["Kıymetli Evrak Sayım Kontrolü", "Kıymetli evrak (çek karnesi, kart, teminat mektubu formu) sayımını kayıtlarla karşılaştırıp seri numarası farklarını listeler.", 2],
      ]},
      { id: "bireysel-portfoy-yoneticisi", name: "Bireysel Portföy Yöneticisi", level: "Uzman", tasks: [
        ["Portföy Müşteri Fırsat Listesi", "Portföydeki müşterilerin ürün sahipliği ve hareketlerinden çapraz satış fırsatlarını (kart, BES, sigorta, mevduat) önceliklendirir.", 4],
        ["Vadesi Dolan Mevduat Takibi", "Vadesi yaklaşan mevduat ve yatırım ürünlerini müşteri bazında listeler, görüşme önceliği önerir.", 2],
        "Ziyaret Notu Özeti",
      ]},
      { id: "ticari-portfoy-yoneticisi", name: "Ticari Portföy Yöneticisi", level: "Uzman", tasks: [
        ["Kredi Teklif Dosyası Hazırlama", "Firma bilgisi, mali tablolar, istihbarat ve teminat bilgilerinden krediler tahsis birimine gönderilecek teklif özetini hazırlar.", 5],
        "Mali Tablo Rasyo Analizi",
        ["Firma İstihbarat Raporu Derleme", "Ticaret sicili, banka istihbaratı, piyasa görüşmeleri ve risk bilgilerini tek bir istihbarat raporunda toplar.", 4],
        ["Teminat Takip Listesi", "Kredi teminatlarının (ipotek, çek, kefalet, teminat mektubu) değer, vade ve ekspertiz yenileme tarihlerini takip eder.", 2],
      ]},
      { id: "sube-muduru", name: "Şube Müdürü", level: "Yönetici", tasks: [
        ["Şube Hedef-Gerçekleşme Panosu", "Mevduat, kredi, kart ve sigorta hedeflerinin portföy yöneticisi bazında gerçekleşmesini tek panoda toplar.", 3],
        ["Şube Risk ve Gecikme Listesi", "Şube portföyündeki gecikmeye düşen ve erken uyarı veren kredileri tutar ve gün bazında önceliklendirir.", 2],
      ]},
    ],
  },

  "banka-krediler-tahsis": {
    name: "Krediler Tahsis", short: "KRT",
    desc: "Kurumsal, ticari ve perakende kredi tekliflerinin değerlendirilmesi.",
    roles: [
      { id: "kredi-tahsis-uzman-yardimcisi", name: "Kredi Tahsis Uzman Yardımcısı", level: "Giriş", tasks: [
        "Mali Tablo Rasyo Analizi",
        ["Kredi Teklifi Eksik Belge Kontrolü", "Şubeden gelen kredi teklifini teklif tipine göre zorunlu belge ve bilgi listesiyle karşılaştırır, eksikleri şubeye iletilecek formatta çıkarır.", 3],
        ["KKB ve Risk Merkezi Raporu Özeti", "Kredi kayıt ve risk merkezi raporlarından toplam risk, banka sayısı, gecikme ve limit kullanım özetini çıkarır.", 3],
      ]},
      { id: "kredi-tahsis-uzmani", name: "Kredi Tahsis Uzmanı", level: "Uzman", tasks: [
        ["Geri Ödeme Kapasitesi Analizi", "Firma nakit akışı, borç servis tutarları ve önerilen kredi yapısından borç servis karşılama oranını ve senaryo analizini hesaplar.", 4],
        ["Sektör Karşılaştırmalı Rasyo Analizi", "Firmanın rasyolarını sektör ortalamalarıyla (TCMB sektör bilançoları) karşılaştırıp sapmaları yorumlar.", 3],
        ["Kredi Komitesi Değerlendirme Notu", "Teklif, analiz ve istihbarat bulgularından güçlü-zayıf yönler, riskler ve öneriyi içeren komite notu taslağı hazırlar.", 4],
      ]},
      { id: "kredi-tahsis-muduru", name: "Kredi Tahsis Müdürü", level: "Yönetici", tasks: [
        ["Kredi Portföy Yoğunlaşma Analizi", "Kredi portföyünü sektör, risk grubu ve teminat türüne göre dağıtıp yoğunlaşma (HHI) ve limit aşımlarını raporlar.", 2],
        ["Tahsis Süreç Performans Raporu", "Tekliflerin bekleme süreleri, onay/ret oranları ve iade nedenlerini raporlar.", 2],
      ]},
    ],
  },

  "banka-krediler-izleme": {
    name: "Krediler İzleme ve Takip", short: "KRİ",
    desc: "Erken uyarı, gecikme yönetimi, yapılandırma ve yasal takip.",
    roles: [
      { id: "kredi-izleme-uzmani", name: "Kredi İzleme Uzmanı", level: "Uzman", tasks: [
        ["Erken Uyarı Sinyalleri Listesi", "Gecikme, karşılıksız çek, protestolu senet, haciz, limit aşımı ve ciro düşüşü sinyallerini birleştirip firma bazında risk puanı üretir.", 4],
        ["Gecikme Yaşlandırma Raporu", "Kredileri gecikme gün aralıklarına (1-30, 31-60, 61-90, 90+) göre dağıtıp şube ve ürün bazında raporlar.", 3],
        ["TFRS 9 Aşama Sınıflandırması", "Gecikme günü ve risk göstergelerine göre kredileri Aşama 1-2-3 olarak sınıflandırır, aşama değişenleri listeler.", 3],
        ["Yapılandırma Takip Listesi", "Yeniden yapılandırılan kredilerin ödeme planı uyumunu ve izleme süresini takip eder.", 2],
      ]},
      { id: "yasal-takip-uzmani", name: "Yasal Takip Uzmanı", level: "Uzman", tasks: [
        ["Takibe Aktarılacak Krediler Listesi", "90 günü aşan veya kat edilen kredileri gerekli evrak durumuyla birlikte yasal takibe aktarım listesine dönüştürür.", 2],
        "İcra Takip Durum Raporu",
        ["Tahsilat ve Avukat Performans Raporu", "Takipteki alacakların tahsilat tutarlarını sözleşmeli avukat ve dosya yaşı bazında raporlar.", 2],
      ]},
    ],
  },

  "banka-hazine": {
    name: "Hazine ve Aktif-Pasif Yönetimi", short: "HZN",
    desc: "Likidite, döviz pozisyonu, faiz riski ve piyasa işlemleri.",
    roles: [
      { id: "hazine-uzmani", name: "Hazine Uzmanı", level: "Uzman", tasks: [
        ["Günlük Likidite Raporu", "Vadesi gelen varlık ve yükümlülüklerden günlük ve haftalık likidite açığı/fazlası tablosunu hazırlar.", 3],
        ["Döviz Pozisyonu Raporu", "Bilanço içi ve dışı döviz varlık ve yükümlülüklerinden döviz cinsi bazında net pozisyonu hesaplar.", 2],
        ["Piyasa Günlüğü Özeti", "Gün içi piyasa verilerinden (kur, faiz, CDS, endeks) kısa piyasa özeti metni hazırlar.", 2],
      ]},
      { id: "aktif-pasif-yonetimi-uzmani", name: "Aktif-Pasif Yönetimi Uzmanı", level: "Uzman", tasks: [
        ["Faiz Riski Vade Uyumsuzluk (Gap) Analizi", "Faize duyarlı varlık ve yükümlülükleri yeniden fiyatlama vadelerine göre dağıtıp gap tablosu ve faiz şoku etkisini hesaplar.", 3],
        ["Mevduat Maliyeti ve Kredi Getirisi Analizi", "Ürün bazında ağırlıklı ortalama mevduat maliyetini ve kredi getirisini hesaplayıp marj analizini yapar.", 2],
      ]},
    ],
  },

  "banka-operasyon": {
    name: "Bankacılık Operasyonları", short: "OPR",
    desc: "Merkezi operasyon: ödeme sistemleri, dış ticaret ve kredi operasyonları.",
    roles: [
      { id: "odeme-sistemleri-operasyon-uzmani", name: "Ödeme Sistemleri Operasyon Uzmanı", level: "Uzman", tasks: [
        ["İade Gelen EFT/FAST Analizi", "İade gelen ödemeleri iade kodu, şube ve hata türüne göre sınıflandırıp tekrarlayan hata nedenlerini raporlar.", 3],
        ["Toplu Ödeme Dosyası Kontrolü", "Müşterilerin yüklediği toplu ödeme (maaş, tedarikçi) dosyalarında format, IBAN, mükerrer satır ve toplam kontrolü yapar.", 3],
      ]},
      { id: "dis-ticaret-operasyon-uzmani", name: "Dış Ticaret Operasyon Uzmanı", level: "Uzman", tasks: [
        "Akreditif Evrak Uygunluk Kontrolü",
        ["Teminat Mektubu Metin Kontrolü", "Teminat mektubu taleplerinde lehtar, tutar, süre ve özel şartları banka standart metniyle karşılaştırıp sapmaları işaretler.", 2],
      ]},
      { id: "kredi-operasyon-uzmani", name: "Kredi Operasyon Uzmanı", level: "Uzman", tasks: [
        ["Kredi Kullandırım Ön Kontrolü", "Onay şartları, teminat tesis durumu ve sözleşme bilgilerini kullandırım talebiyle karşılaştırıp eksikleri listeler.", 3],
        ["Ödeme Planı Hesaplama", "Kredi tutarı, faiz, vade ve ödeme tipine (eşit taksit, eşit anapara, balonlu) göre BSMV ve KKDF dahil ödeme planı çıkarır.", 2],
      ]},
    ],
  },

  "banka-uyum": {
    name: "Kurumsal Uyum", short: "UYM",
    desc: "MASAK, suç gelirlerinin aklanmasının önlenmesi, yaptırımlar ve mevzuat uyumu.",
    roles: [
      { id: "uyum-uzmani", name: "Uyum Uzmanı", level: "Uzman", tasks: [
        ["Şüpheli İşlem Senaryo Taraması", "İşlem verisini yapılandırılabilir senaryolara (parçalı işlem, hızlı para transferi, nakit yoğunluğu, riskli ülke) göre tarayıp inceleme listesi çıkarır.", 6],
        ["Yaptırım Listesi Taraması", "Müşteri ve karşı taraf listesini güncel yaptırım listeleriyle Türkçe karakter ve yazım farklarını dikkate alarak bulanık eşleştirir.", 4],
        ["Şüpheli İşlem Bildirim Taslağı", "İnceleme notları ve işlem dökümünden MASAK şüpheli işlem bildirimi için olay özeti ve şüphe gerekçesi taslağı hazırlar.", 3],
      ]},
      { id: "uyum-muduru", name: "Uyum Müdürü", level: "Yönetici", tasks: [
        "Mevzuat Değişikliği Özeti",
        ["Uyum Programı Gösterge Raporu", "İnceleme, bildirim, eğitim ve yaptırım taraması göstergelerini dönemsel uyum raporunda toplar.", 2],
      ]},
    ],
  },

  "banka-ic-sistemler": {
    name: "İç Kontrol ve Risk Yönetimi", short: "İÇS",
    desc: "İç kontrol testleri, kredi, piyasa ve operasyonel risk ölçümü.",
    roles: [
      { id: "ic-kontrol-uzmani", name: "İç Kontrol Uzmanı", level: "Uzman", tasks: [
        ["Kontrol Testi Örneklem Seçimi", "İşlem popülasyonundan risk ve tutar ağırlıklı, tekrarlanabilir (tohumlu) örneklem seçer.", 2],
        ["Kontrol Bulgu Takibi", "İç kontrol bulgularını sorumlu birim, aksiyon ve kapanış tarihine göre takip edip gecikenleri raporlar.", 2],
      ]},
      { id: "risk-yonetimi-uzmani", name: "Risk Yönetimi Uzmanı", level: "Uzman", tasks: [
        ["Tarihsel Riske Maruz Değer (VaR)", "Portföy pozisyonları ve geçmiş fiyat serilerinden tarihsel simülasyonla günlük VaR ve beklenen kaybı hesaplar.", 3],
        ["Operasyonel Risk Kayıp Veri Analizi", "Operasyonel kayıp olaylarını olay türü ve iş kolu bazında sınıflandırıp sıklık ve şiddet analizi yapar.", 2],
        ["Stres Testi Senaryo Hesabı", "Kur, faiz ve takip oranı şoklarının sermaye yeterliliği ve kârlılık üzerindeki etkisini senaryo bazında hesaplar.", 3],
      ]},
    ],
  },

  "banka-teftis": {
    name: "Teftiş Kurulu", short: "TFT",
    desc: "Şube ve genel müdürlük birimlerinin denetimi.",
    roles: [
      { id: "mufettis-yardimcisi", name: "Müfettiş Yardımcısı", level: "Giriş", tasks: [
        "Kontrol Testi Örneklem Seçimi",
        ["Şube Teftiş Veri Analizi", "Şube işlem verisinde mesai dışı işlem, personel-müşteri hesap ilişkisi, sık iptal ve limit aşımı gibi riskli örüntüleri tarar.", 4],
      ]},
      { id: "mufettis", name: "Müfettiş", level: "Uzman", tasks: [
        ["Teftiş Raporu Bulgu Taslağı", "Çalışma kâğıtlarındaki tespitlerden mevzuat dayanağı, risk ve öneri içeren bulgu metinleri taslağı hazırlar.", 4],
      ]},
    ],
  },
};
