/* Profesyonel hizmet büroları ve diğer sektörler.
   Mali müşavirlik: eleman.net/secretcv büro ilanları (Muhasebe Elemanı, Mali Müşavir Stajyeri, KDV/Muhtasar-Prim
   Hizmet/Geçici Vergi, e-Fatura/e-Arşiv/e-Defter, SGK bordro). Hukuk bürosu: Uçar Hukuk, BSE Avukatlık, Can Şerbetçi
   Hukuk ilanları (İcra Takip Elemanı/Avukat Katibi, Hukuk Sekreteri, Avukat Asistanı, UYAP).
   Otomotiv bayi: Kamer Plaza, Yüce Auto, Buhari Otomotiv ilanları. Eğitim: Uğur Okulları, Özel Yüce Okulları ilanları. */

export const SMMM = {
  "smmm-muhasebe": {
    name: "Muhasebe ve Defter İşleri", short: "DFT",
    desc: "Mükelleflerin kayıtları, e-belgeler, e-defter ve mutabakatlar.",
    roles: [
      { id: "muhasebe-elemani-buro", name: "Muhasebe Elemanı (Büro)", level: "Giriş", tasks: [
        "e-Fatura Okuma ve Listeleme",
        ["Banka Hareketlerinden Muhasebe Fişi Önerisi", "Mükellefin banka ekstresindeki hareketleri açıklama kurallarına göre hesap kodlarıyla eşleştirip muhasebe fişi önerisi çıkarır.", 6],
        "Banka Mutabakatı",
        "Gelen e-Fatura Kayıt Kontrolü",
        ["e-Defter Berat Yükleme Takibi", "Mükellef bazında e-defter dönemlerinin oluşturma, imzalama ve berat yükleme son tarihlerini takip eder.", 2],
      ]},
      { id: "mali-musavir-stajyeri", name: "Mali Müşavir Stajyeri", level: "Giriş", tasks: [
        ["Mükellef Evrak Eksik Takibi", "Ay sonu her mükelleften gelmesi gereken belgeleri (fatura, banka ekstresi, puantaj, kasa) gelenlerle karşılaştırıp eksik listesini çıkarır.", 3],
        "Fatura KDV Tutarlılık Kontrolü",
      ]},
    ],
  },
  "smmm-beyanname": {
    name: "Beyanname ve Vergi", short: "BYN",
    desc: "KDV, muhtasar-prim hizmet, geçici ve yıllık beyannameler.",
    roles: [
      { id: "smmm", name: "Serbest Muhasebeci Mali Müşavir (SMMM)", level: "Yönetici", tasks: [
        ["Mükellef Beyanname Takvimi", "Mükellef listesi ve yükümlülüklerinden aylık beyanname ve ödeme takvimini mükellef bazında çıkarır, tamamlanma durumunu izler.", 3],
        "KDV Beyannamesi Ön Kontrolü",
        ["Geçici Vergi Hesaplama", "Dönem mizanından kanunen kabul edilmeyen giderler ve geçmiş yıl zararlarıyla geçici vergi matrahını ve ödenecek vergiyi hesaplar.", 3],
        ["Mükellef Bilgilendirme Yazısı", "Mevzuat değişikliği veya beyan sonucunu mükellefe anlaşılır dille açıklayan bilgilendirme yazısı hazırlar.", 2],
        "Mevzuat Değişikliği Özeti",
      ]},
    ],
  },
  "smmm-bordro": {
    name: "Bordro ve SGK İşlemleri", short: "SGK",
    desc: "Mükelleflerin bordro, SGK bildirgeleri ve teşvikleri.",
    roles: [
      { id: "bordro-elemani", name: "Bordro Elemanı", level: "Giriş", tasks: [
        "Brüt-Net Maaş Hesaplama",
        "Puantaj Kontrolü",
        "SGK Giriş-Çıkış Bildirge Kontrolü",
        ["SGK Teşvik Uygunluk Ön Kontrolü", "Çalışan ve işyeri bilgilerine göre yürürlükteki SGK prim teşviklerinin uygunluk koşullarını (yaş, kayıt süresi, borç durumu vb.) ön kontrolden geçirir.", 2],
      ]},
    ],
  },
};

export const HUKUK_BUROSU = {
  "hukuk-dava": {
    name: "Dava Takip", short: "DVA",
    desc: "Dava dosyaları, süreler, dilekçeler ve müvekkil raporlaması.",
    roles: [
      { id: "avukat-buro", name: "Avukat", level: "Uzman", tasks: [
        "Sözleşme Ön İnceleme",
        ["Dilekçe Taslağı", "Olay özeti, talepler ve deliller listesinden dava dilekçesi taslağı hazırlar (avukat kontrolü zorunludur).", 4],
        ["Karar ve İçtihat Özeti", "Mahkeme kararını taraflar, olay, hukuki mesele ve hüküm başlıklarıyla özetler.", 3],
        ["Müvekkil Dosya Durum Raporu", "Dava ve icra dosyalarının son durumunu müvekkile gönderilecek rapor formatında hazırlar.", 2],
      ]},
      { id: "hukuk-sekreteri", name: "Hukuk Sekreteri", level: "Giriş", tasks: [
        ["Duruşma ve Kesin Süre Takvimi", "Dosya listesinden duruşma tarihleri ile tebliğ tarihine göre işleyen kesin süreleri (cevap, istinaf, temyiz) hesaplayıp takvim çıkarır.", 3],
        ["Vekâlet Ücreti ve Masraf Hesabı", "Avukatlık Asgari Ücret Tarifesi ve harç tarifesi parametreleriyle dava değerine göre vekâlet ücreti ve yargılama masraflarını hesaplar.", 2],
      ]},
    ],
  },
  "hukuk-icra": {
    name: "İcra Takip", short: "İCR",
    desc: "Alacak tahsili, icra dosyaları ve borçlu iletişimi.",
    roles: [
      { id: "icra-takip-elemani", name: "İcra Takip Elemanı (Avukat Katibi)", level: "Giriş", tasks: [
        ["İcra Dosyası Faiz ve Alacak Hesabı", "Asıl alacak, takip tarihi ve faiz türüne (yasal, ticari avans) göre dönemsel oranlarla işlemiş faizi ve güncel borcu hesaplar.", 3],
        "İcra Takip Durum Raporu",
        ["Borçlu Ödeme Planı Hesaplama", "Toplam borç ve taksit sayısına göre faiz dahil ödeme planı ve protokol özeti çıkarır.", 2],
      ]},
    ],
  },
};

export const OTOMOTIV = {
  "oto-satis": {
    name: "Satış", short: "STŞ",
    desc: "Sıfır ve ikinci el araç satışı, stok ve kampanyalar.",
    roles: [
      { id: "satis-danismani-otomotiv", name: "Satış Danışmanı", level: "Uzman", tasks: [
        ["Araç Satış Teklifi Hesaplama", "Liste fiyatı, kampanya, takas değeri, kredi ve kasko seçeneklerinden müşteriye sunulacak toplam maliyet ve aylık ödeme teklifini hesaplar.", 3],
        ["Test Sürüşü ve Teklif Takibi", "Test sürüşü ve teklif verilen müşterileri takip edip geri dönüş önceliği belirler.", 2],
      ]},
      { id: "satis-muduru-otomotiv", name: "Satış Müdürü", level: "Yönetici", tasks: [
        ["Araç Stok Yaşlandırma", "Stoktaki araçları stoğa giriş tarihi, model ve renk bazında yaşlandırıp finansman maliyetini hesaplar.", 2],
      ]},
      { id: "ikinci-el-satis-danismani", name: "İkinci El Satış Danışmanı", level: "Uzman", tasks: [
        ["İkinci El Araç Fiyat Değerleme", "Benzer ilan verisinden (marka, model, yıl, km, donanım) piyasa fiyat aralığı ve alım fiyatı önerisi üretir.", 3],
      ]},
    ],
  },
  "oto-servis": {
    name: "Satış Sonrası Hizmetler (Servis)", short: "SRV",
    desc: "Araç kabul, iş emri, randevu, garanti ve müşteri memnuniyeti.",
    roles: [
      { id: "servis-danismani", name: "Servis Danışmanı", level: "Uzman", tasks: [
        ["Servis Randevu Kapasite Planı", "Teknisyen kapasitesi ve iş türü sürelerine göre günlük randevu kapasitesini ve boşlukları hesaplar.", 2],
        ["Periyodik Bakım Hatırlatma Listesi", "Son bakım tarihi ve kilometre tahminine göre bakımı yaklaşan müşterileri listeler, hatırlatma mesajı hazırlar.", 2],
        ["İş Emri Ön Fiyat Tahmini", "Şikâyet ve bakım türüne göre işçilik saati ve parça fiyat listesinden müşteriye ön fiyat tahmini çıkarır.", 2],
      ]},
      { id: "garanti-uzmani", name: "Garanti Uzmanı", level: "Uzman", tasks: [
        ["Garanti Başvuru Dosyası Kontrolü", "Garanti başvurusunda şasi, km, arıza kodu, işçilik kodu ve parça bilgilerinin marka kurallarına uygunluğunu kontrol eder.", 3],
        ["Geri Çağırma Kampanyası Araç Eşleştirme", "Marka kampanya şasi listesini servis müşteri veritabanıyla eşleştirip kampanyası yapılmamış araçları listeler.", 2],
      ]},
      { id: "servis-muduru", name: "Satış Sonrası Hizmetler Müdürü", level: "Yönetici", tasks: [
        ["Teknisyen Verimlilik Raporu", "İş emirlerinden teknisyen bazında satılan saat, çalışılan saat ve verimlilik oranlarını hesaplar.", 2],
        ["Müşteri Memnuniyeti (CSI) Analizi", "Servis memnuniyet anketlerini danışman ve iş türü bazında analiz eder.", 2],
      ]},
    ],
  },
  "oto-yedek-parca": {
    name: "Yedek Parça", short: "YDK",
    desc: "Parça stokları, siparişler ve satış.",
    roles: [
      { id: "yedek-parca-danismani", name: "Yedek Parça Danışmanı", level: "Uzman", tasks: [
        ["Yedek Parça Sipariş Önerisi", "Parça bazında tüketim, bekleyen iş emirleri ve tedarik süresine göre stok siparişi önerir.", 3],
        "Stok ABC-XYZ Analizi",
      ]},
    ],
  },
};

export const EGITIM = {
  "egitim-ogrenci-isleri": {
    name: "Kayıt Kabul ve Öğrenci İşleri", short: "ÖİŞ",
    desc: "Kayıt, veli iletişimi, ücret ve burs işlemleri.",
    roles: [
      { id: "ogrenci-isleri-sorumlusu", name: "Öğrenci İşleri Sorumlusu", level: "Uzman", tasks: [
        ["Devamsızlık Takip Raporu", "Yoklama verisinden öğrenci bazında özürlü/özürsüz devamsızlığı hesaplar, sınıra yaklaşanları listeler.", 2],
        ["Veli Bilgilendirme Mesajları", "Duyuru, etkinlik ve öğrenci durum bilgilerinden kişiselleştirilmiş veli mesajları hazırlar.", 3],
        ["Öğrenim Ücreti ve İndirim Hesabı", "Kardeş, erken kayıt ve burs indirim kurallarına göre öğrenci bazında ücret ve taksit planı hesaplar.", 2],
      ]},
      { id: "kayit-kabul-sorumlusu", name: "Kayıt Kabul Sorumlusu", level: "Uzman", tasks: [
        ["Aday Öğrenci Takip Listesi", "Kayıt görüşmesi yapılan aday öğrencileri aşama, kademe ve takip tarihine göre izler, dönüşüm oranını hesaplar.", 2],
      ]},
    ],
  },
  "egitim-akademik": {
    name: "Ölçme Değerlendirme ve Rehberlik", short: "ÖLÇ",
    desc: "Sınav analizi, akademik takip ve rehberlik.",
    roles: [
      { id: "olcme-degerlendirme-uzmani", name: "Ölçme Değerlendirme Uzmanı", level: "Uzman", tasks: [
        ["Sınav Sonuç ve Madde Analizi", "Optik form veya cevap anahtarı sonuçlarından öğrenci netlerini, soru güçlük ve ayırt edicilik indekslerini hesaplar.", 4],
        ["Kazanım Bazlı Başarı Raporu", "Soruların eşleştiği kazanımlara göre sınıf ve öğrenci bazında başarı yüzdelerini raporlar.", 3],
      ]},
      { id: "rehber-ogretmen", name: "Rehber Öğretmen", level: "Uzman", tasks: [
        ["Öğrenci Gelişim Raporu Taslağı", "Not, devamsızlık ve öğretmen gözlemlerinden veli görüşmesi için öğrenci gelişim raporu taslağı hazırlar.", 3],
      ]},
    ],
  },
};
