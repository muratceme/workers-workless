/* =========================================================================
   Gerisi Bende — Sizden Gelenler
   Ziyaretçilerin ve müşterilerin bize ilettiği otomasyon isteklerine göre yazılan paketler.
   Ana katalogdan (sektör → departman → rol → görev) ayrıdır; sayılarına ve listelerine girmez.

   - İsteği gönderen anonimdir: kurum adı, kişi adı ve iletişim bilgisi yazılmaz.
   - Madde: { baslik, aciklama, paket (library/istekler/<paket>), durum: "hazir" | "yapiliyor", siklik }
   - Paketler canlı bağlantı kurmaz; dosyadan çalışır. Canlı veri çekme ve zamanlama kurulum hizmetidir.
   ========================================================================= */

export const ISTEKLER = [
  {
    id: "mobil-uygulama-ekibi",
    baslik: "Abonelikli mobil uygulama için 7 otomasyon",
    kim: "Bir mobil uygulama ekibi (abonelikli, yapay zekâ özellikli tüketici uygulaması)",
    tarih: "21.10.2026",
    istek:
      "Destek e-postaları, kullanıcı analizi ve tutma, abonelik gelirleri, çökmeler, mağaza yorumları ve yapay zekâ " +
      "API maliyetleri her gün elle takip ediliyordu. Ekip bunların otomatik raporlanmasını ve her sabah 09.00'da " +
      "tek bir yönetim özeti gelmesini istedi.",
    kosul:
      "İade, veri silme ve hassas taleplerde son onay her zaman ekipte kalmalı; otomasyon karar vermemeli, " +
      "önermeli. Ekip n8n / Make, bir veritabanı ve Qlik kullanıyor.",
    nasil:
      "Her madde ayrı, ücretsiz ve açık kaynak bir paket oldu. Paketler mağaza, Firebase, Crashlytics ve e-posta " +
      "dökümlerini dosyadan okur; internete bağlanmaz (yapay zekâ kullanan ikisi hariç, onlar da yalnız sizin " +
      "onayınızla). Her paket ayrıca küçük bir özet dosyası yazar; Yönetim Özeti bunları birleştirir. Verinin her " +
      "gün otomatik çekilmesi, zamanlama ve raporun e-posta / Slack ile gönderilmesi kurulum hizmetidir.",
    maddeler: [
      {
        baslik: "Destek e-postaları",
        aciklama:
          "E-postaları kategori, duygu ve aciliyete ayırır; P1–P3 öncelik ve ilk yanıt hedefi verir, cevap taslağı " +
          "yazar. İade, veri silme, yasal tehdit ve hesap güvenliği e-postalarını kod bulur ve onaya ayırır; " +
          "taslakta söz verme varsa işaretler.",
        paket: "uygulama-destek-epostasi",
        durum: "hazir",
        siklik: "Her sabah",
      },
      {
        baslik: "Kullanıcı analizi ve tutma",
        aciklama:
          "Firebase / GA4 olay dökümünden günlük aktif ve yeni kullanıcı, oturum, D1 / D7 / D30 tutma (kohort), " +
          "haftalık / aylık aktif, en çok kullanılan özellikler, platform ve sürüm kırılımı.",
        paket: "uygulama-kullanici-ve-tutma-analizi",
        durum: "hazir",
        siklik: "Her gün",
      },
      {
        baslik: "Abonelik gelirleri",
        aciklama:
          "Google Play kazanç raporu ve App Store satış raporundan satış, iade, mağaza komisyonu ve net gelir; " +
          "günlük, aylık, ürün ve ülke kırılımı, abonelik iptalleri. Döviz tutarları verdiğiniz kurla TL'ye çevrilir.",
        paket: "uygulama-abonelik-gelir-raporu",
        durum: "hazir",
        siklik: "Her gün",
      },
      {
        baslik: "Çökme takibi",
        aciklama:
          "Crashlytics olay dökümünden sorunları P1–P3 önceliklendirir: etkilenen kurulum, yeni sürümde ilk kez " +
          "görülen ve sıçrama yapan sorunlar. Ekibe gönderilecek kısa bildirim metnini de yazar.",
        paket: "uygulama-cokme-onceliklendirme",
        durum: "hazir",
        siklik: "Her sabah",
      },
      {
        baslik: "Mağaza yorumları",
        aciklama:
          "Google Play ve App Store yorumlarını kategori ve duyguya ayırır, özellik taleplerini gruplar, cevap " +
          "taslağı yazar. 1–2 yıldızı olumlu sayamaz, yorumda geçmeyen alıntıyı siler, uzun cevabı işaretler.",
        paket: "magaza-yorumu-siniflandirma",
        durum: "hazir",
        siklik: "Her gün",
      },
      {
        baslik: "Yapay zekâ API maliyetleri",
        aciklama:
          "Uygulamanın kendi API çağrı günlüğünden günlük maliyet, model ve özellik kırılımı, kullanıcı başına " +
          "maliyet, kullanım limiti aşımları ve ay sonu tahmini.",
        paket: "yapay-zeka-api-maliyet-raporu",
        durum: "hazir",
        siklik: "Her gün",
      },
      {
        baslik: "Günlük, haftalık ve aylık rapor",
        aciklama:
          "Diğer altı paketin özetlerini birleştirip önceki dönemle karşılaştırmalı kısa bir yönetim raporu yazar " +
          "(metin + Excel). Dikkat listesi: kötü yönde büyük değişim, P1 çökme, bütçe aşımı, güncellenmemiş veri.",
        paket: "uygulama-yonetim-ozeti",
        durum: "hazir",
        siklik: "Her sabah 09.00",
      },
    ],
  },
];
