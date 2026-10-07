/* =========================================================================
   Workers / Workless — katalog
   Sektör → Departman → Rol (çalışan) → Görev

   - Departmanlar: catalog/ortak.mjs (tüm sektörlerde ortak) + catalog/sektorler/*.mjs
   - Görev formatı: [ad, açıklama, haftalık tahmini kazanç (saat)]. Aynı görev başka bir rolde
     yalnızca adıyla anılır ("Banka Mutabakatı"); kodu library/<tür>/<görev-id>/ altında tek yerdedir.
   - Görev kimliği (id) = görev adının slug'ı.
   - Adlandırma kaynakları: catalog/KAYNAKLAR.md
   ========================================================================= */
import { ORTAK } from "./ortak.mjs";
import { BANKA } from "./sektorler/banka.mjs";
import { SIGORTA, ACENTE } from "./sektorler/sigorta.mjs";
import { URETIM } from "./sektorler/uretim.mjs";
import { TEKSTIL } from "./sektorler/tekstil.mjs";
import { PERAKENDE } from "./sektorler/perakende.mjs";
import { ETICARET } from "./sektorler/eticaret.mjs";
import { SAGLIK } from "./sektorler/saglik.mjs";
import { LOJISTIK } from "./sektorler/lojistik.mjs";
import { INSAAT } from "./sektorler/insaat.mjs";
import { GIDA } from "./sektorler/gida.mjs";
import { OTEL } from "./sektorler/otel.mjs";
import { SMMM, HUKUK_BUROSU, OTOMOTIV, EGITIM } from "./sektorler/hizmet.mjs";

export const CONFIG = {
  githubUser: "muratceme",
  repo: "workers-workless",
  branch: "main",
};

const PARCALAR = { ORTAK, BANKA, SIGORTA, ACENTE, URETIM, TEKSTIL, PERAKENDE, ETICARET, SAGLIK, LOJISTIK, INSAAT, GIDA, OTEL, SMMM, HUKUK_BUROSU, OTOMOTIV, EGITIM };

export const DEPTS = {};
for (const [parca, depts] of Object.entries(PARCALAR)) {
  for (const [id, d] of Object.entries(depts)) {
    if (DEPTS[id]) throw new Error(`Departman kimliği "${id}" iki kez tanımlanmış (${parca}).`);
    DEPTS[id] = d;
  }
}

const ORTAK_TEMEL = ["ik", "muhasebe", "finans", "bt"];

export const SECTORS = [
  { id: "banka", name: "Bankacılık", icon: "bank", tagline: "Şube, krediler, hazine, operasyon, uyum ve teftiş.",
    depts: ["banka-sube", "banka-krediler-tahsis", "banka-krediler-izleme", "banka-hazine", "banka-operasyon", "banka-uyum", "banka-ic-sistemler", "banka-teftis", ...ORTAK_TEMEL, "hukuk", "musteri-hizmetleri", "pazarlama"] },
  { id: "sigorta", name: "Sigortacılık", icon: "shield", tagline: "Risk kabul, hasar, sağlık, aktüerya ve reasürans.",
    depts: ["sigorta-teknik", "sigorta-hasar", "sigorta-saglik", "sigorta-akturya", "sigorta-reasurans", "sigorta-satis", "sigorta-tahsilat", ...ORTAK_TEMEL, "hukuk", "musteri-hizmetleri"] },
  { id: "sigorta-acentesi", name: "Sigorta Acenteliği", icon: "umbrella", tagline: "Teklif karşılaştırma, poliçe yenileme ve hasar takibi.",
    depts: ["acente-teknik", "muhasebe", "satis", "musteri-hizmetleri"] },
  { id: "uretim", name: "Üretim ve Sanayi", icon: "factory", tagline: "Üretim, planlama, kalite, bakım, Ar-Ge ve tedarik.",
    depts: ["uretim-uretim", "uretim-planlama", "uretim-kalite", "uretim-bakim", "uretim-endustri", "uretim-arge", "satinalma", "depo", "dis-ticaret", "isg", ...ORTAK_TEMEL, "satis", "idari"] },
  { id: "tekstil", name: "Tekstil ve Hazır Giyim", icon: "thread", tagline: "Merchandising, modelhane, planlama, fason ve kalite.",
    depts: ["tekstil-merchandising", "tekstil-tasarim", "tekstil-modelhane", "tekstil-planlama", "tekstil-tedarik", "tekstil-fason", "tekstil-kalite", "tekstil-sevkiyat", "tekstil-uygunluk", "dis-ticaret", "isg", ...ORTAK_TEMEL] },
  { id: "perakende", name: "Perakende", icon: "bag", tagline: "Mağaza, kategori, alokasyon, CRM ve kayıp önleme.",
    depts: ["perakende-magaza", "perakende-kategori", "perakende-planlama", "perakende-gorsel", "perakende-crm", "perakende-kayip-onleme", "depo", ...ORTAK_TEMEL, "pazarlama", "musteri-hizmetleri"] },
  { id: "eticaret", name: "E-ticaret", icon: "cart", tagline: "Pazaryerleri, içerik, operasyon ve performans pazarlama.",
    depts: ["eticaret-pazaryeri", "eticaret-icerik", "eticaret-operasyon", "eticaret-performans", "eticaret-musteri", "depo", "muhasebe", "finans", "pazarlama"] },
  { id: "saglik", name: "Sağlık (Özel Hastane)", icon: "pulse", tagline: "Hasta hizmetleri, faturalama, kalite ve uluslararası hasta.",
    depts: ["saglik-hasta-hizmetleri", "saglik-tibbi-sekreterlik", "saglik-anlasmali-kurumlar", "saglik-kalite", "saglik-uluslararasi", "saglik-eczane", "saglik-yonetim", "satinalma", ...ORTAK_TEMEL, "isg"] },
  { id: "lojistik", name: "Lojistik ve Taşımacılık", icon: "truck", tagline: "Karayolu, parsiyel, deniz-hava, gümrük, depo ve filo.",
    depts: ["lojistik-karayolu", "lojistik-parsiyel", "lojistik-forwarding", "lojistik-gumruk", "lojistik-depo", "lojistik-filo", "lojistik-satis", ...ORTAK_TEMEL, "isg", "musteri-hizmetleri"] },
  { id: "insaat", name: "İnşaat", icon: "crane", tagline: "Teknik ofis, şantiye, planlama, ihale ve kalite.",
    depts: ["insaat-teknik-ofis", "insaat-santiye", "insaat-planlama", "insaat-ihale", "insaat-kalite", "satinalma", "isg", ...ORTAK_TEMEL, "hukuk"] },
  { id: "gida", name: "Gıda Üretimi", icon: "wheat", tagline: "Gıda güvenliği, laboratuvar, üretim, Ar-Ge ve kanal satışı.",
    depts: ["gida-kalite-guvence", "gida-laboratuvar", "gida-uretim", "gida-arge", "gida-satis", "uretim-planlama", "satinalma", "depo", "isg", ...ORTAK_TEMEL] },
  { id: "otel", name: "Turizm ve Otelcilik", icon: "bed", tagline: "Ön büro, gelir yönetimi, kat hizmetleri ve yiyecek-içecek.",
    depts: ["otel-on-buro", "otel-rezervasyon", "otel-kat-hizmetleri", "otel-yiyecek-icecek", "otel-misafir-iliskileri", "satinalma", "ik", "muhasebe", "pazarlama"] },
  { id: "mali-musavirlik", name: "Mali Müşavirlik Bürosu", icon: "calc", tagline: "Defter, beyanname, bordro ve SGK işlemleri.",
    depts: ["smmm-muhasebe", "smmm-beyanname", "smmm-bordro"] },
  { id: "hukuk-burosu", name: "Hukuk Bürosu", icon: "scale", tagline: "Dava ve icra takibi, dilekçe ve sözleşmeler.",
    depts: ["hukuk-dava", "hukuk-icra", "muhasebe"] },
  { id: "otomotiv", name: "Otomotiv Bayi ve Servis", icon: "car", tagline: "Araç satışı, servis, garanti ve yedek parça.",
    depts: ["oto-satis", "oto-servis", "oto-yedek-parca", "muhasebe", "ik", "musteri-hizmetleri"] },
  { id: "egitim", name: "Eğitim (Özel Okul)", icon: "school", tagline: "Kayıt kabul, öğrenci işleri, ölçme değerlendirme.",
    depts: ["egitim-ogrenci-isleri", "egitim-akademik", "muhasebe", "ik"] },
];

export const SOON = [
  { name: "Enerji", icon: "bolt" },
  { name: "Telekomünikasyon", icon: "signal" },
  { name: "Faktoring ve Leasing", icon: "coins" },
  { name: "Kimya ve İlaç", icon: "flask" },
];
