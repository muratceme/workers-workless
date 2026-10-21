# Gerisi Bende

> Sen kahveni iç, gerisi bende. (Proje eskiden *Workers / Workless* adını taşıyordu; depo adresi aynı kaldı.)

Her sektörün, her departmanın, her rolün tekrar eden işleri için **açık kaynak kod blokları** ve
**AI agent'lar**. Ücretsiz.

- **Kod Blokları:** kural tabanlı Python betikleri. API anahtarı gerekmez, internete bağlanmaz.
- **AI Agents:** muhakeme gereken işler için. Kendi API anahtarınızla (Claude, OpenAI) ya da
  yerel modelle (Ollama) kendi bilgisayarınızda çalışır.

Sitede Sektör → Departman → Rol → Görev akışıyla ilerleyip görevi indirirsiniz.

## Hazır paketler

357 kod bloğu ve 109 AI agent hazırdır; katalog 16 sektör, 95 departman ve 226 rolü kapsar; her rolün iki modda da
en az bir hazır paketi vardır. Her ZIP'te terminal gerektirmeyen bir başlatıcı bulunur: Windows'ta `BASLAT.bat`,
macOS / Linux'ta `BASLAT.command` çift tıklanır; Python yoksa kurmaya yardım eder, kütüphaneleri kendisi yükler,
"örnek veriyle dene" ya da "kendi dosyalarımla çalıştır" seçeneklerini sunar (kaynak: `scripts/baslatici/`). Bir görevi talep etmek ya da kurulum ve uyarlama için yardım almak üzere sitedeki
"Yardım al" bölümünü kullanabilir veya "Görev talebi" şablonuyla bir issue açabilirsiniz.

Amaç, her meslekte tekrar eden rutin işi azaltmaktır: araçlar ücretsiz ve açık kaynaktır, veri kullanıcının
bilgisayarında kalır, son söz kullanıcıdadır.

## Sizden Gelenler

Sitedeki **Sizden Gelenler** sekmesi (`#/sizden`), ekiplerin bize ilettiği otomasyon isteklerine göre yazılan paketleri
gösterir. Bu paketler ana katalogdan ayrıdır ve yukarıdaki sayılara girmez: istekler `catalog/istekler.mjs`'de,
paketler `library/istekler/` altındadır. İsteği gönderenin adı ve iletişim bilgisi yazılmaz.

İlk istek, abonelikli bir mobil uygulama ekibinin 7 otomasyonudur: destek e-postaları (agent), kullanıcı ve tutma
analizi, abonelik gelirleri, Crashlytics çökmeleri, mağaza yorumları (agent), yapay zekâ API maliyeti ve bunları
birleştiren günlük / haftalık / aylık yönetim özeti. Paketler mağaza, Firebase ve Crashlytics dökümlerini dosyadan
okur; ortak okuma katmanı `library/_ortak/tablo.py`'dir. Her paket `<paket>.ozet.json` yazar, yönetim özeti bunları
birleştirir.

## Kaynaklar

Mevzuata dayanan paketlerin README'lerinde **"Kaynaklar ve dayanaklar"** bölümü vardır: resmî metin bağlantıları
(mevzuat.gov.tr), ikincil kaynaklar, kontrol tarihi (15.10.2026) ve hangi maddelerin resmî metinle karşılaştırıldığı
burada yazar. Teyit edilemeyen noktalar açıkça "Teyit edilemedi" olarak işaretlenmiştir.

`main` dalına her push'ta GitHub Actions çalışır: standart kontrolü ve testler (Windows ve Linux,
Python 3.10 ve 3.13) ile `pip-audit` güvenlik taraması geçerse site GitHub Pages'e yayınlanır.

## Geliştirme

```bash
npm start                                  # yerel sunucu (http://localhost:4321), değişiklikte yeniden derler
node scripts/build.mjs                     # dist/ (site + ZIP'ler)
python scripts/kontrol.py --test           # standart kontrolü (başlatıcı ve önizleme güncelliği dahil) + her paketin testleri
.venv/Scripts/python scripts/onizleme.py   # kod bloklarını (Sizden Gelenler dahil: istekler/<id>) örnek veriyle çalıştırıp çıktı önizlemelerini üretir → data/onizlemeler.json
node scripts/build-seo.mjs                 # deneme: görev / meslek / sektör başına statik sayfalar + sitemap → dist-seo/
node scripts/sunucu-seo.mjs                # SEO sürümünü önizle (http://localhost:4322)
node scripts/og_gorsel.mjs                 # paylaşım görseli (site/assets/og.png; Playwright gerekir)
```

Site ayarları `catalog/catalog.mjs` içindeki `CONFIG`'de: iletişim adresleri, görev talep formu adresi
(`talepFormu`) ve çerezsiz ziyaretçi ölçümü kodu (`goatcounter`). Boş bırakılanlar sitede görünmez.

## Lisans ve sorumluluk

MIT Lisansı. Yazılım ücretsiz ve "olduğu gibi" sunulur; **tüm risk kullanıcıdadır.**

- Hatalı veya eksik sonuç, kaçırılan süre, ceza, veri kaybı ve diğer her türlü zarardan yazarlar sorumlu tutulamaz.
- Çıktılar karar desteğidir; kullanmadan önce yetkin bir kişi kontrol etmelidir.
- Hukuki, mali, vergisel, tıbbi veya yatırım danışmanlığı değildir.
- Mevzuat değişir; süre ve oranların güncelliğini kontrol etmek kullanıcıya aittir.
- AI agent'larda model hata yapabilir. Maskeleme yalnız telefon, e-posta, TCKN, IBAN ve LinkedIn adresini kapsar.
  API ücretleri ve KVKK uyumu kullanıcıya aittir.

Ayrıntılar: [LICENSE](LICENSE) ve [SORUMLULUK_REDDI.md](SORUMLULUK_REDDI.md). Her paketin README'sinde de aynı
"Sorumluluk" bölümü bulunur.
