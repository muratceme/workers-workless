# Workers / Workless

Her sektörün, her departmanın, her rolün tekrar eden işleri için **açık kaynak kod blokları** ve
**AI agent'lar**. Ücretsiz.

- **Kod Blokları:** kural tabanlı Python betikleri. API anahtarı gerekmez, internete bağlanmaz.
- **AI Agents:** muhakeme gereken işler için. Kendi API anahtarınızla (Claude, OpenAI) ya da
  yerel modelle (Ollama) kendi bilgisayarınızda çalışır.

Sitede Sektör → Departman → Rol → Görev akışıyla ilerleyip görevi indirirsiniz.

## Hazır paketler

357 kod bloğu ve 76 AI agent hazırdır; katalog 16 sektör, 95 departman ve 226 rolü kapsar. Hazır olmayan görevler
sitede "Yakında" olarak görünür. Bir görevi talep etmek ya da kurulum ve uyarlama için yardım almak üzere sitedeki
"Yardım al" bölümünü kullanabilir veya "Görev talebi" şablonuyla bir issue açabilirsiniz.

Amaç, her meslekte tekrar eden rutin işi azaltmaktır: araçlar ücretsiz ve açık kaynaktır, veri kullanıcının
bilgisayarında kalır, son söz kullanıcıdadır.

## Kaynaklar

Mevzuata dayanan paketlerin README'lerinde **"Kaynaklar ve dayanaklar"** bölümü vardır: resmî metin bağlantıları
(mevzuat.gov.tr), ikincil kaynaklar, kontrol tarihi (15.10.2026) ve hangi maddelerin resmî metinle karşılaştırıldığı
burada yazar. Teyit edilemeyen noktalar açıkça "Teyit edilemedi" olarak işaretlenmiştir.

`main` dalına her push'ta GitHub Actions çalışır: standart kontrolü ve testler (Windows ve Linux,
Python 3.10 ve 3.13) ile `pip-audit` güvenlik taraması geçerse site GitHub Pages'e yayınlanır.

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
