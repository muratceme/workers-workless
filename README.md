# Workers / Workless

Her sektörün, her departmanın, her rolün tekrar eden işleri için **açık kaynak kod blokları** ve
**AI agent'lar**. Ücretsiz.

- **Kod Blokları:** kural tabanlı Python betikleri. API anahtarı gerekmez, internete bağlanmaz.
- **AI Agents:** muhakeme gereken işler için. Kendi API anahtarınızla (Claude, OpenAI) ya da
  yerel modelle (Ollama) kendi bilgisayarınızda çalışır.

Sitede Sektör → Departman → Rol → Görev akışıyla ilerleyip görevi indirirsiniz.

## Hazır paketler

| Görev | Kod Bloğu | AI Agent |
|---|---|---|
| CV Raporlama (İK › İK Uzmanı) | ✅ | ✅ |
| Banka Mutabakatı (Finans › Muhasebe Elemanı) | ✅ | ✅ |

Katalogdaki diğer görevler sitede "Yakında" olarak görünür. Bir görevi talep etmek için
"Görev talebi" şablonuyla bir issue açabilirsiniz.

## Depo yapısı

```
catalog/catalog.mjs   Sektör → departman → rol → görev kataloğu
library/              Görev paketleri (bkz. STANDART.md)
site/                 Statik site kaynağı (bağımlılıksız HTML/CSS/JS)
scripts/build.mjs     Siteyi dist/ içine derler, her paket için ZIP + SHA-256 üretir
scripts/kontrol.py    Kalite kapısı: standart, sabit bağımlılık, ağ yasağı, testler
server.js             Yerel geliştirme sunucusu (değişiklikte yeniden derler)
```

## Yerel geliştirme

Node.js 18+ ve Python 3.10+ gerekir.

```bash
npm start                         # http://localhost:4321
python scripts/kontrol.py --test  # standart kontrolü + tüm testler
```

## Yayın

`main` dalına her push'ta GitHub Actions çalışır: standart kontrolü ve testler (Windows ve Linux,
Python 3.10 ve 3.13) ile `pip-audit` güvenlik taraması geçerse site GitHub Pages'e yayınlanır.

## Lisans ve sorumluluk

MIT Lisansı. Yazılım "olduğu gibi" sunulur; kullanımı kullanıcının sorumluluğundadır.
Ayrıntılar: [LICENSE](LICENSE) ve [SORUMLULUK_REDDI.md](SORUMLULUK_REDDI.md).
