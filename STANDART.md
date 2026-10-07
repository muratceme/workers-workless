# Görev Paketi Standardı

Kütüphanedeki her görev aynı yapıya uyar. `scripts/kontrol.py` bu kuralları CI'da otomatik denetler;
uymayan paket yayınlanmaz.

## Klasör yapısı

```
library/
├── _ortak/llm.py                     # tüm agent'ların paylaştığı sağlayıcı katmanı (tek kaynak)
├── code-blocks/<görev-id>/
│   ├── task.json                     # manifest
│   ├── README.md                     # ne yapar, girdi standardı, kullanım, sınırlamalar
│   ├── main.py                       # tek giriş noktası
│   ├── requirements.txt              # paket==sürüm (sabit)
│   ├── ornek_veri/                   # kurgusal örnek veri + boş şablonlar
│   └── tests/test_*.py               # unittest; örnek veriyle uçtan uca test şart
└── agents/<görev-id>/
    ├── (yukarıdakilerin hepsi; main.py yerine agent.py)
    ├── prompt.md                     # modele verilen talimat
    ├── .env.example                  # gerçek .env asla depoya girmez
    └── llm.py, *_cekirdek.py         # ortak_dosyalar ile kopyalanır, elle düzenlenmez
```

`<görev-id>`, katalogdaki görev adının slug'ıdır (`catalog/catalog.mjs`; ör. "Banka Mutabakatı" → `banka-mutabakati`).
Bir görev birden çok sektörde görünse de kodu tek bir klasördedir.

## task.json

| Alan | Açıklama |
|---|---|
| `id` | Klasör adıyla aynı |
| `tur` | `kod` veya `agent` |
| `ad`, `surum` | Görünen ad, anlamlı sürüm (`1.0.0`) |
| `python` | Desteklenen Python (`>=3.10`) |
| `ana_dosya` | `main.py` veya `agent.py` |
| `calistir` | Kendi verisiyle örnek komut |
| `girdiler`, `ciktilar` | Sitede gösterilen listeler |
| `internet` | Kod bloklarında zorunlu olarak `false` |
| `ortak_dosyalar` | Yalnızca agent'larda: `{"llm.py": "_ortak/llm.py", ...}` |

## Kurallar

**Genel**
- Girdi, sektörden ve şirketten bağımsız **genel bir standarda** göre tanımlanır (ör. `tarih, aciklama, tutar`).
  Yaygın sütun adı varyasyonları kabul edilir; şirkete özel uyarlama paketin işi değildir.
- Argümansız çalıştırıldığında `ornek_veri/` ile çalışmalıdır (`python main.py`).
- Çıktılar `cikti/` klasörüne yazılır; girdiler asla değiştirilmez.
- Okunamayan bir dosya tüm işi durdurmaz, raporda uyarı olarak görünür.
- Konsol çıktısı Windows (cp1254) terminalinde çökmemelidir: `[OK]`, `[!]`, `[X]` kullanın.
- Bağımlılıklar az ve **sabit sürümlü** olmalıdır (`paket==sürüm`). CI'da `pip-audit` ile taranır.

**Kod blokları**
- İnternete bağlanamaz, kabuk komutu çalıştıramaz: `socket`, `urllib`, `http`, `requests`, `httpx`,
  `subprocess`, `os.system` vb. yasaktır (CI engeller).

**Agent'lar**
- Deterministik yapılabilen her adım kodla yapılır; model yalnızca muhakeme gereken kısma girer.
- Modele gitmeden önce kişisel veriler **yerelde maskelenir** (`llm.maskele`).
- Veri gönderilmeden önce ne gönderileceği gösterilir ve **onay** alınır (`llm.onay_al`; `--evet` ile atlanır).
- Model çıktısı **JSON şeması** ile alınır; modelin uydurduğu kimlikler yok sayılır.
- Raporda boş bir **İnsan Kararı / İnsan Onayı** sütunu bulunur.
- Testler gerçek API çağırmaz (`llm.json_iste` sahte fonksiyonla değiştirilir).
- `anthropic`, `openai` ve `ollama` sağlayıcıları desteklenir (`.env` ile seçilir).

## Yeni görev eklemek

1. Katalogda görevin id'sini bulun (yoksa `catalog/catalog.mjs` içine ekleyin).
2. Var olan bir paketi kopyalayıp uyarlayın (`library/code-blocks/banka-mutabakati` iyi bir başlangıçtır).
3. Agent ise `python scripts/kontrol.py --senkronla`, ardından `python scripts/kontrol.py --test`.
4. `npm start` ile sitede kontrol edin ve pull request açın.
