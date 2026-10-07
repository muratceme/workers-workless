# CV Raporlama · AI Agent

> İnsan Kaynakları › İK Uzmanı · Workers / Workless

CV'leri yapay zekâ ile okur: son pozisyon, şirket, deneyim, eğitim, diller, beceriler ve kısa
bir özet çıkarır; iş ilanı verirseniz **0–100 uyum puanını gerekçesiyle** yazar. Düzeni bozuk,
serbest metinli CV'lerde kural tabanlı kod bloğundan daha başarılıdır.

## Gizlilik ve kontrol

- E-posta, telefon, LinkedIn ve ad soyad **bilgisayarınızda** çıkarılır ve modele gönderilen
  metinden **maskelenir**. Rapora yerelde eklenir.
- Herhangi bir veri gönderilmeden önce ne gönderileceği gösterilir ve **onayınız** istenir.
- Model; yaş, cinsiyet, medeni durum gibi bilgileri değerlendirmemesi için yönlendirilmiştir
  (`prompt.md`). Raporda boş bir **İnsan Kararı** sütunu bulunur; son kararı insan verir.
- Veri hiç dışarı çıkmasın istiyorsanız `WW_PROVIDER=ollama` ile yerel model kullanın.

## Kurulum

```bash
pip install -r requirements.txt
cp .env.example .env        # Windows: copy .env.example .env
```

`.env` dosyasını açıp kendi API anahtarınızı yazın. API ücretleri size aittir.

| Değişken | Açıklama |
|---|---|
| `WW_PROVIDER` | `anthropic` (varsayılan), `openai` veya `ollama` |
| `ANTHROPIC_API_KEY` / `OPENAI_API_KEY` | Sağlayıcı anahtarınız |
| `WW_MODEL` | Model adı. Anthropic için varsayılan `claude-opus-5-5` |
| `WW_EFFORT` | Anthropic çaba düzeyi: `low`, `medium` (varsayılan), `high`. Düşük = daha ucuz |

## Kullanım

```bash
# Örnek veri + örnek ilan ile deneyin
python agent.py

# Kendi CV'leriniz
python agent.py --girdi "C:/İK/Başvurular" --ilan ilan.txt --cikti cikti/aday_raporu_ai.xlsx
```

| Parametre | Açıklama |
|---|---|
| `--girdi` | CV klasörü (`.pdf`, `.docx`, `.txt`) |
| `--ilan` | İş ilanı metni; verilmezse puanlama yapılmaz |
| `--cikti` | Excel çıktısı (varsayılan `cikti/aday_raporu_ai.xlsx`) |
| `--paralel` | Aynı anda analiz edilen CV sayısı (varsayılan 4) |
| `--evet` | Onay sorusunu atla (otomasyon için) |
| `--maskeleme-kapali` | Kişisel bilgileri maskelemeden gönder (önerilmez) |

## Maliyet

Her CV için bir istek gönderilir. Çalışma sonunda kullanılan token sayısı yazdırılır.
Maliyeti düşürmek için `WW_EFFORT=low` veya daha küçük bir model (`WW_MODEL`) seçebilirsiniz.

## Testler

Testler gerçek API'yi çağırmaz:

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur; çıktılar karar desteğidir. Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
