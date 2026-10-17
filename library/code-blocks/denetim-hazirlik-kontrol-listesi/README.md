# Denetim Hazırlık Kontrol Listesi · Kod Bloğu

> Kalite Güvence ve Gıda Güvenliği › Kalite Güvence Uzmanı · Gerisi Bende

Sertifikasyon (BRCGS, IFS vb.), resmî kontrol ya da müşteri denetimi öncesinde **madde bazında doküman ve kayıt hazırlık
durumunu** takip eder. Standart ve bölüm bazında hazırlık yüzdesini hesaplar; denetime yetişmeyecek, sahipsiz ya da eski
kalmış maddeleri işaretleyip sorumlu ve terminli bir aksiyon listesi çıkarır. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Kural |
|---|---|
| Madde puanı | Hazır 1 · Kısmen / Güncellenmeli `--kismen` (%50) · Eksik 0 · Uygulanamaz hesaba girmez |
| Hazırlık % | Puan toplamı ÷ uygulanabilir madde × 100 (genel, standart, standart × bölüm) |
| Eski doküman | "Hazır" maddenin son güncelleme tarihi denetim gününde `--gecerlilik` (365) günden eskiyse Güncellenmeli sayılır |
| Aksiyon listesi | Hazır olmayan maddeler: önce kritikler, sonra termine göre; termin ile denetim arasındaki gün |

Durum metinden tanınır: Hazır / Tamam / Var · Kısmen / Taslak / Devam · Güncellenmeli / Revize · Eksik / Yok (boş = Eksik) ·
Uygulanamaz / Kapsam dışı. "Kritik" sütununa Evet yazılan maddeler (ör. standardın temel / KO gereklilikleri) ayrıca izlenir.

**Uyarılar:** hazır olmayan kritik madde · termini denetimden sonra (Yüksek) · termini geçmiş · termini ya da sorumlusu
yazılmamış · eski doküman (Orta) · denetim tarihi geçmiş (Bilgi).

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                                          # örnek veri (denetim 20.11.2026)
python main.py --liste denetim_listesi.xlsx --denetim 20.11.2026
python main.py --liste denetim_listesi.xlsx --denetim 20.11.2026 --kismen 25 --gecerlilik 180
```

Kendi listenizde `--denetim` zorunludur. Zorunlu sütunlar: Madde No ve Durum.

## Çıktı

`cikti/denetim_hazirlik.xlsx`: `Özet` (hazırlık %90 ve üstü yeşil, %70–90 sarı, altı kırmızı), `Aksiyon Listesi`,
`Tüm Maddeler` (girilen ve değerlendirilen durum), `Uyarılar`.

## Dikkat

- **Madde listesini standardın güncel sürümünden siz hazırlayın.** Paket herhangi bir standardın metnini içermez; örnek
  listedeki madde numaraları (S-1.1, R-3, M-2…) ve başlıklar kurgusaldır, gerçek standart maddeleriyle eşleşmez.
- **Hazırlık yüzdesi resmî denetim puanı değildir.** Sertifikasyon kuruluşlarının not / sınıflandırma kuralları farklıdır;
  bu araç yalnız iç hazırlık takibi içindir.
- **Doküman geçerlilik süresi** (varsayılan 365 gün) genel bir gözden geçirme varsayımıdır; standardınızın ve iç
  prosedürünüzün öngördüğü süreleri kullanın. Süresiz belgelerde son güncelleme sütununu boş bırakın.
- **Örnek veri:** Maddeler, sorumlular ve tarihler kurgusaldır.

## Testler

- Puanlama, uygulanamaz maddeler, eski dokümanın Güncellenmeli sayılması, standart / bölüm / genel hazırlık %, aksiyon
  sıralaması ve denetime kalan gün.
- Kritik, termin, sahipsiz aksiyon ve geçmiş denetim uyarıları; durum tanıma, okuma ve komut satırı.

```bash
python -m unittest discover -s tests
```

## Sorumluluk

Bu paketi indiren ve çalıştıran kişi aşağıdakileri kabul etmiş sayılır:

- **Tüm risk kullanıcıdadır.** Paket ücretsizdir ve "olduğu gibi" sunulur; doğruluk, eksiksizlik veya belirli bir amaca
  uygunluk konusunda hiçbir garanti verilmez. Paketi kendi bilgisayarınızda, kendi verinizle ve kendi kararınızla
  çalıştırırsınız. Hatalı veya eksik sonuç, kaçırılan süre, vergi / SGK cezası, veri kaybı, mali kayıp ve diğer her türlü
  doğrudan ya da dolaylı zarardan yazarlar ve katkıda bulunanlar sorumlu tutulamaz.
- **Sonuçları kontrol etmek sizin işinizdir.** Çıktılar karar desteği içindir. Beyan vermeden, ödeme yapmadan, sözleşme
  imzalamadan, personel / kredi / hasar / tedavi kararı almadan önce sonuçları yetkin bir kişi kontrol etmelidir.
- **Danışmanlık değildir.** Hukuki, mali, vergisel, tıbbi veya yatırım danışmanlığı yerine geçmez.
- **Mevzuat değişir.** Koddaki süre, oran, eşik ve kurallar yazıldığı tarihteki bilgiye dayanır ve eskiyebilir; özel
  durumları kapsamayabilir. Güncelliğini kontrol etmek kullanıcının sorumluluğundadır.
- **Örnek veri kurgusaldır.** Kişi, kurum, tutar ve tarihler gerçek değildir; kendi verinizle sonuçlar farklı olur.
- **Kişisel veri.** Verinizin KVKK, şirket politikaları ve sektör düzenlemelerine uygun işlenmesinden, saklanmasından ve
  paylaşılmasından siz sorumlusunuz.

Ayrıntılar: `LICENSE` (MIT) ve `SORUMLULUK_REDDI.md`.
