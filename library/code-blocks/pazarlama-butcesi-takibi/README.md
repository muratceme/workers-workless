# Pazarlama Bütçesi Takibi · Kod Bloğu

> Pazarlama › Pazarlama Müdürü · Gerisi Bende

Kanal ve kampanya bazında **bütçe, harcama, kalan tutar ve getiriyi** tek tabloda toplar; bütçesini aşan, dönemin
gerisinde ya da ilerisinde harcayan ve harcamasını karşılamayan kampanyaları işaretler. İnternete bağlanmaz.

## Nasıl hesaplar?

| Gösterge | Formül |
|---|---|
| Kalan bütçe, kullanım % | Bütçe − harcama; harcama ÷ bütçe × 100 |
| ROAS | Gelir ÷ geliri yazılmış harcama |
| ROI % | (Gelir − geliri yazılmış harcama) ÷ geliri yazılmış harcama × 100 |
| Dönüşüm başı maliyet | Dönüşümü yazılmış harcama ÷ dönüşüm sayısı |
| Beklenen kullanım % | Dönemin geçen gün sayısı ÷ toplam gün sayısı × 100 |
| Hız | Kullanım % − beklenen kullanım %; fark `--hiz` (15) puanı aşarsa "Hızlı" ya da "Yavaş" |
| Dönem sonu tahmini | Harcama × toplam gün ÷ geçen gün |

Geliri boş bırakılan harcama satırları ROAS ve ROI hesabına girmez; böylece geliri ölçülmeyen bir fuar, ölçülen
kanalların getirisini düşük göstermez.

**Örnek:** Bütçe 600.000, harcama 585.000 → kullanım **%97,50**. Dönemin %78,08'i geçmiş → fark 19,42 puan:
**Hızlı**. Gelir 1.755.000 → ROAS **3,00**, ROI **%200**.

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                            # örnek veri, dönem 2026, durum tarihi 12.10.2026
python main.py --butce butce.xlsx --harcamalar harcamalar.xlsx
python main.py --butce butce.xlsx --harcamalar harcamalar.xlsx --donem-baslangic 01.01.2026 --donem-bitis 31.12.2026 --hiz 10 --roas 2
```

Zorunlu sütunlar: bütçe dosyasında Kanal ve Bütçe; harcama dosyasında Kanal ve Harcama. Kampanya boşsa "(genel)"
yazılır. Dönem verilmezse harcama hızı hesaplanmaz.

## Çıktı

`cikti/pazarlama_butcesi_takibi.xlsx`:

- `Kampanyalar`: her kampanya için bütçe, harcama, kalan, kullanım, gelir, ROAS, ROI, dönüşüm, hız ve tahmin.
- `Kanallar`: aynı göstergeler kanal bazında ve toplam.
- `Aylar`: aylık harcama, gelir, ROAS ve dönüşüm.
- `Uyarılar`: bütçe aşımı, bütçesiz harcama, hızlı ve yavaş harcama, düşük ROAS, geliri yazılmamış kampanya.

## Dikkat

- **Gelir sizin atfettiğiniz tutardır.** Bir satışın hangi kampanyadan geldiğine kod karar vermez; reklam
  panelinden ya da CRM'den aldığınız değeri yazarsınız. Atıf yöntemi değişirse ROAS da değişir.
- **ROI brüt gelir üzerindendir.** Satılan malın maliyeti düşülmez; kâr bazlı getiri için gelir sütununa brüt kârı
  yazın. ROAS eşiği (`--roas`, varsayılan 1) yönetim tercihidir; kâr marjınıza göre belirleyin.
- **Harcama hızı doğrusal varsayar.** Bütçenin yıl içine eşit yayıldığı kabul edilir. Fuar, lansman gibi tek
  seferlik ya da mevsimlik kalemler bu yüzden "Hızlı" ya da "Yavaş" görünebilir; bunları ayrı yorumlayın.
- **Tutarlar tek para biriminde olmalı** ve KDV dahil ya da hariç olması iki dosyada aynı olmalıdır.
- **Kapsam dışı:** gösterim, tıklama ve tıklama oranı gibi reklam ölçütleri, müşteri yaşam boyu değeri, çok
  dokunuşlu atıf, taahhüt edilmiş ama henüz faturalanmamış harcamalar.
- **Örnek veri:** Kanallar, kampanyalar ve tutarlar kurgusaldır.

## Testler

- Kullanım, kalan, ROAS, ROI ve dönüşüm başı maliyet; geliri ya da dönüşümü boş satırların hesaba etkisi.
- Harcama hızı sınırları (tam eşikte, eşiğin altında), dönem başlamadan ve bittikten sonra, dönem dışı harcama.
- Bütçe aşımı, bütçesiz harcama, tekrarlanan bütçe satırı, kanal ve ay özeti, okuma ve komut satırı.

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
