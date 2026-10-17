# Kredi Kullandırım Ön Kontrolü · Kod Bloğu

> Bankacılık Operasyonları › Kredi Operasyon Uzmanı · Gerisi Bende

Kullandırım taleplerini **tahsis onayındaki şartlar, teminat tesis durumu ve sözleşme bilgileriyle** karşılaştırır;
kullandırıma engel olan eksikleri listeler. İnternete bağlanmaz.

## Nasıl hesaplar?

| Kontrol | Kural | Düzey |
|---|---|---|
| Onay | Onay kaydı yok; müşteri numarası uyuşmuyor; kullandırım tarihi geçerlilik bitişinden sonra | Engel |
| Onay | Geçerlilik `--yaklasan` gün (varsayılan 15) içinde doluyor | Kontrol |
| Limit | Mevcut risk + bu listede aynı onaya bağlı önceki talepler + talep > onaylı limit | Engel |
| Para birimi, ürün | Onaydakinden farklı (`*` veya boş ürün = tüm ürünler) | Engel |
| Vade | Azami vadeyi aşıyor | Engel |
| Fiyatlama | Faiz asgari oranın altında | Engel |
| Sözleşme | İmzalı değil; sözleşme tutarı (mevcut risk + talep) tutarını karşılamıyor | Engel |
| Sözleşme | Onay kaydında sözleşme bilgisi yok | Kontrol |
| Şart | "Kullandırım öncesi" şart tamamlanmamış | Engel |
| Teminat | Tesis edilen tutar istenen tutarın altında | Engel |
| Şart | Tamamlanma tarihi kullandırım tarihinden sonra | Kontrol |
| Şart | "Kullandırım sonrası" şart henüz tamamlanmamış | Bilgi (takip listesi) |
| Tarih | Kullandırım tarihi hafta sonu veya geçmiş | Kontrol |

- **Sıralı limit düşümü:** Talepler dosyadaki sırayla işlenir. Engeli olmayan her talep, aynı onaya bağlı sonraki
  talepler için kullanılabilir limitten düşülür.
- Sınır değerler geçerlidir: talep limite eşitse, vade azami vadeye eşitse, faiz asgari faize eşitse engel yoktur.
- **Sonuç:** Engel varsa "Engel var"; yalnız kontrol bulgusu varsa "Kontrol edin"; yoksa "Kullandırılabilir".

## Kurulum

```bash
pip install -r requirements.txt
```

## Kullanım

```bash
python main.py                                    # örnek veri
python main.py --talepler talepler.xlsx --onaylar onaylar.xlsx --sartlar sartlar.xlsx --tarih 09.10.2026
```

| Dosya | Sütunlar |
|---|---|
| Talepler | Talep No, Müşteri No, Müşteri, Şube, Onay No, Ürün, Tutar, Para Birimi, Vade (ay), Faiz Oranı, Kullandırım Tarihi |
| Onaylar | Onay No, Müşteri No, Onaylı Limit, Para Birimi, Ürünler, Azami Vade, Asgari Faiz, Geçerlilik Bitiş, Mevcut Risk, Sözleşme Tutarı, Sözleşme Durumu |
| Şartlar | Onay No, Şart, Tür, Zamanlama (Kullandırım öncesi / sonrası), Gereken Tutar, Tesis Edilen Tutar, Durum, Tamamlanma Tarihi |

"Durum" sütununda `Tamamlandı`, `Tesis edildi`, `Alındı`, `İmzalı` veya `Evet` ile başlayan değerler tamamlanmış sayılır.

## Çıktı

`cikti/kullandirim_on_kontrol.xlsx`:

- `Talep Sonuçları`: talep bazında sonuç, kullanılabilir limit, engel ve kontrol sayısı; boş "Onaylayan" sütunu.
- `Bulgular`: her bulgunun ayrıntısı ve yapılacak iş; boş "Giderildi" sütunu.
- `Şart Takibi`: kullandırım sonrası tamamlanacak şartlar; boş sorumlu ve hedef tarih sütunları.
- `Uyarılar`: okunamayan satırlar.

## Dikkat

- **Kayıt karşılaştırmasıdır.** Kod teminatın hukuken geçerli tesis edildiğini, ipotek derecesini, sözleşme
  imzalarının yetkili kişilerce atıldığını veya belgelerin aslını doğrulamaz.
- **Kullandırım kararı değildir.** "Kullandırılabilir" yalnız girilen kayıtlarda engel bulunmadığını gösterir; nihai
  yetki bankanızın onay akışındadır.
- **Mevcut risk** dosyadaki değerdir. Gün içinde yapılan diğer kullandırımlar ve geri ödemeler dosyaya yansımadıysa
  limit hesabı eksik olur.
- **Çapraz para birimi:** Onay ve talep para birimi farklıysa engel verilir; kur çevrimi yapılmaz.
- **Grup limitleri**, yasal kredi sınırları ve risk grubu kontrolleri kapsam dışıdır.
- **Örnek veri:** Firmalar, onaylar ve şartlar kurgusaldır.

## Testler

- Temiz talep; limit ve sıralı düşüm; sınır değerler.
- Onay koşulları (müşteri, geçerlilik, para birimi, ürün, vade, faiz), sözleşme kontrolleri.
- Şartlar: tamamlanmamış ön koşul, eksik teminat tutarı, kullandırım sonrası şart ve takip listesi.
- Okuma hataları ve örnek veri uçtan uca.

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
