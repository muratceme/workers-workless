# Banka Hareketlerinden Muhasebe Fişi Önerisi · Kod Bloğu

> Muhasebe ve Defter İşleri (Mali Müşavirlik) › Muhasebe Elemanı (Büro) · Workers / Workless

Banka ekstresindeki her hareket için **karşı hesabı bulur** ve **çift taraflı muhasebe fişi satırları** önerir.
Mali müşavirlik bürolarında mükellef ekstrelerinin elle tek tek işlenmesinin yerine geçer. İnternete bağlanmaz.

## Nasıl eşleştirir?

1. **Cari eşleşmesi:** Hareket açıklaması sırasıyla şu bilgilerle cari kart listesine bağlanır:
   - **IBAN**
   - **VKN/TCKN**
   - **Unvan:** "A.Ş., Ltd. Şti., San. ve Tic." gibi ekler atılır. Ayırt edici kelimelerin tamamı ya da en az 2'si
     ve %60'ı eşleşmelidir. İki cari eşit eşleşirse karar verilmez.

   Giriş müşteriden tahsilat (120 alacak), çıkış tedarikçiye ödeme (320 borç) sayılır. Cari kartta hesap kodu
   varsa o kullanılır. Müşteriye ödeme ve tedarikçiden tahsilat (olası iade) işaretlenir.
2. **Kural eşleşmesi:** `kurallar.csv` dosyasında anahtar kelime (`a | b` = a veya b), yön ve tutar aralığına göre
   hesap kodu verilir. İlk uyan kural uygulanır. Örnek kurallar: maaş → 335, SGK → 361, vergi → 360,
   kira → 336, banka masrafı/BSMV → 780, faiz geliri → 642, POS → 108.
3. **Manuel:** Eşleşmeyen hareketler fişe girmez, ayrı sayfada listelenir.

**Fiş:** Giriş → banka (102) **borç** / karşı hesap **alacak**; çıkış → karşı hesap **borç** / banka **alacak**.

> Örnek kurallar genel uygulamayı yansıtır. **Hesap kodlarını kendi hesap planınıza ve mali müşavirinizin
> tercihlerine göre düzenleyin.** KDV, stopaj ve BSMV ayrıştırması yapılmaz.

Ekstre okuma çekirdeği **Banka Mutabakatı** kod bloğuyla ortaktır (`ekstre_cekirdek.py`).

## Kurulum ve kullanım

```bash
pip install -r requirements.txt

python main.py                                                      # örnek 15 hareket, 5 cari, 8 kural
python main.py --ekstre ekstre.xlsx --cariler cari_kartlar.xlsx --kurallar kurallar.csv --banka-hesap 102.01.001 --fis-no 1250
```

## Çıktı

- `Hareket Eşleşmeleri`: karşı hesap, eşleşme yolu, güven (yüksek = IBAN/VKN), "Onay"
- `Fiş Aktarım`: genel biçim (Fiş No, Tarih, Hesap Kodu, Açıklama, Borç, Alacak); muhasebe programınızın içe
  aktarma şablonuna uyarlayın
- `Manuel`
- `Bilgi`

## Testler

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Öneriler taslaktır; muhasebeleştirmeden önce kontrol edilmelidir.
Bkz. `LICENSE` ve `SORUMLULUK_REDDI.md`.
