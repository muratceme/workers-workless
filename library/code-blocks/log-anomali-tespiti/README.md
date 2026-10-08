# Log Anomali Tespiti · Kod Bloğu

> Bilgi Teknolojileri › Sistem Uzmanı · Workers / Workless

Sunucu ve uygulama loglarında **olağandışı hata, oturum ve erişim örüntülerini** bulur ve önem sırasıyla
özetler. Log biçimi satırlardan otomatik tanınır. İnternete bağlanmaz; logları hiçbir yere göndermez.

## Ne bulur?

| Log | Bulgu | Önem |
|---|---|---|
| auth.log / secure (sshd) | Bir IP'den pencere içinde çok sayıda başarısız giriş; var olmayan kullanıcı adlarıyla deneme | Yüksek |
| | Denemelerin **ardından aynı IP'den başarılı giriş** | Kritik |
| | root hesabına parolayla giriş | Yüksek |
| | Kullanıcının logda ilk kez görülen IP'den girişi, mesai dışı giriş | Bilgi |
| | Kullanıcı başına 3+ sudo/su parola hatası | Dikkat |
| Web erişim (Apache/Nginx) | Bir IP'den yoğun 401/403; ardından aynı yola başarılı yanıt | Yüksek / Kritik |
| | Çok sayıda farklı yolda 404: dizin/zafiyet taraması | Yüksek |
| | Saldırı imzaları: SQL enjeksiyonu, XSS, dizin aşma, Log4Shell, komut çalıştırma, .env/.git/phpmyadmin araması | Dikkat; uygulamaya ulaşmışsa (2xx/5xx) Yüksek |
| | Dakikada aşırı istek | Dikkat |
| | Mesai dışında başarılı yönetim paneli isteği (/admin, /yonetim…) | Dikkat |
| Web ve uygulama | 5xx ve uygulama hatası sayısında **ani artış** | Yüksek |
| Uygulama | Kısa sürede ortaya çıkan yeni hata türü | Dikkat |

- **Ani artış:** Log süresi dilimlere bölünür (`--pencere`, varsayılan 10 dk). Bir dilim şu iki koşulu birlikte
  sağlarsa işaretlenir: tüm dilimlerin **medyanı + 4 × (1,4826 × MAD)** eşiğini aşmak ve en az `--min-artis`
  olaya sahip olmak. Ortalama yerine medyan kullanıldığı için tek bir tepe eşiği bozmaz.
- **Hata imzası:** Uygulama hata mesajlarındaki sayılar, kimlikler, tırnaklı değerler ve adresler
  normalleştirilerek gruplanır ("Timeout calling <url:pos…> after <n> ms"). Logun geneline yayılmayıp en fazla
  bir saat içinde 5+ kez görülen imza "ani ortaya çıkan hata türü" sayılır.

## Kullanım

```bash
pip install -r requirements.txt
python main.py                                                       # örnek loglar (07.10.2026)
python main.py --girdi /var/log/nginx/access.log /var/log/auth.log uygulama.log --yil 2026
python main.py --girdi ./loglar --pencere 5 --esik-giris 8 --esik-404 30 --mesai 08-18
```

**Desteklenen biçimler:**
- **Web erişim:** Apache/Nginx `combined` ve `common` biçimleri.
- **Syslog (auth.log):** Klasik `Oct  7 03:00:00 host sshd[123]: …` ya da ISO zaman damgalı satırlar.
  - Klasik syslog satırında yıl yazmaz; yılı `--yil` ile verin.
- **Uygulama:** `2026-10-07 14:30:01,123 ERROR [modül] mesaj` (Log4j, Logback, Python logging benzeri).
- **Sıkıştırılmış dosyalar:** `.gz` uzantılı dosyalar da okunur.
- **Okunamayan satırlar:** Tanınmayan satırlar sayılır ve `Özet` sayfasında gösterilir.

## Çıktı

`log_anomali_raporu.xlsx`:
- `Özet`: dosyalar, çözülen satır sayıları, önem dağılımı.
- `Bulgular`: önem sırasıyla, örnek log satırı ve "İnceleme Notu" sütunuyla.
- `Zaman Çizelgesi`: dilim bazında istek, 4xx, 5xx, başarısız giriş ve uygulama hatası sayıları, grafikli.
- `IP Özeti`.
- `Hata İmzaları`.

## Dikkat

- **Doğrulama:** Bulgular **olası** anomalilerdir. Olay müdahalesine geçmeden önce ilgili sistemlerde
  doğrulayın.
- **Eşikler:** Eşik değerleri örnektir; kendi trafiğinize göre ayarlayın. Yük dengeleyici veya proxy arkasındaki
  sunucularda gerçek istemci IP'si logda olmayabilir.
- **Kişisel veri:** Loglar IP adresi ve kullanıcı adı gibi kişisel veriler içerebilir. Raporu yetkisiz kişilerle
  paylaşmayın; saklama süresini kurum politikanıza göre belirleyin.
- **Örnek veri:** Örnek loglar kurgusaldır. IP'ler belgeleme için ayrılmış bloklardandır (RFC 5737:
  192.0.2.0/24, 198.51.100.0/24, 203.0.113.0/24).

## Testler

```bash
python -m unittest discover -s tests
```

---
Bu paket "olduğu gibi" sunulur. Güvenlik izleme ve olay müdahale süreçlerinin yerini almaz. Bkz. `LICENSE` ve
`SORUMLULUK_REDDI.md`.
