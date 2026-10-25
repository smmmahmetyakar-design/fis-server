# Fiş Aktarım Aracı — Kurulum

Bağımsız bir uygulamadır (masraf-server'dan ayrı, port 8091).

## Docker ile (önerilen — OCR dahil)
```
docker build -t fis-otomasyon .
docker run -d --name fis-otomasyon --restart unless-stopped \
  -p 8091:8091 -v /path/to/fis-server/data:/data fis-otomasyon
```
Tarayıcı: http://localhost:8091

## Kullanım
1. Firma oluştur.
2. Mizan yükle (tüm hesaplar).
3. Kural dosyası yükle (SEZGIN_OZBAY gibi) — banka/fatura/çek için ayrı.
4. Belge yükle (Excel/PDF/resim) — ilgili sekmeye.
5. "İşle" ile önizle; yanlış eşleşmeleri "Düzenle" ile düzelt (öğrenilir).
6. Excel (Fiş Aktarım Şablonu) veya XML olarak indir.

OCR: PDF/resim Türkçe okunur (tesseract-ocr-tur).

## Fatura fişi kuralları

Fiş kontrol aracının (fis-kontrol-server) Luca fiş listelerinde bulduğu kayıt hatalarının
üretim tarafındaki karşılığı. Kod: `app/fatura_kural.py`.

**Genel kurallar** — her firmada, ayar gerekmeden:
- Reddedilen (Uygulama Yanıtı RED, PDF'te RET damgası) ve iptal edilen fatura fişe girmez.
- Aynı e-Arşiv numarası farklı düzenleyicilerde olabilir: liste tekilleştirmesi ve PDF eşleştirmesi
  numara + unvan (PDF'lerde ETTN) ile yapılır.
- KDV hesabının adındaki oran faturanın oranını tutmalı; tevkifat/sorumlu KDV hesapları normal
  KDV satırına yazılmaz.
- Alış tevkifatı: 360 hesabı tevkifat koduna (606 işgücü, 612 temizlik, 624 taşıma …) ve oranına
  göre adından seçilir; uygun hesap yoksa satır KONTROL alır.
- Satış tevkifatı: adında başka oran yazan (ör. "2/10") 391 / 600 hesabı kullanılmaz ya da KONTROL alır.
- Cari hesabın mizan adı fatura unvanını tutmuyorsa KONTROL.
- İade: alıştan iade (satış listesinde) satıcı carisi borç / asıl alış hesabı + 391 alacak;
  satıştan iade (alış listesinde) 610 + 191 borç / alıcı carisi alacak.

**Firma kuralları** — isteğe bağlı, Alış/Satış Fatura sekmesindeki "Fatura kuralları (bu firma)"
kutusundan düzenlenir; sunucuda `data/<firma>/fatura_kurallari.yaml` olarak durur (repoya girmez).
Kalem açıklaması → hesap, tevkifat oranı → satış/KDV hesabı, tevkifatlı alışın KDV hesabı, kurala
uymayan kalem için hesap ve uyarı, iade hesapları. Şablon: `app/fatura_kurallari_ornek.yaml`.
Öncelik: kullanıcının kalem için birebir öğrettiği hesap > firma kuralı > benzer öğrenme / geçmiş
fiş / yapay zekâ.

Test: `python tests/test_fatura_kural.py`
