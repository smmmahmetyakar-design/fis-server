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
