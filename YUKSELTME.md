# Fiş Aktarım Aracı - ML Öğrenen Sistem Yükseltmesi

**Sürüm:** v2.0  
**Tarih:** 2026-09-11  
**Durum:** ✅ Tamamlandı

## Neler Eklendi?

### 1. 4-Katmanlı Makine Öğrenmesi
- Tam eşleştirme (100%)
- Bulanık eşleştirme (60-85%)
- Tedarikçi tipi sınıflandırması (65%)
- Türkçe muhasebe kuralları (75-90%)

### 2. 7 Türk Bankası Desteği
İş, Vakıf, YKB, Garanti, Ziraat, Akbank, Deniz

### 3. 7 Yeni FastAPI Endpoint'i
- GET /api/banks/supported
- POST /api/banks/parse
- POST /api/learning/suggest-account
- POST /api/learning/record-feedback
- POST /api/learning/train-from-history
- GET /api/learning/export-mappings
- POST /api/learning/batch-suggest

### 4. Web UI
http://100.74.86.128:8091

### 5. Aktif Öğrenme
Geri bildirim → Model güncellenir → Daha doğru öneriler

## Test
✅ Tüm endpoint'ler çalışıyor
✅ Öğrenme sistemi test edildi
✅ 7 banka formatı destekleniyor
✅ Docker image dağıtıldı
