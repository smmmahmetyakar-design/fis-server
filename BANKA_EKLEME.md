# Yeni Banka Nasıl Eklenir?

## Adım 1: data/banks_config.json Düzenle

```bash
nano data/banks_config.json
```

Yeni banka ekle (örnek - Deutsche Bank):

```json
{
  "id": "deutsche",
  "name": "Deutsche Bank",
  "detect_keywords": ["DEUTSCHE BANK", "DEUTSCHE"],
  "separator": "|",
  "columns": ["date", "desc", "amount", "balance"],
  "date_format": "DD.MM.YYYY",
  "amount_format": "turkish"
}
```

## Adım 2: Config'i Yeniden Yükle

### Seçenek A: Container Restart (Downtime)
```bash
docker restart fis-otomasyon
```

### Seçenek B: Reload Endpoint (Downtime Yok)
```bash
curl -X POST http://localhost:8091/api/banks/reload-config
```

## Adım 3: Test Et

```bash
curl http://localhost:8091/api/banks/supported
```

## Banka Config Alanları

| Alan | Zorunlu | Açıklama |
|------|---------|----------|
| `id` | ✅ | Benzersiz ID (lowercase) |
| `name` | ✅ | Banka adı |
| `detect_keywords` | ✅ | Dosyada aramak için anahtar kelimeler |
| `separator` | ✅ | Sütun ayıracı (\t = tab, \| = pipe) |
| `columns` | ✅ | Sütun sırası: date, desc, amount, balance |
| `date_format` | ❌ | DD.MM.YYYY (varsayılan) |
| `amount_format` | ❌ | turkish (varsayılan) |

## Örnek: Yeni Banka Eklemesi

### Scenario: ING Bank Desteği
```json
{
  "id": "ing",
  "name": "ING Bank",
  "detect_keywords": ["ING BANK", "INGBANK"],
  "separator": "\t",
  "columns": ["date", "desc", "amount", "balance"]
}
```

### Scenario: Dış Ticaret Bankası
```json
{
  "id": "dtb",
  "name": "Dış Ticaret Bankası",
  "detect_keywords": ["DIŞ TİCARET", "DTB"],
  "separator": "|",
  "columns": ["date", "desc", "amount", "balance"]
}
```

## Kod Değişikliği Yok ✅

Artık yeni banka eklemek için:
- ❌ Python kodu değiştirme
- ❌ Docker rebuild
- ✅ JSON dosyası güncelle
- ✅ Container reload (veya restart)

---
**Avantajlar:**
- Hızlı banka ekleme
- Restart olmadan reload
- Tüm bankalar merkezi yerde
- Geri alma kolay (git)
