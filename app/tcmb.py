# -*- coding: utf-8 -*-
"""
TCMB Döviz Kuru Servisi
Merkez Bankası'ndan günlük döviz kurlarını çeker ve cache'ler.
Kullanım:
    from app.tcmb import kur_getir
    kur = kur_getir("EUR", "2026-07-15")  # → 38.5432 (döviz alış)
"""
import os, json, xml.etree.ElementTree as ET
from datetime import datetime, timedelta
from pathlib import Path
from urllib import request, error

# Cache dizini — tekrar tekrar TCMB'ye gitmesin
CACHE_DIR = Path(os.environ.get("FIS_DATA", "/data")) / "_tcmb_cache"


def _cache_yolu(tarih_str: str) -> Path:
    """Cache dosya yolu: _tcmb_cache/2026-07-15.json"""
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    return CACHE_DIR / f"{tarih_str}.json"


def _tcmb_url(tarih_str: str) -> str:
    """TCMB XML URL: https://www.tcmb.gov.tr/kurlar/YYMM/DDMMYYYY.xml"""
    y, m, d = tarih_str.split("-")
    return f"https://www.tcmb.gov.tr/kurlar/{y[2:]}{m}/{d}{m}{y}.xml"


def _xml_parse(xml_text: str) -> dict:
    """TCMB XML'den döviz kurlarını çıkar.
    Dönüş: {"USD": {"alis": 38.12, "satis": 38.25}, "EUR": {...}, ...}"""
    root = ET.fromstring(xml_text)
    kurlar = {}
    for cur in root.findall(".//Currency"):
        kod = cur.get("CurrencyCode", "")
        if not kod:
            continue
        alis = cur.findtext("ForexBuying", "").strip()
        satis = cur.findtext("ForexSelling", "").strip()
        if alis and satis:
            try:
                kurlar[kod] = {
                    "alis": float(alis.replace(",", ".")),
                    "satis": float(satis.replace(",", ".")),
                }
            except ValueError:
                pass
    return kurlar


def _gunluk_kurlar(tarih_str: str) -> dict:
    """Belirtilen tarih için TCMB kurlarını getirir.
    Önce cache'e bakar, yoksa TCMB'den çeker.
    Hafta sonu/tatil ise önceki iş gününü dener (3 gün geriye kadar)."""
    # Cache kontrolü
    cache = _cache_yolu(tarih_str)
    if cache.exists():
        try:
            return json.loads(cache.read_text(encoding="utf-8"))
        except Exception:
            pass

    # TCMB'den çek — hafta sonu/tatil olabilir, 5 gün geriye dene
    dt = datetime.strptime(tarih_str, "%Y-%m-%d")
    for geri in range(6):
        gun = dt - timedelta(days=geri)
        gun_str = gun.strftime("%Y-%m-%d")
        # Cache'te bu gün var mı?
        gc = _cache_yolu(gun_str)
        if gc.exists():
            try:
                kurlar = json.loads(gc.read_text(encoding="utf-8"))
                if kurlar:
                    # Orijinal tarih için de cache'le
                    if gun_str != tarih_str:
                        _cache_yolu(tarih_str).write_text(
                            json.dumps(kurlar, ensure_ascii=False), encoding="utf-8")
                    return kurlar
            except Exception:
                pass
        # TCMB'den çek
        url = _tcmb_url(gun_str)
        try:
            req = request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            resp = request.urlopen(req, timeout=10)
            xml_text = resp.read().decode("utf-8", errors="replace")
            kurlar = _xml_parse(xml_text)
            if kurlar:
                # Cache'le
                _cache_yolu(gun_str).write_text(
                    json.dumps(kurlar, ensure_ascii=False), encoding="utf-8")
                if gun_str != tarih_str:
                    _cache_yolu(tarih_str).write_text(
                        json.dumps(kurlar, ensure_ascii=False), encoding="utf-8")
                return kurlar
        except (error.URLError, error.HTTPError, Exception):
            continue

    return {}


def kur_getir(doviz_kodu: str, tarih_str: str, yon: str = "alis") -> float | None:
    """Belirtilen döviz ve tarih için TCMB kurunu döndürür.
    doviz_kodu: "USD", "EUR", "GBP" vs.
    tarih_str: "2026-07-15" (ISO format)
    yon: "alis" (döviz alış, varsayılan) veya "satis" (döviz satış)
    Dönüş: kur değeri (float) veya None (bulunamazsa)
    """
    doviz_kodu = doviz_kodu.upper().strip()
    kurlar = _gunluk_kurlar(tarih_str)
    if doviz_kodu in kurlar:
        return kurlar[doviz_kodu].get(yon)
    return None


def toplu_kur_getir(doviz_kodu: str, tarihler: list, yon: str = "alis") -> dict:
    """Birden fazla tarih için kur getirir.
    Dönüş: {"2026-07-15": 38.54, "2026-07-16": 38.62, ...}"""
    sonuc = {}
    for t in set(tarihler):
        k = kur_getir(doviz_kodu, t, yon)
        if k:
            sonuc[t] = k
    return sonuc
