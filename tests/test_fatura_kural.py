"""Fatura kuralları — müşteri verisi gerektirmeyen hızlı denetim:  python tests/test_fatura_kural.py"""
import sys
import tempfile
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app import fatura_kural, isleyici  # noqa: E402
from app.kurallar import KuralMotoru  # noqa: E402

MIZAN = [
    ("153.01.001", "TİCARİ MALLAR"), ("153.01.002", "TAŞ TOZU"),
    ("191.01.01", "SATIŞTAN İADE EDİLEN KDV"), ("191.02.10", "%10 İNDİRİLECEK KDV"), ("191.02.20", "%20 İNDİRİLECEK KDV"),
    ("320.01.001", "ALFA MADENCİLİK LTD"), ("320.01.002", "BETA NAKLİYE A.Ş."), ("320.01.003", "GAMA TEMİZLİK"),
    ("120.01.001", "DELTA YAPI A.Ş."), ("120.01.002", "EPSİLON İNŞAAT"),
    ("360.01.001", "9/10 TEMİZLİK İŞLERİ TEVK."), ("360.01.002", "9/10 İŞGÜCÜ TEMİNİ TEVK."),
    ("391.01.10", "%10 HESAPLANAN KDV"), ("391.01.20", "%20 HESAPLANAN KDV"), ("391.04.01", "2/10 TEVKİFATLI KDV"),
    ("600.01", "NAKLİYE GELİRLERİ"), ("600.02", "MAL SATIŞI"), ("600.04", "2/10 TEVKİFATLI SATIŞLAR"),
    ("610.01", "SATIŞTAN İADELER"),
    ("740.01.001", "NAKLİYE GİDERLERİ"), ("740.01.002", "HAFRİYAT GİDERLERİ"),
]
# geçmiş fişler: (fiş no, hesap, detay, borç, alacak)
GECMIS = [
    ("1", "153.01.001", "ALFA MADENCİLİK LTD", 100, 0), ("1", "191.02.20", "ALFA MADENCİLİK LTD", 20, 0),
    ("1", "320.01.001", "ALFA MADENCİLİK LTD", 0, 120),
    ("2", "740.01.001", "BETA NAKLİYE A.Ş.", 100, 0), ("2", "191.02.20", "BETA NAKLİYE A.Ş.", 20, 0),
    ("2", "320.01.002", "BETA NAKLİYE A.Ş.", 0, 120),
    ("3", "740.01.002", "GAMA TEMİZLİK", 100, 0), ("3", "191.02.20", "GAMA TEMİZLİK", 20, 0),
    ("3", "360.01.001", "GAMA TEMİZLİK", 0, 18), ("3", "320.01.003", "GAMA TEMİZLİK", 0, 102),
    ("4", "120.01.001", "DELTA YAPI A.Ş.", 102, 0), ("4", "600.04", "DELTA YAPI A.Ş.", 0, 100),
    ("4", "391.04.01", "DELTA YAPI A.Ş.", 0, 16),
    # ZETA'nın cari hesabı yanlışlıkla Epsilon'unki
    ("5", "120.01.002", "ZETA GIDA LTD", 120, 0), ("5", "600.02", "ZETA GIDA LTD", 0, 100),
    ("5", "391.01.20", "ZETA GIDA LTD", 0, 20),
]
BASLIK = ["Fatura No", "Fatura Tarihi", "{ad}", "Senaryo", "Tür", "Toplam Tutar", "KDV Toplamı",
          "KDV %10", "KDV %20", "KDV %10 Matrah", "KDV %20 Matrah", "Tevkifat Toplamı", "İlk Tevkifat Kodu",
          "Uygulama Yanıtı"]


def _km(d: Path) -> KuralMotoru:
    wb = openpyxl.Workbook(); ws = wb.active
    for k, a in MIZAN:
        ws.append([k, a])
    wb.save(d / "mizan.xlsx")
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Fiş Aktarım Şablonu"
    ws.append(["FİŞ NO", "FİŞ TARİHİ", "FİŞ AÇIKLAMA", "HESAP KODU", "DETAY AÇIKLAMA", "BORÇ", "ALACAK"])
    for no, h, det, b, a in GECMIS:
        ws.append([no, "01.06.2026", det, h, det, b, a])
    wb.save(d / "gecmis.xlsx")
    return KuralMotoru(d / "mizan.xlsx", None, d / "ogr.json", d / "gecmis.xlsx")


def _liste(ad, satirlar):
    return [{"dosya": "liste.xlsx", "tablolar": [[[c.format(ad=ad) for c in BASLIK]] + satirlar]}]


def _fis(fisler, no, cari=None):
    return [s for s in fisler if s["evrak_no"] == no and (cari is None or cari in s["detay"])]


def test_alis():
    d = Path(tempfile.mkdtemp()); km = _km(d)
    #            no                  tarih         taraf                senaryo  tür     toplam   kdv    k10   k20   m10   m20    tevk  kod   yanıt
    satirlar = [
        ["GIB2026000000005", "01.07.2026", "ALFA MADENCİLİK LTD", "EARSIV", "SATIS", 1200, 200, 0, 200, 0, 1000, 0, "", ""],
        ["GIB2026000000005", "02.07.2026", "BETA NAKLİYE A.Ş.", "EARSIV", "SATIS", 2400, 400, 0, 400, 0, 2000, 0, "", ""],
        ["AAA2026000000001", "03.07.2026", "ALFA MADENCİLİK LTD", "TICARI", "SATIS", 1100, 100, 100, 0, 1000, 0, 0, "", ""],
        ["AAA2026000000002", "04.07.2026", "ALFA MADENCİLİK LTD", "TICARI", "SATIS", 1200, 200, 0, 200, 0, 1000, 0, "", "RED"],
        ["GTM2026000000001", "05.07.2026", "GAMA TEMİZLİK", "TICARI", "TEVKIFAT", 1020, 200, 0, 200, 0, 1000, 180, "606", ""],
    ]
    fisler, uy = isleyici.isle_fatura(_liste("Gönderici Adı", satirlar), km, 1, yon="alis", gider_ogrenme={})
    # aynı e-Arşiv numarası, iki farklı satıcı: ikisi de fişe girer
    assert _fis(fisler, "GIB2026000000005", "ALFA") and _fis(fisler, "GIB2026000000005", "BETA"), uy
    # reddedilen fatura fişe girmez
    assert not _fis(fisler, "AAA2026000000002") and any("reddedilmiş" in u for u in uy)
    # %10 KDV, carinin geçmişteki %20 hesabına değil %10 hesabına
    assert {s["hesap"] for s in _fis(fisler, "AAA2026000000001") if s["rol"] == "kdv"} == {"191.02.10"}
    # 606 (işgücü) tevkifatı temizlik hesabına değil işgücü hesabına
    assert {s["hesap"] for s in _fis(fisler, "GTM2026000000001") if s["rol"] == "tevkifat"} == {"360.01.002"}


def test_satis_ve_iade():
    d = Path(tempfile.mkdtemp()); km = _km(d)
    satirlar = [
        # 9/10 tevkifat: geçmişteki 2/10 hesaplarına yazılmaz
        ["DEN2026000000001", "01.07.2026", "DELTA YAPI A.Ş.", "TICARI", "TEVKIFAT", 1020, 200, 0, 200, 0, 1000, 180, "606", ""],
        # alıştan iade (satış listesinde İADE): satıcının 320'si borç, asıl alış hesabı alacak
        ["DEN2026000000002", "02.07.2026", "BETA NAKLİYE A.Ş.", "TICARI", "IADE", 600, 100, 0, 100, 0, 500, 0, "", ""],
        # cari hesap adı unvanı tutmuyor
        ["DEN2026000000003", "03.07.2026", "ZETA GIDA LTD", "TICARI", "SATIS", 1200, 200, 0, 200, 0, 1000, 0, "", ""],
    ]
    fisler, uy = isleyici.isle_fatura(_liste("Alıcı Adı", satirlar), km, 1, yon="satis", gider_ogrenme={})
    t = _fis(fisler, "DEN2026000000001")
    assert {s["hesap"] for s in t if s["rol"] == "tevkifat"} == {"391.01.20"}, t
    assert any("tevkifat_tur" in (s.get("kontrol") or []) for s in t if s["rol"] == "tevkifat")
    assert any("tevkifat_tur" in (s.get("kontrol") or []) for s in t if s["rol"] == "gider")   # 600.04 "2/10"
    i = {s["hesap"]: (s["borc"], s["alacak"]) for s in _fis(fisler, "DEN2026000000002")}
    assert i["320.01.002"] == (600, 0) and i["740.01.001"] == (0, 500) and i["391.01.20"] == (0, 100), i
    z = [s for s in _fis(fisler, "DEN2026000000003") if s["rol"] == "cari"]
    assert z and "cari_ad" in (z[0].get("kontrol") or [])


def test_firma_kurallari():
    d = Path(tempfile.mkdtemp()); km = _km(d)
    metin = """
alis:
  kalem:
    - {anahtar: 'NAKL[İI]YE', hesap: "740.01.001"}
    - {anahtar: 'TAŞ ?TOZU', hesap: "153.01.002"}
    - {anahtar: 'KUM|MICIR', hesap: "153.01.001"}
satis:
  tevkifat:
    "2/10": {satis: "600.04", kdv: "391.04.01"}
"""
    k, hatalar = fatura_kural.dogrula(metin, km.hesaplar)
    assert not hatalar, hatalar
    satirlar = [["ALF2026000000009", "01.07.2026", "ALFA MADENCİLİK LTD", "TICARI", "SATIS", 3600, 600, 0, 600, 0, 3000, 0, "", ""]]
    pdf = [{"fatura_no": "ALF2026000000009", "cari_ad": "ALFA MADENCİLİK LTD", "toplam": 3600, "dosya": "a.pdf",
            "kalemler": [{"oran": 20, "matrah": 3000, "kdv": 600}], "eksik": [],
            "kalem_detay": [{"a": "MICIR", "t": 1500, "o": 20}, {"a": "Açıklama:", "t": None},
                            {"a": "TAŞ TOZU", "t": 1000, "o": 20}, {"a": "NAKLİYE BEDELİ", "t": 500, "o": 20}]}]
    fisler, uy = isleyici.isle_fatura(_liste("Gönderici Adı", satirlar), km, 1, yon="alis", pdf_faturalar=pdf,
                                      gider_ogrenme={}, kurallar=k)
    g = {s["hesap"]: s["borc"] for s in fisler if s["rol"] == "gider"}
    assert g == {"153.01.001": 1500, "153.01.002": 1000, "740.01.001": 500}, g
    assert fatura_kural.anahtar_okunur("NAKL[İI]YE|TAŞ ?TOZU|\\bKUM\\b") == "NAKLİYE, TAŞ TOZU, KUM"


if __name__ == "__main__":
    for ad, fn in list(globals().items()):
        if ad.startswith("test_"):
            fn()
            print("✓", ad)
