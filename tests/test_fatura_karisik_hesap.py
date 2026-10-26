"""Aynı müşteriye farklı gelir hesaplarıyla (mal satışı / hizmet) kesilen faturalar:
python tests/test_fatura_karisik_hesap.py"""
import sys
import tempfile
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app import isleyici  # noqa: E402
from app.kurallar import KuralMotoru  # noqa: E402

MIZAN = [("120.01.001", "KAPPA PLASTİK LTD"), ("391.01.20", "%20 HESAPLANAN KDV"),
         ("600.05", "TİCARİ MAL SATIŞI"), ("600.06", "HAFRİYAT İŞLERİ")]
# KAPPA'ya 3 hafriyat faturası (600.06), 1 mal satışı faturası (600.05: XYZ2026000000004)
GECMIS = []
for i, h in enumerate(("600.06", "600.06", "600.06", "600.05"), 1):
    det = f"XYZ202600000000{i} KAPPA PLASTİK LTD"
    GECMIS += [(str(i), "120.01.001", det, 120, 0), (str(i), h, det, 0, 100), (str(i), "391.01.20", det, 0, 20)]
BASLIK = ["Fatura No", "Fatura Tarihi", "Alıcı Adı", "Senaryo", "Tür", "Toplam Tutar", "KDV Toplamı",
          "KDV %20", "KDV %20 Matrah"]


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


def _gelir(fisler, no):
    return [s for s in fisler if s["evrak_no"] == no and s["rol"] == "gider"]


def test_karisik_musteri():
    d = Path(tempfile.mkdtemp()); km = _km(d)
    es = km.fatura_gecmis_eslestir("KAPPA PLASTİK LTD", "satis")
    assert es["ana"] == "600.06" and es["ana_dagilim"] == {"600.06": 3, "600.05": 1}, es
    satirlar = [
        # zaten kayıtlı mal satışı faturası: carinin en sık hesabı (600.06) değil, kendi fişindeki 600.05
        ["XYZ2026000000004", "05.07.2026", "KAPPA PLASTİK LTD", "TICARI", "SATIS", 120, 20, 20, 100],
        # yeni fatura: geçmiş hesabı önerilir ama müşteri karışık olduğu için KONTROL
        ["XYZ2026000000009", "06.07.2026", "KAPPA PLASTİK LTD", "TICARI", "SATIS", 240, 40, 40, 200],
    ]
    fisler, uy = isleyici.isle_fatura([{"dosya": "liste.xlsx", "tablolar": [[BASLIK] + satirlar]}], km, 10,
                                      yon="satis", gider_ogrenme={})
    kayitli = _gelir(fisler, "XYZ2026000000004")
    assert [s["hesap"] for s in kayitli] == ["600.05"] and "kayıtlı" in kayitli[0].get("not", ""), kayitli
    yeni = _gelir(fisler, "XYZ2026000000009")
    assert [s["hesap"] for s in yeni] == ["600.06"] and "gider_karisik" in (yeni[0].get("kontrol") or []), yeni
    assert "600.05 (1)" in yeni[0].get("not", ""), yeni


def test_kayitli_fatura_ogrenmeden_once():
    """Kalem 'BALANS TAŞ' bir kez 600.06'ya öğretilmiş; ama fatura fiş listesinde 600.05 ile kayıtlı:
    kayıt kesindir, öğrenme onu ezmez. Kayıtlı olmayan yeni faturada öğrenme geçerli."""
    from app.kurallar import norm
    d = Path(tempfile.mkdtemp()); km = _km(d)
    satirlar = [["XYZ2026000000004", "05.07.2026", "KAPPA PLASTİK LTD", "TICARI", "SATIS", 120, 20, 20, 100],
                ["XYZ2026000000009", "06.07.2026", "KAPPA PLASTİK LTD", "TICARI", "SATIS", 240, 40, 40, 200]]
    pdfler = [{"fatura_no": no, "cari_ad": "KAPPA PLASTİK LTD", "dosya": f"{no}.pdf", "kalem_aciklamalari": ["BALANS TAŞ"],
               "kalem_detay": [{"a": "BALANS TAŞ", "t": t}]} for no, t in (("XYZ2026000000004", 100.0),
                                                                          ("XYZ2026000000009", 200.0))]
    ogr = {f"kalem|satis|{norm('BALANS TAŞ')}": "600.06"}
    fisler, uy = isleyici.isle_fatura([{"dosya": "liste.xlsx", "tablolar": [[BASLIK] + satirlar]}], km, 10,
                                      yon="satis", pdf_faturalar=pdfler, gider_ogrenme=ogr)
    assert [s["hesap"] for s in _gelir(fisler, "XYZ2026000000004")] == ["600.05"], _gelir(fisler, "XYZ2026000000004")
    assert [s["hesap"] for s in _gelir(fisler, "XYZ2026000000009")] == ["600.06"], _gelir(fisler, "XYZ2026000000009")


if __name__ == "__main__":
    test_karisik_musteri()
    test_kayitli_fatura_ogrenmeden_once()
    print("2/2 kontrol gecti")
