"""Çek listesi -> fiş: çek numarası açıklamada (Luca biçimi 'ÇEKNO-BANKA-VADEVDLİ-KİŞİ').
python tests/test_cek.py — sahte veri."""
import sys
import tempfile
from datetime import datetime
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app import isleyici  # noqa: E402
from app.kurallar import KuralMotoru  # noqa: E402

MIZAN = [("101.01.001", "TL ÇEKLER"), ("120.01.001", "KAPPA PLASTİK LTD"), ("320.01.001", "OMEGA DEMİR AŞ")]
BASLIK = ["NO", "ALIM TARİH", "KİMDEN ALINDI", "BANKA", "ŞUBE", "VADE", "ÇEK NO", "TUTARI", "KİME VERİLDİ",
          "ÇIKIŞ TARİH"]


def _km(d: Path) -> KuralMotoru:
    wb = openpyxl.Workbook(); ws = wb.active
    for k, a in MIZAN:
        ws.append([k, a])
    wb.save(d / "mizan.xlsx")
    return KuralMotoru(d / "mizan.xlsx", None, d / "ogr.json", None)


def test_cek_no_aciklamada():
    km = _km(Path(tempfile.mkdtemp()))
    satirlar = [
        [1, datetime(2026, 4, 28), "KAPPA PLASTİK LTD", "KUVEYT TURK", "MERKEZ", datetime(2026, 8, 30), 23758.0,
         350000, "OMEGA DEMİR AŞ", "30.04.2026"],
        [2, "05.05.2026", "KAPPA PLASTİK LTD", "HALKBANK", "", "11.09.2026", "5341489", 1000, "KASA", None],
    ]
    fisler, uy = isleyici.isle_cek([{"dosya": "cek.xlsx", "tablolar": [[BASLIK] + satirlar]}], km, 1)
    g = [s for s in fisler if s["evrak_no"] == "23758"]
    assert len(g) == 4, fisler
    giris, cikis = g[:2], g[2:]
    assert all(s["detay"] == "23758-KUVEYT TURK-30.08.2026VDLİ-KAPPAPLASTİKLTD" for s in giris), giris
    assert all(s["fis_aciklama"] == s["detay"] for s in g), g
    assert all(s["detay"] == "23758-KUVEYT TURK-30.08.2026VDLİ-OMEGADEMİRAŞ" for s in cikis), cikis
    assert [(s["hesap"], s["borc"], s["alacak"]) for s in cikis] == [("320.01.001", 350000, 0),
                                                                      ("101.01.001", 0, 350000)], cikis
    k = [s for s in fisler if s["evrak_no"] == "5341489"]
    assert len(k) == 2 and k[0]["detay"].startswith("5341489-HALKBANK-11.09.2026VDLİ-"), k


if __name__ == "__main__":
    test_cek_no_aciklamada()
    print("1/1 kontrol gecti")
