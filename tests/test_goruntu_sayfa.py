"""Çok faturalı tek PDF (ör. entegratör portalından toplu indirilen) + aynı faturaların XML'leri:
her satırın 'Faturayı aç'ı PDF'in kendi sayfasını açmalı.  python tests/test_goruntu_sayfa.py — sahte veri."""
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from app import fatura_pdf, isleyici  # noqa: E402
from test_fatura_karisik_hesap import _km  # noqa: E402

PDF = "Entegrator Portali.pdf"


def _kayit(no, ad, kaynak, dosya, sayfa=None):
    k = {"fatura_no": no, "tarih": "2026-11-05", "cari_ad": ad, "kalemler": [{"oran": 20, "matrah": 100, "kdv": 20}],
         "toplam": 120, "tevkifat": 0, "tevkifat_kod": "", "ek_vergi": 0.0, "senaryo": "TICARIFATURA", "tur": "SATIS", "yon": "satis", "dosya": dosya, "kaynak": kaynak, "eksik": [], "yz": []}
    if sayfa:
        k["sayfa"] = sayfa
    return k


def test_xml_faturasi_pdf_sayfasini_acar():
    tum = [_kayit("XYZ2026000000004", "KAPPA PLASTİK LTD", "xml", "a.xml"),
           _kayit("XYZ2026000000009", "KAPPA PLASTİK LTD", "xml", "b.xml"),
           _kayit("XYZ2026000000009", "KAPPA PLASTİK LTD", "pdf", PDF, 1),
           _kayit("XYZ2026000000004", "KAPPA PLASTİK LTD", "pdf", PDF, 3)]
    pdfler = fatura_pdf.xml_oncelikli(tum)
    assert len(pdfler) == 2 and {p["goruntu_sayfa"] for p in pdfler} == {1, 3}, pdfler
    km = _km(Path(tempfile.mkdtemp()))
    fisler, _ = isleyici.isle_fatura([], km, 10, yon="satis", pdf_faturalar=pdfler, gider_ogrenme={})
    sayfa = {s["evrak_no"]: (s.get("pdf"), s.get("pdf_sayfa")) for s in fisler if s.get("pdf")}
    assert sayfa == {"XYZ2026000000004": (PDF, 3), "XYZ2026000000009": (PDF, 1)}, sayfa


if __name__ == "__main__":
    test_xml_faturasi_pdf_sayfasini_acar()
    print("1/1 kontrol gecti")
