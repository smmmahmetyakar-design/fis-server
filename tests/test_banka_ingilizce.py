"""İngilizce banka ekstresi (Ziraat 'Date / Explanation / Transaction Amount / Balance'), stil dosyası bozuk xlsx,
KMV, döviz hesabı olmayan döviz satışı, ATM, birden çok hesapta geçen ortak adı.  python tests/test_banka_ingilizce.py
Sahte veri."""
import io
import re
import sys
import tempfile
import zipfile
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app import isleyici  # noqa: E402
from app.belge_oku import belge_oku  # noqa: E402
from app.kurallar import KuralMotoru  # noqa: E402

MIZAN = [("100.01.001", "TL KASA"), ("102.01.001", "ZİRAAT 123 TL"), ("331.01.001", "KEMAL ORTAKOĞLU"),
         ("500.01.001", "KEMAL ORTAKOĞLU"), ("501.01.001", "KEMAL ORTAKOĞLU"), ("770.07.001", "BANKA MASRAFLARI")]
# geçmiş banka fişleri: ortağa giden/gelen havaleler hep 331, ATM para çekme 331, havale komisyonu 770
GECMIS = []
for i, (h, det, b, a) in enumerate([
        ("331.01.001", "KEMAL ORTAKOĞLU Ziraat Mobil Havale", 1000, 0),
        ("331.01.001", "KEMAL ORTAKOĞLU Ziraat Mobil Havale", 2000, 0),
        ("331.01.001", "KEMAL ORTAKOĞLU Ziraat Mobil Havale", 1500, 0),
        ("331.01.001", "KEMAL ORTAKOĞLU Ziraat Mobil Havale", 0, 3000),
        ("331.01.001", "KEMAL ORTAKOĞLU Ziraat Mobil Havale", 0, 500),
        ("331.01.001", "ATM PARA ÇEKME KART NO: 1234", 700, 0),
        ("770.07.001", "Havale Komisyonu Tahsilatı", 7.98, 0)], 1):
    ay = f"{i:02d}"
    GECMIS += [(str(i), f"2026-{ay}-05", h, det, b, a), (str(i), f"2026-{ay}-05", "102.01.001", det, a, b)]

EKSTRE = [
    ["Date", "Invoice No.", "Explanation", "Transaction Amount", "Balance"],
    ["26.10.2026", "F25025", " KEMAL ORTAKOĞLU Ziraat Mobil Havale", -1300, 4000],
    ["26.10.2026", "F24809", "İnternet Döviz Satış ; 0488GKDA26008122 Ref,EUR 250 Kur: 55,3168", 13829.2, 5300],
    ["12.10.2026", "F02643", "İnternet Döviz Alış İşlemi;0488GKDS26001042 Ref,EUR 100 Kur: 53,9306", -5393.06, -8529.2],
    ["12.10.2026", "F02644", "0488GKDS26001042 Ref,EUR 100TL.5393,06 KMV Matrahı İnternet Döviz Satış", -10.79, -3136.14],
    ["11.10.2026", "F02700", "ATM YATAN KART NO:1234 ATM:Z1749437", 3125.35, -3125.35],
]


def _bozuk_stilli_xlsx(satirlar) -> bytes:
    """openpyxl'in okuyamadığı boş <x:fill /> içeren stil dosyası (Ziraat İngilizce dışa aktarımı gibi)."""
    wb = openpyxl.Workbook(); ws = wb.active
    for r in satirlar:
        ws.append(r)
    b = io.BytesIO(); wb.save(b)
    zin = zipfile.ZipFile(io.BytesIO(b.getvalue())); out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for n in zin.namelist():
            d = zin.read(n)
            if n == "xl/styles.xml":
                d = re.sub(rb"<fills[^>]*>.*?</fills>", b'<fills count="1"><fill/></fills>', d, flags=re.S)
            z.writestr(n, d)
    return out.getvalue()


def _km(d: Path) -> KuralMotoru:
    wb = openpyxl.Workbook(); ws = wb.active
    for k, a in MIZAN:
        ws.append([k, a])
    wb.save(d / "mizan.xlsx")
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Fiş Aktarım Şablonu"
    ws.append(["FİŞ NO", "FİŞ TARİHİ", "FİŞ AÇIKLAMA", "HESAP KODU", "DETAY AÇIKLAMA", "BORÇ", "ALACAK"])
    for no, t, h, det, b, a in GECMIS:
        ws.append([no, t, "BANKA", h, det, b, a])
    wb.save(d / "gecmis.xlsx")
    return KuralMotoru(d / "mizan.xlsx", None, d / "ogr.json", d / "gecmis.xlsx")


def test_ingilizce_ekstre():
    d = Path(tempfile.mkdtemp())
    p = d / "ekstre_102.01.001.xlsx"
    p.write_bytes(_bozuk_stilli_xlsx(EKSTRE))
    try:
        openpyxl.load_workbook(p)
        bozuk = False
    except Exception:
        bozuk = True
    assert bozuk, "test dosyası openpyxl'i düşürmeli"
    ham = belge_oku(p, p.name)
    assert ham["tablolar"] and len(ham["tablolar"][0]) == len(EKSTRE), ham
    km = _km(d)
    fisler, uy = isleyici.isle("banka", [{"dosya": p.name, **ham}], km, 1)
    karsi = {s["detay"][:20]: s for s in fisler if s["hesap"] != "102.01.001"}
    assert len(karsi) == 5, (fisler, uy)
    # ortak adı mizanda 331/500/501'de ama geçmişte hep 331: KONTROL yok
    s = next(v for k, v in karsi.items() if "KEMAL" in k)
    assert s["hesap"] == "331.01.001" and not s.get("kontrol"), s
    # KMV: banka masrafı
    s = next(v for v in karsi.values() if "KMV" in v["detay"])
    assert s["hesap"] == "770.07.001", s
    # döviz satışı, mizanda EUR hesabı yok: 198 + açıklayıcı not
    s = next(v for v in karsi.values() if "Döviz Satış" in v["detay"] and "KMV" not in v["detay"])
    assert s["hesap"].startswith("198") and "EUR banka hesabı yok" in (s.get("not") or ""), s
    # ATM'ye yatan: geçmişteki ATM işlemlerinin hesabı (ortak), KONTROL'lü
    s = next(v for v in karsi.values() if "ATM" in v["detay"])
    assert s["hesap"] == "331.01.001" and s.get("kontrol"), s


if __name__ == "__main__":
    test_ingilizce_ekstre()
    print("1/1 kontrol gecti")
