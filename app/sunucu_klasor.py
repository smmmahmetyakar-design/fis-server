"""
Sunucu firma klasörleri (/srv/veri/firmalar -> container içinde /firmalar).

Klasör düzeni (her firma için):
    NN_FIRMA_ADI/
        mizan/          <- mizan Excel'i (xlsx/xls)
        fislistesi/     <- muhasebe programından alınan fiş listesi (xlsx/xls/pdf)
        banka/ dekont/ fatura/ fis/   (gelen / cikan alt klasörleri)

Bu modül klasörü yalnızca OKUR. Dosyalar araç veri klasörüne kopyalanarak
kullanılır; sunucudaki asıl dosyaya hiç yazılmaz.

Dosyalar her zaman doğru alt klasörde durmuyor (ör. mizan fislistesi/ içinde,
fiş listesi PDF'i banka/ içinde). Bu yüzden önce beklenen klasöre bakılır,
sonra firma klasörünün geri kalanında dosya ADINDAN tanınanlar da aday
gösterilir ve "yanlış yerde" diye işaretlenir.
"""
import io, os, re
from datetime import datetime
from pathlib import Path

from app.kurallar import norm, kelimeler

FIRMALAR_DIR = Path(os.environ.get("FIS_FIRMALAR", "/firmalar"))

TURLER = ("mizan", "fis-listesi", "gecmis-fis-pdf")
_EXCEL = (".xlsx", ".xls")

# tür -> (beklenen alt klasör, kabul edilen uzantılar)
_BEKLENEN = {
    "mizan": ("mizan", _EXCEL),
    "fis-listesi": ("fislistesi", _EXCEL),
    "gecmis-fis-pdf": ("fislistesi", (".pdf",)),
}


def _ad_mizan_mi(ad: str) -> bool:
    return "MIZAN" in norm(ad).replace(" ", "")


def _ad_fis_listesi_mi(ad: str) -> bool:
    """Fiş listesi ya da muavin defter (ikisi de geçmiş kayıt kaynağı)."""
    n = norm(ad).replace("_", " ")
    return bool(re.search(r"FIS\s*LISTESI|MUAVIN|YEVMIYE", n))


def klasor_var_mi() -> bool:
    return FIRMALAR_DIR.is_dir()


def guzel_ad(klasor: str) -> str:
    """'02_ARC_LTD' -> 'ARC LTD'"""
    s = re.sub(r"^\d+[_\-\s]*", "", klasor)
    return re.sub(r"[_]+", " ", s).strip() or klasor


def klasorler() -> list:
    """Sunucudaki firma klasörleri (gizli ve '_' ile başlayanlar hariç)."""
    if not klasor_var_mi():
        return []
    out = []
    for p in sorted(FIRMALAR_DIR.iterdir()):
        if p.is_dir() and not p.name.startswith((".", "_")):
            out.append({"klasor": p.name, "ad": guzel_ad(p.name)})
    return out


def klasor_yolu(klasor: str) -> Path | None:
    """Klasör adını doğrular; yalnızca FIRMALAR_DIR'in doğrudan alt klasörü olabilir."""
    if not klasor or "/" in klasor or "\\" in klasor or klasor in (".", ".."):
        return None
    p = FIRMALAR_DIR / klasor
    return p if p.is_dir() else None


def klasor_oner(firma_ad: str, kod: str = "") -> str:
    """Bağlanmamış bir firma için ada en çok benzeyen klasörü önerir.
    Kelime örtüşmesine bakar; tek kelimelik zayıf eşleşmede öneri vermez."""
    hedef = kelimeler(firma_ad) | kelimeler(kod.replace("_", " "))
    if not hedef:
        return ""
    en_iyi, en_skor = "", 0.0
    for k in klasorler():
        kw = kelimeler(k["ad"])
        if not kw:
            continue
        ortak = len(hedef & kw)
        skor = ortak / max(len(kw), 1)
        if ortak and skor > en_skor:
            en_iyi, en_skor = k["klasor"], skor
    return en_iyi if en_skor >= 0.5 else ""


def adaylar(klasor_path: Path, tur: str) -> list:
    """Firma klasöründe verilen tür için aday dosyaları listeler (en yeni önce).
    'cikan' klasörleri (araç çıktıları) hiç taranmaz."""
    if tur not in _BEKLENEN:
        return []
    alt, uzantilar = _BEKLENEN[tur]
    bulunan = {}

    def ekle(p: Path, dogru_yer: bool):
        st = p.stat()
        rel = p.relative_to(klasor_path).as_posix()
        bulunan[rel] = {
            "dosya": p.name, "yol": rel, "dogru_yer": dogru_yer,
            "boyut_kb": round(st.st_size / 1024, 1),
            "tarih": datetime.fromtimestamp(st.st_mtime).isoformat(timespec="seconds"),
            "mtime": st.st_mtime,
        }

    # 1) beklenen klasör
    beklenen = klasor_path / alt
    if beklenen.is_dir():
        for p in beklenen.iterdir():
            if not p.is_file() or p.suffix.lower() not in uzantilar or p.name.startswith((".", "~$")):
                continue
            if tur == "fis-listesi" and _ad_mizan_mi(p.name):
                continue        # fislistesi/ içine konmuş mizan
            ekle(p, True)

    # 2) firma klasörünün geri kalanı — addan tanınanlar
    for p in klasor_path.rglob("*"):
        if not p.is_file() or p.suffix.lower() not in uzantilar or p.name.startswith((".", "~$")):
            continue
        parcalar = p.relative_to(klasor_path).parts
        if "cikan" in parcalar or (len(parcalar) > 1 and parcalar[0] == alt):
            continue
        if tur == "mizan" and _ad_mizan_mi(p.name):
            ekle(p, False)
        elif tur in ("fis-listesi", "gecmis-fis-pdf") and _ad_fis_listesi_mi(p.name) \
                and not _ad_mizan_mi(p.name):
            ekle(p, False)

    return sorted(bulunan.values(), key=lambda x: (x["dogru_yer"], x["mtime"]), reverse=True)


def en_uygun(klasor_path: Path, tur: str) -> dict | None:
    """Otomatik alım için seçilecek dosya: doğru klasördeki en yeni dosya;
    doğru klasörde hiç yoksa başka yerdeki en yeni dosya."""
    liste = adaylar(klasor_path, tur)
    return liste[0] if liste else None


def guvenli_dosya(klasor_path: Path, rel: str) -> Path:
    """Göreli yolu firma klasörü içinde çözer; dışarı çıkmaya izin vermez."""
    p = (klasor_path / rel).resolve()
    if not p.is_file() or not p.is_relative_to(klasor_path.resolve()):
        raise ValueError("Dosya firma klasörünün içinde değil")
    return p


def excel_xlsx_bytes(data: bytes) -> bytes:
    """Excel içeriğini openpyxl'in okuyabileceği .xlsx'e çevirir.
    Mizan ve fiş listesi okuyucuları openpyxl kullanıyor; eski .xls (BIFF)
    dosyalarını okuyamıyor. .xlsx zaten ZIP'tir ('PK'), olduğu gibi döner."""
    if data[:2] == b"PK":
        return data
    if data[:8] != b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1":
        return data             # tanınmayan içerik — okuyucu ne yapacağına karar versin
    import xlrd, openpyxl
    kaynak = xlrd.open_workbook(file_contents=data)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for si in range(kaynak.nsheets):
        ks = kaynak.sheet_by_index(si)
        ws = wb.create_sheet(title=(ks.name or f"Sayfa{si + 1}")[:31])
        for r in range(ks.nrows):
            satir = []
            for c in range(ks.ncols):
                hucre = ks.cell(r, c)
                if hucre.ctype == xlrd.XL_CELL_DATE:
                    try:
                        satir.append(xlrd.xldate_as_datetime(hucre.value, kaynak.datemode))
                        continue
                    except Exception:
                        pass
                if hucre.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                    satir.append(None)
                else:
                    satir.append(hucre.value)
            ws.append(satir)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
