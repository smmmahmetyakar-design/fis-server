"""
Belge okuyucu — gelen dosyayı ham satır/tablo verisine çevirir.
Desteklenen: .xlsx/.xls (openpyxl), .pdf (pdfplumber; metin yoksa OCR),
             .png/.jpg/.jpeg (OCR: pytesseract).

Çıktı: {"tur": "excel|pdf|pdf_ocr|resim_ocr", "tablolar": [[[hücre,...],...]],
         "ham_metin": "...", "uyari": "..."}
Not: OCR sonuçları hatalı olabilir; kullanıcı önizlemede düzeltir.
"""
import io
from pathlib import Path


def _ext(name: str) -> str:
    return Path(name).suffix.lower()


def excel_oku(path: Path):
    import openpyxl
    # openpyxl dosya YOLUNDAKİ uzantıya bakar: '.xls' uzantılı ama içeriği
    # gerçekte xlsx (ZIP) olan dosyaları reddeder. İçeriği BytesIO ile verince
    # uzantıdan bağımsız, doğrudan içeriğe göre okur.
    with open(path, "rb") as f:
        raw = f.read()
    wb = openpyxl.load_workbook(io.BytesIO(raw), data_only=True)
    tablolar = []
    for sn in wb.sheetnames:
        ws = wb[sn]
        satirlar = []
        for row in ws.iter_rows(values_only=True):
            if any(c is not None for c in row):
                satirlar.append([("" if c is None else c) for c in row])
        if satirlar:
            tablolar.append(satirlar)
    return {"tur": "excel", "tablolar": tablolar, "ham_metin": "", "uyari": ""}


def xls_oku(path: Path):
    """Eski .xls formatı (xlrd)."""
    try:
        import xlrd
    except ImportError:
        return {"tur": "excel", "tablolar": [], "ham_metin": "",
                "uyari": "xls okumak için xlrd gerekli"}
    wb = xlrd.open_workbook(path)
    tablolar = []
    for si in range(wb.nsheets):
        ws = wb.sheet_by_index(si)
        satirlar = []
        for r in range(ws.nrows):
            row = [ws.cell_value(r, c) for c in range(ws.ncols)]
            if any(str(v).strip() for v in row):
                satirlar.append(["" if v == "" else v for v in row])
        if satirlar:
            tablolar.append(satirlar)
    return {"tur": "excel", "tablolar": tablolar, "ham_metin": "", "uyari": ""}


def pdf_oku(path: Path):
    """Önce metin/tablo dener; metin çıkmazsa OCR'a düşer."""
    import pdfplumber
    tablolar = []
    ham = []
    metin_var = False
    with pdfplumber.open(path) as pdf:
        for sayfa in pdf.pages:
            t = sayfa.extract_text() or ""
            if t.strip():
                metin_var = True
                ham.append(t)
            # tablo çıkarımı
            for tb in (sayfa.extract_tables() or []):
                temiz = [[("" if c is None else str(c)) for c in row] for row in tb]
                if temiz:
                    tablolar.append(temiz)
    if metin_var or tablolar:
        return {"tur": "pdf", "tablolar": tablolar, "ham_metin": "\n".join(ham), "sayfalar": ham, "uyari": ""}
    # metin yok -> taranmış PDF, OCR gerek
    return pdf_ocr(path)


def pdf_ocr(path: Path):
    """Taranmış PDF: her sayfayı görüntüye çevirip OCR."""
    try:
        import pdfplumber
        import pytesseract
        from PIL import Image
    except ImportError as e:
        return {"tur": "pdf_ocr", "tablolar": [], "ham_metin": "",
                "uyari": f"OCR kütüphanesi yok: {e}"}
    ham = []
    with pdfplumber.open(path) as pdf:
        for sayfa in pdf.pages:
            im = sayfa.to_image(resolution=300).original
            txt = pytesseract.image_to_string(im, lang=_ocr_lang())
            ham.append(txt)
    return {"tur": "pdf_ocr", "tablolar": [], "ham_metin": "\n".join(ham), "sayfalar": ham,
            "uyari": "Taranmış PDF OCR ile okundu; verileri kontrol edin."}


def resim_oku(path: Path):
    """Fotoğraf/tarama: OCR (Türkçe fişler için ön işleme dahil)."""
    try:
        import pytesseract
        from PIL import Image, ImageOps
    except ImportError as e:
        return {"tur": "resim_ocr", "tablolar": [], "ham_metin": "",
                "uyari": f"OCR kütüphanesi yok: {e}"}
    im = Image.open(path)
    # ön işleme: gri + 2x büyüt + kontrast (küçük yazıları iyileştirir)
    try:
        im = im.convert("L")
        w, h = im.size
        # çok büyükse büyütme, çok küçükse 2x
        if max(w, h) < 2000:
            im = im.resize((w * 2, h * 2), Image.LANCZOS)
        im = ImageOps.autocontrast(im, cutoff=2)
    except Exception:
        pass
    txt = pytesseract.image_to_string(im, lang=_ocr_lang())
    return {"tur": "resim_ocr", "tablolar": [], "ham_metin": txt,
            "uyari": "Görüntü OCR ile okundu; verileri kontrol edin."}


def _ocr_lang():
    """Türkçe varsa tur+eng, yoksa eng."""
    try:
        import pytesseract
        langs = pytesseract.get_languages()
        if "tur" in langs:
            return "tur+eng"
    except Exception:
        pass
    return "eng"


def html_tablo_oku(path: Path):
    """Logo/e-Arşiv bazen '.xls' adıyla aslında HTML tablo üretir."""
    with open(path, "rb") as f:
        raw = f.read()
    metin = None
    for enc in ("utf-8", "cp1254", "latin-1"):
        try:
            metin = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if metin is None:
        metin = raw.decode("utf-8", "replace")
    tablolar = []
    try:  # varsa lxml/bs4 ile
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(metin, "html.parser")
        for tbl in soup.find_all("table"):
            satirlar = []
            for tr in tbl.find_all("tr"):
                hucreler = [td.get_text(strip=True) for td in tr.find_all(["td", "th"])]
                if any(h for h in hucreler):
                    satirlar.append(hucreler)
            if satirlar:
                tablolar.append(satirlar)
    except Exception:
        import re
        for tbl in re.findall(r"<table.*?</table>", metin, re.I | re.S):
            satirlar = []
            for tr in re.findall(r"<tr.*?</tr>", tbl, re.I | re.S):
                hucreler = [re.sub(r"<[^>]+>", "", c).strip()
                            for c in re.findall(r"<t[dh].*?</t[dh]>", tr, re.I | re.S)]
                if any(h for h in hucreler):
                    satirlar.append(hucreler)
            if satirlar:
                tablolar.append(satirlar)
    return {"tur": "excel", "tablolar": tablolar, "ham_metin": "",
            "uyari": "" if tablolar else "HTML tablo bulunamadı"}


def csv_oku(path: Path):
    """Banka CSV dökümü (Akbank vb.): kodlama (UTF-8 / Windows-1254) ve ayraç (; , sekme |)
    kendiliğinden bulunur. Satırlar tek tablo olarak döner."""
    import csv
    ham = path.read_bytes()
    metin = None
    for kod in ("utf-8-sig", "cp1254", "latin-1"):
        try:
            metin = ham.decode(kod)
            break
        except UnicodeDecodeError:
            continue
    satirlar = [s for s in metin.splitlines()]
    ornek = "\n".join(satirlar[:40])
    ayrac = max((";", ",", "\t", "|"), key=lambda a: ornek.count(a))
    tablo = [[h.strip() for h in r] for r in csv.reader(satirlar, delimiter=ayrac)]
    tablo = [r for r in tablo if any(r)] or []
    return {"tur": "csv", "tablolar": [tablo] if tablo else [], "ham_metin": "", "uyari": ""}


def _magic_tur(path: Path) -> str:
    """Dosyanın GERÇEK türünü ilk baytlarından anlar (uzantıdan bağımsız)."""
    try:
        with open(path, "rb") as f:
            head = f.read(512)
    except OSError:
        return ""
    if head[:4] == b"PK\x03\x04":
        return "xlsx"                     # ZIP -> modern xlsx/xlsm
    if head[:4] == b"\xd0\xcf\x11\xe0":
        return "xls"                      # OLE2 -> eski xls
    if head[:5] == b"%PDF-":
        return "pdf"
    if (head[:8] == b"\x89PNG\r\n\x1a\n" or head[:3] == b"\xff\xd8\xff"
            or head[:6] in (b"GIF87a", b"GIF89a") or head[:4] == b"RIFF"
            or head[:2] in (b"II", b"MM") and b"\x2a" in head[:4]):
        return "resim"
    dusuk = head.lstrip().lower()
    if dusuk[:5] == b"<html" or dusuk[:6] == b"<table" or dusuk[:5] == b"<?xml" \
            or dusuk[:9] == b"<!doctype":
        return "html"
    return ""


def belge_oku(path: Path, orijinal_ad: str = ""):
    ext = _ext(orijinal_ad or path.name)
    tur = _magic_tur(path)              # önce GERÇEK içerik türü

    # İçerik türü kesinse ona güven (uzantı yalan söyleyebilir: Logo .xls der ama xlsx yazar)
    if tur == "xlsx":
        return excel_oku(path)
    if tur == "xls":
        return xls_oku(path)
    if tur == "pdf":
        return pdf_oku(path)
    if tur == "resim":
        return resim_oku(path)
    if tur == "html":
        return html_tablo_oku(path)

    # İçerikten anlaşılmadıysa uzantıya düş
    if ext in (".xlsx", ".xlsm"):
        return excel_oku(path)
    if ext == ".xls":
        # uzantı .xls ama içerik OLE2 değil: yine de xlsx dene, olmazsa HTML
        try:
            return excel_oku(path)
        except Exception:
            return html_tablo_oku(path)
    if ext == ".pdf":
        return pdf_oku(path)
    if ext in (".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff"):
        return resim_oku(path)
    if ext in (".csv", ".txt"):
        return csv_oku(path)
    return {"tur": "bilinmeyen", "tablolar": [], "ham_metin": "",
            "uyari": f"Dosya türü tanınamadı (uzantı: {ext or 'yok'})"}
