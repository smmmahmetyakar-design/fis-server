"""
Sunucu firma klasörleri (/srv/veri/firmalar -> container içinde /firmalar).

Klasör düzeni (her firma için):
    NN_FIRMA_ADI/
        mizan/          <- mizan Excel'i (xlsx/xls)
        fislistesi/     <- muhasebe programından alınan fiş listesi / muavin (xlsx/xls)
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

TURLER = ("mizan", "fis-listesi")
_EXCEL = (".xlsx", ".xls")

# tür -> (beklenen alt klasör, kabul edilen uzantılar)
_BEKLENEN = {
    "mizan": ("mizan", _EXCEL),
    "fis-listesi": ("fislistesi", _EXCEL),
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
        elif tur == "fis-listesi" and _ad_fis_listesi_mi(p.name) \
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


# ------------------------------------------------------------------ fatura PDF'leri
import threading

FATURA_NO_RE = re.compile(r"[A-Z0-9]{3}20\d{2}\d{9}")
_OCR_ADAY_RE = re.compile(r"[A-Z0-9]{3}[2Z][0O][0-9OILSB]{11}")
_EN_FAZLA_SAYFA = 300


def fatura_nolari(metin: str, ocr: bool = False) -> list:
    """Metindeki e-fatura / e-arşiv numaraları (GIB2026000000100 gibi).
    OCR metninde boşluklar ve O/0, I/1 karışıklıkları da düzeltilir."""
    up = (metin or "").upper()
    bulunan = set(FATURA_NO_RE.findall(up))
    if ocr:
        tablo = str.maketrans({"O": "0", "I": "1", "L": "1", "S": "5", "B": "8", "Z": "2"})
        for satir in up.splitlines():
            for aday in _OCR_ADAY_RE.findall(satir.replace(" ", "")):
                duz = aday[:3] + aday[3:].translate(tablo)
                if FATURA_NO_RE.fullmatch(duz):
                    bulunan.add(duz)
    return sorted(bulunan)


def _pdf_metin_nolari(p: Path) -> tuple:
    """PDF metnindeki (tüm sayfalar) fatura numaraları. Döner: (nolar, metin_var)."""
    try:
        import pdfplumber
        with pdfplumber.open(p) as pdf:
            metin = "\n".join((s.extract_text() or "") for s in pdf.pages[:_EN_FAZLA_SAYFA])
    except Exception:
        return [], False
    return fatura_nolari(metin), bool(metin.strip())


def _pdf_ocr_nolari(p: Path) -> list:
    """Taranmış PDF: sayfaları OCR ile okuyup fatura numaralarını bulur (yavaş)."""
    import pdfplumber, pytesseract
    nolar = set()
    with pdfplumber.open(p) as pdf:
        for s in pdf.pages[:_EN_FAZLA_SAYFA]:
            im = s.to_image(resolution=200).original
            nolar.update(fatura_nolari(pytesseract.image_to_string(im), ocr=True))
    return sorted(nolar)


_ocr_kilit = threading.Lock()
_ocr_sirada: set = set()


def _ocr_arka_plan(isler: list, onbellek_yolu: Path):
    """[(p, rel, mtime)] — sırayla OCR'lar, sonucu önbelleğe yazar."""
    import json
    for p, rel, mt in isler:
        try:
            nolar = _pdf_ocr_nolari(p)
        except Exception:
            nolar = []
        with _ocr_kilit:
            try:
                onb = json.loads(onbellek_yolu.read_text(encoding="utf-8"))
            except Exception:
                onb = {}
            onb[rel] = [mt, nolar, "ocr"]
            try:
                onbellek_yolu.write_text(json.dumps(onb, ensure_ascii=False), encoding="utf-8")
            except Exception:
                pass
            _ocr_sirada.discard(str(p))


def fatura_pdf_dizini(klasor_path: Path, onbellek_yolu: Path) -> tuple:
    """Firma klasöründeki fatura PDF'leri: fatura no -> göreli yol.
    Dosya adındaki ve PDF metnindeki (tüm sayfalar — taranmış toplu PDF'te her sayfa
    ayrı fatura olabilir) numaralar. Metni olmayan (taranmış) PDF'ler arka planda OCR'lanır.
    Sonuç (yol + değişiklik zamanı) önbelleğe yazılır; her seferinde yeniden okunmaz.
    'cikan' (araç çıktıları) taranmaz.
    Döner: (dizin, bilgi) — bilgi: taranan, numarali, ocr_bekleyen, numarasiz (dosya listesi)"""
    import json
    with _ocr_kilit:
        try:
            onb = json.loads(onbellek_yolu.read_text(encoding="utf-8"))
        except Exception:
            onb = {}
    dizin, bilgi = {}, {"taranan": 0, "numarali": 0, "ocr_bekleyen": 0, "numarasiz": [], "klasor": ""}
    # Fatura PDF'leri: "fatura/" altı ve firma klasöründe adı FATURA / ARŞİV geçen her klasör
    # ("Alış Faturaları", "e-Arsiv" gibi). Banka/dekont klasörleri taranmaz: ekstre
    # açıklamalarında fatura numarası geçebilir, faturanın kendisi değildir.
    def _fatura_yolu_mu(parcalar):
        return any(("FATUR" in norm(x).replace(" ", "") or "ARSIV" in norm(x).replace(" ", ""))
                   for x in parcalar)
    yeni_onb, ocr_isleri = {}, []
    adaylar_ = []
    for p in sorted(klasor_path.rglob("*.[pP][dD][fF]")):
        if not p.is_file() or p.name.startswith((".", "~$")):
            continue
        parcalar = p.relative_to(klasor_path).parts[:-1]
        if "cikan" in parcalar or not _fatura_yolu_mu(parcalar):
            continue
        adaylar_.append(p)
    if not adaylar_ and not (klasor_path / "fatura").is_dir():
        bilgi["klasor"] = "yok"
    for p in adaylar_:
        rel = p.relative_to(klasor_path).as_posix()
        mt = p.stat().st_mtime
        bilgi["taranan"] += 1
        adda = FATURA_NO_RE.findall(norm(p.stem).replace(" ", ""))
        kayit = onb.get(rel)
        if kayit and kayit[0] == mt and kayit[1] is not None:
            nolar = kayit[1]
            yeni_onb[rel] = kayit
        else:
            nolar, metin_var = _pdf_metin_nolari(p)
            if not nolar and not metin_var:
                # taranmış PDF — OCR arka planda; şimdilik yalnız dosya adındaki numara
                yeni_onb[rel] = [mt, None]
                if str(p) not in _ocr_sirada:
                    _ocr_sirada.add(str(p))
                    ocr_isleri.append((p, rel, mt))
                bilgi["ocr_bekleyen"] += 1
                nolar = []
            else:
                yeni_onb[rel] = [mt, nolar]
        nolar = sorted(set(nolar) | set(adda))
        if nolar:
            bilgi["numarali"] += 1
        elif str(p) not in _ocr_sirada:
            bilgi["numarasiz"].append(rel)
        for n in nolar:
            dizin.setdefault(n, rel)
    with _ocr_kilit:
        # bu arada biten OCR sonuçları ezilmesin
        try:
            simdiki = json.loads(onbellek_yolu.read_text(encoding="utf-8"))
        except Exception:
            simdiki = {}
        for rel, x in yeni_onb.items():
            y = simdiki.get(rel)
            if x[1] is None and y and y[0] == x[0] and y[1] is not None:
                yeni_onb[rel] = y
                for n in y[1]:
                    dizin.setdefault(n, rel)
        bilgi["ocr_bekleyen"] = sum(1 for x in yeni_onb.values() if x[1] is None)
        try:
            onbellek_yolu.parent.mkdir(parents=True, exist_ok=True)
            onbellek_yolu.write_text(json.dumps(yeni_onb, ensure_ascii=False), encoding="utf-8")
        except Exception:
            pass
    if ocr_isleri:
        threading.Thread(target=_ocr_arka_plan, args=(ocr_isleri, onbellek_yolu), daemon=True,
                         name="pdf-ocr-dizin").start()
    return dizin, bilgi


def fatura_klasor_ozeti(klasor_path: Path) -> list:
    """Teşhis: firma klasöründe adında FATURA/ARŞİV geçen klasörlerdeki dosya türleri
    ve okunamayan (izin) klasörler. PDF bulunamadığında nedenini göstermek için."""
    out = {}
    hatalar = []

    def _hata(e):
        hatalar.append(f"{getattr(e, 'filename', '')}: {e.strerror or e}")

    for kok, dizinler, dosyalar in os.walk(klasor_path, onerror=_hata):
        rel = Path(kok).relative_to(klasor_path)
        if "cikan" in rel.parts:
            dizinler[:] = []
            continue
        if not any(("FATUR" in norm(x).replace(" ", "") or "ARSIV" in norm(x).replace(" ", "")) for x in rel.parts):
            continue
        uz = {}
        for f in dosyalar:
            if f.startswith((".", "~$")):
                continue
            e = (Path(f).suffix.lower() or "(uzantısız)")
            uz[e] = uz.get(e, 0) + 1
        out[rel.as_posix()] = uz
    sonuc = [{"yol": k, "uzantilar": v, "dosya": sum(v.values())} for k, v in sorted(out.items())]
    for h in hatalar[:10]:
        try:
            yol = Path(h.split(":")[0]).relative_to(klasor_path).as_posix()
        except Exception:
            yol = h.split(":")[0]
        sonuc.append({"yol": yol, "uzantilar": {}, "dosya": 0, "hata": h.split(":", 1)[-1].strip()})
    return sonuc
