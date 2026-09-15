"""
Fiş Aktarım Aracı — bağımsız server
Banka dökümü / Fatura / Çek listesi -> muhasebe fişi (Fiş Aktarım Şablonu).
Her firma için: mizan + 3 kural dosyası (banka/fatura/cek) + gelen belgeler + öğrenme.

Klasör yapısı:
  data/<firma>/
    mizan.xlsx
    kural_banka.xlsx  kural_fatura.xlsx  kural_cek.xlsx
    banka/  fatura/  cek/          <- gelen belgeler (excel/pdf/resim)
    banka_ogrenme.json  ...        <- öğrenilen eşleştirmeler
    cikti/                          <- üretilen fiş aktarım dosyaları
    meta.json
"""
import io, json, os, re, shutil, unicodedata
from datetime import datetime
from pathlib import Path

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse, FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
from starlette.middleware.base import BaseHTTPMiddleware

from app.kurallar import KuralMotoru, norm, kural_excel_oku, mizan_hesaplar, fis_listesi_ogren, fatura_gider_ogren, fatura_gelir_ogren
from app.belge_oku import belge_oku
from app import isleyici
from app.routes.enhanced import router as enhanced_router

DATA_DIR = Path(os.environ.get("FIS_DATA", "/data"))
DATA_DIR.mkdir(parents=True, exist_ok=True)

FIRMALAR_KOK = Path(os.environ.get("FIRMALAR_KOK", "/firmalar"))
KLASOR_ESLESTIRME_DOSYASI = DATA_DIR / "firma_klasor_eslestirme.json"


def _firma_klasor_adi(kod: str):
    return _read_json(KLASOR_ESLESTIRME_DOSYASI, {}).get(kod)


def _en_yeni_xlsx(dizin: Path):
    if not dizin.is_dir():
        return None
    adaylar = [f for f in dizin.iterdir()
               if f.is_file() and f.suffix.lower() in (".xlsx", ".xls") and not f.name.startswith("~$")]
    if not adaylar:
        return None
    return max(adaylar, key=lambda f: f.stat().st_mtime)

app = FastAPI(title="Fiş Aktarım Aracı")
app.include_router(enhanced_router)


# ----------------------------------------------------------------- önbellek kapalı
class NoCacheMiddleware(BaseHTTPMiddleware):
    """Arayüz (index.html) tarayıcı önbelleğinde takılıp kalmasın diye her
    yanıta 'önbelleğe alma' başlığı ekler. Her deploy'dan sonra kullanıcı
    ?v=2 gibi bir numara eklemeden / sert yenileme yapmadan en güncel
    sürümü otomatik görür."""
    async def dispatch(self, request, call_next):
        response = await call_next(request)
        response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
        response.headers["Pragma"] = "no-cache"
        response.headers["Expires"] = "0"
        return response


app.add_middleware(NoCacheMiddleware)

TIPLER = ("banka", "fatura", "fatura_satis", "masraf", "cek")


# ----------------------------------------------------------------- yardımcı
def slugify(s: str) -> str:
    s = norm(s).lower().replace(" ", "_")
    return re.sub(r"[^a-z0-9_]", "", s) or "firma"


def firma_dir(kod: str) -> Path:
    d = DATA_DIR / kod
    if not d.exists():
        raise HTTPException(404, "Firma yok")
    return d


def _read_json(p: Path, default):
    if p.exists():
        try:
            return json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            return default
    return default


def _write_json(p: Path, data):
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


# ----------------------------------------------------------------- firma yönetimi
@app.get("/api/firmalar")
def firmalar():
    out = []
    for d in sorted(DATA_DIR.iterdir()):
        if not d.is_dir() or d.name.startswith("_"):
            continue
        meta = _read_json(d / "meta.json", {})
        out.append({
            "kod": d.name,
            "ad": meta.get("ad", d.name),
            "mizan_var": (d / "mizan.xlsx").exists(),
            "kurallar": {t: (d / f"kural_{t}.xlsx").exists() for t in TIPLER},
        })
    return out


@app.post("/api/firma")
def firma_olustur(body: dict):
    ad = (body.get("ad") or "").strip()
    if not ad:
        raise HTTPException(400, "Firma adı gerekli")
    kod = slugify(ad)
    d = DATA_DIR / kod
    d.mkdir(parents=True, exist_ok=True)
    for t in TIPLER:
        (d / t).mkdir(exist_ok=True)
    (d / "cikti").mkdir(exist_ok=True)
    meta = _read_json(d / "meta.json", {})
    meta["ad"] = ad
    meta.setdefault("olusturma", datetime.now().isoformat(timespec="seconds"))
    _write_json(d / "meta.json", meta)
    return {"kod": kod, "ad": ad}


# ----------------------------------------------------------------- mizan & kural yükleme
@app.post("/api/firma/{kod}/mizan")
async def mizan_yukle(kod: str, file: UploadFile = File(...)):
    d = firma_dir(kod)
    data = await file.read()
    (d / "mizan.xlsx").write_bytes(data)
    hes = mizan_hesaplar(d / "mizan.xlsx")
    return {"ok": True, "hesap_sayisi": len(hes)}


@app.delete("/api/firma/{kod}/mizan")
def mizan_sil(kod: str):
    """Mizanı siler; yeni mizan yüklenene kadar hesap listesi/eşleştirme boş kalır."""
    d = firma_dir(kod)
    (d / "mizan.xlsx").unlink(missing_ok=True)
    return {"ok": True}


@app.post("/api/firma/{kod}/klasor-eslestirme")
def klasor_eslestirme_kaydet(kod: str, body: dict):
    klasor = (body.get("klasor") or "").strip()
    if not klasor:
        raise HTTPException(400, "Klasor adi gerekli")
    eslestirme = _read_json(KLASOR_ESLESTIRME_DOSYASI, {})
    eslestirme[kod] = klasor
    _write_json(KLASOR_ESLESTIRME_DOSYASI, eslestirme)
    return {"ok": True, "kod": kod, "klasor": klasor}


@app.post("/api/firma/{kod}/mizan/sunucudan")
def mizan_sunucudan_yukle(kod: str):
    d = firma_dir(kod)
    klasor = _firma_klasor_adi(kod)
    if not klasor:
        raise HTTPException(404, f"'{kod}' icin sunucu klasor eslestirmesi tanimli degil ({KLASOR_ESLESTIRME_DOSYASI.name})")
    kaynak = _en_yeni_xlsx(FIRMALAR_KOK / klasor / "mizan")
    if not kaynak:
        raise HTTPException(404, f"{FIRMALAR_KOK}/{klasor}/mizan altinda Excel dosyasi bulunamadi")
    (d / "mizan.xlsx").write_bytes(kaynak.read_bytes())
    hes = mizan_hesaplar(d / "mizan.xlsx")
    return {"ok": True, "hesap_sayisi": len(hes), "kaynak_dosya": kaynak.name}


@app.delete("/api/firma/{kod}/ogrenme/{tip}")
def ogrenme_sil(kod: str, tip: str):
    """Fiş Listesi'nden (veya elle düzenlemeden) öğrenilmiş eşleştirmeleri sıfırlar,
    böylece yeni bir Fiş Listesi baştan öğretilebilir."""
    if tip not in TIPLER:
        raise HTTPException(400, "Geçersiz tip")
    d = firma_dir(kod)
    (d / f"{tip}_ogrenme.json").unlink(missing_ok=True)
    if tip in ("fatura", "fatura_satis"):
        (d / f"{tip}_gider_eslestirme.json").unlink(missing_ok=True)
    return {"ok": True}


@app.post("/api/firma/{kod}/kural/{tip}")
async def kural_yukle(kod: str, tip: str, file: UploadFile = File(...)):
    if tip not in TIPLER:
        raise HTTPException(400, "Geçersiz tip")
    d = firma_dir(kod)
    data = await file.read()
    (d / f"kural_{tip}.xlsx").write_bytes(data)
    k = kural_excel_oku(d / f"kural_{tip}.xlsx")
    return {"ok": True, "hesap_kurali": len(k["hesaplar"]), "talimat": len(k["talimatlar"])}


@app.post("/api/firma/{kod}/ogren-fis-listesi/{tip}")
async def ogren_fis_listesi(kod: str, tip: str, banka_hesap_kodu: str = Form(""), file: UploadFile = File(...)):
    """
    Gerçek muhasebe fiş geçmişinizi (Logo Tiger 'fiş listesi' export'u) yükleyin:
    - banka/çek: hedef hesap kodunu da verin, karşı hesap eşleştirmelerini
      öğrenip {tip}_ogrenme.json'a ekler.
    - fatura (alış): hedef hesap kodu gerekmez — e-Fatura listesinde ürün/
      hizmet açıklaması olmadığından, hangi CARİ (tedarikçi) hesabının hangi
      GİDER hesabına işlendiği doğrudan geçmiş fişlerden öğrenilip
      fatura_gider_eslestirme.json'a yazılır.
    - fatura_satis: aynı mantığın aynası — hangi CARİ (müşteri) hesabının
      hangi GELİR hesabına işlendiği öğrenilip fatura_satis_gider_eslestirme.json'a yazılır.
    Manuel 'Hesap Kodu Eşleştirme' dosyası hazırlamaya gerek kalmaz.
    """
    if tip not in TIPLER:
        raise HTTPException(400, "Geçersiz tip")
    d = firma_dir(kod)
    tmp = d / "_gecici_fis_listesi.xlsx"
    tmp.write_bytes(await file.read())
    try:
        if tip in ("fatura", "fatura_satis"):
            yeni = fatura_gelir_ogren(tmp) if tip == "fatura_satis" else fatura_gider_ogren(tmp)
            og_path = d / f"{tip}_gider_eslestirme.json"
        else:
            hesap_kodu = (banka_hesap_kodu or "").strip()
            if not hesap_kodu:
                raise HTTPException(400, "Hedef hesap kodu gerekli")
            yeni = fis_listesi_ogren(tmp, hesap_kodu)
            og_path = d / f"{tip}_ogrenme.json"
    finally:
        tmp.unlink(missing_ok=True)
    mevcut = _read_json(og_path, {})
    mevcut.update(yeni)
    _write_json(og_path, mevcut)
    return {"ok": True, "yeni_kural": len(yeni), "toplam_ogrenilen": len(mevcut)}


@app.get("/api/firma/{kod}/durum")
def firma_durum(kod: str):
    d = firma_dir(kod)
    hes = mizan_hesaplar(d / "mizan.xlsx") if (d / "mizan.xlsx").exists() else []
    out = {"mizan_hesap": len(hes), "tipler": {}, "sunucu_klasoru": _firma_klasor_adi(kod)}
    for t in TIPLER:
        kp = d / f"kural_{t}.xlsx"
        k = kural_excel_oku(kp) if kp.exists() else {"hesaplar": [], "talimatlar": []}
        belgeler = [f.name for f in (d / t).iterdir() if f.is_file()] if (d / t).exists() else []
        og = _read_json(d / f"{t}_ogrenme.json", {})
        gider_og = _read_json(d / f"{t}_gider_eslestirme.json", {}) if t in ("fatura", "fatura_satis") else {}
        out["tipler"][t] = {
            "kural_var": kp.exists(),
            "kural_hesap": len(k["hesaplar"]),
            "belge_sayisi": len(belgeler),
            "belgeler": belgeler,
            "ogrenilen": len(og),
            "gider_ogrenilen": len(gider_og),
        }
    return out


# ----------------------------------------------------------------- belge yükleme
@app.post("/api/firma/{kod}/belge/{tip}")
async def belge_yukle(kod: str, tip: str, file: UploadFile = File(...)):
    if tip not in TIPLER:
        raise HTTPException(400, "Geçersiz tip")
    d = firma_dir(kod)
    hedef = d / tip
    hedef.mkdir(exist_ok=True)
    fname = re.sub(r"[^\w.\- ]", "_", file.filename or "belge")
    (hedef / fname).write_bytes(await file.read())
    return {"ok": True, "dosya": fname}


@app.delete("/api/firma/{kod}/belge/{tip}/{fname}")
def belge_sil(kod: str, tip: str, fname: str):
    if tip not in TIPLER or ".." in fname or "/" in fname or "\\" in fname:
        raise HTTPException(400, "Geçersiz")
    d = firma_dir(kod)
    p = d / tip / fname
    if not p.exists():
        raise HTTPException(404, "Yok")
    p.unlink()
    return {"ok": True}


@app.get("/api/firma/{kod}/belge/{tip}/{fname}/sayfa/{sayfa}")
def belge_sayfa_gorsel(kod: str, tip: str, fname: str, sayfa: int):
    """Bir fatura satırının kaynağı olan PDF sayfasını PNG olarak döner —
    önizleme tablosundaki 🖼 butonu bununla faturanın orijinal görselini
    yeni sekmede açar. Sadece PDF kaynaklı fatura satırlarında anlamlı
    (Excel kaynaklı satırların 'sayfa'sı yok, buton zaten gösterilmez)."""
    if tip not in TIPLER or ".." in fname or "/" in fname or "\\" in fname:
        raise HTTPException(400, "Geçersiz")
    d = firma_dir(kod)
    p = d / tip / fname
    if not p.exists() or p.suffix.lower() != ".pdf":
        raise HTTPException(404, "PDF bulunamadı")
    try:
        import pdfplumber
        with pdfplumber.open(p) as pdf:
            if sayfa < 1 or sayfa > len(pdf.pages):
                raise HTTPException(404, "Sayfa yok")
            im = pdf.pages[sayfa - 1].to_image(resolution=150).original
            buf = io.BytesIO()
            im.save(buf, format="PNG")
            buf.seek(0)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(500, f"Sayfa görüntüsü oluşturulamadı: {e}")
    return StreamingResponse(buf, media_type="image/png")


# ----------------------------------------------------------------- işleme
class IsleBody(BaseModel):
    firma_kod: str
    tip: str
    dosyalar: list[str] | None = None   # None: eski istemci uyumu (tümü); []: kullanıcı hiçbirini seçmedi
    fis_baslangic: int = 1         # başlangıç fiş no
    banka_hesap_kodu: str = ""     # opsiyonel: banka hesabını elle belirt (örn. 102.01.004)


@app.post("/api/firma/{kod}/isle/{tip}")
def isle(kod: str, tip: str, body: IsleBody):
    """Seçili belgeleri okuyup fiş satırları üretir (önizleme için)."""
    if tip not in TIPLER:
        raise HTTPException(400, "Geçersiz tip")
    d = firma_dir(kod)
    km = KuralMotoru(d / "mizan.xlsx", d / f"kural_{tip}.xlsx", d / f"{tip}_ogrenme.json",
                      d / "banka_hesap_eslestirme.json", d / f"{tip}_gider_eslestirme.json")

    dosyalar = body.dosyalar if body.dosyalar is not None else [f.name for f in (d / tip).iterdir() if f.is_file()]
    tum_ham = []
    for fn in dosyalar:
        p = d / tip / fn
        if not p.exists():
            continue
        okundu = belge_oku(p, fn)
        tum_ham.append({"dosya": fn, **okundu})

    # tip'e göre işleyiciye ver
    fisler, uyarilar = isleyici.isle(tip, tum_ham, km, body.fis_baslangic,
                                      banka_hesap_kodu=body.banka_hesap_kodu)
    tb = round(sum(f["borc"] for f in fisler), 2)
    ta = round(sum(f["alacak"] for f in fisler), 2)
    return {
        "satirlar": fisler,
        "toplam_borc": tb, "toplam_alacak": ta, "dengeli": abs(tb - ta) < 0.01,
        "uyarilar": uyarilar,
        "okunan_belgeler": [{"dosya": h["dosya"], "tur": h["tur"], "uyari": h.get("uyari", "")}
                            for h in tum_ham],
    }


class OgretBody(BaseModel):
    firma_kod: str
    tip: str
    aciklama: str
    kod: str


@app.post("/api/firma/{kod}/ogret/{tip}")
def ogret(kod: str, tip: str, body: OgretBody):
    if tip not in TIPLER:
        raise HTTPException(400, "Geçersiz tip")
    d = firma_dir(kod)
    km = KuralMotoru(d / "mizan.xlsx", d / f"kural_{tip}.xlsx", d / f"{tip}_ogrenme.json")
    km.ogret(body.aciklama, body.kod)
    return {"ok": True, "ogrenilen": len(km.ogrenme)}


class OgretGiderBody(BaseModel):
    firma_kod: str
    cari_kod: str
    gider_kod: str
    tip: str = "fatura"


@app.post("/api/firma/{kod}/ogret-gider")
def ogret_gider(kod: str, body: OgretGiderBody):
    """Fatura (alış) veya Fatura Satış Düzenle modunda bir fişin CARİ veya
    GİDER/GELİR hesabı elle düzeltilince, aynı fişteki cari<->gider(gelir)
    hesap çiftini doğrudan {tip}_gider_eslestirme.json'a yazar — geçmiş fiş
    listesi yüklemeyi beklemeden, düzeltme yapıldıkça öğrenir (bkz.
    index.html edit())."""
    if body.tip not in ("fatura", "fatura_satis"):
        raise HTTPException(400, "Geçersiz tip")
    if not body.cari_kod or not body.gider_kod:
        raise HTTPException(400, "cari_kod ve gider_kod gerekli")
    d = firma_dir(kod)
    gider_path = d / f"{body.tip}_gider_eslestirme.json"
    gider_og = _read_json(gider_path, {})
    gider_og[body.cari_kod] = body.gider_kod
    _write_json(gider_path, gider_og)
    return {"ok": True, "gider_ogrenilen": len(gider_og)}


@app.get("/api/firma/{kod}/hesaplar")
def hesaplar(kod: str):
    """Firmanın tüm hesapları (eşleştirme kutusu için)."""
    d = firma_dir(kod)
    hes = mizan_hesaplar(d / "mizan.xlsx") if (d / "mizan.xlsx").exists() else []
    kod_ad = {k: a for k, a in hes}
    for t in TIPLER:
        kp = d / f"kural_{t}.xlsx"
        if kp.exists():
            for h in kural_excel_oku(kp)["hesaplar"]:
                kod_ad.setdefault(h["kod"], h["ad"])
    return {"hesaplar": [{"kod": k, "ad": a} for k, a in sorted(kod_ad.items())]}


class ExportBody(BaseModel):
    firma_kod: str
    tip: str
    satirlar: list[dict]           # (düzenlenmiş) fiş satırları
    donem: str = ""


@app.post("/api/firma/{kod}/export")
def export(kod: str, body: ExportBody, format: str = "xlsx"):
    d = firma_dir(kod)
    if not body.satirlar:
        raise HTTPException(400, "Satır yok")
    if format == "xml":
        data = isleyici.fis_xml(body.satirlar)
        fname = f"FIS_{body.tip}_{datetime.now():%Y%m%d_%H%M%S}.xml"
        (d / "cikti" / fname).write_bytes(data)
        return StreamingResponse(io.BytesIO(data), media_type="application/xml",
            headers={"Content-Disposition": f'attachment; filename="{fname}"'})
    bio = isleyici.fis_xlsx(body.satirlar)
    fname = f"FIS_{body.tip}_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    (d / "cikti" / fname).write_bytes(bio.getvalue()); bio.seek(0)
    return StreamingResponse(bio,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{fname}"'})


@app.get("/api/firma/{kod}/ciktilar")
def ciktilar(kod: str):
    d = firma_dir(kod)
    cd = d / "cikti"
    if not cd.exists():
        return []
    out = []
    for f in cd.iterdir():
        if f.suffix in (".xlsx", ".xml"):
            st = f.stat()
            out.append({"dosya": f.name,
                        "tarih": datetime.fromtimestamp(st.st_mtime).isoformat(timespec="seconds"),
                        "boyut": st.st_size})
    out.sort(key=lambda x: x["tarih"], reverse=True)
    return out


@app.get("/api/firma/{kod}/cikti/{fname}")
def cikti_indir(kod: str, fname: str):
    d = firma_dir(kod)
    if ".." in fname or "/" in fname or "\\" in fname:
        raise HTTPException(400, "Geçersiz")
    p = d / "cikti" / fname
    if not p.exists():
        raise HTTPException(404, "Yok")
    return FileResponse(p, filename=fname)


# ----------------------------------------------------------------- statik
app.mount("/", StaticFiles(directory=str(Path(__file__).parent / "static"), html=True), name="static")
