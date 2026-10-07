"""
Fiş Aktarım Aracı — bağımsız server
Banka dökümü / Fatura / Çek listesi -> muhasebe fişi (Fiş Aktarım Şablonu).
Her firma için: mizan + TEK kural dosyası (kural.xlsx) + belgeler + öğrenme.

Klasör yapısı:
  data/<firma>/
    mizan.xlsx
    kural.xlsx                      <- SEZGIN gibi tek dosya (banka/fatura/çek ortak)
                                       Sayfalar: Talimatlar, Hesap Kodu Eşleştirme,
                                                 Fiş Aktarım Şablonu (kümülatif)
    banka/  fatura/  cek/          <- gelen belgeler (excel/pdf/resim)
    banka_ogrenme.json  ...        <- öğrenilen eşleştirmeler
    cikti/                          <- üretilen fiş aktarım dosyaları
    meta.json

  Eski kurulumlarla uyum: kural.xlsx yoksa kural_<tip>.xlsx'e düşer.
"""
import io, json, os, re, shutil, unicodedata
from urllib.parse import quote
from datetime import datetime
from pathlib import Path

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse, FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from app.kurallar import KuralMotoru, norm, kural_excel_oku, mizan_hesaplar
from app.belge_oku import belge_oku
from app import isleyici
from app import sunucu_klasor as sk
from app import fatura_pdf, yapay_zeka, gider_yz

DATA_DIR = Path(os.environ.get("FIS_DATA", "/data"))
DATA_DIR.mkdir(parents=True, exist_ok=True)

app = FastAPI(title="Fiş Aktarım Aracı")

TIPLER = ("banka", "fatura", "cek", "fis")


def ortak_fis_listesi_yolu(d: Path) -> Path:
    """Sunucu klasöründen (fislistesi/) alınan, tüm sekmelerde geçerli fiş listesi."""
    return d / "_gecmis_ortak.xlsx"


def fis_listesi_yolu(d: Path, tip: str) -> Path:
    """Firma+tip için geçmiş fiş listesi Excel'inin yolu.
    İki kaynak olabilir: sekmeye elle yüklenen <tip>/_gecmis.xlsx ve sunucu
    klasöründen gelen ortak _gecmis_ortak.xlsx. Hangisi daha YENİ yüklendiyse
    o kullanılır — elle yüklenen sonra geldiyse sunucudakini geçersiz kılar,
    sunucuya daha yeni liste gelirse o devralır.
    Geriye dönük: ikisi de yoksa kural.xlsx veya kural_<tip>.xlsx."""
    yeni = d / tip / "_gecmis.xlsx"
    mevcut = [p for p in (yeni, ortak_fis_listesi_yolu(d)) if p.exists()]
    if mevcut:
        return max(mevcut, key=lambda p: p.stat().st_mtime)
    # geriye dönük
    for eski in [d / "kural.xlsx", d / f"kural_{tip}.xlsx"]:
        if eski.exists():
            return eski
    return yeni  # yoksa yeni yol döndür


# Eski kod uyumluluğu için alias
def kural_yolu(d: Path, tip: str = "banka") -> Path:
    return fis_listesi_yolu(d, tip)


def _gecmis_kaynak(d: Path, tip: str) -> Path | None:
    """Hesap öğrenmesi için geçmiş kayıt kaynağı: Excel fiş listesi / muavin defter.
    (Geçmiş Fiş PDF'i kaldırıldı; eski firmalarda kalan gecmis_fisler.pdf okunmaz.)
    Fatura sekmesinde kural dosyası verilmediği için liste burada verilir;
    banka/çekte None — KuralMotoru kural dosyasındaki (aynı liste) geçmişi okur."""
    return fis_listesi_yolu(d, tip) if tip == "fatura" else None


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


def _kaynak_yaz(d: Path, anahtar: str, tur: str, dosya: str = ""):
    """Bir dosyanın nereden geldiğini meta.json'a not eder.
    anahtar: 'mizan' | 'fis-listesi' (ortak) | 'fis-listesi:<tip>'
    tur: 'sunucu' | 'elle' | 'firma' (başka firmadan kopya)"""
    meta = _read_json(d / "meta.json", {})
    meta.setdefault("kaynak", {})[anahtar] = {
        "tur": tur, "dosya": dosya,
        "zaman": datetime.now().isoformat(timespec="seconds"),
    }
    _write_json(d / "meta.json", meta)


def _fis_listesi_arka_plan(hedef: Path):
    """Fiş listesinin eski önbelleğini silip parse'ı arka planda başlatır."""
    cache_path = hedef.with_suffix(hedef.suffix + ".cache.json")
    if cache_path.exists():
        try: cache_path.unlink()
        except Exception: pass
    import threading
    from app.kurallar import gecmis_fisler_oku
    def _arka_plan_parse():
        try:
            gecmis_fisler_oku(hedef)  # cache'i yazar
        except Exception as e:
            print(f"[fiş listesi arka plan parse hatası] {hedef}: {e}")
    threading.Thread(target=_arka_plan_parse, daemon=True).start()


def _ice_al(d: Path, tur: str, data: bytes, mtime: float | None = None) -> dict:
    """Mizan / ortak fiş listesini firma veri klasörüne yazar.
    Sunucu klasöründen alınan dosyalar buradan geçer; .xls otomatik .xlsx'e çevrilir.

    mtime verilirse kopyaya kaynak dosyanın tarihi işlenir. Böylece "hangisi daha
    yeni" kıyası yükleme anına göre değil BELGENİN tarihine göre yapılır: dün elle
    yüklenen mizanı, geçen haftadan kalma sunucu dosyası otomatik eşitlemede ezmez."""
    def _yaz(hedef: Path, icerik: bytes):
        hedef.write_bytes(icerik)
        if mtime:
            os.utime(hedef, (mtime, mtime))

    if tur == "mizan":
        _yaz(d / "mizan.xlsx", sk.excel_xlsx_bytes(data))
        return {"hesap_sayisi": len(mizan_hesaplar(d / "mizan.xlsx"))}
    if tur == "fis-listesi":
        hedef = ortak_fis_listesi_yolu(d)
        _yaz(hedef, sk.excel_xlsx_bytes(data))
        _fis_listesi_arka_plan(hedef)
        return {"boyut_kb": round(len(data) / 1024, 1)}
    raise HTTPException(400, "Geçersiz tür: mizan | fis-listesi")


def _firma_klasoru(d: Path) -> Path | None:
    """Firmanın bağlı olduğu sunucu klasörü (yoksa None)."""
    return sk.klasor_yolu(_read_json(d / "meta.json", {}).get("klasor", ""))


def _bagli_klasorler() -> dict:
    """klasör adı -> o klasöre bağlı firma kodu"""
    out = {}
    for d in DATA_DIR.iterdir():
        if d.is_dir() and not d.name.startswith("_"):
            k = _read_json(d / "meta.json", {}).get("klasor")
            if k:
                out[k] = d.name
    return out


# ----------------------------------------------------------------- firma yönetimi
@app.get("/api/firmalar")
def firmalar():
    out = []
    for d in sorted(DATA_DIR.iterdir()):
        if not d.is_dir() or d.name.startswith("_"):
            continue
        meta = _read_json(d / "meta.json", {})
        # Her tip için fiş listesi var mı (elle yüklenen veya sunucudan gelen ortak)
        fis_listesi_var = {t: fis_listesi_yolu(d, t).exists() for t in TIPLER}
        # Geriye dönük: eski kural dosyaları varsa onlar da fiş listesi sayılır
        eski_kural_var = (d / "kural.xlsx").exists() or any((d / f"kural_{t}.xlsx").exists() for t in TIPLER)
        out.append({
            "kod": d.name,
            "ad": meta.get("ad", d.name),
            "klasor": meta.get("klasor", ""),
            "mizan_var": (d / "mizan.xlsx").exists(),
            "fis_listesi_var": fis_listesi_var,
            "kural_var": eski_kural_var or any(fis_listesi_var.values()),
        })
    return out


@app.post("/api/firma")
def firma_olustur(body: dict):
    """Yeni firma. body: {ad?, klasor?} — klasor verilirse ad klasörden türetilir
    ve firma o sunucu klasörüne bağlanır."""
    klasor = (body.get("klasor") or "").strip()
    if klasor and not sk.klasor_yolu(klasor):
        raise HTTPException(404, f"Sunucuda '{klasor}' klasörü yok")
    ad = (body.get("ad") or "").strip() or (sk.guzel_ad(klasor) if klasor else "")
    if not ad:
        raise HTTPException(400, "Firma adı gerekli")
    if klasor:
        bagli = _bagli_klasorler().get(klasor)
        if bagli:
            raise HTTPException(409, f"'{klasor}' klasörü zaten '{bagli}' firmasına bağlı")
    kod = slugify(klasor or ad)
    d = DATA_DIR / kod
    d.mkdir(parents=True, exist_ok=True)
    for t in TIPLER:
        (d / t).mkdir(exist_ok=True)
    (d / "cikti").mkdir(exist_ok=True)
    meta = _read_json(d / "meta.json", {})
    meta["ad"] = ad
    if klasor:
        meta["klasor"] = klasor
    meta.setdefault("olusturma", datetime.now().isoformat(timespec="seconds"))
    _write_json(d / "meta.json", meta)
    return {"kod": kod, "ad": ad, "klasor": meta.get("klasor", "")}


# ----------------------------------------------------------------- sunucu firma klasörleri
@app.get("/api/klasorler")
def klasorler_listesi():
    """/srv/veri/firmalar altındaki firma klasörleri ve bağlı oldukları firma."""
    bagli = _bagli_klasorler()
    return {
        "erisim": sk.klasor_var_mi(),
        "klasorler": [{**k, "firma_kod": bagli.get(k["klasor"], "")} for k in sk.klasorler()],
    }


@app.post("/api/firma/{kod}/klasor")
def firma_klasor_bagla(kod: str, body: dict):
    """Firmayı bir sunucu klasörüne bağlar. body: {klasor: str} — boş string bağlantıyı kaldırır."""
    d = firma_dir(kod)
    klasor = (body.get("klasor") or "").strip()
    meta = _read_json(d / "meta.json", {})
    if klasor:
        if not sk.klasor_yolu(klasor):
            raise HTTPException(404, f"Sunucuda '{klasor}' klasörü yok")
        bagli = _bagli_klasorler().get(klasor)
        if bagli and bagli != kod:
            raise HTTPException(409, f"'{klasor}' klasörü zaten '{bagli}' firmasına bağlı")
        meta["klasor"] = klasor
    else:
        meta.pop("klasor", None)
    _write_json(d / "meta.json", meta)
    return {"ok": True, "klasor": meta.get("klasor", "")}


@app.get("/api/firma/{kod}/klasor-dosyalar")
def klasor_dosyalar(kod: str, tur: str = "mizan"):
    """Firmanın sunucu klasöründe verilen tür için bulunan dosyalar (en yeni önce)."""
    d = firma_dir(kod)
    kp = _firma_klasoru(d)
    if not kp:
        return {"klasor": "", "dosyalar": []}
    if tur not in sk.TURLER:
        raise HTTPException(400, "Geçersiz tür")
    secili = _read_json(d / "meta.json", {}).get("kaynak", {}).get(tur, {})
    return {"klasor": kp.name, "dosyalar": sk.adaylar(kp, tur),
            "secili": secili.get("dosya", "") if secili.get("tur") == "sunucu" else ""}


@app.post("/api/firma/{kod}/klasordan-al")
def klasordan_al(kod: str, body: dict):
    """Sunucu klasöründeki belirli bir dosyayı içe alır.
    body: {tur: 'mizan'|'fis-listesi', dosya: '<firma klasörüne göre yol>'}"""
    d = firma_dir(kod)
    kp = _firma_klasoru(d)
    if not kp:
        raise HTTPException(400, "Firma bir sunucu klasörüne bağlı değil")
    tur = body.get("tur", "")
    try:
        src = sk.guvenli_dosya(kp, body.get("dosya", ""))
    except ValueError as e:
        raise HTTPException(400, str(e))
    sonuc = _ice_al(d, tur, src.read_bytes())
    _kaynak_yaz(d, tur, "sunucu", src.relative_to(kp.resolve()).as_posix())
    return {"ok": True, "tur": tur, "dosya": src.name, **sonuc}


@app.post("/api/firma/{kod}/sunucu-esitle")
def sunucu_esitle(kod: str):
    """Firma seçildiğinde çağrılır. Sunucu klasöründe mizan / fiş listesi
    için araçtakinden DAHA YENİ dosya varsa otomatik içe alır.
    Elle yüklenen dosya sunucudakinden yeniyse dokunulmaz."""
    d = firma_dir(kod)
    kp = _firma_klasoru(d)
    if not kp:
        return {"klasor": "", "alinan": [], "yanlis_yer": []}

    # Mizanda elle yükleme de aynı dosyaya yazar, kıyas doğrudan onunla.
    # Fiş listesinde sunucu kopyası ortak dosyadır; sekmeye elle yüklenen liste
    # ayrı durur ve fis_listesi_yolu() ikisinden tarihi yeni olanı seçer.
    mevcut = {
        "mizan": d / "mizan.xlsx",
        "fis-listesi": ortak_fis_listesi_yolu(d),
    }
    alinan, yanlis_yer = [], []
    for tur in sk.TURLER:
        liste = sk.adaylar(kp, tur)
        yanlis_yer += [{"tur": tur, "yol": a["yol"]} for a in liste if not a["dogru_yer"]]
        if not liste:
            continue
        aday, hedef = liste[0], mevcut[tur]
        if hedef.exists() and aday["mtime"] <= hedef.stat().st_mtime:
            continue        # araçtaki dosya zaten aynı ya da daha yeni
        try:
            src = sk.guvenli_dosya(kp, aday["yol"])
            sonuc = _ice_al(d, tur, src.read_bytes(), mtime=aday["mtime"])
        except Exception as e:
            yanlis_yer.append({"tur": tur, "yol": aday["yol"], "hata": str(e)})
            continue
        _kaynak_yaz(d, tur, "sunucu", aday["yol"])
        alinan.append({"tur": tur, "dosya": aday["dosya"], "yol": aday["yol"], **sonuc})
    return {"klasor": kp.name, "alinan": alinan, "yanlis_yer": yanlis_yer}


@app.delete("/api/firma/{kod}")
def firma_sil(kod: str):
    """Firmayı ve tüm verilerini kalıcı olarak siler."""
    d = firma_dir(kod)
    shutil.rmtree(d)
    return {"ok": True, "silinen": kod}


# ----------------------------------------------------------------- sunucudan dosya listele & kopyala
@app.get("/api/sunucu-dosyalar")
def sunucu_dosyalar(tip: str = "mizan"):
    """Sunucudaki tüm firmaların belirli tip dosyalarını listeler.
    tip: mizan | fis-listesi
    Firma adı, dosya boyutu ve tarih bilgisiyle döner."""
    sonuc = []
    for d in sorted(DATA_DIR.iterdir()):
        if not d.is_dir() or d.name.startswith("_"):
            continue
        meta = _read_json(d / "meta.json", {})
        firma_ad = meta.get("ad", d.name)

        if tip == "mizan":
            p = d / "mizan.xlsx"
            if p.exists():
                st = p.stat()
                sonuc.append({"firma_kod": d.name, "firma_ad": firma_ad,
                              "dosya": "mizan.xlsx", "yol": str(p),
                              "boyut_kb": round(st.st_size / 1024, 1),
                              "tarih": datetime.fromtimestamp(st.st_mtime).isoformat(timespec="seconds")})
        elif tip == "fis-listesi":
            gorulen = set()
            for t in TIPLER:
                p = fis_listesi_yolu(d, t)      # o sekmede geçerli liste
                if not p.exists() or p in gorulen:
                    continue
                gorulen.add(p)
                st = p.stat()
                sonuc.append({"firma_kod": d.name, "firma_ad": firma_ad,
                              "dosya": p.name, "tip": t, "yol": str(p),
                              "boyut_kb": round(st.st_size / 1024, 1),
                              "tarih": datetime.fromtimestamp(st.st_mtime).isoformat(timespec="seconds")})
    return sonuc


@app.post("/api/firma/{kod}/sunucudan-kopyala")
def sunucudan_kopyala(kod: str, body: dict):
    """Sunucudaki başka bir firmanın dosyasını bu firmaya kopyalar.
    body: {tip: 'mizan'|'fis-listesi', kaynak_firma: str, kaynak_tip?: str}
    """
    d = firma_dir(kod)
    tip = body.get("tip", "")
    kaynak = body.get("kaynak_firma", "")
    if not kaynak:
        raise HTTPException(400, "kaynak_firma gerekli")
    kaynak_d = DATA_DIR / kaynak
    if not kaynak_d.exists():
        raise HTTPException(404, "Kaynak firma yok")

    if tip == "mizan":
        src = kaynak_d / "mizan.xlsx"
        if not src.exists():
            raise HTTPException(404, "Kaynak firmada mizan yok")
        shutil.copy(src, d / "mizan.xlsx")      # copy: tarih = kopyalama anı
        hes = mizan_hesaplar(d / "mizan.xlsx")
        _kaynak_yaz(d, "mizan", "firma", kaynak)
        return {"ok": True, "tip": "mizan", "kaynak": kaynak, "hesap_sayisi": len(hes)}

    elif tip == "fis-listesi":
        kaynak_tip = body.get("kaynak_tip", "banka")
        src = fis_listesi_yolu(kaynak_d, kaynak_tip)
        if not src.exists():
            raise HTTPException(404, f"Kaynak firmada {kaynak_tip} fiş listesi yok")
        (d / kaynak_tip).mkdir(exist_ok=True)
        hedef = d / kaynak_tip / "_gecmis.xlsx"
        shutil.copy(src, hedef)
        _fis_listesi_arka_plan(hedef)
        _kaynak_yaz(d, f"fis-listesi:{kaynak_tip}", "firma", kaynak)
        return {"ok": True, "tip": "fis-listesi", "kaynak": kaynak, "kaynak_tip": kaynak_tip,
                "boyut_kb": round(hedef.stat().st_size / 1024, 1)}

    else:
        raise HTTPException(400, "Geçersiz tip: mizan | fis-listesi")


# ----------------------------------------------------------------- mizan & kural yükleme
@app.post("/api/firma/{kod}/mizan")
async def mizan_yukle(kod: str, file: UploadFile = File(...)):
    d = firma_dir(kod)
    sonuc = _ice_al(d, "mizan", await file.read())
    _kaynak_yaz(d, "mizan", "elle", file.filename or "")
    return {"ok": True, **sonuc}


@app.post("/api/firma/{kod}/fis-listesi/{tip}")
async def fis_listesi_yukle(kod: str, tip: str, file: UploadFile = File(...)):
    """Firma+tip için Fiş Aktarım Şablonu (veya Yevmiye Defteri) formatında
    Excel yüklenir. Bu Excel; geçmiş fiş no'ları, hesap eşleştirmelerini
    ve öğrenmeyi sağlar. Yer: <firma>/<tip>/_gecmis.xlsx

    Parse arka planda başlar; kullanıcı beklemez. Sonuç ilk 'İşle' tıklamasında
    hazır cache'ten alınır (büyük yevmiye defterleri için 3+ sn beklemek yerine
    yükleme anında ~100 ms).
    """
    if tip not in TIPLER:
        raise HTTPException(400, "Geçersiz tip")
    d = firma_dir(kod)
    (d / tip).mkdir(exist_ok=True)
    data = await file.read()
    hedef = d / tip / "_gecmis.xlsx"
    hedef.write_bytes(sk.excel_xlsx_bytes(data))   # .xls de kabul edilir
    _kaynak_yaz(d, f"fis-listesi:{tip}", "elle", file.filename or "")
    # eski kural dosyalarını arşivle
    if (d / "kural.xlsx").exists():
        (d / "kural.xlsx").rename(d / "kural.xlsx.eski")
    for t in TIPLER:
        eski = d / f"kural_{t}.xlsx"
        if eski.exists():
            eski.rename(d / f"kural_{t}.xlsx.eski")

    # Parse'ı arka plana at — kullanıcı beklemez
    _fis_listesi_arka_plan(hedef)

    return {"ok": True, "gecmis_satir": None, "son_fis_no": None,
            "ogrenilen_eslesme": None,
            "durum": "arka planda işleniyor",
            "boyut_kb": round(len(data) / 1024, 1)}


# Geriye dönük: eski /kural ucunu da destekle
@app.post("/api/firma/{kod}/kural")
async def kural_yukle_eski(kod: str, file: UploadFile = File(...)):
    """[ESKİ] Kural yüklemesi. Yeni akış /fis-listesi/{tip} kullanmalı.
    Bu uç eski istemciler için tutuldu; yüklenen dosyayı banka+fatura+cek için
    _gecmis.xlsx olarak kopyalar."""
    d = firma_dir(kod)
    data = await file.read()
    for t in TIPLER:
        (d / t).mkdir(exist_ok=True)
        (d / t / "_gecmis.xlsx").write_bytes(data)
    from app.kurallar import gecmis_fisler_oku
    g = gecmis_fisler_oku(d / "banka" / "_gecmis.xlsx")
    return {"ok": True, "hesap_kurali": 0,
            "talimat": 0,
            "gecmis_satir": len(g.get("satirlar", []))}


@app.post("/api/firma/{kod}/kural/{tip}")
async def kural_yukle_tipli_eski(kod: str, tip: str, file: UploadFile = File(...)):
    """[ESKİ] Tip'li kural yüklemesi; artık /fis-listesi/{tip}'e yönleniyor."""
    return await fis_listesi_yukle(kod, tip, file)


@app.get("/api/firma/{kod}/durum")
def firma_durum(kod: str):
    from app.kurallar import gecmis_fisler_oku
    d = firma_dir(kod)
    meta = _read_json(d / "meta.json", {})
    hes = mizan_hesaplar(d / "mizan.xlsx") if (d / "mizan.xlsx").exists() else []
    kp = _firma_klasoru(d)
    out = {
        "mizan_hesap": len(hes), "tipler": {},
        "klasor": kp.name if kp else "",
        # bağlı değilse ada en çok benzeyen sunucu klasörü (arayüz önerir, kendisi bağlamaz)
        "klasor_oneri": "" if kp else sk.klasor_oner(meta.get("ad", ""), kod),
        "kaynak": meta.get("kaynak", {}),
    }
    for t in TIPLER:
        yol = fis_listesi_yolu(d, t)
        g = gecmis_fisler_oku(yol) if yol.exists() else {"satirlar": [], "son_fis_no": 0, "eslesmeler": {}}
        # belgeler: _gecmis.xlsx dışındaki dosyalar
        belgeler = []
        if (d / t).exists():
            for f in (d / t).iterdir():
                if f.is_file() and not f.name.startswith("_"):
                    belgeler.append(f.name)
        og = _read_json(d / f"{t}_ogrenme.json", {})
        out["tipler"][t] = {
            "fis_listesi_var": yol.exists(),
            # bu sekmede hangi liste geçerli: 'ortak' (sunucu) | 'sekme' (elle) | 'eski' (kural.xlsx)
            "fis_listesi_kaynak": ("" if not yol.exists() else
                                   "ortak" if yol == ortak_fis_listesi_yolu(d) else
                                   "sekme" if yol.name == "_gecmis.xlsx" else "eski"),
            "gecmis_satir": len(g.get("satirlar", [])),
            "son_fis_no": g.get("son_fis_no", 0),
            "belge_sayisi": len(belgeler),
            "belgeler": belgeler,
            "ogrenilen": len(og),
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


@app.get("/api/firma/{kod}/belge/{tip}/{fname}")
def belge_ac(kod: str, tip: str, fname: str):
    """2. bölüme yüklenmiş belgeyi tarayıcıda açar (önizlemedeki 📄 bağlantısı)."""
    if tip not in TIPLER:
        raise HTTPException(400, "Geçersiz")
    p = firma_dir(kod) / tip / _guvenli_ad(fname)
    if not p.is_file():
        raise HTTPException(404, "Yok")
    tur = "application/pdf" if p.suffix.lower() == ".pdf" else None
    return FileResponse(p, media_type=tur,
                        headers={"Content-Disposition": f"inline; filename*=UTF-8''{quote(p.name)}"})


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


# ----------------------------------------------------------------- fatura PDF'leri (alış/satış)
def _yon_kontrol(yon: str):
    if yon not in fatura_pdf.YONLER:
        raise HTTPException(400, "yon alis veya satis olmalı")


def _guvenli_ad(fname: str) -> str:
    if not fname or ".." in fname or "/" in fname or "\\" in fname or fname.startswith("_"):
        raise HTTPException(400, "Geçersiz dosya adı")
    return fname


@app.post("/api/firma/{kod}/fatura-pdf/{yon}")
async def fatura_pdf_yukle(kod: str, yon: str, file: UploadFile = File(...)):
    """Tek bir fatura PDF'i yükler, hemen metinden okur; eksik alan kalırsa
    yapay zekâ kuyruğuna atar (arka planda)."""
    _yon_kontrol(yon)
    d = firma_dir(kod)
    data = await file.read()
    if data[:5] != b"%PDF-":
        raise HTTPException(400, f"{file.filename}: PDF değil")
    k = fatura_pdf.klasor(d, yon)
    fname = re.sub(r"[^\w.\- ]", "_", file.filename or "fatura.pdf").lstrip("_") or "fatura.pdf"
    if not fname.lower().endswith(".pdf"):
        fname += ".pdf"
    (k / fname).write_bytes(data)
    sunucudan = _read_json(k / "_sunucudan.json", [])
    if fname in sunucudan:          # elle yüklendi: artık normal fatura PDF'i
        _write_json(k / "_sunucudan.json", [x for x in sunucudan if x != fname])
    kayit = fatura_pdf.oku(k, k / fname, yon)
    return {"ok": True, "dosya": fname, "durum": fatura_pdf._durum_hesapla(kayit),
            "fatura_sayisi": len(kayit.get("faturalar") or [])}


@app.get("/api/firma/{kod}/fatura-pdf/{yon}")
def fatura_pdf_liste(kod: str, yon: str):
    _yon_kontrol(yon)
    d = firma_dir(kod)
    return {"dosyalar": fatura_pdf.liste(fatura_pdf.klasor(d, yon), yon),
            "yapay_zeka": yapay_zeka.durum(), "kuyruk": fatura_pdf.kuyruk_bilgisi(),
            "pdf_otomatik": fatura_pdf.YZ_OTOMATIK}


@app.get("/api/firma/{kod}/fatura-pdf/{yon}/{fname}")
def fatura_pdf_ac(kod: str, yon: str, fname: str):
    """Fatura PDF'ini tarayıcıda açar (önizlemedeki 📄 bağlantısı)."""
    _yon_kontrol(yon)
    d = firma_dir(kod)
    p = fatura_pdf.klasor(d, yon) / _guvenli_ad(fname)
    if not p.is_file():
        raise HTTPException(404, "PDF yok")
    return FileResponse(p, media_type="application/pdf",
                        headers={"Content-Disposition": f"inline; filename*=UTF-8''{quote(p.name)}"})


@app.delete("/api/firma/{kod}/fatura-pdf/{yon}/{fname}")
def fatura_pdf_sil(kod: str, yon: str, fname: str):
    _yon_kontrol(yon)
    d = firma_dir(kod)
    fatura_pdf.sil(fatura_pdf.klasor(d, yon), _guvenli_ad(fname))
    return {"ok": True}


@app.post("/api/firma/{kod}/fatura-pdf/{yon}/{fname}/yapay-zeka")
def fatura_pdf_yz(kod: str, yon: str, fname: str):
    """Eksik kalan bir PDF'i yapay zekâya (yeniden) gönderir."""
    _yon_kontrol(yon)
    d = firma_dir(kod)
    k = fatura_pdf.klasor(d, yon)
    fname = _guvenli_ad(fname)
    if not (k / fname).exists():
        raise HTTPException(404, "Dosya yok")
    if not fatura_pdf.yz_kuyruga_al(k, fname, yon, zorla=True):
        raise HTTPException(503, "Yapay zekâ kullanılamıyor: " + yapay_zeka.durum(tazele=True)["hata"])
    return {"ok": True}


@app.post("/api/firma/{kod}/yz-sinav/{yon}")
def yz_sinav_baslat(kod: str, yon: str):
    """Yapay zekâ sınavı: muavindeki carileri tek tek saklayıp gider hesabını sorar,
    cevapları muavindeki gerçek kayıtla karşılaştırır (arka planda)."""
    _yon_kontrol(yon)
    d = firma_dir(kod)
    if not yapay_zeka.durum(tazele=True)["etkin"]:
        raise HTTPException(503, "Yapay zekâ kullanılamıyor: " + (yapay_zeka.durum()["hata"] or "kapalı"))
    km = KuralMotoru(d / "mizan.xlsx", None, d / "fatura_ogrenme.json", _gecmis_kaynak(d, "fatura"))
    if not km.hesaplar:
        raise HTTPException(400, "Önce mizanı yükleyin")
    if not km.gecmis.get("satirlar"):
        raise HTTPException(400, "Önce 1. bölümden muavin defteri (ya da Excel fiş listesini) yükleyin")
    v = gider_yz.sinav_baslat(d, km.hesaplar, isleyici.alt_hesap_kodlari(km.hesaplar), yon,
                              km.gecmis["satirlar"])
    if v.get("durum") == "yok":
        raise HTTPException(400, v.get("hata", "Soru yok"))
    return v


@app.get("/api/firma/{kod}/yz-sinav/{yon}")
def yz_sinav_durum(kod: str, yon: str):
    _yon_kontrol(yon)
    return gider_yz.sinav_oku(firma_dir(kod), yon)


@app.get("/api/yapay-zeka")
def yapay_zeka_durum():
    return {**yapay_zeka.durum(tazele=True), "kuyruk": fatura_pdf.kuyruk_bilgisi()}


# ----------------------------------------------------------------- işleme
@app.get("/api/firma/{kod}/onerilen-fisno/{tip}")
def onerilen_fisno(kod: str, tip: str):
    """Geçmiş fişlerdeki son numaradan +1 önerir."""
    if tip not in TIPLER:
        raise HTTPException(400, "Geçersiz tip")
    d = firma_dir(kod)
    from app.kurallar import gecmis_fisler_oku
    g = gecmis_fisler_oku(kural_yolu(d, tip))
    son = g.get("son_fis_no", 0)
    return {"son_fis_no": son, "onerilen": son + 1, "gecmis_satir": len(g.get("satirlar", []))}


def _sunucu_fatura_pdf_al(d: Path, tum_ham: list, yon: str) -> dict:
    """Listedeki fatura numaralarıyla eşleşen PDF'leri firmanın sunucu klasöründen
    (fatura/ altı) bulup fatura PDF'leri alanına kopyalar. Böylece listede olmayan
    KDV dağılımı PDF'ten tamamlanır ve önizlemede fatura açılabilir.
    Yalnızca listedeki numaralar alınır (klasördeki her PDF fişe eklenmesin)."""
    kp = _firma_klasoru(d)
    sonuc = {"bagli": bool(kp), "klasor": kp.name if kp else "", "mesaj": "", "alinan": 0,
             "bulunamayan": [], "dizin": {}}
    if not kp:
        return sonuc
    metin = json.dumps(tum_ham, ensure_ascii=False, default=str).upper()
    nolar = set(sk.FATURA_NO_RE.findall(metin))
    if not nolar:
        return sonuc
    k = fatura_pdf.klasor(d, yon)
    zaten = {f.get("fatura_no") for f in fatura_pdf.faturalar(d, yon)[0]}
    eksik = nolar - zaten
    dizin, sonuc["dizin"] = sk.fatura_pdf_dizini(kp, d / "fatura" / "_sunucu_pdf_dizini.json")
    sonuc["dizin"]["eslesen_numara"] = len(nolar & set(dizin))
    if not eksik:
        return sonuc
    sunucudan = set(_read_json(k / "_sunucudan.json", []))
    alinan = []
    for no in sorted(eksik):
        rel = dizin.get(no)
        if not rel:
            sonuc["bulunamayan"].append(no)
            continue
        src = sk.guvenli_dosya(kp, rel)
        hedef = k / src.name
        if hedef.exists() and hedef.stat().st_size != src.stat().st_size:
            hedef = k / f"{no}_{src.name}"
        if not hedef.exists():
            shutil.copy2(src, hedef)
        alinan.append(no)
        sunucudan.add(hedef.name)
    _write_json(k / "_sunucudan.json", sorted(sunucudan))
    sonuc["alinan"] = len(alinan)
    if alinan:
        sonuc["mesaj"] = f"Sunucu klasöründen {len(alinan)} fatura PDF'i eşleşip alındı (KDV dağılımı ve fatura görüntüsü için)"
    if sonuc["dizin"].get("ocr_bekleyen"):
        sonuc["mesaj"] = (sonuc["mesaj"] + " · " if sonuc["mesaj"] else "") + \
            f"sunucu klasöründe {sonuc['dizin']['ocr_bekleyen']} taranmış PDF okunuyor (OCR) — birkaç dakika sonra tekrar deneyin"
    return sonuc


class KarsilastirBody(BaseModel):
    dosyalar: list[str] | None = None     # None: klasördeki tüm belgeler


@app.post("/api/firma/{kod}/karsilastir/{yon}")
def karsilastir(kod: str, yon: str, body: KarsilastirBody):
    """Fatura listesi (Excel) ile fatura PDF'lerinin matrah/KDV tutarlarını
    %1 / %10 / %20 bazında karşılaştırır (önizleme/fiş üretmeden)."""
    import copy
    _yon_kontrol(yon)
    d = firma_dir(kod)
    klasor = d / "fatura"
    dosyalar = body.dosyalar if body.dosyalar is not None else \
        [f.name for f in klasor.iterdir() if f.is_file() and not f.name.startswith("_")] if klasor.exists() else []
    tum_ham = []
    for fn in dosyalar:
        p = klasor / fn
        if p.exists():
            tum_ham.append({"dosya": fn, **belge_oku(p, fn)})
    try:
        sunucu = _sunucu_fatura_pdf_al(d, tum_ham, yon)
    except Exception as e:
        sunucu = {"hata": str(e), "dizin": {}}
    pdfler, bekleyen = fatura_pdf.faturalar(d, yon)
    sunucudan = set(_read_json(fatura_pdf.klasor(d, yon) / "_sunucudan.json", []))
    for pf in pdfler:
        pf["pdf_yer"] = "alan"
        if pf.get("dosya") in sunucudan:
            pf["sunucudan"] = True
    teshis = {}
    liste = isleyici._elogo_fatura_satirlari(copy.deepcopy(tum_ham), yon, teshis=teshis)
    ek_pdf = isleyici._pdf_gercek_faturalar(copy.deepcopy(tum_ham), yon) if liste else []
    sonuc = isleyici.liste_pdf_karsilastir(liste, pdfler + ek_pdf)
    sonuc["liste_var"] = bool(liste)
    sonuc["pdf_sayisi"] = len(pdfler) + len(ek_pdf)
    sonuc["yz_bekleyen"] = bekleyen
    # neden boş kaldığını anlatmak için: listede hangi sütunlar tanındı, PDF'ler nereden aranıyor
    kdv_anahtar = ("kdv_1", "kdv_8", "kdv_10", "kdv_18", "kdv_20", "mat_1", "mat_8", "mat_10",
                   "mat_18", "mat_20", "kdv_top", "mat_toplam")
    for t in teshis.get("tablolar", []):
        t["kdv_sutunu_var"] = any(k in t["taninan"] for k in kdv_anahtar)
    kp = _firma_klasoru(d)
    sonuc["teshis"] = {
        "tablolar": teshis.get("tablolar", []),
        "klasor": kp.name if kp else "",
        "klasor_fatura_var": bool(kp and ((kp / "fatura").is_dir() or (sunucu.get("dizin") or {}).get("taranan"))),
        "klasor_pdf_sayisi": (sunucu.get("dizin") or {}).get("taranan", 0),
        "klasor_numarali_pdf": (sunucu.get("dizin") or {}).get("numarali", 0),
        "klasor_ocr_bekleyen": (sunucu.get("dizin") or {}).get("ocr_bekleyen", 0),
        "klasor_numarasiz": (sunucu.get("dizin") or {}).get("numarasiz", [])[:20],
        "klasor_eslesen": (sunucu.get("dizin") or {}).get("eslesen_numara", 0),
        "sunucu_mesaj": sunucu.get("mesaj") or sunucu.get("hata", ""),
        "alan_pdf_sayisi": len(pdfler),
    }
    return sonuc


class IsleBody(BaseModel):
    firma_kod: str
    tip: str
    # işlenecek belgeler. None (gönderilmezse): klasördeki tümü. Arayüz yalnızca bu sayfa
    # açıkken yüklenenleri gönderir; boş liste = hiç belge yok (önceki oturumun dosyaları işlenmez).
    dosyalar: list[str] | None = None
    fis_baslangic: int = 0         # 0 ise otomatik (son+1)
    gecmis_ekle: bool = False      # çıktıya geçmiş fişleri de kat
    yon: str = "alis"              # fatura için: alis | satis


@app.post("/api/firma/{kod}/isle/{tip}")
def isle(kod: str, tip: str, body: IsleBody):
    """Seçili belgeleri okuyup fiş satırları üretir (önizleme için)."""
    if tip not in TIPLER:
        raise HTTPException(400, "Geçersiz tip")
    d = firma_dir(kod)
    km = KuralMotoru(d / "mizan.xlsx", None if tip == "fatura" else kural_yolu(d, tip), d / f"{tip}_ogrenme.json", _gecmis_kaynak(d, tip))

    dosyalar = body.dosyalar if body.dosyalar is not None else \
        [f.name for f in (d / tip).iterdir() if f.is_file() and not f.name.startswith("_")]
    tum_ham = []
    for fn in dosyalar:
        p = d / tip / fn
        if not p.exists():
            continue
        okundu = belge_oku(p, fn)
        tum_ham.append({"dosya": fn, **okundu})

    # fiş başlangıç: 0/negatifse otomatik (geçmiş son+1)
    fis_bas = body.fis_baslangic
    if fis_bas <= 0:
        fis_bas = km.son_fis_no() + 1

    # tip'e göre işleyiciye ver
    pdf_faturalar, yz_bekleyen = [], 0
    sunucu_pdf_not = ""
    if tip == "fatura" and body.yon in fatura_pdf.YONLER and tum_ham:
        try:
            sunucu_pdf_not = _sunucu_fatura_pdf_al(d, tum_ham, body.yon).get("mesaj", "")
        except Exception as e:
            sunucu_pdf_not = f"Sunucu klasöründe fatura PDF'leri aranamadı: {e}"
    if tip == "fatura" and body.yon in fatura_pdf.YONLER:
        pdf_faturalar, yz_bekleyen = fatura_pdf.faturalar(d, body.yon)
        # sunucu klasöründen otomatik alınan PDF'ler yalnızca listedeki faturayı tamamlar;
        # listede yoksa (ör. geçen ayın listesinden kalan) fişe eklenmez
        sunucudan = set(_read_json(fatura_pdf.klasor(d, body.yon) / "_sunucudan.json", []))
        for pf in pdf_faturalar:
            if pf.get("dosya") in sunucudan:
                pf["sunucudan"] = True
    gider_ogrenme, gider_onerici = None, None
    if tip == "fatura":
        gider_ogrenme = _read_json(d / "fatura_gider_ogrenme.json", {})
        gecmis_s = km.gecmis.get("satirlar", [])
        gider_onerici = gider_yz.onerici(d, km.hesaplar, isleyici.alt_hesap_kodlari(km.hesaplar), body.yon,
                                         gider_yz.kullanim_ozeti(gecmis_s, body.yon),
                                         gider_yz.gecmis_ornekler(gecmis_s, body.yon))
    fisler, uyarilar = isleyici.isle(tip, tum_ham, km, fis_bas, yon=body.yon, pdf_faturalar=pdf_faturalar,
                                     gider_ogrenme=gider_ogrenme, gider_onerici=gider_onerici)
    if sunucu_pdf_not:
        uyarilar.append(sunucu_pdf_not)
    if yz_bekleyen:
        uyarilar.insert(0, f"{yz_bekleyen} PDF hâlâ yapay zekâ ile okunuyor — bitince tekrar İşle'ye basın")
    if not km.hesaplar:
        uyarilar.insert(0, "Bu firmada mizan yüklü değil ya da okunamadı — hesaplar eşleştirilemedi. "
                           "Önce 1. bölümden mizanı yükle; bu önizlemeyi aktarma.")
    if km.hesaplar and fisler and not km.gecmis.get("satirlar"):
        uyarilar.insert(0, "Bu firmada geçmiş kayıt (Excel fiş listesi / muavin defter) yüklü değil — "
                           "hesap kodları tahmin, fiş no 1'den başladı ve çift kayıt denetimi yapılamadı. "
                           "1. bölümden firmanın muavin defterini yükleyip tekrar İşle'ye bas.")
    if tip == "fatura":
        # Karşı taraf firmanın kendisiyse fatura büyük ihtimalle yanlış sekmede
        # (alış faturası satışa ya da tersi): satışta cari = alıcı = biz olurdu.
        from app.kurallar import kelimeler
        firma_k = kelimeler(_read_json(d / "meta.json", {}).get("ad", ""))
        if firma_k:
            kendisi = sorted({s["evrak_no"] for s in fisler
                              if len(firma_k & kelimeler(s.get("detay", ""))) / len(firma_k) >= 0.6})
            if kendisi:
                diger = "alış" if body.yon == "satis" else "satış"
                uyarilar.insert(0, f"Şu faturalarda karşı taraf firmanın kendisi görünüyor — "
                                   f"{diger} sekmesine ait olabilir: {', '.join(kendisi[:6])}"
                                   + (f" ve {len(kendisi) - 6} fatura daha" if len(kendisi) > 6 else ""))

    # geçmiş fişleri çıktıya kat (istenirse)
    gecmis_satir = []
    if body.gecmis_ekle:
        for gs in km.gecmis.get("satirlar", []):
            gecmis_satir.append({
                "fisno": gs["fisno"], "fis_tarih": "", "fis_aciklama": gs.get("detay", ""),
                "hesap": gs["hesap"], "evrak_no": "", "evrak_tarih": "",
                "detay": gs.get("detay", ""), "borc": gs["borc"], "alacak": gs["alacak"],
                "belge_turu": "MF", "kaynak": "gecmis_kayit",
            })
    tum_satir = gecmis_satir + fisler

    uyarilar = list(dict.fromkeys(uyarilar))
    tb = round(sum(f["borc"] for f in tum_satir), 2)
    ta = round(sum(f["alacak"] for f in tum_satir), 2)
    return {
        "satirlar": tum_satir,
        "yeni_satir_sayisi": len(fisler),
        "gecmis_satir_sayisi": len(gecmis_satir),
        "toplam_borc": tb, "toplam_alacak": ta, "dengeli": abs(tb - ta) < 0.01,
        "uyarilar": uyarilar,
        "kullanilan_fis_bas": fis_bas,
        "okunan_belgeler": [{"dosya": h["dosya"], "tur": h["tur"], "uyari": h.get("uyari", "")}
                            for h in tum_ham],
    }


class FisIsleBody(BaseModel):
    """Fiş sekmesi (elle giriş) için istek gövdesi."""
    firma_kod: str
    kalemler: list[dict] = []           # elle girilen kalemler
    karsi_hesap: str = "198.01.001"     # 198 karşı hesap
    fis_baslangic: int = 0              # 0 = otomatik


@app.post("/api/firma/{kod}/isle-fis")
def isle_fis_uc(kod: str, body: FisIsleBody):
    """Fiş sekmesi: elle girilen kalemler → muhasebe fişi satırları (önizleme)."""
    d = firma_dir(kod)
    tip = "fis"
    km = KuralMotoru(d / "mizan.xlsx", None, d / f"{tip}_ogrenme.json",
                     fis_listesi_yolu(d, tip))
    fis_bas = body.fis_baslangic if body.fis_baslangic > 0 else (km.son_fis_no() + 1)
    fisler, uyarilar = isleyici.isle_fis(body.kalemler, km, fis_bas, karsi_hesap=body.karsi_hesap)
    tb = round(sum(f["borc"] for f in fisler), 2)
    ta = round(sum(f["alacak"] for f in fisler), 2)
    return {
        "satirlar": fisler,
        "toplam_borc": tb, "toplam_alacak": ta, "dengeli": abs(tb - ta) < 0.01,
        "uyarilar": uyarilar,
        "kullanilan_fis_bas": fis_bas,
    }


@app.get("/api/firma/{kod}/karsi-hesaplar")
def karsi_hesaplar(kod: str):
    """Fiş sekmesi için 198 karşı hesap listesi + varsayılan (ilk alt kırılım)."""
    d = firma_dir(kod)
    if not (d / "mizan.xlsx").exists():
        return {"h198": [], "varsayilan198": ""}
    hes = mizan_hesaplar(d / "mizan.xlsx")
    h198 = [{"kod": k, "ad": a} for k, a in hes if k.startswith("198")]
    # varsayılan: en detaylı ilk 198 alt hesabı
    altlar = [h for h in h198 if h["kod"].count(".") >= 2]
    if not altlar:
        altlar = [h for h in h198 if "." in h["kod"]]
    varsayilan = altlar[0]["kod"] if altlar else (h198[0]["kod"] if h198 else "")
    return {"h198": h198, "varsayilan198": varsayilan}


class OgretBody(BaseModel):
    firma_kod: str
    tip: str
    aciklama: str
    kod: str
    rol: str = ""          # fatura satırının rolü: gider | cari | kdv | diger
    oran: int = 0          # KDV satırı için oran (KDV hesabı orana göre öğrenilir)
    yon: str = "alis"      # fatura için alis | satis


@app.post("/api/firma/{kod}/ogret/{tip}")
def ogret(kod: str, tip: str, body: OgretBody):
    if tip not in TIPLER:
        raise HTTPException(400, "Geçersiz tip")
    d = firma_dir(kod)
    if tip == "fatura" and body.rol in ("gider", "kdv", "diger"):
        # Faturada bütün satırların açıklaması cari adıdır. Gider satırındaki
        # düzeltme "cari adı -> kod" diye genel öğrenmeye yazılırsa cari
        # eşleştirmesini bozar ve gider hiç hatırlanmaz. Gider kendi dosyasına,
        # cari + yön anahtarıyla öğrenilir; KDV/diğer satırlar öğrenilmez.
        p = d / "fatura_gider_ogrenme.json"
        og = _read_json(p, {})
        if body.rol == "kdv":
            # KDV hesabı cariye değil ORANA bağlıdır: firma genelinde bu oranın hesabı
            if not body.oran:
                return {"ok": True, "ogrenilen": 0, "not": "oran bilinmiyor"}
            og[f"kdv|{body.yon}|{body.oran}"] = body.kod
            _write_json(p, og)
            return {"ok": True, "ogrenilen": len(og), "tur": "kdv"}
        if body.rol != "gider":
            return {"ok": True, "ogrenilen": 0, "not": "bu satır türü öğrenilmez"}
        og[f"{body.yon}|{norm(body.aciklama)}"] = body.kod
        _write_json(p, og)
        return {"ok": True, "ogrenilen": len(og), "tur": "gider"}
    km = KuralMotoru(d / "mizan.xlsx", None if tip == "fatura" else kural_yolu(d, tip), d / f"{tip}_ogrenme.json", _gecmis_kaynak(d, tip))
    km.ogret(body.aciklama, body.kod)
    return {"ok": True, "ogrenilen": len(km.ogrenme)}


@app.get("/api/firma/{kod}/hesaplar")
def hesaplar(kod: str):
    """Firmanın tüm hesapları (eşleştirme kutusu için)."""
    d = firma_dir(kod)
    hes = mizan_hesaplar(d / "mizan.xlsx") if (d / "mizan.xlsx").exists() else []
    # Kural dosyası hesapları yalnızca banka/çek için kullanılır.
    kod_ad = {k: a for k, a in hes}
    for t in ("banka", "cek"):
        kp = kural_yolu(d, t)
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
    body.satirlar = isleyici.kontrol_isaretle(body.satirlar)
    if format == "xml":
        data = isleyici.fis_xml(body.satirlar)
        fname = f"FIS_{body.tip}_{datetime.now():%Y%m%d_%H%M%S}.xml"
        (d / "cikti" / fname).write_bytes(data)
        return StreamingResponse(io.BytesIO(data), media_type="application/xml",
            headers={"Content-Disposition": f'attachment; filename="{fname}"'})
    bio = isleyici.fis_xlsx(body.satirlar)
    fname = f"FIS_{body.tip}_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    (d / "cikti" / fname).write_bytes(bio.getvalue()); bio.seek(0)

    # Banka/çek tarafında mevcut kural dosyasına kümülatif aktarım devam eder.
    # Fatura tarafında kural listesi yoktur; öğrenme fatura_ogrenme.json üzerinden yürür.
    if body.tip != "fatura":
        try:
            from app.kurallar import kural_dosyasina_fis_ekle
            yeni = [r for r in body.satirlar if r.get("kaynak") != "gecmis_kayit"]
            if yeni:
                kural_dos = kural_yolu(d, body.tip)
                if kural_dos.exists():
                    kural_dosyasina_fis_ekle(kural_dos, yeni)
        except Exception as e:
            print(f"[uyarı] kural.xlsx'e ekleme başarısız: {e}")

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


@app.delete("/api/firma/{kod}/ciktilar")
def ciktilar_sil(kod: str):
    """Firmanın kayıtlı çıktılarının (xlsx/xml) tümünü siler."""
    d = firma_dir(kod)
    cd = d / "cikti"
    silinen = 0
    if cd.exists():
        for f in cd.iterdir():
            if f.is_file() and f.suffix in (".xlsx", ".xml"):
                f.unlink()
                silinen += 1
    return {"ok": True, "silinen": silinen}


@app.get("/api/firma/{kod}/cikti/{fname}")
def cikti_indir(kod: str, fname: str):
    d = firma_dir(kod)
    if ".." in fname or "/" in fname or "\\" in fname:
        raise HTTPException(400, "Geçersiz")
    p = d / "cikti" / fname
    if not p.exists():
        raise HTTPException(404, "Yok")
    return FileResponse(p, filename=fname)


@app.delete("/api/firma/{kod}/cikti/{fname}")
def cikti_sil(kod: str, fname: str):
    d = firma_dir(kod)
    if ".." in fname or "/" in fname or "\\" in fname:
        raise HTTPException(400, "Geçersiz")
    p = d / "cikti" / fname
    if not p.exists():
        raise HTTPException(404, "Yok")
    p.unlink()
    return {"ok": True, "silinen": fname}


# ----------------------------------------------------------------- statik
app.mount("/", StaticFiles(directory=str(Path(__file__).parent / "static"), html=True), name="static")
