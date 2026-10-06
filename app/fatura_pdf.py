"""
Fatura PDF'leri — alış/satış sekmesinde fatura listesinin (Excel) yanına
tek tek fatura PDF'leri eklenir. Bunlar:
  * listede eksik kalan alanları (cari, tarih, KDV dağılımı) tamamlar,
  * listede hiç olmayan faturaları ekler.

Okuma iki katmanlı:
  1) Metin + düzenli ifade (isleyici._pdf_tekil_fatura_ayikla) — anında, kesin.
  2) Eksik kalan alanlar için yerel yapay zekâ (Ollama), ARKA PLANDA tek
     işçili kuyrukta. Sonuç yapay_zeka.dogrula()'dan geçmeden kullanılmaz.

Durum: <firma>/fatura/pdf_<yon>/_okuma.json  (dosya adı -> okuma sonucu)
Aynı dosya (sha1) ve aynı okuyucu sürümü için tekrar okunmaz.
"""
import hashlib, json, queue, re, threading, time
from datetime import datetime
from pathlib import Path

from app import yapay_zeka
from app.kurallar import norm

OKUYUCU_SURUM = 2   # 2: fatura kalem açıklamaları (gider seçimi için)
YONLER = ("alis", "satis")

_kilit = threading.RLock()
_kuyruk: "queue.Queue" = queue.Queue()
_kuyrukta: set = set()
_isci = None

# Yapay zekâya ne zaman gidilir: bu alanlardan biri eksikse
_YZ_ALANLARI = ("cari", "tarih", "fatura_no", "tutar", "kdv")


# ------------------------------------------------------------------ yollar / durum
def klasor(d: Path, yon: str) -> Path:
    if yon not in YONLER:
        raise ValueError("yon alis|satis olmalı")
    k = d / "fatura" / f"pdf_{yon}"
    k.mkdir(parents=True, exist_ok=True)
    return k


def _durum_yolu(k: Path) -> Path:
    return k / "_okuma.json"


def _oku(k: Path) -> dict:
    p = _durum_yolu(k)
    if p.exists():
        try:
            return json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def _yaz(k: Path, veri: dict):
    tmp = _durum_yolu(k).with_suffix(".tmp")
    tmp.write_text(json.dumps(veri, ensure_ascii=False, indent=1), encoding="utf-8")
    tmp.replace(_durum_yolu(k))


def _guncelle(k: Path, dosya: str, **alanlar):
    with _kilit:
        v = _oku(k)
        if dosya in v:
            v[dosya].update(alanlar)
            _yaz(k, v)


def pdf_dosyalari(k: Path) -> list:
    return sorted(p for p in k.iterdir() if p.is_file() and p.suffix.lower() == ".pdf")


def _sha1(p: Path) -> str:
    h = hashlib.sha1()
    with open(p, "rb") as f:
        for parca in iter(lambda: f.read(1 << 16), b""):
            h.update(parca)
    return h.hexdigest()


# ------------------------------------------------------------------ metin çıkarma
def _sayfalar(p: Path) -> list:
    """PDF'in sayfa sayfa metni ve tabloları. Metin yoksa (taranmış) OCR'a düşer;
    OCR'da sayfa ayrımı tutulmaz, tüm dosya tek sayfa sayılır."""
    import pdfplumber
    out = []
    with pdfplumber.open(p) as pdf:
        for s in pdf.pages:
            t = s.extract_text() or ""
            tablolar = [[[("" if c is None else str(c)) for c in row] for row in tb]
                        for tb in (s.extract_tables() or [])]
            out.append({"ham_metin": t, "tablolar": tablolar})
    if not any(s["ham_metin"].strip() for s in out):
        from app.belge_oku import pdf_ocr
        o = pdf_ocr(p)
        return [{"ham_metin": o.get("ham_metin", ""), "tablolar": [], "ocr": True}]
    return out


# ------------------------------------------------------------------ düzenli ifade okuma
def _kalemleri_duzelt(ham_kalemler, ek_vergi):
    """Eski ayrıştırıcı çok oranlı faturada matrahın TAMAMINI ilk orana yazar
    (diğer oranlarda matrah 0). KDV tutarlarından oran bazında matrah türetilir;
    toplam matrahı tutuyorsa onlar kullanılır, kalan kısım %0 kalem olur."""
    kalemler = [{"oran": int(o), "matrah": round(m or 0, 2), "kdv": round(k or 0, 2)}
                for o, m, k in ham_kalemler if (m or k)]
    kdvli = [x for x in kalemler if x["oran"] > 0 and x["kdv"] > 0]
    if len(kdvli) < 2:
        return kalemler
    matrah_toplam = round(sum(x["matrah"] for x in kalemler), 2)
    turetilen = sorted(({"oran": x["oran"], "matrah": round(x["kdv"] * 100 / x["oran"], 2), "kdv": x["kdv"]}
                        for x in kdvli), key=lambda x: x["oran"])
    t_top = round(sum(x["matrah"] for x in turetilen), 2)
    fark = round(matrah_toplam - t_top, 2)
    # KDV kuruşa yuvarlandığı için geri hesaplanan matrah orana göre sapar:
    # %1'de en çok ±0,50 TL, %20'de ±0,03 TL. Bu sınırdaki fark yuvarlamadır ve
    # en düşük oranlı kaleme yazılır; böylece matrah toplamı kuruşu kuruşuna tutar.
    tolerans = sum(0.6 / x["oran"] for x in turetilen)
    if abs(fark) <= tolerans:
        turetilen[0]["matrah"] = round(turetilen[0]["matrah"] + fark, 2)
        return turetilen
    if fark > 0:      # kalan kısım KDV'siz (istisna) kalem
        return turetilen + [{"oran": 0, "matrah": fark, "kdv": 0.0}]
    return kalemler


_KDV_IZI_RE = re.compile(r"(KDV|KATMA\s*DE[GĞ]ER)[^\n]{0,40}%\s*(1|8|10|18|20)\b|%\s*(1|8|10|18|20)\b[^\n]{0,25}KDV",
                         re.IGNORECASE)


def _eksikler(f: dict, metin: str) -> list:
    e = []
    if not f.get("cari_ad"):
        e.append("cari")
    if not f.get("tarih"):
        e.append("tarih")
    if not f.get("fatura_no"):
        e.append("fatura_no")
    if not f.get("kalemler") and not f.get("ek_vergi"):
        e.append("tutar")
    elif not any(k["kdv"] > 0 for k in f.get("kalemler", [])) and _KDV_IZI_RE.search(metin or ""):
        e.append("kdv")        # metinde KDV oranı geçiyor ama tutarı okunamadı
    return e


_KALEM_BASLIK = ("MAL HIZMET", "MAL/HIZMET", "ACIKLAMA", "URUN", "HIZMET", "CINSI", "STOK ADI", "MALZEME")


def _kalem_aciklamalari(tablolar: list) -> list:
    """Fatura kalem tablosundan mal/hizmet açıklamaları ("Nakliye bedeli",
    "A4 fotokopi kağıdı"...). Gider hesabı seçiminde yapay zekâya verilir."""
    out = []
    for tb in tablolar or []:
        for ri, row in enumerate(tb[:3]):
            nrow = [norm(str(c or "")).replace("/", " ") for c in row]
            j = next((j for j, c in enumerate(nrow)
                      if any(b.replace("/", " ") in c for b in _KALEM_BASLIK)
                      and not any(x in c for x in ("TOPLAM", "TUTAR", "KDV", "FIYAT", "ORAN", "MIKTAR"))), None)
            if j is None:
                continue
            for r in tb[ri + 1:]:
                if j < len(r):
                    v = re.sub(r"\s+", " ", str(r[j] or "")).strip()
                    if len(v) > 2 and not re.fullmatch(r"[\d.,%\s]+(TL)?", v):
                        out.append(v[:120])
            break
    tekil = []
    for v in out:
        if v not in tekil:
            tekil.append(v)
    return tekil[:15]


def _sayfa_oku(sayfa: dict, dosya: str, yon: str) -> dict | None:
    from app.isleyici import _pdf_tekil_fatura_ayikla
    k = _pdf_tekil_fatura_ayikla({**sayfa, "dosya": dosya}, yon=yon)
    metin = sayfa.get("ham_metin", "")
    if not k:
        # Ayrıştırıcı hiçbir şey bulamadı — yine de boş bir kayıt aç; yapay zekâ doldurabilir
        if len(metin.strip()) < 80:
            return None
        k = {"fatura_no": "", "tarih": "", "gonderici": "", "kalemler": [], "ek_vergiler": 0, "iade": False}
    ek = round(k.get("ek_vergiler") or 0, 2)
    kalemler = _kalemleri_duzelt(k.get("kalemler") or [], ek)
    f = {
        "fatura_no": (k.get("fatura_no") or "").strip(),
        "tarih": k.get("tarih", ""),
        "cari_ad": (k.get("gonderici") or "").strip(),
        "tur": "IADE" if k.get("iade") else "SATIS",
        "senaryo": "TEMELFATURA" if (not kalemler and ek > 0) else "",
        "kalemler": kalemler,
        "toplam": round(sum(x["matrah"] + x["kdv"] for x in kalemler) + ek, 2),
        "tevkifat": 0.0, "tevkifat_kod": "",
        "ek_vergi": ek,
        "yon": yon, "dosya": dosya, "kaynak": "pdf", "yz": [],
        "kalem_aciklamalari": _kalem_aciklamalari(sayfa.get("tablolar")),
    }
    f["eksik"] = _eksikler(f, metin)
    return f


def _dosya_oku(p: Path, yon: str) -> tuple:
    """Döner: (faturalar, sayfa_metinleri). Çok sayfalı PDF önce sayfa sayfa
    denenir; her sayfa ayrı fatura no'lu tam bir fatura çıkıyorsa birden çok
    fatura sayılır, yoksa dosyanın tamamı tek faturadır."""
    sayfalar = _sayfalar(p)
    if len(sayfalar) > 1:
        tek_tek = [_sayfa_oku(s, p.name, yon) for s in sayfalar]
        nolar = [f["fatura_no"] for f in tek_tek if f]
        if all(tek_tek) and all(nolar) and len(set(nolar)) == len(nolar) \
                and all(f["toplam"] > 0 for f in tek_tek):
            for i, f in enumerate(tek_tek):
                f["sayfa"] = i + 1
            return tek_tek, [s["ham_metin"] for s in sayfalar]
    birlesik = {"ham_metin": "\n".join(s["ham_metin"] for s in sayfalar),
                "tablolar": [t for s in sayfalar for t in s["tablolar"]]}
    f = _sayfa_oku(birlesik, p.name, yon)
    return ([f] if f else []), [birlesik["ham_metin"]]


def _durum_hesapla(kayit: dict) -> str:
    if kayit.get("hata"):
        return "okunamadi"
    fl = kayit.get("faturalar") or []
    if not fl:
        return "okunamadi"
    if any(f.get("eksik") for f in fl):
        return kayit.get("yz_durum") or "eksik"
    return "yz_tamamladi" if any(f.get("yz") for f in fl) else "tam"


def oku(k: Path, p: Path, yon: str, yz_kuyruga: bool = True) -> dict:
    """Bir PDF'i (gerekiyorsa) okur, durumu kaydeder, eksik varsa yapay zekâ kuyruğuna atar."""
    sha = _sha1(p)
    with _kilit:
        v = _oku(k)
        eski = v.get(p.name)
    if eski and eski.get("sha1") == sha and eski.get("surum") == OKUYUCU_SURUM:
        kayit = eski
    else:
        kayit = {"sha1": sha, "surum": OKUYUCU_SURUM, "zaman": datetime.now().isoformat(timespec="seconds"),
                 "faturalar": [], "hata": "", "yz_durum": "", "yz_notlar": []}
        try:
            kayit["faturalar"], _ = _dosya_oku(p, yon)
        except Exception as e:
            kayit["hata"] = f"{e.__class__.__name__}: {e}"
        if eski and eski.get("sha1") == sha:
            _yz_alanlarini_tasi(eski, kayit)
        with _kilit:
            v = _oku(k)
            v[p.name] = kayit
            _yaz(k, v)
    if yz_kuyruga and any(f.get("eksik") for f in kayit.get("faturalar") or []) \
            and kayit.get("yz_durum") in ("", "yz_sirada", "yz_okunuyor"):
        yz_kuyruga_al(k, p.name, yon)
    return kayit


def _yz_alanlarini_tasi(eski: dict, yeni: dict):
    """Okuyucu sürümü değişip PDF yeniden okunduğunda, aynı dosya için yapay
    zekânın daha önce doldurduğu alanları yeni kayda aktarır (tekrar sorulmasın)."""
    eski_fl = eski.get("faturalar") or []
    for i, f in enumerate(yeni.get("faturalar") or []):
        e = next((x for x in eski_fl if x.get("fatura_no") and x.get("fatura_no") == f.get("fatura_no")),
                 eski_fl[i] if i < len(eski_fl) else None)
        if not e or not e.get("yz"):
            continue
        for alan in e["yz"]:
            if alan == "cari" and "cari" in f["eksik"]:
                f["cari_ad"] = e["cari_ad"]
            elif alan == "tarih" and "tarih" in f["eksik"]:
                f["tarih"] = e["tarih"]
            elif alan == "fatura_no" and "fatura_no" in f["eksik"]:
                f["fatura_no"] = e["fatura_no"]
            elif alan in ("tutar", "kdv") and ({"tutar", "kdv"} & set(f["eksik"])):
                f["kalemler"], f["ek_vergi"], f["toplam"] = e["kalemler"], e["ek_vergi"], e["toplam"]
                f["senaryo"] = e.get("senaryo", f.get("senaryo", ""))
        f["yz"] = list(e["yz"])
        f["eksik"] = [x for x in f["eksik"] if x not in e["yz"] and not (x == "tutar" and "kdv" in e["yz"])]
    yeni["yz_durum"] = eski.get("yz_durum", "")
    yeni["yz_notlar"] = eski.get("yz_notlar", [])


def sil(k: Path, dosya: str):
    p = k / dosya
    if p.exists():
        p.unlink()
    with _kilit:
        v = _oku(k)
        v.pop(dosya, None)
        _yaz(k, v)


def liste(k: Path, yon: str) -> list:
    """Arayüz için dosya başına özet. Okunmamış dosyalar burada okunur."""
    out = []
    for p in pdf_dosyalari(k):
        kayit = oku(k, p, yon)
        fl = kayit.get("faturalar") or []
        out.append({
            "dosya": p.name,
            "durum": _durum_hesapla(kayit),
            "hata": kayit.get("hata", ""),
            "yz_notlar": kayit.get("yz_notlar", []),
            "faturalar": [{"fatura_no": f["fatura_no"], "tarih": f["tarih"], "cari_ad": f["cari_ad"],
                           "toplam": f["toplam"], "eksik": f.get("eksik", []), "yz": f.get("yz", [])}
                          for f in fl],
        })
    return out


def faturalar(d: Path, yon: str) -> tuple:
    """isle_fatura için tüm PDF faturaları. Döner: (fatura listesi, yz bekleyen dosya sayısı)."""
    k = klasor(d, yon)
    out, bekleyen = [], 0
    for p in pdf_dosyalari(k):
        kayit = oku(k, p, yon)
        if kayit.get("yz_durum") in ("yz_sirada", "yz_okunuyor") and \
                any(f.get("eksik") for f in kayit.get("faturalar") or []):
            bekleyen += 1
        out.extend(json.loads(json.dumps(kayit.get("faturalar") or [])))   # derin kopya
    return out, bekleyen


# ------------------------------------------------------------------ yapay zekâ kuyruğu
def yz_kuyruga_al(k: Path, dosya: str, yon: str, zorla: bool = False) -> bool:
    """Yapay zekâ açıksa dosyayı kuyruğa atar. Kapalıysa durumu not eder."""
    global _isci
    yz = yapay_zeka.durum()
    if not yz["etkin"]:
        not_ = [f"Yapay zekâ kullanılamıyor: {yz['hata']}"]
        with _kilit:
            if _oku(k).get(dosya, {}).get("yz_notlar") != not_:
                _guncelle(k, dosya, yz_durum="", yz_notlar=not_)
        return False
    anahtar = ("pdf", k.as_posix(), dosya)
    with _kilit:
        if anahtar in _kuyrukta:
            return True
        if zorla:
            v = _oku(k)
            if dosya in v:
                v[dosya]["yz_durum"] = ""
                _yaz(k, v)
        _guncelle(k, dosya, yz_durum="yz_sirada")
    is_ekle(anahtar, lambda: _yz_isle(k, dosya, yon),
            lambda e: _guncelle(k, dosya, yz_durum="yz_eksik",
                                yz_notlar=[f"Yapay zekâ hatası: {e.__class__.__name__}: {e}"]))
    return True


def is_ekle(anahtar, fn, hata_fn=None) -> bool:
    """Yapay zekâ işini tek işçili kuyruğa ekler (Ollama'yı aynı anda tek işle meşgul et).
    Aynı anahtar zaten kuyruktaysa tekrar eklenmez."""
    global _isci
    with _kilit:
        if anahtar in _kuyrukta:
            return False
        _kuyrukta.add(anahtar)
        _kuyruk.put((anahtar, fn, hata_fn))
        if _isci is None or not _isci.is_alive():
            _isci = threading.Thread(target=_isci_dongu, daemon=True, name="yapay-zeka")
            _isci.start()
    return True


def kuyrukta_mi(anahtar) -> bool:
    return anahtar in _kuyrukta


def kuyruk_bilgisi() -> dict:
    return {"bekleyen": _kuyruk.qsize(), "toplam": len(_kuyrukta)}


def _isci_dongu():
    while True:
        anahtar, fn, hata_fn = _kuyruk.get()
        try:
            fn()
        except Exception as e:
            if hata_fn:
                try:
                    hata_fn(e)
                except Exception:
                    pass
        finally:
            with _kilit:
                _kuyrukta.discard(anahtar)
            _kuyruk.task_done()


def _yz_isle(k: Path, dosya: str, yon: str):
    p = k / dosya
    if not p.exists():
        return
    _guncelle(k, dosya, yz_durum="yz_okunuyor")
    with _kilit:
        kayit = _oku(k).get(dosya)
    if not kayit:
        return
    _, metinler = _dosya_oku(p, yon)
    fl = kayit.get("faturalar") or []
    notlar = []
    for i, f in enumerate(fl):
        if not f.get("eksik"):
            continue
        metin = metinler[f["sayfa"] - 1] if f.get("sayfa") and len(metinler) >= f["sayfa"] else "\n".join(metinler)
        t0 = time.time()
        ham = yapay_zeka.fatura_cikar(metin)
        sure = round(time.time() - t0, 1)
        d = yapay_zeka.dogrula(ham, metin)
        a = d["alanlar"]
        doldurulan = []
        if "cari" in f["eksik"]:
            unvan = a.get("satici_unvan" if yon == "alis" else "alici_unvan")
            if unvan:
                f["cari_ad"] = unvan; doldurulan.append("cari")
        if "tarih" in f["eksik"] and a.get("tarih"):
            f["tarih"] = a["tarih"]; doldurulan.append("tarih")
        if "fatura_no" in f["eksik"] and a.get("fatura_no"):
            f["fatura_no"] = a["fatura_no"]; doldurulan.append("fatura_no")
        if ("tutar" in f["eksik"] or "kdv" in f["eksik"]) and d["rakam_gecerli"]:
            # metinden okunan bir toplam varsa model onunla aynı toplamı bulmuş olmalı
            if f["toplam"] and abs(f["toplam"] - a["toplam"]) > 0.05:
                notlar.append(f"{f['fatura_no'] or dosya}: model toplamı ({a['toplam']}) metindeki toplamla ({f['toplam']}) uyuşmuyor — tutarlar kullanılmadı")
            else:
                f["kalemler"], f["ek_vergi"], f["toplam"] = a["kalemler"], a["ek_vergi"], a["toplam"]
                if not f["kalemler"] and f["ek_vergi"] > 0:
                    f["senaryo"] = "TEMELFATURA"
                doldurulan.append("tutar" if "tutar" in f["eksik"] else "kdv")
        if a.get("iade"):
            f["tur"] = "IADE"
        f["yz"] = sorted(set(f.get("yz", []) + doldurulan))
        f["eksik"] = [e for e in f["eksik"] if e not in doldurulan and not (e == "tutar" and "kdv" in doldurulan)]
        etiket = f["fatura_no"] or dosya
        notlar.append(f"{etiket}: {sure} sn" + (f", tamamlanan: {', '.join(doldurulan)}" if doldurulan else ", alan tamamlanamadı"))
        notlar.extend(f"{etiket}: {n}" for n in d["notlar"])
        fl[i] = f
    with _kilit:
        v = _oku(k)
        if dosya in v and v[dosya].get("sha1") == kayit.get("sha1"):
            v[dosya]["faturalar"] = fl
            v[dosya]["yz_durum"] = "yz_eksik" if any(x.get("eksik") for x in fl) else "yz_tamamladi"
            v[dosya]["yz_notlar"] = notlar
            _yaz(k, v)
