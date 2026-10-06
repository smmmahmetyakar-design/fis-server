"""
Faturalarda gider (alış) / gelir (satış) hesabını yerel yapay zekâyla seçme.

isle_fatura bu modüle yalnızca, kullanıcının bir düzeltmesi ve geçmiş fişlerde
bir karşılığı OLMAYAN cariler için sorar. Yapay zekâ cevabı arka planda,
PDF okumayla aynı tek işçili kuyrukta hesaplanır ve önbelleğe yazılır:
    <firma>/fatura/_gider_yz.json
Anahtar = yön + cari + kalem açıklamaları + aday hesap listesi. Mizan değişirse
(aday listesi değişir) öneri kendiliğinden yeniden istenir.
"""
import hashlib, json, threading
from datetime import datetime, timedelta
from pathlib import Path

from app import yapay_zeka, fatura_pdf
from app.kurallar import norm

_kilit = threading.RLock()

# Alışta gider / maliyet / stok / duran varlık; satışta gelir hesap grupları
_ONEKLER = {
    "alis": ("150", "151", "152", "153", "157", "253", "254", "255", "256", "258", "260", "264",
             "710", "720", "730", "740", "750", "760", "770", "780", "659", "689"),
    "satis": ("600", "601", "602", "649", "671", "679"),
}
_EN_FAZLA_ADAY = 150


def adaylar(hesaplar: list, alt_kodlar: set, yon: str) -> list:
    """Mizandaki kayıt atılabilir alt hesaplardan, yöne uygun gruplardakiler."""
    onek = _ONEKLER.get(yon, ())
    out = [(k, a) for k, a in hesaplar
           if k in alt_kodlar and any(k == o or k.startswith(o + ".") for o in onek)]
    return out[:_EN_FAZLA_ADAY]


def _yol(d: Path) -> Path:
    p = d / "fatura"
    p.mkdir(parents=True, exist_ok=True)
    return p / "_gider_yz.json"


def _oku(d: Path) -> dict:
    try:
        return json.loads(_yol(d).read_text(encoding="utf-8"))
    except Exception:
        return {}


def _yaz(d: Path, veri: dict):
    tmp = _yol(d).with_suffix(".tmp")
    tmp.write_text(json.dumps(veri, ensure_ascii=False, indent=1), encoding="utf-8")
    tmp.replace(_yol(d))


def _anahtar(cari: str, aciklamalar: list, yon: str, aday: list) -> str:
    h = hashlib.sha1()
    h.update("|".join([yon, norm(cari), "\n".join(sorted(aciklamalar)),
                       ",".join(k for k, _ in aday)]).encode("utf-8"))
    return h.hexdigest()[:20]


def onerici(d: Path, hesaplar: list, alt_kodlar: set, yon: str):
    """isle_fatura'ya verilecek fonksiyon: (cari, aciklamalar, yon) -> öneri.
    Önbellekte varsa {'kod','gerekce'}; yoksa işi kuyruğa atar ve {'bekliyor': True};
    yapay zekâ kapalıysa None."""
    aday = adaylar(hesaplar, alt_kodlar, yon)
    if not aday:
        return None
    etkin = yapay_zeka.durum()["etkin"]
    onbellek = _oku(d)

    def _oner(cari, aciklamalar, yon_):
        a = _anahtar(cari, aciklamalar, yon_, aday)
        kayit = onbellek.get(a)
        if kayit and kayit.get("kod") in alt_kodlar:
            return {"kod": kayit["kod"], "gerekce": kayit.get("gerekce", "")}
        if kayit and kayit.get("red"):
            return None          # model geçersiz cevap verdi: varsayılan kullanılır, tekrar sorulmaz
        if kayit and kayit.get("hata"):
            # bağlantı/zaman aşımı gibi geçici hata: bir saat dolmadan tekrar deneme
            try:
                if datetime.now() - datetime.fromisoformat(kayit["zaman"]) < timedelta(hours=1):
                    return None
            except Exception:
                return None
        if not etkin:
            return None
        fatura_pdf.is_ekle(("gider", d.as_posix(), a),
                           lambda: _calistir(d, a, cari, list(aciklamalar), yon_, aday),
                           lambda e: _kaydet(d, a, {"kod": "", "hata": f"{e.__class__.__name__}: {e}",
                                                    "cari": cari, "yon": yon_}))
        return {"bekliyor": True}
    return _oner


def _kaydet(d: Path, anahtar: str, kayit: dict):
    with _kilit:
        v = _oku(d)
        v[anahtar] = {**kayit, "model": yapay_zeka.OLLAMA_MODEL,
                      "zaman": datetime.now().isoformat(timespec="seconds")}
        _yaz(d, v)


def _calistir(d: Path, anahtar: str, cari: str, aciklamalar: list, yon: str, aday: list):
    sonuc = yapay_zeka.gider_sec(cari, aciklamalar, yon, aday)
    if not sonuc:
        # Model listede olmayan kod döndürdü — bunu da kaydet ki her İşle'de yeniden sorulmasın
        _kaydet(d, anahtar, {"kod": "", "red": True, "cari": cari, "yon": yon,
                             "gerekce": "model mizanda olmayan bir hesap döndürdü"})
        return
    _kaydet(d, anahtar, {**sonuc, "cari": cari, "aciklamalar": aciklamalar[:5], "yon": yon})
