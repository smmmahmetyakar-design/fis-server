"""
Faturalarda gider (alış) / gelir (satış) hesabını yerel yapay zekâyla seçme.

isle_fatura bu modüle yalnızca, kullanıcının bir düzeltmesi ve geçmiş fişlerde
bir karşılığı OLMAYAN cariler için sorar. Yapay zekâ cevabı arka planda,
PDF okumayla aynı tek işçili kuyrukta hesaplanır ve önbelleğe yazılır:
    <firma>/fatura/_gider_yz.json
Anahtar = yön + cari + kalem açıklamaları + aday hesap listesi. Mizan değişirse
(aday listesi değişir) öneri kendiliğinden yeniden istenir.
"""
import hashlib, json, re, threading
from datetime import datetime, timedelta
from pathlib import Path

from app import yapay_zeka, fatura_pdf
from app.kurallar import norm, kelimeler

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
    h.update("|".join([f"t{yapay_zeka.GIDER_TALIMAT_SURUM}", yon, norm(cari), "\n".join(sorted(aciklamalar)),
                       ",".join(k for k, _ in aday)]).encode("utf-8"))
    return h.hexdigest()[:20]


def kullanim_ozeti(gecmis_satirlar: list, yon: str) -> dict:
    """Geçmiş kayıtlardan (fiş listesi / muavin) hesap -> (kullanım sayısı, örnek açıklamalar).
    Yapay zekâ firmanın kendi alışkanlığını görsün diye aday listesine eklenir."""
    import re
    out = {}
    for r in gecmis_satirlar:
        tutar = float((r.get("borc") if yon == "alis" else r.get("alacak")) or 0)
        if tutar <= 0:
            continue
        kod = str(r.get("hesap", "")).strip()
        sayi, ornek = out.get(kod, (0, []))
        # "03/06/2026-YKA2026002941352-YURTİÇİ KARGO..." -> "YURTİÇİ KARGO..."
        acik = re.sub(r"^\s*\d{1,2}[./]\d{1,2}[./]\d{4}\s*-\s*[^-]*-\s*", "", str(r.get("detay", ""))).strip()[:40]
        if acik and acik not in ornek and len(ornek) < 3:
            ornek = ornek + [acik]
        out[kod] = (sayi + 1, ornek)
    return out


_CARI_ONEK = {"alis": ("320", "329", "331", "335", "336"), "satis": ("120", "121")}
_TARAF_DISI = ("100", "101", "102", "103", "108", "120", "121", "191", "192", "193", "194", "195",
               "300", "320", "329", "331", "335", "336", "360", "361", "370", "380", "391")
_ACIKLAMA_ONEK = re.compile(r"^\s*\d{1,2}[./]\d{1,2}[./]\d{4}\s*-\s*(?:[A-Z0-9]{3}20\d{11}\s*-\s*)?", re.I)


def _cari_adi(detay: str) -> str:
    """'13/08/2026-GIB2026000000009-LİORA MATBAA ...' -> 'LİORA MATBAA ...'"""
    return _ACIKLAMA_ONEK.sub("", str(detay or "")).strip()


def gecmis_ornekler(gecmis_satirlar: list, yon: str) -> list:
    """Geçmiş kayıtlardan (muavin / fiş listesi) fatura örnekleri çıkarır:
    [{'ad': karşı taraf, 'cari': cari hesap, 'gider': ana gider/gelir hesabı, 'sayi': kaç fatura}]
    Aynı cari + aynı gider hesabı tek örnek olarak birleşir. Ödeme/tahsilat fişleri
    (gider satırı olmayanlar) atlanır."""
    from collections import OrderedDict
    gruplar = OrderedDict()
    for r in gecmis_satirlar:
        detay = str(r.get("detay", ""))
        fno = re.search(r"[A-Z0-9]{3}20\d{2}\d{9}", detay.upper())
        anahtar = (str(r.get("fisno", "")), fno.group(0) if fno else norm(_cari_adi(detay))[:30])
        gruplar.setdefault(anahtar, []).append(r)
    cari_on = _CARI_ONEK.get(yon, ())
    birlesik = OrderedDict()
    for satirlar in gruplar.values():
        cari = next((r for r in satirlar if str(r.get("hesap", "")).startswith(cari_on)
                     and float((r.get("alacak") if yon == "alis" else r.get("borc")) or 0) > 0), None)
        if not cari or str(cari.get("hesap", "")).startswith("335"):   # personel/maaş fişi fatura değil
            continue
        taraf = "borc" if yon == "alis" else "alacak"
        giderler = [r for r in satirlar if float(r.get(taraf) or 0) > 0
                    and not str(r.get("hesap", "")).startswith(_TARAF_DISI)]
        if yon == "satis":
            giderler = [r for r in giderler if str(r.get("hesap", "")).startswith("6")]
        if not giderler:
            continue
        gider = max(giderler, key=lambda r: float(r.get(taraf) or 0))
        ad = _cari_adi(cari.get("detay", "")) or _cari_adi(gider.get("detay", ""))
        if not ad:
            continue
        k = (norm(ad), str(gider["hesap"]).strip())
        if k in birlesik:
            birlesik[k]["sayi"] += 1
        else:
            birlesik[k] = {"ad": ad[:60], "cari": str(cari["hesap"]).strip(),
                           "gider": str(gider["hesap"]).strip(), "sayi": 1}
    return list(birlesik.values())


def benzer_ornekler(ornekler: list, cari_ad: str, aciklamalar: list, aday_kodlar: set,
                    haric: str = "", en_fazla: int = 8) -> list:
    """Sorulan faturaya en çok benzeyen geçmiş örnekler (unvan + kalem kelimeleri ortaklığı).
    Hiç benzeyen yoksa firmanın en sık örnekleri döner (modelin firmanın tarzını görmesi için).
    haric: bu cari hesabının örnekleri dışarıda bırakılır (sınavda cevabı göstermemek için)."""
    hedef = kelimeler(cari_ad) | {w for a in aciklamalar[:15] for w in kelimeler(a)}
    q = norm(cari_ad)
    puanli = []
    for o in ornekler:
        if o["gider"] not in aday_kodlar or (haric and o["cari"] == haric) or norm(o["ad"]) == q:
            continue
        ortak = len(hedef & kelimeler(o["ad"]))
        puanli.append((ortak, o["sayi"], o))
    puanli.sort(key=lambda x: (x[0], x[1]), reverse=True)
    secilen, gorulen = [], set()
    for ortak, _, o in puanli:
        if len(secilen) >= en_fazla:
            break
        if o["ad"] in gorulen:
            continue
        gorulen.add(o["ad"]); secilen.append(o)
    return secilen


def onerici(d: Path, hesaplar: list, alt_kodlar: set, yon: str, kullanim: dict | None = None,
            ornekler: list | None = None):
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
        if kayit and kayit.get("red") and kayit.get("model") == yapay_zeka.aktif_model():
            return None          # bu model geçersiz cevap verdi: varsayılan kullanılır, tekrar sorulmaz
                                 # (model değişirse yeni modele bir kez daha sorulur)
        if kayit and kayit.get("hata"):
            # bağlantı/zaman aşımı gibi geçici hata: bir saat dolmadan tekrar deneme
            try:
                if datetime.now() - datetime.fromisoformat(kayit["zaman"]) < timedelta(hours=1):
                    return None
            except Exception:
                return None
        if not etkin:
            return None
        benzer = benzer_ornekler(ornekler or [], cari, list(aciklamalar), {k for k, _ in aday})
        fatura_pdf.is_ekle(("gider", d.as_posix(), a),
                           lambda: _calistir(d, a, cari, list(aciklamalar), yon_, aday, kullanim, benzer),
                           lambda e: _kaydet(d, a, {"kod": "", "hata": f"{e.__class__.__name__}: {e}",
                                                    "cari": cari, "yon": yon_}))
        return {"bekliyor": True}
    return _oner


def _kaydet(d: Path, anahtar: str, kayit: dict):
    with _kilit:
        v = _oku(d)
        v[anahtar] = {**kayit, "model": yapay_zeka.aktif_model(),
                      "zaman": datetime.now().isoformat(timespec="seconds")}
        _yaz(d, v)


def _calistir(d: Path, anahtar: str, cari: str, aciklamalar: list, yon: str, aday: list, kullanim=None,
              benzer=None):
    sonuc = yapay_zeka.gider_sec(cari, aciklamalar, yon, aday, kullanim, benzer)
    if not sonuc:
        # Model listede olmayan kod döndürdü — bunu da kaydet ki her İşle'de yeniden sorulmasın
        _kaydet(d, anahtar, {"kod": "", "red": True, "cari": cari, "yon": yon,
                             "gerekce": "model mizanda olmayan bir hesap döndürdü"})
        return
    _kaydet(d, anahtar, {**sonuc, "cari": cari, "aciklamalar": aciklamalar[:5], "yon": yon})


# ------------------------------------------------------------------ sınav
def _sinav_yolu(d: Path, yon: str) -> Path:
    p = d / "fatura"
    p.mkdir(parents=True, exist_ok=True)
    return p / f"_yz_sinav_{yon}.json"


def sinav_oku(d: Path, yon: str) -> dict:
    try:
        return json.loads(_sinav_yolu(d, yon).read_text(encoding="utf-8"))
    except Exception:
        return {"durum": "yok"}


def _sinav_yaz(d: Path, yon: str, veri: dict):
    p = _sinav_yolu(d, yon)
    tmp = p.with_suffix(".tmp")
    tmp.write_text(json.dumps(veri, ensure_ascii=False, indent=1), encoding="utf-8")
    tmp.replace(p)


def sinav_sorulari(gecmis_satirlar: list, yon: str, aday_kodlar: set, en_fazla: int = 40) -> list:
    """Geçmişteki her cari (hesap) bir soru: doğru cevap = o carinin en çok kullanılan
    gider/gelir hesabı; carinin kullandığı diğer hesaplar da 'kabul' sayılır."""
    from collections import OrderedDict
    cariler = OrderedDict()
    for o in gecmis_ornekler(gecmis_satirlar, yon):
        c = cariler.setdefault(o["cari"], {"cari": o["cari"], "adlar": {}, "giderler": {}})
        c["adlar"][o["ad"]] = c["adlar"].get(o["ad"], 0) + o["sayi"]
        c["giderler"][o["gider"]] = c["giderler"].get(o["gider"], 0) + o["sayi"]
    sorular = []
    for c in cariler.values():
        dogru = max(c["giderler"], key=c["giderler"].get)
        if dogru not in aday_kodlar:
            continue
        sorular.append({"cari": c["cari"], "ad": max(c["adlar"], key=c["adlar"].get), "dogru": dogru,
                        "kabul": sorted(c["giderler"]), "sayi": sum(c["giderler"].values())})
    sorular.sort(key=lambda x: -x["sayi"])
    return sorular[:en_fazla]


def sinav_baslat(d: Path, hesaplar: list, alt_kodlar: set, yon: str, gecmis_satirlar: list,
                 en_fazla: int = 40) -> dict:
    """Muavindeki her cariyi sırayla 'hiç görülmemiş' sayıp yapay zekâya sorar, cevabı
    muavindeki gerçek kayıtla karşılaştırır. Sorulan carinin bütün fişleri o soru için
    geçmişten çıkarılır (kullanım özeti ve örneklerde cevabı görmesin)."""
    aday = adaylar(hesaplar, alt_kodlar, yon)
    sorular = sinav_sorulari(gecmis_satirlar, yon, {k for k, _ in aday}, en_fazla)
    if not sorular:
        return {"durum": "yok", "hata": "Geçmiş kayıtlarda sorulabilecek fatura bulunamadı "
                                         "(muavin / fiş listesi yüklü mü?)"}
    anahtar = ("sinav", d.as_posix(), yon)
    if fatura_pdf.kuyrukta_mi(anahtar):
        return sinav_oku(d, yon)
    ad_of = dict(hesaplar)
    veri = {"durum": "sirada", "yon": yon, "model": yapay_zeka.aktif_model(),
            "baslangic": datetime.now().isoformat(timespec="seconds"),
            "toplam": len(sorular), "yapilan": 0, "dogru": 0, "kabul": 0, "sonuclar": []}
    _sinav_yaz(d, yon, veri)

    def calistir():
        veri["durum"] = "calisiyor"
        _sinav_yaz(d, yon, veri)
        for q in sorular:
            fisleri = {str(r.get("fisno", "")) for r in gecmis_satirlar if str(r.get("hesap", "")).strip() == q["cari"]}
            kalan = [r for r in gecmis_satirlar if str(r.get("fisno", "")) not in fisleri]
            kullanim = kullanim_ozeti(kalan, yon)
            benzer = benzer_ornekler(gecmis_ornekler(kalan, yon), q["ad"], [], {k for k, _ in aday},
                                     haric=q["cari"])
            try:
                cevap = yapay_zeka.gider_sec(q["ad"], [], yon, aday, kullanim, benzer) or {}
                hata = "" if cevap else "geçersiz cevap"
            except Exception as e:
                cevap, hata = {}, f"{e.__class__.__name__}: {e}"[:150]
            kod = cevap.get("kod", "")
            sonuc = {"ad": q["ad"], "cari": q["cari"], "dogru": q["dogru"], "dogru_ad": ad_of.get(q["dogru"], ""),
                     "kabul": q["kabul"], "cevap": kod, "cevap_ad": ad_of.get(kod, ""),
                     "gerekce": cevap.get("gerekce", ""), "hata": hata, "sayi": q["sayi"],
                     "sonuc": "dogru" if kod == q["dogru"] else "kabul" if kod in q["kabul"] else "yanlis"}
            veri["sonuclar"].append(sonuc)
            veri["yapilan"] += 1
            veri["dogru"] += sonuc["sonuc"] == "dogru"
            veri["kabul"] += sonuc["sonuc"] in ("dogru", "kabul")
            _sinav_yaz(d, yon, veri)
        veri["durum"] = "bitti"
        veri["bitis"] = datetime.now().isoformat(timespec="seconds")
        _sinav_yaz(d, yon, veri)

    def hata_fn(e):
        veri["durum"] = "hata"
        veri["hata"] = f"{e.__class__.__name__}: {e}"[:200]
        _sinav_yaz(d, yon, veri)

    fatura_pdf.is_ekle(anahtar, calistir, hata_fn)
    return veri
