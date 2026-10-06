"""
Yerel yapay zekâ (Ollama) — fatura PDF'inde düzenli ifadelerin okuyamadığı
alanları tamamlamak için YEDEK okuyucu.

İlke: rakamlara asla körü körüne güvenilmez. Modelin verdiği her sonuç
dogrula() ile denetlenir:
  * matrah + KDV (+ ek vergi) = ödenecek toplam
  * her oranda KDV ≈ matrah × oran
  * toplam ve fatura no fatura METNİNDE gerçekten geçiyor olmalı
  * unvan kelimeleri metinde geçiyor olmalı
Denetimden geçmeyen rakamlar kullanılmaz; yalnızca metinle doğrulanan unvan /
tarih / fatura no alınabilir.

Ayarlar (ortam değişkeni):
  OLLAMA_URL           örn. http://host.docker.internal:11434  (boşsa kapalı)
  OLLAMA_MODEL         tercih sırası, virgülle: varsayılan "qwen3:14b,qwen2.5:14b"
                       (sunucuda ilk bulunan kullanılır; büyük model henüz
                       indirilmediyse araç yedekle çalışmaya devam eder)
  OLLAMA_ZAMAN_ASIMI   saniye, varsayılan 600 (CPU'da 14b bir fatura dakikalar sürebilir)
"""
import json, os, re, time, urllib.request, urllib.error
from datetime import datetime

OLLAMA_URL = os.environ.get("OLLAMA_URL", "").rstrip("/")
OLLAMA_MODEL = os.environ.get("OLLAMA_MODEL", "qwen3:14b,qwen2.5:14b")
MODELLER = [m.strip() for m in OLLAMA_MODEL.split(",") if m.strip()]
_aktif = {"model": MODELLER[0] if MODELLER else ""}


def aktif_model() -> str:
    """Sunucuda bulunan, tercih sırasındaki ilk model (durum() belirler)."""
    return _aktif["model"]


def _model_var(model: str, yuklu: list) -> bool:
    # "qwen3:14b" ile "qwen3:14b-q4_K_M" gibi etiketleri de kabul et
    return any(m == model or m.startswith(model + "-")
               or (":" not in model and m.split(":")[0] == model) for m in yuklu)


def _govde_hazirla(govde: dict) -> dict:
    """Qwen3 varsayılan olarak cevaptan önce uzun bir 'düşünme' metni üretir:
    yavaşlatır ve JSON çıktısını bozabilir. Kısa, yapılandırılmış işlerde kapatılır
    (Ollama 'think' parametresi + Qwen3'ün /no_think anahtarı; eski Ollama
    sürümleri bilmediği alanı yok sayar)."""
    model = govde["model"]
    if model.startswith("qwen3"):
        govde["think"] = False
        govde["messages"][-1]["content"] += "\n/no_think"
    return govde
ZAMAN_ASIMI = int(os.environ.get("OLLAMA_ZAMAN_ASIMI", "600"))

_durum_cache = {"zaman": 0.0, "veri": None}


def _istek(yol: str, govde: dict | None = None, zaman_asimi: float = 5):
    veri = json.dumps(govde).encode("utf-8") if govde is not None else None
    req = urllib.request.Request(
        OLLAMA_URL + yol, data=veri,
        headers={"Content-Type": "application/json"} if veri else {},
        method="POST" if veri is not None else "GET")
    with urllib.request.urlopen(req, timeout=zaman_asimi) as r:
        return json.loads(r.read().decode("utf-8"))


def durum(tazele: bool = False) -> dict:
    """Ollama erişilebilir mi, model yüklü mü. 30 sn önbellekli."""
    if not tazele and _durum_cache["veri"] and time.time() - _durum_cache["zaman"] < 30:
        return _durum_cache["veri"]
    out = {"etkin": False, "url": OLLAMA_URL, "model": aktif_model(), "tercih": MODELLER,
           "model_var": False, "modeller": [], "hata": "", "not": ""}
    if not OLLAMA_URL:
        out["hata"] = "OLLAMA_URL ayarlı değil"
    else:
        try:
            tags = _istek("/api/tags", zaman_asimi=3)
            out["modeller"] = [m.get("name", "") for m in tags.get("models", [])]
            secilen = next((m for m in MODELLER if _model_var(m, out["modeller"])), "")
            if secilen:
                _aktif["model"] = secilen
                out.update(model=secilen, model_var=True, etkin=True)
                if secilen != MODELLER[0]:
                    out["not"] = (f"{MODELLER[0]} sunucuda yok, şimdilik {secilen} kullanılıyor "
                                  f"(sunucuda: ollama pull {MODELLER[0]})")
            else:
                out["hata"] = f"Ollama'da şu modellerin hiçbiri yok: {', '.join(MODELLER)}"
        except Exception as e:
            out["hata"] = f"Ollama'ya ulaşılamadı: {e.__class__.__name__}"
    _durum_cache.update(zaman=time.time(), veri=out)
    return out


# Modelin dönmek ZORUNDA olduğu şekil (Ollama "format" ile zorlanır)
_SEMA = {
    "type": "object",
    "properties": {
        "fatura_no": {"type": "string"},
        "tarih": {"type": "string", "description": "GG.AA.YYYY"},
        "satici_unvan": {"type": "string"},
        "alici_unvan": {"type": "string"},
        "kalemler": {
            "type": "array",
            "items": {"type": "object",
                      "properties": {"oran": {"type": "number"},
                                     "matrah": {"type": "number"},
                                     "kdv": {"type": "number"}},
                      "required": ["oran", "matrah", "kdv"]},
        },
        "ek_vergiler": {"type": "number"},
        "odenecek_toplam": {"type": "number"},
        "iade": {"type": "boolean"},
    },
    "required": ["fatura_no", "tarih", "satici_unvan", "alici_unvan",
                 "kalemler", "ek_vergiler", "odenecek_toplam", "iade"],
}

_TALIMAT = """Sen bir Türk muhasebe asistanısın. Aşağıda bir e-Fatura / e-Arşiv fatura PDF'inden çıkarılmış ham metin var.
Bu metinden şu alanları çıkar ve YALNIZCA istenen JSON'u döndür:
- fatura_no: Fatura No / Fatura Numarası (örn. GIB2026000000448, ABC2026000012345)
- tarih: fatura (düzenlenme) tarihi, GG.AA.YYYY biçiminde
- satici_unvan: faturayı DÜZENLEYEN firmanın tam unvanı (genelde sayfanın en üstünde)
- alici_unvan: "SAYIN" ile başlayan bölümdeki alıcının unvanı
- kalemler: her KDV oranı için bir kayıt: oran (1, 10, 20 gibi), o orandaki matrah ve hesaplanan KDV tutarı
- ek_vergiler: KDV dışındaki vergiler toplamı (ÖTV, BSMV, konaklama vergisi vb.), yoksa 0
- odenecek_toplam: Ödenecek Tutar / Vergiler Dahil Toplam Tutar
- iade: fatura tipi İADE ise true
Kurallar: Tutarları Türkçe biçimden sayıya çevir (1.234,56 -> 1234.56). Metinde olmayan hiçbir bilgiyi uydurma;
bulamadığın metin alanını boş bırak, bulamadığın tutarı 0 yaz."""


def fatura_cikar(metin: str) -> dict | None:
    """Ollama'ya fatura metnini verip yapılandırılmış alanları ister. Ham model çıktısını döner."""
    if not OLLAMA_URL:
        return None
    metin = (metin or "").strip()
    if len(metin) > 12000:          # bağlam penceresini aşmasın; özet tablolar genelde sonda
        metin = metin[:6000] + "\n...\n" + metin[-6000:]
    govde = {
        "model": aktif_model(),
        "stream": False,
        "format": _SEMA,
        "options": {"temperature": 0, "num_ctx": 8192},
        "messages": [
            {"role": "system", "content": _TALIMAT},
            {"role": "user", "content": "FATURA METNİ:\n" + metin},
        ],
    }
    yanit = _istek("/api/chat", _govde_hazirla(govde), zaman_asimi=ZAMAN_ASIMI)
    icerik = (yanit.get("message") or {}).get("content", "")
    return json.loads(icerik)


# ------------------------------------------------------------------ doğrulama
def _metin_rakamlari(metin: str) -> set:
    """Metinde geçen tüm parasal değerler (Türkçe/İngilizce biçimler) kuruş cinsinden."""
    out = set()
    for m in re.finditer(r"\d{1,3}(?:[.\s]\d{3})*,\d{2}|\d+,\d{2}|\d{1,3}(?:,\d{3})*\.\d{2}|\d+\.\d{2}", metin):
        s = m.group(0).replace(" ", "")
        if "," in s and ("." not in s or s.rfind(",") > s.rfind(".")):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
        try:
            out.add(round(float(s) * 100))
        except ValueError:
            pass
    return out


def _sade(s: str) -> str:
    tr = str.maketrans("ÇĞİIÖŞÜçğıiöşü", "CGIIOSUCGIIOSU")
    return re.sub(r"[^A-Z0-9 ]", " ", (s or "").translate(tr).upper())


def _metinde_var(parca: str, metin: str) -> bool:
    """Unvan gibi bir parçanın kelimelerinin çoğu metinde geçiyor mu."""
    kelimeler = [k for k in _sade(parca).split() if len(k) > 2]
    if not kelimeler:
        return False
    sm = _sade(metin)
    bulunan = sum(1 for k in kelimeler if k in sm)
    return bulunan / len(kelimeler) >= 0.7


def dogrula(sonuc: dict, metin: str) -> dict:
    """Model çıktısını metne ve aritmetiğe karşı denetler.
    Döner: {'alanlar': {güvenilen alanlar}, 'rakam_gecerli': bool, 'notlar': [...]}"""
    notlar, alanlar = [], {}
    if not isinstance(sonuc, dict):
        return {"alanlar": {}, "rakam_gecerli": False, "notlar": ["model geçersiz yanıt verdi"]}
    sm = _sade(metin).replace(" ", "")

    fno = (sonuc.get("fatura_no") or "").strip()
    if fno and _sade(fno).replace(" ", "") in sm:
        alanlar["fatura_no"] = fno
    elif fno:
        notlar.append(f"fatura no '{fno}' metinde yok — kullanılmadı")

    tarih = (sonuc.get("tarih") or "").strip()
    m = re.match(r"^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$", tarih)
    if m:
        try:
            d = datetime(int(m.group(3)), int(m.group(2)), int(m.group(1)))
            gg_aa = f"{d.day:02d}.{d.month:02d}.{d.year}"
            ham = re.sub(r"\D", "", metin)
            if f"{d.day:02d}{d.month:02d}{d.year}" in ham or gg_aa in metin \
                    or f"{d.day:02d}-{d.month:02d}-{d.year}" in metin or f"{d.day:02d}/{d.month:02d}/{d.year}" in metin:
                alanlar["tarih"] = f"{d.year:04d}-{d.month:02d}-{d.day:02d}"
            else:
                notlar.append(f"tarih {tarih} metinde yok — kullanılmadı")
        except ValueError:
            notlar.append(f"geçersiz tarih {tarih}")

    for alan in ("satici_unvan", "alici_unvan"):
        v = (sonuc.get(alan) or "").strip()
        if v and _metinde_var(v, metin):
            alanlar[alan] = re.sub(r"\s+", " ", v)
        elif v:
            notlar.append(f"{alan} metinle uyuşmuyor — kullanılmadı")

    alanlar["iade"] = bool(sonuc.get("iade"))

    # --- rakamlar
    rakamlar = _metin_rakamlari(metin)
    kalemler = []
    for k in sonuc.get("kalemler") or []:
        try:
            oran = int(round(float(k.get("oran", 0))))
            matrah = round(float(k.get("matrah", 0)), 2)
            kdv = round(float(k.get("kdv", 0)), 2)
        except (TypeError, ValueError):
            continue
        if matrah == 0 and kdv == 0:
            continue
        kalemler.append({"oran": oran, "matrah": matrah, "kdv": kdv})
    ek = round(float(sonuc.get("ek_vergiler") or 0), 2)
    toplam = round(float(sonuc.get("odenecek_toplam") or 0), 2)

    rakam_gecerli = bool(kalemler) and toplam > 0
    if rakam_gecerli and abs(sum(k["matrah"] + k["kdv"] for k in kalemler) + ek - toplam) > 0.05:
        rakam_gecerli = False
        notlar.append("matrah + KDV + ek vergi ödenecek toplamı tutmuyor")
    if rakam_gecerli:
        for k in kalemler:
            if k["oran"] not in (0, 1, 8, 10, 18, 20):
                rakam_gecerli = False
                notlar.append(f"geçersiz KDV oranı %{k['oran']}")
            elif abs(k["matrah"] * k["oran"] / 100 - k["kdv"]) > max(0.05, k["matrah"] * 0.0005):
                rakam_gecerli = False
                notlar.append(f"%{k['oran']} için KDV matrahla uyuşmuyor")
    if rakam_gecerli and round(toplam * 100) not in rakamlar:
        rakam_gecerli = False
        notlar.append(f"ödenecek toplam {toplam} fatura metninde geçmiyor")
    if rakam_gecerli:
        alanlar.update(kalemler=kalemler, ek_vergi=ek, toplam=toplam)
    return {"alanlar": alanlar, "rakam_gecerli": rakam_gecerli, "notlar": notlar}


# ------------------------------------------------------------------ gider / gelir hesabı seçimi
_GIDER_TALIMAT = """Sen deneyimli bir Türk muhasebecisin (Tekdüzen Hesap Planı).
Sana bir faturanın karşı tarafı ve fatura kalemlerinin açıklamaları verilecek.
Bu faturanın {ne} için, firmanın mizanındaki hesaplardan EN UYGUN TEK hesabı seç.
Yalnızca verilen listedeki kodlardan birini seçebilirsin. Hesap adlarına ve kalem açıklamalarına bak:
ör. nakliye/taşıma -> nakliye gideri, akaryakıt -> akaryakıt/taşıt gideri, kırtasiye -> kırtasiye gideri,
kira -> kira gideri, ticari mal alışı -> 153 Ticari Mallar, demirbaş/makine -> 255/253.
Kalem açıklaması yoksa yalnızca karşı tarafın unvanından (sektöründen) çıkarım yap.
gerekce alanına seçimin nedenini tek kısa Türkçe cümleyle yaz."""


def gider_sec(cari_ad: str, aciklamalar: list, yon: str, adaylar: list) -> dict | None:
    """Fatura için gider (alış) / gelir (satış) hesabı önerir.
    adaylar: [(kod, ad)] — firmanın mizanındaki ALT hesaplar. Model yalnızca bunlardan
    seçebilir (JSON şemasında enum). Döner: {'kod', 'gerekce'} veya None."""
    if not OLLAMA_URL or not adaylar:
        return None
    kodlar = [k for k, _ in adaylar]
    sema = {"type": "object",
            "properties": {"kod": {"type": "string", "enum": kodlar},
                           "gerekce": {"type": "string"}},
            "required": ["kod", "gerekce"]}
    ne = "gider/maliyet/stok kaydı (alış faturası)" if yon == "alis" else "gelir kaydı (satış faturası)"
    kalemler = "\n".join(f"- {a}" for a in aciklamalar[:15]) or "(kalem açıklaması yok)"
    hesaplar = "\n".join(f"{k} — {a}" for k, a in adaylar)
    govde = {
        "model": aktif_model(), "stream": False, "format": sema,
        "options": {"temperature": 0, "num_ctx": 8192},
        "messages": [
            {"role": "system", "content": _GIDER_TALIMAT.format(ne=ne)},
            {"role": "user", "content": f"KARŞI TARAF: {cari_ad}\n\nFATURA KALEMLERİ:\n{kalemler}\n\n"
                                        f"SEÇEBİLECEĞİN HESAPLAR (kod — ad):\n{hesaplar}"},
        ],
    }
    yanit = _istek("/api/chat", _govde_hazirla(govde), zaman_asimi=ZAMAN_ASIMI)
    sonuc = json.loads((yanit.get("message") or {}).get("content", "") or "{}")
    kod = str(sonuc.get("kod", "")).strip()
    if kod not in kodlar:            # şema zorlasa da ikinci kez denetle
        return None
    return {"kod": kod, "gerekce": str(sonuc.get("gerekce", "")).strip()[:200]}
