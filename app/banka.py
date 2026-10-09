"""
Banka ekstresi -> muhasebe fişi.

Okuma
  * Excel / CSV / HTML tabloları: başlık satırı (Tarih, Tutar ya da Borç/Alacak, Bakiye,
    Açıklama) aranır. Tarih hücresi saat de taşıyabilir ("2026-08-31-11.26.44.328029",
    "08/06/2026-14:06:26", "15.06.2026 10:46").
  * PDF metni: tarihle başlayan satır yeni harekettir; altındaki tarihsiz satırlar o
    hareketin açıklamasının devamıdır (Halkbank). Satırda açıklama yoksa (YKB) önceki
    satırlar açıklamadır.
  * Bakiye sütunu varsa tutarın işareti bakiye zinciriyle doğrulanır; ters çıkarsa düzeltilir.
  * Mükerrer: yalnız AYNI hareket iki ayrı dosyada geldiyse atılır (aynı gün aynı tutarlı
    iki EFT tek dosyada ise ikisi de kalır).

Banka hesabı
  * Dosya adında mizanda olan bir 102 kodu varsa o (ör. "Akbank TRY_102.01.001.csv").
  * Yoksa IBAN / banka adı / hesap no ipuçları (isleyici._banka_hesabi_bul).

Karşı hesap (öncelik)
  1. Kullanıcının öğrettiği (açıklamanın anlamlı kelimeleri)
  2. Kendi hesapları arası virman ve döviz alım/satım: yüklenen diğer ekstrede karşılığı
     varsa iki 102 arasında TEK kayıt (TL tarafından / gönderen hesaptan)
  3. EFT/havale masrafı ve BSMV (aynı anda yapılan transferin "- Komisyon", "- Vergi (BSMV)"
     satırları) ve banka masraf/ücret satırları -> geçmişte banka masrafına kullanılan hesap
  4. Vergi tahsilatı (vergi kodu: 0015 KDV, 4017 KDV tevkifat, 0003 stopaj, 0033/0032
     geçici vergi, 0010 kurumlar, 1047/1048 damga) ve SGK -> mizandaki ilgili hesap
  5. Geçmiş fişler: aynı yönde (giriş/çıkış) geçmiş banka satırlarının açıklama kelimeleri
     + mizandaki cari adının kelimeleri; en özgül eşleşme kazanır
  6. Bulunamazsa 198 + KONTROL
"""
import re
from collections import defaultdict, Counter
from datetime import datetime

from app.kurallar import norm, STOP

# ------------------------------------------------------------------ yardımcılar
_GENEL = STOP | {"LIMITED", "SIRKETI", "SIRKET", "TICARET", "SANAYI", "ANONIM", "LTD", "STI", "AS",
                 "PVT", "GMBH", "CO", "BV", "NV", "INC", "LLC", "SRL", "SPA", "AG", "SA", "PLC",
                 "INDUSTRY", "INDUSTRIES", "INTERNATIONAL", "IMPORT", "EXPORT", "THE", "AND",
                 "IC", "DIS", "ITH", "IHR", "ITHALAT", "IHRACAT", "PAZ", "PAZARLAMA", "INS", "INSAAT",
                 "HIZ", "HIZMET", "HIZMETLERI", "MAD", "OTOMOTIV", "TIC", "SAN"}

# mizan adlarında cari olmayan, açıklamada sık geçen kelimeler (hesap adı eşleşmesinde sayılmaz)
_HESAP_GENEL = {"MASRAFLAR", "MASRAF", "KIRALAMA", "ARAC", "OFIS", "EUR", "USD", "GBP", "TL", "TRY",
                "AVANSI", "IS", "HUZUR", "HAKKI", "ORTAK", "ORTAKLAR", "PERSONEL", "BORC", "ALACAK"}

_MASRAF_KELIME = {"KOMISYON", "KOMISYONU", "BSMV", "MASRAF", "MASRAFI", "MASRAFLARI", "UCRET", "UCRETI",
                  "UCRETLERI", "ISLETIM", "AIDAT", "AIDATI"}
_DOVIZ_KELIME = {"DOVIZ", "KAMBIYO", "ARBITRAJ", "FX"}
_VIRMAN_KELIME = {"VIRMAN", "HESAPLARARASI", "HESAPLAR"}


def _kelimeler(metin: str) -> list:
    """Açıklamanın eşleştirmeye yarayan kelimeleri: rakam içermeyen, 2+ harfli, sırayı koruyarak."""
    out = []
    for w in norm(metin).split():
        if len(w) < 2 or any(ch.isdigit() for ch in w):
            continue
        out.append(w)
    return out


def banka_anahtar(metin: str) -> str:
    """Öğrenme anahtarı: tarih, referans ve numaralar atılmış açıklama."""
    return " ".join(_kelimeler(_tarih_onekini_at(metin)))


def _tarih_onekini_at(metin: str) -> str:
    return re.sub(r"^\s*\d{1,2}[./-]\d{1,2}[./-]\d{2,4}\s*", "", str(metin or ""))


def _bir_harf_farki(a: str, b: str) -> bool:
    """Tek harf eklenmiş/eksik/farklı (ASLANTAS ~ ARSLANTAS)."""
    if abs(len(a) - len(b)) > 1 or a == b:
        return a == b
    if len(a) == len(b):
        return sum(x != y for x, y in zip(a, b)) == 1
    if len(a) > len(b):
        a, b = b, a
    for i in range(len(b)):
        if b[:i] + b[i + 1:] == a:
            return True
    return False


def _w_esler(hw: str, wset: set, wlist: list, bulanik: bool = False) -> bool:
    if hw in wset:
        return True
    if len(hw) >= 4 and any(w.startswith(hw) for w in wlist):
        return True                      # Logo açıklamayı ~40 karakterde keser: "BANKAS" ~ "BANKASI"
    if bulanik and len(hw) >= 6:
        return any(len(w) >= 5 and _bir_harf_farki(hw, w) for w in wlist)
    return False


def _hepsi_var(hws, wset, wlist, bulanik=False) -> bool:
    return bool(hws) and all(_w_esler(h, wset, wlist, bulanik) for h in hws)


def _sayi(s):
    if s is None or s == "":
        return None
    if isinstance(s, (int, float)):
        return float(s)
    s = str(s).strip().replace(" ", "").replace(" ", "")
    neg = s.startswith("-") or s.endswith("-") or (s.startswith("(") and s.endswith(")"))
    s = re.sub(r"[^0-9.,]", "", s)
    if not s:
        return None
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")      # 1.234,56
        else:
            s = s.replace(",", "")                          # 1,234.56
    elif "," in s:
        s = s.replace(",", ".")
    elif s.count(".") > 1:
        s = s.replace(".", "")
    try:
        v = float(s)
    except ValueError:
        return None
    return -v if neg else v


_TARIH_ISO = re.compile(r"(?<!\d)(20\d{2})[-./](\d{1,2})[-./](\d{1,2})(?!\d)")
_TARIH_TR = re.compile(r"(?<!\d)(\d{1,2})[-./](\d{1,2})[-./](\d{4}|\d{2})(?!\d)")
_SAAT = re.compile(r"(\d{1,2})[:.](\d{2})(?:[:.](\d{2}))?(?:[.,](\d+))?")


def tarih_saat(v):
    """Hücreden (tarih_iso, saat_str). Saat yoksa ''. Bulunamazsa ('', '')."""
    if v is None:
        return "", ""
    if isinstance(v, datetime):
        return f"{v.year:04d}-{v.month:02d}-{v.day:02d}", (v.strftime("%H:%M:%S") if (v.hour or v.minute or v.second) else "")
    s = str(v).strip()
    m = _TARIH_ISO.match(s) or _TARIH_ISO.search(s[:12])
    if m:
        y, a, g = int(m.group(1)), int(m.group(2)), int(m.group(3))
    else:
        m = _TARIH_TR.match(s) or _TARIH_TR.search(s[:12])
        if not m:
            return "", ""
        g, a, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if y < 100:
            y += 2000
    try:
        datetime(y, a, g)
    except ValueError:
        return "", ""
    saat = ""
    sm = _SAAT.search(s[m.end():])
    if sm:
        saat = f"{int(sm.group(1)):02d}:{sm.group(2)}:{sm.group(3) or '00'}" + (f".{sm.group(4)}" if sm.group(4) else "")
    return f"{y:04d}-{a:02d}-{g:02d}", saat


# ------------------------------------------------------------------ okuma
def _baslik_bul(tablo):
    for i, row in enumerate(tablo[:30]):
        nrow = [norm(str(c or "")) for c in row]
        hh = {}
        for j, c in enumerate(nrow):
            if not c:
                continue
            if c in ("TARIH SAAT", "TARIH", "ISLEM TARIHI", "ISLEM TARIHI SAAT", "TARIH VE SAAT") or \
                    (c.startswith("TARIH") and "ARALIGI" not in c) or c in ("ISLEM ZAMANI", "ZAMAN"):
                hh.setdefault("tarih", j)
            elif c in ("VALOR", "VALOR TARIHI"):
                hh.setdefault("tarih_valor", j)
            elif "BAKIYE" in c and "ONCEKI" not in c:
                hh.setdefault("bakiye", j)
            elif ("GELIR" in c and "TUTAR" in c) or c in ("GIREN", "GELEN", "YATAN", "ALACAK TUTARI", "ALACAK"):
                hh["giris"] = j
            elif ("GIDER" in c and "TUTAR" in c) or c in ("CIKAN", "GIDEN", "CEKILEN", "BORC TUTARI", "BORC"):
                hh["cikis"] = j
            elif "ISLEM TUTARI" in c or "TUTAR TL" in c or c == "TUTAR" or c == "MIKTAR" or c.startswith("TUTAR"):
                hh.setdefault("tutar", j)
            elif "ACIKLAMA" in c or c in ("ISLEM ACIKLAMASI", "DETAY"):
                hh.setdefault("aciklama", j)
            elif ("GONDER" in c or "UNVAN" in c or "KARSI TARAF" in c) and "aciklama" not in hh:
                hh["aciklama_ek"] = j
            elif "REFERANS" in c or "DEKONT" in c or "FIS NO" in c:
                hh.setdefault("referans", j)
            elif "ISLEM TIPI" in c or c == "ISLEM" or c == "KANAL":
                hh.setdefault("islem_tipi", j)
        if ("tarih" in hh or "tarih_valor" in hh) and ("tutar" in hh or "giris" in hh or "cikis" in hh):
            return i, hh
    return None, None


def _tablo_kayitlari(h):
    """Bir belgenin tablolarından hareketler."""
    out = []
    for tablo in h.get("tablolar", []) or []:
        if not tablo:
            continue
        bas, hh = _baslik_bul(tablo)
        if bas is None:
            continue
        tj = hh.get("tarih", hh.get("tarih_valor"))
        for row in tablo[bas + 1:]:
            def g(k):
                j = hh.get(k)
                return row[j] if j is not None and j < len(row) else None
            tarih, saat = tarih_saat(row[tj] if tj < len(row) else None)
            if not tarih:
                continue
            acik = str(g("aciklama") or "").strip()
            ek = str(g("aciklama_ek") or "").strip()
            if ek and norm(ek) not in norm(acik):
                acik = (ek + " " + acik).strip()
            if not acik and "aciklama" in hh and hh["aciklama"] + 1 < len(row):
                acik = str(row[hh["aciklama"] + 1] or "").strip()
            if "giris" in hh or "cikis" in hh:
                gi, ci = _sayi(g("giris")) or 0, _sayi(g("cikis")) or 0
                tutar = abs(gi) - abs(ci)
            else:
                tutar = _sayi(g("tutar")) or 0
            if not tutar:
                continue
            out.append({"tarih": tarih, "saat": saat, "aciklama": re.sub(r"\s+", " ", acik),
                        "tutar": round(tutar, 2), "bakiye": _sayi(g("bakiye")),
                        "referans": str(g("referans") or "").strip(),
                        "islem_tipi": str(g("islem_tipi") or "").strip()})
        if out:
            break
    return out


_TUTAR_RE = re.compile(r"(?<![\d.,])-?\d{1,3}(?:[.\s]\d{3})*[.,]\d{2}(?![\d])")
_ALTBILGI = ("SAYFA NO", "SAYFA", "DEKONT YERINE", "UYUSMAZLIK HALINDE", "MERSIS", "TICARET SICIL",
             "GENEL MUDURLUK", "VERGI DAIRESI", "WWW", "BANKASI A S", "T A S")


def _metin_kayitlari(h):
    """PDF/OCR metninden hareketler (tarihle başlayan satır = yeni hareket)."""
    metin = h.get("ham_metin") or ""
    if not metin:
        return []
    satirlar = [s.rstrip() for s in metin.splitlines()]
    tarih_bas = re.compile(r"^\s*(\d{1,2}[./-]\d{1,2}[./-](?:20)?\d{2}|20\d{2}[./-]\d{1,2}[./-]\d{1,2})(?:[\s\-]|$)")
    idx = [i for i, s in enumerate(satirlar) if tarih_bas.match(s) and _TUTAR_RE.search(s[10:])]
    if not idx:
        return []
    kayit = []
    for k, i in enumerate(idx):
        s = satirlar[i]
        m = tarih_bas.match(s)
        tarih, saat = tarih_saat(s[:m.end() + 12])
        if not tarih:
            continue
        geri = s[m.end():]
        sm = re.match(r"\s*(\d{1,2}:\d{2}(?::\d{2})?)", geri)
        if sm:
            saat = saat or sm.group(1); geri = geri[sm.end():]
        # açıklamadan önce gelen tutarlar (Halkbank: tutar, bakiye, açıklama)
        bas_tutarlar = []
        kalan = geri
        while True:
            tm = re.match(r"\s*(-?\d{1,3}(?:[.\s]\d{3})*[.,]\d{2})(?:\s*(?:TL|TRY|USD|EUR|GBP))?(?=\s|$)", kalan)
            if not tm:
                break
            bas_tutarlar.append(tm.group(1)); kalan = kalan[tm.end():]
        if bas_tutarlar:
            acik = kalan.strip()
            tutar = _sayi(bas_tutarlar[0])
            bakiye = _sayi(bas_tutarlar[1]) if len(bas_tutarlar) > 1 else None
        else:
            # açıklama önce, tutar(lar) sonda: sondan bir önceki işlem tutarı, sonuncusu bakiye
            tl = _TUTAR_RE.findall(geri)
            if not tl:
                continue
            if len(tl) >= 2:
                tutar, bakiye = _sayi(tl[-2]), _sayi(tl[-1])
            else:
                tutar, bakiye = _sayi(tl[-1]), None
            acik = geri
            for t in tl[-2:]:
                acik = acik.replace(t, " ")
            acik = re.sub(r"\b(TL|TRY)\b", " ", acik).strip()
        sonraki = idx[k + 1] if k + 1 < len(idx) else len(satirlar)
        devam = []
        for j in range(i + 1, sonraki):
            d = satirlar[j].strip()
            nd = norm(d)
            if not d or any(nd.startswith(a) or a in nd for a in _ALTBILGI[:5]) or nd.startswith(_ALTBILGI[5:]):
                if nd.startswith(("SAYFA", "TURKIYE")) or "DEKONT YERINE" in nd or "UYUSMAZLIK" in nd:
                    break
                continue
            if re.match(r"^[-\s|*]+$", d):
                continue
            devam.append(d)
        kayit.append({"tarih": tarih, "saat": saat, "aciklama": acik, "tutar": tutar, "bakiye": bakiye,
                      "devam": devam, "referans": "", "islem_tipi": ""})
    # devam satırları: hareket satırında açıklama varsa altındakiler onundur; yoksa (YKB) üstündekiler
    for k, r in enumerate(kayit):
        if r["aciklama"]:
            r["aciklama"] = re.sub(r"\s+", " ", " ".join([r["aciklama"]] + r["devam"])).strip()
        else:
            onceki = kayit[k - 1]["devam"] if k > 0 else []
            r["aciklama"] = re.sub(r"\s+", " ", " ".join(onceki + r["devam"])).strip()
    return [{k: v for k, v in r.items() if k != "devam"} for r in kayit if r["tutar"]]


def _isaret_dogrula(kayitlar):
    """Bakiye zinciriyle tutar işaretini ve sıralamayı doğrular.
    Döner: (kayitlar, zincir_kirik_sayisi). İşaret ters okunmuşsa çevirir."""
    b = [(r["tutar"], r["bakiye"]) for r in kayitlar]
    if sum(1 for _, x in b if x is not None) < 2:
        return kayitlar, 0
    say = Counter()
    for i in range(1, len(b)):
        (t0, b0), (t1, b1) = b[i - 1], b[i]
        if b0 is None or b1 is None:
            continue
        for isaret in (1, -1):
            if abs(b1 - (b0 + isaret * t1)) < 0.011:      # artan sıra
                say[("artan", isaret)] += 1
            if abs(b0 - (b1 + isaret * t0)) < 0.011:      # azalan sıra (en yeni üstte)
                say[("azalan", isaret)] += 1
    if not say:
        return kayitlar, 0
    (sira, isaret), n = say.most_common(1)[0]
    if isaret == -1:
        for r in kayitlar:
            r["tutar"] = -r["tutar"]
    kirik = (len(b) - 1) - n
    return kayitlar, max(kirik, 0)


def ekstre_kayitlari(hamlar):
    """Tüm belgelerden hareketler. Her kayda dosya ve sıra eklenir. Döner: (kayitlar, uyarilar)."""
    uyarilar, tum, gorulen = [], [], {}
    for h in hamlar:
        dosya = h.get("dosya", "")
        k = _tablo_kayitlari(h)
        if not k:
            k = _metin_kayitlari(h)
        if not k:
            if h.get("ham_metin") or h.get("tablolar"):
                uyarilar.append(f"{dosya}: hareket satırı bulunamadı — başlık (Tarih / Tutar / Açıklama) tanınmadı")
            continue
        k, kirik = _isaret_dogrula(k)
        if kirik:
            uyarilar.append(f"{dosya}: bakiye zinciri {kirik} yerde tutmuyor — eksik ya da yanlış okunan satır olabilir, kontrol edin")
        # aynı hareket başka bir dosyada da varsa (aynı ekstre iki kez yüklendi) atla
        imza = [(r["tarih"], r["saat"], r["tutar"], norm(r["aciklama"])[:40], r["bakiye"]) for r in k]
        ayni = next((d for d, im in gorulen.items() if im == imza), None)
        if ayni:
            uyarilar.append(f"{dosya}: {ayni} ile aynı ekstre — ikinci kez alınmadı")
            continue
        gorulen[dosya] = imza
        for i, r in enumerate(k):
            r["dosya"] = dosya; r["sira"] = i
            tum.append(r)
    return tum, uyarilar


# ------------------------------------------------------------------ geçmiş indeksi
class GecmisBanka:
    """Geçmiş fişlerdeki banka satırlarından (102 + karşı hesap aynı açıklamayla)
    açıklama kelimeleri -> karşı hesap örnekleri."""

    def __init__(self, satirlar):
        self.ornekler = []          # {kel, yon, banka, hesap, tutar}
        self.fis_aciklama = {}      # banka hesabı -> son fiş açıklaması
        fisler = defaultdict(list)
        for r in satirlar or []:
            fisler[str(r.get("fisno", ""))].append(r)
        for no, rs in fisler.items():
            # fişin bankası: en çok satırı olan 102 (virmanın karşı 102'si değil)
            say = Counter(r["hesap"] for r in rs if str(r.get("hesap", "")).startswith("102"))
            if say:
                ana_b = say.most_common(1)[0][0]
                fa, ft = rs[0].get("fis_aciklama") or "", rs[0].get("fis_tarih") or ""
                if fa and ft >= self.fis_aciklama.get(ana_b, ("",))[0]:
                    self.fis_aciklama[ana_b] = (ft, fa)
            gruplar = defaultdict(list)
            for r in rs:
                gruplar[_tarih_onekini_at(r.get("detay", "")).strip()].append(r)
            for detay, gs in gruplar.items():
                bankalar = [g for g in gs if str(g.get("hesap", "")).startswith("102")]
                if not bankalar:
                    continue
                kel = _kelimeler(detay)
                if not kel:
                    continue
                b0 = bankalar[0]
                yon = "giris" if float(b0.get("borc") or 0) > 0 else "cikis"
                karsilar = [g for g in gs if g is not b0 and g.get("hesap") != b0.get("hesap")]
                for kr in karsilar:
                    self.ornekler.append({"kel": kel, "yon": yon, "banka": b0["hesap"], "hesap": kr["hesap"],
                                          "tutar": round(float(kr.get("borc") or 0) + float(kr.get("alacak") or 0), 2)})

    def masraf_hesabi(self):
        c = Counter(o["hesap"] for o in self.ornekler
                    if o["yon"] == "cikis" and set(o["kel"]) & _MASRAF_KELIME and not o["hesap"].startswith("102"))
        return c.most_common(1)[0][0] if c else ""

    def adaylar(self, kel, yon, banka, tutar):
        """-> {hesap: (skor, adet)}"""
        wset = set(kel)
        sonuc = {}
        for o in self.ornekler:
            if o["yon"] != yon:
                continue
            ok = o["kel"]
            # çok genel tek kelimelik açıklamalar ("EFT") yalnız birebir aynıysa sayılır
            if not _hepsi_var(ok, wset, kel):
                continue
            skor = len(set(ok)) + (0.3 if o["banka"] == banka else 0) + (0.6 if abs(o["tutar"] - abs(tutar)) < 0.01 else 0)
            eski = sonuc.get(o["hesap"], (0, 0))
            sonuc[o["hesap"]] = (max(eski[0], skor), eski[1] + 1)
        return sonuc


# ------------------------------------------------------------------ vergi / SGK
_VERGI_KOD = {
    "0015": ("KDV", ("ODENECEK KDV", "KDV"), ("SORUMLU", "TEVKIF", "INDIRILECEK", "HESAPLANAN")),
    "4017": ("KDV tevkifatı", ("SORUMLU", "TEVKIF"), ()),
    "0003": ("muhtasar stopaj", ("STOPAJ", "GELIR VERGISI"), ("DAMGA",)),
    "0010": ("kurumlar vergisi", ("KURUMLAR",), ()),
    "0011": ("kurumlar vergisi", ("KURUMLAR",), ()),
    "1047": ("damga vergisi", ("DAMGA",), ()),
    "1048": ("damga vergisi", ("DAMGA",), ()),
    "0071": ("ÖTV", ("OTV", "OZEL TUKETIM"), ()),
    "9047": ("MTV", ("MTV", "MOTORLU TASIT"), ()),
    "0001": ("gelir vergisi", ("GELIR VERGISI",), ()),
}


def _vergi_bilgisi(aciklama):
    """'... / 0033-K.GEÇICI 3 AY / DÖNEM:04260626 ...' -> (kod, ad, donem_ay1)"""
    n = norm(aciklama)
    if not ("VERGI" in n and ("TAHSIL" in n or "ODEME" in n or "THK" in n)) and "GIB" not in n.split():
        return None
    m = re.search(r"(?<!\d)(\d{4})\s*-\s*([A-ZÇĞİÖŞÜa-zçğıöşü.]+)", aciklama)
    if not m:
        return ("", "", 0)
    donem = re.search(r"D[ÖO]NEM\s*:\s*(\d{2})(\d{2})(\d{2})(\d{2})", aciklama, re.I)
    return (m.group(1), norm(m.group(2)), int(donem.group(1)) if donem else 0)


def mizan_unvani(mizan_path) -> str:
    """Mizan raporunun başlığındaki firma unvanı (ilk satırlarda LTD/A.Ş./ŞİRKETİ geçen hücre)."""
    try:
        import openpyxl
        wb = openpyxl.load_workbook(mizan_path, read_only=True, data_only=True)
        ws = wb.worksheets[0]
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i > 8:
                break
            for c in row:
                n = norm(c)
                if len(n) > 8 and re.search(r"\b(LTD|LIMITED|ANONIM|AS|STI|SIRKETI|SIR)\b", n) and "HESAP" not in n:
                    return str(c).strip()
    except Exception:
        pass
    return ""


# ------------------------------------------------------------------ ana işleyici
def isle_banka(hamlar, km, fis0, banka_hesabi_bul=None, kur_getir=None):
    uyarilar = []
    hesaplar = list(km.hesaplar)
    kodlar = {k for k, _ in hesaplar}
    ad = dict(hesaplar)
    alt = {k for k in kodlar if not any(x != k and x.startswith(k + ".") for x in kodlar)}

    def hesap_ad(k):
        return km.hesap_adi(k) or ad.get(k, "")

    # ---------- dosya -> banka hesabı ve döviz
    dosya_hesap, dosya_doviz = {}, {}
    for h in hamlar:
        dosya = h.get("dosya", "")
        kod = ""
        for m in re.finditer(r"(?<![\d.])(\d{3}(?:\.\d{1,3}){1,3})(?![\d])", dosya):
            if m.group(1) in kodlar and m.group(1).startswith("10"):
                kod = m.group(1); break
        if not kod and banka_hesabi_bul:
            kod = banka_hesabi_bul([h], [{"kod": k, "ad": a} for k, a in hesaplar], uyarilar)
        dosya_hesap[dosya] = kod or "102.01.001"
        hn = norm(hesap_ad(dosya_hesap[dosya])) + " " + norm(dosya)
        dv = None
        for anahtar, kodd in (("USD", "USD"), ("DOLAR", "USD"), ("EUR", "EUR"), ("EURO", "EUR"),
                              ("GBP", "GBP"), ("STERLIN", "GBP"), ("CHF", "CHF")):
            if re.search(rf"\b{anahtar}\b", hn):
                dv = kodd; break
        dosya_doviz[dosya] = dv

    kayitlar, u = ekstre_kayitlari(hamlar)
    uyarilar += u
    if not kayitlar:
        return [], uyarilar
    for r in kayitlar:
        r["banka"] = dosya_hesap.get(r["dosya"], "102.01.001")
        r["doviz"] = dosya_doviz.get(r["dosya"])
        r["kel"] = _kelimeler(r["aciklama"])
        r["yon"] = "giris" if r["tutar"] > 0 else "cikis"

    gecmis = GecmisBanka(km.gecmis.get("satirlar", []))
    ogrenme = getattr(km, "ogrenme", {}) or {}
    unvan_k = [w for w in _kelimeler(getattr(km, "firma_unvan", "") or "") if w not in _GENEL][:3]

    def kendisi_mi(r):
        if len(unvan_k) < 1:
            return False
        es = sum(1 for w in unvan_k if _w_esler(w, set(r["kel"]), r["kel"]))
        return es >= min(2, len(unvan_k))

    # ---------- virman ve döviz alım/satım eşleri
    for r in kayitlar:
        r["es"] = None
    bankalar = defaultdict(list)
    for r in kayitlar:
        bankalar[r["banka"]].append(r)

    def saniye(r):
        m = re.match(r"(\d{2}):(\d{2}):(\d{2})", r.get("saat") or "")
        return int(m.group(1)) * 3600 + int(m.group(2)) * 60 + int(m.group(3)) if m else None

    adaylar_es = []
    for i, r in enumerate(kayitlar):
        doviz_satiri = bool(set(r["kel"]) & _DOVIZ_KELIME)
        for j in range(i + 1, len(kayitlar)):
            s_ = kayitlar[j]
            if s_["banka"] == r["banka"] or s_["tarih"] != r["tarih"] or (r["tutar"] > 0) == (s_["tutar"] > 0):
                continue
            sr, ss = saniye(r), saniye(s_)
            fark = abs(sr - ss) if (sr is not None and ss is not None) else 0
            if (r["doviz"] or "TL") == (s_["doviz"] or "TL"):
                if abs(abs(r["tutar"]) - abs(s_["tutar"])) > 0.005:
                    continue
                if not (kendisi_mi(r) or kendisi_mi(s_) or set(r["kel"] + s_["kel"]) & _VIRMAN_KELIME):
                    continue
            else:
                if not (doviz_satiri and set(s_["kel"]) & _DOVIZ_KELIME):
                    continue
                if fark > 120:
                    continue
            adaylar_es.append((fark, i, j))
    for fark, i, j in sorted(adaylar_es):
        r, s_ = kayitlar[i], kayitlar[j]
        if r["es"] is None and s_["es"] is None:
            r["es"], s_["es"] = s_, r

    masraf_hesap = gecmis.masraf_hesabi()
    if not masraf_hesap:
        masraf_hesap = next((k for k, a in hesaplar if k in alt and k.startswith(("770", "780", "653", "689"))
                             and "BANKA" in norm(a)), "")

    # ---------- karşı hesap
    cari_onek = ("120", "121", "126", "131", "136", "159", "195", "196", "300", "303", "320", "321",
                 "329", "331", "335", "336", "340", "400")
    mizan_adaylari = []
    for k, a in hesaplar:
        if k in alt and k.startswith(cari_onek):
            kel = [w for w in _kelimeler(a) if w not in _GENEL and w not in _HESAP_GENEL]
            if kel and (len(kel) >= 2 or len(kel[0]) >= 5):
                mizan_adaylari.append((k, kel))

    def vergi_hesabi(r):
        vb = _vergi_bilgisi(r["aciklama"])
        if vb is None:
            return None
        kod, vad, donem = vb
        if not kod:
            return ("", "vergi kodu okunamadı")
        if "GECICI" in vad or kod in ("0032", "0033"):
            ceyrek = (donem - 1) // 3 + 1 if donem else 0
            gec = [(k, a) for k, a in hesaplar if k in alt and k.startswith(("193", "371"))
                   and "GECICI" in norm(a)]
            sec = next((k for k, a in gec if ceyrek and re.search(rf"(?<!\d){ceyrek}\s*\.?\s*DONEM", norm(a))), "")
            return (sec or (gec[0][0] if len(gec) == 1 else ""), f"{kod} geçici vergi {ceyrek}. dönem")
        tanim = _VERGI_KOD.get(kod)
        if not tanim:
            return ("", f"{kod} vergi kodu tanımlı değil")
        ad_, ara, haric = tanim
        adaylar = [(k, norm(a)) for k, a in hesaplar if k in alt and k.startswith(("360", "361", "368", "369"))]
        for anahtar in ara:
            bul = [k for k, a in adaylar if anahtar in a and not any(x in a for x in haric)]
            if bul:
                # birden çoksa geçmişte en çok kullanılan
                if len(bul) > 1:
                    say = Counter(o["hesap"] for o in gecmis.ornekler if o["hesap"] in bul)
                    bul.sort(key=lambda k: -say.get(k, 0))
                return (bul[0], f"{kod} {ad_}" + (" — birden çok aday" if len(bul) > 1 else ""))
        return ("", f"{kod} {ad_} — mizanda hesap yok")

    def sgk_hesabi(r):
        n = " " + norm(r["aciklama"]) + " "
        if not (re.search(r"[^A-Z]SGK", n) or "SOSYAL GUVENLIK" in n):
            return None
        bul = [k for k, a in hesaplar if k in alt and k.startswith("361") and ("SGK" in norm(a) or "SOSYAL" in norm(a))]
        return bul[0] if bul else ""

    def karsi_bul(r):
        """-> (hesap, kaynak, kontrol_nedeni|None, not)"""
        kel, wset = r["kel"], set(r["kel"])
        # 1) öğrenilen
        n_ = norm(r["aciklama"])
        masraf_mi = r.get("masraf") or bool(wset & _MASRAF_KELIME)
        en_iyi = None
        for anahtar, kod in ogrenme.items():
            if kod not in kodlar or not anahtar:
                continue
            ak = anahtar.split()
            # transferin öğrenilmiş hesabı onun komisyon/BSMV satırına geçmesin
            if masraf_mi and not (set(ak) & _MASRAF_KELIME):
                continue
            if (ak and all(not any(c.isdigit() for c in w) for w in ak) and _hepsi_var(ak, wset, kel)) or anahtar in n_:
                if en_iyi is None or len(ak) > en_iyi[0]:
                    en_iyi = (len(ak), kod)
        if en_iyi:
            return en_iyi[1], "ogrenme", None, ""
        # 3) masraf / BSMV / komisyon
        if r.get("masraf") or (set(kel) & _MASRAF_KELIME and r["yon"] == "cikis" and abs(r["tutar"]) < 5000
                               and not kendisi_mi(r)):
            if masraf_hesap:
                return masraf_hesap, "masraf", None, ""
        # 4) vergi / SGK
        v = vergi_hesabi(r) if r["yon"] == "cikis" else None
        if v is not None:
            kod, aciklama_ = v
            return (kod or ""), "vergi", (None if kod and "aday" not in aciklama_ else "vergi"), \
                f"Vergi ödemesi ({aciklama_}); gecikme zammı varsa ayrı satıra ayırın"
        s = sgk_hesabi(r) if r["yon"] == "cikis" else None
        if s is not None:
            return s, "sgk", (None if s else "vergi"), "SGK ödemesi"
        # 5) geçmiş + mizan adı
        tutar = r["tutar"]
        gec = gecmis.adaylar(kel, r["yon"], r["banka"], tutar)
        skor = {}
        for k, (sk, adet) in gec.items():
            # Logo açıklamayı kestiği için transferin komisyon satırı geçmişte transferle aynı
            # görünür ("ALP ARAYICI GARANTI" -> 770); masraf olmayan satırda masraf hesabı aday değil
            if k in kodlar and k != r["banka"] and not (k == masraf_hesap and not masraf_mi):
                skor[k] = [sk, adet, "gecmis"]
        bulanik_ = set()
        for k, hk in mizan_adaylari:
            tam = _hepsi_var(hk, wset, kel)
            if tam or _hepsi_var(hk, wset, kel, bulanik=True):
                ms = len(hk) + 0.5
                if not tam:
                    bulanik_.add(k)
                if k in skor:
                    skor[k][0] = max(skor[k][0], ms) + 0.5; skor[k][2] = "gecmis"
                else:
                    skor[k] = [ms, 0, "mizan"]
        # mizanda tek kelimesi tutan (TURKCELL İLETİŞİM ...): ilk ayırt edici kelime yalnız o hesapta
        if not skor:
            ilk = defaultdict(list)
            for k, hk in mizan_adaylari:
                if len(hk[0]) >= 5:
                    ilk[hk[0]].append(k)
            for w_, ks in ilk.items():
                if len(set(ks)) == 1 and _w_esler(w_, wset, kel):
                    skor[ks[0]] = [1.5, 0, "mizan"]
        if not skor:
            return "", "", "eslesmedi", ""
        sirali = sorted(skor.items(), key=lambda x: (-x[1][0], -x[1][1]))
        en, (sk, adet, kaynak) = sirali[0]
        kontrol = None
        not_ = ""
        if len(sirali) > 1 and sirali[1][1][0] >= sk - 0.25:
            kontrol = "belirsiz"
            not_ = "Aynı açıklama geçmişte başka hesaplara da yazılmış: " + \
                   ", ".join(k for k, _ in sirali[1:4])
        elif sk < 2:
            kontrol = "zayif"
            not_ = "Açıklamanın yalnız bir kelimesi geçmişle / hesap adıyla tuttu"
        elif en in bulanik_ and kaynak == "mizan":
            kontrol = "benzer"
            not_ = f"Hesap adı ({hesap_ad(en)}) açıklamayla bir harf farklı"
        return en, kaynak, kontrol, not_

    # masraf satırları: aynı anda yapılan transferin "- Komisyon" / "- Vergi (BSMV)" satırları
    for b, rs in bankalar.items():
        for r in rs:
            if r["yon"] != "cikis":
                continue
            nr = norm(r["aciklama"])
            for s in rs:
                if s is r or s["tarih"] != r["tarih"] or abs(s["tutar"]) <= abs(r["tutar"]):
                    continue
                ns = norm(s["aciklama"])
                if len(ns) >= 6 and nr != ns and nr.startswith(ns[:max(6, len(ns) - 4)]) and \
                        (set(_kelimeler(nr[len(ns) - 4:])) & (_MASRAF_KELIME | {"VERGI"})):
                    sr, ss = saniye(r), saniye(s)
                    if sr is None or ss is None or abs(sr - ss) <= 5:
                        r["masraf"] = True
                        break

    # ---------- fiş açıklaması (geçmişteki biçimle)
    def fis_aciklamasi(banka, tarih):
        onceki = gecmis.fis_aciklama.get(banka)
        gun = f"{tarih[8:10]}.{tarih[5:7]}.{tarih[:4]}"
        if onceki:
            _, fa = onceki
            m = re.search(r"(\d{2})([./-])(\d{2})\2(\d{4})", fa)
            if m:
                ay = m.group(2)
                return fa[:m.start()] + f"{tarih[8:10]}{ay}{tarih[5:7]}{ay}{tarih[:4]}" + fa[m.end():]
        return f"{hesap_ad(banka) or 'BANKA'} - {gun}"

    # ---------- kur
    kur_onbellek = {}

    def kur(dv, tarih):
        if not dv or not kur_getir:
            return None
        if (dv, tarih) not in kur_onbellek:
            try:
                kur_onbellek[(dv, tarih)] = kur_getir(dv, tarih)
            except Exception:
                kur_onbellek[(dv, tarih)] = None
        return kur_onbellek[(dv, tarih)]

    # ---------- fişler: (banka, tarih) başına bir fiş
    def satir(fisno, tarih, fis_ac, hesap, borc, alacak, detay, kaynak, rol, r=None, dv=None, kur_=None,
              dv_tutar=None, kontrol=None, not_=""):
        s = {"fisno": fisno, "fis_tarih": tarih, "fis_aciklama": fis_ac, "hesap": hesap,
             "evrak_no": (r or {}).get("referans", "") if r else "", "evrak_tarih": tarih,
             "detay": detay, "borc": round(borc, 2), "alacak": round(alacak, 2), "belge_turu": "MF",
             "kaynak": kaynak, "para_birimi": dv or "", "kur": kur_, "doviz_tutar": dv_tutar, "rol": rol}
        if kontrol:
            s["kontrol"] = [kontrol]
        if not_:
            s["not"] = not_
        return s

    hareketler = []          # (fiş bankası, tarih, saat, sıra, [satırlar])
    kur_yok, eslesmeyen, ozet = set(), [], Counter()
    for r in kayitlar:
        if r.get("_atla"):
            continue
        es = r["es"]
        if es is not None:
            es["_atla"] = True
            # döviz alım/satımı: TL tarafının fişine; virman: çıkan hesabın fişine
            tl_mi = (r["doviz"] or "TL") == "TL"
            es_tl = (es["doviz"] or "TL") == "TL"
            if tl_mi != es_tl:
                tl, dvz = (r, es) if tl_mi else (es, r)
                ana = tl
            else:
                tl = dvz = None
                ana = r if r["tutar"] < 0 else es
            giren, cikan = (r, es) if r["tutar"] > 0 else (es, r)
            tutar_tl = abs(tl["tutar"]) if tl else abs(ana["tutar"])
            dv_k = dvz["doviz"] if dvz else (ana["doviz"])
            dv_t = abs(dvz["tutar"]) if dvz else (abs(ana["tutar"]) if ana["doviz"] else None)
            kur_ = round(tutar_tl / dv_t, 6) if (dvz and dv_t) else None
            if not dvz and ana["doviz"]:
                kur_ = kur(ana["doviz"], ana["tarih"])
                if kur_:
                    tutar_tl = round(abs(ana["tutar"]) * kur_, 2)
                else:
                    kur_yok.add(ana["doviz"])
            detay = ana["aciklama"]
            satirlar = []
            for x, b_, a_ in ((giren, tutar_tl, 0), (cikan, 0, tutar_tl)):
                xdv = x["doviz"]
                satirlar.append((x["banka"], b_, a_, "banka" if x is ana else "virman", xdv,
                                 kur_ if xdv else None, abs(x["tutar"]) if xdv else None))
            hareketler.append((ana["banka"], ana["tarih"], ana.get("saat") or "", ana["sira"], ana, detay, satirlar,
                               None, "Döviz alım/satımı" if dvz else "Hesaplar arası virman"))
            ozet["döviz alım/satım" if dvz else "virman"] += 1
            continue
        hesap, kaynak, kontrol, not_ = karsi_bul(r)
        if not hesap:
            hesap = "198.01.001" if "198.01.001" in kodlar or not kodlar else \
                next((k for k in sorted(kodlar) if k.startswith("198") and k in alt), "198.01.001")
            kontrol = kontrol or "eslesmedi"
            eslesmeyen.append(r["aciklama"][:40])
        ozet[kaynak or "eşleşmedi"] += 1
        tutar = abs(r["tutar"])
        kur_, dv_t = None, None
        if r["doviz"]:
            kur_ = kur(r["doviz"], r["tarih"])
            if kur_:
                dv_t = tutar; tutar = round(tutar * kur_, 2)
            else:
                kur_yok.add(r["doviz"])
        if r["tutar"] > 0:
            satirlar = [(r["banka"], tutar, 0, "banka", r["doviz"], kur_, dv_t),
                        (hesap, 0, tutar, "karsi", r["doviz"], kur_, dv_t)]
        else:
            satirlar = [(hesap, tutar, 0, "karsi", r["doviz"], kur_, dv_t),
                        (r["banka"], 0, tutar, "banka", r["doviz"], kur_, dv_t)]
        hareketler.append((r["banka"], r["tarih"], r.get("saat") or "", r["sira"], r, r["aciklama"], satirlar,
                           (kontrol, not_, kaynak), None))

    fisler = []
    fis = fis0
    gruplu = defaultdict(list)
    for h in hareketler:
        gruplu[(h[0], h[1])].append(h)
    for (banka, tarih) in sorted(gruplu, key=lambda x: (x[1], x[0])):
        fisno = f"{fis:05d}"
        fa = fis_aciklamasi(banka, tarih)
        for _, _, saat, sira, r, detay, satirlar, kk, not_virman in sorted(gruplu[(banka, tarih)],
                                                                            key=lambda h: (h[2], -h[3])):
            for hesap, b_, a_, rol, dv, kur_, dv_t in satirlar:
                kontrol, not_, kaynak = (kk or (None, not_virman or "", "virman"))
                if rol != "karsi":
                    kontrol_ = None
                    kaynak_ = "banka" if rol == "banka" else "virman"
                    n_ = not_virman if rol == "virman" else ""
                else:
                    kontrol_, kaynak_, n_ = kontrol, kaynak, not_
                fisler.append(satir(fisno, tarih, fa, hesap, b_, a_, detay, kaynak_, rol, r, dv, kur_, dv_t,
                                    kontrol_, n_))
        fis += 1

    if ozet:
        parca = ", ".join(f"{v} {k}" for k, v in ozet.most_common())
        uyarilar.insert(0, f"{len(kayitlar)} banka hareketi okundu ({len(set(r['dosya'] for r in kayitlar))} ekstre): {parca}")
    if kur_yok:
        uyarilar.append(f"{', '.join(sorted(kur_yok))} kuru TCMB'den alınamadı — bu satırlar döviz tutarıyla yazıldı, kontrol edin")
    if eslesmeyen:
        uyarilar.append(f"{len(eslesmeyen)} hareketin karşı hesabı bulunamadı (198'e yazıldı, KONTROL): "
                        + "; ".join(eslesmeyen[:5]) + (" …" if len(eslesmeyen) > 5 else ""))
    kontrollu = sorted({s["fisno"] for s in fisler if s.get("kontrol")})
    if kontrollu:
        uyarilar.append(f"{sum(1 for s in fisler if s.get('kontrol'))} satırda karşı hesap kesin değil — sarı KONTROL "
                        f"işaretlileri düzelt ya da ✓ ile onayla; düzeltme öğrenilir")
    return fisler, uyarilar
