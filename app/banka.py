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
                  "UCRETLERI", "ISLETIM", "AIDAT", "AIDATI", "KESINTI"}
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
_SAAT = re.compile(r"(?<![\d.,])(\d{1,2})[:.](\d{2})(?:[:.](\d{2}))?(?:\.(\d+))?(?![\d,])")


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
    if sm and (int(sm.group(1)) > 23 or int(sm.group(2)) > 59):
        sm = None                      # "73.745,17" bir tutarın başı, saat değil
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
            elif "ISLEM TIPI" in c or c == "ISLEM" or c == "ETIKET" or c == "ISLEM TURU":
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


_SAYFA_ALTI = re.compile(r"^(\(\*\)|SAYFA\b|SF\b|\d+\s*/\s*\d+$|WWW|HTTPS?|BU HESAP OZETI|ISLEM SAATLERI|DEKONT YERINE|"
                         r"UYUSMAZLIK|MERSIS|TICARET SICIL|GENEL MUDURLUK|BUYUK MUKELLEFLER|FINANSKENT|"
                         r"T GARANTI BANKASI|TURKIYE IS BANKASI|YAPI VE KREDI BANKASI|TURKIYE HALK BANKASI|"
                         r"TURKIYE VAKIFLAR BANKASI|AKBANK T A S|DENIZBANK A S|2026 DENIZBANK)")
_BASLIK_ETIKET = ("MUSTERI ADI", "MUSTERI NUMARASI", "MUSTERI BILGI", "HESAP BILGI", "HESAP NO", "HESAP ADI", "HESAP TURU",
                  "IBAN", "TARIH ARALIGI", "KULLANILABILIR BAKIYE", "UNVAN", "TCKN", "VKN", "SUBE", "URUN DOVIZ",
                  "MEVCUT BAKIYE", "NET BAKIYE", "ISLEM YERI", "BELGE DUZENLEME", "SORGULAMA", "HESAP OZETI",
                  "HESAP HAREKETLERI", "BAKIYE BILGI", "EK HESAP LIMITI", "ORTAK HESAP", "RUMUZ", "VB MUS NO",
                  "DONEMI", "URETIM ZAMANI", "BLOKE BAKIYE", "KREDI LIMITI", "KREDILI BAKIYE", "DOVIZ CINSI")
_TARIH_BAS = re.compile(r"^\s*(\d{1,2}[./-]\d{1,2}[./-](?:20)?\d{2}|20\d{2}[./-]\d{1,2}[./-]\d{1,2})(?=[\s\-]|$)")
_TUTAR_TOKEN = re.compile(r"(?<![\d.,])([+-]?\d{1,3}(?:[.\s]\d{3})*[.,]\d{2})(?![\d])(?:\s*(?:TL|TRY|USD|EUR|GBP)\b)?")


def _baslik_satiri(s):
    n = norm(s)
    return ("TARIH" in n and ("BAKIYE" in n or "TUTAR" in n or "MIKTAR" in n)) or \
        (":" in s and any(n.startswith(e) or (" " + e) in (" " + n[:40]) for e in _BASLIK_ETIKET)) or \
        "BEKLEYEN ISLEM" in n or bool(re.match(r"^[-\s|*]+$", s.strip()))


def _metin_kayitlari(h):
    """PDF/OCR metninden hareketler.
    * Her sayfanın sütun başlığına kadarki kısmı (müşteri/hesap bilgisi) ve altbilgisi atılır.
    * Tarihle başlayan ve tutar içeren satır = hareket. Tutar tarih satırında yoksa bir üstteki
      tarihsiz tutar satırından alınır (Garanti PDF: açıklama+tutar / tarih / açıklama devamı).
    * Açıklama yönü belgeden anlaşılır: son hareketin altında devam satırı varsa açıklamalar
      hareketin ALTINDA (Halkbank, Vakıfbank); ilk hareketin üstünde satır varsa ya da son hareketin
      altında yoksa ÜSTÜNDE (İş Bankası, Yapı Kredi).
    * Satırda birden çok tutar varsa (tutar, bakiye, ek hesap bakiyesi) hangisinin tutar olduğu
      bakiye zincirine en iyi uyan sırayla seçilir."""
    sayfalar = h.get("sayfalar") or ([h.get("ham_metin")] if h.get("ham_metin") else [])
    if not sayfalar:
        return []
    satirlar = []          # (metin, sayfa_basi_mi)
    for sayfa in sayfalar:
        sat = [x.rstrip() for x in (sayfa or "").splitlines()]
        bas = next((i for i, x in enumerate(sat) if _baslik_satiri(x) and "TARIH" in norm(x)), None)
        if bas is not None:
            sat = sat[bas + 1:]
        for x in sat:
            if _SAYFA_ALTI.match(norm(x)) or norm(x).startswith("***"):
                break       # sayfanın altbilgisi
            if x.strip() and not _baslik_satiri(x):
                satirlar.append(x)
    # Garanti PDF: tarihsiz tutar satırı + tutarsız tarih satırı -> tek satır
    birlesik, i = [], 0
    while i < len(satirlar):
        x = satirlar[i]
        if (i + 1 < len(satirlar) and not _TARIH_BAS.match(x) and _TUTAR_TOKEN.search(x)
                and _TARIH_BAS.match(satirlar[i + 1]) and not _TUTAR_TOKEN.search(satirlar[i + 1][12:])):
            m = _TARIH_BAS.match(satirlar[i + 1])
            birlesik.append(satirlar[i + 1][:m.end()] + " " + x.strip() + " " + satirlar[i + 1][m.end():].strip())
            i += 2
            continue
        birlesik.append(x)
        i += 1
    satirlar = birlesik
    idx = [i for i, s in enumerate(satirlar) if _TARIH_BAS.match(s) and _TUTAR_TOKEN.search(s[_TARIH_BAS.match(s).end():])]
    if not idx:
        return []

    def parcala(s):
        m = _TARIH_BAS.match(s)
        tarih, saat = tarih_saat(s[:m.end() + 15])
        geri = s[m.end():]
        sm = re.match(r"[\s\-]*(\d{1,2}:\d{2}(?::\d{2})?)(?![\d,])", geri)
        if sm:
            saat = saat or sm.group(1); geri = geri[sm.end():]
        tl = [(t.group(1), t.span()) for t in _TUTAR_TOKEN.finditer(geri)]
        return tarih, saat, geri, tl

    ham = [parcala(satirlar[i]) for i in idx]

    def secim(tl, strateji):
        if not tl:
            return None, None, []
        if len(tl) == 1:
            return _sayi(tl[0][0]), None, [0]
        if strateji == "ilk":
            return _sayi(tl[0][0]), _sayi(tl[1][0]), [0, 1]
        return _sayi(tl[-2][0]), _sayi(tl[-1][0]), [len(tl) - 2, len(tl) - 1]

    def zincir(strateji):
        d = [secim(x[3], strateji)[:2] for x in ham]
        say = 0
        for (t0, b0), (t1, b1) in zip(d, d[1:]):
            if None in (t0, b0, t1, b1):
                continue
            if any(abs(b1 - (b0 + i * t1)) < 0.011 or abs(b0 - (b1 + i * t0)) < 0.011 for i in (1, -1)):
                say += 1
        return say
    strateji = "ilk" if zincir("ilk") >= zincir("son") else "son"

    # açıklama yönü
    once = sum(1 for s in satirlar[:idx[0]])
    sonra = sum(1 for s in satirlar[idx[-1] + 1:])
    nm = norm(h.get("ham_metin", "")[:3000])
    if once > sonra:
        yon = "ust"
    elif sonra > once:
        yon = "alt"
    else:
        yon = "ust" if ("YAPI VE KREDI" in nm or "YAPIKREDI" in nm or "IS BANKASI" in nm or "ISCEP" in nm) else "alt"

    kayit = []
    for k, (i, (tarih, saat, geri, tl)) in enumerate(zip(idx, ham)):
        if not tarih:
            continue
        tutar, bakiye, kullan = secim(tl, strateji)
        if not tutar:
            continue
        acik = geri
        for j in sorted(kullan, reverse=True):
            a, b = tl[j][1]
            acik = acik[:a] + " " + acik[b:]
        acik = re.sub(r"\b(TL|TRY)\b", " ", acik)
        if yon == "alt":
            sonraki = idx[k + 1] if k + 1 < len(idx) else len(satirlar)
            ek = satirlar[i + 1:sonraki]
            parca = [acik] + [x.strip() for x in ek]
        else:
            onceki = idx[k - 1] + 1 if k > 0 else 0
            ek = satirlar[onceki:i]
            parca = [x.strip() for x in ek] + [acik]
        kayit.append({"tarih": tarih, "saat": saat, "aciklama": re.sub(r"\s+", " ", " ".join(parca)).strip(" -|"),
                      "tutar": tutar, "bakiye": bakiye, "referans": "", "islem_tipi": ""})
    return kayit


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
        if h.get("tur") in ("pdf", "pdf_ocr") and h.get("ham_metin"):
            # PDF tablosu çoğu zaman yalnız bir sayfada tanınır; metinden okunan daha eksiksizse o
            km_ = _metin_kayitlari(h)
            if len(km_) > len(k):
                k = km_
        if not k:
            k = _metin_kayitlari(h)
        if not k:
            tum_metin = norm((h.get("ham_metin") or "") + " " + " ".join(
                str(c) for t in (h.get("tablolar") or []) for row in t for c in row if c))
            if "HESAP HAREKETLERINDE ARA" in tum_metin or ("TUTAR ARALIGI" in tum_metin and "GORUNTULE" in tum_metin):
                uyarilar.append(f"{dosya}: bu PDF ekstre değil, internet şubesindeki arama ekranının çıktısı — hareket listesi yok. "
                                f"Bankadan 'Hesap Hareketleri'ni Excel ya da ekstre PDF'i olarak indirin")
            elif re.search(r"KAYIT BULUNMU|HAREKET\w* BULUNMA|ISLEM BULUNMA|KAYIT YOK|HAREKET YOK", tum_metin):
                uyarilar.append(f"{dosya}: bu dönemde hareket yok (ekstre boş)")
            elif h.get("ham_metin") or h.get("tablolar"):
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
                if len(karsilar) > 4 or re.search(r"ACILIS|KAPANIS|DEVIR", norm(detay)):
                    continue        # açılış/kapanış fişi: banka hareketi değil
                for kr in karsilar:
                    self.ornekler.append({"kel": kel, "yon": yon, "banka": b0["hesap"], "hesap": kr["hesap"],
                                          "tutar": round(float(kr.get("borc") or 0) + float(kr.get("alacak") or 0), 2),
                                          "tarih": b0.get("fis_tarih") or ""})

    def ayikla(self, cikar):
        """Firmanın kendi unvan kelimelerini örneklerden çıkarır (müşteri açıklamalarına
        'AREL NAKLİYAT CARİ ÖDEME' gibi bizim adımız yazılır; eşleşmeyi bozmasın)."""
        if not cikar:
            return
        for o in self.ornekler:
            k2 = [w for w in o["kel"] if w not in cikar]
            o["kel"] = k2 if k2 else ["__KENDI__"]

    def nadir(self, w):
        """Kelime geçmiş banka açıklamalarının azında geçiyor (firma/kişi adı; 'CARI', 'HAVALE' değil)."""
        if not hasattr(self, "_df"):
            self._df = Counter(x for o in self.ornekler for x in set(o["kel"]))
            self._n = max(1, len(self.ornekler))
        return 0 < self._df.get(w, 0) <= max(6, self._n * 0.03)

    def masraf_hesabi(self, hesap_adlari=None):
        """Geçmişte banka masrafına kullanılan hesap. Yalnız küçük tutarlı (< 1.000) masraf/komisyon
        satırları sayılır — 'Şirket içi masrafları' açıklamalı büyük ortak ödemeleri sayılmasın;
        adında BANKA / KOMİSYON geçen hesap öne geçer."""
        c = Counter(o["hesap"] for o in self.ornekler
                    if o["yon"] == "cikis" and set(o["kel"]) & _MASRAF_KELIME and 0 < o["tutar"] < 1000
                    and not o["hesap"].startswith("102"))
        if not c:
            return ""
        ad = hesap_adlari or {}
        return max(c, key=lambda h: (c[h] * (3 if re.search(r"BANKA|KOMISYON|MASRAF", norm(ad.get(h, ""))) else 1), c[h]))

    def adaylar(self, kel, yon, banka, tutar, tarih=""):
        """-> {hesap: {"taban", "tutar_b", "adet", "agirlik"}}
        taban   : tutan kelime sayısı (+0,3 aynı banka)
        tutar_b : tutar yakınlığı 0..0,6 (aynı açıklama hem transfer hem masrafı için kullanılmışsa:
                  42.600 -> avans, 16,76 -> masraf)
        agirlik : yakın aylarda kullanım (her ay geriye 1/4) — hesap zamanla değişmişse son kullanılan"""
        import math
        wset = set(kel)
        out = {}
        a = abs(tutar)
        try:
            ty, ta = int(tarih[:4]), int(tarih[5:7])
        except (ValueError, TypeError):
            ty = ta = None
        for o in self.ornekler:
            if o["yon"] != yon:
                continue
            ok = o["kel"]
            if not _hepsi_var(ok, wset, kel):
                continue
            h = o["hesap"]
            d = out.setdefault(h, {"taban": 0, "tutar_b": 0, "adet": 0, "agirlik": 0.0, "kelimeler": set()})
            d["kelimeler"].update(ok)
            d["taban"] = max(d["taban"], len(set(ok)) + (0.3 if o["banka"] == banka else 0))
            if abs(o["tutar"] - a) < 0.01:
                tb = 1.0                       # birebir aynı tutar (aylık kira gibi): güçlü işaret
                d["ayni_tutar"] = True
            elif o["tutar"] > 0 and a > 0:
                tb = 0.4 * max(0.0, 1 - abs(math.log10(a / o["tutar"])) / 2)
            else:
                tb = 0
            d["tutar_b"] = max(d["tutar_b"], tb)
            d["adet"] += 1
            ot = o.get("tarih") or ""
            if ty and len(ot) >= 7:
                ay_once = max(0, (ty - int(ot[:4])) * 12 + ta - int(ot[5:7]))
                d["agirlik"] += 0.25 ** ay_once
            else:
                d["agirlik"] += 0.05
        return out


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
    m = re.search(r"(?<![\d/])(\d{4})\s*[-/]\s*([A-ZÇĞİÖŞÜa-zçğıöşü][A-ZÇĞİÖŞÜa-zçğıöşü. ]+)", aciklama)
    if not m:
        return ("", "", 0)
    # "DÖNEM:04260626" (ayyy ayyy) ya da "Dönem :04/2026/06/2026"
    donem = re.search(r"D[ÖO]NEM\s*:\s*(\d{2})(?:\d{2}|/\d{4})", aciklama, re.I)
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
            # hesap satırlarına gelindiyse başlık bitti (müşteri adı firmanın unvanı sanılmasın)
            if any(re.match(r"^\s*\d{3}([.\s]|$)", str(c or "")) for c in row[:2]):
                break
            for c in row:
                n = norm(c)
                if len(n) > 8 and re.search(r"\b(LTD|LIMITED|ANONIM|AS|STI|SIRKETI|SIR)\b", n) and "HESAP" not in n:
                    return str(c).strip()
    except Exception:
        pass
    return ""


def ekstre_unvani(hamlar) -> str:
    """Ekstre başlığındaki hesap sahibi unvanı ("Ad Soyad/Ünvan", "Şirket Ünvanı", "Müşteri Adı").
    En az iki anlamlı kelime yoksa (PDF'te satır kayması: "LIMITED SIR") kabul edilmez."""
    def gecerli(v):
        return len([w for w in _kelimeler(v) if w not in _GENEL and w not in ("SIR", "SIRK", "LIMITE")]) >= 2
    for h in hamlar:
        for t in h.get("tablolar") or []:
            for row in t[:25]:
                hucre = [str(c).strip() for c in row if c not in (None, "") and str(c).strip()]
                for i, c in enumerate(hucre[:-1]):
                    n = norm(c)
                    if ("UNVAN" in n or n in ("AD SOYAD", "MUSTERI ADI", "HESAP SAHIBI")) and len(n) < 40:
                        if gecerli(hucre[i + 1]):
                            return re.split(r"\s{2,}|\bBakiye\b", hucre[i + 1])[0].strip()
        m = re.search(r"(?:[ÜU]nvan[ıi]?|M[üu][şs]teri Ad[ıi])[^:\n]{0,15}:\s*([^\n]{6,80})", h.get("ham_metin") or "")
        if m and gecerli(m.group(1)):
            return re.split(r"\s{2,}|\bBakiye\b|\bBAKIYE\b", m.group(1))[0].strip()
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
    unvan = getattr(km, "firma_unvan", "") or ekstre_unvani(hamlar) or getattr(km, "firma_adi", "") or ""
    unvan_k = [w for w in _kelimeler(unvan) if w not in _GENEL][:3]

    def kendisi_mi(r):
        if len(unvan_k) < 1:
            return False
        kt = r.get("kel_tam", r["kel"])
        es = sum(1 for w in unvan_k if _w_esler(w, set(kt), kt))
        return es >= min(2, len(unvan_k))

    # müşteri açıklamalarında geçen bizim unvanımız ("... AREL NAKLİYAT CARİ ÖDEME") eşleşmeyi bozmasın
    cikar = set(unvan_k)
    gecmis.ayikla(cikar)
    for r in kayitlar:
        r["kel_tam"] = r["kel"]
        if cikar:
            k2 = [w for w in r["kel"] if w not in cikar]
            r["kel"] = k2 + (["__KENDI__"] if kendisi_mi(r) else [])

    # hesap no ipucu: "60363919 - 354 Hesaba Para Transferi" -> 354 yalnız 102.01.08'in dosya adında/adında
    def _nolar(metin):
        return set(re.findall(r"(?<![\d.])(\d{3,})(?![\d.])", metin or ""))
    hesap_nolari = {}
    for dsy, kod in dosya_hesap.items():
        hesap_nolari.setdefault(kod, set()).update(_nolar(re.sub(r"_?10\d(?:\.\d+)+", " ", dsy)) | _nolar(hesap_ad(kod)))
    for kod in list(hesap_nolari):
        digerleri = set().union(*[v for k, v in hesap_nolari.items() if k != kod]) if len(hesap_nolari) > 1 else set()
        hesap_nolari[kod] = hesap_nolari[kod] - digerleri

    def _no_esit(a, b):
        a, b = a.lstrip("0"), b.lstrip("0")
        return bool(a and b) and (a == b or (min(len(a), len(b)) >= 6 and (a.endswith(b) or b.endswith(a))))

    def _banka_adi(metin):
        n = " " + norm(metin) + " "
        out = set()
        for anahtar, ad_ in (("VAKIF", "VAKIF"), ("GARANTI", "GARANTI"), (" YKB", "YKB"), ("YAPI KREDI", "YKB"),
                             ("YAPIKREDI", "YKB"), ("ISBANK", "IS"), (" IS BANK", "IS"), ("HALKBANK", "HALK"),
                             (" HALK ", "HALK"), ("AKBANK", "AKBANK"), ("DENIZ", "DENIZ"), ("ZIRAAT", "ZIRAAT"),
                             (" QNB", "QNB"), ("FINANSBANK", "QNB"), (" TEB ", "TEB"), (" ING ", "ING"),
                             ("KUVEYT", "KUVEYT"), ("ALBARAKA", "ALBARAKA"), ("SEKERBANK", "SEKER")):
            if anahtar in n:
                out.add(ad_)
        return out

    def kendi_banka_tahmini(r):
        """'UMH İŞBANKASI HESABINA' -> mizandaki İş Bankası vadesiz TL hesabı."""
        istenen = _banka_adi(r["aciklama"])
        if not istenen:
            return ""
        adaylar = [k for k, a in hesaplar if k in alt and k.startswith("102") and k != r["banka"]
                   and _banka_adi(a) & istenen]
        adaylar.sort(key=lambda k: (not k.startswith("102.01"), bool(re.search(r"VADELI|USD|EUR|GBP", norm(hesap_ad(k)))), k))
        return adaylar[0] if adaylar else ""

    def hesap_ipucu(r, kod):
        return any(_no_esit(a, b) for a in _nolar(r["aciklama"]) for b in hesap_nolari.get(kod, ()))

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
                if not (kendisi_mi(r) or kendisi_mi(s_) or set(r["kel"] + s_["kel"]) & _VIRMAN_KELIME
                        or hesap_ipucu(r, s_["banka"]) or hesap_ipucu(s_, r["banka"])):
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

    masraf_hesap = gecmis.masraf_hesabi(ad)
    if not masraf_hesap:
        masraf_hesap = next((k for k, a in hesaplar if k in alt and k.startswith(("770", "780", "653", "689"))
                             and "BANKA" in norm(a)), "")

    # ---------- karşı hesap
    cari_onek = ("120", "121", "126", "131", "136", "159", "195", "196", "300", "303", "320", "321",
                 "329", "331", "335", "336", "340", "400")
    kullanilan = {str(x.get("hesap", "")) for x in km.gecmis.get("satirlar", [])}
    mizan_adaylari = []
    for k, a in hesaplar:
        if k in alt and k.startswith(cari_onek):
            kel = [w for w in _kelimeler(a) if w not in _GENEL and w not in _HESAP_GENEL]
            if kel and (len(kel) >= 2 or len(kel[0]) >= 5):
                mizan_adaylari.append((k, kel))

    # aynı kişi/firma adı mizanda birden çok hesapta (ortak: 131, 331, 195, 500 …): her zaman KONTROL
    ad_hesaplari = defaultdict(set)
    for k, a in hesaplar:
        if k in alt:
            kk_ = frozenset(w for w in _kelimeler(a) if w not in _GENEL and w not in _HESAP_GENEL)
            if len(kk_) >= 2:
                ad_hesaplari[kk_].add(k)
    cok_hesapli = {kk_: v for kk_, v in ad_hesaplari.items() if len(v) >= 2}

    def cok_hesapli_ad(r, hesap):
        # seçilen hesap o adı taşımasa da ("ORTAK HAREKETLERİ") açıklamadaki kişinin birden çok hesabı varsa
        for kk_, v in cok_hesapli.items():
            if _hepsi_var(list(kk_), set(r["kel"]), r["kel"]):
                return sorted(v | {hesap})
        return None

    def vergi_hesabi(r):
        vb = _vergi_bilgisi(r["aciklama"])
        if vb is None:
            return None
        kod, vad, donem = vb
        if not kod:
            return ("", "vergi kodu okunamadı")
        if "GECICI" in vad or "GEC" in vad.split() or kod in ("0032", "0033"):
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

    def kredi_karti_hesabi(r):
        n = norm(r["aciklama"])
        maskeli = re.search(r"\d{4,6}\s*\*{2,}\s*\d{4}", r["aciklama"])
        if not ((re.search(r"\b(KREDI|K|KRE) KART", n) and ("BORC" in n or "ODEME" in n)) or maskeli):
            return None
        ebeveyn_kart = {k for k, a in hesaplar if "KREDI KART" in norm(a) and k not in alt}
        adaylar = [k for k, a in hesaplar if k in alt and k.startswith(("309", "300", "329", "336")) and
                   ("KART" in norm(a) or "CARD" in norm(a) or any(k.startswith(e + ".") for e in ebeveyn_kart))]
        if not adaylar:
            return ("", "kredi_karti", "eslesmedi", "Kredi kartı ödemesi — mizanda kredi kartı hesabı yok")
        son4 = re.findall(r"\*+\s*(\d{4})", r["aciklama"])
        tutan = [k for k in adaylar if son4 and son4[-1] in (hesap_ad(k) or "")]
        if tutan:
            return (tutan[0], "kredi_karti", None, "")
        if len(adaylar) == 1:
            return (adaylar[0], "kredi_karti", None, "Kredi kartı ödemesi")
        return (adaylar[0], "kredi_karti", "belirsiz", "Birden çok kredi kartı hesabı var: " + ", ".join(adaylar[:4]))

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
        # 2) aynı anda yapılan transferin komisyon / BSMV satırı
        if r.get("masraf") and masraf_hesap:
            return masraf_hesap, "masraf", None, ""
        # 3) geçmiş ve mizan cari adı — güçlüyse (2+ kelime, tek aday) genel kurallardan önce gelir:
        #    firma vergi ödemesini hep 770'e yazıyorsa öyle kalır
        aday = gecmis_mizan(r, masraf_mi)
        kk = kredi_karti_hesabi(r) if r["yon"] == "cikis" else None
        if kk is not None and not kk[0] and aday:
            # kart hesabı mizanda yok: geçmişteki hesap, ama kontrol edilsin (ortağın kartı olabilir)
            return aday[0], aday[1], "belirsiz", "Kredi kartı ödemesi — mizanda kredi kartı hesabı yok, geçmişteki hesap önerildi"
        if aday and aday[2] is None and not (kk is not None and kk[0] and kk[0] != aday[0]):
            return aday[:4]
        # 4) banka masraf / ücret satırı (bir kişiye/cariye giden ödeme değilse: "PERSONEL MASRAF ÖDEMESİ")
        if wset & _MASRAF_KELIME and r["yon"] == "cikis" and abs(r["tutar"]) < 5000 and not kendisi_mi(r) and masraf_hesap \
                and not (aday and aday[4] >= 2):
            return masraf_hesap, "masraf", None, ""
        # 5) kredi kartı borç ödemesi
        kk = kredi_karti_hesabi(r) if r["yon"] == "cikis" else None
        if kk is not None:
            return kk
        # 6) vergi / SGK
        v = vergi_hesabi(r) if r["yon"] == "cikis" else None
        if v is not None:
            kod, aciklama_ = v
            return (kod or ""), "vergi", (None if kod and "aday" not in aciklama_ else "vergi"), \
                f"Vergi ödemesi ({aciklama_}); gecikme zammı varsa ayrı satıra ayırın"
        s_ = sgk_hesabi(r) if r["yon"] == "cikis" else None
        if s_ is not None:
            return s_, "sgk", (None if s_ else "vergi"), "SGK ödemesi"
        # 7) zayıf / belirsiz geçmiş eşleşmesi (KONTROL'lü)
        if aday:
            return aday[:4]
        return "", "", "eslesmedi", ""

    def gecmis_mizan(r, masraf_mi):
        """Geçmiş fiş + mizan cari adı. -> (hesap, kaynak, kontrol|None, not) ya da None"""
        kel, wset = r["kel"], set(r["kel"])
        tutar = r["tutar"]
        gec = gecmis.adaylar(kel, r["yon"], r["banka"], tutar, r["tarih"])
        taban = {}
        for k, d in gec.items():
            # Logo açıklamayı kestiği için transferin komisyon satırı geçmişte transferle aynı
            # görünür ("ALP ARAYICI GARANTI" -> 770); masraf olmayan satırda masraf hesabı aday değil
            if k in kodlar and k != r["banka"] and not (k == masraf_hesap and not masraf_mi):
                taban[k] = d["taban"]
        mizan_s, bulanik_ = {}, set()
        for k, hk in mizan_adaylari:
            tam = _hepsi_var(hk, wset, kel)
            if tam or _hepsi_var(hk, wset, kel, bulanik=True):
                ms = len(hk) + 0.5
                # aynı ad hem 131 hem 331'de olabilir: giden ödeme borcu (3xx), gelen tahsilat
                # alacağı (1xx) kapatır; geçmişte kullanılmış hesap öne geçer
                ms += 0.3 if (r["yon"] == "cikis") == k.startswith(("3", "4")) else 0
                ms += 0.3 if k in kullanilan else 0
                mizan_s[k] = ms
                if not tam:
                    bulanik_.add(k)
        # mizanda tek kelimesi tutan (TURKCELL İLETİŞİM ...): ilk ayırt edici kelime yalnız o hesapta
        if not taban and not mizan_s:
            ilk = defaultdict(list)
            for k, hk in mizan_adaylari:
                if len(hk[0]) >= 5:
                    ilk[hk[0]].append(k)
            for w_, ks in ilk.items():
                if len(set(ks)) == 1 and _w_esler(w_, wset, kel):
                    mizan_s[ks[0]] = 1.5
        adaylar_ = set(taban) | set(mizan_s)
        if not adaylar_:
            return None
        temel = {k: max(taban.get(k, 0), mizan_s.get(k, 0)) + (0.5 if k in taban and k in mizan_s else 0)
                 for k in adaylar_}
        en_t = max(temel.values())
        # yarışanlar: açıklamayı en iyiye yakın tutanlar; aralarında son aylarda kullanılan ve tutarı benzeyen
        yaris = [k for k in adaylar_ if temel[k] >= max(en_t * 0.6, min(en_t, 2))]
        # aynı firma geçmişte başka hesaba yazılmışsa (MÜZİKÜSSÜ: mizanda/önce 120, sonra 349) o hesap da yarışır
        varlik = {w for k in yaris if k not in taban for w in dict(mizan_adaylari).get(k, []) if len(w) >= 5}
        varlik |= {w for w in wset if len(w) >= 5 and gecmis.nadir(w)
                   and any(w in gec[k]["kelimeler"] for k in yaris if k in gec)}
        for k in taban:
            if k not in yaris and varlik & gec[k]["kelimeler"]:
                yaris.append(k)
        varlik_catisma = bool(varlik) and len(yaris) > 1 and \
            len({k for k in yaris if (k in gec and varlik & gec[k]["kelimeler"]) or k not in taban}) > 1
        top_ag = sum(gec[k]["agirlik"] for k in yaris if k in gec) or 1
        puan = {k: 0.6 * temel[k] / en_t + 1.5 * (gec[k]["agirlik"] / top_ag if k in gec else 0)
                + 1.5 * (gec[k]["tutar_b"] if k in gec else 0) for k in yaris}
        sirali = sorted(yaris, key=lambda k: (-puan[k], -(gec[k]["adet"] if k in gec else 0)))
        en = sirali[0]
        sk = temel[en]
        kaynak = "gecmis" if en in taban else "mizan"
        kontrol, not_ = None, ""
        pay = (gec[en]["agirlik"] / top_ag) if en in gec else 0
        ayni_tutar = en in gec and gec[en].get("ayni_tutar")
        # açıklama geçmişte birçok hesaba dağılmışsa (her vergi ödemesinin bir de gecikme satırı
        # olması gibi) güvenilir değil: kesin kurallar (vergi, masraf) öne geçsin, değilse KONTROL
        dagink = len(sirali) > 1 and en in gec and pay < 0.6 and not ayni_tutar
        if (len(sirali) > 1 and puan[sirali[1]] >= puan[en] - 0.3) or dagink or varlik_catisma:
            kontrol = "belirsiz"
            not_ = "Bu açıklama geçmişte başka hesaplara da yazılmış: " + ", ".join(sirali[1:4])
        elif sk < 2:
            kontrol = "zayif"
            not_ = "Açıklamanın yalnız bir kelimesi geçmişle / hesap adıyla tuttu"
        elif en in bulanik_ and kaynak == "mizan":
            kontrol = "benzer"
            not_ = f"Hesap adı ({hesap_ad(en)}) açıklamayla bir harf farklı"
        coklu = cok_hesapli_ad(r, en)
        if coklu and kontrol is None:
            kontrol = "belirsiz"
            not_ = f"Bu ad mizanda {len(coklu)} hesapta var ({', '.join(coklu[:5])}) — doğru hesabı seçin"
        return en, kaynak, kontrol, not_, sk

    # masraf satırları: aynı anda yapılan transferin "- Komisyon" / "- Vergi (BSMV)" satırları
    for b, rs in bankalar.items():
        for r in rs:
            if r["yon"] != "cikis":
                continue
            it = norm(r.get("islem_tipi"))
            if re.search(r"KOMISYON|MASRAF|UCRET", it) and "TRANSFER" not in it:
                r["masraf"] = True        # Garanti "Etiket: Faiz / Komisyon"
                continue
            nr = norm(r["aciklama"])
            for s in rs:
                if s is r or s["tarih"] != r["tarih"] or abs(s["tutar"]) <= abs(r["tutar"]):
                    continue
                ns = norm(s["aciklama"])
                # Halkbank: masraf satırı transferle birebir aynı açıklamalı ("Şirket içi masrafları" 42.600 / 16,76)
                if nr == ns and abs(r["tutar"]) < 250 and abs(s["tutar"]) >= 20 * abs(r["tutar"]):
                    r["masraf"] = True
                    break
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
                yeni = fa[:m.start()] + f"{tarih[8:10]}{ay}{tarih[5:7]}{ay}{tarih[:4]}" + fa[m.end():]
                # geçmişteki açıklamada kalmış "KONTROL -" öneki / baştaki boşluk yeni fişe taşınmasın
                return re.sub(r"^\s*KONTROL\s*-\s*", "", yeni).strip()
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
        # kendi hesaplarımız arası transfer, ama karşı bankanın ekstresi bu tarihi kapsamıyor
        if kendisi_mi(r) and not r.get("masraf"):
            if hesap.startswith("102") and not kontrol:
                kontrol = "belirsiz"
                not_ = "Kendi hesaplarınız arası transfer — karşı bankanın ekstresi yüklenmedi; banka hesabını kontrol edin"
            elif (not hesap or hesap.startswith("198") or kontrol == "eslesmedi"):
                t_ = kendi_banka_tahmini(r)
                if t_:
                    hesap, kaynak, kontrol = t_, "virman", "belirsiz"
                    not_ = "Kendi hesaplarınız arası transfer (açıklamadaki banka adından) — kontrol edin"
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
