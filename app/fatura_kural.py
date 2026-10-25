"""
Fatura fişi üretim kuralları (alış / satış).

Fiş kontrol aracının (fis-kontrol-server) Luca fiş listelerinde bulduğu kayıt hatalarının
üretim tarafındaki karşılığı. İki katman:

GENEL KURALLAR — her firmada, ayar dosyası gerekmeden (mizan hesap adlarıyla):
  * RET / reddedilen / iptal edilen fatura fişe girmez.
  * KDV hesabının adında yazan oran faturanın KDV oranını tutmalı.
  * Alış tevkifatı: 360 hesabı tevkifat koduna (606 işgücü, 612 temizlik, 624 taşıma …)
    ve oranına (9/10, 2/10 …) göre hesap adından seçilir; tutmuyorsa KONTROL.
  * Satış tevkifatı: 391 tevkifat hesabının adındaki oran faturadaki oranı tutmalı.
  * Cari hesabın mizan adı fatura unvanıyla uyuşmuyorsa KONTROL.
  * İade: alıştan iade (satış listesinde) satıcı carisine ve asıl alış hesabına, satıştan
    iade (alış listesinde) alıcı carisine ve 610 satıştan iadeler hesabına yazılır.

FİRMA KURALLARI — isteğe bağlı, data/<firma>/fatura_kurallari.yaml (arayüzden düzenlenir):
  * kalem açıklaması → hesap (mal → 153 alt hesabı, hizmet → 740 alt hesabı,
    satış kategorisi → 600 alt hesabı); karışık fatura kalem kalem bölünür,
  * tevkifat oranı → satış (600) ve KDV (391) hesabı, tevkifatlı alışın KDV hesabı,
  * hiçbir kurala uymayan kalem için hesap ve uyarı (ör. yedek parça satışı),
  * iade hesapları, eşdeğer cari adları.
Firma dosyası yoksa yalnız genel kurallar çalışır; mevcut işleyiş (geçmiş fiş, öğrenme,
yapay zekâ) değişmez.
"""
import re
from pathlib import Path

import yaml

from app.kurallar import norm

DOSYA_ADI = "fatura_kurallari.yaml"
ORNEK_YOLU = Path(__file__).with_name("fatura_kurallari_ornek.yaml")

_TR = str.maketrans({"i": "İ", "ı": "I"})


def tr_ust(s) -> str:
    """Türkçe kurallara uygun büyük harf (kalem açıklaması kurallarla bunun üzerinde eşleşir)."""
    return str(s or "").translate(_TR).upper()


# Unvan karşılaştırmasında ayırt edici sayılmayan kelimeler (şirket türü, sektör ekleri)
YOKSAY = {
    "LTD", "STI", "LIMITED", "SIRKETI", "SIRKET", "AS", "A", "S", "ANONIM", "VE", "SAN", "SANAYI", "TIC",
    "TICARET", "IC", "DIS", "ITH", "IHR", "ITHALAT", "IHRACAT", "IHRACAAT", "PAZARLAMA", "INS", "INSAAT",
    "OTOMOTIV", "NAK", "NAKLIYAT", "TURZ", "TURIZM", "MAK", "MAKINE", "MAKINA", "ELK", "MEK", "DOG", "SANAYII",
    "TICARETI", "GRUP", "HIZMETLERI", "HIZMET", "LTDSTI", "KOLL", "KOMANDIT", "TEKSTIL", "MADENCILIK", "YAPI",
}


def ad_kelimeleri(s, ek_yoksay=None) -> list:
    """Unvandaki anlamlı kelimeler, sırasıyla."""
    yok = YOKSAY | set(ek_yoksay or ())
    return [k for k in norm(s).split() if len(k) > 1 and k not in yok and not k.isdigit()]


# PDF'ten unvan yerine okunabilen belge sözcükleri (bunlarla taraf karşılaştırılmaz)
_BELGE_KELIME = {"FATURA", "ARSIV", "EFATURA", "EARSIV", "SAYIN", "VKN", "TCKN", "TC", "VERGI", "DAIRESI",
                 "MUDURLUGU", "MUD", "ADRES", "TEL", "FAX", "WEB", "SITESI", "POSTA", "SENARYO", "OZELLESTIRME"}


def ayni_taraf_genel(a: str, b: str) -> bool | None:
    """Kural dosyası olmadan iki unvan karşılaştırması (PDF ↔ liste eşleştirmesi için).
    True: aynı taraf · False: farklı taraf · None: karşılaştırılamıyor."""
    ka = [k for k in ad_kelimeleri(a) if k not in _BELGE_KELIME and not re.fullmatch(r"[A-Z0-9]{3}20\d{11}", k)]
    kb = [k for k in ad_kelimeleri(b) if k not in _BELGE_KELIME and not re.fullmatch(r"[A-Z0-9]{3}20\d{11}", k)]
    if not ka or not kb:
        return None
    if ka[0] in kb or kb[0] in ka or "".join(ka[:2]) == kb[0] or "".join(kb[:2]) == ka[0]:
        return True
    return False


# KDV tevkifat kodu → hesap adında aranacak kelimeler (GİB kod listesi; norm() edilmiş)
TEVKIFAT_KOD_ANAHTAR = {
    "601": ("YAPIM",),
    "602": ("ETUT", "PLAN", "PROJE", "DANISMANLIK", "MUSAVIRLIK"),
    "603": ("BAKIM", "ONARIM"),
    "604": ("YEMEK",),
    "605": ("ORGANIZASYON",),
    "606": ("ISGUCU", "IS GUCU", "PERSONEL"),
    "607": ("GUVENLIK",),
    "608": ("YAPI DENETIM",),
    "609": ("FASON", "KONFEKSIYON"),
    "610": ("TURISTIK", "MAGAZA"),
    "611": ("SPOR",),
    "612": ("TEMIZLIK",),
    "613": ("CEVRE", "BAHCE"),
    "614": ("SERVIS",),
    "615": ("BASKI", "BASIM"),
    "616": ("DIGER HIZMET",),
    "617": ("HURDA",), "618": ("HURDA",), "619": ("HURDA",), "620": ("HURDA",), "621": ("HURDA",),
    "622": ("PAMUK", "TIFTIK", "YUN"),
    "623": ("AGAC", "ORMAN"),
    "624": ("TASIMA", "NAKLIYE", "NAKLIYAT"),
    "625": ("REKLAM",),
    "626": ("DIGER TESLIM",),
    "627": ("DEMIR", "CELIK"),
}
TEVKIFAT_ORANLARI = (0.2, 0.3, 0.4, 0.5, 0.7, 0.9, 1.0)
KDV_ORANLARI = (1, 8, 10, 18, 20)
_ESKI_ORAN = {20: 18, 10: 8}           # Temmuz 2023 öncesi açılmış hesap adları
_TEVK_DISI = ("GELIR", "STOPAJ", "DAMGA", "SGK", "SIGORTA", "KURUMLAR", "MUHTASAR")


def oran_metin(r) -> str:
    """0.9 → '9/10'"""
    if not r:
        return ""
    return "tam" if r >= 0.999 else f"{round(r * 10)}/10"


def oran_coz(s) -> float | None:
    """'9/10', '%90', '0,9', '90' → 0.9"""
    t = str(s or "").strip().replace("%", "").replace(" ", "").replace(",", ".")
    if not t:
        return None
    m = re.fullmatch(r"(\d+)/(\d+)", t)
    try:
        r = int(m.group(1)) / int(m.group(2)) if m else float(t)
    except (ValueError, ZeroDivisionError):
        return None
    if r > 1:
        r = r / 100
    return min(TEVKIFAT_ORANLARI, key=lambda x: abs(x - r)) if 0 < r <= 1 else None


def tevkifat_orani(tevk: float, kdv: float) -> float | None:
    """Tevkifat tutarı / KDV → standart orana yuvarlanmış (0.2 … 1.0)."""
    if not tevk or not kdv or kdv <= 0:
        return None
    r = tevk / kdv
    en = min(TEVKIFAT_ORANLARI, key=lambda x: abs(x - r))
    return en if abs(en - r) <= 0.02 else None


def hesap_tevkifat_orani(ad: str) -> float | None:
    """Hesap adında yazan tevkifat oranı: '9/10', '2/10', '%90' (KDV oranlarıyla karışan
    %1/%8/%10/%18/%20 tevkifat oranı sayılmaz — '%20 TEVKİFATLI KDV' belirsizdir)."""
    n = str(ad or "").upper()
    m = re.search(r"(\d{1,2})\s*/\s*10\b", n)
    if m and 0 < int(m.group(1)) <= 10:
        return int(m.group(1)) / 10
    m = re.search(r"%\s*(30|40|50|70|90|100)\b", n)
    if m:
        return int(m.group(1)) / 100
    return None


def hesap_kdv_orani(ad: str) -> int | None:
    """KDV hesabının adında yazan KDV oranı (yalnız 1/8/10/18/20)."""
    n = norm(ad)
    if "SORUMLU" in n or "TEVK" in n:
        return None
    m = re.search(r"%\s*(\d{1,2})(?!\d)", str(ad or ""))
    if m and int(m.group(1)) in KDV_ORANLARI:
        return int(m.group(1))
    for m in re.finditer(r"(?<!\d)(\d{1,2})(?!\d)", str(ad or "")):
        o = int(m.group(1))
        if o in KDV_ORANLARI:
            return o
    return None


def kdv_orani_uyar(ad: str, oran: int) -> bool:
    """Hesap adı oran yazmıyorsa ya da faturanın oranını (eski karşılığıyla) yazıyorsa True."""
    o = hesap_kdv_orani(ad)
    return o is None or o == oran or _ESKI_ORAN.get(oran) == o


def anahtar_okunur(rx: str, n: int = 3) -> str:
    """Düzenli ifadeyi önizleme notu için sadeleştirir: 'NAKL[İI]YE|TAŞIMA' → 'NAKLİYE, TAŞIMA'."""
    s = re.sub(r"\[([^\]])[^\]]*\]", r"\1", str(rx or ""))          # [İI] → İ
    s = re.sub(r"\\[bBdsw]|[\^$()?*+{}\\]", "", s)
    parcalar = [p.strip() for p in s.split("|") if p.strip()]
    tekil = list(dict.fromkeys(parcalar))
    return ", ".join(tekil[:n]) + (" …" if len(tekil) > n else "")


def _alt_hesaplar(hesaplar) -> set:
    ust = set()
    for k, _ in hesaplar:
        p = k.split(".")
        for i in range(1, len(p)):
            ust.add(".".join(p[:i]))
    return {k for k, _ in hesaplar if k not in ust}


def _bos(v) -> bool:
    return v is None or str(v).strip().lower() in ("", "null", "~", "yok", "none", "-")


def _evet(v) -> bool:
    return str(v or "").strip().lower() in ("true", "evet", "1", "yes", "acik", "açık")


def _tur(hesap: str) -> str:
    """Alış kalemi için hesap türü: stok (15x) ya da gider/hizmet."""
    return "mal" if str(hesap).startswith("15") else "hizmet"


class Kurallar:
    """Bir firma için kural seti. veri: fatura_kurallari.yaml içeriği (yoksa {})."""

    def __init__(self, veri: dict | None, hesaplar):
        self.veri = veri or {}
        self.var = bool(self.veri)
        self.kod_ad = {k: a for k, a in hesaplar}
        self.alt = _alt_hesaplar(hesaplar)
        self.hatalar: list = []          # kural dosyasındaki sorunlar (arayüzde gösterilir)
        self.kalem = {"alis": [], "satis": []}
        self.diger_kalem = {"alis": None, "satis": None}
        self.tevkifat_anahtar = {k: list(v) for k, v in TEVKIFAT_KOD_ANAHTAR.items()}
        self.alis_tevkifat_kdv: dict = {}     # oran (None: hepsi) -> hesap
        self.satis_tevkifat: dict = {}        # oran -> {"satis": hesap|None, "kdv": hesap|None}
        self.iade = {"alis_iadesi": "asil", "satis_iadesi": ""}
        self.cari_esdeger: list = []
        self.cari_yoksay: set = set()
        self._oku()

    # ------------------------------------------------------------------ dosya
    def _hesap(self, v, yer: str) -> str:
        if _bos(v):
            return ""
        h = str(v).strip()
        if self.kod_ad and h not in self.alt:
            self.hatalar.append(f"{yer}: {h} mizanda {'üst hesap (kayıt atılamaz)' if h in self.kod_ad else 'yok'} — kullanılmadı")
            return ""
        return h

    def _oku(self):
        v = self.veri
        for yon in ("alis", "satis"):
            b = v.get(yon) or {}
            if not isinstance(b, dict):
                self.hatalar.append(f"{yon}: bölüm yapısı hatalı")
                continue
            for i, k in enumerate(b.get("kalem") or [], 1):
                if not isinstance(k, dict) or _bos(k.get("anahtar")):
                    self.hatalar.append(f"{yon}.kalem {i}: 'anahtar' yok")
                    continue
                try:
                    rx = re.compile(str(k["anahtar"]), re.IGNORECASE)
                except re.error as e:
                    self.hatalar.append(f"{yon}.kalem {i}: düzenli ifade hatalı ({e})")
                    continue
                hesap = self._hesap(k.get("hesap"), f"{yon}.kalem {i}")
                tur = str(k.get("tur") or "").strip().lower()
                if tur not in ("", "mal", "hizmet"):
                    self.hatalar.append(f"{yon}.kalem {i}: tur 'mal' ya da 'hizmet' olmalı")
                    tur = ""
                if not hesap and not tur:
                    if _bos(k.get("hesap")):
                        self.hatalar.append(f"{yon}.kalem {i}: 'hesap' ya da 'tur' gerekli")
                    continue
                self.kalem[yon].append({"rx": rx, "anahtar": str(k["anahtar"]), "hesap": hesap,
                                        "tur": tur or (_tur(hesap) if yon == "alis" else ""),
                                        "not": str(k.get("not") or "").strip()})
            d = b.get("diger_kalem")
            if isinstance(d, dict):
                h = self._hesap(d.get("hesap"), f"{yon}.diger_kalem")
                self.diger_kalem[yon] = {"hesap": h, "uyari": str(d.get("uyari") or "").strip(),
                                         "tur": _tur(h) if (h and yon == "alis") else ""}
            if yon == "alis":
                for kod, kelime in (b.get("tevkifat_360") or {}).items():
                    kod = str(kod).strip()
                    for w in re.split(r"[|,]", str(kelime or "")):
                        if norm(w):
                            self.tevkifat_anahtar.setdefault(kod, [])
                            if norm(w) not in self.tevkifat_anahtar[kod]:
                                self.tevkifat_anahtar[kod].insert(0, norm(w))
                tk = b.get("tevkifat_kdv")
                if isinstance(tk, dict):
                    for o, h in tk.items():
                        r = oran_coz(o)
                        h = self._hesap(h, f"alis.tevkifat_kdv {o}")
                        if r and h:
                            self.alis_tevkifat_kdv[r] = h
                elif not _bos(tk):
                    h = self._hesap(tk, "alis.tevkifat_kdv")
                    if h:
                        self.alis_tevkifat_kdv[None] = h
            else:
                for o, t in (b.get("tevkifat") or {}).items():
                    r = oran_coz(o)
                    if not r:
                        self.hatalar.append(f"satis.tevkifat: '{o}' oranı okunamadı (ör. \"2/10\")")
                        continue
                    t = t if isinstance(t, dict) else {"satis": t}
                    self.satis_tevkifat[r] = {"satis": self._hesap(t.get("satis"), f"satis.tevkifat {o}"),
                                              "kdv": self._hesap(t.get("kdv"), f"satis.tevkifat {o}")}
        iade = v.get("iade") or {}
        if isinstance(iade, dict):
            a = str(iade.get("alis_iadesi") or "asil").strip()
            self.iade["alis_iadesi"] = "asil" if a.lower() in ("asil", "asıl", "") else self._hesap(a, "iade.alis_iadesi") or "asil"
            self.iade["satis_iadesi"] = self._hesap(iade.get("satis_iadesi"), "iade.satis_iadesi")
        for g in v.get("cari_esdeger") or []:
            if isinstance(g, list) and len(g) > 1:
                self.cari_esdeger.append([norm(x) for x in g if norm(x)])
        self.cari_yoksay = {norm(x) for x in (v.get("cari_ad_yoksay") or []) if norm(x)}

    # ------------------------------------------------------------------ kalem
    def kalem_kurali(self, yon: str, aciklama: str) -> dict | None:
        """Kalem açıklaması için kural: {hesap, tur, anahtar, not, diger}."""
        if not self.var:
            return None
        metin = tr_ust(aciklama)
        if not metin.strip() or re.fullmatch(r"[^:]{0,25}:\s*", metin):
            return None
        for k in self.kalem.get(yon, []):
            if k["rx"].search(metin):
                return {**{x: k[x] for x in ("hesap", "tur", "anahtar", "not")}, "diger": False}
        d = self.diger_kalem.get(yon)
        if d and (d["hesap"] or d["uyari"]):
            return {"hesap": d["hesap"], "tur": d["tur"], "anahtar": "", "not": d["uyari"], "diger": True}
        return None

    # ------------------------------------------------------------------ cari
    def ayni_taraf(self, a: str, b: str) -> bool | None:
        """İki unvan aynı tarafı mı gösteriyor? Birinin ilk ayırt edici kelimesi diğerinde
        geçmeli (ya da eşdeğer tanımlı olmalı). Karşılaştırılamıyorsa None."""
        ka, kb = ad_kelimeleri(a, self.cari_yoksay), ad_kelimeleri(b, self.cari_yoksay)
        if not ka or not kb:
            return None
        if ka[0] in kb or kb[0] in ka:
            return True
        # birleşik yazım: "EN AS" / "ENAS", "MAY365" / "MAY 365"
        if "".join(ka[:2]) == kb[0] or "".join(kb[:2]) == ka[0]:
            return True
        na, nb = norm(a), norm(b)
        for grup in self.cari_esdeger:
            if any(g in na for g in grup) and any(g in nb for g in grup):
                return True
        return False

    # ------------------------------------------------------------------ tevkifat
    def _kod_kelimeleri(self, kod: str) -> list:
        return self.tevkifat_anahtar.get(str(kod or "").strip(), [])

    def _diger_kod_kelimesi(self, ad_n: str, kod: str) -> bool:
        kendi = set(self._kod_kelimeleri(kod))
        for k, kl in self.tevkifat_anahtar.items():
            if k == kod:
                continue
            if any(w in ad_n and w not in kendi for w in kl):
                return True
        return False

    def tevkifat_360(self, mevcut: str, kod: str, oran: float | None):
        """Alış tevkifat (360) hesabı tevkifat koduna/oranına uyuyor mu?
        Döner: (hesap, durum) — durum: '' uygun · 'kural' kurala göre değiştirildi ·
        'uyumsuz' uymuyor ve mizanda uygun hesap yok."""
        def bilgi(h):
            ad_n = norm(self.kod_ad.get(h, ""))
            kw = self._kod_kelimeleri(kod)
            return {"kw": bool(kod) and any(w in ad_n for w in kw),
                    "diger": bool(kod) and self._diger_kod_kelimesi(ad_n, kod),
                    "oran": hesap_tevkifat_orani(self.kod_ad.get(h, ""))}

        def uyumsuz(b):
            return bool((b["oran"] and oran and abs(b["oran"] - oran) > 0.01) or (b["diger"] and not b["kw"]))

        if mevcut:
            b = bilgi(mevcut)
            if not uyumsuz(b):
                return mevcut, ""
        adaylar = []
        for h in self.alt:
            if not h.startswith("360"):
                continue
            ad_n = norm(self.kod_ad.get(h, ""))
            if any(x in ad_n for x in _TEVK_DISI) or not any(x in ad_n for x in ("TEVK", "SORUMLU", "KDV")):
                continue
            b = bilgi(h)
            if uyumsuz(b):
                continue
            puan = (2 if b["kw"] else 0) + (1 if (b["oran"] and oran and abs(b["oran"] - oran) <= 0.01) else 0)
            if puan:
                adaylar.append((puan, h))
        if adaylar:
            return max(adaylar)[1], "kural"
        return mevcut, ("uyumsuz" if mevcut else "")

    def alis_tevkifat_kdv_hesabi(self, oran: float | None) -> str:
        return self.alis_tevkifat_kdv.get(oran) or self.alis_tevkifat_kdv.get(None) or ""

    def satis_tevkifat_ayari(self, oran: float | None) -> dict:
        return self.satis_tevkifat.get(oran) or {}

    def baska_orana_ayrilmis(self, hesap: str, oran: float | None, alan: str) -> bool:
        """Hesap kural dosyasında başka bir tevkifat oranına ayrılmış mı (alan: satis|kdv)?"""
        return any(t.get(alan) == hesap and r != oran for r, t in self.satis_tevkifat.items())

    def tevkifata_ayrilmis(self, hesap: str, alan: str) -> bool:
        return any(t.get(alan) == hesap for t in self.satis_tevkifat.values())

    def satis_391_uyar(self, hesap: str, oran: float | None) -> bool:
        """391 tevkifat hesabının adındaki oran (yazıyorsa) faturanın oranını tutuyor mu?"""
        r = hesap_tevkifat_orani(self.kod_ad.get(hesap, ""))
        return not (r and oran and abs(r - oran) > 0.01)

    def satis_391_bul(self, oran: float | None) -> str:
        """Adında bu tevkifat oranı yazan 391 alt hesabı."""
        for h in sorted(self.alt):
            if h.startswith("391") and "TEVK" in norm(self.kod_ad.get(h, "")):
                r = hesap_tevkifat_orani(self.kod_ad.get(h, ""))
                if r and oran and abs(r - oran) <= 0.01:
                    return h
        return ""

    # ------------------------------------------------------------------ iade
    def satis_iadesi_hesabi(self) -> str:
        if self.iade.get("satis_iadesi"):
            return self.iade["satis_iadesi"]
        adaylar = sorted(h for h in self.alt if h.startswith("610"))
        tercih = [h for h in adaylar if "IADE" in norm(self.kod_ad.get(h, ""))]
        return (tercih or adaylar or [""])[0]

    def satis_iadesi_kdv(self) -> str:
        """Satıştan iade KDV'si için 191 alt hesabı (adında İADE geçen), yoksa ''."""
        for h in sorted(self.alt):
            ad_n = norm(self.kod_ad.get(h, ""))
            if h.startswith("191") and "IADE" in ad_n and "SAT" in ad_n:
                return h
        return ""


# ---------------------------------------------------------------------- dosya işlemleri
def yukle_metin(metin: str) -> dict:
    """YAML metnini okur (tüm değerler metin olarak: hesap kodları sayıya dönmesin)."""
    if not str(metin or "").strip():
        return {}
    veri = yaml.load(metin, Loader=yaml.BaseLoader)
    if veri is None:
        return {}
    if not isinstance(veri, dict):
        raise ValueError("Kural dosyası 'alis:', 'satis:' gibi bölümlerden oluşmalı")
    return veri


def firma_kurallari(d: Path, hesaplar) -> Kurallar:
    """Firma klasöründeki kural dosyasıyla (yoksa yalnız genel kurallar) kural seti."""
    p = Path(d) / DOSYA_ADI
    veri, hata = {}, ""
    if p.exists():
        try:
            veri = yukle_metin(p.read_text(encoding="utf-8"))
        except Exception as e:
            hata = f"Fatura kuralları dosyası okunamadı, yalnız genel kurallar uygulandı: {e}"
    k = Kurallar(veri, hesaplar)
    if hata:
        k.hatalar.insert(0, hata)
    return k


def dogrula(metin: str, hesaplar) -> tuple:
    """Kaydetmeden önce: (Kurallar | None, hatalar)."""
    try:
        veri = yukle_metin(metin)
    except Exception as e:
        return None, [f"YAML okunamadı: {e}"]
    k = Kurallar(veri, hesaplar)
    return k, list(k.hatalar)


def ornek_metin() -> str:
    try:
        return ORNEK_YOLU.read_text(encoding="utf-8")
    except Exception:
        return ""
