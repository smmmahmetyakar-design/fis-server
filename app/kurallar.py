"""
Kural motoru — firma bazlı hesap eşleştirme.
İki kaynak:
  1) Firma kural Excel'i (SEZGIN_OZBAY gibi): "Hesap Kodu Eşleştirme" + "Talimatlar"
  2) Öğrenilen JSON (data/<firma>/<tip>_ogrenme.json): anahtar kelime -> hesap kodu
  3) Gerçek fiş geçmişi (Logo Tiger export) -> fis_listesi_ogren() ile toplu öğrenme

Belge işleme sırasında bir açıklama/satır için hesap kodu ararken:
  önce öğrenilen kurallar, sonra Excel kuralları, sonra anahtar kelime tahmini.
"""
import json, re, unicodedata
from pathlib import Path
import openpyxl


def norm(s: str) -> str:
    """Türkçe duyarsız normalize: büyük harf, aksan yok, sadece harf/rakam/boşluk."""
    if s is None:
        return ""
    s = str(s)
    repl = {"İ": "I", "I": "I", "ı": "I", "Ş": "S", "ş": "S", "Ğ": "G", "ğ": "G",
            "Ü": "U", "ü": "U", "Ö": "O", "ö": "O", "Ç": "C", "ç": "C"}
    for a, b in repl.items():
        s = s.replace(a, b)
    s = s.upper()
    s = re.sub(r"[^A-Z0-9 ]", " ", s)
    return re.sub(r"\s+", " ", s).strip()


STOP = {"ANONIM", "SIRKETI", "LIMITED", "LTD", "STI", "SAN", "TIC", "VE", "A", "S",
        "AS", "TICARET", "SANAYI", "MItedh", "HIZM", "HIZMETLERI", "MALI", "STI."}


def kelimeler(s: str) -> set:
    return {w for w in norm(s).split() if w not in STOP and len(w) > 2}


# ----------------------------------------------------------------- mizan
def mizan_hesaplar(mizan_path: Path):
    """Mizandaki tüm hesapları [(kod, ad)] döndürür. Başlık satırını otomatik bulur."""
    if not mizan_path.exists():
        return []
    try:
        wb = openpyxl.load_workbook(mizan_path, data_only=True)
    except Exception:
        return []
    ws = wb[wb.sheetnames[0]]
    out = []
    for r in range(1, ws.max_row + 1):
        k = ws.cell(r, 1).value
        a = ws.cell(r, 2).value
        if k is None or a is None:
            continue
        k = str(k).strip(); a = str(a).strip()
        if re.match(r"^\d{3}(\.\d+)*$", k):
            out.append((k, a))
    return out


# ----------------------------------------------------------------- kural excel
def kural_excel_oku(kural_path: Path):
    """
    Firma kural Excel'inden hesap eşleştirme kurallarını okur.
    Dosya bozuk/geçersizse artık çökmez — boş kural + "uyari" döner.
    """
    sonuc = {"hesaplar": [], "anahtar_kod": {}, "talimatlar": [], "uyari": ""}
    if not kural_path.exists():
        return sonuc

    try:
        wb = openpyxl.load_workbook(kural_path, data_only=True)
    except Exception as e:
        sonuc["uyari"] = f"Kural dosyası okunamadı (bozuk/desteklenmeyen format): {e}"
        return sonuc

    for sn in wb.sheetnames:
        if "eslest" in norm(sn).lower() or "hesap kodu" in norm(sn).lower():
            ws = wb[sn]
            for r in range(1, ws.max_row + 1):
                kod = ws.cell(r, 1).value
                ad = ws.cell(r, 2).value
                kull = ws.cell(r, 3).value
                if kod and re.match(r"^\d{3}(\.\d+)*$", str(kod).strip()):
                    sonuc["hesaplar"].append({
                        "kod": str(kod).strip(),
                        "ad": str(ad).strip() if ad else "",
                        "kullanim": str(kull).strip() if kull else "",
                    })

    for sn in wb.sheetnames:
        if "talimat" in norm(sn).lower():
            ws = wb[sn]
            for r in range(1, ws.max_row + 1):
                a = ws.cell(r, 1).value
                b = ws.cell(r, 2).value
                if a or b:
                    sonuc["talimatlar"].append({
                        "konu": str(a).strip() if a else "",
                        "aciklama": str(b).strip() if b else "",
                    })

    return sonuc


# ----------------------------------------------------------------- fiş listesinden öğrenme
def fis_listesi_ogren(fis_path: Path, banka_hesap_kodu: str) -> dict:
    """
    Firmanın gerçek muhasebe fiş geçmişinden (Logo Tiger 'fiş listesi' export'u,
    HESAP KODU | HESAP ADI | AÇIKLAMA | BORÇ | ALACAK sütunlu) banka hareketleri
    için karşı hesap kodunu öğrenir.

    Mantık: her fiş bloğunda hedef banka hesabı (banka_hesap_kodu) ile birlikte
    TEK bir farklı hesap kodu varsa (basit, tek karşı hesaplı fiş), o eşleşmeyi
    normalize edilmiş açıklama -> karşı hesap kodu olarak kaydeder. Karmaşık
    (çok karşı hesaplı) fişler yanlış öğrenme riskine karşı atlanır.

    Döndürür: {norm(aciklama): karsi_hesap_kodu}
    """
    ogrenme = {}
    if not fis_path.exists():
        return ogrenme
    try:
        wb = openpyxl.load_workbook(fis_path, data_only=True)
    except Exception:
        return ogrenme
    ws = wb[wb.sheetnames[0]]

    blok = []

    def isle_blok(blok):
        hedef = [b for b in blok if b[0] == banka_hesap_kodu]
        if not hedef:
            return
        karsi_kodlar = {b[0] for b in blok if b[0] != banka_hesap_kodu}
        if len(karsi_kodlar) != 1:
            return
        kkod = karsi_kodlar.pop()
        for kod, aciklama in hedef:
            anahtar = norm(aciklama)
            if anahtar:
                ogrenme[anahtar] = kkod

    for r in range(1, ws.max_row + 1):
        kod = ws.cell(r, 2).value
        aciklama = ws.cell(r, 7).value
        if kod == "HESAP KODU":
            blok = []
            continue
        if isinstance(kod, str) and kod.strip().startswith("FİŞ TOPLAM"):
            isle_blok(blok)
            blok = []
            continue
        if kod and re.match(r"^\d{3}(\.\d+)*$", str(kod).strip()):
            blok.append((str(kod).strip(), aciklama))

    return ogrenme


# ----------------------------------------------------------------- öğrenme JSON
def ogrenme_oku(path: Path) -> dict:
    if path.exists():
        try:
            return json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def ogrenme_yaz(path: Path, data: dict):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


# ----------------------------------------------------------------- eşleştirme
class KuralMotoru:
    """Bir firma + belge tipi (banka/fatura/cek) için hesap eşleştirme yapar."""

    def __init__(self, mizan_path: Path, kural_path: Path, ogrenme_path: Path):
        self.hesaplar = mizan_hesaplar(mizan_path)
        self.kural = kural_excel_oku(kural_path)
        self.ogrenme = ogrenme_oku(ogrenme_path)
        self.ogrenme_path = ogrenme_path
        self.kod_ad = {k: a for k, a in self.hesaplar}
        for h in self.kural["hesaplar"]:
            self.kod_ad.setdefault(h["kod"], h["ad"])

    def hesap_adi(self, kod: str) -> str:
        return self.kod_ad.get(kod, "")

    def eslestir(self, aciklama: str):
        nq = norm(aciklama)
        if not nq:
            return "", ""

        gw = kelimeler(aciklama)
        best_og = None; best_og_sc = 0
        for anahtar, kod in self.ogrenme.items():
            if not anahtar:
                continue
            if anahtar in nq:
                return kod, "ogrenme"
            aw = kelimeler(anahtar)
            if aw:
                ort = len(gw & aw) / len(aw)
                if ort >= 0.6 and len(gw & aw) > best_og_sc:
                    best_og_sc = len(gw & aw); best_og = kod
        if best_og:
            return best_og, "ogrenme"

        best = None; bs = 0
        for h in self.kural["hesaplar"]:
            hedef = kelimeler(h["ad"] + " " + h.get("kullanim", ""))
            sc = len(gw & hedef)
            if sc > bs:
                bs = sc; best = h["kod"]
        if bs >= 1:
            return best, "excel"

        best = None; bs = 0
        for kod, ad in self.hesaplar:
            sc = len(gw & kelimeler(ad))
            if sc > bs:
                bs = sc; best = kod
        if bs >= 2:
            return best, "tahmin"

        return "", ""

    def ogret(self, aciklama: str, kod: str):
        anahtar = norm(aciklama)
        if not anahtar or not kod:
            return
        self.ogrenme[anahtar] = kod
        ogrenme_yaz(self.ogrenme_path, self.ogrenme)
