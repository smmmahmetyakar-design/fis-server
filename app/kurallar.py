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
from difflib import SequenceMatcher
from pathlib import Path
import openpyxl


# ----------------------------------------------------------------- dosyadan banka hesabı tanıma
_IBAN_RE = re.compile(r"TR\d{2}(?:[ .]?\d{4}){5}[ .]?\d{2}")
_HESAPNO_RE = re.compile(r"HESAP\s*NO\S*\s*[:\-]?\s*([0-9][0-9 .\-]{3,20}[0-9])")


def dosya_banka_anahtari(h: dict) -> str:
    """Bir belgenin (tablolar/ham_metin) içinden IBAN ya da hesap numarasını bulup
    bu fiziksel banka hesabını tekil tanımlayan bir anahtar döndürür (yoksa '').
    Aynı anahtar, öğrenilmiş banka_hesap_eslestirme.json içinde bir hesap koduna
    bağlanınca, sonraki aylarda bu dosya otomatik doğru hesaba yazılır."""
    parcalar = []
    for tablo in (h.get("tablolar") or [])[:3]:
        for row in tablo[:40]:
            for c in row:
                if c not in (None, ""):
                    parcalar.append(str(c))
    if h.get("ham_metin"):
        parcalar.append(h["ham_metin"][:4000])
    metin = " \n ".join(parcalar).upper()

    m = _IBAN_RE.search(metin)
    if m:
        return re.sub(r"[ .\-]", "", m.group(0))

    m = _HESAPNO_RE.search(metin)
    if m:
        rakam = re.sub(r"[^0-9]", "", m.group(1))
        if len(rakam) >= 4:
            return rakam

    return ""


# ----------------------------------------------------------------- kısa banka adı
# Mizan/kural excel'deki hesap adı genelde uzun ve dağınık olur (ör.
# "GARANTİ BANK 1047-6296738 TL"). Fiş Açıklama'da "BANKA-TARİH" biçimi için
# kısa, tanıdık bir banka adına indirgiyoruz. Yeni bir banka mizan'a eklendiğinde
# bu listeye eklemeye gerek yok — eşleşmeyen adlarda ilk anlamlı kelime kullanılır.
_BANKA_KISALTMALARI = [
    (re.compile(r"\bGARANTI\b"), "GARANTİ"),
    (re.compile(r"\bVAKIF"), "VAKIFBANK"),
    (re.compile(r"YAPI\s+(VE\s+)?KREDI|\bYKB\b"), "YKB"),
    (re.compile(r"\bIS\s*BANKASI\b|\bISBANK\b"), "İŞBANKASI"),
    (re.compile(r"\bAKBANK\b"), "AKBANK"),
    (re.compile(r"\bDENIZBANK\b"), "DENİZBANK"),
    (re.compile(r"\bZIRAAT\b"), "ZİRAAT"),
    (re.compile(r"\bHALKBANK\b"), "HALKBANK"),
    (re.compile(r"\bQNB\b|\bFINANSBANK\b"), "QNB FİNANSBANK"),
    (re.compile(r"\bTEB\b"), "TEB"),
    (re.compile(r"\bING\b"), "ING"),
    (re.compile(r"\bODEA"), "ODEABANK"),
    (re.compile(r"\bENPARA\b"), "ENPARA"),
    (re.compile(r"\bKUVEYT\b"), "KUVEYTTÜRK"),
    (re.compile(r"\bALBARAKA\b"), "ALBARAKA"),
]


def banka_kisa_adi(hesap_adi: str, hesap_kodu: str = "") -> str:
    """Mizan/kural hesap adından kısa, tanıdık banka adı çıkarır
    (ör. 'GARANTİ BANK 1047-6296738 TL' -> 'GARANTİ')."""
    ad_norm = norm(hesap_adi)
    for pat, kisa in _BANKA_KISALTMALARI:
        if pat.search(ad_norm):
            return kisa
    ilk_kelimeler = [w for w in ad_norm.split() if w not in STOP]
    if ilk_kelimeler:
        return ilk_kelimeler[0]
    return hesap_kodu or "BANKA"


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


def _yakin(a: str, b: str) -> bool:
    """İki kelime aynı ya da OCR'dan kaynaklı ufak bir yazım farkıyla (ör.
    'MUHASEBE'/'MUHASEBA', 'ODEMESI'/'ODENESI') neredeyse aynıysa True döner."""
    if a == b:
        return True
    if abs(len(a) - len(b)) > 2:
        return False
    return SequenceMatcher(None, a, b).ratio() >= 0.82


def _ortak_sayisi(gw: set, aw: set) -> int:
    """aw (öğrenilen anahtarın kelimeleri) içindeki kaç kelimenin gw (sorgu
    kelimeleri) içinde aynısı veya yakın bir eşi olduğunu sayar."""
    n = 0
    for a in aw:
        if a in gw or any(_yakin(a, g) for g in gw):
            n += 1
    return n


def kelimeler(s: str) -> set:
    """Eşleştirmede kullanılacak anlamlı kelimeler. Referans/evrak no, tutar,
    oran ve sicil gibi işlem bazında değişen sayısal token'lar (ör. 'H2604200942507',
    '26920050511927770340630', '5000') elenir — aksi halde her işlemde farklı olan
    bu kodlar yüzünden aynı türden hareketler (BSMV, ÜCRET, BRÜT FAİZ ... gibi) her
    seferinde eşleşmeyen kabul edilirdi."""
    out = set()
    for w in norm(s).split():
        if w in STOP or len(w) <= 2:
            continue
        rakam = sum(c.isdigit() for c in w)
        if rakam == len(w):
            continue  # salt sayı: tutar/referans/evrak/sicil no
        if len(w) >= 4 and rakam / len(w) >= 0.5:
            continue  # harf+rakam karışık referans kodu (ör. H2604200942507)
        out.add(w)
    return out


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


def _fatura_kalem_ogren(fis_path: Path, cari_onek: str, kdv_onek: str) -> dict:
    """fatura_gider_ogren ve fatura_gelir_ogren'in ortak mantığı — bkz. onların
    docstring'i. Tek fark cari_onek/kdv_onek: alışta '32'/'191' (Ticari Borçlar/
    İndirilecek KDV), satışta '12'/'391' (Ticari Alacaklar/Hesaplanan KDV)."""
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
        kodlar = {b[0] for b in blok}
        kdv_haric = {k for k in kodlar if not k.startswith(kdv_onek)}
        cari_kodlar = {k for k in kdv_haric if k.startswith(cari_onek)}
        diger_kodlar = kdv_haric - cari_kodlar
        if len(cari_kodlar) == 1 and len(diger_kodlar) == 1:
            ogrenme[next(iter(cari_kodlar))] = next(iter(diger_kodlar))

    for r in range(1, ws.max_row + 1):
        kod = ws.cell(r, 2).value
        if kod == "HESAP KODU":
            blok = []
            continue
        if isinstance(kod, str) and kod.strip().startswith("FİŞ TOPLAM"):
            isle_blok(blok)
            blok = []
            continue
        if kod and re.match(r"^\d{3}(\.\d+)*$", str(kod).strip()):
            blok.append((str(kod).strip(),))

    return ogrenme


def fatura_gider_ogren(fis_path: Path, cari_onek: str = "32", kdv_onek: str = "191") -> dict:
    """
    ALIŞ faturaları: hangi CARİ (tedarikçi) hesabının hangi GİDER/STOK
    hesabıyla eşleştiğini geçmiş fiş listesinden (Logo Tiger 'fiş listesi'
    export'u) öğrenir. Gelen e-Fatura listesinde ürün/hizmet açıklaması
    olmadığı için (sadece gönderici/tedarikçi adı var) gider hesabı açıklama
    kelimelerinden değil, CARİ HESABIN KENDİSİNDEN öğrenilir: aynı tedarikçi
    (aynı cari kod) geçmişte hangi gider hesabına işlenmişse, yeni
    faturalarda da varsayılan olarak o hesap önerilir.

    cari_onek varsayılan olarak '32' (Ticari Borçlar ANA GRUBU: 320 Satıcılar,
    321/322 Borç Senetleri, 329 Diğer Ticari Borçlar vb.) — sadece '320' değil,
    çünkü ofisler faktoring/kira gibi bazı tedarikçileri 329.x gibi farklı bir
    alt hesapta izleyebiliyor (gerçek AREL verisinde GARANTİ FAKTORİNG'in cari
    kodu 329.02.006 çıktı, eski '320' sabiti bunu hiç yakalamıyordu).

    Mantık: her fiş bloğunda KDV (191.x) hesabı hariç tutulduğunda geriye TEK
    bir cari (32x) ve TEK bir gider hesabı kalıyorsa (basit, tek kalemli
    fatura fişi), bu ikisini eşleştirip kaydeder. Birden çok gider kalemli
    (KDV hariç 2+ farklı gider hesabı olan) fişler yanlış öğrenme riskine
    karşı atlanır — bu tür çok kalemli faturalar otomatik öğrenilemez,
    Düzenle modunda elle düzeltilince (bkz. main.py /ogret-gider) öğrenilir.

    Döndürür: {cari_hesap_kodu: gider_hesap_kodu}
    """
    return _fatura_kalem_ogren(fis_path, cari_onek, kdv_onek)


def fatura_gelir_ogren(fis_path: Path, cari_onek: str = "12", kdv_onek: str = "391") -> dict:
    """
    SATIŞ faturaları: hangi CARİ (müşteri/alıcı) hesabının hangi GELİR
    hesabıyla eşleştiğini geçmiş fiş listesinden öğrenir — fatura_gider_ogren
    ile birebir aynı mantık, sadece yön ters: cari_onek '12' (Ticari
    Alacaklar ANA GRUBU: 120 Alıcılar, 121 Alacak Senetleri, 126/127 diğer
    ticari alacaklar), kdv_onek '391' (Hesaplanan KDV).

    Döndürür: {cari_hesap_kodu: gelir_hesap_kodu}
    """
    return _fatura_kalem_ogren(fis_path, cari_onek, kdv_onek)


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

    def __init__(self, mizan_path: Path, kural_path: Path, ogrenme_path: Path,
                 banka_eslestirme_path: Path = None, gider_eslestirme_path: Path = None,
                 gelir_eslestirme_path: Path = None):
        self.hesaplar = mizan_hesaplar(mizan_path)
        self.kural = kural_excel_oku(kural_path)
        self.ogrenme = ogrenme_oku(ogrenme_path)
        self.ogrenme_path = ogrenme_path
        self.banka_eslestirme_path = banka_eslestirme_path
        self.banka_eslestirme = ogrenme_oku(banka_eslestirme_path) if banka_eslestirme_path else {}
        # Fatura (ALIŞ): CARİ hesap kodu -> GİDER hesap kodu (geçmiş fiş
        # listesinden öğrenilir, bkz. fatura_gider_ogren — açıklama
        # kelimesinden değil, doğrudan cari hesabın kendisinden öğrenilir).
        self.gider_eslestirme_path = gider_eslestirme_path
        self.gider_eslestirme = ogrenme_oku(gider_eslestirme_path) if gider_eslestirme_path else {}
        # Fatura (SATIŞ): CARİ (müşteri) hesap kodu -> GELİR hesap kodu — aynı
        # mantık, ters yön (bkz. fatura_gelir_ogren).
        self.gelir_eslestirme_path = gelir_eslestirme_path
        self.gelir_eslestirme = ogrenme_oku(gelir_eslestirme_path) if gelir_eslestirme_path else {}
        self.kod_ad = {k: a for k, a in self.hesaplar}
        for h in self.kural["hesaplar"]:
            self.kod_ad.setdefault(h["kod"], h["ad"])

    def _kalem_hesabi(self, eslestirme: dict, cari_kod: str, ad: str) -> str:
        """gider_hesabi/gelir_hesabi'nin ortak mantığı: önce cari koduna göre
        öğrenilenlere bakar, yoksa Kural Dosyası'ndaki hesap adı/kullanım
        metniyle karşı taraf adı arasında kelime eşleşmesi dener."""
        if cari_kod and cari_kod in eslestirme:
            return eslestirme[cari_kod]
        if ad:
            gw = kelimeler(ad)
            best = None; bs = 0
            for h in self.kural["hesaplar"]:
                hedef = kelimeler(h["ad"] + " " + h.get("kullanim", ""))
                sc = len(gw & hedef)
                if sc > bs:
                    bs = sc; best = h["kod"]
            if bs >= 1:
                return best
        return ""

    def gider_hesabi(self, cari_kod: str, gonderici: str = "") -> str:
        """Bir cari (tedarikçi) hesap koduna öğrenilmiş varsayılan GİDER hesabı
        (ALIŞ faturaları için) — bkz. _kalem_hesabi."""
        return self._kalem_hesabi(self.gider_eslestirme, cari_kod, gonderici)

    def gelir_hesabi(self, cari_kod: str, alici: str = "") -> str:
        """Bir cari (müşteri/alıcı) hesap koduna öğrenilmiş varsayılan GELİR
        hesabı (SATIŞ faturaları için) — bkz. _kalem_hesabi."""
        return self._kalem_hesabi(self.gelir_eslestirme, cari_kod, alici)

    def banka_hesap_ogren(self, anahtar: str, kod: str):
        """IBAN/hesap no anahtarını bir hesap koduna bağlar ve kalıcı olarak kaydeder
        (sonraki işlemlerde bu dosya otomatik doğru hesaba düşer)."""
        if not anahtar or not kod or not self.banka_eslestirme_path:
            return
        self.banka_eslestirme[anahtar] = kod
        ogrenme_yaz(self.banka_eslestirme_path, self.banka_eslestirme)

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
                ortak = _ortak_sayisi(gw, aw)
                ort = ortak / len(aw)
                if ort >= 0.6 and ortak > best_og_sc:
                    best_og_sc = ortak; best_og = kod
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
