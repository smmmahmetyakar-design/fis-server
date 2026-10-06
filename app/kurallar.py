"""
Kural motoru — firma bazlı hesap eşleştirme.
İki kaynak:
  1) Firma kural Excel'i (SEZGIN_OZBAY gibi): "Hesap Kodu Eşleştirme" + "Talimatlar"
  2) Öğrenilen JSON (data/<firma>/<tip>_ogrenme.json): anahtar kelime -> hesap kodu

Belge işleme sırasında bir açıklama/satır için hesap kodu ararken:
  banka/çek: öğrenilen eşleşmeler -> geçmiş fişler -> Excel kuralları -> tahmin.
  fatura: cari/gider/gelir/KDV için doğrudan geçmiş fişler referans alınır;
           geçmişte kayıt yoksa mizan varsayılanları kullanılır.
"""
import json, re, unicodedata
from datetime import datetime
from pathlib import Path
import openpyxl

# Bellek cache: {kaynak_str: (mtime, sonuc_dict)}
# Fiş listesi Excel'i her istekte yeniden diskten JSON okumasın diye RAM'de tutar.
_bellek_cache = {}


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
_mizan_bellek_cache = {}

def mizan_hesaplar(mizan_path: Path):
    """Mizandaki tüm hesapları [(kod, ad)] döndürür. Başlık satırını otomatik bulur.
    Bellek cache: dosya mtime değişmezse Excel yeniden okunmaz (~1 ms yerine ~1500 ms)."""
    if not mizan_path.exists():
        return []
    kaynak_str = str(mizan_path)
    mtime = mizan_path.stat().st_mtime
    if kaynak_str in _mizan_bellek_cache:
        cached_mtime, cached_out = _mizan_bellek_cache[kaynak_str]
        if cached_mtime == mtime:
            return cached_out
    wb = openpyxl.load_workbook(mizan_path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    out = []
    for r in range(1, ws.max_row + 1):
        k = ws.cell(r, 1).value
        a = ws.cell(r, 2).value
        if k is None or a is None:
            continue
        k = str(k).strip(); a = str(a).strip()
        # hesap kodu deseni: rakam ve nokta
        if re.match(r"^\d{3}(\.\d+)*$", k):
            out.append((k, a))
    _mizan_bellek_cache[kaynak_str] = (mtime, out)
    return out


# ----------------------------------------------------------------- kural excel
def kural_excel_oku(kural_path: Path | None):
    """
    Firma kural Excel'inden hesap eşleştirme kurallarını okur.
    "Hesap Kodu Eşleştirme" sayfası: Kod | Ad | İşlem Türü/Kullanım
    Döndürür: {
      "hesaplar": [{kod, ad, kullanim}],
      "anahtar_kod": {anahtar_kelime: kod},   # kullanım açıklamasından türetilir
      "talimatlar": [ham metin satırları],
    }
    """
    sonuc = {"hesaplar": [], "anahtar_kod": {}, "talimatlar": []}
    if not kural_path or not kural_path.exists():
        return sonuc
    wb = openpyxl.load_workbook(kural_path, data_only=True)

    # hesap kodu eşleştirme sayfası
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

    # talimatlar sayfası (ham metin — ileride kural motoru için referans)
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


# ----------------------------------------------------------------- geçmiş fişler

def _tr_num(s):
    """Türkçe muhasebe tutarını float'a çevirir."""
    if s is None:
        return 0.0
    s = str(s).strip().replace(".", "").replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0


def gecmis_fisler_oku_pdf(pdf_path: Path):
    """
    Geçmiş fiş listesini PDF'den okur.
    Beklenen rapor düzeni Logo/Luca tarzı fiş listesidir:
      Fiş No / Tarih / Belge Düzenleme Nedeni
      HESAP KODU ... BORÇ ALACAK
      hesap kodu ... borç alacak
    Metin tabanlı PDF tercih edilir; taranmış PDF için OCR fallback kullanılır.
    """
    sonuc = {"satirlar": [], "son_fis_no": 0, "eslesmeler": {}}
    if not pdf_path or not pdf_path.exists():
        return sonuc

    try:
        import pdfplumber
    except Exception:
        return sonuc

    lines = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                txt = page.extract_text() or ""
                if txt.strip():
                    lines.extend(txt.splitlines())
    except Exception:
        return sonuc

    # Taranmış PDF ise mevcut belge okuyucusunun OCR yolunu kullan.
    if not lines:
        try:
            from app.belge_oku import pdf_ocr
            raw = pdf_ocr(pdf_path).get("ham_metin", "")
            lines = raw.splitlines()
        except Exception:
            return sonuc

    current_fis = ""
    current_date = ""
    current_aciklama = ""
    max_fis = 0

    date_re = re.compile(r"\b\d{2}/\d{2}/\d{4}\b")
    fis_re = re.compile(r"Fiş\s*No\s*:\s*(\S+)", re.I)
    neden_re = re.compile(r"Belge Düzenleme Nedeni\s*:\s*(.*)$", re.I)
    row_re = re.compile(
        r"^(\d{3}(?:\.\d+)+)\s+(.+?)\s+(\d{2}/\d{2}/\d{4})\s+(.*?)\s+"
        r"([\d.]+,\d{2})\s+([\d.]+,\d{2})\s*$"
    )

    # Bazı PDF'lerde hesap adı/açıklama satır sonundan taşar. Hesap koduyla başlayan
    # satırları biriktirip tutarlarla biten mantıksal satıra dönüştürüyoruz.
    logical = []
    buf = ""
    for raw in lines:
        line = " ".join(str(raw).split()).strip()
        if not line:
            continue
        is_new_account = bool(re.match(r"^\d{3}(?:\.\d+)+\s+", line))
        is_control = (bool(fis_re.search(line)) or line.upper().startswith("TARİH") or
                      "HESAP KODU" in norm(line) or line.startswith("FİŞ TOPLAM") or
                      line.startswith("FIS TOPLAM") or line.startswith("DÜZENLEYEN") or
                      line.startswith("DUZENLEYEN"))
        if is_new_account:
            if buf:
                logical.append(buf)
            buf = line
        elif buf and not is_control:
            buf += " " + line
        else:
            if buf:
                logical.append(buf); buf = ""
            logical.append(line)
    if buf:
        logical.append(buf)

    for line in logical:
        m = fis_re.search(line)
        if m:
            current_fis = m.group(1).strip()
            mm = re.match(r"0*(\d+)", current_fis)
            if mm:
                max_fis = max(max_fis, int(mm.group(1)))
            continue

        if line.upper().startswith("TARİH"):
            md = date_re.search(line)
            if md:
                current_date = md.group(0)
            mn = neden_re.search(line)
            if mn:
                current_aciklama = mn.group(1).strip()
            continue

        mn = neden_re.search(line)
        if mn:
            current_aciklama = mn.group(1).strip()
            continue

        if "HESAP KODU" in norm(line) or line.startswith("FİŞ TOPLAM") or line.startswith("FIS TOPLAM"):
            continue
        if line.startswith("DÜZENLEYEN") or line.startswith("DUZENLEYEN"):
            continue

        m = row_re.match(line)
        if not m:
            mc = re.match(r"^(\d{3}(?:\.\d+)+)\s+(.*?)\s+(\d{2}/\d{2}/\d{4})\s+(.*?)\s+([\d.]+,\d{2})\s+([\d.]+,\d{2})$", line)
            if mc:
                m = mc
        if not m or not current_fis:
            continue

        hesap = m.group(1).strip(); orta = m.group(2).strip(); tarih = m.group(3).strip()
        detay = m.group(4).strip(); borc = _tr_num(m.group(5)); alacak = _tr_num(m.group(6))
        detay_full = f"{orta} {detay}".strip()

        sonuc["satirlar"].append({
            "fisno": current_fis, "tarih": tarih, "hesap": hesap,
            "detay": detay_full, "fis_aciklama": current_aciklama,
            "borc": borc, "alacak": alacak,
        })
        if detay_full and not hesap.startswith(("102", "100")):
            sonuc["eslesmeler"][norm(f"{current_aciklama} {detay_full}")] = hesap

    sonuc["son_fis_no"] = max_fis
    return sonuc


def _yevmiye_defteri_oku(wb):
    """
    Logo/Luca 'Yevmiye Defteri' raporunu parse eder.
    Format:
      Fiş başlığı satırı: "00001-----00001-----AÇILIŞ-----01/01/2026"
      Yaprak hesap satırları (kod 2+ noktalı): kod | ad | detay | tutar | '' | ''
      Ara satırlar (kod 0-1 noktalı): kod | ad | '' | özet_borç veya özet_alacak | ...
    Tek geçişli okuma (read_only workbook uyumlu, hızlı).

    Çıktı: standart {"satirlar": [...], "son_fis_no": N, "eslesmeler": {...}}
    """
    def _sayi_local(v):
        if v is None or v == "":
            return None
        if isinstance(v, (int, float)):
            return float(v)
        try:
            s = str(v).replace(".", "").replace(",", ".") if ("," in str(v) and str(v).count(".") > 1) else str(v).replace(",", ".")
            return float(s)
        except Exception:
            return None

    sonuc = {"satirlar": [], "son_fis_no": 0, "eslesmeler": {}}
    if not wb.sheetnames:
        return None
    ws = wb[wb.sheetnames[0]]
    bas_re = re.compile(r'^0*(\d+)-{2,}0*(\d+)-{2,}(.+?)-{2,}(\d{1,2}/\d{1,2}/\d{4})\s*$')

    # Tek geçişte tüm satırları belleğe al (sadece 7 sütun lazım)
    tum_satirlar = []
    for row in ws.iter_rows(min_col=1, max_col=7, values_only=True):
        tum_satirlar.append(row)

    # Fiş başlıklarını bul
    fis_baslari = []  # (satir_indeksi, fis_no, aciklama, tarih_iso)
    for idx, row in enumerate(tum_satirlar):
        v = row[0] if row else None
        if v and isinstance(v, str):
            m = bas_re.match(v.strip())
            if m:
                gun, ay, yil = m.group(4).split("/")
                tarih_iso = f"{yil}-{int(ay):02d}-{int(gun):02d}"
                fis_baslari.append((idx, int(m.group(1)), m.group(3).strip(), tarih_iso))

    if not fis_baslari:
        return None

    max_fis = 0
    for i, (bas_idx, fis_no, aciklama, tarih_iso) in enumerate(fis_baslari):
        max_fis = max(max_fis, fis_no)
        son_idx = fis_baslari[i + 1][0] if i + 1 < len(fis_baslari) else len(tum_satirlar)

        # Fiş bloğu içindeki hesap satırlarını topla
        blok = []
        mod = None
        for idx in range(bas_idx + 1, son_idx):
            row = tum_satirlar[idx]
            if not row:
                continue
            kod = str(row[0] or "").strip()
            if not kod or not re.match(r'^\d', kod):
                continue
            derinlik = kod.count(".")
            if derinlik == 0:
                # Ana hesap: yön belirler (sütun 4 = borç, sütun 5 = alacak)
                b = _sayi_local(row[4]) if len(row) > 4 else None
                a = _sayi_local(row[5]) if len(row) > 5 else None
                if b and b > 0:
                    mod = "borc"
                elif a and a > 0:
                    mod = "alacak"
                continue
            blok.append((kod, row, mod))

        # Yaprakları belirle: başka bir kodun prefix'i olmayanlar
        kodlar = [k for k, _, _ in blok]
        kod_seti = set(kodlar)
        for kod, row, kayit_mod in blok:
            if kayit_mod is None:
                continue
            prefix = kod + "."
            if any(k.startswith(prefix) for k in kod_seti if k != kod):
                continue
            tutar = 0
            for ci in (3, 4, 5):
                if len(row) > ci:
                    v = _sayi_local(row[ci])
                    if v:
                        tutar = v; break
            if tutar == 0:
                continue
            detay = str(row[2] or "").strip() if len(row) > 2 else ""
            borc = tutar if kayit_mod == "borc" else 0.0
            alacak = tutar if kayit_mod == "alacak" else 0.0
            sonuc["satirlar"].append({
                "fisno": f"{fis_no:05d}",
                "hesap": kod,
                "detay": detay,
                "fis_aciklama": aciklama,
                "borc": borc,
                "alacak": alacak,
                "fis_tarih": tarih_iso,
            })
            if detay and not kod.startswith(("102", "100")):
                sonuc["eslesmeler"][norm(detay)] = kod

    sonuc["son_fis_no"] = max_fis
    return sonuc


def _mikro_fis_listesi_oku(wb):
    """Mikro/Zirve tipi fiş listesi formatı. Sayfa: 'fis_listesi'.
    Fiş başlığı: Tarih : dd/mm/yyyy | Belge Düzenleme Nedeni : ... | Fiş No : 00001
    Sonra: HESAP KODU | HESAP ADI | AÇIKLAMA | BORÇ | ALACAK satırları.
    FİŞ TOPLAM ile biter."""
    def _sayi(v):
        if v is None or v == "": return None
        if isinstance(v, (int, float)): return float(v)
        try: return float(str(v).replace(",", "."))
        except: return None

    sonuc = {"satirlar": [], "son_fis_no": 0, "eslesmeler": {}}
    ws = None
    for sn in wb.sheetnames:
        if "fis_listesi" in norm(sn).replace(" ", "_").lower() or "fis listesi" in norm(sn).lower():
            ws = wb[sn]; break
    if ws is None:
        ws = wb[wb.sheetnames[0]]

    tum = []
    for row in ws.iter_rows(values_only=True):
        tum.append(row)

    # Format kontrolü
    format_var = False
    for row in tum[:100]:
        for c in (row or ()):
            if c and str(c).strip() in ("Fiş No", "FİŞ NO", "FIS NO"):
                format_var = True; break
        if format_var: break
    if not format_var:
        return None

    # Fiş bloklarını bul
    fis_baslari = []
    for i, row in enumerate(tum):
        if not row: continue
        for j, c in enumerate(row):
            if c is None: continue
            cs = str(c).strip()
            if cs not in ("Fiş No", "FİŞ NO", "FIS NO"): continue
            fno_str = None
            for k in range(j+1, len(row)):
                v = row[k]
                if v is not None and str(v).strip() and str(v).strip() != ":":
                    fno_str = str(v).strip(); break
            if not fno_str: continue
            tarih_iso = ""; aciklama = ""
            for ui in range(max(0, i-3), i+1):
                uprow = tum[ui]
                if not uprow: continue
                for uj, uc in enumerate(uprow):
                    if uc is None: continue
                    ucs = str(uc).strip()
                    if ucs == "Tarih":
                        for uk in range(uj+1, len(uprow)):
                            uv = uprow[uk]
                            if uv and str(uv).strip() and str(uv).strip() != ":":
                                m = re.match(r'(\d{1,2})[./](\d{1,2})[./](\d{4})', str(uv).strip())
                                if m: tarih_iso = f"{m.group(3)}-{int(m.group(2)):02d}-{int(m.group(1)):02d}"
                                break
                    elif "Belge D" in ucs or "Nedeni" in ucs:
                        for uk in range(uj+1, len(uprow)):
                            uv = uprow[uk]
                            if uv and str(uv).strip() and str(uv).strip() != ":":
                                aciklama = str(uv).strip(); break
            m = re.match(r'0*(\d+)', fno_str)
            if m: fis_baslari.append((i, int(m.group(1)), tarih_iso, aciklama))
            break

    if not fis_baslari:
        return None

    max_fis = 0
    for idx, (bas_i, fis_no, tarih_iso, aciklama) in enumerate(fis_baslari):
        max_fis = max(max_fis, fis_no)
        son_i = fis_baslari[idx+1][0] if idx+1 < len(fis_baslari) else len(tum)
        baslik_i = None; sut = {}
        for i in range(bas_i, min(bas_i+10, son_i)):
            row = tum[i]
            if not row: continue
            for j, c in enumerate(row):
                if c is None: continue
                cs = norm(str(c))
                if cs == "HESAP KODU": sut["hesap"] = j; baslik_i = i
                elif cs == "HESAP ADI": sut["ad"] = j
                elif cs == "ACIKLAMA": sut["detay"] = j
                elif cs == "BORC": sut["borc"] = j
                elif cs == "ALACAK": sut["alacak"] = j
            if baslik_i is not None and "hesap" in sut and "borc" in sut: break
        if baslik_i is None: continue
        for i in range(baslik_i+1, son_i):
            row = tum[i]
            if not row: continue
            hkod = row[sut["hesap"]] if sut["hesap"] < len(row) else None
            if hkod is None: continue
            hkod_s = str(hkod).strip()
            if not hkod_s: continue
            if "TOPLAM" in norm(hkod_s): break
            if not re.match(r'^\d{3}(\.\d+)*$', hkod_s): continue
            b = _sayi(row[sut["borc"]]) or 0 if "borc" in sut and sut["borc"] < len(row) else 0
            a = _sayi(row[sut["alacak"]]) or 0 if "alacak" in sut and sut["alacak"] < len(row) else 0
            if b == 0 and a == 0: continue
            detay = str(row[sut["detay"]] or "").strip() if "detay" in sut and sut["detay"] < len(row) else ""
            sonuc["satirlar"].append({
                "fisno": f"{fis_no:05d}", "hesap": hkod_s, "detay": detay,
                "fis_aciklama": aciklama, "borc": float(b), "alacak": float(a),
                "fis_tarih": tarih_iso,
            })
            if detay and not hkod_s.startswith(("102", "100", "108")):
                sonuc["eslesmeler"][norm(detay)] = hkod_s

    sonuc["son_fis_no"] = max_fis
    return sonuc


def _muavin_oku(wb):
    """Muavin defter (Logo/Luca/Mikro tarzı) okuyucu.
    Yapı: her hesap için bir başlık satırı ("102.01.001 AKBANK ÇEŞME ..."), ardından
    TARİH / TİP / FİŞ NO / AÇIKLAMA / BORÇ / ALACAK başlıkları, "Nakli Yekün" devir
    satırı, hareketler ve ara toplamlar. Hesap kodu satırlarda değil bölüm
    başlığında yazar; her hareket satırına o bölümün hesabı verilir.
    Çıktı diğer okuyucularla aynı: satirlar / son_fis_no / eslesmeler."""
    tarih_re = re.compile(r"^\s*(\d{1,2})[./-](\d{1,2})[./-](\d{4})")
    hesap_re = re.compile(r"^\s*(\d{3}(?:\.\d+)*)(?:\s+(\S.*))?$")
    for ws in wb.worksheets:
        sut, hesap, satirlar, baslik_goruldu = None, None, [], 0
        for row in ws.iter_rows(values_only=True):
            hucreler = list(row)
            if not hucreler:
                continue
            a = hucreler[0]
            nrow = [norm(str(c or "")) for c in hucreler]
            # sütun başlık satırı (her hesap bölümünde tekrar eder)
            if "BORC" in nrow and "ALACAK" in nrow and any("ACIKLAMA" in c for c in nrow) \
                    and any(c.startswith("TARIH") for c in nrow):
                sut = {"tarih": next(j for j, c in enumerate(nrow) if c.startswith("TARIH")),
                       "borc": nrow.index("BORC"), "alacak": nrow.index("ALACAK"),
                       "aciklama": next(j for j, c in enumerate(nrow) if "ACIKLAMA" in c)}
                for j, c in enumerate(nrow):
                    if c in ("FIS NO", "FIS NUMARASI", "YEVMIYE NO", "FISNO", "EVRAK NO") and "fisno" not in sut:
                        sut["fisno"] = j
                    elif c in ("TIP", "FIS TIPI", "FIS TURU", "TURU"):
                        sut["tip"] = j
                baslik_goruldu += 1
                continue
            # hesap bölüm başlığı: "770.10.001 BANKA GİDERLERİ" (tek hücre) ya da kod + ad ayrı hücrede
            if isinstance(a, str) and not tarih_re.match(a):
                m = hesap_re.match(a)
                if m and (m.group(2) or (len(hucreler) > 1 and isinstance(hucreler[1], str))):
                    hesap = m.group(1)
                    continue
            if sut is None or hesap is None:
                continue
            t = hucreler[sut["tarih"]] if sut["tarih"] < len(hucreler) else None
            if isinstance(t, datetime):
                tarih = f"{t.year:04d}-{t.month:02d}-{t.day:02d}"
            elif isinstance(t, str) and tarih_re.match(t):
                g_, a_, y_ = tarih_re.match(t).groups()
                tarih = f"{y_}-{int(a_):02d}-{int(g_):02d}"
            else:
                continue          # devir, ara toplam, boş satır
            def _n(j):
                v = hucreler[j] if j is not None and j < len(hucreler) else None
                if isinstance(v, (int, float)):
                    return float(v)
                return (_tr_num(v) or 0.0) if v not in (None, "") else 0.0
            borc, alacak = _n(sut["borc"]), _n(sut["alacak"])
            if not borc and not alacak:
                continue
            def _h(anahtar):
                j = sut.get(anahtar)
                v = hucreler[j] if j is not None and j < len(hucreler) else None
                return str(v if v is not None else "").strip()
            satirlar.append({"fisno": _h("fisno"), "fis_tip": _h("tip"), "hesap": hesap, "tarih": tarih,
                             "detay": _h("aciklama"), "fis_aciklama": _h("aciklama"),
                             "borc": borc, "alacak": alacak})
        if satirlar and baslik_goruldu >= 1:
            # Son fiş no: mahsup fişlerinden (araç mahsup fişi üretir); yoksa hepsinden
            def _no(r):
                m = re.match(r"0*(\d+)", r["fisno"])
                return int(m.group(1)) if m else 0
            mahsup = [r for r in satirlar if "MAHSUP" in norm(r["fis_tip"])]
            son = max((_no(r) for r in (mahsup or satirlar)), default=0)
            eslesmeler = {norm(r["detay"]): r["hesap"] for r in satirlar
                          if r["detay"] and not r["hesap"].startswith(("102", "100"))}
            return {"satirlar": satirlar, "son_fis_no": son, "eslesmeler": eslesmeler, "tur": "muavin"}
    return None


def gecmis_fisler_oku(kaynak_path: Path):
    """Geçmiş fiş listesini okur. Üç format desteklenir:
    1) Mikro/Zirve tipi ('fis_listesi' sayfası)
    2) Fiş Aktarım Şablonu (14 sütunlu standart)
    3) Logo/Luca Yevmiye Defteri
    3 seviye cache: bellek → disk JSON → Excel parse."""
    if kaynak_path and kaynak_path.suffix.lower() == ".pdf":
        return gecmis_fisler_oku_pdf(kaynak_path)

    sonuc = {"satirlar": [], "son_fis_no": 0, "eslesmeler": {}}
    if not kaynak_path or not kaynak_path.exists():
        return sonuc

    kaynak_str = str(kaynak_path)
    kaynak_mtime = kaynak_path.stat().st_mtime

    # 1. BELLEK CACHE (en hızlı)
    global _bellek_cache
    if kaynak_str in _bellek_cache:
        cached_mtime, cached_data = _bellek_cache[kaynak_str]
        if cached_mtime == kaynak_mtime:
            return cached_data

    # 2. DİSK CACHE (.cache.json)
    cache_path = kaynak_path.with_suffix(kaynak_path.suffix + ".cache.json")
    if cache_path.exists():
        try:
            cache = json.loads(cache_path.read_text(encoding="utf-8"))
            if cache.get("_mtime") == kaynak_mtime:
                cache.pop("_mtime", None)
                _bellek_cache[kaynak_str] = (kaynak_mtime, cache)
                return cache
        except Exception:
            pass

    # 3. Excel parse (en yavaş) — devamında cache oluşturulacak
    try:
        # NOT: read_only bazı Logo/Luca çıktılarında iter_rows'un erken durmasına
        # neden oluyor; normal modda okunur ama iter_rows ile hızlı geçilir.
        wb = openpyxl.load_workbook(kaynak_path, data_only=True)
    except Exception:
        return sonuc

    def _cache_yaz(sonuc_dict):
        """Sonucu bellek + disk cache'e yazar."""
        # Bellek cache
        _bellek_cache[kaynak_str] = (kaynak_mtime, sonuc_dict)
        # Disk cache (JSON)
        try:
            veri = {"_mtime": kaynak_mtime, **sonuc_dict}
            cache_path.write_text(json.dumps(veri, ensure_ascii=False), encoding="utf-8")
        except Exception:
            pass

    ws = None
    # Öncelik 1: Mikro/Zirve (fis_listesi sayfası)
    for sn in wb.sheetnames:
        if "fis_listesi" in norm(sn).replace(" ", "_").lower() or "fis listesi" in norm(sn).lower():
            mikro = _mikro_fis_listesi_oku(wb)
            if mikro and mikro.get("satirlar"):
                _cache_yaz(mikro); return mikro
            break
    # Muavin defter (hesap bölümlü) — Mikro fiş listesi değilse ve şablon sayfası yoksa
    muavin = _muavin_oku(wb)
    if muavin:
        _cache_yaz(muavin); return muavin
    # Öncelik 2: Fiş Aktarım Şablonu
    for sn in wb.sheetnames:
        if "aktarim" in norm(sn).lower() or ("fis" in norm(sn).lower() and "listesi" not in norm(sn).lower()):
            ws = wb[sn]; break
    if ws is None:
        # Öncelik 3: Logo/Luca Yevmiye Defteri
        yev = _yevmiye_defteri_oku(wb)
        if yev:
            _cache_yaz(yev); return yev
        # Son çare: Mikro tekrar dene (sayfa adı farklı olabilir)
        mikro = _mikro_fis_listesi_oku(wb)
        if mikro and mikro.get("satirlar"):
            _cache_yaz(mikro); return mikro
        return sonuc

    # Tek geçişte tüm satırları belleğe al (read_only uyumlu, hızlı)
    tum_satirlar = []
    for row in ws.iter_rows(values_only=True):
        tum_satirlar.append(row)

    bas = None; sut = {}
    for r_idx in range(min(5, len(tum_satirlar))):
        row = tum_satirlar[r_idx]
        for c_idx, v in enumerate(row):
            vn = norm(str(v or ""))
            if vn == "FIS NO": sut["fisno"] = c_idx
            elif vn == "FIS TARIHI": sut["tarih"] = c_idx
            elif vn == "HESAP KODU": sut["hesap"] = c_idx
            elif "DETAY" in vn: sut["detay"] = c_idx
            elif vn == "BORC": sut["borc"] = c_idx
            elif vn == "ALACAK": sut["alacak"] = c_idx
            elif vn == "FIS ACIKLAMA": sut["fis_aciklama"] = c_idx
        if "fisno" in sut and "hesap" in sut:
            bas = r_idx; break
    if bas is None:
        return sonuc

    def gc(row, key):
        c = sut.get(key)
        return row[c] if c is not None and c < len(row) else None

    max_fis = 0
    for r_idx in range(bas + 1, len(tum_satirlar)):
        row = tum_satirlar[r_idx]
        fisno = gc(row, "fisno"); hesap = gc(row, "hesap")
        if fisno is None or hesap is None: continue
        fstr = str(fisno).strip(); hstr = str(hesap).strip()
        if not fstr or not hstr: continue
        detay = str(gc(row, "detay") or gc(row, "fis_aciklama") or "").strip()
        borc = gc(row, "borc") or 0; alacak = gc(row, "alacak") or 0
        sonuc["satirlar"].append({"fisno": fstr, "hesap": hstr, "detay": detay,
                                  "fis_aciklama": str(gc(row, "fis_aciklama") or ""),
                                  "borc": float(borc or 0), "alacak": float(alacak or 0)})
        m = re.match(r"0*(\d+)", fstr)
        if m: max_fis = max(max_fis, int(m.group(1)))
        if detay and not hstr.startswith(("102", "100")):
            sonuc["eslesmeler"][norm(detay)] = hstr
    sonuc["son_fis_no"] = max_fis
    _cache_yaz(sonuc)
    return sonuc

def kural_dosyasina_fis_ekle(kural_path: Path, yeni_satirlar: list):
    """
    Üretilen fişleri kural.xlsx'in 'Fiş Aktarım Şablonu' sayfasına ekler (kümülatif).
    Dosya yoksa oluşturur; sayfa yoksa ekler.
    yeni_satirlar: isleyici.fis_xlsx ile aynı sözlük formatı.
    """
    import openpyxl
    from datetime import datetime as _dt

    basliklar = ["Fiş No", "Fiş Tarihi", "Fiş Açıklama", "Hesap Kodu", "Evrak No",
                 "Evrak Tarihi", "Detay Açıklama", "Borç", "Alacak", "Miktar",
                 "Belge Türü", "Para Birimi", "Kur", "Döviz Tutar"]

    if kural_path.exists():
        wb = openpyxl.load_workbook(kural_path)
    else:
        wb = openpyxl.Workbook()
        # varsayılan boş Sheet'i sil
        if "Sheet" in wb.sheetnames and len(wb.sheetnames) == 1:
            del wb["Sheet"]

    # "Fiş Aktarım" sayfasını bul veya oluştur
    hedef = None
    for sn in wb.sheetnames:
        if "aktarim" in norm(sn).lower() or "aktar" in norm(sn).lower():
            hedef = wb[sn]; break
    if hedef is None:
        hedef = wb.create_sheet("Fiş Aktarım Şablonu")
        hedef.append(basliklar)

    # başlık yoksa ekle
    ilk_satir = [hedef.cell(1, c).value for c in range(1, len(basliklar) + 1)]
    if not ilk_satir[0] or "FIS" not in norm(str(ilk_satir[0])):
        hedef.insert_rows(1)
        for c, b in enumerate(basliklar, 1):
            hedef.cell(1, c).value = b

    def dt(iso):
        try:
            y, m, d = map(int, iso.split("-")); return _dt(y, m, d)
        except Exception:
            return iso

    for r in yeni_satirlar:
        evno = r.get("evrak_no", "")
        try:
            evno = int(evno) if str(evno).isdigit() else evno
        except Exception:
            pass
        hedef.append([
            r.get("fisno", ""), dt(r.get("fis_tarih", "")), r.get("fis_aciklama", ""),
            r.get("hesap", ""), evno, dt(r.get("evrak_tarih", "")), r.get("detay", ""),
            r.get("borc") or None, r.get("alacak") or None, None,
            r.get("belge_turu", "MF"), "", "", "",
        ])
    wb.save(kural_path)


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
    """
    Bir firma + belge tipi (banka/fatura/cek) için hesap eşleştirme yapar.
    """
    def __init__(self, mizan_path: Path, kural_path: Path | None, ogrenme_path: Path, gecmis_pdf_path: Path | None = None):
        self.hesaplar = mizan_hesaplar(mizan_path)                 # [(kod, ad)]
        self.kural = kural_excel_oku(kural_path) if kural_path else {"hesaplar": [], "anahtar_kod": {}, "talimatlar": []}
        self.ogrenme = ogrenme_oku(ogrenme_path)                  # {anahtar: kod}
        self.ogrenme_path = ogrenme_path
        # geçmiş fişlerden öğrenilen eşleşmeler (kural excel'deki Fiş Aktarım sayfası)
        self.gecmis = gecmis_fisler_oku(gecmis_pdf_path if gecmis_pdf_path and gecmis_pdf_path.exists() else kural_path)
        # hızlı arama için kod->ad
        self.kod_ad = {k: a for k, a in self.hesaplar}
        for h in self.kural["hesaplar"]:
            self.kod_ad.setdefault(h["kod"], h["ad"])

    _FATURA_NO_RE = re.compile(r"(?<![A-Z0-9])([A-Z0-9]{3}20\d{2}\d{9})(?![A-Z0-9])")

    def kayitli_faturalar(self) -> dict:
        """Geçmiş kayıtların açıklamalarında geçen e-fatura numaraları -> fiş no.
        (e-Fatura/e-Arşiv no: 3 harf/rakam + yıl + 9 hane, ör. GIB2026000000448.)
        Yeni yüklenen bir fatura burada varsa daha önce muhasebeleşmiştir."""
        if getattr(self, "_kayitli", None) is None:
            self._kayitli = {}
            for r in self.gecmis.get("satirlar", []):
                metin = str(r.get("detay", "")).upper() + " " + str(r.get("fis_aciklama", "")).upper()
                for no in self._FATURA_NO_RE.findall(metin):
                    self._kayitli.setdefault(no, str(r.get("fisno", "")))
        return self._kayitli

    def son_fis_no(self):
        """Geçmiş fişlerdeki en yüksek fiş numarası."""
        return self.gecmis.get("son_fis_no", 0)

    def hesap_adi(self, kod: str) -> str:
        return self.kod_ad.get(kod, "")

    def fatura_gecmis_eslestir(self, cari_ad: str, yon: str = "alis"):
        """
        Fatura için geçmişte gerçekten kesilmiş/kaydedilmiş fişlerden hesap seçer.
        Bulunan hesap kodları mizanda kontrol edilir — mizanda yoksa kullanılmaz.
        """
        sonuc = {"cari": "", "ana": "", "ek": "", "kdv": [], "tevkifat": [], "kaynak": ""}
        q = norm(cari_ad)
        if not q:
            return sonuc

        # Mizan hesap kodları seti
        mizan_kodlari = {k for k, _ in self.hesaplar}

        # Cari adının tamamı veya anlamlı kelimeleriyle geçmiş satırları bul.
        # SADECE TAM EŞLEŞMELERİ kabul et — bulanık eşleşme yanlış cariye yol açıyor.
        qwords = kelimeler(cari_ad)
        aday = []
        for idx, r in enumerate(self.gecmis.get("satirlar", [])):
            metin = norm(f"{r.get('detay','')} {r.get('fis_aciklama','')}")
            if not metin:
                continue
            # Tam eşleşme: cari adının tamamı metin içinde geçmeli
            if q in metin:
                sc = 100 + len(qwords)
                aday.append((sc, idx, r))
            # Unvan farklı yazılmış olabilir ("A.Ş." / "ANONİM ŞİRKETİ"): ekler hariç
            # en az 2 anlamlı kelimenin HEPSİ satırda tam kelime olarak geçmeli.
            elif len(qwords) >= 2 and qwords <= set(metin.split()):
                aday.append((50 + len(qwords), idx, r))

        if not aday:
            return sonuc

        aday.sort(key=lambda x: (x[0], x[1]), reverse=True)
        satirlar = [x[2] for x in aday]

        def say_sec(rows, kod_filtresi, borc_mu=None):
            say = {}
            son = {}
            for i, r in enumerate(rows):
                kod = str(r.get("hesap", "")).strip()
                if not kod or not kod_filtresi(kod):
                    continue
                # Mizanda kontrol — yoksa atla
                if kod not in mizan_kodlari:
                    continue
                if borc_mu is True and float(r.get("borc", 0) or 0) <= 0:
                    continue
                if borc_mu is False and float(r.get("alacak", 0) or 0) <= 0:
                    continue
                say[kod] = say.get(kod, 0) + 1
                son[kod] = i
            if not say:
                return ""
            return max(say, key=lambda k: (say[k], son[k]))

        def tutar_sirasi(rows, kod_filtresi, borc_mu=True):
            """Hesapları o carideki TOPLAM tutara göre sıralar (büyükten küçüğe).
            Satır sayısıyla seçmek, Turkcell'deki gibi aynı fişte iki gider hesabı
            (iletişim + ÖİV) olduğunda ana gideri şansa bırakıyordu."""
            top = {}
            for r in rows:
                kod = str(r.get("hesap", "")).strip()
                if not kod or not kod_filtresi(kod) or kod not in mizan_kodlari:
                    continue
                v = float((r.get("borc") if borc_mu else r.get("alacak")) or 0)
                if v > 0:
                    top[kod] = top.get(kod, 0) + v
            return sorted(top, key=lambda k: top[k], reverse=True)

        if yon == "alis":
            sonuc["cari"] = say_sec(satirlar, lambda k: k.startswith(("320", "329", "331", "335", "336")), borc_mu=False)
            def ana(k):
                return (not k.startswith(("100", "101", "102", "103", "108", "120", "121", "191", "192", "193",
                                          "194", "195", "300", "320", "329", "331", "335", "336", "360", "361",
                                          "370", "380", "391")))
            sira = tutar_sirasi(satirlar, ana)
            sonuc["ana"] = sira[0] if sira else ""
            # ikinci gider hesabı (ör. Turkcell'de 689 ÖİV) — faturadaki ek vergi buraya
            sonuc["ek"] = sira[1] if len(sira) > 1 else ""
            sonuc["kdv"] = list(dict.fromkeys(k for k in (str(r.get("hesap", "")).strip() for r in satirlar)
                                               if k.startswith("191") and k in mizan_kodlari))
            sonuc["tevkifat"] = list(dict.fromkeys(k for k in (str(r.get("hesap", "")).strip() for r in satirlar)
                                                     if k.startswith("191") and k in mizan_kodlari and k not in sonuc["kdv"]))
        else:
            sonuc["cari"] = say_sec(satirlar, lambda k: k.startswith(("120", "121")), borc_mu=True)
            sonuc["ana"] = say_sec(satirlar, lambda k: k.startswith(("600", "601", "602", "603", "610", "611", "612")), borc_mu=False)
            sonuc["kdv"] = list(dict.fromkeys(k for k in (str(r.get("hesap", "")).strip() for r in satirlar)
                                               if k.startswith("391") and k in mizan_kodlari))
            sonuc["tevkifat"] = list(dict.fromkeys(k for k in (str(r.get("hesap", "")).strip() for r in satirlar)
                                                     if k.startswith("391") and k in mizan_kodlari and k not in sonuc["kdv"]))

        if any(sonuc[k] for k in ("cari", "ana", "kdv", "tevkifat")):
            sonuc["kaynak"] = "gecmis_fatura"
        return sonuc

    def _mizan_isim_eslestir(self, aciklama: str):
        """
        Mizan hesap adlarıyla kelime bazlı kesin eşleştirme.
        Açıklamadaki (LTD/ŞTİ/SAN/TİC vb. ekler hariç) TÜM anlamlı kelimeler
        hesap adında geçmeli. En az 2 anlamlı kelime şart (tek kelimeyle
        yanlış hesaba düşme riski yüksek). Birden fazla aday eşit ölçüde
        uyuyorsa belirsiz kabul edilip boş döner (yanlış hesaba düşmesin).
        """
        qwords = kelimeler(aciklama)
        if len(qwords) < 2:
            return ""
        adaylar = []
        for kod, ad in self.hesaplar:
            adwords = kelimeler(ad)
            if qwords <= adwords:
                adaylar.append((kod, len(adwords)))
        if not adaylar:
            return ""
        adaylar.sort(key=lambda x: x[1])
        if len(adaylar) > 1 and adaylar[0][1] == adaylar[1][1]:
            return ""
        return adaylar[0][0]

    def eslestir(self, aciklama: str, onekler: tuple | None = None):
        """
        onekler verilirse yalnızca bu öneklerle başlayan hesaplar kabul edilir
        (örn. fatura carisi için 320/120). Böylece aynı açıklamaya öğrenilmiş
        bir gider kodu cari eşleştirmesini kapatmaz.
        Hesap kodu eşleştirme:
        1) Kullanıcı öğretmişse (öğrenme) — TAM eşleşme
        2) Geçmiş fişlerde birebir geçiyorsa — TAM eşleşme
        3) Mizan hesap adında açıklamanın TÜM anlamlı kelimeleri geçiyorsa
           (LTD/ŞTİ/SAN/TİC gibi ekler hariç) — tek/en yakın aday kabul
           edilir; birden fazla eşit aday varsa belirsiz sayılıp atlanır.
        Tahmin/yarı-eşleşme YOK — üstteki 3 kesin kural dışında bulunamazsa
        boş döner → 198.01.001 fallback.
        """
        nq = norm(aciklama)
        if not nq:
            return "", ""

        mizan_kodlari = {k for k, _ in self.hesaplar}

        def _mizan_kontrol(kod, kaynak):
            if not kod:
                return "", ""
            if onekler and not kod.startswith(onekler):
                return "", ""
            if kod in mizan_kodlari:
                return kod, kaynak
            return "", ""

        # 1) Öğrenilen eşleşme — TAM eşleşme (anahtar açıklamada geçmeli)
        for anahtar, kod in self.ogrenme.items():
            if not anahtar:
                continue
            if anahtar in nq or nq in anahtar:
                r = _mizan_kontrol(kod, "ogrenme")
                if r[0]: return r

        # 2) Geçmiş fişlerden — TAM eşleşme
        gecmis_es = self.gecmis.get("eslesmeler", {})
        for anahtar, kod in gecmis_es.items():
            if not anahtar:
                continue
            if anahtar in nq or nq in anahtar:
                r = _mizan_kontrol(kod, "gecmis")
                if r[0]: return r

        # 3) Mizan hesap adıyla kelime eşleşmesi — TÜM anlamlı kelimeler
        #    (LTD/ŞTİ/SAN/TİC gibi ekler hariç) hesap adında geçmeli.
        kod = self._mizan_isim_eslestir(aciklama)
        if kod and (not onekler or kod.startswith(onekler)):
            return kod, "mizan"

        # 4) Hiçbir kesin eşleşme yok → boş döner → 198.01.001
        return "", ""

    def ogret(self, aciklama: str, kod: str):
        """Bir açıklama -> kod eşleşmesini öğrenir (kalıcı)."""
        anahtar = norm(aciklama)
        if not anahtar or not kod:
            return
        self.ogrenme[anahtar] = kod
        ogrenme_yaz(self.ogrenme_path, self.ogrenme)
