"""
İşleyici — ham belge verisini Fiş Aktarım Şablonu satırlarına çevirir.
Her satır sözlük (14 sütun karşılığı):
  fisno, fis_tarih(iso), fis_aciklama, hesap, evrak_no, evrak_tarih(iso),
  detay, borc, alacak, belge_turu(MF)

Şimdilik iskelet: banka/fatura/cek için temel akış. OCR/metin satırlarını
kural motoruyla eşleştirir, dengeli çift kayıt üretir. Kural dosyasındaki
banka hesabı / karşı hesap mantığı geliştirilecek.
"""
import io, re
from datetime import datetime
from app.kurallar import norm, dosya_banka_anahtari, banka_kisa_adi

# Gerçek işlem değil, hesap özeti/metadata satırı olduğu belli olan açıklamalar
# (ör. "Sicil: 26920050511927770340630") — bunlar tabloya karışırsa hem hesap
# eşleşmez hem de anlamsız devasa bir "tutar" üretebilir.
_METADATA_RE = re.compile(
    r"^(SICIL|IBAN|SUBE|VKN|TCKN|MUSTERI NO|HESAP NO|HESAP SAHIBI)\b"
)
_MAKUL_TUTAR_UST_SINIR = 1_000_000_000  # bu üstü gerçek bir banka hareketi değil, hatalı okunmuş sayıdır
_SAAT_ONEK_RE = re.compile(r"^\d{1,2}:\d{2}(:\d{2})?\s+")  # açıklama başındaki "12:31:21 " gibi saat
_ISLEMNO_RE = re.compile(r"\b\d{10,}\b")  # İşlem/referans no gibi uzun sayı dizileri açıklamaya karışmasın


def _ddmmyyyy(iso: str) -> str:
    try:
        y, m, d = map(int, iso.split("-"))
        return f"{d:02d}.{m:02d}.{y}"
    except Exception:
        return iso or ""


# ----------------------------------------------------------------- tarih/sayı ayrıştırma
def _tarih_iso(s):
    if s is None:
        return ""
    if isinstance(s, datetime):
        return f"{s.year:04d}-{s.month:02d}-{s.day:02d}"
    s = str(s).strip()
    for fmt in ("%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d.%m.%y"):
        try:
            d = datetime.strptime(s, fmt)
            return f"{d.year:04d}-{d.month:02d}-{d.day:02d}"
        except Exception:
            pass
    return ""


def _sayi(s):
    if s is None or s == "":
        return None
    if isinstance(s, (int, float)):
        return float(s)
    s = str(s).strip().replace(" ", "")
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    s = re.sub(r"[^0-9.\-]", "", s)
    try:
        return float(s)
    except Exception:
        return None


# ----------------------------------------------------------------- fiş satırı
def _sat(fisno, tarih, aciklama, hesap, borc, alacak, evrak_no="", detay="", kaynak="", kaynak_dosya=""):
    return {
        "fisno": fisno, "fis_tarih": tarih, "fis_aciklama": aciklama,
        "hesap": hesap, "evrak_no": evrak_no, "evrak_tarih": tarih,
        "detay": detay or aciklama, "borc": round(borc, 2), "alacak": round(alacak, 2),
        "belge_turu": "MF", "kaynak": kaynak, "kaynak_dosya": kaynak_dosya,
    }


# ----------------------------------------------------------------- excel tablo -> satırlar
def _tablo_satirlari(hamlar):
    """Excel/PDF tablolarından (tarih, açıklama, tutar, borc/alacak) çıkarımı dener.
    Başlık satırını bulup sütunları eşler. Genel amaçlı, banka dökümü odaklı."""
    kayitlar = []
    for h in hamlar:
        for tablo in h.get("tablolar", []):
            if not tablo:
                continue
            bas_idx = None; harita = {}
            for i, row in enumerate(tablo[:30]):
                nrow = [norm(str(c)) for c in row]
                for j, c in enumerate(nrow):
                    if c in ("TARIH", "ISLEM TARIHI", "VALOR", "VALOR TARIHI"):
                        harita["tarih"] = j
                    elif "ACIKLAMA" in c or "ISLEM" == c:
                        harita["aciklama"] = j
                    elif c in ("TUTAR", "ISLEM TUTARI", "MIKTAR"):
                        harita["tutar"] = j
                    elif "BORC" in c:
                        harita["borc"] = j
                    elif "ALACAK" in c:
                        harita["alacak"] = j
                    elif "GONDER" in c or "UNVAN" in c or "FIRMA" in c:
                        harita.setdefault("aciklama", j)
                if "tarih" in harita and ("tutar" in harita or "borc" in harita or "alacak" in harita):
                    bas_idx = i; break
            if bas_idx is None:
                continue
            for row in tablo[bas_idx + 1:]:
                def g(key):
                    j = harita.get(key)
                    return row[j] if j is not None and j < len(row) else None
                tarih = _tarih_iso(g("tarih"))
                acik = _SAAT_ONEK_RE.sub("", str(g("aciklama") or "").strip())
                if "borc" in harita or "alacak" in harita:
                    borc = _sayi(g("borc")) or 0
                    alacak = _sayi(g("alacak")) or 0
                    tutar = borc - alacak
                else:
                    tutar = _sayi(g("tutar")) or 0
                if not tarih and not acik:
                    continue
                if tutar == 0:
                    continue
                if _METADATA_RE.match(norm(acik)):
                    continue
                if abs(tutar) > _MAKUL_TUTAR_UST_SINIR:
                    continue
                kayitlar.append({"tarih": tarih, "aciklama": acik, "tutar": tutar,
                                 "dosya": h.get("dosya", "")})
    return kayitlar


def _ham_metin_satirlari(hamlar):
    """OCR/PDF ham metninden (tarih ... açıklama ... tutar) desenli satırları çıkarır."""
    kayitlar = []
    tarih_re = re.compile(r"\b(\d{1,2}[./]\d{1,2}[./]\d{2,4})\b")
    tutar_re = re.compile(r"-?\d{1,3}(?:[.\s]\d{3})*(?:[.,]\d{2})|-?\d+[.,]\d{2}")
    for h in hamlar:
        metin = h.get("ham_metin", "")
        if not metin:
            continue
        for hat in metin.splitlines():
            hat = hat.strip()
            if not hat:
                continue
            tm = tarih_re.search(hat)
            # Tarih satırın hemen başında değilse (küçük OCR gürültüsü hariç)
            # bu bir işlem satırı değil, bir üstbilgi/etiket satırıdır
            # (ör. "Tarih Aralığı : 01.08.2026 - 31.08.2026") — atla.
            if not tm or tm.start() > 3:
                continue
            # Tutarları SADECE tarihten SONRAKİ kısımda ara; aksi halde
            # "03.08.2026" gibi tarihin kendi ".08" parçası yanlışlıkla bir
            # tutar sanılabiliyordu.
            kalan = hat[tm.end():]
            tutarlar = tutar_re.findall(kalan)
            if not tutarlar:
                continue
            tarih = _tarih_iso(tm.group(1))
            # Ekstrede aynı satırda hem İŞLEM TUTARI hem de (o hareket sonrası)
            # BAKİYE gösteriliyorsa (ör. Vakıfbank: "... -5.000,00 183.681,28
            # FAST Anlık Ödeme") sıradaki İLK tutar gerçek işlem tutarıdır;
            # sonraki(ler) bakiyedir — SON değeri almak bakiyeyi tutar sanardı.
            tutar = _sayi(tutarlar[0])
            if tutar is None or tutar == 0:
                continue
            if abs(tutar) > _MAKUL_TUTAR_UST_SINIR:
                continue
            acik = _ISLEMNO_RE.sub("", kalan)
            for t in tutarlar:
                acik = acik.replace(t, "")
            acik = re.sub(r"\s+", " ", acik).strip(" \t-|")
            acik = _SAAT_ONEK_RE.sub("", acik)
            if _METADATA_RE.match(norm(acik)):
                continue
            kayitlar.append({"tarih": tarih, "aciklama": acik, "tutar": tutar,
                             "dosya": h.get("dosya", "")})
    return kayitlar


def _kayitlar(hamlar):
    """Önce tablo, tablo yoksa ham metin (OCR) satırları."""
    k = _tablo_satirlari(hamlar)
    if not k:
        k = _ham_metin_satirlari(hamlar)
    return k


# ----------------------------------------------------------------- BANKA
def isle_banka(hamlar, km, fis0, banka_hesap_kodu=""):
    """
    Banka dökümü -> fiş. Aynı banka hesabına ait, aynı tarihli hareketler TEK
    fişte toplanır (fiş no ve fiş tarihi ortak, her hareket kendi Borç/Alacak
    satır çiftiyle fişe eklenir); farklı tarih veya farklı banka hesabı yeni
    bir fiş no başlatır.
    Bankaya para GİRİŞİ (alacak, +): banka borç / karşı alacak
    Bankadan ÇIKIŞ (borç, -): karşı borç / banka alacak

    Banka hesabı DOSYA BAZINDA belirlenir:
      1) dosyadaki IBAN/hesap no daha önce öğrenilmiş bir hesaba bağlıysa -> otomatik
      2) değilse elle verilen banka_hesap_kodu / kural excel (vadesiz ana) / mizanda
         tek 102.x hesap -> varsayılan olarak TÜM dosyalara uygulanır
      3) hâlâ yoksa 102.01.001'e düşer ve uyarı verir
    Aynı işlemde tek bir YENİ (henüz öğrenilmemiş) IBAN/hesap no ile birlikte elle
    banka_hesap_kodu verilmişse, bu eşleşme kalıcı öğrenilir — sonraki aylarda bu
    dosya, diğer bankalarla karışık bir toplu işlemde bile otomatik doğru hesaba düşer.
    """
    uyarilar = []
    kayitlar = _kayitlar(hamlar)
    if not kayitlar:
        for h in hamlar:
            if h.get("ham_metin"):
                uyarilar.append(f"{h.get('dosya')}: tablo çıkarılamadı, ham metin var — elle düzenleme gerekebilir")
        return [], uyarilar

    # 1) her dosyanın IBAN/hesap no anahtarını çıkar, öğrenilmiş eşleşmeye bak
    dosya_anahtar = {}
    dosya_hesap = {}
    for h in hamlar:
        fn = h.get("dosya", "")
        anahtar = dosya_banka_anahtari(h)
        dosya_anahtar[fn] = anahtar
        if anahtar and anahtar in km.banka_eslestirme:
            dosya_hesap[fn] = km.banka_eslestirme[anahtar]

    # 2) elle verilen kod + tek bir YENİ anahtar varsa -> kalıcı öğren
    manuel_hesap = (banka_hesap_kodu or "").strip()
    if manuel_hesap:
        yeni_anahtarlar = {a for a in dosya_anahtar.values() if a and a not in km.banka_eslestirme}
        if len(yeni_anahtarlar) == 1:
            km.banka_hesap_ogren(yeni_anahtarlar.pop(), manuel_hesap)
            for fn, a in dosya_anahtar.items():
                if a and a in km.banka_eslestirme:
                    dosya_hesap[fn] = km.banka_eslestirme[a]

    # 3) hiçbir şekilde belirlenemeyen dosyalar için varsayılan zincir
    varsayilan_hesap = manuel_hesap
    if not varsayilan_hesap:
        for hh in km.kural["hesaplar"]:
            if hh["kod"].startswith("102") and ("VADESIZ" in norm(hh.get("kullanim", "")) or "ANA" in norm(hh.get("kullanim", ""))):
                varsayilan_hesap = hh["kod"]; break
    if not varsayilan_hesap:
        for hh in km.kural["hesaplar"]:
            if hh["kod"].startswith("102"):
                varsayilan_hesap = hh["kod"]; break
    if not varsayilan_hesap:
        mizan_banka = [k for k, a in km.hesaplar if k.startswith("102")]
        if len(mizan_banka) == 1:
            varsayilan_hesap = mizan_banka[0]
    varsayilan_kullanildi = False
    if not varsayilan_hesap:
        varsayilan_hesap = "102.01.001"
        varsayilan_kullanildi = True

    for fn, kod in sorted(dosya_hesap.items()):
        uyarilar.append(f"{fn}: {kod} olarak otomatik tanındı (IBAN/hesap no eşleşmesi)")
    eksik_dosyalar = sorted(fn for fn in dosya_anahtar if fn not in dosya_hesap)
    if eksik_dosyalar and varsayilan_kullanildi:
        uyarilar.append("Banka hesabı belirlenemedi (kural/mizan'da net değil), 102.01.001 varsayıldı — "
                        "isteğe banka_hesap_kodu vererek düzeltebilirsiniz: " + ", ".join(eksik_dosyalar))

    fisler = []
    fis = fis0
    grup_fisno = {}  # (banka_hesap, tarih) -> fisno — aynı bankanın aynı tarihli hareketleri tek fişte toplanır
    for k in sorted(kayitlar, key=lambda x: (x["tarih"] or "", x.get("dosya", ""))):
        kd = k.get("dosya", "")
        banka_hesap = dosya_hesap.get(kd, varsayilan_hesap)
        grup_anahtar = (banka_hesap, k["tarih"])
        if grup_anahtar not in grup_fisno:
            grup_fisno[grup_anahtar] = f"{fis:05d}"
            fis += 1
        fisno = grup_fisno[grup_anahtar]
        karsi, kaynak = km.eslestir(k["aciklama"])
        if not karsi:
            karsi = ""
            uyarilar.append(f"{k['aciklama'][:30]}: hesap eşleşmedi")
        tutar = abs(k["tutar"])
        # Fiş Açıklama = "BANKA-TARİH" (ör. "YKB-05.04.2026"); Detay Açıklama
        # ise gerçek işlem metni. Tüm bankalarda aynı kural geçerli.
        fis_aciklama = f"{banka_kisa_adi(km.hesap_adi(banka_hesap), banka_hesap)}-{_ddmmyyyy(k['tarih'])}"
        aciklama = k["aciklama"]
        if k["tutar"] >= 0:
            fisler.append(_sat(fisno, k["tarih"], fis_aciklama, banka_hesap, tutar, 0, detay=aciklama, kaynak="banka", kaynak_dosya=kd))
            fisler.append(_sat(fisno, k["tarih"], fis_aciklama, karsi, 0, tutar, detay=aciklama, kaynak=kaynak, kaynak_dosya=kd))
        else:
            fisler.append(_sat(fisno, k["tarih"], fis_aciklama, karsi, tutar, 0, detay=aciklama, kaynak=kaynak, kaynak_dosya=kd))
            fisler.append(_sat(fisno, k["tarih"], fis_aciklama, banka_hesap, 0, tutar, detay=aciklama, kaynak="banka", kaynak_dosya=kd))
    return fisler, uyarilar


# ----------------------------------------------------------------- FATURA
def isle_fatura(hamlar, km, fis0):
    uyarilar = []
    kayitlar = _kayitlar(hamlar)
    if not kayitlar:
        for h in hamlar:
            if h.get("ham_metin"):
                uyarilar.append(f"{h.get('dosya')}: tablo çıkarılamadı (OCR ham metin) — elle düzenleme gerekebilir")
        return [], uyarilar
    fisler = []
    fis = fis0
    for k in sorted(kayitlar, key=lambda x: x["tarih"] or ""):
        fisno = f"{fis:05d}"
        gider, kaynak = km.eslestir(k["aciklama"])
        cari, _ = km.eslestir(k["aciklama"])
        toplam = abs(k["tutar"])
        matrah = round(toplam / 1.20, 2); kdv = round(toplam - matrah, 2)
        if not gider:
            uyarilar.append(f"{k['aciklama'][:30]}: gider hesabı eşleşmedi")
        fisler.append(_sat(fisno, k["tarih"], k["aciklama"], gider, matrah, 0, detay=k["aciklama"], kaynak=kaynak))
        fisler.append(_sat(fisno, k["tarih"], k["aciklama"], "191.01.020", kdv, 0, detay=k["aciklama"] + " (%20 KDV)", kaynak="kdv"))
        fisler.append(_sat(fisno, k["tarih"], k["aciklama"], cari or "320.01.001", 0, toplam, detay=k["aciklama"], kaynak=kaynak))
        fis += 1
    return fisler, uyarilar


# ----------------------------------------------------------------- ÇEK
def isle_cek(hamlar, km, fis0):
    uyarilar = []
    kayitlar = _kayitlar(hamlar)
    if not kayitlar:
        uyarilar.append("Çek listesi tablo olarak okunamadı (Excel bekleniyor)")
        return [], uyarilar
    fisler = []
    fis = fis0
    for k in sorted(kayitlar, key=lambda x: x["tarih"] or ""):
        fisno = f"{fis:05d}"
        cari, kaynak = km.eslestir(k["aciklama"])
        tutar = abs(k["tutar"])
        fisler.append(_sat(fisno, k["tarih"], k["aciklama"], "101.01.001", tutar, 0, detay=k["aciklama"], kaynak="cek"))
        fisler.append(_sat(fisno, k["tarih"], k["aciklama"], cari or "120.01.001", 0, tutar, detay=k["aciklama"], kaynak=kaynak))
        fis += 1
    return fisler, uyarilar


def isle(tip, hamlar, km, fis0, banka_hesap_kodu=""):
    if tip == "banka":
        return isle_banka(hamlar, km, fis0, banka_hesap_kodu)
    if tip == "fatura":
        return isle_fatura(hamlar, km, fis0)
    if tip == "cek":
        return isle_cek(hamlar, km, fis0)
    return [], ["Bilinmeyen tip"]


# ----------------------------------------------------------------- çıktı (Fiş Aktarım Şablonu)
def fis_xlsx(satirlar):
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Fiş Aktarım Şablonu"
    basliklar = ["Fiş No", "Fiş Tarihi", "Fiş Açıklama", "Hesap Kodu", "Evrak No",
                 "Evrak Tarihi", "Detay Açıklama", "Borç", "Alacak", "Miktar",
                 "Belge Türü", "Para Birimi", "Kur", "Döviz Tutar"]
    ws.append(basliklar)
    hf = Font(bold=True, color="FFFFFF"); hfill = PatternFill("solid", start_color="1F4E78")
    thin = Side(style="thin", color="D0D0D0"); bd = Border(left=thin, right=thin, top=thin, bottom=thin)
    for c in ws[1]:
        c.font = hf; c.fill = hfill; c.alignment = Alignment(horizontal="center", vertical="center"); c.border = bd

    def dt(iso):
        try:
            y, m, d = map(int, iso.split("-")); return datetime(y, m, d)
        except Exception:
            return iso

    for r in satirlar:
        evno = r.get("evrak_no", "")
        evno = int(evno) if str(evno).isdigit() else evno
        ws.append([
            r.get("fisno", ""), dt(r.get("fis_tarih", "")), r.get("fis_aciklama", ""),
            r.get("hesap", ""), evno, dt(r.get("evrak_tarih", "")), r.get("detay", ""),
            r.get("borc") or None, r.get("alacak") or None, None,
            r.get("belge_turu", "MF"), "", "", "",
        ])
    sfill = PatternFill("solid", start_color="FFF3CD")
    for row in ws.iter_rows(min_row=2, max_row=len(satirlar) + 1):
        for c in row:
            c.border = bd
        row[1].number_format = "dd/mm/yyyy"; row[5].number_format = "dd/mm/yyyy"
        row[7].number_format = "#,##0.00"; row[8].number_format = "#,##0.00"
        if not row[3].value:
            row[3].fill = sfill
    for col, w in zip("ABCDEFGHIJKLMN", [10, 12, 26, 14, 20, 12, 30, 13, 13, 8, 10, 10, 8, 12]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    bio = io.BytesIO(); wb.save(bio); bio.seek(0)
    return bio


def fis_xml(satirlar, vir0=80001, tir0=950001):
    from collections import OrderedDict
    gruplar = OrderedDict()
    for r in satirlar:
        gruplar.setdefault(r.get("fisno", ""), []).append(r)
    now = datetime.now()
    def esc(s): return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    def ds(iso):
        try:
            y, m, d = map(int, iso.split("-")); return f"{d:02d}.{m:02d}.{y}"
        except Exception:
            return iso
    out = ['<?xml version="1.0" encoding="ISO-8859-9"?>', "<GL_VOUCHERS>"]
    vir = vir0; tir = tir0; ln = 0
    for fis, lines in gruplar.items():
        d = ds(lines[0].get("fis_tarih", "")); parcali = d.split(".")
        mm = parcali[1] if len(parcali) > 1 else "1"; yy = parcali[2] if len(parcali) > 2 else "2026"
        td = round(sum(l["borc"] for l in lines), 2); tc = round(sum(l["alacak"] for l in lines), 2)
        out += ['  <GL_VOUCHER DBOP="INS">',
                "    <INTERNAL_REFERENCE>%d</INTERNAL_REFERENCE>" % vir, "    <TYPE>4</TYPE>",
                "    <NUMBER>%s</NUMBER>" % fis, "    <DATE>%s</DATE>" % d,
                "    <TOTAL_DEBIT>%.2f</TOTAL_DEBIT>" % td, "    <TOTAL_CREDIT>%.2f</TOTAL_CREDIT>" % tc,
                "    <CREATED_BY>6</CREATED_BY>", "    <DATE_CREATED>%s</DATE_CREATED>" % now.strftime("%d.%m.%Y"),
                "    <HOUR_CREATED>%d</HOUR_CREATED>" % now.hour, "    <MIN_CREATED>%d</MIN_CREATED>" % now.minute,
                "    <SEC_CREATED>%d</SEC_CREATED>" % now.second, "    <CURRSEL_TOTALS>1</CURRSEL_TOTALS>",
                "    <DATA_REFERENCE>%d</DATA_REFERENCE>" % vir, "    <TRANSACTIONS>"]
        for l in lines:
            ln += 1; hes = l.get("hesap", ""); parent = hes.split(".")[0] if hes else ""
            out.append("      <TRANSACTION>")
            out.append("        <INTERNAL_REFERENCE>%d</INTERNAL_REFERENCE>" % tir)
            if l["alacak"] > 0: out.append("        <SIGN>1</SIGN>")
            out.append("        <GL_CODE>%s</GL_CODE>" % hes)
            out.append("        <PARENT_GLCODE>%s</PARENT_GLCODE>" % parent)
            if l["borc"] > 0: out.append("        <DEBIT>%.2f</DEBIT>" % l["borc"]); amt = l["borc"]
            else: out.append("        <CREDIT>%.2f</CREDIT>" % l["alacak"]); amt = l["alacak"]
            out += ["        <LINENO>%d</LINENO>" % ln, "        <DESCRIPTION>%s</DESCRIPTION>" % esc(l.get("detay", "")),
                    "        <TC_XRATE>1</TC_XRATE>", "        <TC_AMOUNT>%.2f</TC_AMOUNT>" % amt,
                    "        <QUANTITY>0</QUANTITY>", "        <DATA_REFERENCE>%d</DATA_REFERENCE>" % tir,
                    "        <DETLIST/>", "        <DEFNFLDLIST/>", "        <MONTH>%s</MONTH>" % int(mm),
                    "        <YEAR>%s</YEAR>" % int(yy), "        <DOC_DATE>%s</DOC_DATE>" % d,
                    "        <GUID>00000000-0000-0000-0000-%012d</GUID>" % tir, "      </TRANSACTION>"]
            tir += 1
        out += ["    </TRANSACTIONS>", "  </GL_VOUCHER>"]; vir += 1
    out.append("</GL_VOUCHERS>")
    return "\n".join(out).encode("iso-8859-9", "replace")
