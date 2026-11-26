"""
e-Fatura / e-Arşiv UBL-TR XML okuyucu.

GİB portalından ya da entegratörden indirilen faturalar (tek .xml ya da içinde
.xml + .html olan .zip) bu modülle okunur. XML'de tüm bilgiler kesin alanlarda
durur: fatura no, tarih, satıcı/alıcı VKN-TCKN ve unvan, her KDV oranının matrahı
ve vergisi, tevkifat, diğer vergiler, kalemler. Bu yüzden XML okunan faturada
düzenli ifadeye ya da yapay zekâya gerek yoktur ve aynı numaralı PDF'in önüne geçer.

Çıktı, fatura_pdf._sayfa_oku ile AYNI şekildedir (isle_fatura hiç değişmeden çalışır):
  fatura_no, tarih (ISO), cari_ad, tur, senaryo, kalemler [{oran, matrah, kdv}],
  toplam (ÖDENECEK: KDV dahil − tevkifat), tevkifat, tevkifat_kod, ek_vergi,
  yon, dosya, kaynak="xml", eksik, yz, kalem_aciklamalari, kalem_detay
Ek alanlar: vkn, profil, fatura_tipi, uyari
"""
import io
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

_NS = {
    "cbc": "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2",
    "cac": "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2",
}
_KOK = ("Invoice", "CreditNote")
KDV_KODU = "0015"
UZANTILAR = (".xml", ".zip")


# ------------------------------------------------------------------ dosya türü
def _zip_mi(data: bytes) -> bool:
    return data[:4] == b"PK\x03\x04"


def _xml_fatura_mi(data: bytes) -> bool:
    bas = data[:4096].lstrip(b"\xef\xbb\xbf \r\n\t")
    if not bas.startswith(b"<"):
        return False
    return b"urn:oasis:names:specification:ubl:schema:xsd:Invoice-2" in data[:20000] \
        or re.search(rb"<(\w+:)?Invoice[\s>]", data[:20000]) is not None


def xmlleri(data: bytes) -> list:
    """Dosya içeriğinden fatura XML'lerini çıkarır: [(ad, bytes)].
    .xlsx/.docx gibi Office ZIP'leri ([Content_Types].xml içerir) fatura sayılmaz."""
    if _zip_mi(data):
        try:
            z = zipfile.ZipFile(io.BytesIO(data))
        except zipfile.BadZipFile:
            return []
        adlar = z.namelist()
        if "[Content_Types].xml" in adlar:
            return []
        out = []
        for ad in adlar:
            if not ad.lower().endswith(".xml") or ad.endswith("/"):
                continue
            bilgi = z.getinfo(ad)
            if bilgi.file_size > 30 * 1024 * 1024:      # makul olmayan büyüklük
                continue
            icerik = z.read(ad)
            if _xml_fatura_mi(icerik):
                out.append((ad, icerik))
        return out
    return [("", data)] if _xml_fatura_mi(data) else []


def ubl_mu(data: bytes) -> bool:
    """İçerik (tek XML ya da ZIP) en az bir UBL fatura içeriyor mu."""
    return bool(xmlleri(data))


def html_gorunum(data: bytes) -> bytes | None:
    """Faturanın okunur görüntüsü (HTML), şu sırayla:
    1) ZIP içinde GİB/entegratörün ürettiği HTML,
    2) XML'in içine gömülü XSLT şablonu (e-Fatura'da cac:AdditionalDocumentReference) ile dönüştürme,
    3) o da yoksa / çalışmazsa XML alanlarından kurulan sade fatura sayfası."""
    if _zip_mi(data):
        try:
            z = zipfile.ZipFile(io.BytesIO(data))
            ad = next((a for a in z.namelist() if a.lower().endswith((".html", ".htm"))), None)
            if ad:
                return z.read(ad)
        except zipfile.BadZipFile:
            return None
    xmller = xmlleri(data)
    if not xmller:
        return None
    xml = xmller[0][1]
    return xslt_gorunum(xml) or sade_gorunum(xml)


def _gomulu_xslt(kok) -> bytes | None:
    """UBL içindeki gömülü XSLT (base64) şablonu."""
    import base64
    for ref in kok.findall("cac:AdditionalDocumentReference", _NS):
        ek = ref.find("cac:Attachment/cbc:EmbeddedDocumentBinaryObject", _NS)
        if ek is None or not (ek.text or "").strip():
            continue
        ad = (ek.get("filename") or "").lower()
        tur = (_t(ref, "cbc:DocumentType") + " " + _t(ref, "cbc:DocumentTypeCode")).upper()
        if ad.endswith((".xslt", ".xsl")) or "XSLT" in tur:
            try:
                return base64.b64decode(re.sub(r"\s+", "", ek.text))
            except ValueError:
                continue
    return None


def xslt_gorunum(xml_bytes: bytes) -> bytes | None:
    """Faturaya gömülü XSLT ile HTML üretir (lxml gerekir). Şablon dosya/ağ erişemez."""
    try:
        from lxml import etree as LET
    except ImportError:
        return None
    try:
        kok = ET.fromstring(xml_bytes)
        sablon = _gomulu_xslt(kok)
        if not sablon:
            return None
        guvenli = LET.XMLParser(resolve_entities=False, no_network=True, huge_tree=True)
        xslt = LET.XSLT(LET.fromstring(sablon, guvenli),
                        access_control=LET.XSLTAccessControl.DENY_ALL)
        sonuc = xslt(LET.fromstring(xml_bytes, guvenli))
        html = bytes(sonuc)
        return html if html.strip() else None
    except Exception:          # bozuk / desteklenmeyen şablon: sade görünüme düşülür
        return None


def sade_gorunum(xml_bytes: bytes) -> bytes | None:
    """Gömülü şablonu olmayan XML için sade fatura sayfası."""
    from html import escape as e
    try:
        kok = ET.fromstring(xml_bytes)
        h = ayristir(xml_bytes)
    except Exception:
        return None
    para = h["para"]

    def nf(v):
        return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    def taraf(baslik, t):
        return f"<div class='taraf'><div class='bas'>{baslik}</div><b>{e(t['ad'] or '—')}</b>" \
               f"{'<div>VKN/TCKN: ' + e(t['vkn']) + '</div>' if t['vkn'] else ''}</div>"

    satir = []
    for i, ln in enumerate(kok.findall("cac:InvoiceLine", _NS) + kok.findall("cac:CreditNoteLine", _NS), 1):
        mik = ln.find("cbc:InvoicedQuantity", _NS)
        if mik is None:
            mik = ln.find("cbc:CreditedQuantity", _NS)
        miktar = f"{nf(float(mik.text))} {e(mik.get('unitCode') or '')}" if mik is not None and mik.text else ""
        oran = ""
        for st in ln.findall("cac:TaxTotal/cac:TaxSubtotal", _NS):
            if _t(st, "cac:TaxCategory/cac:TaxScheme/cbc:TaxTypeCode") == KDV_KODU:
                oran = f"%{_n(st, 'cbc:Percent'):g}"
        ad = _t(ln, "cac:Item/cbc:Name") or _t(ln, "cac:Item/cbc:Description")
        satir.append(f"<tr><td>{i}</td><td class='l'>{e(ad)}</td><td>{miktar}</td>"
                     f"<td>{nf(_n(ln, 'cac:Price/cbc:PriceAmount'))}</td><td>{oran}</td>"
                     f"<td>{nf(_n(ln, 'cbc:LineExtensionAmount'))}</td></tr>")
    lmt = kok.find("cac:LegalMonetaryTotal", _NS)
    top = [("Mal/hizmet toplamı", _n(lmt, "cbc:LineExtensionAmount")),
           ("Toplam iskonto", _n(lmt, "cbc:AllowanceTotalAmount"))]
    for k in h["kalemler"]:
        top.append((f"Hesaplanan KDV (%{k['oran']})", k["kdv"] / (h["kur"] or 1)))
    if h["tevkifat"]:
        top.append(("KDV tevkifatı", h["tevkifat"] / (h["kur"] or 1)))
    top += [("Vergiler dahil toplam", _n(lmt, "cbc:TaxInclusiveAmount")),
            ("Ödenecek tutar", _n(lmt, "cbc:PayableAmount"))]
    notlar = "".join(f"<div>{e(n.text)}</div>" for n in kok.findall("cbc:Note", _NS) if (n.text or "").strip())
    tarih = ".".join(reversed(h["tarih"].split("-"))) if h["tarih"] else ""
    sayfa = f"""<!DOCTYPE html><html lang="tr"><head><meta charset="utf-8"><title>{e(h['fatura_no'])}</title>
<style>body{{font-family:Arial,sans-serif;font-size:13px;color:#222;max-width:900px;margin:20px auto;padding:0 16px}}
h1{{font-size:18px;margin:0 0 4px}} .ust{{display:flex;gap:16px;margin:14px 0}} .taraf{{flex:1;border:1px solid #ccc;padding:8px;border-radius:4px}}
.bas{{font-size:11px;color:#777;margin-bottom:3px}} table{{border-collapse:collapse;width:100%;margin-top:10px}}
td,th{{border:1px solid #ccc;padding:4px 6px;text-align:right}} th{{background:#f2f2f2}} .l{{text-align:left}}
.top{{width:auto;margin-left:auto}} .top tr:last-child td{{font-weight:bold}} .not{{margin-top:12px;color:#555;font-size:12px}}
.uyari{{color:#8a6d1a;font-size:12px;margin-top:6px}}</style></head><body>
<h1>{e(h['profil'] or 'e-FATURA')} · {e(h['fatura_tipi'])}</h1>
<div>Fatura no: <b>{e(h['fatura_no'])}</b> · Tarih: <b>{tarih}</b>{' · Para birimi: ' + e(para) if para not in ('TRY', 'TL') else ''}</div>
<div class="hint" style="color:#888;font-size:11px">ETTN: {e(h['uuid'])}</div>
<div class="ust">{taraf('Satıcı', h['satici'])}{taraf('Alıcı', h['alici'])}</div>
<table><thead><tr><th>#</th><th class="l">Mal / hizmet</th><th>Miktar</th><th>Birim fiyat</th><th>KDV</th><th>Tutar</th></tr></thead>
<tbody>{''.join(satir)}</tbody></table>
<table class="top">{''.join(f"<tr><td class='l'>{e(a)}</td><td>{nf(v)} {e(para)}</td></tr>" for a, v in top if v or a in ('Vergiler dahil toplam', 'Ödenecek tutar'))}</table>
{f'<div class="not"><b>Notlar</b>{notlar}</div>' if notlar else ''}
{f'<div class="uyari">{e(h["uyari"])}</div>' if h['uyari'] else ''}
<div class="not">Bu sayfa e-Fatura XML'inden üretilmiş sade görünümdür (faturada görüntü şablonu yoktu).</div>
</body></html>"""
    return sayfa.encode("utf-8")


# ------------------------------------------------------------------ yardımcılar
def _t(el, yol: str) -> str:
    if el is None:
        return ""
    x = el.find(yol, _NS)
    return (x.text or "").strip() if x is not None and x.text else ""


def _n(el, yol: str) -> float:
    s = _t(el, yol)
    try:
        return float(s) if s else 0.0
    except ValueError:
        return 0.0


def _taraf(party) -> dict:
    if party is None:
        return {"ad": "", "vkn": ""}
    ad = _t(party, "cac:PartyName/cbc:Name")
    if not ad:
        kisi = party.find("cac:Person", _NS)
        if kisi is not None:
            ad = " ".join(x for x in (_t(kisi, "cbc:FirstName"), _t(kisi, "cbc:MiddleName"),
                                      _t(kisi, "cbc:FamilyName")) if x)
    vkn = ""
    for pid in party.findall("cac:PartyIdentification/cbc:ID", _NS):
        if (pid.get("schemeID") or "").upper() in ("VKN", "TCKN") and (pid.text or "").strip():
            vkn = pid.text.strip()
            break
    return {"ad": re.sub(r"\s+", " ", ad).strip(), "vkn": vkn}


# ------------------------------------------------------------------ ayrıştırma
def ayristir(xml_bytes: bytes) -> dict:
    """Tek bir UBL faturayı yöne bağlı olmayan ham sözlüğe çevirir (satıcı ve alıcı ikisi de)."""
    kok = ET.fromstring(xml_bytes)
    ad = kok.tag.split("}")[-1]
    if ad not in _KOK:
        raise ValueError(f"UBL fatura değil (kök: {ad})")

    para = _t(kok, "cbc:DocumentCurrencyCode").upper() or "TRY"
    kur = 1.0
    uyari = ""
    if para not in ("TRY", "TL"):
        kur = _n(kok, "cac:PricingExchangeRate/cbc:CalculationRate")
        if kur <= 0:
            uyari = f"döviz ({para}) faturasında kur yok — tutarlar {para} cinsinden, TL'ye çevrilmedi"
            kur = 1.0

    def tl(v):
        return round(v * kur, 2)

    # KDV dağılımı ve diğer vergiler (belge düzeyi)
    oranlar, diger_vergi = {}, 0.0
    for tt in kok.findall("cac:TaxTotal", _NS):
        for st in tt.findall("cac:TaxSubtotal", _NS):
            kod = _t(st, "cac:TaxCategory/cac:TaxScheme/cbc:TaxTypeCode")
            vergi = _n(st, "cbc:TaxAmount")
            if kod == KDV_KODU:
                oran = int(round(_n(st, "cbc:Percent")))
                o = oranlar.setdefault(oran, [0.0, 0.0])
                o[0] += _n(st, "cbc:TaxableAmount")
                o[1] += vergi
            else:
                diger_vergi += vergi

    tevkifat, tevkifat_kod = 0.0, ""
    for wt in kok.findall("cac:WithholdingTaxTotal", _NS):
        tevkifat += _n(wt, "cbc:TaxAmount")
        if not tevkifat_kod:
            tevkifat_kod = _t(wt, "cac:TaxSubtotal/cac:TaxCategory/cac:TaxScheme/cbc:TaxTypeCode")

    lmt = kok.find("cac:LegalMonetaryTotal", _NS)
    vergi_dahil = _n(lmt, "cbc:TaxInclusiveAmount")
    odenecek = _n(lmt, "cbc:PayableAmount")

    kalemler = [{"oran": o, "matrah": tl(m), "kdv": tl(k)}
                for o, (m, k) in sorted(oranlar.items(), key=lambda x: -x[0]) if m or k]
    # KDV matrahına girmeyen vergiler (konaklama vergisi vb.) ve yuvarlama: KDV dahil
    # toplamla aradaki fark. ÖTV gibi KDV matrahına giren vergiler zaten matrahın içindedir.
    ek_vergi = round(tl(vergi_dahil) - sum(x["matrah"] + x["kdv"] for x in kalemler), 2) if vergi_dahil else 0.0
    if abs(ek_vergi) < 0.01:
        ek_vergi = 0.0
    elif ek_vergi < 0:
        uyari = (uyari + " · " if uyari else "") + \
            f"KDV dağılımı vergiler dahil toplamı {abs(ek_vergi):,.2f} aşıyor — kontrol edin"
        ek_vergi = 0.0

    satirlar = []
    for ln in kok.findall("cac:InvoiceLine", _NS) + kok.findall("cac:CreditNoteLine", _NS):
        aciklama = _t(ln, "cac:Item/cbc:Name") or _t(ln, "cac:Item/cbc:Description") or _t(ln, "cbc:Note")
        oran = None
        for st in ln.findall("cac:TaxTotal/cac:TaxSubtotal", _NS):
            if _t(st, "cac:TaxCategory/cac:TaxScheme/cbc:TaxTypeCode") == KDV_KODU:
                oran = int(round(_n(st, "cbc:Percent")))
                break
        k = {"a": re.sub(r"\s+", " ", aciklama)[:120], "t": tl(_n(ln, "cbc:LineExtensionAmount"))}
        if oran is not None:
            k["o"] = oran
        if k["a"]:
            satirlar.append(k)

    tip = _t(kok, "cbc:InvoiceTypeCode").upper()
    tarih = _t(kok, "cbc:IssueDate")
    return {
        "fatura_no": _t(kok, "cbc:ID"),
        "uuid": _t(kok, "cbc:UUID"),
        "tarih": tarih if re.fullmatch(r"\d{4}-\d{2}-\d{2}", tarih) else "",
        "profil": _t(kok, "cbc:ProfileID").upper(),
        "fatura_tipi": tip,
        "para": para, "kur": kur,
        "satici": _taraf(kok.find("cac:AccountingSupplierParty/cac:Party", _NS)),
        "alici": _taraf(kok.find("cac:AccountingCustomerParty/cac:Party", _NS)),
        "kalemler": kalemler,
        "ek_vergi": ek_vergi,
        "diger_vergi": tl(diger_vergi),
        "tevkifat": tl(tevkifat), "tevkifat_kod": tevkifat_kod,
        "odenecek": tl(odenecek),
        "satirlar": satirlar[:30],
        "uyari": uyari,
    }


def kayda_cevir(h: dict, yon: str, dosya: str = "") -> dict:
    """ayristir() çıktısını isle_fatura'nın beklediği fatura kaydına çevirir.
    Alışta karşı taraf satıcı, satışta alıcıdır."""
    taraf = h["satici"] if yon == "alis" else h["alici"]
    kalemler = [dict(k) for k in h["kalemler"]]
    ek = h["ek_vergi"]
    tevk = h["tevkifat"]
    toplam = round(sum(k["matrah"] + k["kdv"] for k in kalemler) + ek - tevk, 2)
    uyarilar = [h["uyari"]] if h.get("uyari") else []
    if h["odenecek"] and abs(toplam - h["odenecek"]) > 0.05:
        uyarilar.append(f"XML ödenecek tutarı {h['odenecek']:,.2f}, hesaplanan {toplam:,.2f}")
    tip = h["fatura_tipi"]
    if "IADE" in tip:
        tur = "IADE"
    elif tevk > 0:
        tur = "TEVKIFAT"
    else:
        tur = "SATIS"
    eksik = []
    if not taraf["ad"]:
        eksik.append("cari")
    if not h["tarih"]:
        eksik.append("tarih")
    if not h["fatura_no"]:
        eksik.append("fatura_no")
    if not kalemler and not ek:
        eksik.append("tutar")
    return {
        "fatura_no": h["fatura_no"],
        "tarih": h["tarih"],
        "cari_ad": taraf["ad"],
        "vkn": taraf["vkn"],
        "tur": tur,
        # fatura_pdf ile aynı kural: KDV'siz, yalnız ek vergili belge faktoring/BSMV yolundan yazılır
        "senaryo": "TEMELFATURA" if (not kalemler and ek > 0) else "",
        "profil": h["profil"], "fatura_tipi": tip,
        "kalemler": kalemler,
        "toplam": toplam,
        "tevkifat": tevk, "tevkifat_kod": h["tevkifat_kod"],
        "ek_vergi": ek,
        "yon": yon, "dosya": dosya, "kaynak": "xml", "yz": [],
        "eksik": eksik,
        "kalem_aciklamalari": list(dict.fromkeys(s["a"] for s in h["satirlar"]))[:15],
        "kalem_detay": h["satirlar"],
        "uyari": (f"{h['fatura_no']}: " + "; ".join(uyarilar)) if uyarilar else "",
    }


def oku_bytes(data: bytes, yon: str, dosya: str = "") -> list:
    """Dosya içeriğindeki tüm faturalar (ZIP'te birden çok olabilir). Aynı UUID bir kez."""
    out, gorulen = [], set()
    for ad, x in xmlleri(data):
        h = ayristir(x)
        anahtar = h["uuid"] or h["fatura_no"] or ad
        if anahtar in gorulen:
            continue
        gorulen.add(anahtar)
        out.append(kayda_cevir(h, yon, dosya))
    return out


def dosya_oku(p: Path, yon: str) -> list:
    return oku_bytes(p.read_bytes(), yon, p.name)


def fatura_nolari(data: bytes) -> list:
    """Dizinleme için: içerikteki fatura numaraları (tam ayrıştırma yapmadan, hızlı)."""
    out = []
    for _, x in xmlleri(data):
        try:
            kok = ET.fromstring(x)
            no = _t(kok, "cbc:ID")
            if no:
                out.append(no.upper())
        except ET.ParseError:
            continue
    return out
