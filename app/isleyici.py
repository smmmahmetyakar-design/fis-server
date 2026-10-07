"""
İşleyici — ham belge verisini Fiş Aktarım Şablonu satırlarına çevirir.
Her satır sözlük (14 sütun karşılığı):
  fisno, fis_tarih(iso), fis_aciklama, hesap, evrak_no, evrak_tarih(iso),
  detay, borc, alacak, belge_turu(MF)

Şimdilik iskelet: banka/fatura/cek için temel akış. OCR/metin satırlarını
kural motoruyla eşleştirir, dengeli çift kayıt üretir. Kural dosyasındaki
banka hesabı / karşı hesap mantığı geliştirilecek.
"""
import io, re, itertools
from datetime import datetime
from app.kurallar import norm


# ----------------------------------------------------------------- tarih/sayı ayrıştırma
def _tarih_iso(s):
    if s is None:
        return ""
    if isinstance(s, datetime):
        return f"{s.year:04d}-{s.month:02d}-{s.day:02d}"
    s = str(s).strip()
    # "02.09.2026 14:33:12" / "2026-09-02T00:00:00" gibi saatli değerlerde saati at
    adaylar = [s]
    parca = re.split(r"[ T]", s, maxsplit=1)[0]
    if parca != s:
        adaylar.append(parca)
    for aday in adaylar:
        for fmt in ("%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d.%m.%y", "%d/%m/%y"):
            try:
                d = datetime.strptime(aday, fmt)
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
    # Türkçe binlik/ondalık: 1.234,56 -> 1234.56
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
def _sat(fisno, tarih, aciklama, hesap, borc, alacak, evrak_no="", detay="", kaynak="",
         para_birimi="", kur=None, doviz_tutar=None):
    return {
        "fisno": fisno, "fis_tarih": tarih, "fis_aciklama": aciklama,
        "hesap": hesap, "evrak_no": evrak_no, "evrak_tarih": tarih,
        "detay": detay or aciklama, "borc": round(borc, 2), "alacak": round(alacak, 2),
        "belge_turu": "MF", "kaynak": kaynak,
        "para_birimi": para_birimi, "kur": kur, "doviz_tutar": doviz_tutar,
    }


# ----------------------------------------------------------------- excel tablo -> satırlar
def _tablo_satirlari(hamlar):
    """Excel/PDF tablolarından (tarih, açıklama, tutar) çıkarımı.
    Banka başlık satırını (ilk 20 satırda) bulup sütunları eşler.
    İş Bankası, genel banka formatları desteklenir."""
    kayitlar = []
    for h in hamlar:
        for tablo in h.get("tablolar", []):
            if not tablo:
                continue
            bas_idx = None; harita = {}
            for i, row in enumerate(tablo[:20]):
                nrow = [norm(str(c)) for c in row]
                hh = {}
                for j, c in enumerate(nrow):
                    # tarih sütunu (öncelik: İşlem tarihi/Tarih-Saat, sonra Valör)
                    if c in ("TARIH SAAT", "TARIH", "ISLEM TARIHI") or c.startswith("TARIH"):
                        hh.setdefault("tarih", j)
                    elif c in ("VALOR", "VALOR TARIHI") and "tarih" not in hh:
                        hh["tarih_valor"] = j
                    # tutar (İş Bankası: "İşlem Tutarı", Denizbank: "Tutar (TL)")
                    elif "ISLEM TUTARI" in c or "TUTAR TL" in c or c == "TUTAR" or c == "MIKTAR" or c.startswith("TUTAR"):
                        hh["tutar"] = j
                    elif "GELIR" in c and "TUTAR" in c:
                        hh["borc"] = j    # Gelir Tutar = para girişi (bankaya gelen)
                    elif "GIDER" in c and "TUTAR" in c:
                        hh["alacak"] = j  # Gider Tutar = para çıkışı (bankadan giden)
                    elif "BORC" in c:
                        hh["borc"] = j
                    elif "ALACAK" in c:
                        hh["alacak"] = j
                    # açıklama
                    elif "ACIKLAMA" in c:
                        hh["aciklama"] = j
                    elif ("GONDER" in c or "UNVAN" in c or "FIRMA" in c) and "aciklama" not in hh:
                        hh["aciklama"] = j
                    # referans / evrak
                    elif "REFERANS" in c or "DEKONT" in c or "FIS NO" in c:
                        hh["referans"] = j
                    elif "ISLEM TIPI" in c or c == "ISLEM":
                        hh.setdefault("islem_tipi", j)
                # geçerli başlık: tarih + (tutar/borc/alacak) + açıklama
                if ("tarih" in hh or "tarih_valor" in hh) and ("tutar" in hh or "borc" in hh or "alacak" in hh):
                    harita = hh; bas_idx = i; break
            if bas_idx is None:
                continue

            tar_j = harita.get("tarih", harita.get("tarih_valor"))
            for row in tablo[bas_idx + 1:]:
                def g(key):
                    j = harita.get(key)
                    return row[j] if j is not None and j < len(row) else None
                # tarih: "08/06/2026-14:06:26" (İş Bank) veya "15.06.2026 10:46"
                # (Denizbank) gibi olabilir; saat kısmını kes.
                ham_tarih = row[tar_j] if tar_j is not None and tar_j < len(row) else None
                if isinstance(ham_tarih, str):
                    ht = ham_tarih.strip()
                    # tire veya boşluk ayracı varsa saat kısmını kes
                    if "-" in ht and len(ht) > 10:
                        ht = ht.split("-")[0].strip()
                    if " " in ht and len(ht) > 10:
                        ht = ht.split(" ")[0].strip()
                    ham_tarih = ht
                tarih = _tarih_iso(ham_tarih)
                acik = str(g("aciklama") or "").strip()
                # Bazı formatlarda açıklama bir sütun kaymış olabiliyor (başlık sütun 10, veri sütun 11)
                if not acik and "aciklama" in harita:
                    j_next = harita["aciklama"] + 1
                    if j_next < len(row):
                        acik = str(row[j_next] or "").strip()
                if "borc" in harita or "alacak" in harita:
                    borc = _sayi(g("borc")) or 0
                    alacak = _sayi(g("alacak")) or 0
                    tutar = borc - alacak
                else:
                    tutar = _sayi(g("tutar")) or 0
                if not tarih or not acik:
                    continue
                if tutar == 0:
                    continue
                ref = str(g("referans") or "").strip()
                islem_tipi = str(g("islem_tipi") or "").strip()
                kayitlar.append({"tarih": tarih, "aciklama": acik, "tutar": tutar,
                                 "referans": ref, "islem_tipi": islem_tipi,
                                 "dosya": h.get("dosya", "")})
    return kayitlar


def _ham_metin_satirlari(hamlar):
    """OCR/PDF ham metninden ekstre satırlarını çıkarır.
    Sağlam desen: satır tarihi ile başlarsa veri satırıdır. Bir sonraki tarih
    satırına kadar olan devam satırları açıklamanın parçasıdır.
    YKB tarzı çok satırlı ekstrelerde açıklama önceki satır(lar)a taşabilir —
    bu durumda önceki devam satırları da açıklamaya eklenir.
    Tutar tanıma: 'İşlem Tutarı' + 'Bakiye' varsa (2 tutar), sondan biri BAKİYE'dir;
    işlem tutarı sondan bir öncesidir.
    """
    kayitlar = []
    # dd/mm/yyyy, dd.mm.yyyy, dd-mm-yyyy tam tarih deseni (satır başında)
    tarih_bas_re = re.compile(r"^\s*(\d{1,2}[./\-]\d{1,2}[./\-](?:20)?\d{2})(?:\s|$)")
    tutar_re = re.compile(r"-?\d{1,3}(?:[.\s]\d{3})*[.,]\d{2}")

    for h in hamlar:
        metin = h.get("ham_metin", "")
        if not metin:
            continue
        hatlar = [ln.rstrip() for ln in metin.splitlines()]
        # veri satırlarının indekslerini bul (tarihle başlayanlar)
        veri_idx = []
        for i, ln in enumerate(hatlar):
            if tarih_bas_re.match(ln):
                veri_idx.append(i)
        if not veri_idx:
            continue

        for k, i in enumerate(veri_idx):
            hat = hatlar[i]
            tm = tarih_bas_re.match(hat)
            if not tm:
                continue
            tarih = _tarih_iso(tm.group(1))
            if not tarih:
                continue
            # bu satırdaki tutarları bul (TL gibi para birimi yazısını da temizle)
            sat_temiz = re.sub(r"\bTL\b|\bTRY\b", "", hat)
            tutarlar = tutar_re.findall(sat_temiz)
            if not tutarlar:
                continue
            # Sondan bir öncesi = işlem tutarı (2+ tutar varsa), tek varsa o
            if len(tutarlar) >= 2:
                islem_str = tutarlar[-2]
            else:
                islem_str = tutarlar[-1]
            tutar = _sayi(islem_str)
            if tutar is None or tutar == 0:
                continue

            # açıklama: satırın kendisinden tarih ve tutarları çıkar
            acik_parts = []
            # bu satırın tarih sonrası kısmı
            sonrasi = hat[tm.end():]
            # tüm tutarları temizle
            for t in tutarlar:
                sonrasi = sonrasi.replace(t, " ")
            sonrasi = re.sub(r"\bTL\b|\bTRY\b", " ", sonrasi)
            # HH:MM:SS saat kalıbını temizle (YKB'de tarihten sonra saat var)
            sonrasi = re.sub(r"\b\d{1,2}:\d{2}(?::\d{2})?\b", " ", sonrasi)
            sonrasi = re.sub(r"\s+", " ", sonrasi).strip(" \t-|:")
            if sonrasi:
                acik_parts.append(sonrasi)

            # Önceki devam satırları (bir önceki veri satırı ile bu arası)
            onceki_veri_idx = veri_idx[k - 1] if k > 0 else -1
            for j in range(onceki_veri_idx + 1, i):
                ln = hatlar[j].strip()
                if not ln:
                    continue
                # başlık/altbilgi filtrele
                if tarih_bas_re.match(ln):
                    continue
                # sadece işaretler/çizgiler
                if re.match(r"^[-\s|]+$", ln):
                    continue
                # başlık satırı (Tarih | Saat | ... | Bakiye)
                nl = norm(ln)
                if "TARIH" in nl and "BAKIYE" in nl:
                    continue
                if "HESAP HAREKETLERI" in nl or "MUSTERI ADI" in nl or "MUSTERI NUMARASI" in nl or "IBAN" in nl or "SUBE" in nl.split() or "HESAP ADI" in nl or "TARIH ARALIGI" in nl or "KULLANILABILIR BAKIYE" in nl:
                    continue
                if nl.startswith(("YAPI VE KREDI", "TICARET SICIL", "MERSIS", "ISLETMENIN")):
                    continue
                # tutar satırı ise geç (ama içinde harf varsa geçme, açıklama olabilir)
                if tutar_re.search(ln) and not re.search(r"[A-Za-zĞÜŞİÖÇğüşıöç]{4,}", ln):
                    continue
                # "- - - - Diğer Bekleyen İşlemler 0,00 TL -" gibi placeholder satırlar
                if "BEKLEYEN ISLEM" in nl or re.match(r"^-\s+-\s+-\s+-", ln):
                    continue
                acik_parts.append(ln)

            # Sonraki satır her zaman başka kaydın parçasıdır — dokunma.
            acik = " ".join(acik_parts).strip()
            # Sütun adı kalıntılarını temizle (her yerde)
            for kalinti in ["Para Gönder", "Yatırım Fonu", "Internet - Mobil", "Internet Mobil",
                            "Şube", "ATM", "EFT", "Havale", "Diğer"]:
                # başta ise
                acik = re.sub(rf"^{re.escape(kalinti)}\s+", "", acik, flags=re.IGNORECASE)
                # ortada ise (etrafında boşluk varsa)
                acik = re.sub(rf"\s+{re.escape(kalinti)}\s+", " ", acik, flags=re.IGNORECASE)
            acik = re.sub(r"\s+", " ", acik).strip(" -")

            kayitlar.append({"tarih": tarih, "aciklama": acik, "tutar": tutar,
                             "dosya": h.get("dosya", "")})
    return kayitlar


def _kayitlar(hamlar):
    """Önce tablo, tablo yoksa ham metin (OCR) satırları.
    Mükerrer kayıtları otomatik filtreler (aynı tarih + tutar + açıklama)."""
    k = _tablo_satirlari(hamlar)
    if not k:
        k = _ham_metin_satirlari(hamlar)
    if not k:
        return k
    # Mükerrer filtresi: aynı (tarih, tutar, açıklama[:40]) birden fazla varsa tekini tut
    gorulen = set()
    temiz = []
    mukerrer_sayisi = 0
    for kayit in k:
        anahtar = (
            kayit.get("tarih", ""),
            round(kayit.get("tutar", 0), 2),
            (kayit.get("aciklama", "") or "")[:40],
        )
        if anahtar in gorulen:
            mukerrer_sayisi += 1
            continue
        gorulen.add(anahtar)
        temiz.append(kayit)
    if mukerrer_sayisi > 0:
        print(f"[mükerrer filtre] {mukerrer_sayisi} mükerrer satır silindi, {len(temiz)} benzersiz kaldı")
    return temiz


# ----------------------------------------------------------------- BANKA
def _banka_hesabi_bul(hamlar, hesaplar_list, uyarilar):
    """Belge içeriğinden ve dosya adından banka+hesap numarası ipuçları çıkarır,
    mizandaki en uygun 102 alt hesabını seçer.
    hesaplar_list: [{"kod": "102.01.08", "ad": "DENİZBANK-354"}, ...]

    Tanıma öncelik sırası:
    1. IBAN (TR + 2 kontrol + 5-8. hane = banka kodu) — en güvenilir
    2. Metin/dosya adında banka adı (varyantlarıyla)
    3. Bulunamazsa mizandaki ilk 102 hesabı
    """
    import re as _re

    # IBAN → banka kodu → ana banka eşleşmesi
    IBAN_BANKA_KODU = {
        "00010": "ZIRAAT", "00012": "HALKBANK", "00015": "VAKIF",
        "00032": "TEB", "00046": "AKBANK", "00059": "TURK TICARET",
        "00062": "GARANTI", "00064": "ISBANK", "00067": "YAPI KREDI",
        "00092": "AKTIF", "00099": "ODEA", "00103": "FIBABANK",
        "00111": "QNB", "00123": "HSBC", "00124": "ALTERNATIF",
        "00125": "BURGAN", "00134": "DENIZBANK", "00135": "ANADOLUBANK",
        "00143": "ING", "00146": "ODEABANK", "00203": "ALBARAKA",
        "00206": "KUVEYT TURK", "00210": "TURKIYE FINANS",
    }

    # 1. Belge ipuçlarını topla — BAŞLIK ve İŞLEM ayrı tutulur
    # IBAN sadece başlık alanından alınır (işlem açıklamalarındaki karşı taraf IBAN'ları karıştırmasın)
    baslik_ipuclari = []  # sadece dosya adı + ilk 10 satır (başlık bölgesi)
    tum_ipuclari = []     # tüm metin (banka adı aramak için)
    for h in hamlar:
        dosya = h.get("dosya", "")
        baslik_ipuclari.append(dosya)
        tum_ipuclari.append(dosya)
        if h.get("ham_metin"):
            baslik_ipuclari.append(h["ham_metin"][:600])
            tum_ipuclari.append(h["ham_metin"][:600])
        for tab in h.get("tablolar", []):
            for i, r in enumerate(tab[:20]):
                for c in r:
                    if c:
                        s = str(c)
                        tum_ipuclari.append(s)
                        if i < 10:  # ilk 10 satır başlık
                            baslik_ipuclari.append(s)
    baslik_ham = " ".join(baslik_ipuclari)
    baslik_norm = norm(baslik_ham)
    metin_ham = " ".join(tum_ipuclari)
    metin = norm(metin_ham)

    # 2. Mizanda 102 ile başlayan tüm hesaplar
    banka_hesaplari = [(h["kod"], norm(h["ad"])) for h in hesaplar_list if h["kod"].startswith("102")]
    if not banka_hesaplari:
        uyarilar.append("Mizanda 102 hesabı yok, 102.01.001 varsayıldı")
        return "102.01.001"

    # 3. Banka adı listesi + varyantları
    banka_adaylari = [
        ("DENIZBANK", ["DENIZ"]),
        ("ISBANK", ["IS BANK", "TURKIYE IS", "ISBANKASI", "IS BANKASI"]),
        ("ZIRAAT", ["TC ZIRAAT", "ZIRAAT BANK"]),
        ("VAKIF", ["VAKIFBANK", "VAKIF BANK", "VAKIFLAR"]),
        ("GARANTI", ["GARANTI BBVA", "GARANTI BANK"]),
        ("YAPI KREDI", ["YKB", "YAPIKREDI", "YAPI VE KREDI"]),
        ("HALKBANK", ["HALK BANK", "TC HALK", "HALKBANKA"]),
        ("AKBANK", []),
        ("QNB", ["FINANSBANK", "FINANS BANK"]),
        ("ING", ["ING BANK"]),
        ("HSBC", []),
        ("TEB", []),
        ("SEKERBANK", []),
        ("KUVEYT TURK", ["KUVEYTTURK"]),
        ("TURKIYE FINANS", ["TURKIYEFINANS"]),
        ("ALBARAKA", []),
    ]

    # 4. Hesap numarası ipuçlarını çıkar (dosya adında "354", metinde "9290 - 60363919 - 354" gibi)
    hesap_no_ipuclari = set()
    for h in hamlar:
        for parca in _re.findall(r'\d{3,}', h.get("dosya", "")):
            hesap_no_ipuclari.add(parca)
    for m in _re.finditer(r'(\d{3,8})(?!\d)', metin[:3000]):
        hesap_no_ipuclari.add(m.group(1))

    # 5. IBAN'ı yakala — SADECE BAŞLIK bölgesinden (işlem açıklamalarındaki
    # karşı taraf IBAN'ları yanlış banka tespitine yol açmasın)
    tespit_banka = None
    iban_bulundu = None
    # IBAN: TR + 2 kontrol + 5 banka kodu — ama boşluklu yazılabilir (TR14 0006 4000...)
    # Bu yüzden 4+1 olarak yakala (boşluk arada olabilir)
    iban_re = _re.compile(r'TR\s*(\d{2})\s*(\d{4})\s*(\d)')
    for m in iban_re.finditer(baslik_ham):
        banka_kod = m.group(2) + m.group(3)  # 4+1 = 5 haneli banka kodu
        if banka_kod in IBAN_BANKA_KODU:
            tespit_banka = IBAN_BANKA_KODU[banka_kod]
            iban_bulundu = m.group(0)
            break

    # 6. IBAN yoksa BAŞLIK metninde banka adı ara (tüm metin değil)
    if not tespit_banka:
        for ana, varyantlar in banka_adaylari:
            if ana in baslik_norm or any(v in baslik_norm for v in varyantlar):
                tespit_banka = ana; break

    # 7. Uygun hesabı seç
    if tespit_banka:
        # Tespit edilen banka için varyantlar
        varyant_listesi = [tespit_banka]
        for ana, vv in banka_adaylari:
            if ana == tespit_banka:
                varyant_listesi.extend(vv)
                break

        # Hesap adında tespit edilen bankanın adı veya varyantı geçenler
        eslesenler = [(k, ad) for k, ad in banka_hesaplari
                      if any(v in ad for v in varyant_listesi)]

        if eslesenler:
            # Birden fazla eşleşme varsa hesap numarası ipuçları ile daralt
            if len(eslesenler) > 1 and hesap_no_ipuclari:
                for k, ad in eslesenler:
                    hesap_rakamlari = set(_re.findall(r'\d{3,}', ad))
                    if hesap_rakamlari & hesap_no_ipuclari:
                        return k
            return eslesenler[0][0]

        # Banka tespit edildi ama mizanda bu bankaya ait 102 alt hesap yok
        uyarilar.append(f"Belge {tespit_banka} bankasından ama mizanda uygun 102 hesabı yok, varsayılan kullanıldı")

    # 8. Hiçbir banka tespit edilemedi — ilk 102 alt hesabını al
    for k, ad in banka_hesaplari:
        if k.count(".") >= 2:
            uyarilar.append(f"Banka tanınamadı, mizandaki ilk 102 alt hesabı ({k}) kullanıldı")
            return k
    uyarilar.append("Uygun 102 hesabı bulunamadı, 102.01.001 varsayıldı")
    return "102.01.001"


def isle_banka(hamlar, km, fis0):
    """
    Banka dökümü -> muhasebe fişi.
    Aynı banka + aynı tarih = tek fiş.
    Dövizli hesaplar: TCMB kuruyla TL'ye çevrilir.
    """
    uyarilar = []
    hesaplar_list = [{"kod": k, "ad": a} for k, a in km.hesaplar]

    # Her belge için ayrı banka hesabı belirle
    dosya_to_hesap = {}
    for h in hamlar:
        dosya = h.get("dosya", "")
        hesap = _banka_hesabi_bul([h], hesaplar_list, uyarilar)
        dosya_to_hesap[dosya] = hesap

    # Döviz tespiti: banka hesabının adından veya belge içeriğinden
    DOVIZ_KODLARI = {"EUR": "EUR", "USD": "USD", "GBP": "GBP", "CHF": "CHF",
                     "EURO": "EUR", "DOLAR": "USD", "STERLIN": "GBP"}
    dosya_to_doviz = {}
    for h in hamlar:
        dosya = h.get("dosya", "")
        banka_hesap = dosya_to_hesap.get(dosya, "")
        hesap_ad = (km.hesap_adi(banka_hesap) or "").upper()
        doviz = None
        # Hesap adında döviz kodu var mı? (örn. "GARANTİ EUR")
        for anahtar, kod in DOVIZ_KODLARI.items():
            if anahtar in hesap_ad:
                doviz = kod; break
        # Dosya adında da ara
        if not doviz:
            for anahtar, kod in DOVIZ_KODLARI.items():
                if anahtar in dosya.upper():
                    doviz = kod; break
        # Belge başlığında ara
        if not doviz:
            icerik = (h.get("ham_metin") or "")[:500].upper()
            for tab in h.get("tablolar", []):
                for r in tab[:8]:
                    for c in r:
                        if c: icerik += " " + str(c).upper()
            for anahtar, kod in DOVIZ_KODLARI.items():
                if anahtar in icerik:
                    doviz = kod; break
        dosya_to_doviz[dosya] = doviz

    # Dövizli dosyalar varsa TCMB kurlarını hazırla
    kur_getir_fn = None
    dovizli_tarihler = {}  # cache: (doviz, tarih) -> kur
    if any(v for v in dosya_to_doviz.values()):
        try:
            from app.tcmb import kur_getir
            kur_getir_fn = kur_getir
        except ImportError:
            uyarilar.append("TCMB modülü yüklenemedi, döviz kurları çevrilmedi")

    kayitlar = _kayitlar(hamlar)
    if not kayitlar:
        for h in hamlar:
            if h.get("ham_metin"):
                uyarilar.append(f"{h.get('dosya')}: tablo çıkarılamadı, ham metin var")
        return [], uyarilar

    fisler = []
    fis = fis0

    # Aynı bankanın aynı tarihteki hareketlerini TEK FİŞTE topla
    from collections import defaultdict
    gruplar = defaultdict(list)
    for k in kayitlar:
        banka_hesap = dosya_to_hesap.get(k.get("dosya", ""), "102.01.001")
        tarih = k.get("tarih", "") or ""
        gruplar[(banka_hesap, tarih)].append(k)

    for (banka_hesap, tarih), grup_kayitlar in sorted(gruplar.items(), key=lambda x: (x[0][1], x[0][0])):
        fisno = f"{fis:05d}"
        banka_ad = km.hesap_adi(banka_hesap) or "BANKA"

        for k in grup_kayitlar:
            karsi, kaynak = km.eslestir(k["aciklama"])
            if not karsi:
                karsi = "198.01.001"
                uyarilar.append(f"{k['aciklama'][:30]}: hesap eşleşmedi")
            tutar = abs(k["tutar"])
            evno = k.get("referans", "")

            # Döviz kontrolü
            dosya = k.get("dosya", "")
            doviz = dosya_to_doviz.get(dosya)
            para_birimi = doviz or ""
            kur_degeri = None
            doviz_tutar = None

            if doviz and kur_getir_fn:
                cache_key = (doviz, k["tarih"])
                if cache_key not in dovizli_tarihler:
                    try:
                        dovizli_tarihler[cache_key] = kur_getir_fn(doviz, k["tarih"])
                    except Exception:
                        dovizli_tarihler[cache_key] = None
                kur_degeri = dovizli_tarihler.get(cache_key)
                if kur_degeri:
                    doviz_tutar = tutar
                    tutar = round(tutar * kur_degeri, 2)
                else:
                    uyarilar.append(f"{doviz} kuru bulunamadı ({k['tarih']}), tutar çevrilmedi")

            if k["tutar"] >= 0:
                fisler.append(_sat(fisno, k["tarih"], banka_ad, banka_hesap, tutar, 0,
                    evrak_no=evno, detay=k["aciklama"], kaynak="banka",
                    para_birimi=para_birimi, kur=kur_degeri, doviz_tutar=doviz_tutar))
                fisler.append(_sat(fisno, k["tarih"], banka_ad, karsi, 0, tutar,
                    evrak_no=evno, detay=k["aciklama"], kaynak=kaynak,
                    para_birimi=para_birimi, kur=kur_degeri, doviz_tutar=doviz_tutar))
            else:
                fisler.append(_sat(fisno, k["tarih"], banka_ad, karsi, tutar, 0,
                    evrak_no=evno, detay=k["aciklama"], kaynak=kaynak,
                    para_birimi=para_birimi, kur=kur_degeri, doviz_tutar=doviz_tutar))
                fisler.append(_sat(fisno, k["tarih"], banka_ad, banka_hesap, 0, tutar,
                    evrak_no=evno, detay=k["aciklama"], kaynak="banka",
                    para_birimi=para_birimi, kur=kur_degeri, doviz_tutar=doviz_tutar))
        fis += 1
    return fisler, uyarilar



# ----------------------------------------------------------------- FATURA
def _elogo_fatura_satirlari(hamlar, yon="alis", teshis=None):
    """
    eLogo entegratör Excel formatındaki fatura listesini okur.
    Sütun anahtarları (Türkçe duyarsız):
      FATURA NO, FATURA TARIHI, GONDERICI ADI (alış) / ALICI ADI (satış),
      SENARYO, TUR, TOPLAM TUTAR, KDV TOPLAMI, KDV %10, KDV %20,
      KDV %10 MATRAH, KDV %20 MATRAH, TEVKIFAT TOPLAMI, EK VERGILER
    Döndürür: liste of fatura dict.
    """
    faturalar = []
    for h in hamlar:
        for tablo in h.get("tablolar", []):
            if not tablo:
                continue
            # başlık satırını bul: içinde "FATURA NO" geçen ilk satır
            bas_idx = None; sut = {}
            for i, row in enumerate(tablo[:5]):
                nrow = [norm(str(c).replace("\n"," ")) for c in row]
                for j, c in enumerate(nrow):
                    if c == "FATURA NO" or "FATURA" in c and ("NO" in c or "NUMARASI" in c):
                        sut["fatura_no"] = j
                    elif c == "FATURA TARIHI" or ("FATURA" in c and "TARIH" in c):
                        sut["tarih"] = j
                    elif c in ("GONDERICI ADI", "ALICI ADI", "GONDERICI UNVAN", "ALICI UNVAN",
                                "UNVAN", "GONDEREN ADI", "GONDEREN UNVAN",
                                "SATICI ADI", "SATICI UNVAN", "SATICI UNVANI",
                                "TEDARIKCI ADI", "TEDARIKCI UNVAN",
                                "MUSTERI ADI", "MUSTERI UNVAN"):
                        sut["cari_ad"] = j
                    elif c == "ACIKLAMA" and "cari_ad" not in sut:
                        # AÇIKLAMA sütunu sadece başka cari adı sütunu yoksa fallback olarak kullanılır.
                        # GİB formatında bu sütun yazıyla tutarı içerir (örn. "Yalnız …TL")
                        sut["cari_ad"] = j
                        sut["_cari_aciklamadan"] = True
                    elif c == "SENARYO":
                        sut["senaryo"] = j
                    elif c in ("TUR", "FATURA TURU"):
                        sut["tur"] = j
                    elif c == "TOPLAM TUTAR" or c == "TOPLAM":
                        sut["toplam"] = j
                    elif c == "KDV TOPLAMI":
                        sut["kdv_top"] = j
                    # İnteraktif Vergi Dairesi e-Arşiv listesi: her satırda tek oran —
                    # "KDV Oranı" + "Matrah Tutar" + "KDV Tutar" (ÖİV ayrı sütunlarda)
                    elif c in ("KDV ORANI", "KDV ORAN", "KDV YUZDESI"):
                        sut["kdv_orani"] = j
                    elif c in ("MATRAH TUTAR", "MATRAH TUTARI", "KDV MATRAH TUTARI"):
                        sut["matrah_tutar"] = j
                    elif c in ("KDV TUTAR", "KDV TUTARI") and "kdv_tutar" not in sut:
                        sut["kdv_tutar"] = j
                    elif "VKN" in c or "TCKN" in c:
                        sut["vkn"] = j
                    # KDV sütunları: hem "KDV 20" hem "%20'lik KDV" formatı
                    elif c == "KDV 1" or ("1LIK" in c.replace(" ","") and "KDV" in c and "MATRAH" not in c):
                        sut["kdv_1"] = j
                    elif c == "KDV 8" or ("8LIK" in c.replace("'","").replace(" ","") and "KDV" in c and "MATRAH" not in c):
                        sut["kdv_8"] = j
                    elif c == "KDV 10" or ("10LUK" in c.replace("'","").replace(" ","") and "KDV" in c and "MATRAH" not in c) or ("10LIK" in c.replace("'","").replace(" ","") and "KDV" in c and "MATRAH" not in c):
                        sut["kdv_10"] = j
                    elif c == "KDV 18" or ("18LIK" in c.replace("'","").replace(" ","") and "KDV" in c and "MATRAH" not in c):
                        sut["kdv_18"] = j
                    elif c == "KDV 20" or ("20LIK" in c.replace("'","").replace(" ","") and "KDV" in c and "MATRAH" not in c):
                        sut["kdv_20"] = j
                    elif c == "KDV 1 MATRAH" or ("1LIK" in c.replace("'","").replace(" ","") and "MATRAH" in c):
                        sut["mat_1"] = j
                    elif c == "KDV 8 MATRAH" or ("8LIK" in c.replace("'","").replace(" ","") and "MATRAH" in c):
                        sut["mat_8"] = j
                    elif c == "KDV 10 MATRAH" or ("10LUK" in c.replace("'","").replace(" ","") and "MATRAH" in c) or ("10LIK" in c.replace("'","").replace(" ","") and "MATRAH" in c):
                        sut["mat_10"] = j
                    elif c == "KDV 18 MATRAH" or ("18LIK" in c.replace("'","").replace(" ","") and "MATRAH" in c):
                        sut["mat_18"] = j
                    elif c == "KDV 20 MATRAH" or ("20LIK" in c.replace("'","").replace(" ","") and "MATRAH" in c):
                        sut["mat_20"] = j
                    elif "TEVKIFAT" in c and "KOD" in c:
                        sut.setdefault("tevkifat_kod", j)
                    elif "TEVKIFAT" in c and "ORAN" in c:
                        sut.setdefault("tevkifat_oran", j)     # 7/10, %70 ya da 0,7
                    elif "TEVKIFAT" in c and "MATRAH" not in c and "tevkifat" not in sut:
                        sut["tevkifat"] = j                     # tevkifat TUTARI (KDV'nin alıcının ödeyeceği kısmı)
                    elif c in ("FATURA TIPI", "FATURA TIP") and "tur" not in sut:
                        sut["tur"] = j
                    elif c in ("EK VERGILER", "EK VERGI", "DIGER VERGILER", "OIV", "OIV TUTARI", "OIV TUTAR",
                               "OTV", "OTV TUTARI", "OTV TUTAR") \
                            or "OZEL ILETISIM" in c or "OZEL TUKETIM" in c:
                        sut["ek_vergi"] = j
                    elif "OZEL" in c and "MATRAH" in c:
                        sut["ozel_matrah"] = j
                # Genişletilmiş tanıma — GİB portalı ve farklı entegratör listeleri
                # başlıkları farklı yazıyor ("Düzenlenme Tarihi", "Alıcı Unvanı/Adı Soyadı").
                if "fatura_no" in sut and "tarih" not in sut:
                    for j, c in enumerate(nrow):
                        if "TARIH" in c and not any(x in c for x in (
                                "VADE", "ODEME", "ALINMA", "KAYIT", "GONDERIM", "ONAY", "IPTAL", "YANIT", "SEVK", "IRSALIYE")):
                            sut["tarih"] = j
                            break
                if "fatura_no" in sut:
                    # Tutar sütunları ÖNCELİĞE göre seçilir, sütun sırasına göre değil.
                    # "Mal Hizmet Toplam Tutarı" / "Vergiler Hariç" KDV HARİÇ matrahtır —
                    # toplam sanılırsa cariye matrah yazılır, KDV satırı kaybolur.
                    def _ilk(kosul):
                        return next((j for j, c in enumerate(nrow) if kosul(c)), None)
                    mat_j = _ilk(lambda c: ("MAL HIZMET" in c and ("TOPLAM" in c or "TUTAR" in c))
                                 or "VERGILER HARIC" in c or c in ("MATRAH", "TOPLAM MATRAH", "KDV MATRAHI"))
                    top_j = None
                    for kosul in (lambda c: "VERGILER DAHIL" in c,
                                  lambda c: "ODENECEK" in c,
                                  lambda c: c in ("TOPLAM TUTAR", "TOPLAM", "GENEL TOPLAM", "FATURA TUTARI"),
                                  lambda c: "TOPLAM" in c and "TUTAR" in c
                                  and not any(x in c for x in ("KDV", "MATRAH", "MAL HIZMET", "HARIC", "VERGI"))):
                        top_j = _ilk(kosul)
                        if top_j is not None and top_j != mat_j:
                            break
                        top_j = None
                    if top_j is not None:
                        sut["toplam"] = top_j
                    elif sut.get("toplam") == mat_j:
                        sut.pop("toplam", None)
                    if mat_j is not None:
                        sut["mat_toplam"] = mat_j
                if "fatura_no" in sut and "kdv_top" not in sut and "kdv_orani" not in sut:
                    j = next((j for j, c in enumerate(nrow) if c in ("TOPLAM KDV", "HESAPLANAN KDV", "KDV TUTARI", "KDV")
                              or ("KDV" in c and "TOPLAM" in c and "MATRAH" not in c)), None)
                    if j is None:
                        # GİB portal / entegratör listelerinde KDV "Vergiler Toplamı" adıyla gelebilir
                        # (ÖİV gibi ek vergi de içerebilir: oran türetilemezse fatura PDF ister)
                        j = next((j for j, c in enumerate(nrow) if c in (
                            "VERGILER TOPLAMI", "VERGI TOPLAMI", "TOPLAM VERGI", "TOPLAM VERGILER",
                            "HESAPLANAN VERGI", "HESAPLANAN VERGILER", "VERGI TUTARI", "VERGILER")), None)
                    if j is not None and j not in {v for k, v in sut.items() if k.startswith(("kdv_", "mat_"))}:
                        sut["kdv_top"] = j
                if "fatura_no" in sut and ("cari_ad" not in sut or sut.get("_cari_aciklamadan")):
                    # alışta karşı taraf faturayı gönderen, satışta alan
                    tercih = (("GONDERICI", "GONDEREN", "SATICI", "TEDARIKCI", "DUZENLEYEN")
                              if yon == "alis" else ("ALICI", "MUSTERI"))
                    secilen = next((j for j, c in enumerate(nrow)
                                    if any(t in c for t in tercih) and ("UNVAN" in c or "ADI" in c or "AD SOYAD" in c)
                                    and "VKN" not in c and "TCKN" not in c), None)
                    if secilen is None and "cari_ad" not in sut:
                        secilen = next((j for j, c in enumerate(nrow) if "UNVAN" in c and "VKN" not in c), None)
                    if secilen is not None:
                        sut["cari_ad"] = secilen
                        sut.pop("_cari_aciklamadan", None)
                if "fatura_no" in sut:
                    # Listede hem gönderen hem alan sütunu varsa cari YÖNE göre seçilir:
                    # alışta karşı taraf gönderen (satıcı), satışta alan (müşteri).
                    # Eskiden sağdaki sütun kazanıyordu; alış listesinde cari = firmanın kendisi oluyordu.
                    gonderen = ("GONDERICI", "GONDEREN", "SATICI", "TEDARIKCI", "DUZENLEYEN")
                    alan = ("ALICI", "MUSTERI")
                    def _taraf(c, kelimeler_):
                        return (any(t in c for t in kelimeler_) and ("UNVAN" in c or "ADI" in c or "AD SOYAD" in c)
                                and "VKN" not in c and "TCKN" not in c)
                    g_j = next((j for j, c in enumerate(nrow) if _taraf(c, gonderen)), None)
                    a_j = next((j for j, c in enumerate(nrow) if _taraf(c, alan)), None)
                    istenen = g_j if yon == "alis" else a_j
                    if istenen is not None:
                        sut["cari_ad"] = istenen
                        sut.pop("_cari_aciklamadan", None)
                # GİB formatında fatura_no+tarih+cari_ad yeterli; eLogo'da fatura_no+tarih+cari_ad
                if ("fatura_no" in sut or "tarih" in sut) and ("cari_ad" in sut or "toplam" in sut):
                    bas_idx = i; break
            if bas_idx is None:
                continue
            if teshis is not None:
                # karşılaştırma penceresi için: listenin başlıkları ve hangilerinin tanındığı
                teshis.setdefault("tablolar", []).append({
                    "dosya": h.get("dosya", ""),
                    "basliklar": [str(c).replace("\n", " ").strip() for c in tablo[bas_idx] if c not in (None, "")],
                    "taninan": {k: str(tablo[bas_idx][v]).replace("\n", " ").strip()
                                for k, v in sut.items() if not k.startswith("_") and v < len(tablo[bas_idx])},
                })

            def g(row, key):
                j = sut.get(key)
                return row[j] if j is not None and j < len(row) else None

            # Listede hiç KDV bilgisi (oran/matrah/KDV toplamı sütunu) yoksa KDV
            # dağılımı bilinmiyor demektir — PDF'ten tamamlanması gerekir.
            kdv_sutun_var = any(k in sut for k in (
                "kdv_1", "kdv_8", "kdv_10", "kdv_18", "kdv_20",
                "mat_1", "mat_8", "mat_10", "mat_18", "mat_20", "kdv_top", "kdv_orani")) \
                or ("mat_toplam" in sut and "toplam" in sut)

            for row in tablo[bas_idx + 1:]:
                fno = g(row, "fatura_no")
                if not fno or not str(fno).strip():
                    continue
                # Listenin altındaki not / açıklama satırları ("E-Fatura içeriğine konu olan
                # kargo gönderi detay bilgilerine ... adresinden ulaşabilirsiniz") fatura değildir.
                # Fatura no boşluksuz, kısa bir koddur (GIB2026000000448).
                fno_s = str(fno).strip()
                if len(fno_s) > 30 or len(fno_s.split()) > 2 or "http" in fno_s.lower() or "www." in fno_s.lower():
                    continue
                eksik = []
                cari = str(g(row, "cari_ad") or "").strip()
                # "Yalnız ... TL" gibi yazıyla tutar içeren açıklama cari ad değil
                if cari and re.match(r"(?i)yaln[ıi]z\s", cari):
                    cari = ""
                if not cari:
                    # cari adı boşsa fatura_no kullan (en azından bir referans olsun)
                    cari = str(fno).strip()
                    eksik.append("cari")
                # kalemler: sıfır olmayan her KDV oranı = bir kalem
                kalemler = []
                for oran, mat_k, kdv_k in [
                    (1, "mat_1", "kdv_1"), (8, "mat_8", "kdv_8"),
                    (10, "mat_10", "kdv_10"), (18, "mat_18", "kdv_18"),
                    (20, "mat_20", "kdv_20"),
                ]:
                    matrah = _sayi(g(row, mat_k)) or 0
                    kdv = _sayi(g(row, kdv_k)) or 0
                    if matrah > 0 or kdv > 0:
                        kalemler.append({"oran": oran, "matrah": round(matrah, 2), "kdv": round(kdv, 2)})
                # Oran sütunlu liste (satır başına tek oran): KDV Oranı + Matrah + KDV Tutar
                if not kalemler and "kdv_orani" in sut:
                    oran_ = _sayi(g(row, "kdv_orani"))
                    mat_ = _sayi(g(row, "matrah_tutar")) or 0
                    kdv_ = _sayi(g(row, "kdv_tutar")) or 0
                    if oran_ is not None and (mat_ > 0 or kdv_ > 0):
                        oran_ = int(round(oran_ * 100)) if 0 < oran_ < 1 else int(round(oran_))   # 0,20 -> 20
                        if not mat_ and kdv_ and oran_:
                            mat_ = round(kdv_ * 100 / oran_, 2)
                        kalemler.append({"oran": oran_, "matrah": round(mat_, 2), "kdv": round(kdv_, 2)})

                toplam = _sayi(g(row, "toplam")) or 0
                tevkifat = _sayi(g(row, "tevkifat")) or 0
                tevkifat_kod = str(g(row, "tevkifat_kod") or "").strip()
                ek_vergi = _sayi(g(row, "ek_vergi")) or 0
                tur = norm(str(g(row, "tur") or "SATIS")) or "SATIS"
                if "TEVKIFAT" in tur:
                    tur = "TEVKIFAT"        # "Tevkifatlı", "TEVKİFAT" …
                elif "IADE" in tur:
                    tur = "IADE"
                tevkifat_oran = _tevkifat_orani(g(row, "tevkifat_oran"))
                senaryo = str(g(row, "senaryo") or "").strip().upper()

                # toplam sütunu yoksa (satış Excel'i) kalemlerden hesapla
                if toplam == 0 and kalemler:
                    toplam = round(sum(k["matrah"] + k["kdv"] for k in kalemler), 2)
                    # tevkifat varsa toplama eklenmesi lazım (KDV zaten hesaplı ama tevkifat toplam düşüldüğünde alıcının ödeyeceği)
                    # ödenecek = toplam - tevkifat_kdv (alıcı tevkifatı direkt vergiye öder)
                # ek vergiler (BSMV vb.) hala 0 olabilir; eğer toplam 0 hala ise
                if toplam == 0 and ek_vergi > 0:
                    toplam = ek_vergi

                # Oran sütunu yok ama "KDV Toplamı" var: oranı aritmetikle türet
                # (KDV dahil toplam ve KDV tutarı tek bir oranı işaret ediyorsa)
                kdv_top = _sayi(g(row, "kdv_top")) or 0
                # Matrah sütunu varsa KDV = toplam − matrah (− ek vergi)
                mat_toplam = _sayi(g(row, "mat_toplam")) or 0
                if not kalemler and not kdv_top and mat_toplam > 0 and toplam > mat_toplam:
                    kdv_top = round(toplam - mat_toplam - ek_vergi, 2)
                if not toplam and mat_toplam > 0 and not kdv_sutun_var:
                    toplam = mat_toplam          # yalnız matrah var: KDV bilinmiyor, PDF tamamlar
                    eksik.append("kdv")
                if not kalemler and kdv_top > 0 and toplam > 0:
                    t = _yk_oran_turet(round(toplam - ek_vergi, 2), kdv_top)
                    if not t and tevkifat > 0:
                        # toplam ödenecek tutar olabilir (KDV dahil − tevkifat)
                        t = _yk_oran_turet(round(toplam + tevkifat - ek_vergi, 2), kdv_top)
                    if t:
                        kalemler.append({"oran": t[0], "matrah": t[1], "kdv": t[2]})
                    else:
                        eksik.append("kdv")     # çok oranlı ya da tevkifatlı — PDF gerekir

                # eğer hiç kalem yoksa ama toplam varsa (faktoring/BSMV): tek kalem KDV=0
                if not kalemler and toplam > 0:
                    kalemler.append({"oran": 0, "matrah": round(toplam - ek_vergi, 2), "kdv": 0})
                    # KDV sütunları var ama bu satırda hepsi BOŞ: KDV "sıfır" değil, bilinmiyor
                    kdv_hucreleri = [g(row, k) for k in ("kdv_1", "kdv_8", "kdv_10", "kdv_18", "kdv_20",
                                                         "mat_1", "mat_8", "mat_10", "mat_18", "mat_20", "kdv_top")
                                     if k in sut]
                    bos = all(v is None or str(v).strip() in ("", "-") for v in kdv_hucreleri)
                    if (not kdv_sutun_var or bos) and "kdv" not in eksik and senaryo != "TEMELFATURA":
                        eksik.append("kdv")

                tarih = _tarih_iso(g(row, "tarih"))
                if not tarih:
                    eksik.append("tarih")

                # Tevkifat tutarı yok ama oranı var: tevkifat = KDV × oran
                kdv_sum = round(sum(k["kdv"] for k in kalemler), 2)
                if not tevkifat and tevkifat_oran and kdv_sum > 0:
                    tevkifat = round(kdv_sum * tevkifat_oran, 2)
                if tevkifat > 0 and kdv_sum and tevkifat > kdv_sum + 0.05:
                    tevkifat = 0      # KDV'den büyük tevkifat olmaz — sütun yanlış okunmuş
                if tevkifat > 0:
                    tur = "TEVKIFAT"

                faturalar.append({
                    "fatura_no": str(fno).strip(),
                    "tarih": tarih,
                    "cari_ad": cari,
                    "tur": tur, "senaryo": senaryo,
                    "toplam": round(toplam, 2),
                    "kalemler": kalemler,
                    "tevkifat": round(tevkifat, 2),
                    "tevkifat_kod": tevkifat_kod,
                    "ek_vergi": round(ek_vergi, 2),
                    "yon": yon, "dosya": h.get("dosya", ""),
                    "eksik": eksik, "kaynak": "liste",
                })
    return faturalar


def _kdv_hesabi_ad(kod_ad_list, oran, yon="alis"):
    """191 (alış indirilecek) veya 391 (satış hesaplanan) hesabını orana göre bulur.
    "SORUMLU/IADE/ITHAL" hariç. Bulamazsa boş döner."""
    if oran <= 0:
        return ""
    prefix = "191" if yon == "alis" else "391"
    YASAK = ("SORUMLU", "IADE", "ITHAL", "VAZGEC", "TEVKIF")
    adaylar = []
    # Temmuz 2023'te %18 -> %20, %8 -> %10 oldu; eski açılmış hesaplar hâlâ
    # "İndirilecek KDV %18" adıyla kullanılıyor. Önce güncel oran, yoksa eski adı.
    for oran_str in [str(oran)] + {20: ["18"], 10: ["8"]}.get(oran, []):
        for h in kod_ad_list:
            kod = h["kod"]
            if not kod.startswith(prefix):
                continue
            nad = norm(h["ad"])
            if any(y in nad for y in YASAK):
                continue
            nad_sp = nad.replace(" ", "")
            # tam oran eşleşmesi: %20 ile %10 karışmasın
            if re.search(rf"%{oran_str}(?!\d)", nad_sp) or re.search(rf"(?<!\d){oran_str}(?!\d)", nad_sp):
                adaylar.append((kod.count("."), kod))
        if adaylar:
            break
    alt = alt_hesap_kodlari([(h["kod"], h["ad"]) for h in kod_ad_list])
    if adaylar:
        # kayıt atılabilir (alt) hesabı tercih et
        yapraklar = [a for a in adaylar if a[1] in alt]
        (yapraklar or adaylar).sort(reverse=True)
        return (yapraklar or adaylar)[0][1]
    # Adında hiç oran yazmayan KDV hesabı (çoğu mizanda tek "İndirilecek KDV"):
    # yalnızca TEK böyle hesap varsa onu kullan.
    genel = []
    for h in kod_ad_list:
        kod = h["kod"]
        if kod not in alt or not kod.startswith(prefix):
            continue
        nad = norm(h["ad"])
        if any(y in nad for y in YASAK):
            continue
        if re.search(r"%\s*\d|(?<!\d)(1|8|10|18|20)(?!\d)", nad.replace(".", " ")):
            continue        # adında başka bir oran yazıyor
        genel.append(kod)
    if len(genel) == 1:      # birden çoksa hangisi olduğu belli değil — tahmin etme
        return genel[0]
    return ""


TEVKIFAT_ORANLARI = (0.2, 0.3, 0.4, 0.5, 0.7, 0.9, 1.0)   # 2/10 … 9/10, tam tevkifat


def _tevkifat_orani(v):
    """'7/10', '%70', '70', '0,7' -> 0.7; okunamazsa 0."""
    if v is None:
        return 0.0
    s = str(v).strip().replace("%", "").replace(" ", "")
    m = re.match(r"^(\d+)\s*/\s*(\d+)$", s)
    if m and int(m.group(2)):
        r = int(m.group(1)) / int(m.group(2))
    else:
        x = _sayi(s)
        if not x:
            return 0.0
        r = x / 100 if x > 1 else x
    return r if 0 < r <= 1 else 0.0


def _tevkifat_tahmin(kalemler, toplam, ek_vergi=0.0):
    """Listede tevkifat yok ama toplam KDV dahil tutardan KDV'nin standart bir
    tevkifat payı (2/10 … 9/10, tamamı ya da bir oranın KDV'sinin payı) kadar azsa
    o fark tevkifattır (entegratör listesinde 'Toplam' = ödenecek tutar). Döner: tutar ya da 0."""
    if not kalemler or not toplam:
        return 0.0
    kdv_dahil = sum(k["matrah"] + k["kdv"] for k in kalemler) + (ek_vergi or 0)
    fark = round(kdv_dahil - toplam, 2)
    kdv = sum(k["kdv"] for k in kalemler)
    if fark <= 0.05 or kdv <= 0 or fark > kdv + 0.05:
        return 0.0
    tabanlar = [kdv] + [k["kdv"] for k in kalemler if k["kdv"] > 0]
    for t in tabanlar:
        for r in TEVKIFAT_ORANLARI:
            if abs(t * r - fark) <= 0.05 + 0.0005 * t:
                return fark
    return 0.0


def _tevkifat_hesabi(kod_ad_list, yon="satis", alt_kodlar=None):
    """Tevkifat hesabı.
    Alışta (alıcı sorumlu): tevkif edilen KDV'yi alıcı 2 No.lu beyanla öder —
    360 ÖDENECEK VERGİ VE FONLAR altındaki 'sorumlu/tevkifat KDV' hesabı (alacak).
    Satışta: 391 HESAPLANAN KDV altındaki 'tevkifat' hesabı (borç; hesaplanan KDV'yi azaltır)."""
    if yon == "satis":
        prefix = ("391",); oncelik = (("TEVKIF",),)
    else:
        prefix = ("360",); oncelik = (("TEVKIF",), ("SORUMLU",), ("2NOLU", "2NO", "KDV2"), ("KDV",))
    hesaplar = [h for h in kod_ad_list if h["kod"].startswith(prefix)
                and (alt_kodlar is None or h["kod"] in alt_kodlar)]
    for anahtarlar in oncelik:
        adaylar = []
        for h in hesaplar:
            nad = norm(h["ad"]); bit = nad.replace(" ", "").replace(".", "")
            if any(a in nad or a in bit for a in anahtarlar):
                if yon == "alis" and anahtarlar == ("KDV",) and ("GELIR" in nad or "STOPAJ" in nad or "DAMGA" in nad):
                    continue
                adaylar.append((h["kod"].count("."), h["kod"]))
        if adaylar:
            adaylar.sort(key=lambda x: (-x[0], x[1]))
            return adaylar[0][1]
    return ""


def _bsmv_hesabi(kod_ad_list):
    """Faktoring/finans giderleri için hesap. Öncelik: 780.02 (finansman/faktoring
    gideri), sonra BSMV içeren 780, en son herhangi bir 780."""
    # 1) 780.02 ile başlayan (en detaylı alt hesap)
    adaylar_02 = [h for h in kod_ad_list if h["kod"].startswith("780.02")]
    if adaylar_02:
        adaylar_02.sort(key=lambda h: -h["kod"].count("."))
        return adaylar_02[0]["kod"]
    # 2) BSMV / banka sigorta içeren 780
    for h in kod_ad_list:
        kod = h["kod"]; nad = norm(h["ad"])
        if kod.startswith("780") and ("BSMV" in nad or ("BANKA" in nad and "SIGORTA" in nad)):
            return kod
    # 3) herhangi bir 780
    for h in kod_ad_list:
        if h["kod"].startswith("780"):
            return h["kod"]
    return "780.02.001"


# ------------------------------------------------- YAZAR KASA (perakende) FİŞİ
# Market/esnaf yazar kasa fişi e-faturadan yapıca farklıdır:
#   * KDV oranı departman adının yanındadır ("GIDA %01", "TEMİZLİK %10") ve
#     OCR tam burayı bozar ("GIDA “01", "TEMİZLİK y4Q", "TOPKDV ağ").
#   * Tutarlar "*" ile başlar, OCR sık sık virgülü düşürür ("*525 00").
#   * Şahıs firmalarında unvan eki (LTD/A.Ş.) yoktur; bir taramada 4 fiş varsa
#     unvandan bölme denemesi hepsini tek fiş sanır.
# Bu yüzden:
#   1) Fişler, her fişte bir kez geçen TOPKDV satırı ÇAPA alınarak bölünür.
#   2) KDV oranı OCR'dan değil TOPLAM/TOPKDV ARİTMETİĞİNDEN türetilir — oran
#      okunamasa bile sonuç doğru çıkar.
#   3) POS slipleri (banka onay fişleri) TOPKDV içermediği için ayrı belge
#      sayılmaz; alan okuma fiş gövdesiyle (TOPLAM satırına kadar) sınırlıdır.

_YK_TOPKDV_RE = re.compile(r"\bTOP\s*K[DO][VU]\b")
_YK_TOPLAM_RE = re.compile(r"\bTOPLAM\b")
_YK_TARIH_RE = re.compile(r"(\d{1,2})\s*[./-]\s*(\d{1,2})\s*[./-]\s*(20\d{2})")
_YK_FISNO_RE = re.compile(r"\b(?:FIS|TIS|FLS|FIG)\s*NO\b\D{0,8}(\d{1,6})")
_YK_VKN_RE = re.compile(r"\b(\d{10,11})\b")
_YK_ORAN_RE = re.compile(r"[%‰]\s*(\d{1,2})\b")
# Para: "160,00" / "1.234,56" / "525 ,00" / "525 00" (OCR virgülü düşürmüş)
_YK_PARA_RE = re.compile(r"\d{1,3}(?:\.\d{3})*\s*[,.]\s*\d{2}\b|\b\d+\s+\d{2}\b")
# Önceki fişin/POS slibinin bittiğini gösteren satırlar — blok başı ararken dur.
# ("TEŞEKKÜRLER" bilerek yok: fiş başlığında da geçiyor.)
_YK_KUYRUK_RE = re.compile(
    r"EK[UÜ]\s*NO|\bZ\s*NO\b|DEGERI\s*YOKTUR|NUSHASI|SAKLAYINIZ|ONAY\s*KODU|"
    r"ISYERI\s*NO|TERM\s*NO|BATCH|APP\s*LABEL|SIRA\s*NO|TEMASSIZ|KART\s*HAMILI|"
    r"POS\s*NO|\bTOPLAM\b|TOP\s*K[DO][VU]|\bODEAL\b|\bKREDI\b|\bNAKIT\b|\bTUTAR\b")
# Tek başına dursa bile asla atılmayacak fiş alanları (ödeme şekli, toplamlar)
_YK_KORU_RE = re.compile(
    r"\bKREDI\b|\bKART\b|\bNAKIT\b|\bPESIN\b|\bTOPLAM\b|TOP\s*K[DO][VU]|\bKDV\b|"
    r"\bODEAL\b|\bTUTAR\b|\bTARIH\b|\bSAAT\b|\bFIS\b")
# Türkiye'de yürürlükteki KDV oranları (eski 8/18 geçmiş fişler için korunuyor)
_YK_KDV_ORANLARI = (1, 10, 20, 8, 18, 0)


def _yk_para(hat: str) -> list:
    """Yazar kasa satırındaki tutarları OCR toleranslı okur."""
    out = []
    temiz = re.sub(r"[*«»¥#£]", " ", hat)
    for m in _YK_PARA_RE.finditer(temiz):
        t = re.sub(r"\s*([,.])\s*", r"\1", m.group(0))   # "525 ,00" -> "525,00"
        t = re.sub(r"^(\d+)\s+(\d{2})$", r"\1,\2", t)    # "525 00"  -> "525,00"
        v = _sayi(t)
        if v is not None:
            out.append(v)
    return out


def _yk_oran_turet(toplam: float, kdv_tutar: float, ipucu=None):
    """KDV dahil toplam + KDV tutarından oranı/matrahı türetir.
    OCR oranı bozsa bile ("%01", "y4Q") aritmetik doğruyu verir.
    Döner: (oran, matrah, kdv) veya None."""
    if not toplam or toplam <= 0 or kdv_tutar is None:
        return None
    adaylar = list(_YK_KDV_ORANLARI)
    if ipucu in adaylar:            # OCR'dan okunan oranı önce dene
        adaylar.remove(ipucu)
        adaylar.insert(0, ipucu)
    for r in adaylar:
        matrah = round(toplam / (1 + r / 100.0), 2)
        kdv = round(toplam - matrah, 2)
        if abs(kdv - kdv_tutar) <= 0.02:
            return r, matrah, kdv
    return None


_YK_ADRES_RE = re.compile(
    r"\bVD\b|\bV\s*\.?\s*D\b|VERGI\s*DAIRESI|MAHALLE|\bMAH\b|\bMH\b|\bSOK\b|"
    r"\bSK\b|\bCD\b|\bCAD\b|\bNO\s*:|TEL\s*:|\bCEP\b|\bTARIH\b|\bSAAT\b|"
    r"\bFIS\s*NO\b|\bZ\s*NO\b|EK[UÜ]\s*NO")


def _yk_departmanlar(blok: list, ci: int) -> list:
    """TOPKDV'den önceki departman satırlarını (tutar + varsa OCR oran ipucu)
    toplar. Ör: 'GIDA %01 *160,00' -> {tutar: 160.00, ipucu: 1}"""
    out = []
    for i in range(ci):
        h = blok[i]
        if _YK_ADRES_RE.search(norm(h)) or _YK_TARIH_RE.search(h):
            continue
        p = _yk_para(h)
        if not p:
            continue
        m = _YK_ORAN_RE.search(h)
        ipucu = int(m.group(1)) if m and int(m.group(1)) in _YK_KDV_ORANLARI else None
        out.append({"tutar": max(p), "ipucu": ipucu})
    return out


def _yk_kalemler(departmanlar: list, toplam: float, kdv_tutar: float):
    """Çok KDV oranlı fişte her departmanın oranını çözer.
    Market fişlerinde tek fişte %1 gıda + %20 temizlik birlikte olabilir;
    TOPKDV ikisinin toplamıdır. OCR oranları güvenilmez olduğu için oranlar
    departman tutarlarından ARANIR: hangi oran dizilimi TOPKDV'yi tutturuyorsa
    o alınır, eşitlik birden çok dizilimde sağlanıyorsa OCR ipuçlarına en çok
    uyan tercih edilir. Döner: kalem listesi veya None."""
    tutarlar = [d["tutar"] for d in departmanlar]
    n = len(tutarlar)
    if not n or n > 5 or kdv_tutar is None:
        return None
    # Departman tutarları fiş toplamını tutmuyorsa satırları yanlış okumuşuz
    if abs(sum(tutarlar) - toplam) > 0.05:
        return None

    best = None
    for kombin in itertools.product(_YK_KDV_ORANLARI, repeat=n):
        k = 0.0
        for t, r in zip(tutarlar, kombin):
            k += t - round(t / (1 + r / 100.0), 2)
        if abs(round(k, 2) - kdv_tutar) <= 0.02 + 0.01 * n:
            skor = sum(1 for d, r in zip(departmanlar, kombin) if d["ipucu"] == r)
            # eşit skorda daha az farklı oran kullanan dizilim yalındır
            anahtar = (skor, -len(set(kombin)))
            if best is None or anahtar > best[0]:
                best = (anahtar, kombin)
    if not best:
        return None

    birlesik = {}
    for t, r in zip(tutarlar, best[1]):
        matrah = round(t / (1 + r / 100.0), 2)
        kdv = round(t - matrah, 2)
        if r not in birlesik:
            birlesik[r] = {"oran": r, "matrah": 0.0, "kdv": 0.0}
        birlesik[r]["matrah"] = round(birlesik[r]["matrah"] + matrah, 2)
        birlesik[r]["kdv"] = round(birlesik[r]["kdv"] + kdv, 2)
    return sorted(birlesik.values(), key=lambda x: x["oran"])


def _yk_firma_adi(govde: list) -> str:
    """Fiş başlığından firma unvanını çıkarır.
    Başlık yapısı: [unvan] [unvan 2. satır] [adres] [VD] [TARİH] ...
    Adres/VD/TARİH satırına kadar olan satırlar unvandır; aradaki OCR artıkları
    ('2C 56551789', 'np AVa000108999') harf oranına bakılarak elenir."""
    basliklar = []
    for h in govde[:7]:
        t = re.sub(r"\s+", " ", h).strip(" .,:;*-_|")
        nt = norm(t)
        if _YK_TARIH_RE.search(t) or _YK_ADRES_RE.search(nt):
            break
        harf = sum(1 for c in t if c.isalpha())
        if len(t) < 5 or harf < 4 or harf < len(t) * 0.5:
            continue                    # OCR artığı satır
        basliklar.append(t)
        if len(basliklar) == 2:         # unvan en fazla 2 satıra yayılır
            break
    if basliklar:
        return " ".join(basliklar)
    for h in govde[:3]:
        t = re.sub(r"\s+", " ", h).strip()
        if len(t) > 4:
            return t
    return ""


def _yk_tarih(govde: list) -> tuple:
    """Fiş tarihini okur. Gün 31'den büyükse (OCR '0'ı '6'/'8' okumuş olabilir)
    baştaki rakamı '0' yapıp düzeltmeyi dener. Döner: (iso_tarih, uyari)."""
    for h in govde:
        m = _YK_TARIH_RE.search(h)
        if not m:
            continue
        g, a, y = m.group(1), m.group(2), m.group(3)
        iso = _tarih_iso(f"{int(g):02d}.{int(a):02d}.{y}") if int(g) <= 31 and int(a) <= 12 else ""
        if iso:
            return iso, ""
        # gün bozuk: "64/10/2026" -> "04/10/2026"
        if len(g) == 2 and int(g) > 31:
            d = _tarih_iso(f"0{g[1]}.{int(a):02d}.{y}")
            if d:
                return d, f"tarih OCR'da '{g}/{a}/{y}' okundu, '{d}' varsayıldı — kontrol edin"
    return "", ""


def _yk_parse(blok: list, dosya: str = "") -> dict | None:
    """Tek bir yazar kasa fişi bloğunu alanlarına ayırır."""
    nblok = [norm(h) for h in blok]
    ci = next((i for i, h in enumerate(nblok) if _YK_TOPKDV_RE.search(h)), None)
    if ci is None:
        return None
    # TOPLAM satırı çapadan hemen sonra gelir
    ti = next((i for i in range(ci + 1, min(len(blok), ci + 4))
               if _YK_TOPLAM_RE.search(nblok[i])), None)
    govde = blok[:(ti if ti is not None else ci) + 1]

    # KDV tutarı: çapa satırında, yoksa bir sonraki satırda ("TOPKDV ağ" / "2,08")
    kdv_tutar = None
    for i in (ci, ci + 1):
        if i >= len(blok) or (ti is not None and i == ti):
            continue
        p = _yk_para(blok[i])
        if p:
            kdv_tutar = p[0]
            break

    toplam = 0.0
    if ti is not None:
        p = _yk_para(blok[ti])
        if p:
            toplam = max(p)
    if not toplam:          # TOPLAM okunamadı — departman satırlarının toplamı
        for i in range(ci):
            if _YK_ORAN_RE.search(blok[i]) or re.search(r"[*«]", blok[i]):
                p = _yk_para(blok[i])
                if p:
                    toplam = max(toplam, max(p))
    if toplam <= 0:
        return None

    # KDV oranı: önce OCR ipucu, sonra aritmetik doğrulama
    ipucu = None
    for i in range(ci):
        m = _YK_ORAN_RE.search(blok[i])
        if m and int(m.group(1)) in _YK_KDV_ORANLARI:
            ipucu = int(m.group(1))
            break

    uyari = ""
    # Önce departman satırlarından çöz (tek fişte birden çok KDV oranı olabilir),
    # olmazsa fiş toplamından tek oran türet.
    kalemler = _yk_kalemler(_yk_departmanlar(blok, ci), toplam, kdv_tutar) or []
    if not kalemler:
        turetim = _yk_oran_turet(toplam, kdv_tutar, ipucu)
        if turetim:
            oran, matrah, kdv = turetim
            kalemler = [{"oran": oran, "matrah": matrah, "kdv": kdv}]
        else:
            uyari = (f"{dosya}: KDV oranı TOPLAM/TOPKDV'den türetilemedi "
                     f"(toplam={toplam}, topkdv={kdv_tutar}) — KDV'yi elle ayırın")

    firma = _yk_firma_adi(govde)
    if not firma:
        return None

    tarih, t_uyari = _yk_tarih(govde)
    if t_uyari:
        uyari = (uyari + " | " if uyari else "") + f"{firma[:25]}: {t_uyari}"

    fno = ""
    for h in govde:
        m = _YK_FISNO_RE.search(norm(h))
        if m:
            fno = m.group(1)
            break

    vkn = ""
    for h in govde:
        nh = norm(h)
        if re.search(r"\bVD\b|\bV\s*\.?\s*D\b|VERGI\s*DAIRESI", nh):
            m = _YK_VKN_RE.search(nh)
            if m:
                vkn = m.group(1)
                break

    # Ödeme şekli: gövdeden hemen sonraki birkaç satırda
    odeme = ""
    for h in blok[len(govde):len(govde) + 6]:
        nh = norm(h)
        if "NAKIT" in nh or "PESIN" in nh:
            odeme = "NAKIT"
            break
        # ÖDEAL/banka POS satırı da kartla ödemeyi gösterir
        if "KREDI" in nh or "KART" in nh or "ODEAL" in nh:
            odeme = "KREDI KARTI"
            break

    return {
        "fatura_no": fno,
        "tarih": tarih,
        "cari_ad": firma,
        "tur": "SATIS", "senaryo": "PERAKENDE",
        "toplam": round(toplam, 2),
        "kalemler": kalemler,
        "tevkifat": 0.0, "tevkifat_kod": "",
        "ek_vergi": 0.0,
        "yon": "alis", "dosya": dosya,
        "vkn": vkn, "odeme": odeme, "belge_turu": "YAZARKASA",
        "uyari": uyari,
    }


def _yk_satirlari(metin: str) -> list:
    """Ham OCR metnini satırlara ayırırken fiş ayırıcı artıklarını atar.
    Fişleri ayıran '****...****' çizgisi ve POS slip logoları OCR'da tek başına
    duran anlamsız satırlara dönüşür ('EREARRA REARS ATRIA', 'ödeâ,',
    '2C 56551789'); bunlar bir sonraki fişin unvanı sanılmasın diye elenir.
    Ölçüt (hepsi birden): öncesi ve sonrası boş + hiç rakam yok + ':' yok +
    fiş alanı anahtar kelimesi yok. Gerçek başlık satırları ardışık bir öbek
    hâlinde gelir, fiş alanları ise ya rakam ya anahtar kelime taşır."""
    ham = metin.splitlines()
    out = []
    for i, h in enumerate(ham):
        t = h.rstrip()
        if not t.strip():
            continue
        tek_basina = ((i == 0 or not ham[i - 1].strip())
                      and (i == len(ham) - 1 or not ham[i + 1].strip()))
        if (tek_basina and ":" not in t
                and not any(c.isdigit() for c in t)
                and not _YK_KORU_RE.search(norm(t))):
            continue
        out.append(t)
    return out


def _yazarkasa_fisleri(metin: str, dosya: str = "") -> list:
    """Taramadaki yazar kasa fişlerini TOPKDV çapalarına göre böler ve okur."""
    hatlar = _yk_satirlari(metin)
    capalar = [i for i, h in enumerate(hatlar) if _YK_TOPKDV_RE.search(norm(h))]
    if not capalar:
        return []

    baslar = []
    for k, a in enumerate(capalar):
        onceki = capalar[k - 1] if k > 0 else -1
        # çapadan geriye bu fişin TARİH satırını bul
        ti = a
        for j in range(a - 1, onceki, -1):
            if _YK_TARIH_RE.search(hatlar[j]):
                ti = j
                break
        # TARİH'ten geriye başlık/adres satırlarını topla; önceki fişin
        # kuyruğuna (TOPLAM/ÖDEAL/EKÜ NO/POS slip) çarpınca dur
        s = ti
        while (s - 1 > onceki and ti - s < 10
               and not _YK_KUYRUK_RE.search(norm(hatlar[s - 1]))):
            s -= 1
        baslar.append(s)

    fisler = []
    for k, s in enumerate(baslar):
        son = baslar[k + 1] if k + 1 < len(baslar) else len(hatlar)
        f = _yk_parse(hatlar[s:son], dosya)
        if f:
            fisler.append(f)
    return fisler


def _ocr_fatura_ayikla(metin: str, dosya: str = "") -> list:
    """
    OCR ham metninden fatura(lar)ı çıkarır.
    Bir görüntüde birden çok fiş olabilir. Hem perakende hem e-fatura fotoğrafları.
    OCR gürültülü olduğu için parser gevşek: eksik alan olabilir, kullanıcı düzeltir.
    """
    if not metin:
        return []
    hatlar = [h.rstrip() for h in metin.splitlines() if h.strip()]
    if not hatlar:
        return []

    # Önce yazar kasa (perakende) fişi dene — TOPKDV çapası varsa bu yoldur.
    yk = _yazarkasa_fisleri(metin, dosya)
    if yk:
        return yk

    # Fiş başlangıç adayları: "E-Arşiv", "E-Fatura", firma adı olabilecek satırlar
    # (büyük harfli, sonu "A.Ş." veya "LTD" ile biten)
    fis_bas = []
    for i, hat in enumerate(hatlar):
        nh = norm(hat)
        # Fatura başlığı işaretleri
        if any(k in nh for k in ["E ARSIV FATURA", "E FATURA", "E-ARSIV", "E-FATURA", "TICARIFATURA", "TEMELFATURA"]):
            fis_bas.append(i); continue
        # Firma unvanı (A.Ş., LTD, ŞTİ ile biten büyük harfli satır)
        if re.search(r"\b(A\.?\s?[SŞ]\.?|LTD\.?|[SŞ]T[IİI]\.?|LIMITED|LİMİTED)\b", hat, re.IGNORECASE):
            if len(hat) > 10 and (not fis_bas or i - fis_bas[-1] > 5):
                fis_bas.append(i)

    if not fis_bas:
        fis_bas = [0]

    fisler = []
    for k, bas in enumerate(fis_bas):
        son = fis_bas[k + 1] if k + 1 < len(fis_bas) else len(hatlar)
        bolum = hatlar[bas:son]
        f = _parse_perakende(bolum, dosya)
        if f:
            fisler.append(f)
    return fisler


def _parse_perakende(hatlar: list, dosya: str = "") -> dict | None:
    """Bir fiş bloğundan alanları çıkarır. OCR gürültüsüne dayanıklı."""
    if not hatlar:
        return None
    metin = "\n".join(hatlar)

    # Firma adı: ilk 6 satırda A.Ş./LTD/ŞTİ bitişli, en uzun olanı seç
    firma = ""
    for h in hatlar[:6]:
        if re.search(r"\b(A\.?\s?[SŞ]\.?|LTD\.?|[SŞ]T[IİI]\.?|LIMITED|LİMİTED|MAGAZALAR|MARKETCILIK|DISCOUNT|TEKSTIL)\b", h, re.IGNORECASE):
            if len(h.strip()) > len(firma):
                firma = h.strip()
    if not firma:
        # ilk anlamlı büyük harfli satır
        for h in hatlar[:5]:
            if len(h.strip()) > 8 and h.strip().upper() == h.strip():
                firma = h.strip(); break
    if not firma and hatlar:
        firma = hatlar[0].strip()
    firma = re.sub(r"\s+", " ", firma).strip()
    # "E-Arşiv Fatura" gibi başlıkları filtrele
    if norm(firma).startswith(("E ARSIV", "E FATURA")):
        # sonraki satırda firma
        for h in hatlar[1:5]:
            if len(h.strip()) > 8:
                firma = h.strip(); break

    # Tarih: dd.mm.yyyy veya dd/mm/yyyy
    tarih = ""
    tm = re.search(r"(\d{1,2}[./]\d{1,2}[./]20\d{2})", metin)
    if tm:
        tarih = _tarih_iso(tm.group(1))

    # Fiş/Fatura no
    fno = ""
    fnm = re.search(r"(?:FATURA\s*NO|FIS\s*NO|FI[SŞ]\s*NO)\s*:?\s*(\S+)", norm(metin))
    if fnm:
        fno = fnm.group(1)[:24]  # kırp

    # KDV MATRAH tablosu — gevşek yaklaşım
    kalemler = _kdv_tablosu_ayikla(metin)

    # Toplam: "ÖDENECEK" veya "TOPLAM" satırından en büyük sayı
    toplam = 0.0
    for h in hatlar:
        nh = norm(h)
        if any(k in nh for k in ["ODENECEK", "TOPLAM", "TOPKDV", "GENEL"]):
            # Satırdaki tüm sayıları bul, en büyüğünü al
            sayilar = re.findall(r"(\d+[.,]\d{2})", h)
            for s in sayilar:
                v = _sayi(s)
                if v and v > toplam and v > 5:
                    toplam = v

    # Toplam bulunamadıysa kalemlerden
    if toplam == 0 and kalemler:
        toplam = round(sum(k["matrah"] + k["kdv"] for k in kalemler), 2)

    # Geçerlilik: firma VE (toplam > 0 VEYA en az bir kalem)
    if not firma or (toplam == 0 and not kalemler):
        return None

    return {
        "fatura_no": fno,
        "tarih": tarih,
        "cari_ad": firma,
        "tur": "SATIS", "senaryo": "PERAKENDE",
        "toplam": round(toplam, 2),
        "kalemler": kalemler,
        "tevkifat": 0.0, "tevkifat_kod": "",
        "ek_vergi": 0.0,
        "yon": "alis", "dosya": dosya,
    }


def _kdv_tablosu_ayikla(metin: str) -> list:
    """
    Perakende fişteki KDV matrah tablosundan kalemleri çıkarır.
    OCR gürültüsüne dayanıklı: "*", "«", "»" gibi karakterleri temizler.
    """
    kalemler = []
    for hat in metin.splitlines():
        # Temizleme: gürültü karakterlerini boşluğa çevir
        temiz = re.sub(r"[*«»¥#£]", " ", hat).strip()
        # Oran (% ile veya %'siz) + en az 2 ondalıklı sayı
        # Ör: "20  91.67  18.33  110.00" veya "%1 86.14 0.86 87.00"
        m = re.match(r"^[\s%]*(\d{1,2})\s+(\d+[.,]\d{2})\s+(\d+[.,]\d{2})(?:\s+(\d+[.,]\d{2}))?", temiz)
        if m:
            oran = int(m.group(1))
            if oran not in (0, 1, 8, 10, 18, 20):
                continue
            a = _sayi(m.group(2)) or 0
            b = _sayi(m.group(3)) or 0
            c = _sayi(m.group(4)) if m.group(4) else None
            # Üç sayı varsa: matrah, kdv, dahil (en küçük ilk ikisi)
            # İki sayı varsa: matrah, kdv
            if c is not None:
                # matrah = en küçük ondalıklı büyük sayı, kdv = ortadaki
                # Genelde sıra: matrah, kdv_tutar, kdv_dahil
                matrah = a; kdv = b
            else:
                matrah = a; kdv = b
            if matrah > 0:
                kalemler.append({"oran": oran, "matrah": round(matrah, 2),
                                 "kdv": round(kdv, 2)})
    # Aynı oran birden çok kez varsa en büyük matrahı tut
    tek = {}
    for k in kalemler:
        if k["oran"] not in tek or k["matrah"] > tek[k["oran"]]["matrah"]:
            tek[k["oran"]] = k
    return list(tek.values())


def _ham_metin_faturalar(hamlar, yon="alis"):
    """Her ham belgenin metninden perakende/e-fatura ayıkla."""
    out = []
    for h in hamlar:
        metin = h.get("ham_metin", "")
        if not metin:
            continue
        dosya = h.get("dosya", "")
        for f in _ocr_fatura_ayikla(metin, dosya):
            f["yon"] = yon
            out.append(f)
    return out


# ------------------------------------------------------- PDF tekil e-Fatura (gerçek fatura PDF'i)
# eLogo entegratör Excel listesi yoksa ve dosya market-fişi FOTOĞRAFI değil,
# doğrudan bir e-Fatura/e-Arşiv PDF'iyse burası devreye girer. Her tedarikçi
# kendi şablonunu kullandığından ürün/hizmet kalem tablosu yerine TÜM
# şablonlarda ortak üst bilgi alanları hedeflenir: Fatura No/ID, Tarih,
# "Mal/Hizmet Toplam Tutarı", "Hesaplanan KDV(...)", "Ödenecek Tutar" /
# "Vergiler Dahil Toplam Tutar". Bu alanlar öncelikle sayfa sonundaki özet
# TABLOSUNDAN (pdfplumber tablo çıkarımı), yoksa ham metinden regex ile okunur.
# (Kaynak: GitHub'ın daha önce ayrı geliştirilen fatura modülünden — gerçek
# tedarikçi PDF şablonlarıyla doğrulanmış — bu sunucudaki mizan-tabanlı hesap
# eşleştirme mantığına uyacak şekilde uyarlanmıştır.)
_PDF_FATURA_ID_RE = re.compile(r"Fatura\s*ID\s*:?\s*([A-Za-z0-9]+)", re.IGNORECASE)
_PDF_FATURA_NO_RE = re.compile(r"Fatura\s*(?:No|Numaras[ıi])\s*:?\s*([A-Za-z0-9\-]+)", re.IGNORECASE)
_PDF_FATURA_TARIHI_RE = re.compile(r"Fatura\s*Tarih[i]?\s*:?\s*(\d{1,2})[\s./-]+(\d{1,2})[\s./-]+(\d{2,4})", re.IGNORECASE)
_PDF_TARIH_FALLBACK_RE = re.compile(r"(?:^|\n)\s*Tarih\s*:?\s*(\d{1,2})[\s./-]+(\d{1,2})[\s./-]+(\d{2,4})", re.IGNORECASE | re.MULTILINE)
_PDF_TARIH_ANY_RE = re.compile(r"\b(\d{1,2})[.\-](\d{1,2})[.\-](\d{4})\b")
_PDF_IADE_RE = re.compile(r"Fatura\s*Tipi\s*:?\s*IADE", re.IGNORECASE)
_PDF_ODENECEK_RE = re.compile(r"(?:ÖDENECEK\s+(?:TUTAR|TOPLAM)|Ödenecek\s+Tutar)\s*:?\s*([\d.,]+)", re.IGNORECASE)
_PDF_MAL_HIZMET_RE = re.compile(r"Mal\s*/?\s*Hizmet\s*Toplam\s*Tutar[ıi]\s*:?\s*([\d.,]+)\s*(?:TL|TRY)?", re.IGNORECASE)
_PDF_HESAPLANAN_KDV_METIN_RE = re.compile(r"Hesaplanan\s+KDV[^(%\n]*[(%][^)]*?(\d+)[^)]*\)?\s*:?\s*([\d.,]+)\s*(?:TL|TRY)", re.IGNORECASE)
_PDF_KDV_ORAN_RE = re.compile(r"%\s*(\d+)")
_PDF_UNVAN_RE = re.compile(r"(ŞİRKETİ|SIRKETI|A\.\Ş\.|A\.S\.|LTD\.|ŞTİ\.|STI\.|ANONIM)", re.IGNORECASE)
_PDF_GONDERICI_DUR_RE = re.compile(
    r"VKN|TCKN|Vergi\s*Dairesi|Vergi\s*No|Tel:|Faks|Web\s*Sitesi|E-Posta|E-posta|"
    r"Mah\.|Mahallesi|\bMh\.|Sokak|Sok\.|\bSk\.|Cadde|Cad\.|\bCd\.|No\s*:\s*\d", re.IGNORECASE)
_PDF_METADATA_ALAN_RE = re.compile(r"Özelleştirme|Senaryo|Fatura\s*ID|Fatura\s*Tarih|Fatura\s*Tipi", re.IGNORECASE)
_PDF_TUTAR_HUCRE_RE = re.compile(r"^[\d.,]+\s*(?:TL|TRY)?$")
_PDF_SAYIN_ETIKET_RE = re.compile(r"^\s*(?:SAY[İI]N|AL[İI]C[İI])\s*:?\s*(.*)$", re.IGNORECASE)
_PDF_VERGI_SATIRI_RE = re.compile(
    r"^(.{3,60}?)\s*%\s*(\d+)\s*\(Matrah\w*\s*:?\s*([\d.,]+)\s*(?:TRY|TL)?\s*\)\s*([\d.,]+)\s*$",
    re.IGNORECASE | re.MULTILINE)


def _pdf_alici_adi_bul(metin: str) -> str:
    """SATIŞ faturasında 'SAYIN'/'ALICI' etiketinin altındaki (veya aynı
    satırındaki) alıcı unvanını çıkarır."""
    satirlar_all = [s.strip() for s in metin.splitlines()]
    for i, s in enumerate(satirlar_all):
        m = _PDF_SAYIN_ETIKET_RE.match(s)
        if not m:
            continue
        aday = []
        inline = m.group(1).strip()
        if inline:
            aday.append(inline)
        for s2 in satirlar_all[i + 1:i + 12]:
            if len(aday) >= 2:
                break
            if not s2 or len(s2) < 3:
                continue
            if _PDF_GONDERICI_DUR_RE.search(s2):
                break
            aday.append(s2)
        if aday:
            return " ".join(aday)
    return ""


def _pdf_tekil_fatura_ayikla(ham, yon="alis"):
    """Tek bir PDF sayfasını (=tek fatura, tedarikçiye özgü şablon) ayrıştırır.
    Döndürür: {tarih, gonderici, kalemler:[(oran,matrah,kdv)], ek_vergiler,
    fatura_no, iade, dosya} ya da None (ayrıştırılamadıysa)."""
    metin = ham.get("ham_metin", "")
    if not metin.strip() or len(metin) < 150:
        return None

    for tablo in ham.get("tablolar", []):
        for row in tablo[:3]:
            nrow = [norm(str(c)) for c in row]
            if any("GONDERICI ADI" in c for c in nrow) and any("ALICI ADI" in c for c in nrow):
                return None

    m = _PDF_FATURA_ID_RE.search(metin) or _PDF_FATURA_NO_RE.search(metin)
    fatura_no = m.group(1).strip() if m else ""

    m = _PDF_FATURA_TARIHI_RE.search(metin) or _PDF_TARIH_FALLBACK_RE.search(metin)
    d = mo = y = None
    if m:
        d, mo, y = m.groups()
    else:
        satirlar_ = metin.splitlines()
        for i, s in enumerate(satirlar_):
            if re.search(r"Fatura\s*Tarih", s, re.IGNORECASE):
                for s2 in satirlar_[i:i + 3]:
                    m2 = _PDF_TARIH_ANY_RE.search(s2)
                    if m2:
                        d, mo, y = m2.groups()
                        break
                break
    tarih = _tarih_iso(f"{d}.{mo}.{y}") if d else ""

    satirlar_all0 = [s.strip() for s in metin.splitlines()]
    satirlar = satirlar_all0[1:]
    gonderici = ""
    if yon == "satis":
        gonderici = _pdf_alici_adi_bul(metin)
    if not gonderici:
        for s in satirlar_all0[:20]:
            if not s or s.lower() in ("e-fatura", "e-fatura."):
                continue
            if _PDF_SAYIN_ETIKET_RE.match(s):
                break
            if _PDF_METADATA_ALAN_RE.search(s):
                continue
            if re.match(r"^\s*(BANKA|ŞUBE|SUBE|HESAP|IBAN)\s*:", s, re.IGNORECASE):
                continue
            eslesmeler = list(_PDF_UNVAN_RE.finditer(s))
            if not eslesmeler:
                continue
            bitis = eslesmeler[0].end()
            for sonraki in eslesmeler[1:]:
                if sonraki.start() - bitis <= 3:
                    bitis = sonraki.end()
                else:
                    break
            gonderici = s[:bitis].strip()
            break
    if not gonderici:
        aday = []
        for s in satirlar:
            if not s or s.lower() in ("e-fatura", "e-fatura."):
                continue
            if _PDF_GONDERICI_DUR_RE.search(s):
                break
            aday.append(s)
            if len(aday) >= 2:
                break
        gonderici = " ".join(aday)
        if _PDF_METADATA_ALAN_RE.search(gonderici):
            gonderici = ""

    iade = bool(_PDF_IADE_RE.search(metin))

    ozet = {}
    for tablo in ham.get("tablolar", []):
        for row in tablo:
            hucreler = [c for c in row if c not in (None, "")]
            # Özet satırı "etiket | tutar"dır; çok hücreli satırlar fatura KALEMİDİR
            # (kalemin "Diğer Vergiler" hücresindeki "KDVGUT..." KDV özeti sanılmasın).
            if len(hucreler) < 2 or len(hucreler) > 3:
                continue
            # Birleşik hücre: "Hesaplanan KDV(%10)\nHesaplanan KDV(%20)" | "2.124,80TL\n10.199,04TL"
            et_s = [x.strip() for x in str(hucreler[-2]).split("\n") if x.strip()]
            de_s = [x.strip() for x in str(hucreler[-1]).split("\n") if x.strip()]
            ciftler = list(zip(et_s, de_s)) if len(et_s) > 1 and len(et_s) == len(de_s) else \
                [(" ".join(str(hucreler[-2]).split()), " ".join(str(hucreler[-1]).split()))]
            for etiket, son in ciftler:
                son = son.lstrip(":").strip()
                if not _PDF_TUTAR_HUCRE_RE.match(son):
                    continue
                deger = _sayi(son)
                if deger is None:
                    continue
                ozet[etiket] = deger

    matrah = None
    kdv_kalemleri = []
    ek_vergi = 0.0
    toplam = None
    tevkifat = 0.0
    vergiler_dahil = None
    for etiket, deger in ozet.items():
        e = norm(etiket).replace(" ", "")
        if e.startswith("TEVKIFAT") or e.startswith("KDVTEVKIFAT") or e.startswith("HESAPLANANKDVTEVKIFAT"):
            tevkifat += deger          # KDV tevkifatı: KDV değil, ödenecekten düşülen kısım
        elif e.startswith("MALHIZMETTOPLAMTUTARI") or e.startswith("VERGIHARICTUTAR"):
            matrah = deger
        elif e == "TOPLAM" and matrah is None:
            matrah = deger
        elif e.startswith("HESAPLANANKDV") or (e.startswith("HESAPLANAN") and ("KDV" in e or "KATMADEGER" in e)):
            rm = _PDF_KDV_ORAN_RE.search(etiket)
            kdv_kalemleri.append((int(rm.group(1)) if rm else 20, deger))
        elif "KDV" in e:
            rm = _PDF_KDV_ORAN_RE.search(etiket)
            if rm:
                kdv_kalemleri.append((int(rm.group(1)), deger))
            else:
                ek_vergi += deger
        elif e.startswith("HESAPLANAN"):
            ek_vergi += deger
        elif (e.startswith("VERGILERDAHILTOPLAMTUTAR") or e.startswith("ODENECEKTUTAR")
              or e.startswith("GENELTOPLAM") or e.startswith("FATURATUTARI")):
            toplam = deger
            if e.startswith("VERGILERDAHIL"):
                vergiler_dahil = deger
    # Tevkifatlı faturada denge KDV dahil toplamla kurulur (ödenecek = KDV dahil − tevkifat)
    if tevkifat and vergiler_dahil is not None:
        toplam = vergiler_dahil

    if matrah is None:
        m = _PDF_MAL_HIZMET_RE.search(metin)
        if m:
            matrah = _sayi(m.group(1))
    if not kdv_kalemleri:
        m = _PDF_HESAPLANAN_KDV_METIN_RE.search(metin)
        if m:
            kdv_kalemleri.append((int(m.group(1)), _sayi(m.group(2))))
    if toplam is None:
        m = _PDF_ODENECEK_RE.search(metin)
        if m:
            toplam = _sayi(m.group(1))

    if matrah is None and not kdv_kalemleri:
        for lbl, oran_s, matrah_s, tutar_s in _PDF_VERGI_SATIRI_RE.findall(metin):
            oran_i = int(oran_s)
            m_i = _sayi(matrah_s)
            t_i = _sayi(tutar_s) or 0
            lbl_n = norm(lbl)
            if "KDV" in lbl_n or "KATMA DEGER" in lbl_n:
                kdv_kalemleri.append((oran_i, t_i))
                if m_i and (matrah is None or m_i > matrah):
                    matrah = m_i
            else:
                ek_vergi += t_i

    if matrah is None and toplam is not None:
        kdv_toplam = sum(t for _, t in kdv_kalemleri)
        matrah = round(toplam - kdv_toplam - ek_vergi, 2)

    if toplam is not None:
        kdv_toplam = sum(t for _, t in kdv_kalemleri)
        fark = round(toplam - ((matrah or 0) + kdv_toplam + ek_vergi), 2)
        if abs(fark) >= 0.01:
            if matrah is None:
                matrah = round(toplam - kdv_toplam - ek_vergi, 2)
            else:
                ek_vergi = round(ek_vergi + fark, 2)

    if not (fatura_no or gonderici) or (matrah is None and toplam is None):
        return None

    if kdv_kalemleri and len({o for o, _ in kdv_kalemleri}) > 1:
        # Çok oranlı: her oranın matrahı KDV'den türetilir (KDV / oran), yuvarlama farkı
        # en büyük orana yazılır — toplam matrah özetteki "Mal/Hizmet Toplam" ile tutar.
        tur_ = [(o, round(t * 100 / o, 2) if o else 0.0, t) for o, t in kdv_kalemleri]
        if matrah is not None and tur_:
            fark_ = round(matrah - sum(m for _, m, _ in tur_), 2)
            en = max(range(len(tur_)), key=lambda i: tur_[i][0])
            o_, m_, t_ = tur_[en]
            tur_[en] = (o_, round(m_ + fark_, 2), t_)
        kalemler = tur_
    elif kdv_kalemleri:
        kalemler = [(oran, matrah if i == 0 else 0, tutar) for i, (oran, tutar) in enumerate(kdv_kalemleri)]
    elif matrah:
        kalemler = [(0, matrah, 0)]
    else:
        kalemler = []
    return {
        "tarih": tarih, "gonderici": gonderici, "kalemler": kalemler,
        "ek_vergiler": ek_vergi, "fatura_no": fatura_no, "iade": iade,
        "dosya": ham.get("dosya", ""), "tevkifat": round(tevkifat, 2),
    }


def _pdf_gercek_fatura_ham(hamlar, yon="alis"):
    """Excel entegratör listesi değil, tek tek fatura PDF'i/görüntüsü olan
    dosyalar (her SAYFA = bir fatura, tedarikçi başına farklı şablon).
    belge_oku 'sayfalar' alanı varsa sayfa bazlı işlenir (çok faturalı tek
    PDF), yoksa dosyanın tamamı tek fatura sayılır (yaygın durum: avatar-tutor
    her faturayı ayrı dosya olarak yolluyor)."""
    kayitlar = []
    for h in hamlar:
        if h.get("tur") not in ("pdf", "pdf_ocr"):
            continue
        sayfalar = h.get("sayfalar") or [{"ham_metin": h.get("ham_metin", ""), "tablolar": h.get("tablolar", [])}]
        for sayfa_no, sayfa in enumerate(sayfalar, start=1):
            veri = {"ham_metin": sayfa.get("ham_metin", ""), "tablolar": sayfa.get("tablolar", []),
                    "dosya": h.get("dosya", "")}
            k = _pdf_tekil_fatura_ayikla(veri, yon=yon)
            if k:
                k["sayfa"] = sayfa_no
                kayitlar.append(k)
    return kayitlar


def _pdf_gercek_faturalar(hamlar, yon="alis"):
    """_pdf_gercek_fatura_ham'ın (gonderici/kalemler-tuple) ham çıktısını bu
    dosyadaki isle_fatura'nın beklediği sözlük şekline (cari_ad/kalemler-dict)
    çevirir — böylece hesap eşleştirmesi (mizan/geçmiş fiş) hiç değişmeden
    aynı şekilde çalışır."""
    ham_kayitlar = _pdf_gercek_fatura_ham(hamlar, yon=yon)
    out = []
    for k in ham_kayitlar:
        cari_ad = (k.get("gonderici") or "").strip()
        if not cari_ad:
            continue
        ham_kalemler = k.get("kalemler") or []
        ek_vergi = round(k.get("ek_vergiler") or 0, 2)
        iade = bool(k.get("iade"))
        if not ham_kalemler and ek_vergi > 0:
            # KDV kalemi yok, sadece Ek Vergiler (faktoring/BSMV tipik durumu)
            kalemler = []
            senaryo = "TEMELFATURA"
        else:
            kalemler = [{"oran": oran, "matrah": round(matrah, 2), "kdv": round(kdv, 2)}
                        for oran, matrah, kdv in ham_kalemler if matrah or kdv]
            senaryo = ""
        toplam = round(sum(x["matrah"] + x["kdv"] for x in kalemler) + ek_vergi, 2)
        out.append({
            "fatura_no": (k.get("fatura_no") or "").strip(),
            "tarih": k.get("tarih", ""),
            "cari_ad": cari_ad,
            "tur": "IADE" if iade else "SATIS",
            "senaryo": senaryo,
            "toplam": toplam,
            "kalemler": kalemler,
            "tevkifat": 0.0, "tevkifat_kod": "",
            "ek_vergi": ek_vergi,
            "yon": yon, "dosya": k.get("dosya", ""),
            # belge alanına (2. bölüm) yüklenmiş PDF — arayüzde açmak için yeri
            "pdf_yer": "belge" if str(k.get("dosya", "")).lower().endswith(".pdf") else "",
        })
    return out


def alt_hesap_kodlari(hesaplar) -> set:
    """Kayıt atılabilir (en alt kırılım) hesaplar: başka bir hesabın üst hesabı
    olmayanlar. "740" ve "740.01" üst hesaptır, "740.01.001" alt hesaptır."""
    ust = set()
    for k, _ in hesaplar:
        parca = k.split(".")
        for i in range(1, len(parca)):
            ust.add(".".join(parca[:i]))
    return {k for k, _ in hesaplar if k not in ust}


# ------------------------------------------------------- fatura listesi + fatura PDF'leri
_EKSIK_AD = {"cari": "cari", "tarih": "tarih", "kdv": "KDV dağılımı", "tutar": "tutar", "fatura_no": "fatura no"}


def _fno_anahtar(s) -> str:
    """Fatura no karşılaştırma anahtarı: büyük harf, yalnızca harf/rakam."""
    return re.sub(r"[^A-Z0-9]", "", norm(str(s or "")).replace(" ", ""))


def _kisa_liste(nolar, n=6):
    nolar = [x for x in nolar if x]
    return ", ".join(nolar[:n]) + (f" ve {len(nolar) - n} fatura daha" if len(nolar) > n else "")


def _liste_pdf_birlestir(liste, pdfler):
    """Fatura listesini (Excel) fatura PDF'leriyle fatura no üzerinden birleştirir.
    * Listede eksik olan cari / tarih / KDV dağılımı eşleşen PDF'ten tamamlanır.
      KDV dağılımı yalnızca PDF'in ödenecek toplamı listedekiyle aynıysa alınır.
    * Listede olmayan PDF faturaları eklenir.
    * Liste tam ama PDF'le tutar uyuşmuyorsa uyarı verilir (liste esas alınır).
    Döner: (faturalar, uyarilar)"""
    uyarilar = []
    pdf_idx = {}
    for p in pdfler:
        a = _fno_anahtar(p.get("fatura_no"))
        if a and a not in pdf_idx:
            pdf_idx[a] = p
    kullanilan, tamamlanan, tutar_farki, kdv_haric, tevkifat_pdf = set(), [], [], [], []
    for f in liste:
        a = _fno_anahtar(f["fatura_no"])
        p = pdf_idx.get(a)
        if not p:
            continue
        kullanilan.add(a)
        f["pdf"] = p.get("dosya", "")
        f["pdf_sayfa"] = p.get("sayfa", 0)
        f["pdf_yer"] = p.get("pdf_yer") or "alan"
        # fatura kalemlerinin açıklamaları (gider hesabı seçimi için) yalnız PDF'te var
        if p.get("kalem_aciklamalari") and not f.get("kalem_aciklamalari"):
            f["kalem_aciklamalari"] = p["kalem_aciklamalari"]
        if p.get("kalem_detay") and not f.get("kalem_detay"):
            f["kalem_detay"] = p["kalem_detay"]
        eksik = list(f.get("eksik", []))
        dolan = []
        if "cari" in eksik and p.get("cari_ad"):
            f["cari_ad"] = p["cari_ad"]; dolan.append("cari")
        if "tarih" in eksik and p.get("tarih"):
            f["tarih"] = p["tarih"]; dolan.append("tarih")
        pdf_tutar_tam = bool(p.get("kalemler")) and not ({"tutar", "kdv"} & set(p.get("eksik", [])))
        # Tevkifatlı fatura: PDF'in toplamı ÖDENECEK tutar; liste ya ödeneceği ya KDV dahil
        # toplamı verir. İkisi de aynı faturadır — tevkifat ve KDV dağılımı PDF'ten alınır.
        p_tev = round(float(p.get("tevkifat") or 0), 2)
        if p_tev > 0 and pdf_tutar_tam and not f.get("tevkifat"):
            p_dahil = round((p.get("toplam") or 0) + p_tev, 2)
            l_top = f.get("toplam") or 0
            if abs(l_top - p["toplam"]) <= 0.05 or abs(l_top - p_dahil) <= 0.05:
                f["tevkifat"] = p_tev
                f["tur"] = "TEVKIFAT"
                l_dahil = round(sum(k["matrah"] + k["kdv"] for k in f.get("kalemler") or []) + (f.get("ek_vergi") or 0), 2)
                if "kdv" in eksik or abs(l_dahil - p_dahil) > 0.05:
                    f["kalemler"] = p["kalemler"]
                    f["ek_vergi"] = p.get("ek_vergi", 0)
                    if "kdv" in eksik:
                        dolan.append("kdv")
                tevkifat_pdf.append(f["fatura_no"])
                if p.get("yz"):
                    f["yz"] = [x for x in dolan if x in p["yz"] or (x == "kdv" and "tutar" in p["yz"])]
                f["eksik"] = [e for e in eksik if e not in dolan]
                if dolan:
                    tamamlanan.append(f["fatura_no"])
                continue
        ayni_toplam = pdf_tutar_tam and abs((p.get("toplam") or 0) - (f.get("toplam") or 0)) <= 0.05
        # Listedeki tutar PDF'in MATRAHINA eşitse liste KDV hariç tutar vermiş demektir
        # (örn. "Mal Hizmet Toplam Tutarı" sütunu). Bu bir uyuşmazlık değil:
        # KDV dahil toplam ve KDV dağılımı PDF'ten alınır.
        pdf_matrah = round(sum(k["matrah"] for k in p.get("kalemler", [])), 2)
        liste_kdv_haric = (pdf_tutar_tam and not ayni_toplam and f.get("toplam")
                           and abs(pdf_matrah - f["toplam"]) <= 0.05 and p["toplam"] > f["toplam"]
                           and not any(k["kdv"] > 0 for k in f.get("kalemler", [])))
        if liste_kdv_haric:
            f["kalemler"] = p["kalemler"]
            f["ek_vergi"] = p.get("ek_vergi", 0)
            f["toplam"] = p["toplam"]
            if p.get("senaryo"):
                f["senaryo"] = p["senaryo"]
            if "kdv" in eksik:
                dolan.append("kdv")
            kdv_haric.append(f["fatura_no"])
        elif "kdv" not in eksik and ayni_toplam and not any(k["kdv"] > 0 for k in f.get("kalemler", [])) \
                and any(k["kdv"] > 0 for k in p.get("kalemler", [])):
            # Listede KDV sıfır/boş görünüyor ama aynı toplamlı faturanın PDF'inde KDV var:
            # liste KDV dağılımını vermemiş — dağılım PDF'ten alınır.
            f["kalemler"] = p["kalemler"]
            f["ek_vergi"] = p.get("ek_vergi", 0)
            if p.get("senaryo"):
                f["senaryo"] = p["senaryo"]
            dolan.append("kdv")
        elif "kdv" in eksik:
            if ayni_toplam:
                f["kalemler"] = p["kalemler"]
                f["ek_vergi"] = p.get("ek_vergi", 0)
                if p.get("senaryo"):
                    f["senaryo"] = p["senaryo"]
                dolan.append("kdv")
            elif pdf_tutar_tam:
                tutar_farki.append(f"{f['fatura_no']} (liste {f['toplam']:,.2f} / PDF {p['toplam']:,.2f})")
        elif pdf_tutar_tam and f.get("toplam") and not ayni_toplam and not liste_kdv_haric:
            tutar_farki.append(f"{f['fatura_no']} (liste {f['toplam']:,.2f} / PDF {p['toplam']:,.2f})")
        if p.get("yz"):
            # PDF'teki bu alanı yapay zekâ okuduysa listeye de öyle işaretle
            f["yz"] = [x for x in dolan if x in p["yz"] or (x == "kdv" and "tutar" in p["yz"])]
        f["eksik"] = [e for e in eksik if e not in dolan]
        if dolan:
            tamamlanan.append(f["fatura_no"])

    # listede olmayan PDF faturaları
    eklenen, eklenemeyen = [], []
    gorulen = set()
    for p in pdfler:
        a = _fno_anahtar(p.get("fatura_no"))
        if a and (a in kullanilan or a in gorulen):
            continue
        gorulen.add(a)
        if p.get("sunucudan"):
            continue        # sunucudan yalnız listedeki faturayı tamamlamak için alındı
        if "tutar" in p.get("eksik", []) or not p.get("toplam"):
            eklenemeyen.append(p.get("fatura_no") or p.get("dosya", ""))
            continue
        if not p.get("cari_ad"):
            p["cari_ad"] = p.get("fatura_no") or p.get("dosya", "")
        liste.append(p)
        eklenen.append(p.get("fatura_no") or p.get("dosya", ""))

    if tamamlanan:
        uyarilar.append(f"{len(tamamlanan)} faturanın listede eksik bilgisi PDF'ten tamamlandı")
    if kdv_haric:
        uyarilar.append(f"{len(kdv_haric)} faturada listedeki tutar KDV hariç (matrah); "
                        f"KDV dahil toplam ve KDV dağılımı PDF'ten alındı: {_kisa_liste(kdv_haric, 4)}")
    if tevkifat_pdf:
        uyarilar.append(f"{len(tevkifat_pdf)} faturada KDV tevkifatı PDF'ten alındı: {_kisa_liste(tevkifat_pdf, 4)}")
    if eklenen:
        uyarilar.append(f"Listede olmayan {len(eklenen)} fatura PDF'ten eklendi: {_kisa_liste(eklenen)}")
    if eklenemeyen:
        uyarilar.append(f"Tutarı okunamadığı için eklenemeyen PDF: {_kisa_liste(eklenemeyen)}")
    if tutar_farki:
        uyarilar.append(f"Liste ile PDF toplamı farklı (liste esas alındı): {_kisa_liste(tutar_farki, 4)}")
    # hâlâ eksik kalanlar — tür bazında grupla
    for tur, ad in _EKSIK_AD.items():
        nolar = [f["fatura_no"] for f in liste if tur in f.get("eksik", [])]
        if nolar:
            ek = " (KDV satırı yazılmadı, tutarın tamamı gidere gitti)" if tur == "kdv" else ""
            uyarilar.append(f"{ad} eksik, PDF'i yok ya da okunamadı{ek}: {_kisa_liste(nolar)}")
    return liste, uyarilar


def _kalem_ogrenilmis(aciklama: str, ogrenilen: dict):
    """Tek kalem açıklaması için öğrenilmiş hesap: (kod, benzer_mi) ya da None."""
    from app.kurallar import kelimeler
    n = norm(aciklama)
    if not n:
        return None
    if n in ogrenilen:
        return ogrenilen[n], False
    kw = kelimeler(n)
    if len(kw) < 2:
        return None
    en, en_skor = None, 0.0
    for ad, kod in ogrenilen.items():
        ow = kelimeler(ad)
        if ow:
            skor = len(kw & ow) / len(kw | ow)
            if skor > en_skor:
                en, en_skor = kod, skor
    return (en, True) if en and en_skor >= 0.6 else None


def _kalem_ogrenmesi(f: dict, gider_ogrenme: dict | None, yon: str, alt_kodlar: set) -> dict | None:
    """Faturanın kalem açıklamaları için kullanıcının daha önce seçtiği gider hesabı.
    Öğrenme anahtarı: "kalem|<yön>|<NORM(açıklama)>". Birebir eşleşme yoksa kelimelerinin
    çoğu tutan öğrenilmiş açıklama (benzer=True — KONTROL ister). Kalemler farklı hesaplara
    gidiyorsa tutarı en büyük olan kalemin hesabı seçilir."""
    og = gider_ogrenme or {}
    onek = f"kalem|{yon}|"
    ogrenilen = {k[len(onek):]: v for k, v in og.items() if k.startswith(onek) and v in alt_kodlar}
    if not ogrenilen:
        return None
    detay = f.get("kalem_detay") or [{"a": a, "t": None} for a in (f.get("kalem_aciklamalari") or [])]
    if not detay:
        return None
    from app.kurallar import kelimeler
    bulunan = []          # (tutar, kod, açıklama, benzer)
    for k in detay:
        n = norm(k.get("a", ""))
        if not n:
            continue
        if n in ogrenilen:
            bulunan.append((k.get("t") or 0, ogrenilen[n], k["a"], False))
            continue
        kw = kelimeler(n)
        if len(kw) < 2:
            continue
        en, en_skor = None, 0.0
        for ad, kod in ogrenilen.items():
            ow = kelimeler(ad)
            if not ow:
                continue
            skor = len(kw & ow) / len(kw | ow)
            if skor > en_skor:
                en, en_skor = (ad, kod), skor
        if en and en_skor >= 0.6:
            bulunan.append((k.get("t") or 0, en[1], k["a"], True))
    if not bulunan:
        return None
    tam = [b for b in bulunan if not b[3]]
    secim = max(tam or bulunan, key=lambda b: b[0])
    return {"kod": secim[1], "benzer": not tam,
            "not": f"Kalem öğrenmesi: '{secim[2][:50]}' daha önce bu hesaba yazıldı"
                   + (" (benzer açıklama)" if not tam else "")}


KARSILASTIRMA_ORANLARI = (1, 10, 20)


def _oran_dagilim(f: dict) -> dict:
    """Fatura kalemleri -> {'1': [matrah, kdv], '10': [...], '20': [...], 'diger': [...]}"""
    out = {str(o): [0.0, 0.0] for o in KARSILASTIRMA_ORANLARI}
    out["diger"] = [0.0, 0.0]
    for k in f.get("kalemler") or []:
        try:
            oran = int(round(float(k.get("oran") or 0)))
        except Exception:
            oran = -1
        anahtar = str(oran) if oran in KARSILASTIRMA_ORANLARI else "diger"
        out[anahtar][0] += float(k.get("matrah") or 0)
        out[anahtar][1] += float(k.get("kdv") or 0)
    return {a: [round(m, 2), round(v, 2)] for a, (m, v) in out.items()}


def liste_pdf_karsilastir(liste: list, pdfler: list, tolerans: float = 0.05) -> dict:
    """Fatura listesindeki (Excel) matrah/KDV tutarlarını fatura PDF'lerindekiyle
    %1 / %10 / %20 (ve diğer oranlar) bazında, fatura no üzerinden karşılaştırır.
    durum: uyumlu | farkli | pdf_yok | pdf_eksik | liste_kdv_yok | listede_yok"""
    pdf_idx = {}
    for p in pdfler:
        a = _fno_anahtar(p.get("fatura_no"))
        if a and a not in pdf_idx:
            pdf_idx[a] = p

    def tam(f):
        return bool(f.get("kalemler")) and not ({"tutar", "kdv"} & set(f.get("eksik") or []))

    def pdf_bilgi(p):
        return {"dosya": p.get("dosya", ""), "sayfa": p.get("sayfa", 0), "yer": p.get("pdf_yer") or "alan"}

    satirlar, gorulen = [], set()
    for f in liste:
        a = _fno_anahtar(f.get("fatura_no"))
        p = pdf_idx.get(a)
        if a:
            gorulen.add(a)
        ld = _oran_dagilim(f) if tam(f) else None
        pd = _oran_dagilim(p) if p and tam(p) else None
        farklar = []
        if not p:
            durum = "pdf_yok"
        elif pd is None:
            durum = "pdf_eksik"
        elif ld is None:
            durum = "liste_kdv_yok"
        else:
            for o in ld:
                if abs(ld[o][0] - pd[o][0]) > tolerans:
                    farklar.append(f"%{o} matrah" if o != "diger" else "%0/diğer matrah")
                if abs(ld[o][1] - pd[o][1]) > tolerans:
                    farklar.append(f"%{o} KDV" if o != "diger" else "%0/diğer KDV")
            # tevkifatlı faturada PDF toplamı ödenecek tutardır; liste KDV dahil toplamı da verebilir
            p_tev = float(p.get("tevkifat") or 0)
            l_top, p_top = f.get("toplam") or 0, p.get("toplam") or 0
            if abs(l_top - p_top) > tolerans and not (p_tev and abs(l_top - (p_top + p_tev)) <= tolerans):
                farklar.append("toplam")
            l_tev = float(f.get("tevkifat") or 0)
            if l_tev and p_tev and abs(l_tev - p_tev) > tolerans:
                farklar.append("tevkifat")
            durum = "farkli" if farklar else "uyumlu"
        satirlar.append({
            "fatura_no": f.get("fatura_no", ""), "tarih": f.get("tarih", ""), "cari": f.get("cari_ad", ""),
            "liste": ld, "pdf": pd, "liste_toplam": round(f.get("toplam") or 0, 2),
            "pdf_toplam": round(p.get("toplam") or 0, 2) if p else None,
            "pdf_dosya": pdf_bilgi(p) if p else None, "durum": durum, "farklar": farklar,
            "liste_tevkifat": round(float(f.get("tevkifat") or 0), 2),
            "pdf_tevkifat": round(float(p.get("tevkifat") or 0), 2) if p else None,
        })
    for a, p in pdf_idx.items():
        if a in gorulen or p.get("sunucudan"):
            continue
        satirlar.append({
            "fatura_no": p.get("fatura_no", ""), "tarih": p.get("tarih", ""), "cari": p.get("cari_ad", ""),
            "liste": None, "pdf": _oran_dagilim(p) if tam(p) else None, "liste_toplam": None,
            "pdf_toplam": round(p.get("toplam") or 0, 2), "pdf_dosya": pdf_bilgi(p),
            "durum": "listede_yok", "farklar": [],
            "liste_tevkifat": None, "pdf_tevkifat": round(float(p.get("tevkifat") or 0), 2),
        })

    # oran bazında toplamlar: "ortak" = iki tarafı da okunabilen faturalar (karşılaştırılabilir küme);
    # "tum" = her taraf kendi okunabilen faturaları (biri eksikse fark anlamlı değildir)
    anahtarlar = [str(x) for x in KARSILASTIRMA_ORANLARI] + ["diger"]
    def _bos():
        return {o: {"liste": [0.0, 0.0], "pdf": [0.0, 0.0]} for o in anahtarlar}
    toplam, tum = _bos(), _bos()
    ortak_sayi = liste_sayi = pdf_sayi = 0
    for r in satirlar:
        liste_sayi += bool(r["liste"]); pdf_sayi += bool(r["pdf"])
        for taraf in ("liste", "pdf"):
            if r[taraf]:
                for o in anahtarlar:
                    for i in (0, 1):
                        tum[o][taraf][i] += r[taraf][o][i]
        if r["liste"] and r["pdf"]:
            ortak_sayi += 1
            for o in anahtarlar:
                for i in (0, 1):
                    toplam[o]["liste"][i] += r["liste"][o][i]
                    toplam[o]["pdf"][i] += r["pdf"][o][i]
    for t_ in (toplam, tum):
        for o in t_:
            for taraf in ("liste", "pdf"):
                t_[o][taraf] = [round(x, 2) for x in t_[o][taraf]]
    tevkifat = {"liste": 0.0, "pdf": 0.0, "liste_tum": 0.0, "pdf_tum": 0.0, "sayi": 0}
    for r in satirlar:
        lt, pt = r.get("liste_tevkifat") or 0, r.get("pdf_tevkifat") or 0
        tevkifat["liste_tum"] += lt; tevkifat["pdf_tum"] += pt
        tevkifat["sayi"] += bool(lt or pt)
        if r["liste"] and r["pdf"]:
            tevkifat["liste"] += lt; tevkifat["pdf"] += pt
    tevkifat = {k: (round(v, 2) if isinstance(v, float) else v) for k, v in tevkifat.items()}
    sayilar = {}
    for r in satirlar:
        sayilar[r["durum"]] = sayilar.get(r["durum"], 0) + 1
    sira = {"farkli": 0, "liste_kdv_yok": 1, "pdf_eksik": 2, "pdf_yok": 3, "listede_yok": 4, "uyumlu": 5}
    satirlar.sort(key=lambda r: (sira.get(r["durum"], 9), r["tarih"] or "", r["fatura_no"]))
    return {"satirlar": satirlar, "toplam": toplam, "toplam_tum": tum, "sayilar": sayilar, "tevkifat": tevkifat,
            "ortak_sayi": ortak_sayi, "liste_kdv_sayi": liste_sayi, "pdf_kdv_sayi": pdf_sayi}


def isle_fatura(hamlar, km, fis0, yon="alis", pdf_faturalar=None,
                gider_ogrenme=None, gider_onerici=None, kalem_onerici=None):
    """
    Fatura listesi -> muhasebe fişi.
    ALIŞ: gider(7xx) + 191(indirilecek KDV) borç / 320(satıcı) alacak
    SATIŞ: 120(alıcı) borç / 600(satış) + 391(hesaplanan KDV) alacak
    TEVKIFAT: KDV'nin sorumlu kısmı ayrı hesaba (191.03/391.03)
    FAKTORING (TEMELFATURA + Ek Vergi): 780 BSMV borç, KDV yok
    """
    uyarilar = []
    hes_list = [{"kod": k, "ad": a} for k, a in km.hesaplar]
    # tekrar edenleri kaldır
    seen = set(); hes_uniq = []
    for h in hes_list:
        if h["kod"] not in seen:
            seen.add(h["kod"]); hes_uniq.append(h)
    hes_list = hes_uniq

    pdf_faturalar = pdf_faturalar or []
    faturalar = _elogo_fatura_satirlari(hamlar, yon)
    # Aynı fatura birden çok listede olabilir (ör. ay klasöründe hem İVD hem entegratör listesi):
    # bir kez işlenir; KDV dağılımı olan kayıt tercih edilir.
    if faturalar:
        tekil, cift_liste = {}, []
        for f in faturalar:
            a = _fno_anahtar(f["fatura_no"]) or id(f)
            if a in tekil:
                cift_liste.append(f["fatura_no"])
                if "kdv" in (tekil[a].get("eksik") or []) and "kdv" not in (f.get("eksik") or []):
                    tekil[a] = f
            else:
                tekil[a] = f
        if cift_liste:
            faturalar = list(tekil.values())
            uyarilar.append(f"{len(cift_liste)} fatura birden çok listede var, bir kez işlendi: {_kisa_liste(cift_liste)}")
    if faturalar:
        # Liste esas; "Fatura PDF'leri" alanındaki PDF'ler eksikleri tamamlar.
        # Belge alanına listeyle birlikte atılmış PDF'ler de aynı işe yarasın.
        ek_pdf = [f for f in _pdf_gercek_faturalar(hamlar, yon)
                  if _fno_anahtar(f["fatura_no"]) not in {_fno_anahtar(p.get("fatura_no")) for p in pdf_faturalar}]
        faturalar, birlesme_uyari = _liste_pdf_birlestir(faturalar, pdf_faturalar + ek_pdf)
        uyarilar.extend(birlesme_uyari)
    if not faturalar:
        # eLogo listesi yok — gerçek e-Fatura/e-Arşiv PDF'i olabilir (tek tek fatura)
        faturalar = _pdf_gercek_faturalar(hamlar, yon)
        anahtarlar = {_fno_anahtar(f["fatura_no"]) for f in faturalar}
        for p in pdf_faturalar:
            if _fno_anahtar(p.get("fatura_no")) in anahtarlar or not p.get("toplam"):
                continue
            if not p.get("cari_ad"):
                p["cari_ad"] = p.get("fatura_no") or p.get("dosya", "")
            faturalar.append(p)
        eksikli = [f.get("fatura_no") or f.get("dosya", "") for f in faturalar if f.get("eksik")]
        if eksikli:
            uyarilar.append(f"PDF'te eksik alan kalan fatura: {_kisa_liste(eksikli)}")
    if not faturalar:
        # PDF de değilse — OCR ham metninden perakende fiş fotoğrafı olabilir
        faturalar = _ham_metin_faturalar(hamlar, yon)
    if not faturalar:
        for h in hamlar:
            if h.get("ham_metin"):
                uyarilar.append(f"{h.get('dosya')}: fatura verisi çıkarılamadı (OCR gürültülü olabilir) — elle düzeltme gerekebilir")
        return [], uyarilar

    # Parser'ın tek tek belge için ürettiği uyarılar (ör. OCR'da bozuk tarih)
    for f in faturalar:
        if f.get("uyari"):
            uyarilar.append(f["uyari"])

    # Çift kayıt: geçmiş fişlerde (fiş listesi / muavin) açıklamasında bu fatura
    # numarası geçiyorsa fatura daha önce muhasebeleşmiştir.
    try:
        kayitli = km.kayitli_faturalar()
    except Exception:
        kayitli = {}
    if kayitli:
        cift = []
        for f in faturalar:
            no = _fno_anahtar(f.get("fatura_no"))
            if no not in kayitli:
                continue
            fisno = km.fatura_kayitli_mi(no, f.get("cari_ad", "")) if hasattr(km, "fatura_kayitli_mi") \
                else kayitli[no]
            if fisno is not None:
                f["kayitli_fis"] = fisno
                cift.append(f"{f['fatura_no']} (fiş {fisno})" if fisno else f["fatura_no"])
        if cift:
            uyarilar.insert(0, f"{len(cift)} fatura geçmiş kayıtlarda zaten var — çift kayıt olmasın, "
                               f"aktarmadan önce çıkar: {_kisa_liste(cift)}")

    alt_kodlar = alt_hesap_kodlari(km.hesaplar)
    vars_gider = next((k for onek in (("740", "770", "760", "730", "150", "153") if yon == "alis"
                                      else ("600", "601", "602"))
                       for k, _ in km.hesaplar if k in alt_kodlar and (k == onek or k.startswith(onek + "."))),
                      "")
    if not vars_gider and yon == "alis":   # tercihli gruplarda yoksa herhangi bir 7xx gider alt hesabı
        vars_gider = next((k for k, _ in km.hesaplar if k in alt_kodlar and k[:1] == "7" and not k.startswith("79")), "")
    gider_bekleyen = set()
    kalem_bekleyen = set()
    kdv_eksik = {}      # "%20 için 191" -> {fatura no}
    fark_yazilan, fark_dengesiz, yeni_hesap = [], [], {}
    tevkifatli, tevkifat_tahminli, tevkifat_tutarsiz = [], [], []
    tevkifat_hesapsiz, tevkifat_kdvden, tevk_hesap_belirsiz = set(), set(), set()
    # bu firmanın geçmiş kayıtlarında (fiş listesi / muavin) gider-gelir tarafında kullanılmış hesaplar
    kullanilan_gider = {str(r.get("hesap", "")).strip() for r in km.gecmis.get("satirlar", [])
                        if float((r.get("borc") if yon == "alis" else r.get("alacak")) or 0) > 0}

    fisler = []
    fis = fis0
    for f in sorted(faturalar, key=lambda x: x["tarih"] or ""):
        fisno = f"{fis:05d}"
        cari_ad = f["cari_ad"]
        tarih = f["tarih"]
        fatura_no = f["fatura_no"]
        toplam = f["toplam"]

        # KDV tevkifatı: listede/PDF'te yoksa liste toplamı KDV dahil tutardan standart bir
        # tevkifat payı kadar azsa tahmin edilir (KONTROL işaretli)
        tevk = round(float(f.get("tevkifat") or 0), 2)
        tevk_tahmin = False
        if not tevk and f.get("kalemler") and f.get("senaryo") != "TEMELFATURA" \
                and not ({"kdv", "tutar"} & set(f.get("eksik") or [])):
            tevk = _tevkifat_tahmin(f["kalemler"], toplam, f.get("ek_vergi") or 0)
            tevk_tahmin = tevk > 0

        if f.get("tur") == "IADE":
            uyarilar.append(f"{cari_ad[:40]} ({fatura_no}): IADE faturasi - borc/alacak yonu kontrol edilmeli")

        # FATURA HESAPLARI: firma kuralı yerine geçmiş fişler referans alınır.
        # Aynı cari daha önce hangi 320/120 ve hangi gider/gelir hesabıyla
        # kullanılmışsa onu tercih ederiz. Geçmişte kayıt yoksa varsayılana düşeriz.
        gecmis_es = km.fatura_gecmis_eslestir(cari_ad, yon)
        cari_kod = gecmis_es.get("cari", "")
        gider_kod = gecmis_es.get("ana", "")
        cari_kaynak = gecmis_es.get("kaynak", "")
        gecmis_kdv = gecmis_es.get("kdv", [])

        # Geçmiş fişlerde bu cari hiç geçmemişse (yeni tedarikçi/müşteri veya
        # cari adı ilk kez okunuyor) mizan hesap adlarıyla da dene — bugünkü
        # eslestir() düzeltmesiyle aynı mekanizma (kurallar.py).
        cari_onek = ("320", "329", "331", "335", "336") if yon == "alis" else ("120", "121")
        # Kullanıcının bu cari için düzelttiği/onayladığı hesap her şeyden önce gelir
        cari_ogr = km.ogrenme.get(norm(cari_ad)) if getattr(km, "ogrenme", None) else None
        if cari_ogr and cari_ogr in alt_kodlar and cari_ogr.startswith(cari_onek):
            cari_kod, cari_kaynak = cari_ogr, "ogrenme"
        if not cari_kod:
            ek_kod, ek_kaynak = km.eslestir(cari_ad, onekler=cari_onek)
            if ek_kod:
                cari_kod = ek_kod
                cari_kaynak = cari_kaynak or ek_kaynak

        # cari doğru öneke uymuyorsa varsayılan
        vars_cari = "198.01.001"
        if not cari_kod or not cari_kod.startswith(cari_onek):
            cari_kod = vars_cari

        # GİDER / GELİR HESABI — öncelik sırası:
        #   1) kullanıcının bu cari için yaptığı düzeltme (gider öğrenmesi)
        #   2) geçmiş fişlerde bu carinin karşısındaki hesap
        #   3) yapay zekâ önerisi (fatura kalemleri + satıcı adı, yalnız mizandaki alt hesaplar)
        #   4) varsayılan ALT hesap (ana hesap "740" değil — ona kayıt atılamaz)
        gider_kaynak, gider_not = "", ""
        yasak_onek = ("320", "329", "331", "335", "336", "120", "121", "100", "102", "191", "391")
        ogr = (gider_ogrenme or {}).get(f"{yon}|{norm(cari_ad)}")
        kalem_es = _kalem_ogrenmesi(f, gider_ogrenme, yon, alt_kodlar)
        if kalem_es and not kalem_es["benzer"]:
            # kullanıcının bu kalem açıklaması için seçtiği hesap (cariden bağımsız, en özel bilgi)
            gider_kod, gider_kaynak, gider_not = kalem_es["kod"], "ogrenme_kalem", kalem_es["not"]
        elif ogr and ogr in alt_kodlar:
            gider_kod, gider_kaynak = ogr, "ogrenme"
        elif kalem_es:
            gider_kod, gider_kaynak, gider_not = kalem_es["kod"], "ogrenme_kalem_benzer", kalem_es["not"]
        elif gider_kod and gider_kod in alt_kodlar and gider_kod != cari_kod \
                and not gider_kod.startswith(yasak_onek):
            gider_kaynak = "gecmis"
        else:
            gider_kod = ""
            # Cari adı okunamamış ve kalem açıklaması da yoksa modele verilecek bilgi yok
            # (yalnız fatura no) — sormak boşuna Ollama'yı meşgul eder.
            bilgi_var = bool(f.get("kalem_aciklamalari")) or "cari" not in f.get("eksik", [])
            if gider_onerici and bilgi_var:
                oneri = gider_onerici(cari_ad, f.get("kalem_aciklamalari") or [], yon)
                if oneri and oneri.get("kod") in alt_kodlar:
                    gider_kod, gider_kaynak, gider_not = oneri["kod"], "yz", oneri.get("gerekce", "")
                elif oneri and oneri.get("bekliyor"):
                    gider_bekleyen.add(cari_ad)
            if not gider_kod:
                gider_kod, gider_kaynak = vars_gider, "tahmin"

        # Belge kaynağı rozeti: yapay zekâ ile tamamlanan > PDF'ten gelen/tamamlanan
        belge_rozet = ("yz" if f.get("yz") else
                       "pdf" if (f.get("pdf") or f.get("kaynak") == "pdf") else "")

        # Kalem başına hesap: öğrenilmiş (birebir / benzer) ya da yapay zekâ önerisi
        kalem_hesap_detay = [dict(k) for k in (f.get("kalem_detay") or [])[:12]] or \
            [{"a": a, "t": None} for a in (f.get("kalem_aciklamalari") or [])[:12]]
        if kalem_hesap_detay:
            onek_ = f"kalem|{yon}|"
            ogrenilen_ = {k[len(onek_):]: v for k, v in (gider_ogrenme or {}).items()
                          if k.startswith(onek_) and v in alt_kodlar}
            eksik_ = False
            for k in kalem_hesap_detay:
                o_ = _kalem_ogrenilmis(k.get("a", ""), ogrenilen_) if ogrenilen_ else None
                if o_:
                    k["k"], k["kk"] = o_[0], ("benzer" if o_[1] else "ogrenme")
                else:
                    eksik_ = True
            if eksik_ and kalem_onerici:
                yz_ = kalem_onerici(cari_ad, [k.get("a", "") for k in kalem_hesap_detay], yon)
                if yz_ and yz_.get("bekliyor"):
                    kalem_bekleyen.add(fatura_no)
                    for k in kalem_hesap_detay:
                        k.setdefault("kk", "bekliyor")
                elif yz_ and yz_.get("kalemler"):
                    for k, o_ in zip(kalem_hesap_detay, yz_["kalemler"]):
                        if "k" not in k and o_:
                            k["k"], k["kk"], k["g"] = o_["kod"], "yz", o_.get("gerekce", "")
            # Etkin kalem hesabı: öğrenilmiş kalem > (fatura hesabı kullanıcıdan/geçmişten geliyorsa) o
            # > yapay zekâ kalem önerisi > fatura hesabı. Yapay zekânın farklı önerisi "yk"de kalır.
            fatura_guvenilir = gider_kaynak in ("ogrenme", "gecmis")
            # Fatura hesabı güvenilir olsa da yapay zekâ kalemleri FARKLI hesaplara ayırıyorsa
            # (karışık fatura: kiralama + malzeme) bölünür; fatura hesabıyla aynı olanlar onaylı sayılır.
            karisik = len({k["k"] for k in kalem_hesap_detay if k.get("kk") == "yz"}) > 1
            for k in kalem_hesap_detay:
                if k.get("kk") in ("ogrenme", "benzer"):
                    continue
                if k.get("kk") == "yz" and (not fatura_guvenilir or (karisik and k["k"] != gider_kod)):
                    continue
                if k.get("kk") == "yz" and k.get("k") != gider_kod:
                    k["yk"] = k["k"]
                if gider_kod:
                    k["k"], k["kk"] = gider_kod, "fatura"

        def gider_bolumleri(oran, matrah):
            """Bir KDV oranının matrahını kalem hesaplarına böler: [(kod, tutar, açıklamalar, kaynak)].
            Kalem tutarları o oranın matrahını tutmuyorsa (fatura altı iskonto vb.) bölünmez."""
            tek = [(gider_kod or vars_gider, matrah, None, None)]
            det = [k for k in kalem_hesap_detay if k.get("t") is not None and k.get("k")]
            if not det or len(det) != len(kalem_hesap_detay):
                return tek
            if any("o" in k for k in det):
                det = [k for k in det if k.get("o") == oran]
            elif len({x["oran"] for x in f.get("kalemler") or []}) > 1:
                return tek
            if not det or abs(sum(k["t"] for k in det) - matrah) > 0.05 + 0.01 * len(det):
                return tek
            gruplar = {}
            for k in det:
                g = gruplar.setdefault(k["k"], {"t": 0.0, "a": [], "kk": set()})
                g["t"] += k["t"]; g["a"].append(k["a"]); g["kk"].add(k.get("kk", ""))
            if len(gruplar) == 1 and next(iter(gruplar)) == (gider_kod or vars_gider):
                return [(gider_kod or vars_gider, matrah, [k["a"] for k in det], None)]
            out, kalan = [], round(matrah, 2)
            sirali = sorted(gruplar.items(), key=lambda x: -x[1]["t"])
            for i, (kod, g) in enumerate(sirali):
                tutar = kalan if i == len(sirali) - 1 else round(g["t"], 2)
                kalan = round(kalan - tutar, 2)
                kaynak_ = ("yz" if "yz" in g["kk"] else "ogrenme_kalem_benzer" if "benzer" in g["kk"]
                           else "ogrenme_kalem" if g["kk"] <= {"ogrenme"} else None)
                out.append((kod, tutar, g["a"], kaynak_))
            return out

        def gider_yaz(oran, matrah, detay_ek=""):
            """Gider (alış) / gelir (satış) satır(lar)ı — kalem hesaplarına bölünmüş olabilir."""
            for kod, tutar, aciklamalar, kaynak_ in gider_bolumleri(oran, matrah):
                if yon == "alis":
                    fisler.append(sat(kod, tutar, 0, detay_ek, rol="gider", kaynak_=kaynak_, kalemler=aciklamalar))
                else:
                    fisler.append(sat(kod, 0, tutar, detay_ek, rol="gider", kaynak_=kaynak_, kalemler=aciklamalar))

        def sat(hesap, borc, alacak, detay_ek="", rol=None, kaynak_=None, kalemler=None):
            # Açıklama alanları SADECE firma adı (KDV oranı, "(faktoring)" gibi ekler yok).
            s = _sat(fisno, tarih, cari_ad, hesap, borc, alacak,
                     evrak_no=fatura_no, detay=cari_ad, kaynak=cari_kaynak or "fatura")
            if belge_rozet:
                s["belge"] = belge_rozet
            # önizlemede gösterilecek fatura içeriği: kalem açıklamaları ve KDV hariç tutar
            s["fatura_kalemleri"] = [str(a)[:120] for a in (kalemler if kalemler is not None
                                                            else (f.get("kalem_aciklamalari") or []))[:8]]
            s["fatura_kalem_detay"] = kalem_hesap_detay
            s["kdv_haric"] = round(sum(float(k.get("matrah") or 0) for k in f.get("kalemler") or []), 2)
            s["kdv_dagilim"] = [{"oran": k.get("oran"), "matrah": round(float(k.get("matrah") or 0), 2),
                                 "kdv": round(float(k.get("kdv") or 0), 2)} for k in f.get("kalemler") or []]
            if tevk > 0:
                kdv_dahil_ = round(sum(float(k.get("matrah") or 0) + float(k.get("kdv") or 0)
                                       for k in f.get("kalemler") or []) + float(f.get("ek_vergi") or 0), 2)
                s["tevkifat"] = tevk
                s["kdv_dahil"] = kdv_dahil_
                s["odenecek"] = round(kdv_dahil_ - tevk, 2)
                if tevk_tahmin:
                    s["tevkifat_tahmin"] = True
            # faturanın PDF'i (arayüzde açıp bakmak için)
            pdf_ad = f.get("pdf") or (f.get("dosya", "") if (f.get("kaynak") == "pdf" or f.get("pdf_yer")) else "")
            if pdf_ad:
                s["pdf"] = pdf_ad
                s["pdf_yer"] = f.get("pdf_yer") or "alan"     # alan: Fatura PDF'leri · belge: 2. bölüm
                if f.get("pdf_sayfa") or f.get("sayfa"):
                    s["pdf_sayfa"] = f.get("pdf_sayfa") or f.get("sayfa")
            # rol: arayüzde düzeltme hangi öğrenmeye gidecek (gider düzeltmesi cariye öğrenilmesin)
            if rol == "tevkifat":
                s["rol"] = "tevkifat"
            elif "KDV)" in detay_ek or hesap.startswith(("191", "391")):
                s["rol"] = "kdv"
                m_ = re.search(r"%(\d+)", detay_ek)
                if m_:
                    s["oran"] = int(m_.group(1))
            elif rol == "gider" or (hesap == gider_kod and hesap != cari_kod):
                s["rol"] = "gider"
                s["kaynak"] = kaynak_ or gider_kaynak
                if kaynak_ and kaynak_ != gider_kaynak:
                    s["not"] = {"yz": "Kalem bazında yapay zekâ önerisi",
                                "ogrenme_kalem": "Kalem öğrenmesi",
                                "ogrenme_kalem_benzer": "Benzer kalem öğrenmesi"}.get(kaynak_, "")
                    if not s["not"]:
                        del s["not"]
                elif gider_not:
                    s["not"] = gider_not
            elif hesap == cari_kod:
                s["rol"] = "cari"
            else:
                s["rol"] = "diger"
            # Doğrulanmamış hesaplar: Excel'e aktarılır ama fiş açıklamasına KONTROL yazılır
            # (kullanıcı önizlemede düzeltir ya da onaylarsa işaret kalkar).
            kontrol = []
            if s["rol"] == "gider" and s.get("kaynak") in ("yz", "tahmin", "ogrenme_kalem_benzer"):
                kontrol.append("gider")
            if s["rol"] == "kdv" and not hesap:
                kontrol.append("kdv")
            if s["rol"] == "tevkifat" and (not hesap or fatura_no in tevk_hesap_belirsiz):
                kontrol.append("tevkifat_hesap")
            if tevk_tahmin and s["rol"] in ("tevkifat", "cari"):
                kontrol.append("tevkifat_tahmin")
            # Listede/PDF'te KDV dağılımı yoksa tutarın tamamı (KDV dahil) gidere yazılır — kontrol şart
            if s["rol"] == "gider" and (not f.get("kalemler") or {"kdv", "tutar"} & set(f.get("eksik") or [])):
                kontrol.append("kdv_yok")
            if s["rol"] == "cari":
                if hesap == vars_cari:
                    kontrol.append("cari")
                if belge_rozet == "yz":
                    kontrol.append("belge")
            # (kullanıcı cariyi onayladıysa aynı geçmiş kaydından gelen gider de onaylı sayılır)
            if gecmis_es.get("zayif") and cari_kaynak != "ogrenme" and \
                    ((s["rol"] == "gider" and gider_kaynak == "gecmis") or s["rol"] == "cari"):
                kontrol.append("benzer_ad")
            if kontrol:
                s["kontrol"] = kontrol
            return s

        def gecmis_kdv_sec(oran, yon):
            # Kullanıcı bu oranın KDV satırını bir kez düzelttiyse o hesap (firma geneli)
            ogr = (gider_ogrenme or {}).get(f"kdv|{yon}|{oran}")
            if ogr and ogr in alt_kodlar:
                return ogr
            # Geçmişte aynı cari için kullanılan KDV hesabını öncele.
            if gecmis_kdv:
                # Birden fazla KDV hesabı varsa mizan adında oranı tutan hesabı ara.
                for kod in gecmis_kdv:
                    ad = km.hesap_adi(kod)
                    nad = norm(ad).replace(" ", "")
                    if re.search(rf"%{str(oran)}(?!\d)", nad) or re.search(rf"(?<!\d){str(oran)}(?!\d)", nad):
                        return kod
                return gecmis_kdv[0]
            return _kdv_hesabi_ad([{"kod": k, "ad": a} for k, a in km.hesaplar], oran, yon)

        def tevkifat_hesap_sec(yon):
            # kullanıcının düzelttiği (firma geneli) > bu carinin geçmişi > mizan adı
            ogr = (gider_ogrenme or {}).get(f"tevkifat|{yon}")
            if ogr and ogr in alt_kodlar:
                return ogr
            for kod in gecmis_es.get("tevkifat") or []:
                ad_ = norm(km.hesap_adi(kod) or "")
                # 360'ta gelir vergisi stopajı (SMM, kira) da olur — KDV tevkifatı sayılmaz
                if kod in alt_kodlar and not any(x in ad_ for x in ("GELIR", "STOPAJ", "DAMGA", "SGK", "SIGORTA", "KURUMLAR")):
                    return kod
            kod = _tevkifat_hesabi(hes_list, yon, alt_kodlar)
            ad_ = norm(km.hesap_adi(kod) or "") if kod else ""
            if kod and "TEVKIF" not in ad_ and "SORUMLU" not in ad_:
                tevk_hesap_belirsiz.add(fatura_no)   # adı yalnız "KDV" — bir kez onaylanınca öğrenilir
            return kod

        # FAKTORİNG / BSMV (TEMELFATURA + ek vergi)
        # Matrah + BSMV toplamı TEK satır olarak 780.02'ye yazılır.
        # Kalemlerin varlığı önemsiz — TEMELFATURA + ek_vergi kesin faktoring senaryosu.
        if f["senaryo"] == "TEMELFATURA" and f["ek_vergi"] > 0:
            fak_kod = _bsmv_hesabi(hes_list)  # 780.02.xxx öncelikli
            if yon == "alis":
                fisler.append(sat(fak_kod, toplam, 0))
                fisler.append(sat(cari_kod or "198.01.001", 0, toplam))
            else:
                fisler.append(sat(cari_kod or "198.01.001", toplam, 0))
                fisler.append(sat(fak_kod, 0, toplam))
            fis += 1
            continue

        # NORMAL SATIŞ/ALIŞ FATURASI (çok KDV oranlı olabilir; tevkifatlı da bu yoldan:
        # gider/gelir ve KDV'nin TAMAMI normal yazılır, tevkifat ayrı satırda, cari ödenecek tutarla)
        fatura_bas = len(fisler)
        # Kalem yok ama toplam varsa (OCR gürültülü perakende fişi): tek satır
        # gider/gelir yaz. Kullanıcı Düzenle'de KDV'yi ayırabilir.
        if not f["kalemler"] and toplam > 0:
            if yon == "alis":
                fisler.append(sat(gider_kod or vars_gider, toplam, 0, rol="gider"))
            else:
                fisler.append(sat(gider_kod or vars_gider, 0, toplam, rol="gider"))
        # Tevkifatlı SATIŞ: tevkifat, ait olduğu oranın KDV'sinden düşülür; kalan KDV tek satırda
        # tevkifat hesabına (391.04.02 gibi) ALACAK yazılır — ayrı borç satırı yok.
        satis_net_oran = None
        if yon == "satis" and tevk > 0:
            kdvli = [k for k in f["kalemler"] if k["oran"] and k["kdv"] > 0]
            eslesen = [k for k in kdvli if any(abs(k["kdv"] * r - tevk) <= 0.05 + 0.0005 * k["kdv"] for r in TEVKIFAT_ORANLARI)]
            aday = (eslesen or sorted(kdvli, key=lambda k: -k["kdv"]))[:1]
            if aday and aday[0]["kdv"] + 0.01 >= tevk:
                satis_net_oran = aday[0]["oran"]
        for k in f["kalemler"]:
            oran = k["oran"]
            matrah = k["matrah"]; kdv = k["kdv"]
            if oran == 0:
                # KDV'siz kalem
                gider_yaz(0, matrah)
                continue
            if yon == "alis":
                gider_yaz(oran, matrah, f" (%{oran})")
                kdv_kod = gecmis_kdv_sec(oran, "alis")
                if not kdv_kod:
                    kdv_eksik.setdefault(f"%{oran} için 191", set()).add(fatura_no)
                fisler.append(sat(kdv_kod, kdv, 0, f" (%{oran} KDV)"))
            else:
                kdv_kod = gecmis_kdv_sec(oran, "satis")
                if not kdv_kod:
                    kdv_eksik.setdefault(f"%{oran} için 391", set()).add(fatura_no)
                gider_yaz(oran, matrah, f" (%{oran})")
                if oran == satis_net_oran:
                    satis_net_oran = ("yazildi", oran)
                    t_kod = tevkifat_hesap_sec("satis")
                    if not t_kod:
                        t_kod = kdv_kod          # ayrı tevkifat hesabı yok: net KDV normal 391'e
                        if kdv_kod:
                            tevkifat_kdvden.add(fatura_no)
                        else:
                            tevkifat_hesapsiz.add(fatura_no)
                    fisler.append(sat(t_kod, 0, round(kdv - tevk, 2), f" (%{oran} KDV − tevkifat)", rol="tevkifat"))
                else:
                    fisler.append(sat(kdv_kod, 0, kdv, f" (%{oran} KDV)"))

        # ek vergi (BSMV %5) — TTNET, faktoring karışık faturaları için
        if f["ek_vergi"] > 0:
            # geçmişte bu carinin ikinci gider hesabı (ör. ÖİV 689) varsa o; yoksa BSMV
            ek_k = gecmis_es.get("ek", "")
            bsmv_kod = ek_k if ek_k and ek_k in alt_kodlar else _bsmv_hesabi(hes_list)
            if yon == "alis":
                fisler.append(sat(bsmv_kod, f["ek_vergi"], 0))
            else:
                fisler.append(sat(bsmv_kod, 0, f["ek_vergi"]))

        # Denge emniyeti: liste toplamı yazılan kalemlerden fazlaysa (ör. Turkcell'de ÖİV
        # sütunu listede yok) fark kaybolup fiş dengesiz çıkıyordu. Fark carinin geçmişteki
        # ek vergi hesabına (yoksa gider hesabına) yazılır ve uyarılır.
        yazilan = round(sum((x["borc"] if yon == "alis" else x["alacak"]) for x in fisler[fatura_bas:]), 2)
        satis_net = isinstance(satis_net_oran, tuple)
        if satis_net:
            yazilan = round(yazilan + tevk, 2)     # KDV dahil karşılığı (tevkifat KDV satırından düşüldü)
        # Tevkifatlı faturada liste toplamı ya KDV dahil toplam ya da ödenecek tutardır
        # (entegratör listeleri/PDF: ödenecek). Denge KDV dahil toplamla kurulur.
        toplam_dahil = toplam
        if tevk > 0 and abs(toplam + tevk - yazilan) < abs(toplam - yazilan):
            toplam_dahil = round(toplam + tevk, 2)
        fark = round(toplam_dahil - yazilan, 2)
        if fark > 0.01:
            ek_k = gecmis_es.get("ek", "")
            fark_kod = ek_k if ek_k and ek_k in alt_kodlar else (gider_kod or vars_gider)
            fisler.append(sat(fark_kod, fark, 0) if yon == "alis" else sat(fark_kod, 0, fark))
            fark_yazilan.append(f"{fatura_no} ({fark:,.2f} → {fark_kod or 'boş'})")
        elif fark < -0.01:
            fark_dengesiz.append(f"{fatura_no} ({-fark:,.2f})")
            for x in fisler[fatura_bas:]:
                x["kontrol"] = list(dict.fromkeys((x.get("kontrol") or []) + ["dengesiz"]))

        # KDV tevkifatı:
        #  alış : KDV'nin tamamı 191'de (borç); tevkif edilen kısmı alıcı 2 No.lu beyanla
        #         öder -> 360 sorumlu KDV ALACAK; satıcıya (320) yalnız ödenecek tutar.
        #  satış: tevkifat KDV'den düşülür, kalan KDV 391 tevkifat hesabına (391.04.02) alacak
        #         (yukarıda KDV satırında); 120'ye ödenecek tutar.
        cari_tutar = toplam_dahil
        if tevk > 0:
            tevk_kod = tevkifat_hesap_sec(yon)
            if yon == "alis":
                if not tevk_kod:
                    tevkifat_hesapsiz.add(fatura_no)
                fisler.append(sat(tevk_kod, 0, tevk, " (KDV tevkifatı)", rol="tevkifat"))
            elif satis_net:
                pass        # tevkifat KDV satırından düşüldü, kalan KDV tevkifat hesabında
            else:
                if not tevk_kod:
                    # ayrı tevkifat hesabı yok: en büyük KDV'li oranın 391 hesabı borçlanır (net hesaplanan KDV)
                    kdv_satirlari = [x for x in fisler[fatura_bas:] if x.get("rol") == "kdv" and x["hesap"]]
                    if kdv_satirlari:
                        tevk_kod = max(kdv_satirlari, key=lambda x: x["alacak"])["hesap"]
                        tevkifat_kdvden.add(fatura_no)
                    else:
                        tevkifat_hesapsiz.add(fatura_no)
                fisler.append(sat(tevk_kod, tevk, 0, " (KDV tevkifatı)", rol="tevkifat"))
            cari_tutar = round(toplam_dahil - tevk, 2)
            tevkifatli.append(fatura_no)
            if tevk_tahmin:
                tevkifat_tahminli.append(f"{fatura_no} ({tevk:,.2f})")
        elif f.get("tur") == "TEVKIFAT":
            tevkifat_tutarsiz.append(fatura_no)

        # karşı taraf (tek satır)
        if yon == "alis":
            fisler.append(sat(cari_kod or "198.01.001", 0, cari_tutar))
        else:
            fisler.append(sat(cari_kod or "198.01.001", cari_tutar, 0))

        # Bu firmanın geçmişinde hiç kullanılmamış gider hesabı (yapay zekâ/varsayılan seçtiyse)
        if gider_kaynak in ("yz", "tahmin") and kullanilan_gider and gider_kod and gider_kod not in kullanilan_gider:
            yeni_hesap.setdefault(gider_kod, []).append(fatura_no)

        fis += 1

    kontrollu = sorted({s["evrak_no"] or s["fisno"] for s in fisler if s.get("kontrol")})
    if kontrollu:
        uyarilar.append(f"{len(kontrollu)} faturada doğrulanmamış hesap ya da tutar var (yapay zekâ/tahmin, benzer ad, KDV ayrılamadı, dengesiz) — "
                        f"sarı KONTROL işaretli satırları düzelt ya da ✓ ile onayla; onaylanmayanların "
                        f"Excel'de açıklamasına KONTROL yazılır: {_kisa_liste(kontrollu, 4)}")
    if kalem_bekleyen:
        uyarilar.append(f"{len(kalem_bekleyen)} faturanın kalem hesap önerileri yapay zekâda hazırlanıyor — "
                        f"birkaç dakika sonra tekrar İşle'ye basın")
    if tevkifatli:
        if yon == "alis":
            uyarilar.append(f"{len(tevkifatli)} tevkifatlı fatura: KDV'nin tamamı 191'e borç, tevkif edilen kısım "
                            f"360 sorumlu KDV'ye alacak, satıcıya ödenecek tutar yazıldı: {_kisa_liste(tevkifatli, 4)}")
        else:
            uyarilar.append(f"{len(tevkifatli)} tevkifatlı fatura: tevkifat KDV'den düşüldü, kalan KDV 391 tevkifat "
                            f"hesabına alacak, müşteriye ödenecek tutar yazıldı: {_kisa_liste(tevkifatli, 4)}")
    if tevkifat_tahminli:
        uyarilar.append(f"{len(tevkifat_tahminli)} faturada tevkifat listede yok; liste toplamı KDV dahil tutardan "
                        f"tevkifat payı kadar az olduğu için fark tevkifat sayıldı — PDF'ten kontrol et: "
                        f"{_kisa_liste(tevkifat_tahminli, 4)}")
    if tevkifat_tutarsiz:
        uyarilar.append(f"{len(tevkifat_tutarsiz)} fatura tevkifatlı görünüyor ama tevkifat tutarı okunamadı — "
                        f"fiş tevkifatsız yazıldı; faturanın PDF'ini ekle: {_kisa_liste(tevkifat_tutarsiz, 4)}")
    if tevkifat_hesapsiz:
        ne = "360 (sorumlu sıfatıyla ödenecek KDV)" if yon == "alis" else "391 tevkifat"
        uyarilar.append(f"Mizanda {ne} hesabı bulunamadı ({len(tevkifat_hesapsiz)} fatura) — tevkifat satırı boş "
                        f"hesapla bırakıldı; bir satırı doğru hesaba çekersen diğerleri de öğrenilir: "
                        f"{_kisa_liste(sorted(tevkifat_hesapsiz), 4)}")
    if tevkifat_kdvden:
        uyarilar.append(f"Mizanda ayrı 391 tevkifat hesabı yok: {len(tevkifat_kdvden)} faturada tevkifat düşülmüş KDV "
                        f"normal 391 KDV hesabına yazıldı: {_kisa_liste(sorted(tevkifat_kdvden), 4)}")
    if fark_yazilan:
        uyarilar.append(f"{len(fark_yazilan)} faturada liste toplamı kalemlerden fazlaydı (ÖİV/ÖTV gibi ek vergi "
                        f"sütunu olmayabilir); fark ayrı satıra yazıldı, kontrol et: {_kisa_liste(fark_yazilan, 4)}")
    if fark_dengesiz:
        uyarilar.insert(0, f"{len(fark_dengesiz)} faturada kalemler liste toplamından fazla — fiş DENGESİZ, "
                           f"aktarmadan önce düzelt: {_kisa_liste(fark_dengesiz, 4)}")
    for kod, nolar in yeni_hesap.items():
        uyarilar.append(f"{kod} {km.hesap_adi(kod)[:30]} bu firmanın geçmiş kayıtlarında hiç kullanılmamış "
                        f"({len(nolar)} fatura) — kontrol et: {_kisa_liste(nolar, 4)}")
    if kdv_eksik:
        onek_ = "191" if yon == "alis" else "391"
        mevcut = [f"{k} {a}" for k, a in km.hesaplar if k.startswith(onek_) and k in alt_kodlar]
        for ne, nolar in kdv_eksik.items():
            uyarilar.append(f"Mizanda {ne} KDV hesabı bulunamadı ({len(nolar)} fatura) — bu KDV satırları boş "
                            f"hesapla bırakıldı. Bir satırı doğru hesaba çekersen aynı oranlı tüm satırlar "
                            f"o hesaba geçer ve öğrenilir")
        uyarilar.append(f"Mizandaki {onek_} alt hesapları: " + ("; ".join(mevcut[:8]) + (" …" if len(mevcut) > 8 else "")
                                                               if mevcut else "hiç yok"))
    if not vars_gider and fisler:
        uyarilar.append("Mizanda gider/gelir için uygun alt hesap yok — eşleşmeyen gider satırları boş bırakıldı")
    if gider_bekleyen:
        uyarilar.insert(0, f"{len(gider_bekleyen)} cari için gider hesabı yapay zekâ ile seçiliyor; "
                           f"şimdilik varsayılan hesap yazıldı — bitince tekrar İşle'ye basın")
    return fisler, uyarilar


# ----------------------------------------------------------------- ÇEK
def isle_cek(hamlar, km, fis0):
    """
    Çek listesi -> muhasebe fişi.

    Beklenen Excel formatı (her satır bir çek):
      NO | ALIM TARİH | KİMDEN ALINDI | BANKA | ŞUBE | VADE | ÇEK NO | TUTARI | KİME VERİLDİ | ÇIKIŞ TARİH

    Her çek 2 fiş üretir:
      GİRİŞ (alım tarihinde): 101.01.001 borç / kimden_hesap alacak
      ÇIKIŞ (çıkış tarihinde): kime_hesap borç / 101.01.001 alacak

    KİME VERİLDİ = "KASA" veya boş ise sadece GİRİŞ fişi üretilir.
    """
    uyarilar = []
    CEK_HESAP = "101.01.001"

    # Çek satırlarını oku (özel parser — _kayitlar yerine doğrudan tablo)
    cekler_ham = _cek_listesi_oku(hamlar)
    # Mükerrer çek filtresi (çek no + tutar aynıysa tekini tut)
    gorulen_cek = set()
    cekler = []
    for c in cekler_ham:
        key = (c.get("cek_no",""), c["tutar"])
        if key in gorulen_cek:
            continue
        gorulen_cek.add(key)
        cekler.append(c)
    if len(cekler) < len(cekler_ham):
        uyarilar.append(f"{len(cekler_ham)-len(cekler)} mükerrer çek silindi")

    if not cekler:
        # Fallback: _kayitlar ile dene (basit format)
        kayitlar = _kayitlar(hamlar)
        if not kayitlar:
            uyarilar.append("Çek listesi okunamadı")
            return [], uyarilar
        # Basit format: her satır tek fiş (giriş varsayılır)
        fisler = []
        fis = fis0
        for k in sorted(kayitlar, key=lambda x: x["tarih"] or ""):
            fisno = f"{fis:05d}"
            cari, kaynak = km.eslestir(k["aciklama"])
            tutar = abs(k["tutar"])
            if tutar == 0: continue
            fisler.append(_sat(fisno, k["tarih"], "ÇEK GİRİŞİ", CEK_HESAP, tutar, 0,
                evrak_no=k.get("referans",""), detay=k["aciklama"], kaynak="cek"))
            fisler.append(_sat(fisno, k["tarih"], "ÇEK GİRİŞİ", cari or "", 0, tutar,
                evrak_no=k.get("referans",""), detay=k["aciklama"], kaynak=kaynak))
            fis += 1
        return fisler, uyarilar

    fisler = []
    fis = fis0

    for c in cekler:
        tutar = c["tutar"]
        if tutar == 0:
            continue
        cek_no = c.get("cek_no", "")
        banka = c.get("banka", "")
        vade = c.get("vade", "")
        kimden = c.get("kimden", "")
        kime = c.get("kime", "")
        alim_tarih = c.get("alim_tarih", "")
        cikis_tarih = c.get("cikis_tarih", "")

        # Detay formatı
        alim_fmt = _tarih_gg(alim_tarih)
        vade_fmt = _tarih_gg(vade)
        kimden_temiz = _isim_temizle(kimden)
        kime_temiz = _isim_temizle(kime)

        # -- GİRİŞ FİŞİ (alım tarihinde) --
        giris_detay = f"{alim_fmt}-{banka}-{vade_fmt}VDLİ-{kimden_temiz}" if alim_fmt and banka else kimden
        kimden_hesap, kimden_kaynak = km.eslestir(kimden)
        if not kimden_hesap:
            kimden_hesap = "198.01.001"
            uyarilar.append(f"GİRİŞ {kimden[:25]}: hesap eşleşmedi")

        fisno = f"{fis:05d}"
        fisler.append(_sat(fisno, alim_tarih, "ÇEK GİRİŞİ", CEK_HESAP, tutar, 0,
            evrak_no=cek_no, detay=giris_detay, kaynak="cek"))
        fisler.append(_sat(fisno, alim_tarih, "ÇEK GİRİŞİ", kimden_hesap, 0, tutar,
            evrak_no=cek_no, detay=giris_detay, kaynak=kimden_kaynak))
        fis += 1

        # -- ÇIKIŞ FİŞİ (çıkış tarihinde) -- sadece kime verilmişse
        kime_upper = kime.strip().upper()
        if kime_upper and kime_upper not in ("KASA", "KASAYA", "KASADA", ""):
            if not cikis_tarih:
                cikis_tarih = alim_tarih  # çıkış tarihi yoksa alım tarihi kullan

            cikis_detay = f"{_tarih_gg(cikis_tarih)}-{banka}-{vade_fmt}VDLİ-{kimden_temiz}-{kime_temiz}" if banka else f"{kimden_temiz}-{kime_temiz}"

            # Özel durumlar: FİRMA ORTAĞINA ÇIKIŞ / ORTAKTAN GİRİŞ
            kime_arama = kime
            if "ORTA" in kime_upper and ("ÇIKIŞ" in kime_upper or "CIKIS" in kime_upper):
                kime_arama = "ŞİRKET ORTAĞINA ÇIKIŞ"
            elif "FAKTORİNG" in kime_upper or "FAKTORING" in kime_upper:
                kime_arama = kime  # faktoring şirketi adıyla ara

            kime_hesap, kime_kaynak = km.eslestir(kime_arama)
            if not kime_hesap:
                kime_hesap = "198.01.001"
                uyarilar.append(f"ÇIKIŞ {kime[:25]}: hesap eşleşmedi")

            fisno = f"{fis:05d}"
            fisler.append(_sat(fisno, cikis_tarih, "ÇEK ÇIKIŞI", kime_hesap, tutar, 0,
                evrak_no=cek_no, detay=cikis_detay, kaynak=kime_kaynak))
            fisler.append(_sat(fisno, cikis_tarih, "ÇEK ÇIKIŞI", CEK_HESAP, 0, tutar,
                evrak_no=cek_no, detay=cikis_detay, kaynak="cek"))
            fis += 1

    return fisler, uyarilar


def _tarih_gg(tarih_str):
    """ISO tarih (2026-07-01) → gg.aa.yyyy"""
    if not tarih_str: return ""
    import re
    m = re.match(r'(\d{4})-(\d{2})-(\d{2})', str(tarih_str))
    if m: return f"{m.group(3)}.{m.group(2)}.{m.group(1)}"
    m = re.match(r'(\d{1,2})[./](\d{1,2})[./](\d{4})', str(tarih_str))
    if m: return f"{int(m.group(1)):02d}.{int(m.group(2)):02d}.{m.group(3)}"
    return str(tarih_str)


def _isim_temizle(isim):
    """İsimden boşlukları kaldır, büyük harf."""
    if not isim: return ""
    return str(isim).strip().upper().replace(" ", "")


def _cek_listesi_oku(hamlar):
    """Çek listesi Excel'inden satırları özel formatta okur.
    Sütunlar: NO | ALIM TARİH | KİMDEN ALINDI | BANKA | ŞUBE | VADE | ÇEK NO | TUTARI | KİME VERİLDİ | ÇIKIŞ TARİH
    """
    import re as _re
    from datetime import datetime

    def _tarih_iso(v):
        if not v: return ""
        if isinstance(v, datetime): return v.strftime("%Y-%m-%d")
        s = str(v).strip()
        m = _re.match(r'(\d{1,2})[./](\d{1,2})[./](\d{4})', s)
        if m: return f"{m.group(3)}-{int(m.group(2)):02d}-{int(m.group(1)):02d}"
        return s

    def _sayi(v):
        if v is None: return 0
        if isinstance(v, (int, float)): return float(v)
        try: return float(str(v).replace(",", ".").replace(" ", ""))
        except: return 0

    cekler = []
    for h in hamlar:
        for tab in h.get("tablolar", []):
            # Başlık satırını bul
            harita = {}
            baslik_r = -1
            for ri, row in enumerate(tab[:10]):
                for ci, c in enumerate(row):
                    if c is None: continue
                    cn = norm(str(c))
                    if "ALIM" in cn and "TARIH" in cn: harita["alim_tarih"] = ci
                    elif "KIMDEN" in cn or "ALINDI" in cn: harita["kimden"] = ci
                    elif cn.strip() in ("BANKA", "BANKA "): harita["banka"] = ci
                    elif "SUBE" in cn: harita["sube"] = ci
                    elif "VADE" in cn: harita["vade"] = ci
                    elif "CEK NO" in cn or "CEK NUMARASI" in cn or "SENET NO" in cn: harita["cek_no"] = ci
                    elif "TUTAR" in cn or "BEDEL" in cn: harita["tutar"] = ci
                    elif "KIME" in cn or "VERILDI" in cn: harita["kime"] = ci
                    elif "CIKIS" in cn and "TARIH" in cn: harita["cikis_tarih"] = ci
                    elif cn == "NO": harita["no"] = ci
                if "kimden" in harita and "tutar" in harita:
                    baslik_r = ri; break

            if baslik_r < 0:
                continue

            def g(row, alan):
                ci = harita.get(alan)
                if ci is not None and ci < len(row): return row[ci]
                return None

            for ri in range(baslik_r + 1, len(tab)):
                row = tab[ri]
                tutar = _sayi(g(row, "tutar"))
                kimden = str(g(row, "kimden") or "").strip()
                if tutar == 0 and not kimden: continue
                # TOPLAM satırını atla
                if "TOPLAM" in norm(kimden): continue

                cekler.append({
                    "alim_tarih": _tarih_iso(g(row, "alim_tarih")),
                    "kimden": kimden,
                    "banka": str(g(row, "banka") or "").strip().upper(),
                    "sube": str(g(row, "sube") or "").strip(),
                    "vade": _tarih_iso(g(row, "vade")),
                    "cek_no": str(g(row, "cek_no") or "").strip(),
                    "tutar": tutar,
                    "kime": str(g(row, "kime") or "").strip(),
                    "cikis_tarih": _tarih_iso(g(row, "cikis_tarih")),
                })
    return cekler


def isle_fis(kalemler, km, fis0=1, karsi_hesap="198.01.001"):
    """
    Masraf-server tarzı elle girilen fiş kalemlerini muhasebe fişi satırlarına çevirir.
    kalemler: [{tarih, no, firma, matrah, oranFactor, kdv, toplam, kod, ad, grup}]
    Aynı grup = tek fiş no; her kalem gider+KDV borç, tek karşı(198) alacak.
    """
    from collections import OrderedDict
    uyarilar = []
    hes_list = [{"kod": k, "ad": a} for k, a in km.hesaplar]

    # gruplama: grup > no > tarih
    gruplar = OrderedDict()
    for k in kalemler:
        anahtar = k.get("grup") or k.get("no") or k.get("tarih") or "_"
        gruplar.setdefault(anahtar, []).append(k)

    fisler = []
    fis = fis0
    for anahtar, kalem_grubu in gruplar.items():
        ilk = kalem_grubu[0]
        tarih = ilk.get("tarih", "")
        firma = ilk.get("firma", "")
        evrak_no = ilk.get("no", "")
        try:
            y, m, dd = map(int, tarih.split("-"))
            fisno = f"{y}{m:02d}{dd:02d}.{fis}"
        except Exception:
            fisno = f"{tarih or 'FIS'}.{fis}"

        def sat(hesap, borc, alacak):
            return _sat(fisno, tarih, firma, hesap, borc, alacak,
                        evrak_no=evrak_no, detay=firma, kaynak="fis")

        fis_toplam = 0.0
        for k in kalem_grubu:
            gider_kod = k.get("kod", "")
            toplam = round(float(k.get("toplam") or 0), 2)
            oranF = float(k.get("oranFactor") or 1)
            oran_pct = round((oranF - 1) * 100)
            matrah = round(float(k.get("matrah") or (toplam / oranF if oranF else toplam)), 2)
            kdv = round(float(k.get("kdv") or (toplam - matrah)), 2)
            fis_toplam = round(fis_toplam + toplam, 2)

            if not gider_kod:
                uyarilar.append(f"{firma[:24]}: gider hesabı boş")
            if kdv > 0:
                kdv_kod = _kdv_hesabi_ad(hes_list, oran_pct, "alis")
                if not kdv_kod:
                    uyarilar.append(f"{firma[:24]}: mizanda %{oran_pct} için 191 indirilecek KDV hesabı yok — satır boş hesapla bırakıldı")
                fisler.append(sat(gider_kod, matrah, 0))
                fisler.append(sat(kdv_kod, kdv, 0))
            else:
                fisler.append(sat(gider_kod, toplam, 0))

        # fiş başına tek karşı (alacak) satırı — dengeyi kurar
        fisler.append(sat(karsi_hesap, 0.0, fis_toplam))
        fis += 1
    return fisler, uyarilar


def isle(tip, hamlar, km, fis0, yon="alis", pdf_faturalar=None, gider_ogrenme=None, gider_onerici=None,
         kalem_onerici=None):
    if tip == "banka":
        return isle_banka(hamlar, km, fis0)
    if tip == "fatura":
        return isle_fatura(hamlar, km, fis0, yon=yon, pdf_faturalar=pdf_faturalar,
                           gider_ogrenme=gider_ogrenme, gider_onerici=gider_onerici, kalem_onerici=kalem_onerici)
    if tip == "cek":
        return isle_cek(hamlar, km, fis0)
    if tip == "fis":
        # fis sekmesi için hamlar bekelenmiyor; ayrı uç kullanılıyor
        return [], ["Fiş sekmesi için /isle-fis kullanın"]
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
        evno = str(evno) if evno else ""
        ws.append([
            r.get("fisno", ""), dt(r.get("fis_tarih", "")), r.get("fis_aciklama", ""),
            r.get("hesap", ""), evno, dt(r.get("evrak_tarih", "")), r.get("detay", ""),
            r.get("borc") or None, r.get("alacak") or None, None,
            r.get("belge_turu", "MF"),
            r.get("para_birimi", "") or "",
            r.get("kur") or None,
            r.get("doviz_tutar") or None,
        ])
    sfill = PatternFill("solid", start_color="FFF3CD")
    for row in ws.iter_rows(min_row=2, max_row=len(satirlar) + 1):
        for c in row:
            c.border = bd
        row[1].number_format = "dd/mm/yyyy"; row[5].number_format = "dd/mm/yyyy"
        row[4].number_format = "@"  # Evrak No metin
        row[7].number_format = "#,##0.00"; row[8].number_format = "#,##0.00"
        if not row[3].value:
            row[3].fill = sfill
    for col, w in zip("ABCDEFGHIJKLMN", [10, 12, 26, 14, 20, 12, 30, 13, 13, 8, 10, 10, 8, 12]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    bio = io.BytesIO(); wb.save(bio); bio.seek(0)
    return bio


KONTROL_ON = "KONTROL - "


def kontrol_isaretle(satirlar: list) -> list:
    """Dışa aktarımdan önce: içinde doğrulanmamış hesap ('kontrol') olan fişlerin
    tüm satırlarında fiş açıklamasının, işaretli satırların detay açıklamasının
    başına 'KONTROL - ' yazar. Logo'da fiş listesinde aranıp bulunabilsin diye."""
    isaretli = {r.get("fisno") for r in satirlar if r.get("kontrol")}
    out = []
    for r in satirlar:
        r = dict(r)
        if r.get("fisno") in isaretli:
            fa = str(r.get("fis_aciklama") or "")
            if not fa.startswith(KONTROL_ON):
                r["fis_aciklama"] = KONTROL_ON + fa
            if r.get("kontrol"):
                dt = str(r.get("detay") or "")
                if not dt.startswith(KONTROL_ON):
                    r["detay"] = KONTROL_ON + dt
        out.append(r)
    return out


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
