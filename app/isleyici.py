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
from app.kurallar import norm


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
def _sat(fisno, tarih, aciklama, hesap, borc, alacak, evrak_no="", detay="", kaynak=""):
    return {
        "fisno": fisno, "fis_tarih": tarih, "fis_aciklama": aciklama,
        "hesap": hesap, "evrak_no": evrak_no, "evrak_tarih": tarih,
        "detay": detay or aciklama, "borc": round(borc, 2), "alacak": round(alacak, 2),
        "belge_turu": "MF", "kaynak": kaynak,
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
    """Önce tablo, tablo yoksa ham metin (OCR) satırları."""
    k = _tablo_satirlari(hamlar)
    if not k:
        k = _ham_metin_satirlari(hamlar)
    return k


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
    Banka dökümü -> her hareket bir fiş.
    Bankaya para GİRİŞİ (alacak, +): banka borç / karşı alacak
    Bankadan ÇIKIŞ (borç, -): karşı borç / banka alacak
    Her belge için AYRI banka hesabı akıllı seçilir (birden çok banka
    ekstresi aynı anda yüklenmişse her biri kendi 102 hesabına gider).
    """
    uyarilar = []
    hesaplar_list = [{"kod": k, "ad": a} for k, a in km.hesaplar]

    # Her belge için ayrı banka hesabı belirle
    dosya_to_hesap = {}
    for h in hamlar:
        dosya = h.get("dosya", "")
        # Tek dosyayı liste olarak vererek _banka_hesabi_bul'u çalıştır
        hesap = _banka_hesabi_bul([h], hesaplar_list, uyarilar)
        dosya_to_hesap[dosya] = hesap

    # Tüm kayıtları topla, her kaydın hangi dosyadan geldiğini bildiği için
    # o dosyanın banka hesabına yazılır
    kayitlar = _kayitlar(hamlar)
    if not kayitlar:
        for h in hamlar:
            if h.get("ham_metin"):
                uyarilar.append(f"{h.get('dosya')}: tablo çıkarılamadı, ham metin var — elle düzenleme gerekebilir")
        return [], uyarilar

    fisler = []
    fis = fis0
    for k in sorted(kayitlar, key=lambda x: x["tarih"] or ""):
        fisno = f"{fis:05d}"
        # Bu kaydın geldiği dosyayı belirle, ona ait banka hesabını al
        banka_hesap = dosya_to_hesap.get(k.get("dosya", ""), "102.01.001")
        banka_ad = km.hesap_adi(banka_hesap) or "BANKA"

        karsi, kaynak = km.eslestir(k["aciklama"])
        if not karsi:
            karsi = ""  # boş bırak, önizlemede sarı/uyarı
            uyarilar.append(f"{k['aciklama'][:30]}: hesap eşleşmedi")
        tutar = abs(k["tutar"])
        evno = k.get("referans", "")
        if k["tutar"] >= 0:  # giriş: banka borç / karşı alacak
            fisler.append(_sat(fisno, k["tarih"], banka_ad, banka_hesap, tutar, 0, evrak_no=evno, detay=k["aciklama"], kaynak="banka"))
            fisler.append(_sat(fisno, k["tarih"], banka_ad, karsi, 0, tutar, evrak_no=evno, detay=k["aciklama"], kaynak=kaynak))
        else:               # çıkış: karşı borç / banka alacak
            fisler.append(_sat(fisno, k["tarih"], banka_ad, karsi, tutar, 0, evrak_no=evno, detay=k["aciklama"], kaynak=kaynak))
            fisler.append(_sat(fisno, k["tarih"], banka_ad, banka_hesap, 0, tutar, evrak_no=evno, detay=k["aciklama"], kaynak="banka"))
        fis += 1
    return fisler, uyarilar


# ----------------------------------------------------------------- FATURA
def _elogo_fatura_satirlari(hamlar, yon="alis"):
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
                nrow = [norm(str(c)) for c in row]
                for j, c in enumerate(nrow):
                    if c == "FATURA NO":
                        sut["fatura_no"] = j
                    elif c == "FATURA TARIHI":
                        sut["tarih"] = j
                    elif c in ("GONDERICI ADI", "ALICI ADI"):
                        sut["cari_ad"] = j
                    elif c == "SENARYO":
                        sut["senaryo"] = j
                    elif c in ("TUR",):
                        sut["tur"] = j
                    elif c == "TOPLAM TUTAR":
                        sut["toplam"] = j
                    elif c == "KDV TOPLAMI":
                        sut["kdv_top"] = j
                    elif c == "KDV 1":
                        sut["kdv_1"] = j
                    elif c == "KDV 8":
                        sut["kdv_8"] = j
                    elif c == "KDV 10":
                        sut["kdv_10"] = j
                    elif c == "KDV 18":
                        sut["kdv_18"] = j
                    elif c == "KDV 20":
                        sut["kdv_20"] = j
                    elif c == "KDV 1 MATRAH":
                        sut["mat_1"] = j
                    elif c == "KDV 8 MATRAH":
                        sut["mat_8"] = j
                    elif c == "KDV 10 MATRAH":
                        sut["mat_10"] = j
                    elif c == "KDV 18 MATRAH":
                        sut["mat_18"] = j
                    elif c == "KDV 20 MATRAH":
                        sut["mat_20"] = j
                    elif c == "TEVKIFAT TOPLAMI":
                        sut["tevkifat"] = j
                    elif c == "ILK TEVKIFAT KODU":
                        sut["tevkifat_kod"] = j
                    elif c == "EK VERGILER":
                        sut["ek_vergi"] = j
                if "fatura_no" in sut and "tarih" in sut and "cari_ad" in sut:
                    bas_idx = i; break
            if bas_idx is None:
                continue

            def g(row, key):
                j = sut.get(key)
                return row[j] if j is not None and j < len(row) else None

            for row in tablo[bas_idx + 1:]:
                fno = g(row, "fatura_no")
                if not fno or not str(fno).strip():
                    continue
                cari = str(g(row, "cari_ad") or "").strip()
                if not cari:
                    continue
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

                toplam = _sayi(g(row, "toplam")) or 0
                tevkifat = _sayi(g(row, "tevkifat")) or 0
                tevkifat_kod = str(g(row, "tevkifat_kod") or "").strip()
                ek_vergi = _sayi(g(row, "ek_vergi")) or 0
                tur = str(g(row, "tur") or "SATIS").strip().upper()
                senaryo = str(g(row, "senaryo") or "").strip().upper()

                # toplam sütunu yoksa (satış Excel'i) kalemlerden hesapla
                if toplam == 0 and kalemler:
                    toplam = round(sum(k["matrah"] + k["kdv"] for k in kalemler), 2)
                    # tevkifat varsa toplama eklenmesi lazım (KDV zaten hesaplı ama tevkifat toplam düşüldüğünde alıcının ödeyeceği)
                    # ödenecek = toplam - tevkifat_kdv (alıcı tevkifatı direkt vergiye öder)
                # ek vergiler (BSMV vb.) hala 0 olabilir; eğer toplam 0 hala ise
                if toplam == 0 and ek_vergi > 0:
                    toplam = ek_vergi

                # eğer hiç kalem yoksa ama toplam varsa (faktoring/BSMV): tek kalem KDV=0
                if not kalemler and toplam > 0:
                    kalemler.append({"oran": 0, "matrah": round(toplam - ek_vergi, 2), "kdv": 0})

                faturalar.append({
                    "fatura_no": str(fno).strip(),
                    "tarih": _tarih_iso(g(row, "tarih")),
                    "cari_ad": cari,
                    "tur": tur, "senaryo": senaryo,
                    "toplam": round(toplam, 2),
                    "kalemler": kalemler,
                    "tevkifat": round(tevkifat, 2),
                    "tevkifat_kod": tevkifat_kod,
                    "ek_vergi": round(ek_vergi, 2),
                    "yon": yon, "dosya": h.get("dosya", ""),
                })
    return faturalar


def _kdv_hesabi_ad(kod_ad_list, oran, yon="alis"):
    """191 (alış indirilecek) veya 391 (satış hesaplanan) hesabını orana göre bulur.
    "SORUMLU/IADE/ITHAL" hariç. Bulamazsa boş döner."""
    if oran <= 0:
        return ""
    prefix = "191" if yon == "alis" else "391"
    YASAK = ("SORUMLU", "IADE", "ITHAL", "VAZGEC", "TEVKIF")
    oran_str = str(oran)
    adaylar = []
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
        adaylar.sort(reverse=True)
        return adaylar[0][1]
    return ""


def _tevkifat_hesabi(kod_ad_list, yon="satis"):
    """Tevkifat KDV hesabını bulur. Satışta 391.02 gibi 'HESAPLANAN...TEVKIFAT',
    alışta 191.03 gibi 'SORUMLU SIFATIYLA ODENEN KDV'."""
    if yon == "satis":
        prefix = "391"; ara = ("TEVKIF", "HESAPLANAN")
    else:
        prefix = "191"; ara = ("SORUMLU", "TEVKIF")
    adaylar = []
    for h in kod_ad_list:
        kod = h["kod"]
        if not kod.startswith(prefix):
            continue
        nad = norm(h["ad"])
        if any(a in nad for a in ara):
            adaylar.append((kod.count("."), kod))
    if adaylar:
        adaylar.sort(reverse=True)
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


def isle_fatura(hamlar, km, fis0, yon="alis"):
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

    faturalar = _elogo_fatura_satirlari(hamlar, yon)
    if not faturalar:
        # eLogo listesi yok — OCR ham metninden perakende/e-fatura fotoğrafı olabilir
        faturalar = _ham_metin_faturalar(hamlar, yon)
    if not faturalar:
        for h in hamlar:
            if h.get("ham_metin"):
                uyarilar.append(f"{h.get('dosya')}: fatura verisi çıkarılamadı (OCR gürültülü olabilir) — elle düzeltme gerekebilir")
        return [], uyarilar

    fisler = []
    fis = fis0
    for f in sorted(faturalar, key=lambda x: x["tarih"] or ""):
        fisno = f"{fis:05d}"
        cari_ad = f["cari_ad"]
        tarih = f["tarih"]
        fatura_no = f["fatura_no"]
        toplam = f["toplam"]

        # FATURA HESAPLARI: firma kuralı yerine geçmiş fişler referans alınır.
        # Aynı cari daha önce hangi 320/120 ve hangi gider/gelir hesabıyla
        # kullanılmışsa onu tercih ederiz. Geçmişte kayıt yoksa varsayılana düşeriz.
        gecmis_es = km.fatura_gecmis_eslestir(cari_ad, yon)
        cari_kod = gecmis_es.get("cari", "")
        gider_kod = gecmis_es.get("ana", "")
        cari_kaynak = gecmis_es.get("kaynak", "")
        gecmis_kdv = gecmis_es.get("kdv", [])

        # varsayılanlar
        vars_cari = "320.01.001" if yon == "alis" else "120.01.001"
        vars_gider = "740.01.001" if yon == "alis" else "600.01.001"

        # cari doğru öneke uymuyorsa varsayılan
        if yon == "alis":
            if not cari_kod or not cari_kod.startswith(("320", "329", "331", "335")):
                cari_kod = vars_cari
            # gider cari'yle aynıysa (aynı öğrenme sonucu) veya 320'yse varsayılana düş
            if not gider_kod or gider_kod == cari_kod or gider_kod.startswith(("320", "120", "100", "102")):
                gider_kod = vars_gider
        else:  # satış
            if not cari_kod or not cari_kod.startswith(("120", "121")):
                cari_kod = vars_cari
            if not gider_kod or gider_kod == cari_kod or gider_kod.startswith(("120", "320")):
                gider_kod = vars_gider  # 600 gelir

        def sat(hesap, borc, alacak, detay_ek=""):
            # Açıklama alanları SADECE firma adı (KDV oranı, "(faktoring)" gibi ekler yok).
            return _sat(fisno, tarih, cari_ad, hesap, borc, alacak,
                        evrak_no=fatura_no, detay=cari_ad, kaynak=cari_kaynak or "fatura")

        def gecmis_kdv_sec(oran, yon):
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

        # FAKTORİNG / BSMV (TEMELFATURA + ek vergi)
        # Matrah + BSMV toplamı TEK satır olarak 780.02'ye yazılır.
        # Kalemlerin varlığı önemsiz — TEMELFATURA + ek_vergi kesin faktoring senaryosu.
        if f["senaryo"] == "TEMELFATURA" and f["ek_vergi"] > 0:
            fak_kod = _bsmv_hesabi(hes_list)  # 780.02.xxx öncelikli
            if yon == "alis":
                fisler.append(sat(fak_kod, toplam, 0))
                fisler.append(sat(cari_kod or "320.01.001", 0, toplam))
            else:
                fisler.append(sat(cari_kod or "120.01.001", toplam, 0))
                fisler.append(sat(fak_kod, 0, toplam))
            fis += 1
            continue

        # TEVKIFATLI fatura
        if f["tur"] == "TEVKIFAT" and f["tevkifat"] > 0:
            # kalemler normal (gider matrah + normal KDV), sonra tevkifat kısmı ayrı hesaba
            for k in f["kalemler"]:
                oran = k["oran"]
                matrah = k["matrah"]; kdv = k["kdv"]
                if yon == "alis":
                    # alışta: gider + %20 KDV borç
                    fisler.append(sat(gider_kod or "740.01.001", matrah, 0, f" (%{oran})"))
                    kdv_kod = gecmis_kdv_sec(oran, "alis") or "191.02"
                    fisler.append(sat(kdv_kod, kdv, 0, f" (%{oran} KDV)"))
                else:  # satış
                    kdv_kod = gecmis_kdv_sec(oran, "satis") or "391.02"
                    fisler.append(sat("600.01.001", 0, matrah, f" (%{oran})"))
                    fisler.append(sat(kdv_kod, 0, kdv, f" (%{oran} KDV)"))

            # tevkifat: satışta 391 alacaktan düşülür (biz KDV'nin bir kısmını
            # tahsil etmedik, alıcı direkt vergiye ödedi); alışta ise sorumlu KDV.
            tevk_kod = _tevkifat_hesabi(hes_list, yon)
            if yon == "alis":
                # alışta sorumlu sıfatıyla ödenecek KDV borç
                if tevk_kod:
                    fisler.append(sat(tevk_kod, f["tevkifat"], 0, " (tevkifat KDV)"))
                fisler.append(sat(cari_kod or "320.01.001", 0, toplam))
            else:
                # satışta: tevkifat KDV BORÇ (bizde iade edilecek KDV)
                # Ödenecek Tutar = toplam - tevkifat, bu kadar 120 borç
                if tevk_kod:
                    fisler.append(sat(tevk_kod, f["tevkifat"], 0, " (tevkifat/iade KDV)"))
                odenecek = round(toplam - f["tevkifat"], 2)
                fisler.append(sat(cari_kod or "120.01.001", odenecek, 0, " (ödenecek)"))
            fis += 1
            continue

        # NORMAL SATIŞ/ALIŞ FATURASI (çok KDV oranlı olabilir)
        # Kalem yok ama toplam varsa (OCR gürültülü perakende fişi): tek satır
        # gider/gelir yaz. Kullanıcı Düzenle'de KDV'yi ayırabilir.
        if not f["kalemler"] and toplam > 0:
            if yon == "alis":
                fisler.append(sat(gider_kod or "", toplam, 0))
            else:
                fisler.append(sat("600.01.001", 0, toplam))
        for k in f["kalemler"]:
            oran = k["oran"]
            matrah = k["matrah"]; kdv = k["kdv"]
            if oran == 0:
                # KDV'siz kalem
                if yon == "alis":
                    fisler.append(sat(gider_kod or "740.01.001", matrah, 0))
                else:
                    fisler.append(sat("600.01.001", 0, matrah))
                continue
            if yon == "alis":
                fisler.append(sat(gider_kod or "740.01.001", matrah, 0, f" (%{oran})"))
                kdv_kod = gecmis_kdv_sec(oran, "alis")
                if not kdv_kod:
                    uyarilar.append(f"{cari_ad[:20]}: %{oran} için 191 indirilecek KDV bulunamadı")
                    kdv_kod = "191.02"
                fisler.append(sat(kdv_kod, kdv, 0, f" (%{oran} KDV)"))
            else:
                kdv_kod = gecmis_kdv_sec(oran, "satis")
                if not kdv_kod:
                    uyarilar.append(f"{cari_ad[:20]}: %{oran} için 391 hesaplanan KDV bulunamadı")
                    kdv_kod = "391.02"
                fisler.append(sat("600.01.001", 0, matrah, f" (%{oran})"))
                fisler.append(sat(kdv_kod, 0, kdv, f" (%{oran} KDV)"))

        # ek vergi (BSMV %5) — TTNET, faktoring karışık faturaları için
        if f["ek_vergi"] > 0:
            bsmv_kod = _bsmv_hesabi(hes_list)
            if yon == "alis":
                fisler.append(sat(bsmv_kod, f["ek_vergi"], 0))
            else:
                fisler.append(sat(bsmv_kod, 0, f["ek_vergi"]))

        # karşı taraf (tek satır)
        if yon == "alis":
            fisler.append(sat(cari_kod or "320.01.001", 0, toplam))
        else:
            fisler.append(sat(cari_kod or "120.01.001", toplam, 0))

        fis += 1

    return fisler, uyarilar


# ----------------------------------------------------------------- ÇEK
def isle_cek(hamlar, km, fis0):
    """
    Çek listesi -> çek portföyü/borç senetleri. Excel bekler.
    İskelet: tutar + vade + keşideci; 101 çek / 320 cari gibi.
    """
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
        # alınan çek: 101 çekler borç / 120 alıcı alacak (varsayım)
        fisler.append(_sat(fisno, k["tarih"], k["aciklama"], "101.01.001", tutar, 0, detay=k["aciklama"], kaynak="cek"))
        fisler.append(_sat(fisno, k["tarih"], k["aciklama"], cari or "120.01.001", 0, tutar, detay=k["aciklama"], kaynak=kaynak))
        fis += 1
    return fisler, uyarilar


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
                    uyarilar.append(f"{firma[:24]}: %{oran_pct} için 191 indirilecek KDV bulunamadı")
                    kdv_kod = "191.02"
                fisler.append(sat(gider_kod, matrah, 0))
                fisler.append(sat(kdv_kod, kdv, 0))
            else:
                fisler.append(sat(gider_kod, toplam, 0))

        # fiş başına tek karşı (alacak) satırı — dengeyi kurar
        fisler.append(sat(karsi_hesap, 0.0, fis_toplam))
        fis += 1
    return fisler, uyarilar


def isle(tip, hamlar, km, fis0, yon="alis"):
    if tip == "banka":
        return isle_banka(hamlar, km, fis0)
    if tip == "fatura":
        return isle_fatura(hamlar, km, fis0, yon=yon)
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
            r.get("belge_turu", "MF"), "", "", "",
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
