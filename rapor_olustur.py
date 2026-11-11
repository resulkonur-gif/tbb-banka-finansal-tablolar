#!/usr/bin/env python3
"""
tbb_cek.py çıktısından (tbb_veri.csv) etkileşimli HTML raporu ve düz tablo üretir.

Kullanım:
    python rapor_olustur.py                          # tbb_veri.csv -> Turk_Bankacilik_Sistemi.html
    python rapor_olustur.py --girdi tbb_veri.csv --html rapor.html --uzun tbb_uzun.csv
    python rapor_olustur.py --excel tbb_veri.xlsx    # (pandas + openpyxl gerekir)

Yalnız Python standart kütüphanesi yeterlidir (Excel çıktısı hariç).
"""
import argparse
import csv
import json
import re
import statistics
import sys
from collections import defaultdict
from pathlib import Path

HERE = Path(__file__).resolve().parent

# Tablo kökü -> rapordaki tablo adı
TABLO_ADI = {
    "102087": "1. Varlıklar", "102086": "2. Yükümlülükler", "102519": "3. Nazım Hesaplar",
    "102567": "4. Gelir-Gider Tablosu", "16039": "5. Kar veya Zarar ve DKG", "102404": "7. Nakit Akış Tablosu",
}
LIKIDITE = {
    "solo": {"109460": "hqla", "109463": "out", "109466": "lcr", "109467": "lcrYp", "17130": "asf", "17131": "rsf", "17144": "nsfr"},
    "kons": {"359460": "hqla", "359463": "out", "359466": "lcr", "359467": "lcrYp", "487130": "asf", "487131": "rsf", "487144": "nsfr"},
}
# Dipnot kökleri (solo kodları): mevduat vade yapısı ve özkaynak kalemleri
DIPNOT_ADI = {"114171": "8. Mevduat Vade Yapısı", "109865": "9. Yasal Özkaynak"}   # 109865: özkaynak tablosunun "Tutar" sütunu
# Sermaye yeterliliği oranları (yüzde); kalem listesine girmez, ayrı dizi olarak tutulur
SERMAYE_ORAN = {"159980": "cet1", "159981": "t1", "159982": "car", "159984": "buf",
                "159985": "ccb", "159986": "ccyb", "109885": "sib", "159987": "bufR"}
MEVDUAT_TUR = [("114173", "Tasarruf"), ("114181", "Döviz tevdiat"), ("114206", "Ticari kuruluş"), ("114198", "Resmi kuruluş"),
               ("114214", "Diğer kuruluş"), ("114222", "Kıymetli maden"), ("114230", "Bankalar")]
VADE = ["Vadesiz", "7 Gün İhbarlı", "1Aya Kadar", "1-3 Ay", "3-6 Ay", "6 Ay-1 Yıl", "1 Yıl ve Üstü", "Birikimli Mevduat"]
VADE_ETIKET = {"1Aya Kadar": "1 aya kadar", "7 Gün İhbarlı": "7 gün ihbarlı", "1 Yıl ve Üstü": "1 yıl ve üstü", "Birikimli Mevduat": "Birikimli"}
PAY_KOD = {"cet1": "159935", "t1": "159953", "car": "159977"}   # oranın pay kalemi (payda: risk ağırlıklı tutarlar 159978)
TAMPON = {"buf": 6, "ccb": 6, "ccyb": 6, "bufR": 100}            # yüzde olarak makul üst sınır
ORANLAR = ("lcr", "lcrYp", "nsfr", "cet1", "t1", "car", "buf", "ccb", "ccyb", "sib", "bufR")
YENI_AD = {"114172": "Toplam Mevduat", "109865": "Özkaynak kalemleri", "159909": "İndirimler Öncesi Çekirdek Sermaye", "159953": "Ana Sermaye Toplamı (Çekirdek + İlave Ana Sermaye)",
           "159977": "Yasal Özkaynak (Ana Sermaye + Katkı Sermaye)", "159978": "Toplam Risk Ağırlıklı Tutarlar"}
# TBB kayıtlarından grubu doğru okunamayan bankalar için elle düzeltme (banka kodu -> TBB grubu)
GRUP_DUZELTME = {26: "Özel Sermayeli Mevduat Bankaları",   # ING Bank
                 92: "Özel Sermayeli Mevduat Bankaları",   # Fibabanka
                 29: "Özel Sermayeli Mevduat Bankaları"}   # Türk Ticaret Bankası
ILK7 = [3, 6, 32, 5, 30, 8, 37]  # Ziraat, VakıfBank, İş, Halkbank, Garanti BBVA, Akbank, Yapı Kredi
KISA = {3: "Ziraat", 6: "VakıfBank", 32: "İş Bankası", 5: "Halkbank", 30: "Garanti BBVA", 8: "Akbank", 37: "Yapı Kredi"}
AY = {"03": ("Mar", "Mart", 3), "06": ("Haz", "Haziran", 6), "09": ("Eyl", "Eylül", 9), "12": ("Ara", "Aralık", 12)}


def yil_den(y):
    """2018 -> 2018’den, 2019 -> 2019’dan (Türkçe ayrılma hâli eki)."""
    son = int(y) % 10
    onlar = int(y) % 100
    if son == 0:
        ek = {0: "dan", 10: "dan", 20: "den", 30: "dan", 40: "tan", 50: "den", 60: "tan", 70: "ten", 80: "den", 90: "dan"}[onlar]
    else:
        ek = {1: "den", 2: "den", 3: "ten", 4: "ten", 5: "ten", 6: "dan", 7: "den", 8: "den", 9: "dan"}[son]
    return f"{y}’{ek}"


def yil_ya(y):
    """2026 -> 2026’ya, 2025 -> 2025’e (Türkçe yönelme hâli eki)."""
    son, onlar = int(y) % 10, int(y) % 100
    if son == 0:
        ek = {0: "a", 10: "a", 20: "ye", 30: "a", 40: "a", 50: "ye", 60: "a", 70: "e", 80: "e", 90: "a"}[onlar]
    else:
        ek = {1: "e", 2: "ye", 3: "e", 4: "e", 5: "e", 6: "ya", 7: "ye", 8: "e", 9: "a"}[son]
    return f"{y}’{ek}"


def kisa_ad(n):
    for x in ["Türkiye Cumhuriyeti ", " T.A.Ş.", " A.Ş.", " T.A.O.", " (SA)", " S.p.A.", " N.A."]:
        n = n.replace(x, "")
    return n.strip()


def seviye(ad):
    ad = re.sub(r"^(\d+(?:\.\d+)*) (\d+)\.", r"\1.\2.", ad.strip())  # "16.5 1.Yasal" -> "16.5.1."
    if re.match(r"^[IVXLC]+\.", ad):
        return 0
    m = re.match(r"^([0-9]+(?:\.[0-9]+)*)\.?(?![0-9])", ad)
    return m.group(1).count(".") if m else 0


def oku(yol):
    banks, sec, cur, baslik = None, {}, None, ""
    with open(yol, encoding="utf-8-sig") as f:
        for i, l in enumerate(f):
            l = l.rstrip("\n")
            if i == 0:
                baslik = l
                continue
            if l.startswith("#BANKS "):
                banks = json.loads(l[7:])
            elif l.startswith("#ITEMS "):
                _, tur, js = l.split(" ", 2)
                sec[tur] = {"items": json.loads(js), "val": {}}
            elif l.startswith("#DATA "):
                cur = l.split(" ")[1]
            elif l and cur:
                c, b, o, tp, yp = l.split(",")
                sec[cur]["val"][(c, int(b), f"{o[:4]}-{o[4:]}")] = (float(tp), float(yp))
    m = re.search(r"alındı:\s*([0-9.]+)", baslik)
    return banks, sec, (m.group(1) if m else "")


def olcek_tespit(val, kod, donemler):
    """TBB bazı dönemlerde oranları 1.000 ile çarpılmış yayımlıyor; medyanı 1.000'i aşan dönemleri bulur."""
    out = set()
    for p in donemler:
        xs = [sum(v) for (c, b, q), v in val.items() if c == kod and q == p]
        if xs and statistics.median(xs) > 1000:
            out.add(p)
    return out


def veri_kur(tur, sec, banks, solo_kodlar):
    val, its = sec[tur]["val"], sec[tur]["items"]
    P = sorted({k[2] for k in val})
    liq = LIKIDITE[tur]
    off = 250000 if tur == "kons" else 0

    def eslestir(c):  # konsolide kodları solo karşılıklarına çevrilir
        if tur == "solo":
            return c
        if c == "17078":
            return "17077"
        q = str(int(c) - off)
        return q if q in solo_kodlar else c

    oranlar = {(str(int(k) + off)): v for k, v in SERMAYE_ORAN.items()}
    kok_eslestir = {str(int(k) + 250000): k for k in list(TABLO_ADI) + list(DIPNOT_ADI)} if tur == "kons" else {}
    kok_eslestir.update({"266039": "16039"})
    ust_of = {c: (str(u) if u is not None else None) for c, ad, u, sk in its}
    ad_of = {c: ad.strip() for c, ad, u, sk in its}
    kalemler = []
    dipnot = []
    for c, ad, ust, sk in its:
        if c in liq or c in oranlar or ust is None:
            continue
        # dipnot kalemleri: üst zincirini kalem listesinin dışına çıkana kadar izle
        d, a = 0, str(ust)
        onceki = ad_of.get(c, "").startswith("Önceki Dönem")
        while kok_eslestir.get(a, a) not in DIPNOT_ADI and a in ust_of and ust_of[a] is not None:
            onceki = onceki or ad_of.get(a, "").startswith("Önceki Dönem")
            d += 1; a = ust_of[a]
        if onceki:      # dipnotların önceki dönem dalları (her dönemin kendi verisi zaten var)
            continue
        a = kok_eslestir.get(a, a)
        if a in DIPNOT_ADI:
            skey = tuple(int(x) for x in re.findall(r"\d+", sk or "")) or (10**9,)
            mc = eslestir(c)
            if tur == "kons" and mc == c and int(c) < off:
                mc = str(int(c) + 9000000)   # yalnız konsolide tabloda olan kalem
            dipnot.append([skey, mc, c, re.sub(r"^\d+\.\s*", "", ad.strip()), DIPNOT_ADI[a], d, str(ust)])
            continue
        kok = str(ust)
        kok = kok_eslestir.get(kok, kok)
        if kok not in TABLO_ADI:
            continue
        skey = tuple(int(x) for x in re.findall(r"\d+", sk or "")) or (10**9,)
        kalemler.append([skey, eslestir(c), c, ad.strip(), TABLO_ADI[kok]])
    kalemler.sort(key=lambda x: (x[4], x[0]))
    items, onceki = [], {}
    for sk, mc, c, ad, tab in kalemler:
        lv = seviye(ad)
        lv = min(lv, onceki.get(tab, -1) + 1)
        onceki[tab] = lv
        items.append([int(mc), ad, lv, tab])
    # dipnot kalemleri ağaç sırasıyla (derinlik öncelikli) eklenir
    by_parent = {}
    for row in dipnot:
        by_parent.setdefault(row[6], []).append(row)
    dip_codes = {row[2] for row in dipnot}
    def dfs(par):
        for row in sorted(by_parent.get(par, []), key=lambda r: r[0]):
            items.append([int(row[1]), YENI_AD.get(row[1], row[3]), row[5], row[4]])
            dfs(row[2])
    for par in sorted({r[6] for r in dipnot if r[6] not in dip_codes}, key=lambda p: by_parent[p][0][4]):
        dfs(par)
    kaynak = {str(mc): c for sk, mc, c, ad, tab in kalemler}
    kaynak.update({str(r[1]): r[2] for r in dipnot})
    # mevduat yapısı: türler ve vade dilimleri (yaprak kalemler)
    adi = {str(it[0]): it[1].split(" (")[0].strip() for it in items}
    ebeveyn = {str(r[1]): (str(int(r[6]) - off) if tur == "kons" and str(int(r[6]) - off) in adi else r[6]) for r in dipnot}
    def alti(kod):
        k = kod
        while k in ebeveyn:
            k = ebeveyn[k]
            if k == kod: break
            yield k
    mv_vade = {v: [] for v in VADE}
    for kod, nm in adi.items():
        if nm in mv_vade and any(a == "114172" for a in alti(kod)):
            mv_vade[nm].append(int(kod))
    MV = {"top": 114172, "tur": [[ad, int(k)] for k, ad in MEVDUAT_TUR if k in adi],
          "vade": [[VADE_ETIKET.get(v, v), mv_vade[v]] for v in VADE]}

    meta = {b[0]: b for b in banks}
    akod = "16018" if tur == "solo" else "266018"
    bks = sorted({k[1] for k in val}, key=lambda b: -sum(val.get((akod, b, P[-1]), (0, 0))))
    A = {"codes": bks,
         "names": [meta[b][1] if b in meta else str(b) for b in bks],
         "short": [KISA.get(b) or kisa_ad(meta[b][1] if b in meta else str(b)) for b in bks],
         "grp": [GRUP_DUZELTME.get(b) or (meta[b][2] if b in meta and len(meta[b]) > 2 else "Grup bilgisi yok") for b in bks],
         "v": {}, "yp": {}, **{k: {} for k in liq.values()}, **{k: {} for k in SERMAYE_ORAN.values()}}
    num = lambda x: int(round(x)) if abs(x) >= 100 or float(x).is_integer() else round(x, 2)
    olcekli = olcek_tespit(val, [k for k, v in liq.items() if v == "lcr"][0], P)
    olcekli_syo = olcek_tespit(val, [k for k, v in oranlar.items() if v == "car"][0], P)
    for p in P:
        A["v"][p] = {}
        for it in items:
            mc = str(it[0]); c = kaynak[mc]
            arr = [None if (c, b, p) not in val else num(sum(val[(c, b, p)])) for b in bks]
            while arr and arr[-1] is None:
                arr.pop()
            if any(x is not None for x in arr):
                A["v"][p][mc] = arr
        A["yp"][p] = [int(round(val.get((akod, b, p), (0, 0))[1])) for b in bks]
        for c, key in liq.items():
            sc = 1000 if (p in olcekli and key in ORANLAR) else 1
            A[key][p] = [None if (c, b, p) not in val else
                         (round(sum(val[(c, b, p)]) / sc, 2) if key in ORANLAR else int(round(sum(val[(c, b, p)]))))
                         for b in bks]
        for c, key in oranlar.items():
            sc = 1000 if p in olcekli_syo else 1
            arr = []
            for b in bks:
                if (c, b, p) not in val:
                    arr.append(None); continue
                x = sum(val[(c, b, p)]) / sc
                if key in PAY_KOD:      # oranı tutarlardan doğrula; bankanın açıkladığı birim hatalıysa hesaplanan kullanılır
                    pay, rwa = val.get((str(int(PAY_KOD[key]) + off), b, p)), val.get((str(159978 + off), b, p))
                    if pay and rwa and sum(rwa) > 0:
                        h = sum(pay) / sum(rwa) * 100
                        if abs(x - h) > max(0.5, 0.05 * abs(h)):
                            x = h
                elif key in TAMPON:     # tampon oranları birkaç yüzdeyi geçemez; 10'un katlarıyla yayımlananlar düzeltilir
                    while x > TAMPON[key]:
                        x /= 10
                arr.append(round(x, 2))
            A[key][p] = arr
    T7 = [bks.index(c) for c in ILK7 if c in bks]
    if len(T7) != 7:
        sys.exit(f"[{tur}] ilk 7 bankanın hepsi veride yok")
    plabel = {p: f"{AY[p[5:]][0]} {p[:4]}" for p in P}
    months = {p: AY[p[5:]][2] for p in P}
    return {"periods": P, "plabel": plabel, "months": months, "items": items, "A": A, "T7": T7, "MV": MV}, olcekli | olcekli_syo


MALI_BUNYE = "102772"          # 1. Mali bünye ile ilgili dipnotlar (solo kodu; konsolidede + 250.000)
# Sayfalarda gösterilen dipnot grupları (kök kodu, grup adı); yenisi eklenince tbb_cek.py'de de çekilmeli
DIPNOT_KOKLERI = [("102772", "Mali bünye ile ilgili dipnotlar"), ("102774", "Aktiflerle ilgili dipnotlar"),
                  ("102776", "Pasiflerle ilgili dipnotlar")]
DN_HARIC = {"109869"}          # özkaynak kalemleri raporda zaten "9. Yasal Özkaynak" tablosu olarak var


def _n(s):
    """Türkçe küçük harf, sade boşluk (ad eşleştirmek için)."""
    return re.sub(r"\s+", " ", s.replace("İ", "i").replace("I", "ı").lower()).strip()


def dipnot_kur(tur, sec, bks):
    """Mali bünye dipnotlarını ayrı veri paketleri olarak hazırlar (sayfada istenince yüklenir).
    Paket: {"items": [[kod, ad, düzey, oran_mı, (pay, payda)?]], "v": {dönem: {kod: [banka_sırası, değer, ...]}}}
    - "Önceki Dönem" dalları alınmaz (her dönemin kendi verisi zaten var); "Cari Dönem-" öneki atılır.
    - Bankaların dipnotlarında yer alan toplam ve net satırları (TBB'de sütun bazında yok) hesaplanarak eklenir;
      bu satırların kodu 9.500.000'den başlar.
    - Kaldıraç, LKO ve NİFO oranlarına, grup toplamlarında kullanılacak pay/payda kalemleri eklenir."""
    val, its = sec[tur]["val"], sec[tur]["items"]
    off = 250000 if tur == "kons" else 0
    ad = {c: re.sub(r"\s*\(\d+\)\s*$", "", n).strip() for c, n, u, sk in its}
    sk_of = {c: tuple(int(x) for x in re.findall(r"\d+", sk or "")) or (10**9,) for c, n, u, sk in its}
    cocuk = {}
    for c, n, u, sk in its:
        if u is not None:
            cocuk.setdefault(str(u), []).append(c)
    for k in cocuk:
        cocuk[k].sort(key=lambda c: sk_of[c])
    bidx = {b: i for i, b in enumerate(bks)}
    V = {}   # kod -> dönem -> {banka_sırası: değer}
    for (c, b, p), tv in val.items():
        if b in bidx:
            V.setdefault(c, {}).setdefault(p, {})[bidx[b]] = sum(tv)
    num = lambda x: int(round(x)) if abs(x) >= 100 else round(x, 2)
    manifest, paketler = [], {}
    for r, grup in [(r_, g_) for k_, g_ in DIPNOT_KOKLERI for r_ in cocuk.get(str(int(k_) + off), [])]:
        if str(int(r) - off) in DN_HARIC or r in DN_HARIC:
            continue
        # ağaç (Önceki Dönem dalları hariç)
        N = {}
        def kur(c):
            n = ad.get(c, c)
            N[c] = {"c": c, "n": re.sub(r"^Cari Dönem\s*-\s*", "", n).strip(), "k": [k for k in cocuk.get(c, []) if not _n(ad.get(k, "")).startswith("önceki dönem")]}
            for k in N[c]["k"]:
                kur(k)
        kur(r)
        sayac = [9500000]
        def yeni(n, kids=None):
            sayac[0] += 1; c = str(sayac[0]); N[c] = {"c": c, "n": n, "k": kids or [], "h": 1}; return c
        bul = lambda ust, kosul: next((k for k in N[ust]["k"] if kosul(_n(N[k]["n"]))), None)
        def ara(kosul):
            return [c for c in N if kosul(_n(N[c]["n"]))]
        yaprak = lambda c: not N[c]["k"]
        matris = lambda c: N[c]["k"] and all(yaprak(k) for k in N[c]["k"])
        def sutunlar(satir):
            return {_n(N[k]["n"]): k for k in N[satir]["k"]}
        def topla(satirlar, isim, isaretler=None):
            """satırların sütun bazında (işaretli) toplamı olan yeni bir matris satırı"""
            isaretler = isaretler or [1] * len(satirlar)
            sira, etiket = [], {}
            for s_ in satirlar:
                for k in N[s_]["k"]:
                    key = _n(N[k]["n"])
                    if key not in etiket:
                        etiket[key] = N[k]["n"]; sira.append(key)
            yeni_satir = yeni(isim)
            for key in sira:
                yc = yeni(etiket[key])
                N[yeni_satir]["k"].append(yc)
                for p in {p for s_ in satirlar for k in N[s_]["k"] for p in V.get(k, {})}:
                    acc = {}
                    for s_, sg in zip(satirlar, isaretler):
                        k = sutunlar(s_).get(key)
                        for i, x in V.get(k, {}).get(p, {}).items():
                            # "fark": TBB yükümlülükleri ve borçları eksi işaretli tutuyor; işaretten bağımsız A − |B|
                            acc[i] = acc.get(i, 0) + (-abs(x) if sg == "fark" else sg * x)
                    if acc:
                        V.setdefault(yc, {})[p] = acc
                        tot = V.setdefault(yeni_satir, {}).setdefault(p, {})
                        for i, x in acc.items():
                            tot[i] = tot.get(i, 0) + x
            return yeni_satir
        def toplam_ekle(kap, isim):
            satirlar = [k for k in N[kap]["k"] if matris(k)]
            if not satirlar:
                return None
            t = topla(satirlar, isim); N[kap]["k"].append(t); return t
        rn = _n(ad.get(r, ""))
        if rn.startswith("kur riskine"):
            for nb in ara(lambda n: n == "net bilanço pozisyonu"):
                va, yu = bul(nb, lambda n: n == "varlıklar"), bul(nb, lambda n: n == "yükümlülükler")
                tv = toplam_ekle(va, "Toplam Varlıklar") if va else None
                ty = toplam_ekle(yu, "Toplam Yükümlülükler") if yu else None
                if tv and ty:
                    N[nb]["k"].append(topla([tv, ty], "Net Bilanço Pozisyonu (varlıklar − yükümlülükler)", [1, "fark"]))
            for nn in ara(lambda n: n == "net nazım hesap pozisyonu"):
                al, bo = bul(nn, lambda n: "türev" in n and "alacak" in n), bul(nn, lambda n: "türev" in n and "borç" in n)
                if al and bo:
                    N[nn]["k"].append(topla([al, bo], "Net Nazım Hesap Pozisyonu (alacaklar − borçlar)", [1, "fark"]))
        elif rn == "faiz oranı riski":
            for vy in ara(lambda n: "faiz riski-varlık-yüküm" in n):
                for isim, kos in (("Toplam Varlıklar", "varlıklar"), ("Toplam Yükümlülükler", "yükümlülükler")):
                    k = bul(vy, lambda n, kos=kos: n == kos)
                    if k:
                        toplam_ekle(k, isim)
            for tp in ara(lambda n: n.endswith("toplam pozisyon") and "faiz" in n):
                satirlar = [k for k in N[tp]["k"] if matris(k)]
                if satirlar:
                    N[tp]["k"].append(topla(satirlar, "Toplam Pozisyon"))
        elif "kalan vadelerine" in rn:
            for la in ara(lambda n: n == "likidite açığı"):
                va, yu = bul(la, lambda n: n == "varlıklar"), bul(la, lambda n: n == "yükümlülükler")
                tv = toplam_ekle(va, "Toplam Varlıklar") if va else None
                ty = toplam_ekle(yu, "Toplam Yükümlülükler") if yu else None
                if tv and ty:
                    N[la]["k"].append(topla([tv, ty], "Net Likidite Açığı (varlıklar − yükümlülükler)", [1, "fark"]))
            for bd in ara(lambda n: n == "net bilanço dışı pozisyonu"):
                al, bo = bul(bd, lambda n: "türev" in n and "alacak" in n), bul(bd, lambda n: "türev" in n and "borç" in n)
                if al and bo:
                    N[bd]["k"].append(topla([al, bo], "Net Bilanço Dışı Pozisyonu (alacaklar − borçlar)", [1, "fark"]))
        # oran kalemlerinin grup toplamı için pay / payda
        formul = {}
        for c in list(N):
            n = _n(N[c]["n"])
            par = next((q for q in N if c in N[q]["k"]), None)
            if n.startswith("kaldıraç oranı") and par:
                a_, t_ = bul(par, lambda m: m == "ana sermaye"), bul(par, lambda m: m == "toplam risk tutarı")
                if a_ and t_:
                    formul[c] = (a_, t_)
            if "net istikrarlı fonlama oranı (%)" in n and par:
                m_, g_ = bul(par, lambda m: "mevcut istikrarlı fon toplamı" in m), bul(par, lambda m: "gerekli istikrarlı fon toplamı" in m)
                if m_ and g_:
                    formul[c] = (m_, g_)
            if n.startswith("likidite karşılama oranı (%)") and par:
                h_, o_ = bul(par, lambda m: m.startswith("toplam yüksek kaliteli likit varlık")), bul(par, lambda m: m.startswith("toplam net nakit çıkış"))
                if h_ and o_:
                    for k in N[c]["k"]:
                        key = _n(N[k]["n"])
                        hk, ok = sutunlar(h_).get(key), sutunlar(o_).get(key)
                        if hk and ok:
                            formul[k] = (hk, ok)
        # düzleştir
        sira = []
        def dfs(c, d, ust_ad):
            sira.append((c, N[c]["n"], d, ust_ad))
            for k in N[c]["k"]:
                dfs(k, d + 1, ust_ad + " " + N[c]["n"])
        dfs(r, 0, "")
        items, v = [], {}
        for c, n, d, ust_ad in sira:
            # dipnot başlığının TBB'deki değeri alt satırların (toplamlar dahil) toplamı; anlamsız olduğu için boş bırakılır
            seri = {} if c == r else V.get(c, {})
            tum = [abs(x) for pv in seri.values() for x in pv.values() if x]
            med = statistics.median(tum) if tum else 0
            low = _n(n)
            oran = ("(%)" in n) or ("(%)" in ust_ad) or (c in formul) or ("oran" in low and "tutar" not in low and tum and med < 500)
            olc = set()   # oran kalemlerinde 1.000 ile çarpılmış yayımlanan dönemler
            if oran and len(seri) > 2:
                meds = {p: statistics.median([abs(x) for x in pv.values() if x] or [0]) for p, pv in seri.items()}
                for p, m in meds.items():
                    digerleri = [mm for q, mm in meds.items() if q != p and mm]
                    if m and digerleri and m > 300 * statistics.median(digerleri):
                        olc.add(p)
            it = [int(c), n, d, 1 if oran else 0]
            if c in formul:
                it.append([int(formul[c][0]), int(formul[c][1])])
            items.append(it)
            for p, pv in seri.items():
                sc = 1000 if p in olc else 1
                flat = []
                for i, x in sorted(pv.items()):
                    flat += [i, num(x / sc)]
                if flat:
                    v.setdefault(p, {})[c] = flat
        if len(items) <= 1 and not v:
            continue
        key = f"{tur}_{r}"
        paketler[key] = {"items": items, "v": v}
        manifest.append({"id": key, "ad": re.sub(r"\s+", " ", ad.get(r, r)).strip(), "n": len(items), "g": grup})
    return manifest, paketler


def dipnot_yaz(paketler, klasor):
    klasor = Path(klasor)
    klasor.mkdir(parents=True, exist_ok=True)
    for f in klasor.glob("*.js"):
        f.unlink()
    toplam = 0
    for key, pk in paketler.items():
        t = f'__DN("{key}",' + json.dumps(pk, ensure_ascii=False, separators=(",", ":")) + ");\n"
        (klasor / f"{key}.js").write_text(t, encoding="utf-8")
        toplam += len(t.encode("utf-8"))
    return toplam


def html_yaz(D, tarih, olcekli, cikti):
    t = (HERE / "rapor_sablonu.html").read_text(encoding="utf-8")
    P = D["solo"]["periods"] if "solo" in D else D["kons"]["periods"]
    p0, p1 = P[0], P[-1]
    a0, a1 = AY[p0[5:]][1], AY[p1[5:]][1]
    olc = sorted(olcekli)
    oran_notu = ""
    if olc:
        ad = ", ".join(f"{AY[p[5:]][1]} {p[:4]}" for p in olc)
        oran_notu = f"{ad} verisinde TBB oranları 1.000 ile çarpılmış yayımladığı için bu dönem{'lerin' if len(olc) > 1 else 'in'} oranları 1.000’e bölündü. "
    rep = {
        "{{ARALIK}}": f"{a0} {p0[:4]} – {a1} {p1[:4]}",
        "{{ARALIK_DEN_YA}}": f"{a0} {yil_den(p0[:4])} {a1} {yil_ya(p1[:4])}",
        "{{N}}": str(len(P)),
        "{{NBANKA}}": str(len(D["solo"]["A"]["codes"]) if "solo" in D else len(D["kons"]["A"]["codes"])),
        "{{ILK_YIL_DEN}}": yil_den(p0[:4]),
        "{{YILLAR}}": f"{p0[:4]}–{p1[:4]}",
        "{{ORAN_NOTU}}": oran_notu,
        "{{TARIH}}": tarih or "–",
    }
    for k, v in rep.items():
        t = t.replace(k, v)
    if "kons" not in D or "solo" not in D:
        sys.exit("HTML rapor için hem solo hem konsolide veri gerekir (tbb_cek.py --tur solo,kons).")
    t = t.replace("__DATA__", json.dumps(D, ensure_ascii=False, separators=(",", ":")))
    head = t.split("<style>")[0]
    rest = t[len(head):]
    css = rest.split("</style>")[0] + "</style>"
    body = rest.split("</style>", 1)[1]
    html = ('<!doctype html>\n<html lang="tr">\n<head>\n<meta charset="utf-8">\n'
            '<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n'
            + head + css + "\n</head>\n<body>\n" + body + "\n</body>\n</html>\n")
    Path(cikti).parent.mkdir(parents=True, exist_ok=True)
    Path(cikti).write_text(html, encoding="utf-8")


def uzun_yaz(sec, banks, cikti):
    meta = {b[0]: b for b in banks}
    with open(cikti, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["tur", "kalem_kodu", "kalem", "banka_kodu", "banka", "banka_grubu", "donem", "TP", "YP", "TOPLAM"])
        for tur, s in sec.items():
            ad = {c: n for c, n, *_ in s["items"]}
            for (c, b, p), (tp, yp) in sorted(s["val"].items(), key=lambda x: (x[0][2], x[0][1], x[0][0])):
                m = meta.get(b, [b, "", ""])
                w.writerow([tur, c, ad.get(c, ""), b, m[1], m[2] if len(m) > 2 else "", p, tp, yp, tp + yp])


def excel_yaz(sec, banks, cikti):
    try:
        import pandas as pd
    except ImportError:
        sys.exit("Excel çıktısı için:  pip install pandas openpyxl")
    meta = {b[0]: b for b in banks}
    with pd.ExcelWriter(cikti) as xw:
        for tur, s in sec.items():
            ad = {c: n for c, n, *_ in s["items"]}
            rows = [(c, ad.get(c, ""), meta.get(b, [b, ""])[1], p, tp + yp) for (c, b, p), (tp, yp) in s["val"].items()]
            df = pd.DataFrame(rows, columns=["kod", "kalem", "banka", "donem", "toplam"])
            pv = df.pivot_table(index=["kod", "kalem", "banka"], columns="donem", values="toplam", aggfunc="sum")
            pv.to_excel(xw, sheet_name=("Solo" if tur == "solo" else "Konsolide"))


def main():
    ap = argparse.ArgumentParser(description="tbb_veri.csv dosyasından HTML rapor ve tablolar üretir.")
    ap.add_argument("--girdi", default="tbb_veri.csv")
    ap.add_argument("--html", default="Turk_Bankacilik_Sistemi.html", help="HTML rapor dosyası ('' verilirse üretilmez)")
    ap.add_argument("--json", default="", help="Rapor verisini ayrıca JSON olarak yaz")
    ap.add_argument("--uzun", default="", help="Düz (uzun) CSV: her satır bir kalem-banka-dönem")
    ap.add_argument("--excel", default="", help="Excel: her tür için kalem×banka satır, dönem sütun")
    a = ap.parse_args()

    banks, sec, tarih = oku(a.girdi)
    ozet = ", ".join("{} {:,} değer".format(k, len(v["val"])) for k, v in sec.items())
    print("Okundu: " + ozet)
    solo_kodlar = {c for c, *_ in sec["solo"]["items"]} if "solo" in sec else set()
    D, olc, paketler = {}, set(), {}
    for tur in sec:
        D[tur], o = veri_kur(tur, sec, banks, solo_kodlar)
        olc |= o
        D[tur]["DN"], pk = dipnot_kur(tur, sec, D[tur]["A"]["codes"])
        paketler.update(pk)
        A = D[tur]["A"]
        print(f"[{tur}] {len(D[tur]['periods'])} dönem, {len(A['codes'])} banka, {len(D[tur]['items'])} kalem"
              + (f"; ölçeği düzeltilen oran dönemleri: {sorted(o)}" if o else ""))
    if a.json:
        Path(a.json).write_text(json.dumps(D, ensure_ascii=False), encoding="utf-8")
        print("Yazıldı:", a.json)
    if a.html:
        html_yaz(D, tarih, olc, a.html)
        print("Yazıldı:", a.html)
        if paketler:
            kl = Path(a.html).resolve().parent / "dipnot"
            mb = dipnot_yaz(paketler, kl) / 1e6
            print(f"Yazıldı: {kl} ({len(paketler)} dipnot paketi, {mb:.1f} MB)")
    if a.uzun:
        uzun_yaz(sec, banks, a.uzun)
        print("Yazıldı:", a.uzun)
    if a.excel:
        excel_yaz(sec, banks, a.excel)
        print("Yazıldı:", a.excel)


if __name__ == "__main__":
    main()
