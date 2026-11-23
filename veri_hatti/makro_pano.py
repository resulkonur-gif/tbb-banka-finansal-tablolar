#!/usr/bin/env python3
"""Siteye "Makro pano" bölümü ekler: TCMB rezervleri (brüt / net / swap hariç net), rezerv yeterliliği,
reel faiz, enflasyon ve beklentiler, cari denge ve finansmanı, dolarizasyon, yabancı portföy akımları,
reel efektif kur, kredi–mevduat faiz farkı. Her gösterge trafik ışığı ve kısa yorumla.

Kullanım:
    python makro_pano.py --site docs/onizleme                       # makro_veri.json'daki son veriyle
    python makro_pano.py --site docs/onizleme --anahtar <EVDS_API>  # önce EVDS'den günceller
    (anahtar EVDS_API_KEY ortam değişkeninden de okunur; evds2.tcmb.gov.tr → Profil → API Anahtarı)

Veri makro_veri.json dosyasında tutulur (betiğin yanında). EVDS'ye ulaşılamazsa eldeki veriyle devam eder.
Betik idempotenttir: daha önce eklenen bölüm işaretlerinden bulunup değiştirilir.
"""
import argparse
import datetime as dt
import json
import os
import re
import urllib.parse
import urllib.request
from pathlib import Path

BURADA = Path(__file__).resolve().parent
VERI = Path(os.environ.get("MAKRO_VERI") or BURADA / "makro_veri.json")   # veri_hatti.py 5_makro/ klasörünü verir

# (kod, frekans, toplulaştırma, formül, anahtar) — anahtar makro_veri.json'daki ad
SERILER = [
    ("TP.AB.TOPLAM", 5, "last", None, "TP.AB.TOPLAM"),                      # brüt rezerv (milyon $), haftalık
    ("TP.AB.A02", 5, "last", None, "TP.AB.A02"),                            # analitik bilanço dış varlıklar (bin TL)
    ("TP.AB.A10", 5, "last", None, "TP.AB.A10"),                            # toplam döviz yükümlülükleri (bin TL)
    ("TP.DK.USD.A.YTL", 5, "last", None, "TP.DK.USD.A.YTL"),                # USD/TRY alış
    ("TP.DK.EUR.A.YTL", 5, "last", None, "TP.DK.EUR.A.YTL"),                # EUR/TRY alış (sepet kur için)
    ("TP.DOVVARNC.K15", None, None, None, "TP.DOVVARNC.K15"),               # IRFCL 2.2.1 forward/swap açık pozisyon
    ("TP.BISPOLFAIZ.TUR", None, None, None, "TP.BISPOLFAIZ.TUR"),           # politika faizi (BIS)
    ("TP.TUKFIY2025.GENEL", None, None, 3, "TP.TUKFIY2025.GENEL#yillik"),   # TÜFE yıllık %
    ("TP.PKAUO.S01.E.U", None, None, None, "TP.PKAUO.S01.E.U"),             # 12 ay sonrası TÜFE beklentisi
    ("TP.PKAUO.S01.D.U", None, None, None, "TP.PKAUO.S01.D.U"),             # cari yıl sonu TÜFE beklentisi
    ("TP.PKAUO.S01.I.U", None, None, None, "TP.PKAUO.S01.I.U"),             # gelecek yıl sonu TÜFE beklentisi
    ("TP.PKAUO.S04.D.U", None, None, None, "TP.PKAUO.S04.D.U"),             # 12 ay sonrası politika faizi beklentisi
    ("TP.PKAUO.S05.C.U", None, None, None, "TP.PKAUO.S05.C.U"),
    ("TP.PKAUO.S05.B.U", None, None, None, "TP.PKAUO.S05.B.U"),             # cari yıl sonu USD/TRY beklentisi
    ("TP.PKAUO.S04.H.U", None, None, None, "TP.PKAUO.S04.H.U"),             # cari yıl sonu politika faizi beklentisi
    ("TP.PKAUO.S04.I.U", None, None, None, "TP.PKAUO.S04.I.U"),             # gelecek yıl sonu politika faizi beklentisi
    ("TP.PKAUO.S04.E.U", None, None, None, "TP.PKAUO.S04.E.U"),             # 24 ay sonrası politika faizi beklentisi
    ("TP.PKAUO.S01.F.U", None, None, None, "TP.PKAUO.S01.F.U"),             # 24 ay sonrası TÜFE beklentisi             # 12 ay sonrası USD/TRY beklentisi
    ("TP.PKAUO.S06.A.U", None, None, None, "TP.PKAUO.S06.A.U"),             # cari yıl cari denge beklentisi
    ("TP.PKAUO.S06.B.U", None, None, None, "TP.PKAUO.S06.B.U"),             # gelecek yıl cari denge beklentisi
    ("TP.PKAUO.S07.A.U", None, None, None, "TP.PKAUO.S07.A.U"),             # cari yıl büyüme beklentisi
    ("TP.PKAUO.S07.B.U", None, None, None, "TP.PKAUO.S07.B.U"),             # gelecek yıl büyüme beklentisi
    ("TP.ODANA6.Q01", None, None, None, "TP.ODANA6.Q01"),                   # cari işlemler dengesi (aylık, milyon $)
    ("TP.ODEAYRSUNUM6.Q210", None, None, None, "TP.ODEAYRSUNUM6.Q210"),     # net hata ve noksan
    ("TP.KALANVADE.K20", None, None, None, "TP.KALANVADE.K20"),             # kalan vadeye göre KV dış borç
    ("TP.TLDTHVADE.KB12", 5, "last", None, "TP.TLDTHVADE.KB12"),            # DTH toplam (bin TL)
    ("TP.TLDTHVADE.KB18", 5, "last", None, "TP.TLDTHVADE.KB18"),            # toplam mevduat (bin TL)
    ("TP.MKNETHAR.M7", 5, "sum", None, "TP.MKNETHAR.M7#ay"),                # yabancı hisse net alım (milyon $)
    ("TP.MKNETHAR.M8", 5, "sum", None, "TP.MKNETHAR.M8#ay"),                # yabancı DİBS net alım (milyon $)
    ("TP.KTF17", 5, "avg", None, "TP.KTF17"),                               # TL ticari kredi faizi (akım)
    ("TP.TRY.MT06", 5, "avg", None, "TP.TRY.MT06"),                         # TL mevduat faizi (akım)
    ("TP.RK.T1.Y", None, None, None, "TP.RK.T1.Y"),                         # TÜFE bazlı reel efektif kur (2025=100)
]
EVDS_ADRES = ["https://evds3.tcmb.gov.tr/igmevdsms-dis/", "https://evds2.tcmb.gov.tr/service/evds/"]


# ----------------------------------------------------------------------------- EVDS
def _ay(t):
    t = str(t)
    m = re.match(r"^(\d{4})-(\d{1,2})", t)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}"
    m = re.match(r"^(\d{1,2})-(\d{1,2})-(\d{4})", t)
    if m:
        return f"{m.group(3)}-{int(m.group(2)):02d}"
    return None


def evds_seri(anahtar, kod, frek, topl, formul, bas, son):
    q = {"series": kod, "startDate": bas, "endDate": son, "type": "json"}
    if frek:
        q["frequency"] = str(frek)
    if topl:
        q["aggregationTypes"] = topl
    if formul is not None:
        q["formulas"] = str(formul)
    yol = "&".join(f"{k}={urllib.parse.quote(str(v), safe='.-_')}" for k, v in q.items())
    son_hata = None
    for kok in EVDS_ADRES:
        try:
            req = urllib.request.Request(kok + yol, headers={"key": anahtar, "User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=60) as r:
                j = json.loads(r.read().decode("utf-8"))
            out = {}
            sutun = kod.replace(".", "_")
            for it in j.get("items", []):
                a = _ay(it.get("Tarih"))
                deger = None
                for k, v in it.items():
                    if k.startswith(sutun) and v not in (None, "", "null"):
                        deger = float(v)
                        break
                if a and deger is not None:
                    out[a] = deger
            if out:
                return out
        except Exception as e:                                   # noqa: BLE001
            son_hata = e
    raise RuntimeError(f"{kod}: {son_hata}")


def guncelle(anahtar, gunluk=print, bastan=False):
    """Artımlı: serinin arşivde son ayından 6 ay öncesinden (formüllü/toplamlı serilerde 24 ay) bugüne çeker;
    arşivde olmayan seri 2016'dan çekilir. --bastan ile hepsi 2016'dan."""
    V = json.loads(VERI.read_text(encoding="utf-8")) if VERI.exists() else {"seriler": {}}
    son = dt.date.today().strftime("%d-%m-%Y")
    tamam = yeni_gozlem = 0
    for kod, frek, topl, formul, ad in SERILER:
        mevcut = V["seriler"].get(ad) or {}
        if mevcut and not bastan:
            y, m = map(int, max(mevcut).split("-"))
            geri = 24 if (formul is not None or topl) else 6
            m -= geri
            while m <= 0:
                y, m = y - 1, m + 12
            bas = f"01-{m:02d}-{y}"
        else:
            bas = "01-01-2016"                  # site 2018-03'ten başlar; 12 aylık toplamlar için 2016
        try:
            yeni = evds_seri(anahtar, kod, frek, topl, formul, bas, son)
            s_ = V["seriler"].setdefault(ad, {})
            yeni_gozlem += sum(1 for a, x in yeni.items() if s_.get(a) != x)
            s_.update(yeni)
            tamam += 1
        except Exception as e:                                   # noqa: BLE001
            gunluk(f"  ! EVDS {kod}: {e}")
    if tamam:
        V["alindi"] = dt.date.today().isoformat()
        VERI.write_text(json.dumps(V, ensure_ascii=False, indent=0), encoding="utf-8")
    gunluk(f"  EVDS aylık: {tamam}/{len(SERILER)} seri sorgulandı, {yeni_gozlem} yeni/değişen gözlem")
    return tamam


# ----------------------------------------------------------------------------- hesap
def hesapla(V, aylik=False):
    S = V["seriler"]
    g = lambda ad: S.get(ad, {})                                  # noqa: E731
    aylar = sorted({a for s in S.values() for a in s if a >= "2017-01"})

    def deger(ad, a):
        return g(ad).get(a)

    def son_ay(ad, kosul=None):
        for a in reversed(aylar):
            if deger(ad, a) is not None and (kosul is None or kosul(a)):
                return a
        return None

    def ay_kaydir(a, n):
        y, m = map(int, a.split("-"))
        m += n
        while m <= 0:
            y, m = y - 1, m + 12
        while m > 12:
            y, m = y + 1, m - 12
        return f"{y}-{m:02d}"

    def toplam12(ad, a):
        xs = [deger(ad, ay_kaydir(a, -i)) for i in range(12)]
        return None if any(x is None for x in xs) else sum(xs)

    seri = {k: [] for k in ("brut", "net", "swapsiz", "politika", "tufe", "bek12", "reel", "cari12", "nhn12",
                            "efr", "yeterlilik", "dolar", "hisse", "dibs", "reer", "kredi", "mevduat", "makas")}
    for a in aylar:
        usd = deger("TP.DK.USD.A.YTL", a)
        br = deger("TP.AB.TOPLAM", a)
        a02, a10 = deger("TP.AB.A02", a), deger("TP.AB.A10", a)
        net = (a02 - a10) / usd / 1000 if a02 and a10 and usd else None
        sw = deger("TP.DOVVARNC.K15", a)
        pol, tf, bk = deger("TP.BISPOLFAIZ.TUR", a), deger("TP.TUKFIY2025.GENEL#yillik", a), deger("TP.PKAUO.S01.E.U", a)
        c12, n12 = toplam12("TP.ODANA6.Q01", a), toplam12("TP.ODEAYRSUNUM6.Q210", a)
        kv = deger("TP.KALANVADE.K20", a)
        efr = kv + max(0.0, -c12) if kv is not None and c12 is not None else None
        dth, top = deger("TP.TLDTHVADE.KB12", a), deger("TP.TLDTHVADE.KB18", a)
        kr, mv = deger("TP.KTF17", a), deger("TP.TRY.MT06", a)
        r = lambda x, d=1: None if x is None else round(x, d)    # noqa: E731
        seri["brut"].append(r(br / 1000 if br else None))
        seri["net"].append(r(net / 1000 if net is not None else None))
        seri["swapsiz"].append(r((net + sw) / 1000 if net is not None and sw is not None else None))
        seri["politika"].append(pol)
        seri["tufe"].append(r(tf, 2))
        seri["bek12"].append(bk)
        seri["reel"].append(r(pol - bk, 2) if pol is not None and bk is not None else None)
        seri["cari12"].append(r(c12 / 1000 if c12 is not None else None))
        seri["nhn12"].append(r(n12 / 1000 if n12 is not None else None))
        seri["efr"].append(r(efr / 1000 if efr is not None else None))
        seri["yeterlilik"].append(r(br / efr, 2) if br and efr else None)
        seri["dolar"].append(r(dth / top * 100, 1) if dth and top else None)
        seri["hisse"].append(r(deger("TP.MKNETHAR.M7#ay", a) / 1000, 2) if deger("TP.MKNETHAR.M7#ay", a) is not None else None)
        seri["dibs"].append(r(deger("TP.MKNETHAR.M8#ay", a) / 1000, 2) if deger("TP.MKNETHAR.M8#ay", a) is not None else None)
        seri["reer"].append(deger("TP.RK.T1.Y", a))
        seri["kredi"].append(r(kr, 2))
        seri["mevduat"].append(r(mv, 2))
        seri["makas"].append(r(kr - mv, 2) if kr is not None and mv is not None else None)

    AY = {1: "Oca", 2: "Şub", 3: "Mar", 4: "Nis", 5: "May", 6: "Haz", 7: "Tem", 8: "Ağu", 9: "Eyl", 10: "Eki",
          11: "Kas", 12: "Ara"}
    etk = lambda a: f"{AY[int(a[5:])]} {a[:4]}" if a else "–"     # noqa: E731
    son_veri_ayi = aylar[-1] if aylar else None
    alindi_ayi = (V.get("alindi") or "")[:7]

    def donem_gostergeleri(kesim):
        """kesim (YYYY-AA) ve öncesindeki veriyle göstergeler: rapor dönemi neyse makro da o tarihteki hâli."""
        return _gostergeler(kesim, aylar, seri, g, son_ay, etk, alindi_ayi)

    ceyrekler = []
    if aylar:
        y, m, adim = 2018, (1 if aylik else 3), (1 if aylik else 3)
        while f"{y}-{m:02d}" <= son_veri_ayi:
            ceyrekler.append(f"{y}-{m:02d}")
            m += adim
            if m > 12:
                y, m = y + 1, m - 12
    donemler = {q: donem_gostergeleri(q) for q in ceyrekler}
    if son_veri_ayi and son_veri_ayi not in donemler:          # çeyrek içi en son veri (dönem dışı görünüm)
        donemler["son"] = donem_gostergeleri(son_veri_ayi)
    return {"alindi": V.get("alindi"), "aylar": aylar, "etiket": [etk(a) for a in aylar], "seri": seri,
            "donemler": donemler}


def _gostergeler(kesim, aylar, seri, g, son_ay, etk, alindi_ayi):
    def son(k, geri=0):
        xs = [(a, v) for a, v in zip(aylar, seri[k]) if v is not None and a <= kesim]
        return xs[-1 - geri] if len(xs) > geri else (None, None)

    def isik(v, sinirlar, ters=False):
        """sinirlar=(kırmızı/sarı, sarı/yeşil); ters: büyük kötü"""
        if v is None:
            return "gri"
        a, b = sinirlar
        if ters:
            return "kirmizi" if v > b else "sari" if v > a else "yesil"
        return "kirmizi" if v < a else "sari" if v < b else "yesil"

    f1 = lambda x: "–" if x is None else f"{x:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".")  # noqa: E731
    G = []

    a, v = son("swapsiz"); _, v3 = son("swapsiz", 3); ab, vb = son("brut"); an, vn = son("net")
    G.append(dict(k="swapsiz", ad="Swap hariç net rezerv", deger=f"{f1(v)} mlr $", ay=etk(a),
                  isik=isik(v, (20, 50)), deg=None if v3 is None else round(v - v3, 1), birim="mlr $",
                  not_=f"Brüt {f1(vb)} ({etk(ab)}), net {f1(vn)} mlr $. TCMB'nin taahhüt edilmemiş dövizi; "
                       f"20 mlr $ altı müdahale gücünün zayıf olduğu bölge."))
    a, v = son("yeterlilik"); _, v12 = son("yeterlilik", 12); ae, ve = son("efr")
    G.append(dict(k="yeterlilik", ad="Rezerv yeterliliği", deger=f"{f1(v)}x" if v else "–", ay=etk(a),
                  isik=isik(v, (0.8, 1.0)), deg=None if v12 is None else round(v - v12, 2), birim="x",
                  not_=f"Brüt rezerv / dış finansman ihtiyacı ({f1(ve)} mlr $ = kalan vadesi 1 yıla kadar dış borç + "
                       f"12 aylık cari açık). 1'in altı: rezerv, bir yıllık dış ödemeleri karşılamıyor."))
    a, v = son("reel"); _, v3 = son("reel", 3); ap, vp = son("politika"); ab_, vbk = son("bek12")
    G.append(dict(k="reel", ad="Reel politika faizi (beklentiye göre)", deger=f"%{f1(v)}", ay=etk(a),
                  isik=isik(v, (0, 5)), deg=None if v3 is None else round(v - v3, 1), birim="puan",
                  not_=f"Politika faizi %{f1(vp)} ({etk(ap)}) − 12 ay sonrası enflasyon beklentisi %{f1(vbk)}. "
                       f"Pozitif ve yüksek reel faiz TL taşıma işlemini ve dezenflasyonu destekler."))
    a, v = son("bek12"); _, v3 = son("bek12", 3); at, vt = son("tufe")
    dv = None if v3 is None else round(v - v3, 2)
    G.append(dict(k="bek12", ad="Enflasyon beklentisi (12 ay)", deger=f"%{f1(v)}", ay=etk(a),
                  isik="kirmizi" if dv is not None and dv > 0.5 else "sari" if dv is not None and dv > -0.5 else "yesil",
                  deg=dv, birim="puan",
                  not_=f"Yıllık TÜFE %{f1(vt)} ({etk(at)}). Beklentinin son 3 aydaki yönü politikaya güvenin ölçüsü; "
                       f"yükselen beklenti faiz indirim alanını daraltır."))
    a, v = son("cari12"); _, v12 = son("cari12", 12); _, vn12 = son("nhn12")
    G.append(dict(k="cari12", ad="Cari denge (12 aylık)", deger=f"{f1(v)} mlr $", ay=etk(a),
                  isik=isik(-v if v is not None else None, (20, 40), ters=True),
                  deg=None if v12 is None else round(v - v12, 1), birim="mlr $",
                  not_=f"Net hata ve noksan 12 ayda {f1(vn12)} mlr $: " + (
                      "kayıt dışı giriş açığı finanse ediyor; bu kaynak her an tersine dönebilir." if (vn12 or 0) > 0 else
                      "kayıt dışı çıkış var; açığın yanında bu çıkış da rezervlerden ve borçlanmayla karşılanıyor.")))
    a, v = son("dolar"); _, v12 = son("dolar", 12)
    dv = None if v12 is None else round(v - v12, 1)
    G.append(dict(k="dolar", ad="Dolarizasyon (DTH / mevduat)", deger=f"%{f1(v)}", ay=etk(a),
                  isik="kirmizi" if dv is not None and dv > 1 else "sari" if dv is not None and dv > -1 else "yesil",
                  deg=dv, birim="puan",
                  not_="Yurt içi yerleşiklerin TL'ye güveni. Bir yılda 1 puandan çok artış dövize kaçış işaretidir."))
    hs = [x for x in zip(aylar, seri["hisse"], seri["dibs"]) if x[1] is not None and x[2] is not None and x[0] <= kesim]
    son3 = hs[-3:]
    t3 = round(sum(h + d for _, h, d in son3), 2) if son3 else None
    t3d = round(sum(d for _, _, d in son3), 2) if son3 else None
    G.append(dict(k="akim", ad="Yabancı portföy (son 3 ay)", deger=f"{f1(t3)} mlr $", ay=etk(son3[-1][0]) if son3 else "–",
                  isik=isik(t3, (-2, 2)), deg=None, birim="mlr $",
                  not_=f"Hisse + DİBS net alımı; DİBS payı {f1(t3d)} mlr $. Yabancının TL varlıklara iştahı"
                       + (f"; {etk(son3[-1][0]).split()[0]} ayı kısmi." if son3 and son3[-1][0] == alindi_ayi else ".")))
    a, v = son("reer"); _, v12 = son("reer", 12)
    G.append(dict(k="reer", ad="Reel efektif kur (2025=100)", deger=f1(v), ay=etk(a),
                  isik="kirmizi" if v and v > 112 else "sari" if v and (v > 105 or v < 92) else "yesil",
                  deg=None if v12 is None else round(v - v12, 1), birim="puan",
                  not_="100'ün üstü TL'nin 2025 ortalamasına göre reel değerli olduğunu gösterir: taşıma işlemi için "
                       "olumlu, ihracat rekabeti ve cari denge için olumsuz."))
    a, v = son("makas"); _, v3 = son("makas", 3); _, vk = son("kredi"); _, vm = son("mevduat")
    G.append(dict(k="makas", ad="Kredi–mevduat faiz farkı", deger=f"{f1(v)} puan", ay=etk(a),
                  isik=isik(v, (4, 8)), deg=None if v3 is None else round(v - v3, 1), birim="puan",
                  not_=f"TL ticari kredi %{f1(vk)} − TL mevduat %{f1(vm)} (yeni işlemler, aylık ortalama). "
                       f"Bankaların marjının öncü göstergesi; bilançolara 1–2 çeyrek gecikmeyle yansır."))
    def son_ay_k(k):
        return son_ay(k, lambda a: a <= kesim)
    bek = {k: (son_ay_k(k), g(k).get(son_ay_k(k)) if son_ay_k(k) else None) for k in
           ("TP.PKAUO.S07.A.U", "TP.PKAUO.S07.B.U", "TP.PKAUO.S01.D.U", "TP.PKAUO.S01.I.U", "TP.PKAUO.S01.E.U",
            "TP.PKAUO.S01.F.U", "TP.PKAUO.S04.H.U", "TP.PKAUO.S04.I.U", "TP.PKAUO.S04.D.U", "TP.PKAUO.S04.E.U",
            "TP.PKAUO.S05.B.U", "TP.PKAUO.S05.C.U", "TP.PKAUO.S06.A.U", "TP.PKAUO.S06.B.U")}
    return {"gosterge": G, "beklenti": bek, "kesim": kesim}


def tufe_ceyrek():
    """yatirimci_ozeti için çeyrek sonu yıllık TÜFE (makro_veri.json varsa)."""
    if not VERI.exists():
        return {}
    s = json.loads(VERI.read_text(encoding="utf-8"))["seriler"].get("TP.TUKFIY2025.GENEL#yillik", {})
    return {a: round(v, 2) for a, v in s.items() if a[5:] in ("03", "06", "09", "12")}


# ----------------------------------------------------------------------------- sayfa
HTML = r"""<!--MKR-->
  <section aria-labelledby="h-mkr" id="mkrSec">
    <div class="sec-head">
      <h2 id="h-mkr">Makro pano</h2>
      <p class="insight" id="mkrInsight"></p>
    </div>
    <div class="mkr-tiles" id="mkrTiles"></div>
    <div class="panel-grid" style="grid-template-columns:repeat(auto-fit,minmax(min(100%,460px),1fr));margin-top:24px">
      <div class="panel" style="min-width:0"><h3>TCMB rezervleri</h3><div class="chartbox tall"><canvas id="mkrResChart" role="img" aria-label="Brüt, net ve swap hariç net rezerv"></canvas></div><p class="muted" style="font-size:12.5px">Milyar $, ay sonu. Net = analitik bilançoda dış varlıklar − toplam döviz yükümlülükleri (kurla dolara çevrildi). Swap hariç = net − TCMB'nin vadeli işlem ve swaplarla teslim edeceği döviz (Uluslararası Rezervler ve Döviz Likiditesi tablosu 2.2.1).</p></div>
      <div class="panel" style="min-width:0"><div class="controls"><label class="field" for="mkrMetric">Konu<select id="mkrMetric"></select></label></div><div class="chartbox tall"><canvas id="mkrChart" role="img" aria-label="Seçilen makro konunun seyri"></canvas></div><p class="muted" id="mkrNote" style="font-size:12.5px"></p></div>
    </div>
    <div class="panel" style="margin-top:24px"><h3>Piyasa katılımcılarının beklentileri</h3><div style="overflow-x:auto"><table class="ratio-table" id="mkrBek"></table></div><p class="muted" style="font-size:12.5px">TCMB Piyasa Katılımcıları Anketi; seçili dönem sonundaki ya da ondan önceki son anket.<span id="mkrBekAy"></span> GSYH ve cari denge yalnız takvim yılı için, USD/TRY yalnız bu yıl sonu ve 12 ay sonrası için sorulur (·). Gelecek yıl sonu TÜFE sorusu Eylül 2021'den, politika faizinin yıl sonu soruları Haziran 2025'ten beri ankettedir; daha önceki dönemlerde –.</p></div>
    <details class="yoz-yontem"><summary>Kaynak ve eşikler</summary>
      <ul>
        <li>Bütün veriler TCMB EVDS: TP.AB.TOPLAM, TP.AB.A02, TP.AB.A10, TP.DOVVARNC.K15, TP.DK.USD.A.YTL, TP.BISPOLFAIZ.TUR, TP.TUKFIY2025.GENEL, TP.PKAUO.*, TP.ODANA6.Q01, TP.ODEAYRSUNUM6.Q210, TP.KALANVADE.K20, TP.TLDTHVADE.KB12/KB18, TP.MKNETHAR.M7/M8, TP.KTF17, TP.TRY.MT06, TP.RK.T1.Y. <span id="mkrAlindi"></span></li>
        <li>Göstergeler sayfanın üstünde seçilen raporlama döneminin sonu itibarıyladır (ör. Eylül 2026 raporunda Eylül 2026 sonu); grafikler 2018'den o aya kadar çizilir. Çeyrek sonu verisi henüz yayımlanmamış göstergede o tarihten önceki son veri kullanılır ve kartta ayı yazar. Yabancı portföy akımları EVDS'de Eylül 2020'den başlar.</li>
        <li>Trafik ışıkları: swap hariç net rezerv &lt;20 / 50 mlr $; rezerv yeterliliği &lt;0,8 / 1,0; reel faiz &lt;0 / 5 puan; beklentide 3 ayda ±0,5 puan; cari açık &gt;20 / 40 mlr $; dolarizasyonda yılda ±1 puan; yabancı akım ±2 mlr $; reel kur 92–105 arası yeşil, 112 üstü kırmızı; kredi–mevduat farkı &lt;4 / 8 puan. Eşikler yorum kolaylığı içindir, resmi ölçüt değildir.</li>
        <li>Swap hariç rezervde yalnız TCMB'nin yurt içi para karşılığı vadeli/swap açık pozisyonu düşüldü; diğer merkez bankalarıyla yapılan swapları ayrıca düşen hesaplar birkaç milyar $ daha düşük sonuç verir.</li>
        <li>Politika faizi EVDS'deki BIS serisinden; seri gecikirse son PPK kararı elle eklenir (makro_veri.json).</li>
      </ul>
    </details>
  </section>
<!--/MKR-->"""

CSS = r"""<style id="mkrCss">
.mkr-tiles{display:grid;grid-template-columns:repeat(auto-fill,minmax(min(100%,250px),1fr));gap:14px}
.mkr-tile{border:1px solid var(--line);border-radius:12px;padding:14px 16px;background:var(--surface);display:flex;flex-direction:column;gap:4px;position:relative}
.mkr-tile::before{content:"";position:absolute;left:0;top:12px;bottom:12px;width:4px;border-radius:0 3px 3px 0;background:var(--line)}
.mkr-tile.yesil::before{background:var(--good)} .mkr-tile.sari::before{background:var(--warm)} .mkr-tile.kirmizi::before{background:var(--bad)}
.mkr-tile .l{font-size:12.5px;color:var(--ink-2);font-weight:600;display:flex;justify-content:space-between;gap:8px}
.mkr-tile .v{font-family:var(--display);font-size:24px;font-weight:800;color:var(--ink)}
.mkr-tile .c{font-size:12px;color:var(--ink-3)}
.mkr-tile .n{font-size:12.5px;color:var(--ink-2);line-height:1.45}
.mkr-dot{display:inline-block;width:10px;height:10px;border-radius:50%;background:var(--line);flex:none;margin-top:3px}
.mkr-dot.yesil{background:var(--good)} .mkr-dot.sari{background:var(--warm)} .mkr-dot.kirmizi{background:var(--bad)}
</style>"""

JS = r"""/*MKR*/
function renderMakro(){
  const MM=window.MAKRO, sec=document.getElementById("mkrSec"); if(!MM||!sec) return;
  const p=(typeof curP!=="undefined"&&curP)||"", ks=Object.keys(MM.donemler).filter(k=>k!=="son").sort();
  const kk=MM.donemler[p]?p:(p>ks[ks.length-1]&&MM.donemler.son?"son":(ks.filter(k=>k<=p).pop()||ks[0]));
  const M=Object.assign({},MM,MM.donemler[kk]);
  const kes=MM.aylar.indexOf(M.kesim), bas=Math.max(0,MM.aylar.indexOf("2018-01"));
  const cut=a=>a.slice(bas,kes+1), L=cut(MM.etiket), S=Object.fromEntries(Object.entries(MM.seri).map(([k,v])=>[k,cut(v)]));
  const kesL=MM.etiket[kes]||"";
  const f1=x=>x==null?"–":nf1.format(x), sg=x=>x==null?"":(x>=0?"+":"−")+nf1.format(Math.abs(x));
  const ISIK={yesil:"olumlu",sari:"izle",kirmizi:"risk",gri:"veri yok"};
  document.getElementById("mkrTiles").innerHTML=M.gosterge.map(g=>`<div class="mkr-tile ${g.isik}"><span class="l"><span>${esc(g.ad)}</span><span class="mkr-dot ${g.isik}" title="${ISIK[g.isik]}"></span></span><span class="v">${esc(g.deger)}</span><span class="c">${esc(g.ay)}${g.deg!=null?` · ${sg(g.deg)} ${esc(g.birim)} (${g.k==="yeterlilik"||g.k==="cari12"||g.k==="dolar"||g.k==="reer"?"1 yılda":"3 ayda"})`:""}</span><span class="n">${esc(g.not_)}</span></div>`).join("");
  const gs=Object.fromEntries(M.gosterge.map(g=>[g.k,g]));
  const kir=M.gosterge.filter(g=>g.isik==="kirmizi").map(g=>g.ad), yes=M.gosterge.filter(g=>g.isik==="yesil").map(g=>g.ad), sar=M.gosterge.filter(g=>g.isik==="sari").length;
  document.getElementById("mkrInsight").innerHTML=`<b>${esc(kesL)} sonu itibarıyla</b> (raporun dönem sonu): ${M.gosterge.length} göstergeden ${yes.length} olumlu, ${sar} izlenmeli, ${kir.length} risk bölgesinde. `+
    (yes.length?`Güçlü yanlar: ${yes.join(", ")}. `:"")+(kir.length?`Kırılgan yanlar: <b>${kir.join(", ")}</b>. `:"")+
    `Bankalar için bağlantı: kredi–mevduat faiz farkı ${esc(gs.makas?gs.makas.deger:"–")}; bu fark aşağıdaki karnede yer alan swap düzeltmeli marjın öncüsüdür, rezerv ve kur oynaklığı ise swap maliyetini belirler.`;
  document.getElementById("mkrAlindi").textContent=M.alindi?`Veri ${M.alindi} tarihinde alındı.`:"";
  const ds=(lab,data,col,extra={})=>Object.assign({label:lab,data,borderColor:css(col),backgroundColor:css(col),borderWidth:2,tension:.25,spanGaps:true,pointRadius:0},extra);
  const opt=(yf,extra={})=>baseOpts(yf,Object.assign({plugins:{legend:{labels:{color:css("--ink-2"),boxWidth:10,boxHeight:10,font:{family:"Roboto",size:11}}},tooltip:{callbacks:{label:c=>`${c.dataset.label}: ${c.parsed.y==null?'–':yf(c.parsed.y)}`}}}},extra));
  mkChart("mkrResChart",{type:"line",data:{labels:L,datasets:[ds("Brüt",S.brut,"--b0"),ds("Net",S.net,"--b1"),ds("Swap hariç net",S.swapsiz,"--b4",{borderWidth:3})]},options:opt(x=>nf1.format(x)+" mlr $")});
  const T={
    faiz:["Faiz ve enflasyon",()=>({type:"line",data:{labels:L,datasets:[ds("Politika faizi",S.politika,"--b0"),ds("Yıllık TÜFE",S.tufe,"--b4"),ds("12 ay sonrası beklenti",S.bek12,"--b1",{borderDash:[6,4]}),ds("Reel faiz (beklentiye göre)",S.reel,"--b3",{borderWidth:3})]},options:opt(x=>"%"+nf1.format(x))}),"Reel faiz = politika faizi − 12 ay sonrası TÜFE beklentisi."],
    dis:["Dış finansman",()=>({type:"bar",data:{labels:L,datasets:[ds("12 aylık cari denge",S.cari12,"--b0",{type:"bar",borderRadius:2}),ds("12 aylık net hata noksan",S.nhn12,"--b1",{type:"line"}),ds("Dış finansman ihtiyacı",S.efr.map(x=>x==null?null:-x),"--b4",{type:"line",borderDash:[6,4]})]},options:opt(x=>nf1.format(x)+" mlr $")}),"Dış finansman ihtiyacı (eksi işaretle) = kalan vadesi 1 yıla kadar dış borç + 12 aylık cari açık."],
    yet:["Rezerv yeterliliği",()=>({type:"line",data:{labels:L,datasets:[ds("Brüt rezerv / dış finansman ihtiyacı",S.yeterlilik,"--b0",{borderWidth:3}),ds("Eşik (1,0)",L.map(()=>1),"--bad",{borderDash:[4,4],borderWidth:1.5})]},options:opt(x=>nf2.format(x)+"x")}),"Guidotti–Greenspan ölçütü: brüt rezerv bir yıllık dış ödemeleri karşılamalı (≥1)."],
    akim:["Yabancı portföy akımları",()=>({type:"bar",data:{labels:L,datasets:[ds("Hisse",S.hisse,"--b0",{borderRadius:2}),ds("DİBS",S.dibs,"--b1",{borderRadius:2})]},options:opt(x=>nf2.format(x)+" mlr $",{scales:{x:{stacked:true,ticks:{color:css("--ink-2")}},y:{stacked:true,ticks:{color:css("--ink-2")}}}})}),"Aylık net alım (fiyat ve kur etkisinden arındırılmış); son ay kısmi."],
    dolar:["Dolarizasyon ve reel kur",()=>({type:"line",data:{labels:L,datasets:[ds("DTH / toplam mevduat (%)",S.dolar,"--b0",{yAxisID:"y"}),ds("Reel efektif kur (sağ)",S.reer,"--b4",{yAxisID:"y1"})]},options:opt(x=>nf1.format(x),{scales:{y:{position:"left",ticks:{color:css("--ink-2")}},y1:{position:"right",grid:{drawOnChartArea:false},ticks:{color:css("--ink-2")}},x:{ticks:{color:css("--ink-2")}}}})}),"Reel kur yükselişi TL'nin reel değerlenmesi demek."],
    kredi:["Kredi ve mevduat faizleri",()=>({type:"line",data:{labels:L,datasets:[ds("TL ticari kredi",S.kredi,"--b0"),ds("TL mevduat",S.mevduat,"--b1"),ds("Fark",S.makas,"--b3",{borderWidth:3})]},options:opt(x=>"%"+nf1.format(x))}),"Yeni işlemlerin aylık ortalama faizi (akım)."]
  };
  const sel=document.getElementById("mkrMetric");
  if(!sel.options.length){ Object.entries(T).forEach(([k,[l]])=>sel.appendChild(el("option",{value:k},l))); let s0="faiz"; try{ s0=localStorage.getItem("mkr.konu")||s0; }catch(e){} sel.value=T[s0]?s0:"faiz"; sel.onchange=()=>{ try{localStorage.setItem("mkr.konu",sel.value)}catch(e){} renderMakro(); }; }
  const t=T[sel.value]; mkChart("mkrChart",t[1]()); document.getElementById("mkrNote").textContent=t[2];
  const B=M.beklenti, bv=k=>B[k]&&B[k][1]!=null?B[k][1]:null, by=(k)=>B[k]&&B[k][0]?B[k][0]:"";
  const yil=+(by("TP.PKAUO.S07.A.U")||M.kesim).slice(0,4);
  const ay=by("TP.PKAUO.S01.E.U")||by("TP.PKAUO.S07.A.U"), YOK='<span class="muted" title="Ankette sorulmuyor">·</span>';
  const yp=k=>bv(k)==null?"–":"%"+f1(bv(k)), yk=k=>bv(k)==null?"–":nf2.format(bv(k)), yc=k=>bv(k)==null?"–":f1(bv(k)/1000);
  document.getElementById("mkrBek").innerHTML=`<thead><tr><th>Gösterge</th><th>${yil} yıl sonu</th><th>${yil+1} yıl sonu</th><th>12 ay sonra</th><th>24 ay sonra</th></tr></thead><tbody>
   <tr><td>TÜFE (yıllık)</td><td>${yp("TP.PKAUO.S01.D.U")}</td><td>${yp("TP.PKAUO.S01.I.U")}</td><td>${yp("TP.PKAUO.S01.E.U")}</td><td>${yp("TP.PKAUO.S01.F.U")}</td></tr>
   <tr><td>Politika faizi</td><td>${yp("TP.PKAUO.S04.H.U")}</td><td>${yp("TP.PKAUO.S04.I.U")}</td><td>${yp("TP.PKAUO.S04.D.U")}</td><td>${yp("TP.PKAUO.S04.E.U")}</td></tr>
   <tr><td>USD/TRY</td><td>${yk("TP.PKAUO.S05.B.U")}</td><td>${YOK}</td><td>${yk("TP.PKAUO.S05.C.U")}</td><td>${YOK}</td></tr>
   <tr><td>GSYH büyümesi (yıl)</td><td>${yp("TP.PKAUO.S07.A.U")}</td><td>${yp("TP.PKAUO.S07.B.U")}</td><td>${YOK}</td><td>${YOK}</td></tr>
   <tr><td>Cari denge (yıl, mlr $)</td><td>${yc("TP.PKAUO.S06.A.U")}</td><td>${yc("TP.PKAUO.S06.B.U")}</td><td>${YOK}</td><td>${YOK}</td></tr></tbody>`;
  const ayEl=document.getElementById("mkrBekAy"); if(ayEl&&ay) ayEl.textContent=` Anket: ${["Ocak","Şubat","Mart","Nisan","Mayıs","Haziran","Temmuz","Ağustos","Eylül","Ekim","Kasım","Aralık"][+ay.slice(5,7)-1]} ${ay.slice(0,4)}.`;
}
/*/MKR*/
"""


def temizle(t, bas, son):
    while bas in t:
        i = t.index(bas)
        j = t.index(son, i) + len(son)
        t = t[:i] + t[j:]
    return t


TCMB_SAYFA = r"""<!doctype html>
<html lang="tr"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>TCMB verileri · Makro pano</title>
__HEAD__
<style>.tc-top{display:flex;align-items:center;gap:12px;padding:12px 0}.tc-top a{color:var(--accent);font-weight:500;text-decoration:none}
.tc-top a:hover{text-decoration:underline}</style>
</head><body>
<div class="wrap gut fade-in">
  <div class="tc-top"><a href="../">← Başlangıç ekranı</a></div>
  <header class="hdr">
    <div class="hdr-top"><span class="eyebrow">TCMB · Elektronik Veri Dağıtım Sistemi (EVDS)</span></div>
    <h1>TCMB verileri: makro pano</h1>
    <p class="lede">Rezervler, faiz ve enflasyon, dış finansman, dolarizasyon, yabancı portföy akımları, reel kur ve kredi–mevduat faiz farkı; 2018'den bugüne aylık. Bir ay seçin: kartlar o ayın sonundaki değerleri, grafikler 2018'den o aya kadar olan seyri gösterir. Bankaların finansal raporlarında aynı pano raporun dönem sonuna bağlıdır.</p>
    <div class="hdr-meta"><span>Kaynak: TCMB EVDS</span><span id="tcAlindi"></span></div>
  </header>
  <div class="pbar" role="group" aria-label="Ay seçimi">
    <span class="pbar-label">Ay</span>
    <div class="pctl" id="periodSeg"></div>
  </div>
  __BOLUM__
</div>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
__VERI__
<script>
const nf1=new Intl.NumberFormat("tr-TR",{minimumFractionDigits:1,maximumFractionDigits:1});
const nf2=new Intl.NumberFormat("tr-TR",{minimumFractionDigits:2,maximumFractionDigits:2});
const css=n=>getComputedStyle(document.documentElement).getPropertyValue(n).trim();
const el=(tag,attrs={},html="")=>{ const e=document.createElement(tag); for(const k in attrs) e.setAttribute(k,attrs[k]); if(html) e.innerHTML=html; return e; };
const esc=s=>String(s).replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const charts={};
function mkChart(id,cfg){ if(charts[id]) charts[id].destroy(); const e=document.getElementById(id); if(!e||!window.Chart) return; charts[id]=new Chart(e,cfg); return charts[id]; }
__BASEOPTS__
const AYK=["Oca","Şub","Mar","Nis","May","Haz","Tem","Ağu","Eyl","Eki","Kas","Ara"];
const P=Object.keys(window.MAKRO.donemler).filter(k=>k!=="son").sort();
let curP=P[P.length-1];
try{ const s=localStorage.getItem("tcmb.ay"); if(s&&P.includes(s)) curP=s; }catch(e){}
function setP(p){ curP=p; try{localStorage.setItem("tcmb.ay",p)}catch(e){} renderAll(); }
function renderPeriodBar(){
  const box=document.getElementById("periodSeg"); box.innerHTML="";
  const i=P.indexOf(curP), y=curP.slice(0,4);
  const prev=el("button",{type:"button",class:"icon-btn","aria-label":"Önceki ay",title:"Önceki ay"},'<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M15.41 7.41 14 6l-6 6 6 6 1.41-1.41L10.83 12z"/></svg>');
  const next=el("button",{type:"button",class:"icon-btn","aria-label":"Sonraki ay",title:"Sonraki ay"},'<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M10 6 8.59 7.41 13.17 12l-4.58 4.59L10 18l6-6z"/></svg>');
  if(i<=0) prev.disabled=true; if(i>=P.length-1) next.disabled=true;
  prev.onclick=()=>setP(P[i-1]); next.onclick=()=>setP(P[i+1]);
  const ys=el("select",{"aria-label":"Yıl"}); [...new Set(P.map(p=>p.slice(0,4)))].forEach(v=>ys.appendChild(el("option",{value:v},v))); ys.value=y;
  ys.onchange=()=>{ let p=ys.value+curP.slice(4); if(!P.includes(p)) p=P.filter(q=>q.startsWith(ys.value)).pop(); setP(p); };
  const qs=el("div",{class:"seg",role:"group","aria-label":"Ay"});
  AYK.forEach((l,k)=>{ const p=y+"-"+String(k+1).padStart(2,"0"); const b=el("button",{type:"button","aria-pressed":String(p===curP)},l); if(!P.includes(p)) b.disabled=true; b.onclick=()=>setP(p); qs.appendChild(b); });
  box.append(prev,ys,qs,next);
}
__JS__
function renderAll(){ renderPeriodBar(); renderMakro(); }
document.getElementById("tcAlindi").textContent=window.MAKRO.alindi?`Veri ${window.MAKRO.alindi} tarihinde alındı`:"";
renderAll();
</script>
</body></html>
"""


def tcmb_sayfa(site: Path, V: dict):
    """site/tcmb/index.html: makro pano tek başına, aylık seçimle (raporun stil ve grafik ayarlarıyla)."""
    html = (site / "index.html").read_text(encoding="utf-8")
    bas = html[:html.index("</head>")]
    bas = bas[bas.index("<link"):] if "<link" in bas else bas            # yazı tipleri + stiller
    bas = re.sub(r"<script\b.*?</script>", "", bas, flags=re.S)
    i = html.index("function baseOpts")
    base = html[i:html.index("\n}\n", i) + 3]
    M = hesapla(V, aylik=True)
    bolum = HTML.replace('<!--MKR-->', '').replace('<!--/MKR-->', '').replace(
        "seçili dönem sonundaki ya da ondan önceki son anket", "seçili ayın sonundaki ya da ondan önceki son anket").replace(
        "Göstergeler sayfanın üstünde seçilen raporlama döneminin sonu itibarıyladır (ör. Eylül 2026 raporunda Eylül 2026 sonu); grafikler 2018'den o aya kadar çizilir.",
        "Göstergeler seçilen ayın sonu itibarıyladır; grafikler 2018'den o aya kadar çizilir.")
    js = JS.replace(" (raporun dönem sonu):", ":").replace("bu fark aşağıdaki karnede yer alan swap düzeltmeli marjın öncüsüdür",
                                                           "bu fark finansal raporlardaki swap düzeltmeli marjın öncüsüdür")
    veri = "<script>window.MAKRO=" + json.dumps(M, ensure_ascii=False, separators=(",", ":")) + ";</script>"
    out = (TCMB_SAYFA.replace("__HEAD__", bas).replace("__BOLUM__", bolum).replace("__VERI__", veri)
           .replace("__BASEOPTS__", base).replace("__JS__", js))
    (site / "tcmb").mkdir(exist_ok=True)
    (site / "tcmb" / "index.html").write_text(out, encoding="utf-8")
    print(f"TCMB sayfası: {site / 'tcmb' / 'index.html'} ({len(M['donemler'])} ay)")


def uygula(site: Path):
    V = json.loads(VERI.read_text(encoding="utf-8"))
    M = hesapla(V)
    yol = site / "index.html"
    html = yol.read_text(encoding="utf-8")
    html = temizle(html, "<!--MKR-->", "<!--/MKR-->")
    html = temizle(html, "/*MKR*/", "/*/MKR*/")
    html = re.sub(r'<style id="mkrCss">.*?</style>', "", html, flags=re.S)
    html = re.sub(r'<script id="mkrData">.*?</script>\s*', "", html, flags=re.S)
    html = html.replace("renderMakro(); ", "")
    # bölüm: yatırımcı özetinden (karne) önce, o yoksa ilk 7 toplamından önce — makrodan bankalara
    i = html.index("<!--YOZ-->") if "<!--YOZ-->" in html else html.index('<section aria-labelledby="h-sum">')
    html = html[:i] + HTML + "\n\n  " + html[i:]
    html = html.replace("</head>", CSS + "\n</head>", 1)
    veri = "<script id=\"mkrData\">window.MAKRO=" + json.dumps(M, ensure_ascii=False, separators=(",", ":")) + ";</script>\n"
    i = html.index("<script>\nconst DATA = ")
    html = html[:i] + veri + html[i:]
    imza = "function renderAll(){ renderPeriodBar(); "
    assert html.count(imza) == 1, "rapor renderAll bulunamadı"
    html = html.replace(imza, JS + imza + "renderMakro(); ", 1)
    yol.write_text(html, encoding="utf-8")
    try:                                                         # TCMB sayfası: günlük/haftalık arşiv varsa tam sürüm
        import tcmb_sayfa
        tcmb_sayfa.uret(site)
    except (Exception, SystemExit) as e:                         # noqa: BLE001
        print(f"  TCMB sayfası arşivsiz üretildi ({e})")
        tcmb_sayfa(site, V)
    try:
        import giris_sayfasi
        giris_sayfasi.uygula(site)
    except Exception as e:                                       # noqa: BLE001
        print(f"  ! giriş ekranı güncellenemedi: {e}")
    isik = {}
    son_d = M["donemler"].get("son") or M["donemler"][max(k for k in M["donemler"] if k != "son")]
    for g in son_d["gosterge"]:
        isik[g["isik"]] = isik.get(g["isik"], 0) + 1
    print(f"makro pano eklendi: {yol} (veri {M['alindi']}; ışıklar {isik})")
    return M


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--site", required=True)
    ap.add_argument("--anahtar", default=os.environ.get("EVDS_API_KEY"), help="EVDS API anahtarı")
    ap.add_argument("--guncelleme", action="store_true", help="EVDS'den çekme (veri_hatti.py zaten çekti)")
    a = ap.parse_args()
    if a.anahtar and not a.guncelleme and not os.environ.get("VERI_HATTI"):
        guncelle(a.anahtar)
    uygula(Path(a.site))
