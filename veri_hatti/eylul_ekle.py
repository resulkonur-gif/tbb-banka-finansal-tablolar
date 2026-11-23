#!/usr/bin/env python3
"""
KAP PDF'lerinden okunan son çeyreği (ör. Eylül 2026) TBB raporunun verisine ekler ve
raporun önizleme kopyasını üretir: <cikti>/index.html ve değişen dipnot dosyaları
<cikti>/dipnot/*.js. Değişmeyen dipnotlar ana sitedeki ../dipnot/ klasöründen yüklenir;
bu yüzden önizleme, sitenin docs/onizleme/ klasörüne konacak biçimdedir.

  python eylul_ekle.py --donem 202609 --pdf-klasoru banka_karsilastirma_202609 \
         --site docs --cikti docs/onizleme

  --site   : sitenin docs klasörü (index.html ve dipnot/ buradan okunur).
             Verilmezse GitHub'dan indirilir.

Okuma ve eşleme banka_karsilastir.py ile aynıdır (aynı klasörde olmalı):
ana tablolar, kur riski, mali bünye tabloları, göstergeler, öteki dipnotlar.

Kurallar
  - Tutarlar rapordaki birimden bin TL'ye çevrilir (milyon TL × 1000); oranlar olduğu gibi.
  - İşaret: TBB bazı kalemleri raporun tersi işaretle tutar. Her kalem için işaret, aynı
    rapordaki önceki dönem değeri ile TBB'nin o dönemdeki değeri karşılaştırılarak öğrenilir
    (bütün bankalarda çoğunluk). Önceki dönemi raporda olmayan kalemde TBB'nin önceki dönem
    işareti esas alınır.
  - Ara toplamlar: raporda okunmayan bir üst kalem, TBB'de önceki dönemde alt kalemlerinin
    toplamına eşitse, cari dönemde de alt kalemlerin toplamı olarak doldurulur.
  - PDF'te bulunamayan kalemler boş kalır.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import banka_karsilastir as bk   # noqa: E402

AY_KISA = {3: "Mar", 6: "Haz", 9: "Eyl", 12: "Ara"}
AY_UZUN = {3: "Mart", 6: "Haziran", 9: "Eylül", 12: "Aralık"}


def veri_ayir(html: str):
    i = html.index("const DATA = ") + len("const DATA = ")
    D, son = json.JSONDecoder().raw_decode(html, i)
    return D, i, son


def bankanin_pdfleri(klasor: Path):
    out = []
    for b in bk.BANKALAR:
        k = klasor / bk.dosya_adi(b[2])
        if not k.exists():
            continue
        bulunan = {}
        for p in sorted(k.glob("*.pdf")):
            t = bk.pdf_turu(p)
            if t and t not in bulunan:
                bulunan[t] = p
        out += [(b, t, p) for t, p in bulunan.items()]
    return out


def agac(items):
    """kod → çocuk kodları (sv seviyesine göre)."""
    cocuk, yigin = defaultdict(list), []
    for kod, ad, sv, *_ in items:
        while yigin and yigin[-1][1] >= sv:
            yigin.pop()
        if yigin:
            cocuk[yigin[-1][0]].append(kod)
        yigin.append((kod, sv))
    return cocuk


HESAP_KOD = 9500000        # ana sitenin hesapladığı satırlar (toplam / net); TBB'de yok


def hesaplanan_duzelt(ham: dict, kur_riski: bool):
    """Hesaplanan satırlar: (1) alt satırı olan her hesaplanan satır, bütün dönemlerde alt satırlarının toplamı
    (ana sitenin üreticisi aynı kodlu toplamları dipnotlar arasında biriktiriyor: faiz riski 'Toplam Varlıklar' =
    kur riski + faiz riski …). (2) Kur riski: TBB'deki 'Net Bilanço Pozisyonu' ve 'Net Nazım Hesap Pozisyonu' ile
    aynı olan iki hesaplanan satır yerine 'Net genel pozisyon (bilanço içi + bilanço dışı)'."""
    items, V = ham["items"], ham["v"]
    cocuk = agac(items)

    def deger(p, kod):
        d = V.get(p, {}).get(str(kod)) or []
        return {d[i]: d[i + 1] for i in range(0, len(d), 2)}

    def yaz(p, kod, dd):
        if dd:
            V.setdefault(p, {})[str(kod)] = [x for b, v in sorted(dd.items()) for x in (b, v)]
        else:
            V.get(p, {}).pop(str(kod), None)
    for kod, *_ in reversed(items):
        if int(kod) < HESAP_KOD or not cocuk.get(kod):
            continue
        for p in list(V):
            top = {}
            for c in cocuk[kod]:
                for b, v in deger(p, c).items():
                    if v is not None:
                        top[b] = top.get(b, 0) + v
            yaz(p, kod, top)
    if not kur_riski:
        return
    kur_net_genel(ham)


def kur_net_genel(ham: dict):
    """Kur riski: 'Net Genel Pozisyon (bilanço içi + bilanço dışı)' satırı ve döviz cinsi kırılımları, her dönem kaynak
    satırlardan: (toplam varlıklar − |toplam yükümlülükler|) + (türev alacakları − |türev borçları|). Eski iki hesaplanan
    satır (TBB'deki net bilanço / net nazım pozisyonunun tekrarı) kaldırılır; boş TBB net kalemleri de doldurulur."""
    items, V = ham["items"], ham["v"]
    cocuk = agac(items)
    ad_of = {k: a for k, a, *_ in items}
    nk = lambda a: re.sub(r"\s+", " ", a).strip().lower()

    def bul(ust, kosul):
        return next((c for c in cocuk.get(ust, []) if kosul(nk(ad_of[c]))), None)

    def deger(p, kod):
        d = V.get(p, {}).get(str(kod)) or []
        return {d[i]: d[i + 1] for i in range(0, len(d), 2)}

    def yaz(p, kod, dd):
        if dd:
            V.setdefault(p, {})[str(kod)] = [x for b, v in sorted(dd.items()) for x in (b, v)]
    kok = items[0][0]
    nb = bul(kok, lambda n: n == "net bilanço pozisyonu")
    nn = bul(kok, lambda n: n == "net nazım hesap pozisyonu")
    if not (nb and nn):
        return
    va, yu = bul(nb, lambda n: n == "varlıklar"), bul(nb, lambda n: n == "yükümlülükler")
    tv = bul(va, lambda n: n.startswith("toplam varlıklar")) if va else None
    ty = bul(yu, lambda n: n.startswith("toplam yükümlülükler")) if yu else None
    al = bul(nn, lambda n: "türev" in n and "alacak" in n)
    bo = bul(nn, lambda n: "türev" in n and "borç" in n)
    if not (tv and ty and al and bo):
        return
    # eski hesaplanan net satırları kaldır
    sil = set()
    for ust in (nb, nn):
        for c in cocuk.get(ust, []):
            if int(c) >= HESAP_KOD and "(" in ad_of[c] and nk(ad_of[c]).startswith(("net bilanço", "net nazım")):
                sil |= {c} | set(cocuk.get(c, []))
    sutunlar = []
    for kaynak in (tv, ty, al, bo):
        for c in cocuk.get(kaynak, []):
            if ad_of[c] not in sutunlar:
                sutunlar.append(ad_of[c])
    ng = next((k for k, a, *_ in items if a.startswith("Net Genel Pozisyon")), None)
    if ng:
        kodlar = {ad_of[c]: c for c in cocuk.get(ng, [])}
    else:
        ng = 9590001
        kodlar = {}
    yeni_items = [it for it in items if it[0] not in sil]
    ek = list(items[0][3:4])
    if not any(it[0] == ng for it in yeni_items):
        yeni_items.append([ng, "Net Genel Pozisyon (bilanço içi + bilanço dışı)", 1] + ek)
    sira = max([int(k) for k in kodlar.values()] + [9590001])
    for ad in sutunlar:
        if ad not in kodlar:
            sira += 1
            kodlar[ad] = sira
            yeni_items.append([sira, ad, 2] + ek)
    sutun_of = lambda kaynak: {ad_of[c]: c for c in cocuk.get(kaynak, [])}
    S = {x: sutun_of(x) for x in (tv, ty, al, bo)}
    for p in list(V):
        for k in sil:
            V[p].pop(str(k), None)
        top, bi_top, bd_top = {}, {}, {}
        for ad in sutunlar:
            dd = {}
            parca = [(S[tv].get(ad), 1, bi_top), (S[ty].get(ad), -1, bi_top), (S[al].get(ad), 1, bd_top), (S[bo].get(ad), -1, bd_top)]
            for kod, isaret, hedef in parca:
                for b, v in (deger(p, kod) if kod else {}).items():
                    if v is None:
                        continue
                    x = v if isaret > 0 else -abs(v)
                    dd[b] = dd.get(b, 0) + x
                    hedef[b] = hedef.get(b, 0) + x
            yaz(p, kodlar[ad], dd)
            for b, v in dd.items():
                top[b] = top.get(b, 0) + v
        yaz(p, ng, top)
        for kod, dd in ((nb, bi_top), (nn, bd_top)):          # boş TBB net kalemleri
            mevcut = deger(p, kod)
            eksik = {b: v for b, v in dd.items() if mevcut.get(b) is None}
            if eksik and p == max(V):                       # yalnız en yeni dönem (KAP önizlemesi)
                yaz(p, kod, {**mevcut, **eksik})
    ham["items"] = yeni_items


def ara_toplam_doldur(items, al, koy, bs, onc, yeni, oncekiler=None):
    """Cari dönemde boş üst kalemi doldurur. Önceki dönemlerin hepsinde
    (a) üst kalem alt kalemlerin toplamıysa toplamla, (b) belirli bir alt kalemle aynıysa o alt kalemle
    (NİFO'da satır = 'oranı uygulanmış toplam'). Önceki dönemde boş olan alt kalem 0 sayılır."""
    donemler = [onc] + [p for p in (oncekiler or []) if p != onc]
    cocuk_tum = agac(items)
    # hesaplanan satırlar (kod ≥ 9.500.000) TBB kaleminin altında toplanmaz: "Varlıklar" = satırlar, "Toplam Varlıklar" değil;
    # tek alt kalemle eşitlik kuralında ise bakılır (Net Nazım Hesap Pozisyonu = alacaklar − borçlar satırı)
    cocuk = {k: [c for c in v if int(k) >= HESAP_KOD or int(c) < HESAP_KOD] for k, v in cocuk_tum.items()}
    eklenen = 0
    for kod, *_ in reversed(items):              # alttan üste
        cc = cocuk.get(kod)
        if not cc or al(yeni, kod, bs) is not None:
            continue
        gecerli = []
        for p in donemler:
            u = al(p, kod, bs)
            if u is None:
                continue
            gecerli.append((u, [al(p, c, bs) for c in cc]))
        if not gecerli:
            continue
        alt = [al(yeni, c, bs) for c in cc]
        # cari dönemde boş alt kalem: önceki dönemlerde de hep boş/0 ise 0 say
        hep_bos = [all(not ao[i] for _, ao in gecerli) for i in range(len(cc))]
        if all(x is None for x in alt) and all(al(yeni, c, bs) is None for c in cocuk_tum.get(kod, [])):
            continue
        alt0 = [0 if (x is None and hep_bos[i]) else x for i, x in enumerate(alt)]
        tol = lambda n: 2 * n + 1
        if all(x is not None for x in alt0) and all(
                abs(u - sum(x or 0 for x in ao)) <= tol(len(cc)) for u, ao in gecerli):
            koy(kod, bs, sum(alt0))
            eklenen += 1
            continue
        cc2 = cocuk_tum.get(kod, cc)
        for c in cc2:                              # tek bir alt kalemle hep aynı mı?
            x = al(yeni, c, bs)
            if x is None:
                continue
            if all(al(p, c, bs) is not None and abs(al(p, kod, bs) - al(p, c, bs)) <= 2 and abs(al(p, kod, bs)) > 0
                   for p in donemler if al(p, kod, bs) is not None):
                koy(kod, bs, x)
                eklenen += 1
                break
    return eklenen


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--donem", required=True, help="YYYYAA, ör. 202609")
    ap.add_argument("--pdf-klasoru", required=True)
    ap.add_argument("--site", help="sitenin docs klasörü (index.html + dipnot/)")
    ap.add_argument("--cikti", default="onizleme")
    a = ap.parse_args()
    yil, ay = int(a.donem[:4]), int(a.donem[4:])
    yeni = f"{yil}-{ay:02d}"
    onc_yil = f"{yil - 1}-12"

    site = Path(a.site) if a.site else None
    html = (site / "index.html").read_text(encoding="utf-8") if site else \
        bk.istek("GET", bk.TBB_URL).decode("utf-8")
    D, bas, son = veri_ayir(html)

    # dipnotlar: yerel sitedeki dosyalardan (ham hâliyle de saklanır, yazarken gerekiyor)
    dn_ham = {}

    def yerel_dipnot(D_, tur, ad):
        dn = next((x for x in D_[tur].get("DN", []) if x["ad"] == ad), None)
        if not dn:
            return None
        if dn["id"] not in dn_ham:
            if site:
                t = (site / "dipnot" / f"{dn['id']}.js").read_text(encoding="utf-8")
            else:
                t = bk.istek("GET", bk.TBB_URL.replace("index.html", f"dipnot/{dn['id']}.js")).decode("utf-8")
            m = re.match(r'\s*__DN\("([^"]+)",(.*)\);\s*$', t, re.S)
            dn_ham[dn["id"]] = json.loads(m.group(2))
        return dn_ham[dn["id"]]
    bk.tbb_dipnotu = yerel_dipnot

    raporlar = bankanin_pdfleri(Path(a.pdf_klasoru))
    print(f"{len(raporlar)} rapor: " + ", ".join(f"{b[2]} {t}" for b, t, _ in raporlar))
    for tur in ("solo", "kons"):
        if yeni in D[tur]["periods"]:
            raise SystemExit(f"{yeni} TBB verisinde zaten var; önizlemeye gerek yok.")

    # 1) bütün raporları banka_karsilastir'ın karşılaştırma akışıyla oku
    rapor = bk.Rapor(yil, ay, D)
    for b, tur, p in raporlar:
        print(f"  {b[2]} {tur}: okunuyor…", flush=True)
        ok = bk.Okuyucu(p).oku()
        rapor.ekle(b, tur, ok, p.name)
        rapor.ekle_kur(b, tur, ok)
        rapor.ekle_dipnot(b, tur, ok)
        rapor.ekle_gosterge(b, tur, ok)
        rapor.ekle_genel(b, tur, ok)

    kayit = [k for k in rapor.kayit if k.get("hedef") and k["c"] is not None
             and (k["hedef"] != "A" and k["hedef"][0] == "G" or str(k["kod"]).isdigit())]

    # 2) işaret: önceki dönem PDF ile TBB karşılaştırması (bütün bankalarda çoğunluk)
    oy = defaultdict(Counter)
    for k in kayit:
        if k["o"] and k["t_ham"]:
            oy[(k["tur"], str(k["hedef"]), k["kod"])]["ayni" if (k["o"] > 0) == (k["t_ham"] > 0) else "ters"] += 1

    def isaret(k):
        # önce bankanın kendi önceki dönem değeri: bankalar aynı kalemi farklı işaretle yazabiliyor
        # (vergi karşılığını Ziraat/VakıfBank negatif, öteki bankalar pozitif yazıyor)
        if k["o"] and k["t_ham"] and abs(abs(k["o"] * k["carp"]) - abs(k["t_ham"])) <= 0.02 * abs(k["t_ham"]) + k["carp"]:
            return 1 if (k["o"] > 0) == (k["t_ham"] > 0) else -1
        # gelir tablosu yıl başından birikimli: önceki yıl değeri yoksa, Eylül tutarı Haziran'ınkini aşıyorsa
        # bankanın Haziran TBB değerinin işareti esas alınır
        if k["hedef"] == "A" and "Gelir-Gider" in str(k.get("tablo", "")) and k["c"]:
            onc_p = [p_ for p_ in D[k["tur"]]["periods"] if p_ < yeni][-1]
            dizi = (D[k["tur"]]["A"]["v"].get(onc_p) or {}).get(str(k["kod"])) or []
            tv = dizi[k["bs"]] if k["bs"] < len(dizi) else None
            if tv and abs(k["c"] * k["carp"]) >= 0.9 * abs(tv):
                return 1 if (k["c"] > 0) == (tv > 0) else -1
        o = oy.get((k["tur"], str(k["hedef"]), k["kod"]))
        if o:
            return -1 if o["ters"] > o["ayni"] else 1
        if k["t_ham"] and k["c"] and (k["t_ham"] > 0) != (k["c"] > 0):
            return -1
        return 1

    # 3) yeni dönemi ekle
    degisen_dn = {}           # id → (tur, ham)
    ozet = Counter()
    for tur in ("solo", "kons"):
        S, A = D[tur], D[tur]["A"]
        nb = len(A["codes"])
        S["periods"].append(yeni)
        S["plabel"][yeni] = f"{AY_KISA[ay]} {yil}"
        S["months"][yeni] = ay
        onceki = S["periods"][-2]
        A["v"][yeni] = {kod: [None] * nb for kod in A["v"][onceki]}
        for k in ("yp", "hqla", "out", "lcr", "lcrYp", "asf", "rsf", "nsfr", "cet1", "t1", "car", "buf", "ccb",
                  "ccyb", "sib", "bufR"):
            if k in A:
                A[k][yeni] = [None] * nb
        bankalar = set()
        for k in kayit:
            if k["tur"] != tur:
                continue
            bs, h = k["bs"], k["hedef"]
            bankalar.add(bs)
            oran = k["carp"] == 1
            deger = k["c"] * k["carp"] * isaret(k)
            deger = deger if oran else round(deger)
            if h == "A":
                A["v"][yeni].setdefault(str(k["kod"]), [None] * nb)[bs] = deger
                ozet[(tur, "ana tablo")] += 1
            elif h[0] == "G":
                if h[1] in A:
                    A[h[1]][yeni][bs] = deger
                    ozet[(tur, "gösterge")] += 1
            elif h[0] == "DN":
                dn = next((x for x in S["DN"] if x["ad"] == h[1]), None)
                if not dn or dn["id"] not in dn_ham:
                    continue
                # yalnız paketteki kalemler: karşılaştırma için türetilen yardımcı satırlar (kur riski döviz
                # cinsine göre net bilanço/nazım pozisyonu, 9580001+) sayfaya yazılmaz
                if not any(it[0] == k["kod"] for it in dn_ham[dn["id"]]["items"]):
                    continue
                v = dn_ham[dn["id"]]["v"].setdefault(yeni, {})
                dizi = v.setdefault(str(k["kod"]), [])
                for i in range(0, len(dizi), 2):
                    if dizi[i] == bs:
                        dizi[i + 1] = deger
                        break
                else:
                    dizi += [bs, deger]
                degisen_dn[dn["id"]] = (tur, dn_ham[dn["id"]])
                ozet[(tur, "dipnot")] += 1

        # 4) ara toplamlar: ana tablolar (mevduat vade yapısı gibi) ve dipnotlar
        def a_al(dn_, kod, bs):
            x = A["v"].get(dn_, {}).get(str(kod))
            return x[bs] if x and bs < len(x) else None

        def a_koy(kod, bs, val):
            A["v"][yeni].setdefault(str(kod), [None] * nb)[bs] = val
        for bs in bankalar:
            for tablo in {it[3] for it in S["items"]}:
                its = [it for it in S["items"] if it[3] == tablo]
                ozet[(tur, "ara toplam")] += ara_toplam_doldur(its, a_al, a_koy, bs, onc_yil, yeni,
                                                               sorted(p for p in A["v"] if p < yeni)[-2:])
        for did, (t_, ham) in list(degisen_dn.items()):
            if t_ != tur:
                continue

            def d_al(dn_, kod, bs, ham=ham):
                dizi = ham["v"].get(dn_, {}).get(str(kod)) or []
                return next((dizi[i + 1] for i in range(0, len(dizi), 2) if dizi[i] == bs), None)

            def d_koy(kod, bs, val, ham=ham):
                ham["v"].setdefault(yeni, {}).setdefault(str(kod), []).extend([bs, val])
            onc_dn = onc_yil if onc_yil in ham["v"] else max(p for p in ham["v"] if p < yeni)
            for bs in bankalar:
                ozet[(tur, "ara toplam")] += ara_toplam_doldur(ham["items"], d_al, d_koy, bs, onc_dn, yeni,
                                                               sorted(p for p in ham["v"] if p < yeni)[-2:])

    # hesaplanan satırlar (bütün dönemler) ve kur riski net genel pozisyon
    for did, (_, ham) in degisen_dn.items():
        hesaplanan_duzelt(ham, ham["items"][0][1].startswith("Kur riskine"))

    # 5) yaz
    cikti = Path(a.cikti)
    (cikti / "dipnot").mkdir(parents=True, exist_ok=True)
    yeni_html = html[:bas] + json.dumps(D, ensure_ascii=False, separators=(",", ":")) + html[son:]
    # değişen dipnotlar önizleme klasöründen, ötekiler ana sitenin dipnot klasöründen yüklenir
    yeni_html = yeni_html.replace(
        'const url="dipnot/"+key+".js";',
        'const url=(ONIZLEME_DN.has(key)?"dipnot/":"../dipnot/")+key+".js";', 1)
    yeni_html = yeni_html.replace(
        "const DN_CACHE={}, DN_WAIT={};",
        "const DN_CACHE={}, DN_WAIT={};\nconst ONIZLEME_DN=new Set(" + json.dumps(sorted(degisen_dn)) + ");", 1)
    eski = "const depth=SRC.rows.length>120?1:2; SRC.rows.forEach((r,i)=>{ if(r.kids && r.lv<depth) OPEN.add(i); });"
    if eski in yeni_html:          # oran dipnotları (satırların çoğu %) seçilince ağacın tamamı açık gelir
        yeni_html = yeni_html.replace(eski, "const depth=SRC.rows.length>120?1:2, oranT=SRC.rows.filter(r=>r.r).length*2>"
                                      "SRC.rows.length; SRC.rows.forEach((r,i)=>{ if(r.kids && (oranT || r.lv<depth)) "
                                      "OPEN.add(i); });", 1)
    elif "oranT=SRC.rows" not in yeni_html:                # ana sitede zaten varsa sessiz geç
        print("   ! ağaç açma kodu bulunamadı (oran tabloları)")
    # tablo görünümü: düz listede tekrarlanan sütun adları bölüm başlığıyla ayrılır (risk grubu c.2: "GUD … / Riskten
    # korunma …" altında aynı "Dönem Başı / Dönem Sonu / Toplam Kâr-Zarar"); başlık satırı sütun olmaz
    mk = ('const mk=ks=>{ const nm=ks.map(j=>nk(rows[j].n)), cnt={}; nm.forEach(n=>cnt[n]=(cnt[n]||0)+1); let sec=""; '
          'return ks.map((j,q)=>{ const n=nm[q]; if(cnt[n]===1 && q+1<ks.length && cnt[nm[q+1]]>1){ sec=rows[j].n; return null; } '
          'return (cnt[n]>1&&sec)? sec+" · "+rows[j].n : rows[j].n; }); };')
    eski1 = ('const cols=[], ck={}; rows.forEach((r,i)=>{ if(isM(i)) kids[i].forEach(j=>{ const k=nk(rows[j].n); '
             'if(!(k in ck)){ ck[k]=cols.length; cols.push({k,l:rows[j].n}); } }); });\n  return SRC.mx={kids,isM,cols,')
    yeni1 = (mk + ' const cols=[], ck={}; rows.forEach((r,i)=>{ if(isM(i)){ const ky=mk(kids[i]); kids[i].forEach((j,q)=>{ '
             'if(ky[q]==null) return; const k=nk(ky[q]); if(!(k in ck)){ ck[k]=cols.length; cols.push({k,l:ky[q]}); } }); } });\n'
             '  return SRC.mx={kids,isM,cols,mk,')
    eski2 = "kids[i].forEach(j=>{ cell[nk(rows[j].n)]=eVal(rows[j],curP,e); if(rows[j].r) rat=true; });"
    yeni2 = ("const ky=mxInfo().mk(kids[i]); kids[i].forEach((j,q)=>{ if(ky[q]==null) return; cell[nk(ky[q])]=eVal(rows[j],curP,e); "
             "if(rows[j].r) rat=true; });")
    if eski1 in yeni_html and eski2 in yeni_html:
        yeni_html = yeni_html.replace(eski1, yeni1, 1).replace(eski2, yeni2, 1)
    elif "SRC.mx={kids,isM,cols,mk," not in yeni_html:      # ana sitede zaten varsa sessiz geç
        print("   ! tablo görünümü kodu bulunamadı (tekrarlanan sütunlar)")
    n_banka = len({b[2] for b, _, _ in raporlar})
    afis = (f'<div style="position:sticky;top:0;z-index:9999;background:#7a4b00;color:#fff;padding:8px 16px;'
            f'font:500 14px/1.4 Roboto,system-ui,sans-serif;text-align:center">'
            f'ÖNİZLEME — {AY_UZUN[ay]} {yil} dönemi KAP ön verisidir: {n_banka} bankanın {len(raporlar)} raporu, '
            f'banka PDF\'lerinden okundu; TBB bu dönemi henüz yayımlamadı. Öteki bankalar ve raporda bulunamayan '
            f'kalemler bu dönem için boştur; sektör ve grup toplamları {AY_UZUN[ay]} {yil} için yalnız bu bankaları '
            f'kapsar. <a href="../" style="color:#ffd58a">Ana rapor</a></div>')
    yeni_html = re.sub(r"(<body[^>]*>)", lambda m: m.group(1) + afis, yeni_html, count=1)
    yeni_html = re.sub(r"<title>(.*?)</title>", lambda m: f"<title>Önizleme {AY_KISA[ay]} {yil} – {m.group(1)}</title>",
                       yeni_html, count=1)
    (cikti / "index.html").write_text(yeni_html, encoding="utf-8")
    for did, (_, ham) in degisen_dn.items():
        (cikti / "dipnot" / f"{did}.js").write_text(
            f'__DN("{did}",' + json.dumps(ham, ensure_ascii=False, separators=(",", ":")) + ");", encoding="utf-8")
    print(f"→ {cikti / 'index.html'} ve {len(degisen_dn)} dipnot dosyası")
    for (tur, ne), n in sorted(ozet.items()):
        print(f"   {tur:<5} {ne:<11} {n:>6} değer")


if __name__ == "__main__":
    main()
