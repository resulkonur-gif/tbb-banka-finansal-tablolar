#!/usr/bin/env python3
"""Siteye "Yatırımcı özeti" bölümü ekler (karne, çeyreklik kâr/marj, aktif kalitesi, sermaye tamponu, banka notları).

Kullanım:
    python yatirimci_ozeti.py --site docs/onizleme --onizleme 2026-09
    python yatirimci_ozeti.py --site docs

Betik idempotenttir: daha önce eklenen bölüm işaretlerinden bulunur ve yenisiyle değiştirilir.
Aktif kalitesi göstergeleri dipnotlardan (2.f.10 iv, 2.f.2.i, 2.f.9) hesaplanıp sayfaya gömülür;
dipnot dosyaları önce --site/dipnot, yoksa --site/../dipnot klasöründe aranır.
"""
import argparse
import json
import re
from pathlib import Path

# Yıllık TÜFE (%), çeyrek sonları. Kaynak: TCMB EVDS TP.FG.J0 (yıllık % değişim), 2026 için TÜİK duyuruları.
# makro_veri.json (makro_pano.py) varsa oradaki TP.TUKFIY2025.GENEL değerleri bunların üstüne yazılır.
TUFE = {
    "2018-03": 10.23, "2018-06": 15.39, "2018-09": 24.52, "2018-12": 20.30,
    "2019-03": 19.71, "2019-06": 15.72, "2019-09": 9.26, "2019-12": 11.84,
    "2020-03": 11.86, "2020-06": 12.62, "2020-09": 11.75, "2020-12": 14.60,
    "2021-03": 16.19, "2021-06": 17.53, "2021-09": 19.58, "2021-12": 36.08,
    "2022-03": 61.14, "2022-06": 78.62, "2022-09": 83.45, "2022-12": 64.27,
    "2023-03": 50.51, "2023-06": 38.21, "2023-09": 61.53, "2023-12": 64.77,
    "2024-03": 68.50, "2024-06": 71.60, "2024-09": 49.38, "2024-12": 44.38,
    "2025-03": 38.10, "2025-06": 35.05, "2025-09": 33.29, "2025-12": 30.89,
    "2026-03": 30.87, "2026-06": 32.11, "2026-09": 29.73,
}

# dipnot adı → solo kodları (konsolide pakette aynı sıradaki kalemler kullanılır)
DN_NPL = ("2. f.10 iv) Donuk alacakların kullanıcı grup.göre brüt ve net tut.göst.",
          [106309, 106317, 106320, 106327, 106316, 106330, 106365, 106367, 106369])
DN_S2 = ("2.f.2.i) Standart nitelikli ve yakın izle.kre. ile yeniden yapılan.yakın izl.kre..", [113556, 113582])
DN_S3K = ("2. f.9) Ayrılan özel karşılıklar veya temerrüt (3.aşama) karşılıkları", [113713])

BAS_HTML, SON_HTML = "<!--YOZ-->", "<!--/YOZ-->"
BAS_JS, SON_JS = "/*YOZ*/", "/*/YOZ*/"


def veri_ayir(html):
    i = html.index("const DATA = ") + len("const DATA = ")
    D, son = json.JSONDecoder().raw_decode(html, i)
    return D


EK_DIPNOT: list[Path] = []


def dipnot_oku(site: Path, id_):
    for k in [site / "dipnot" / f"{id_}.js", site.parent / "dipnot" / f"{id_}.js"] + [e / f"{id_}.js" for e in EK_DIPNOT]:
        if k.exists():
            t = k.read_text(encoding="utf-8")
            m = re.match(r'\s*__DN\("([^"]+)",(.*)\);\s*$', t, re.S)
            return json.loads(m.group(2))
    return None


def aktif_kalitesi(D, site: Path):
    solo_sira = {}
    for ad, _ in (DN_NPL, DN_S2, DN_S3K):
        dn = next((x for x in D["solo"].get("DN", []) if x["ad"] == ad), None)
        pk = dipnot_oku(site, dn["id"]) if dn else None
        solo_sira[ad] = {it[0]: k for k, it in enumerate(pk["items"])} if pk else {}
    out = {}
    for tur in ("solo", "kons"):
        S = D[tur]
        T7 = S["T7"]
        paket = {}
        for ad, kodlar in (DN_NPL, DN_S2, DN_S3K):
            dn = next((x for x in S.get("DN", []) if x["ad"] == ad), None)
            pk = dipnot_oku(site, dn["id"]) if dn else None
            if not pk:
                continue
            sira = solo_sira[ad]
            kk = [str(pk["items"][sira[c]][0]) for c in kodlar if c in sira and sira[c] < len(pk["items"])]
            paket[ad] = (pk, kk)

        def al(ad, p, kod_sira, i):
            if ad not in paket:
                return None
            pk, kk = paket[ad]
            if kod_sira >= len(kk):
                return None
            f = pk["v"].get(p, {}).get(kk[kod_sira]) or []
            d = dict(zip(f[::2], f[1::2]))
            return d.get(i)

        tv = {}
        for p in S["periods"]:
            r = {"npl": [], "s1": [], "s2": [], "s3k": []}
            for i in T7:
                ps = [al(DN_NPL[0], p, k, i) for k in range(len(DN_NPL[1]))]
                r["npl"].append(sum(abs(x) for x in ps if x) if any(x for x in ps) else None)
                s1, s2 = al(DN_S2[0], p, 0, i), al(DN_S2[0], p, 1, i)
                r["s1"].append(s1 if s1 else None)
                r["s2"].append(s2 if s2 is not None and s1 else None)
                k3 = al(DN_S3K[0], p, 0, i)
                r["s3k"].append(abs(k3) if k3 else None)
            if any(x is not None for x in r["npl"] + r["s1"]):
                tv[p] = r
        out[tur] = tv
    return out


# ----------------------------------------------------------------------------- banka göstergeleri tablosu
TB_SIRA = ["VakıfBank", "Ziraat", "İş Bankası", "Garanti BBVA", "Yapı Kredi", "Akbank", "Halkbank"]   # QNB şimdilik hariç
TB_DN = {   # alan → (dipnot adı, solo kodları; değerler toplanır, mutlak)
    "s1": ("2.f.2.i) Standart nitelikli ve yakın izle.kre. ile yeniden yapılan.yakın izl.kre..", [113556]),
    "s2": ("2.f.2.i) Standart nitelikli ve yakın izle.kre. ile yeniden yapılan.yakın izl.kre..", [113582]),
    "k1": ("2.f.2.i) Standart Nitelikli - Yakın İzlemedeki Krediler", [16598, 16632]),
    "k2": ("2.f.2.i) Standart Nitelikli - Yakın İzlemedeki Krediler", [16614, 16634]),
    "s3": DN_NPL,
    "k3": ("2. f.9) Ayrılan özel karşılıklar veya temerrüt (3.aşama) karşılıkları", [113713]),
    "wo": ("2. f.10 ii) Toplam donuk alacak hareketlerine ilişkin bilgiler", [16823, 113750, 16884, 113759, 16885, 113768]),
    "reklam": ("2.f) Diğer faaliyet giderlerine ilişkin bilgiler", [114893]),
}


def serbest_karsilik(pdf_klasoru, donem):
    """KAP raporlarındaki 'Muhtemel riskler için ayrılan serbest karşılıklara ilişkin bilgiler' cümlesi:
    {(kısa ad, tür): {dönem: tutar (bin TL)}} — cari dönem ve parantezdeki önceki yıl sonu."""
    out = {}
    if not pdf_klasoru:
        return out
    try:
        import banka_karsilastir as bk
        import pymupdf
    except Exception:                                 # noqa: BLE001
        return out
    y, a = int(donem[:4]), int(donem[5:])
    onc = f"{y - 1}-12"
    sayi = r"(\d{1,3}(?:[.,]\d{3})*(?:[.,]\d+)?)"
    for b in bk.BANKALAR:
        k = Path(pdf_klasoru) / bk.dosya_adi(b[2])
        if not k.exists():
            continue
        for pdf in sorted(k.glob("*.pdf")):
            tur = bk.pdf_turu(pdf)
            if not tur or (b[2], tur) in out:
                continue
            try:
                doc = pymupdf.open(pdf)
                bas = " ".join(doc[i].get_text() for i in range(min(6, len(doc))))
                carpan = 1000 if re.search(r"Milyon\s+(TL|Türk\s+Lirası)", bas, re.I) else 1
                for pg in doc:
                    t = re.sub(r"\s+", " ", pg.get_text())
                    m = re.search(r"serbest karşılıklara ilişkin bilgiler\W*(.{0,260})", t, re.I)
                    if not m:
                        continue
                    cumle = m.group(1)
                    par = re.search(r"\(([^)]*?\d{4})\s*:\s*([^)]*)\)", cumle)
                    ana = cumle[:par.start()] if par else cumle
                    d = {}
                    if re.search(r"bulunmamaktadır|yoktur", ana, re.I):
                        d[donem] = 0
                    else:
                        mm = re.search(sayi + r"\s*(?:Milyon\s*)?TL", ana)
                        if mm:
                            d[donem] = float(mm.group(1).replace(".", "").replace(",", "")) * carpan
                    if par:
                        o = par.group(2)
                        if re.search(r"bulunmamaktadır|yoktur", o, re.I):
                            d[onc] = 0
                        else:
                            mm = re.search(sayi, o)
                            if mm:
                                d[onc] = float(mm.group(1).replace(".", "").replace(",", "")) * carpan
                    if d:
                        out[(b[2], tur)] = d
                    break
                doc.close()
            except Exception:                         # noqa: BLE001
                continue
    return out


def ek_veri_serbest(yol) -> dict:
    """pdf_ek_veri deposu (ek_veri.json, elle düzeltmeler dahil 'geçerli' değer) → {(banka, tür): {dönem: bin TL}}"""
    out: dict = {}
    if not yol or not Path(yol).exists():
        return out
    sat = json.loads(Path(yol).read_text(encoding="utf-8")).get("satirlar", [])
    for kod in ("SK_ONCEKI_YS", "SK_BAKIYE"):            # dönem sonu bakiyesi önceki yıl sonu bilgisini ezer
        for r in sat:
            if r.get("kod") != kod or r.get("gecerli") in (None, ""):
                continue
            d = r["donem"] if kod == "SK_BAKIYE" else f"{int(r['donem'][:4]) - 1}-12"
            out.setdefault((r["banka"], r["kapsam"]), {})[d] = float(r["gecerli"]) * 1000
    return out


def banka_tablosu(D, site: Path, pdf_klasoru=None, onizleme=None, ek_veri=None):
    solo_sira = {}
    for ad, _ in set((v[0], None) for v in TB_DN.values()):
        dn = next((x for x in D["solo"].get("DN", []) if x["ad"] == ad), None)
        pk = dipnot_oku(site, dn["id"]) if dn else None
        solo_sira[ad] = {it[0]: k for k, it in enumerate(pk["items"])} if pk else {}
    serbest = serbest_karsilik(pdf_klasoru, onizleme) if onizleme else {}
    for k, v in ek_veri_serbest(ek_veri).items():       # depo (elle düzeltilmiş) değerleri esastır
        serbest.setdefault(k, {}).update(v)
    out = {}
    for tur in ("solo", "kons"):
        S = D[tur]
        A = S["A"]
        bankalar = [A["short"].index(n) for n in TB_SIRA if n in A["short"]]
        paket = {}
        for alan, (ad, kodlar) in TB_DN.items():
            dn = next((x for x in S.get("DN", []) if x["ad"] == ad), None)
            pk = dipnot_oku(site, dn["id"]) if dn else None
            if not pk:
                continue
            sira = solo_sira.get(ad, {})
            kk = [str(pk["items"][sira[c]][0]) for c in kodlar if c in sira and sira[c] < len(pk["items"])]
            paket[alan] = (pk, kk)

        def al(alan, p, i):
            if alan not in paket:
                return None
            pk, kk = paket[alan]
            xs = []
            for k in kk:
                f = pk["v"].get(p, {}).get(k) or []
                x = dict(zip(f[::2], f[1::2])).get(i)
                if x is not None:
                    xs.append(abs(x))
            return sum(xs) if xs else None

        yil_wo = {}                                   # (yıl, banka) → yıl sonu kayıttan düşülen + satılan
        for p in S["periods"]:
            if p.endswith("-12"):
                for i in bankalar:
                    yil_wo[(int(p[:4]), i)] = al("wo", p, i)
        dv = {}
        for p in S["periods"]:
            r = {k: [] for k in list(TB_DN) + ["wocum", "serbest"]}
            for i in bankalar:
                for alan in TB_DN:
                    r[alan].append(al(alan, p, i))
                y = int(p[:4])
                onceki = [yil_wo.get((yy, i)) for yy in range(2018, y)]
                r["wocum"].append(sum(x for x in onceki if x) if any(x is not None for x in onceki) else None)
                kisa = A["short"][i].replace(" BBVA", "").replace(" Bankası", "")
                sb = next((v for (bn, tt), v in serbest.items() if tt == tur and
                           (bn == A["short"][i] or bn.startswith(kisa) or A["short"][i].startswith(bn))), {})
                r["serbest"].append(sb.get(p))
            if any(x is not None for x in r["s1"] + r["s3"]):
                dv[p] = r
        out[tur] = {"bankalar": bankalar, "v": dv}
    return out


HTML_AK = r"""<!--YOZAK-->
  <section aria-labelledby="h-yozak" id="yozAkSec">
    <div class="sec-head">
      <h2 id="h-yozak">Aktif kalitesi: aşamalar ve karşılıklar</h2>
      <p class="insight">Banka banka canlı krediler, aşamalar, karşılıklar ve kayıttan düşülen alacaklar. Tutarlar bin TL; seçili dönem ve solo/konsolide görünüme göre.</p>
    </div>
    <div class="panel" style="min-width:0">
      <div style="display:flex;justify-content:space-between;align-items:baseline;gap:12px;flex-wrap:wrap"><h3 style="margin:0"><span id="yozAkP"></span></h3><button type="button" class="yoz-btn" id="yozAkCsv">CSV indir</button></div>
      <div style="overflow-x:auto"><table class="ratio-table yoz-tb" id="yozAk"></table></div>
      <p class="muted" style="font-size:12.5px">Canlı krediler = 1. aşama (standart nitelikli) + 2. aşama (yakın izlemedeki krediler), 2.f.2.i dipnotu; aşama karşılıkları aynı dipnotun karşılık tablosundan. 3. aşama = III–V. grup brüt tutarlar (2.f.10 iv), karşılığı 2.f.9. Kayıttan düşülen ve satılan donuk alacaklar 2.f.10 ii, yıl başından birikimli. Geçmiş yıllar tutarı TBB verisinden, 2018'den bu yana yıl sonları toplanarak hesaplandı; bankaların kendi açıkladığı birikimli tutardan farklı olabilir.</p>
    </div>
  </section>
<!--/YOZAK-->"""

HTML_GD = r"""<!--YOZGD-->
  <section aria-labelledby="h-yozgd" id="yozGdSec">
    <div class="sec-head">
      <h2 id="h-yozgd">Karşılık ve gider kalemleri</h2>
      <p class="insight">Banka banka serbest karşılıklar, faaliyet giderleri ve reklam harcamaları. Tutarlar bin TL; giderler yıl başından birikimli.</p>
    </div>
    <div class="panel" style="min-width:0">
      <div style="display:flex;justify-content:space-between;align-items:baseline;gap:12px;flex-wrap:wrap"><h3 style="margin:0"><span id="yozGdP"></span></h3><button type="button" class="yoz-btn" id="yozGdCsv">CSV indir</button></div>
      <div style="overflow-x:auto"><table class="ratio-table yoz-tb" id="yozGd"></table></div>
      <p class="muted" style="font-size:12.5px">Serbest karşılıklar TBB verisinde yoktur; Aralık 2025 döneminden itibaren banka raporlarındaki "muhtemel riskler için ayrılan serbest karşılıklar" açıklamasından (cari dönem ve önceki yıl sonu) okunur. OPEX = personel giderleri + diğer faaliyet giderleri. Reklam ve ilan giderleri 2.f dipnotundan.</p>
    </div>
  </section>
<!--/YOZGD-->"""


HTML = r"""<!--YOZ-->
  <section aria-labelledby="h-yoz" id="yozSec">
    <div class="sec-head">
      <h2 id="h-yoz">Çeyreğe ilişkin notlar</h2>
      <p class="insight" id="yozInsight"></p>
    </div>
    <p class="yoz-src" id="yozSrc" hidden></p>
    <div class="panel" style="min-width:0">
      <h3>Karne · <span id="yozKarneP"></span></h3>
      <div style="overflow-x:auto"><table class="ratio-table yoz-karne" id="yozKarne"></table></div>
      <p class="muted" style="font-size:12.5px">Yeşil: yedi banka içinde en iyi üç; kırmızı: en zayıf üç (yönü sütun başlığındaki oka göre). Başlığa tıklayınca o sütuna göre sıralanır.</p>
    </div>
    <div class="panel-grid" style="grid-template-columns:repeat(auto-fit,minmax(min(100%,460px),1fr));margin-top:24px">
      <div class="panel" style="min-width:0"><div class="controls"><label class="field" for="yozMetric">Gösterge<select id="yozMetric"></select></label></div><div class="chartbox tall"><canvas id="yozChart" role="img" aria-label="Seçilen göstergenin çeyreklere göre seyri"></canvas></div><p class="muted" id="yozChartNote" style="font-size:12.5px"></p></div>
      <div class="panel" style="min-width:0"><h3>Çekirdek sermaye tamponu</h3><div class="chartbox tall"><canvas id="yozCapChart" role="img" aria-label="Çekirdek sermaye oranının asgari orana uzaklığı"></canvas></div><p class="muted" style="font-size:12.5px">Çekirdek sermaye oranı eksi asgari oran (%4,5 + bankaya özgü tampon gereksinimi: koruma, döngüsel ve sistemik önem tamponları). Puan cinsinden; düşük tampon kredi büyümesini sınırlar.</p></div>
    </div>
    <h3 style="margin:24px 0 12px">Banka notları</h3>
    <div class="yoz-notes" id="yozNotes"></div>
    <details class="yoz-yontem"><summary>Yöntem ve kısıtlar</summary>
      <ul>
        <li>ROE: yıl başından birikimli net kâr yıllıklaştırılarak, önceki yıl sonu ile dönem sonu özkaynağının ortalamasına bölündü (bankaların ROAE yöntemi). Reel ROE = (1 + ROE) / (1 + yıllık TÜFE) − 1.</li>
        <li>Çeyreklik değerler birikimli gelir tablosunun ardışık çeyrek farkıdır; marjlar çeyreklik tutar × 4 / ortalama aktif.</li>
        <li>Swap düzeltmeli marj: net faiz gelirine türev finansal işlemler ve kambiyo kâr/zararı eklendi. Türk bankalarında TL fonlama için yapılan swapların maliyeti faiz gideri yerine türev zararında izlenir; bu düzeltme bankalar arası karşılaştırmayı adilleştirir, ama alım satım amaçlı türev sonuçlarını da içerir.</li>
        <li>Risk maliyeti: beklenen zarar karşılığı gideri (yıllıklaştırılmış) / ortalama krediler. Takipteki kredi oranı: 3. aşama (III–V. grup) brüt krediler / krediler (2.f.10 iv dipnotu). 2. aşama payı: yakın izlemedeki krediler / (standart + yakın izleme) (2.f.2.i). Karşılık oranı: 3. aşama karşılıkları / 3. aşama krediler (2.f.9).</li>
        <li>Ücret / gider: net ücret-komisyon gelirinin personel ve diğer faaliyet giderleri toplamına oranı. Gider / gelir: aynı giderlerin faaliyet brüt kârına oranı.</li>
        <li>Tutarlar nominaldir; bankalar TMS 29 enflasyon muhasebesi uygulamamaktadır. Piyasa değeri, F/DD, F/K gibi değerleme verileri bu sayfada yoktur. Banka notları verilerden kurallı olarak üretilir; yatırım tavsiyesi değildir.</li>
      </ul>
    </details>
  </section>
<!--/YOZ-->"""

CSS = r"""<style id="yozCss">
.yoz-src{margin:-8px 0 16px;padding:10px 14px;border-radius:8px;background:var(--warm-soft);color:var(--ink);font-size:13.5px}
.yoz-karne thead th button{all:unset;cursor:pointer}
.yoz-karne thead th{vertical-align:bottom;white-space:normal;min-width:64px;line-height:1.25}
.yoz-karne td.g{background:color-mix(in srgb,var(--good) 16%,transparent)}
.yoz-karne td.b{background:color-mix(in srgb,var(--bad) 14%,transparent)}
.yoz-karne tr.tot td{font-weight:700;border-top:2px solid var(--line)}
.yoz-karne td:first-child{font-weight:600}
.yoz-karne .sw{display:inline-block;width:10px;height:10px;border-radius:2px;margin-right:6px;vertical-align:0}
.yoz-notes{display:grid;grid-template-columns:repeat(auto-fill,minmax(min(100%,320px),1fr));gap:16px}
.yoz-note{border:1px solid var(--line);border-radius:12px;padding:14px 16px;background:var(--surface);font-size:13.5px;line-height:1.5}
.yoz-note h4{margin:0 0 6px;font-size:15px;display:flex;align-items:center;gap:8px}
.yoz-note .tag{font-size:11.5px;font-weight:700;border-radius:6px;padding:1px 7px}
.yoz-note .tag.g{color:var(--good);background:color-mix(in srgb,var(--good) 14%,transparent)}
.yoz-note .tag.b{color:var(--bad);background:color-mix(in srgb,var(--bad) 12%,transparent)}
.yoz-note p{margin:4px 0}
.yoz-tb td,.yoz-tb th{font-variant-numeric:tabular-nums}
.yoz-tb th{text-align:right}.yoz-tb th:first-child{text-align:left}
.yoz-tb tr.k td{font-weight:700}
.yoz-tb tr.alt td:first-child{padding-left:22px;color:var(--ink-2)}
.yoz-tb tr.bas td{background:var(--surface-2);font-weight:800;color:var(--bad)}
.yoz-btn{border:1px solid var(--line);background:var(--surface);color:var(--ink);border-radius:8px;padding:6px 12px;font-size:13px;cursor:pointer}
.yoz-yontem{margin-top:20px;font-size:13px;color:var(--ink-2)}
.yoz-yontem summary{cursor:pointer;font-weight:600}
.yoz-yontem li{margin:4px 0}
</style>"""

JS = r"""/*YOZ*/
const YZ = ((window.YOZ||{})[MODE])||{};
const YTUFE = (window.YOZ_TUFE||{});
const YONIZ = new Set(window.YOZ_ONIZ||[]);
const yozState = {sort:"roe", dir:-1, met:"qnet"};
try{ Object.assign(yozState, JSON.parse(localStorage.getItem("yoz.state")||"{}")); }catch(e){}
const yzX = b => idxOf(b);
const yzQ = (p,c,x) => { const q=prevP(p); if(MON[p]===3 || !q || q.slice(0,4)!==p.slice(0,4)) return MON[p]===3? asum(p,c,x) : null; return asum(p,c,x)-asum(q,c,x); };
const yzAvg = (p,c,x) => { const q=prevP(p); return q? (asum(p,c,x)+asum(q,c,x))/2 : asum(p,c,x); };
const yzDQ = (p,cs,x) => { let s=0; for(const c of cs){ const v=yzQ(p,c,x); if(v==null) return null; s+=v; } return s; };
const yzSum = (p,k,x) => { const r=YZ[p]; if(!r||!r[k]) return null; let s=0,n=0; x.forEach(i=>{ const j=T7.indexOf(i), v=j>=0? r[k][j] : null; if(v!=null){ s+=v; n++; } }); return n===x.length? s : null; };
const yzBuf = (p,i) => { let q=p; for(let k=0;k<3&&q;k++){ const v=A.buf&&A.buf[q]? A.buf[q][i] : null; if(v!=null) return v; q=prevP(q); } return null; };
const YM = {
  roe:{l:"ROE", u:"%", up:1, f:(p,b)=>RATIO.roe.f(p,b)},
  rroe:{l:"Reel ROE", u:"%", up:1, f:(p,b)=>{ const r=RATIO.roe.f(p,b), t=YTUFE[p]; return r==null||t==null? null : (1+r)/(1+t/100)-1; }},
  qnet:{l:"Çeyrek net kâr", u:"mlr", up:1, chartOnly:0, f:(p,b)=>{ const v=yzQ(p,102088,yzX(b)); return v==null? null : v; }},
  qnetd:{l:"Çeyrek kâr değişimi", u:"%", up:1, f:(p,b)=>{ const q=prevP(p); if(!q) return null; const a=YM.qnet.f(p,b), c=YM.qnet.f(q,b); return a==null||c==null||c<=0? null : a/c-1; }},
  qnim:{l:"Net faiz marjı (çeyrek)", u:"%", d:2, up:1, f:(p,b)=>{ const x=yzX(b), v=yzQ(p,102213,x); return v==null? null : div(v*4, yzAvg(p,16018,x)); }},
  qsnim:{l:"Swap düzeltmeli marj (çeyrek)", u:"%", d:2, up:1, f:(p,b)=>{ const x=yzX(b), v=yzDQ(p,[102213,107311,102279],x); return v==null? null : div(v*4, yzAvg(p,16018,x)); }},
  cor:{l:"Risk maliyeti", u:"%", d:2, up:0, f:(p,b)=>{ const x=yzX(b); return div(-asum(p,16034,x)*12/MON[p], avgYB(p,16005,x)); }},
  qcor:{l:"Risk maliyeti (çeyrek)", u:"%", d:2, up:0, f:(p,b)=>{ const x=yzX(b), v=yzQ(p,16034,x); return v==null? null : div(-v*4, yzAvg(p,16005,x)); }},
  npl:{l:"Takipteki kredi oranı", u:"%", d:2, up:0, f:(p,b)=>{ const n=yzSum(p,"npl",yzX(b)); return n==null? null : div(n, asum(p,16005,yzX(b))); }},
  s2:{l:"2. aşama payı", u:"%", up:0, f:(p,b)=>{ const x=yzX(b), a=yzSum(p,"s2",x), s=yzSum(p,"s1",x); return a==null||s==null? null : div(a, a+s); }},
  cov:{l:"Takipteki karşılık oranı", u:"%", up:1, f:(p,b)=>{ const x=yzX(b), k=yzSum(p,"s3k",x), n=yzSum(p,"npl",x); return k==null||n==null? null : div(k,n); }},
  cet1:{l:"Çekirdek sermaye oranı", u:"%", d:2, up:1, f:(p,b)=>RATIO.cet1? RATIO.cet1.f(p,b) : null},
  tampon:{l:"Çekirdek sermaye tamponu", u:"pp", up:1, f:(p,b)=>{ if(b===NB) return null; const i=T7[b], c=A.cet1&&A.cet1[p]? A.cet1[p][i] : null, f=yzBuf(p,i); return c==null||f==null? null : c-4.5-f; }},
  km:{l:"Kredi / mevduat", u:"%", up:null, f:(p,b)=>RATIO.km.f(p,b)},
  kbuy:{l:"Kredi büyümesi (yıllık)", u:"%", up:null, f:(p,b)=>{ const q=yoyP(p); return q? div(asum(p,16005,yzX(b)), asum(q,16005,yzX(b)))-1 : null; }},
  ucgid:{l:"Ücret / gider", u:"%", up:1, f:(p,b)=>{ const x=yzX(b); return div(asum(p,102215,x), absSum(p,16033,x)+absSum(p,102221,x)); }},
  gidgel:{l:"Gider / gelir", u:"%", up:0, f:(p,b)=>{ const x=yzX(b); return div(absSum(p,16033,x)+absSum(p,102221,x), asum(p,102360,x)); }}
};
const YK_COLS = ["roe","rroe","qnet","qsnim","cor","npl","s2","cov","cet1","tampon","km","kbuy","ucgid","gidgel"];
const yzFmt = (k,x) => { const m=YM[k]; if(x==null||!isFinite(x)) return '<span class="muted">–</span>'; if(m.u==="mlr") return nf1.format(mlr(x)); if(m.u==="pp") return (x<0?"−":"")+nf1.format(Math.abs(x)); return pct(x, m.d===2?2:1); };
const yzRank = (k,p) => { const m=YM[k], vals=[...Array(NB).keys()].map(b=>m.f(p,b)); const ok=vals.map((v,b)=>[v,b]).filter(([v])=>v!=null&&isFinite(v)); ok.sort((a,b)=>m.up? b[0]-a[0] : a[0]-b[0]); const r={}; ok.forEach(([,b],k2)=>r[b]=k2); return {r, n:ok.length, vals}; };
function renderYoz(){
  const sec=document.getElementById("yozSec"); if(!sec) return;
  const p=curP;
  const src=document.getElementById("yozSrc");
  src.hidden=!YONIZ.has(p);
  src.innerHTML=`<b>${PL[p]} ön veridir:</b> bankaların KAP’ta yayımladığı finansal raporlardan okundu; TBB verisi yayımlanınca güncellenecek. Küçük farklar olabilir.`;
  document.getElementById("yozKarneP").textContent=`${PL[p]}, ${MODE==="kons"?"konsolide":"solo"}`;
  // karne
  const ranks={}; YK_COLS.forEach(k=>ranks[k]=yzRank(k,p));
  const order=[...Array(NB).keys()]; const sk=yozState.sort;
  order.sort((a,b)=>{ const x=YM[sk].f(p,a), y=YM[sk].f(p,b); if(x==null) return 1; if(y==null) return -1; return (x-y)*yozState.dir; });
  const head=`<thead><tr><th>Banka</th>${YK_COLS.map(k=>{ const m=YM[k]; const ar=m.up===1?" ↑":m.up===0?" ↓":""; return `<th scope="col" title="${m.l}"><button type="button" data-k="${k}">${m.l}${m.u==="mlr"?" (mlr TL)":m.u==="pp"?" (puan)":""}${ar}${sk===k?(yozState.dir<0?" ▼":" ▲"):""}</button></th>`; }).join("")}</tr></thead>`;
  const row=(b,cls)=>`<tr class="${cls||''}"><td><span class="sw" style="background:${bcol(b)}"></span>${b===NB?"İlk 7":esc(B[b])}</td>${YK_COLS.map(k=>{ const m=YM[k], x=m.f(p,b); let c=""; if(b<NB && m.up!=null && ranks[k].n>=5 && ranks[k].r[b]!=null){ const rk=ranks[k].r[b]; c= rk<3? "g" : rk>=ranks[k].n-3? "b" : ""; } return `<td class="${c}">${yzFmt(k,x)}</td>`; }).join("")}</tr>`;
  const t=document.getElementById("yozKarne");
  t.innerHTML=head+`<tbody>${order.map(b=>row(b)).join("")}${row(NB,"tot")}</tbody>`;
  t.querySelectorAll("thead button").forEach(bt=>bt.onclick=()=>{ const k=bt.dataset.k; if(yozState.sort===k) yozState.dir*=-1; else { yozState.sort=k; yozState.dir=YM[k].up===0?1:-1; } try{localStorage.setItem("yoz.state",JSON.stringify(yozState))}catch(e){} renderYoz(); });
  // içgörü
  const tot=k=>YM[k].f(p,NB), q=prevP(p);
  const best=k=>{ const r=ranks[k]; const b=Object.keys(r.r).find(x=>r.r[x]===0); return b!=null? B[+b] : null; };
  const worst=k=>{ const r=ranks[k]; const b=Object.keys(r.r).find(x=>r.r[x]===r.n-1); return b!=null? B[+b] : null; };
  const bp=(a,c)=> a==null||c==null? null : Math.round((a-c)*1e4);
  const sgn=x=> x==null? "" : (x>=0?"+":"−")+Math.abs(x);
  let s=[];
  const roe=tot("roe"), rr=tot("rroe");
  if(roe!=null) s.push(`Yedi bankanın yıllık ROE’si <b>${pct(roe)}</b>${YTUFE[p]!=null?`; ${PL[p]} yıllık enflasyonu %${nf1.format(YTUFE[p])} olduğundan reel ROE <b class="${rr>=0?'up':'down'}">${pct(rr)}</b>`:""}. En kârlı ${best("roe")}, en düşük ${worst("roe")}.`);
  const m1=tot("qsnim"), m0=q? YM.qsnim.f(q,NB) : null;
  if(m1!=null) s.push(`Swap düzeltmeli çeyreklik marj <b>${pct(m1,2)}</b>${m0!=null?` (önceki çeyrek ${pct(m0,2)}, ${sgn(bp(m1,m0))} baz puan)`:""}.`);
  const n1=tot("npl"), n0=q? YM.npl.f(q,NB) : null;
  if(n1!=null) s.push(`Takipteki kredi oranı <b>${pct(n1,2)}</b>${n0!=null?` (${sgn(bp(n1,n0))} baz puan)`:""}; risk maliyeti ${pct(tot("cor"),2)}.`);
  const tb=ranks.tampon; const wb=Object.keys(tb.r).find(x=>tb.r[x]===tb.n-1);
  if(wb!=null) s.push(`Sermaye tamponu en ince banka ${B[+wb]} (${nf1.format(YM.tampon.f(p,+wb))} puan).`);
  document.getElementById("yozInsight").innerHTML=s.join(" ");
  renderYozTablo();
  // grafik
  renderYozChart();
  // sermaye tamponu
  const tq=[q,p].filter(Boolean), tbs=[...Array(NB).keys()];
  mkChart("yozCapChart",{type:"bar",data:{labels:tbs.map(b=>B[b]),datasets:tq.map((pp,k)=>({label:PL[pp],data:tbs.map(b=>YM.tampon.f(pp,b)),backgroundColor:k===tq.length-1? tbs.map(b=>bcol(b)) : css("--line"),borderRadius:3}))},
    options:baseOpts(x=>nf1.format(x)+" p",{plugins:{legend:{display:true,labels:{color:css("--ink-2"),boxWidth:10,boxHeight:10,font:{family:"Roboto",size:11}}},tooltip:{callbacks:{label:c=>`${c.dataset.label}: ${c.parsed.y==null?'–':nf2.format(c.parsed.y)+' puan'}`}}}})});
  // notlar
  const notes=document.getElementById("yozNotes");
  const KEYS=["roe","qsnim","cor","npl","s2","cov","tampon","ucgid","gidgel"];
  notes.innerHTML=[...Array(NB).keys()].map(b=>{
    const g=[], w=[];
    KEYS.forEach(k=>{ const r=ranks[k]; if(r.n<5||r.r[b]==null) return; if(r.r[b]===0) g.push(k); else if(r.r[b]===r.n-1) w.push(k); });
    const lab=k=>`${YM[k].l.replace(" (çeyrek)","")} ${yzFmt(k,YM[k].f(p,b)).replace(/<[^>]+>/g,"")}`;
    const ch=[];
    if(q){
      const a=YM.qnet.f(p,b), c=YM.qnet.f(q,b);
      if(a!=null&&c!=null&&c>0) ch.push(`çeyrek kârı ${nf1.format(mlr(c))} → ${nf1.format(mlr(a))} milyar TL (${chg(a/c-1)})`);
      const m=bp(YM.qsnim.f(p,b),YM.qsnim.f(q,b)); if(m!=null) ch.push(`swap düzeltmeli marj ${sgn(m)} bp`);
      const n=bp(YM.npl.f(p,b),YM.npl.f(q,b)); if(n!=null) ch.push(`takipteki oranı ${sgn(n)} bp`);
      const c1=A.cet1&&A.cet1[p]&&A.cet1[q]? [A.cet1[p][T7[b]],A.cet1[q][T7[b]]] : null; if(c1&&c1[0]!=null&&c1[1]!=null) ch.push(`çekirdek sermaye ${sgn(Math.round((c1[0]-c1[1])*100))} bp`);
    }
    const ip=[];
    const qd=YM.qnetd.f(p,b); if(qd!=null && Math.abs(qd)>=0.25) ip.push(`Çeyrek kârındaki ${qd>0?"sıçramanın":"düşüşün"} kaynağını gelir tablosundan (personel, karşılık, vergi, ticari kâr/zarar) kontrol edin; tek seferlik olabilir.`);
    const sn=YM.qsnim.f(p,b), nm=YM.qnim.f(p,b); if(sn!=null&&nm!=null&&nm-sn>0.015) ip.push(`Swap ve kur zararı marjı ${Math.round((nm-sn)*1e4)} bp aşağı çekiyor.`);
    { const TBm=((window.YOZ_TB||{})[MODE])||null; if(TBm){ const j=TBm.bankalar.indexOf(T7[b]), yb=ybP(p);
      const sb=x=> x&&TBm.v[x]&&TBm.v[x].serbest? TBm.v[x].serbest.at(j) : null, s1=sb(p), s0=sb(yb);
      if(j>=0 && s1!=null && s0!=null && Math.abs(s1-s0)>=1e6){ const net=asum(p,102088,[T7[b]]);
        ip.push(s1<s0 ? `Yıl başından ${nf1.format(mlr(s0-s1))} milyar TL serbest karşılık iptal edildi; dönem kârının yaklaşık %${nf0.format((s0-s1)/Math.max(1,net)*100)}'i bu tek seferlik kalemden (vergi etkisi hariç).`
                      : `Yıl başından ${nf1.format(mlr(s1-s0))} milyar TL serbest karşılık ayrıldı; kârı aşağı çeken, ileride iptal edilebilecek bir tampon.`); } } }
    if(MODE==="solo" && asum(p,102223,[T7[b]])>0.2*Math.max(1,asum(p,102088,[T7[b]]))) ip.push(`Kârın önemli kısmı iştiraklerden (özkaynak yöntemi); konsolide görünüme de bakın.`);
    return `<div class="yoz-note"><h4><span class="sw" style="display:inline-block;width:10px;height:10px;border-radius:2px;background:${bcol(b)}"></span>${esc(B[b])}</h4>
      ${g.length?`<p><span class="tag g">Güçlü</span> ${g.map(lab).join(" · ")}</p>`:""}
      ${w.length?`<p><span class="tag b">Zayıf</span> ${w.map(lab).join(" · ")}</p>`:""}
      ${!g.length&&!w.length?`<p class="muted">Ana göstergelerde yedi bankanın ortasında.</p>`:""}
      ${ch.length?`<p><b>${q?PL[q]+" → "+PL[p]:""}:</b> ${ch.join("; ")}.</p>`:""}
      ${ip.map(x=>`<p class="muted">${x}</p>`).join("")}</div>`;
  }).join("");
}
function renderYozTablo(){
  const T=((window.YOZ_TB||{})[MODE])||null, p=curP;
  const etk=`${PL[p]}, ${MODE==="kons"?"konsolide":"solo"} · bin TL`;
  const BK=T? T.bankalar : [], R=T&&T.v? T.v[p]||null : null;
  const a=(c,j)=>{ const i=BK.at(j), r=A.v[p]&&A.v[p][c]; return r&&r[i]!=null? r[i] : null; };
  const d=(k,j)=> R&&R[k]&&(k!=="wocum"||(R.s1&&R.s1.at(j)!=null))? R[k].at(j) : null;
  const sum=(...xs)=> xs.some(x=>x==null)? null : xs.reduce((s,x)=>s+x,0);
  const q=(x,y)=> x==null||!y? null : x/y;
  const canli=j=>sum(d("s1",j),d("s2",j));
  const AK=[
    ["Canlı krediler",canli,"t","k"],
    ["1. aşama",j=>d("s1",j),"t"],["1. aşama karşılıkları",j=>d("k1",j),"t","alt"],["1. aşama karşılık oranı",j=>q(d("k1",j),d("s1",j)),"p","alt"],
    ["2. aşama",j=>d("s2",j),"t"],["2. aşama karşılıkları",j=>d("k2",j),"t","alt"],["2. aşama karşılık oranı",j=>q(d("k2",j),d("s2",j)),"p","alt"],
    ["3. aşama",j=>d("s3",j),"t"],["3. aşama karşılıkları",j=>d("k3",j),"t","alt"],["3. aşama karşılık oranı",j=>q(d("k3",j),d("s3",j)),"p","alt"],
    ["Kayıttan düşülen ve satılan donuk alacaklar",j=>d("wo",j),"t"],
    ["Geçmiş yıllarda kayıttan çıkarılan ya da satılan donuk alacaklar",j=>d("wocum",j),"t"],
    ["Kayıttan silinen alacaklar dikkate alınarak hesaplanan 3. aşama karşılık oranı",j=>{ const w=sum(d("wo",j)||0,d("wocum",j)||0); return q(sum(d("k3",j),w),sum(d("s3",j),w)); },"p","k"]
  ];
  const GD=[
    ["Serbest karşılıklar",j=>d("serbest",j),"s"],
    ["OPEX",j=>{ const x=a(16033,j), y=a(102221,j); return x==null&&y==null? null : Math.abs(x||0)+Math.abs(y||0); },"t"],
    ["Reklam ve ilan giderleri",j=>d("reklam",j),"t"]
  ];
  const fm=(x,t)=> x==null||!isFinite(x)? "–" : t==="s"&&x===0? "yok" : t==="p"? "%"+nf2.format(x*100) : nf0.format(x);
  const js=[...BK.keys()];
  const ciz=(id,etId,csvId,ROWS,dosya)=>{
    const tb=document.getElementById(id); if(!tb) return;
    document.getElementById(etId).textContent=etk;
    if(!T){ tb.innerHTML=""; return; }
    tb.innerHTML=`<thead><tr><th></th>${js.map(j=>`<th>${esc(A.short[BK.at(j)])}</th>`).join("")}</tr></thead><tbody>${ROWS.map(([l,f,t,c])=>`<tr class="${c||""}"><td>${esc(l)}</td>${js.map(j=>`<td>${fm(f(j),t)}</td>`).join("")}</tr>`).join("")}</tbody>`;
    document.getElementById(csvId).onclick=()=>{
      const sat=[["Gösterge",...js.map(j=>A.short[BK.at(j)])].join(";"), ...ROWS.map(([l,f,t])=>[l,...js.map(j=>{ const x=f(j); return x==null||!isFinite(x)? "" : t==="p"? (x*100).toFixed(2).replace(".",",") : String(Math.round(x)); })].join(";"))];
      const blob=new Blob(["﻿"+sat.join("\n")],{type:"text/csv;charset=utf-8"}), u=URL.createObjectURL(blob), e=document.createElement("a");
      e.href=u; e.download=`${dosya}_${MODE}_${p}.csv`; document.body.appendChild(e); e.click(); e.remove(); setTimeout(()=>URL.revokeObjectURL(u),1000);
    };
  };
  ciz("yozAk","yozAkP","yozAkCsv",AK,"aktif_kalitesi");
  ciz("yozGd","yozGdP","yozGdCsv",GD,"karsilik_gider");
}
function renderYozChart(){
  const k=yozState.met, m=YM[k]; if(!m) return;
  const YP=P.slice(-9);
  const ds=[...Array(NB+1).keys()].filter(b=>!(b===NB&&(k==="qnet"||k==="tampon"))).map(b=>({label:b===NB?"İlk 7":B[b],data:YP.map(pp=>{ const x=m.f(pp,b); return x==null||!isFinite(x)? null : (m.u==="mlr"? mlr(x) : m.u==="pp"? x : x*100); }),borderColor:bcol(b),backgroundColor:bcol(b),borderWidth:b===NB?3:2,borderDash:b===NB?[6,4]:[],tension:.25,spanGaps:true,pointRadius:YP.map(pp=>pp===curP?5:2.5)}));
  const yf=x=> m.u==="mlr"? nf1.format(x) : m.u==="pp"? nf1.format(x)+" p" : "%"+(m.d===2?nf2:nf1).format(x);
  mkChart("yozChart",{type:"line",data:{labels:YP.map(pp=>PL[pp]+(YONIZ.has(pp)?"*":"")),datasets:ds},options:baseOpts(yf,{plugins:{legend:{labels:{color:css("--ink-2"),boxWidth:10,boxHeight:10,font:{family:"Roboto",size:11}}},tooltip:{callbacks:{label:c=>`${c.dataset.label}: ${c.parsed.y==null?'–':yf(c.parsed.y)}`}}}})});
  document.getElementById("yozChartNote").textContent=(m.u==="mlr"?"Milyar TL, çeyreklik (birikimli değil). ":"")+(YP.some(pp=>YONIZ.has(pp))?"* KAP ön verisi. ":"")+"Kesik çizgi ilk 7 toplamı.";
}
(function yozInit(){
  const s=document.getElementById("yozMetric"); if(!s) return;
  ["qnet","roe","rroe","qnim","qsnim","qcor","npl","s2","cov","cet1","tampon","km","kbuy","ucgid","gidgel"].forEach(k=>s.appendChild(el("option",{value:k},YM[k].l)));
  if(!YM[yozState.met]) yozState.met="qnet";
  s.value=yozState.met;
  s.onchange=()=>{ yozState.met=s.value; try{localStorage.setItem("yoz.state",JSON.stringify(yozState))}catch(e){} renderYozChart(); };
})();
/*/YOZ*/
"""


def temizle(t, bas, son):
    while bas in t:
        i = t.index(bas)
        j = t.index(son, i) + len(son)
        t = t[:i] + t[j:]
    return t


def dn_donemleri(D, site: Path):
    """Her dipnot paketi için verisi olan dönemler (herhangi bir bankada en az bir değer)."""
    out = {}
    for tur in ("solo", "kons"):
        out[tur] = {}
        for dn in D[tur].get("DN", []):
            pk = dipnot_oku(site, dn["id"])
            if not pk:
                continue
            out[tur][dn["id"]] = sorted(p for p, v in pk["v"].items() if any(len(f) for f in v.values()))
    return out


DN_YAMALAR = [
    # dipnot listesi: seçili dönemde verisi olmayan dipnot seçilemez
    ("function tabOptions(sel){",
     "function dnSecenek(){ const sel=$(\"nvTab\"); if(!sel) return; const V=((window.DN_VAR||{})[MODE])||null; if(!V) return;\n"
     "  let bozuk=false; sel.querySelectorAll('option[value^=\"dn:\"]').forEach(o=>{ const id=MODE+\"_\"+o.value.slice(3), ps=V[id];\n"
     "    const yok=!!ps && !ps.includes(curP); o.disabled=yok; if(!o.dataset.l) o.dataset.l=o.textContent; o.textContent=o.dataset.l+(yok?` · ${PL[curP]} verisi yok`:\"\");\n"
     "    if(yok && o.selected) bozuk=true; });\n"
     "  if(bozuk){ const ilk=[...sel.options].find(o=>!o.disabled); if(ilk){ sel.value=ilk.value; setTab(ilk.value); } } }\n"
     "function tabOptions(sel){"),
    ("  sel.value=want; NVS[TABKEY]=want;\n}",
     "  sel.value=want; NVS[TABKEY]=want; dnSecenek();\n}"),
    ("const onP=e=>{ if(P.includes(e.detail)){ curP=e.detail; renderTable(); } };",
     "const onP=e=>{ if(P.includes(e.detail)){ curP=e.detail; dnSecenek(); renderTable(); } };"),
    # tablo görünümü: kurum listesinde bu dipnotu bu dönemde yayımlamayanlar seçilemez
    ("function mxEnt(){ if(VIEW===\"tek\") return TEK(); const es=ENTS(); return es.find(e=>e.id===NVS[K(\"mxe\")])||es[0]||ent(\"all\"); }",
     "const mxVar=e=>!SRC||SRC.rows.some(r=>eVal(r,curP,e)!=null);\n"
     "function mxEnt(){ if(VIEW===\"tek\") return TEK(); const es=ENTS(); const c=es.find(e=>e.id===NVS[K(\"mxe\")]); return (c&&mxVar(c))?c:(es.find(mxVar)||c||es[0]||ent(\"all\")); }"),
    ("sel.innerHTML=es.map(x=>`<option value=\"${esc(x.id)}\">${esc(x.l)}</option>`).join(\"\"); sel.value=e.id;",
     "sel.innerHTML=es.map(x=>{ const v=mxVar(x); return `<option value=\"${esc(x.id)}\"${v?\"\":\" disabled\"}>${esc(x.l)}${v?\"\":\" · veri yok\"}</option>`; }).join(\"\"); sel.value=e.id;"),
]


def bolum_tasi(html, adlar, sonra):
    """Rapor şablonunda <section aria-labelledby="..."> bölümlerini (iç içe değil) verilen sırayla
    `sonra` bölümünün hemen ardına taşır. Bölüm bulunamazsa ya da zaten yerindeyse dokunmaz."""
    def aralik(h, ad):
        i = h.find(f'<section aria-labelledby="{ad}"')
        if i < 0:
            return None
        j = h.index("</section>", i) + len("</section>")
        return i, j
    parcalar = []
    for ad in adlar:
        r = aralik(html, ad)
        if not r:
            return html
        parcalar.append(html[r[0]:r[1]])
        html = html[:r[0]] + html[r[1]:].lstrip("\n ")
        html = html[:r[0]] + "  " + html[r[0]:] if r[0] and html[r[0] - 1] == "\n" else html
    r = aralik(html, sonra)
    if not r:
        return html
    ek = "".join("\n\n  " + p for p in parcalar)
    return html[:r[1]] + ek + html[r[1]:]


def uygula(site: Path, onizleme=None, ek_dipnot=None, pdf_klasoru=None, ek_veri=None):
    if ek_dipnot:
        EK_DIPNOT[:] = [Path(ek_dipnot)]
    yol = site / "index.html"
    html = yol.read_text(encoding="utf-8")
    D = veri_ayir(html)
    yoz = aktif_kalitesi(D, site)

    # eski eklemeleri kaldır
    html = temizle(html, BAS_HTML, SON_HTML)
    html = temizle(html, "<!--YOZAK-->", "<!--/YOZAK-->")
    html = temizle(html, "<!--YOZGD-->", "<!--/YOZGD-->")
    html = temizle(html, BAS_JS, SON_JS)
    html = re.sub(r'<style id="yozCss">.*?</style>', "", html, flags=re.S)
    html = re.sub(r'<script id="yozData">.*?</script>\s*', "", html, flags=re.S)
    html = html.replace("renderYoz(); ", "")

    # 1) bölüm: dönem çubuğunun hemen altına
    imza = '<section aria-labelledby="h-sum">'
    assert imza in html, "rapor şablonunda h-sum bölümü bulunamadı"
    html = html.replace(imza, HTML + "\n\n  " + imza, 1)
    # dipnot seçicileri: verisi olmayan dönemde seçilemez
    for eski, yeni in DN_YAMALAR:
        if yeni not in html:
            if html.count(eski) != 1:
                print(f"  ! dipnot yaması bulunamadı: {eski[:60]}")
                continue
            html = html.replace(eski, yeni, 1)
    html = re.sub(r'<script id="dnVar">.*?</script>\s*', "", html, flags=re.S)
    i = html.index("<script>\nconst DATA = ")
    html = html[:i] + "<script id=\"dnVar\">window.DN_VAR=" + json.dumps(dn_donemleri(D, site), separators=(",", ":")) + ";</script>\n" + html[i:]
    # bölüm sırası: bankalar → bilanço yapısı → tüm kalemler → oran haritası → yasal özkaynak → aktif kalitesi
    # → mevduat → nakit yaratma kapasitesi → öteki bölümler → (karşılık ve gider kalemleri) → dipnotlar
    html = bolum_tasi(html, ["h-mix", "h-x", "h-heat", "h-oz", "h-mv", "h-cash"], "h-banks")
    html = bolum_tasi(html, ["h-all", "h-vs"], "h-dn")          # tüm bankalar ve iki banka karşılaştırması en altta
    i = html.find('<section aria-labelledby="h-oz"')
    if i >= 0:
        j = html.index("</section>", i) + len("</section>")
        html = html[:j] + "\n\n" + HTML_AK + html[j:]
    hedef = '<section aria-labelledby="h-dn">'
    if hedef in html:
        html = html.replace(hedef, HTML_GD + "\n\n  " + hedef, 1)
    # 2) CSS
    html = html.replace("</head>", CSS + "\n</head>", 1)
    # 3) veri
    tufe = dict(TUFE)
    try:                                            # makro_pano'nun EVDS verisi varsa çeyrek sonu TÜFE oradan
        import makro_pano
        tufe.update(makro_pano.tufe_ceyrek())
    except Exception:                               # noqa: BLE001
        pass
    tb = banka_tablosu(D, site, pdf_klasoru, onizleme, ek_veri)
    veri = ("<script id=\"yozData\">window.YOZ=" + json.dumps(yoz, ensure_ascii=False, separators=(",", ":")) +
            ";window.YOZ_TB=" + json.dumps(tb, ensure_ascii=False, separators=(",", ":")) +
            ";window.YOZ_TUFE=" + json.dumps(tufe) +
            ";window.YOZ_ONIZ=" + json.dumps([onizleme] if onizleme else []) + ";</script>\n")
    i = html.index("<script>\nconst DATA = ")
    html = html[:i] + veri + html[i:]
    # 4) JS: rapor sayfasının renderAll'undan önce
    imza = "function renderAll(){ renderPeriodBar(); "
    assert html.count(imza) == 1, "rapor renderAll bulunamadı"
    html = html.replace(imza, JS + imza + "renderYoz(); ", 1)
    yol.write_text(html, encoding="utf-8")
    n = {t: len(v) for t, v in yoz.items()}
    print(f"yatırımcı özeti eklendi: {yol} (aktif kalitesi dönemleri: {n})")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--site", required=True, help="index.html'in bulunduğu klasör (docs ya da docs/onizleme)")
    ap.add_argument("--onizleme", help="KAP ön verisi olan dönem, ör. 2026-09")
    ap.add_argument("--dipnot", help="sitede bulunmayan dipnot dosyalarının klasörü (ana sitenin docs/dipnot'u)")
    ap.add_argument("--pdf-klasoru", help="KAP raporları (serbest karşılıklar için; yalnız ön veri döneminde)")
    ap.add_argument("--ek-veri", help="pdf_ek_veri deposu ek_veri.json (2025-12'den itibaren serbest karşılıklar, elle düzeltmeler dahil)")
    a = ap.parse_args()
    uygula(Path(a.site), a.onizleme, a.dipnot, a.pdf_klasoru, a.ek_veri)
