#!/usr/bin/env python3
"""
TCMB verileri sayfası (site/tcmb/index.html) — BDDK sayfalarıyla aynı düzen:
  1. Günlük göstergeler: kurlar (USD, EUR, sepet, EUR/USD), faizler (AOFM, TLREF), borsa (BIST 100, Mali),
     TCMB net dış varlıkları — seçilen güne göre kartlar ve grafik (1 ay / 3 ay / 1 yıl / tümü)
  2. Bugün ne değişti?  günlük · 1 hafta · 1 ay · yılbaşından · 1 yıl (CSV)
  3. Haftalık: brüt rezervler ve yabancı portföy akımları (CSV)
  4. Makro pano (aylık; seçilen günün ayı)
  5. Aylık tablolar: bütün aylık seriler ve değişimleri (CSV)

Veri: tcmb_veri.py arşivi (günlük/haftalık) + makro_veri.json (aylık).
    python tcmb_sayfa.py --site docs/onizleme [--arsiv tcmb_arsiv]
makro_pano.uygula() bunu kendiliğinden çağırır.
"""
from __future__ import annotations

import argparse
import csv
import json
import os
import re
from pathlib import Path

BURADA = Path(__file__).resolve().parent

# anahtar, ad, kısa ad, birim, grup, tür (fiyat: % değişim · oran: puan · stok: fark), ondalık, renk yönü (1: artış iyi, -1: artış kötü, 0: nötr)
GOSTERGE = [
    ("usd", "ABD doları", "USD/TRY", "TL", "Kurlar", "fiyat", 4, -1),
    ("eur", "Euro", "EUR/TRY", "TL", "Kurlar", "fiyat", 4, -1),
    ("sepet", "Sepet kur (0,5 USD + 0,5 EUR)", "Sepet", "TL", "Kurlar", "fiyat", 4, -1),
    ("eurusd", "EUR / USD paritesi", "EUR/USD", "", "Kurlar", "fiyat", 4, 0),
    ("aofm", "TCMB ağırlıklı ortalama fonlama maliyeti", "AOFM", "%", "Faizler", "oran", 2, 0),
    ("tlref", "BIST TLREF (gecelik)", "TLREF", "%", "Faizler", "oran", 2, 0),
    ("bist", "BIST 100 endeksi", "BIST 100", "puan", "Borsa", "fiyat", 2, 1),
    ("mali", "BIST Mali endeks", "BIST Mali", "puan", "Borsa", "fiyat", 2, 1),
    ("netrez", "TCMB net dış varlıklar (analitik bilanço)", "Net dış varlık", "mlr $", "Rezervler", "stok", 1, 1),
]


def oku_csv(yol: Path) -> dict:
    out: dict = {}
    if yol.exists():
        with yol.open(encoding="utf-8-sig") as f:
            for r in csv.DictReader(f):
                out.setdefault(r["kod"], {})[r["tarih"]] = float(r["deger"])
    return out


def gunluk_veri(G: dict) -> dict:
    g = lambda k: G.get(k, {})                                           # noqa: E731
    tarihler = sorted({t for s in G.values() for t in s})
    usd, eur = g("TP.DK.USD.A.YTL"), g("TP.DK.EUR.A.YTL")
    a02, a10 = g("TP.AB.A02"), g("TP.AB.A10")
    S = {k[0]: [] for k in GOSTERGE}
    r = lambda x, d: None if x is None else round(x, d)                  # noqa: E731
    for t in tarihler:
        u, e = usd.get(t), eur.get(t)
        S["usd"].append(r(u, 4)); S["eur"].append(r(e, 4))
        S["sepet"].append(r((u + e) / 2, 4) if u and e else None)
        S["eurusd"].append(r(e / u, 4) if u and e else None)
        S["aofm"].append(g("TP.APIFON4").get(t)); S["tlref"].append(r(g("TP.BISTTLREF.ORAN").get(t), 4))
        S["bist"].append(g("TP.MK.F.BILESIK").get(t)); S["mali"].append(g("TP.MK.F.MALI").get(t))
        S["netrez"].append(r((a02[t] - a10[t]) / u / 1e6, 2) if t in a02 and t in a10 and u else None)
    return {"t": tarihler, "s": S}


def haftalik_veri(H: dict) -> dict:
    tarihler = sorted({t for s in H.values() for t in s})
    g = lambda k, t: H.get(k, {}).get(t)                                 # noqa: E731
    r = lambda x, d=2: None if x is None else round(x, d)                # noqa: E731
    return {"t": tarihler,
            "brut": [r(g("TP.AB.TOPLAM", t) / 1000 if g("TP.AB.TOPLAM", t) is not None else None) for t in tarihler],
            "hisse": [r(g("TP.MKNETHAR.M7", t) / 1000 if g("TP.MKNETHAR.M7", t) is not None else None, 3) for t in tarihler],
            "dibs": [r(g("TP.MKNETHAR.M8", t) / 1000 if g("TP.MKNETHAR.M8", t) is not None else None, 3) for t in tarihler]}


AYLIK = [  # anahtar (makro_pano seri), ad, birim, tür, ondalık
    ("sepet", "Sepet kur (ay sonu)", "TL", "fiyat", 4),
    ("usd", "ABD doları (ay sonu)", "TL", "fiyat", 4),
    ("eur", "Euro (ay sonu)", "TL", "fiyat", 4),
    ("reer", "Reel efektif kur (TÜFE bazlı, 2025=100)", "endeks", "fiyat", 1),
    ("politika", "Politika faizi", "%", "oran", 2),
    ("tufe", "Yıllık TÜFE", "%", "oran", 2),
    ("bek12", "12 ay sonrası TÜFE beklentisi", "%", "oran", 2),
    ("reel", "Reel politika faizi (beklentiye göre)", "puan", "oran", 2),
    ("kredi", "TL ticari kredi faizi (akım)", "%", "oran", 2),
    ("mevduat", "TL mevduat faizi (akım)", "%", "oran", 2),
    ("makas", "Kredi–mevduat faiz farkı", "puan", "oran", 2),
    ("brut", "TCMB brüt rezervleri", "mlr $", "stok", 1),
    ("net", "TCMB net rezervleri", "mlr $", "stok", 1),
    ("swapsiz", "Swap hariç net rezerv", "mlr $", "stok", 1),
    ("yeterlilik", "Rezerv yeterliliği (brüt / dış finansman ihtiyacı)", "x", "stok", 2),
    ("cari12", "Cari denge (12 aylık)", "mlr $", "stok", 1),
    ("nhn12", "Net hata ve noksan (12 aylık)", "mlr $", "stok", 1),
    ("efr", "Dış finansman ihtiyacı", "mlr $", "stok", 1),
    ("dolar", "Dolarizasyon (DTH / mevduat)", "%", "oran", 1),
]

HTML = r"""<!doctype html>
<html lang="tr"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>TCMB verileri · günlük, haftalık, aylık</title>
__HEAD__
<style>
.tc-top{display:flex;align-items:center;gap:12px;padding:12px 0}.tc-top a{color:var(--accent);font-weight:500;text-decoration:none}.tc-top a:hover{text-decoration:underline}
.tc-nav{display:flex;flex-wrap:wrap;gap:8px}.tc-nav a{border:1px solid var(--line);border-radius:18px;padding:6px 12px;font-size:13.5px;color:var(--ink-2);text-decoration:none;background:var(--surface)}.tc-nav a:hover{border-color:var(--accent);color:var(--accent)}
.kpi.sel{border-color:var(--accent);background:var(--accent-soft)}
.kpi .k-d{font-size:11.5px;color:var(--ink-3)}
.tc-grp td{background:var(--surface-2);font-weight:700;font-size:12.5px;color:var(--ink-2)}
.tc-tab td.l{padding-left:22px}
.tc-tab .dim{color:var(--ink-3);font-size:12px}
.tc-row{display:flex;flex-wrap:wrap;gap:12px;align-items:center;justify-content:space-between}
.tc-note{font-size:12.5px;color:var(--ink-3)}
section{scroll-margin-top:90px}
</style>
</head><body>
<div class="wrap gut fade-in">
  <div class="tc-top"><a href="../">← Başlangıç ekranı</a></div>
  <header class="hdr">
    <div class="hdr-top"><span class="eyebrow">TCMB · Elektronik Veri Dağıtım Sistemi (EVDS)</span><span class="mode-chip" id="tcKapsam"></span></div>
    <h1>TCMB verileri</h1>
    <p class="lede">Kurlar ve sepet kur, TCMB fonlama maliyeti ve TLREF, borsa endeksleri ve TCMB'nin net dış varlıkları her iş günü; brüt rezervler ve yabancıların portföy akımları her hafta; rezerv yeterliliği, enflasyon, beklentiler, cari denge, dolarizasyon ve faizler her ay. Bir gün seçin: bütün bölümler o gün itibarıyla gösterilir.</p>
    <div class="hdr-meta"><span>Kaynak: TCMB EVDS</span><span id="tcGun"></span><span>Tutarlar nominaldir</span></div>
    <nav class="tc-nav" aria-label="Bölümler"><a href="#h-gun">Günlük göstergeler</a><a href="#h-deg">Bugün ne değişti?</a><a href="#h-haf">Haftalık</a><a href="#h-mkr">Makro pano</a><a href="#h-ayl">Aylık tablolar</a></nav>
  </header>
  <div class="pbar" role="group" aria-label="Gün ve karşılaştırma">
    <span class="pbar-label">Gün</span>
    <div class="pctl" id="gunSec"></div>
    <span class="pbar-label">Karşılaştırma</span>
    <div class="seg" id="cmpSeg" role="group" aria-label="Karşılaştırma"></div>
    <span class="pbar-note" id="cmpNot"></span>
  </div>

  <section aria-labelledby="h-gun" id="secGun">
    <div class="sec-head"><h2 id="h-gun">Günlük göstergeler</h2><p class="insight" id="gunInsight"></p></div>
    <div class="kpis" id="gunKpi"></div>
    <div class="panel">
      <div class="tc-row"><h3 id="gunChTitle"></h3><div class="seg" id="rngSeg" role="group" aria-label="Grafik aralığı"></div></div>
      <div class="chartbox tall"><canvas id="gunChart" role="img" aria-label="Seçilen göstergenin günlük seyri"></canvas></div>
      <p class="tc-note" id="gunChNote"></p>
    </div>
  </section>

  <section aria-labelledby="h-deg" id="secDeg">
    <div class="sec-head"><h2 id="h-deg">Bugün ne değişti?</h2><p class="insight" id="degInsight"></p></div>
    <div class="scroll"><table class="ratio-table tc-tab" id="degTab"></table></div>
    <div class="tc-row"><p class="tc-note">Kurlarda ve fiyatlarda % değişim, faizlerde puan, net dış varlıkta milyar $ fark. Kurlarda artış TL'nin değer kaybıdır (kırmızı). Her satır seçilen güne kadar yayımlanmış son veriyi kullanır; tarihi farklıysa yanında yazar. Yılbaşı: önceki yılın son iş günü.</p><button class="btn" type="button" id="degCsv">CSV indir</button></div>
  </section>

  <section aria-labelledby="h-haf" id="secHaf">
    <div class="sec-head"><h2 id="h-haf">Haftalık: rezervler ve portföy akımları</h2><p class="insight" id="hafInsight"></p></div>
    <div class="panel-grid">
      <div class="panel" style="min-width:0"><h3>TCMB brüt rezervleri (mlr $, cuma)</h3><div class="chartbox"><canvas id="hafRez" role="img" aria-label="Haftalık brüt rezervler"></canvas></div></div>
      <div class="panel" style="min-width:0"><h3>Yabancıların haftalık net alımı (mlr $)</h3><div class="chartbox"><canvas id="hafAkim" role="img" aria-label="Haftalık hisse ve DİBS net alımı"></canvas></div></div>
    </div>
    <div class="scroll"><table class="ratio-table" id="hafTab"></table></div>
    <div class="tc-row"><p class="tc-note">Brüt rezerv: TCMB altın + döviz rezervleri (uluslararası rezervler, haftalık). Portföy akımları fiyat ve kur etkisinden arındırılmış net alımlardır (hisse senedi ve devlet iç borçlanma senetleri).</p><button class="btn" type="button" id="hafCsv">CSV indir</button></div>
  </section>

__MAKRO__

  <section aria-labelledby="h-ayl" id="secAyl">
    <div class="sec-head"><h2 id="h-ayl">Aylık tablolar</h2><p class="insight" id="aylInsight"></p></div>
    <div class="scroll"><table class="ratio-table tc-tab" id="aylTab"></table></div>
    <div class="tc-row"><p class="tc-note">Oranlarda puan, kurlarda ve endekslerde % değişim, tutarlarda fark. Makro panodaki seriler ve ay sonu kurlar; yabancı portföy akımları haftalık bölümde.</p><button class="btn" type="button" id="aylCsv">CSV indir</button></div>
  </section>

  <details class="yoz-yontem"><summary>Kaynak ve yöntem</summary><ul>
    <li>Günlük: TP.DK.USD.A.YTL, TP.DK.EUR.A.YTL (TCMB döviz alış kuru, 15:30), TP.APIFON4 (TCMB ağırlıklı ortalama fonlama maliyeti), TP.BISTTLREF.ORAN, TP.MK.F.BILESIK, TP.MK.F.MALI, TP.AB.A02 ve TP.AB.A10 (analitik bilanço). Sepet kur = 0,5 × USD + 0,5 × EUR; EUR/USD = EUR kuru / USD kuru.</li>
    <li>Net dış varlık = (dış varlıklar − toplam döviz yükümlülükleri) / USD kuru; aylık makro panodaki "net rezerv" ile aynı tanımdır, swap düşülmemiştir. Swap hariç net rezerv yalnız aylık (Uluslararası Rezervler ve Döviz Likiditesi tablosu) yayımlandığı için aylık bölümdedir.</li>
    <li>Haftalık: TP.AB.TOPLAM (brüt rezerv), TP.MKNETHAR.M7 / M8 (yabancı hisse / DİBS net alımı).</li>
    <li id="tcArsivNot"></li>
  </ul></details>
</div>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<script>window.TCMB=__TCMB__;window.MAKRO=__MAKROVERI__;</script>
<script>
const nf1=new Intl.NumberFormat("tr-TR",{minimumFractionDigits:1,maximumFractionDigits:1});
const nf2=new Intl.NumberFormat("tr-TR",{minimumFractionDigits:2,maximumFractionDigits:2});
const nfd=d=>new Intl.NumberFormat("tr-TR",{minimumFractionDigits:d,maximumFractionDigits:d});
const css=n=>getComputedStyle(document.documentElement).getPropertyValue(n).trim();
const el=(tag,attrs={},html="")=>{ const e=document.createElement(tag); for(const k in attrs) e.setAttribute(k,attrs[k]); if(html) e.innerHTML=html; return e; };
const esc=s=>String(s).replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const charts={};
function mkChart(id,cfg){ if(charts[id]) charts[id].destroy(); const e=document.getElementById(id); if(!e||!window.Chart) return; charts[id]=new Chart(e,cfg); return charts[id]; }
__BASEOPTS__
const T=window.TCMB, D=T.gun, W=T.haf, G=Object.fromEntries(T.gosterge.map(g=>[g.k,g]));
const AYK=["Oca","Şub","Mar","Nis","May","Haz","Tem","Ağu","Eyl","Eki","Kas","Ara"];
const gl=t=>`${+t.slice(8)} ${AYK[+t.slice(5,7)-1]} ${t.slice(0,4)}`, gk=t=>`${t.slice(8)}.${t.slice(5,7)}.${t.slice(2,4)}`;
let gi=D.t.length-1, cmp="1g", rng="1y", sel="sepet", curP="";
try{ const s=JSON.parse(localStorage.getItem("tcmb.d")||"{}"); if(s.cmp) cmp=s.cmp; if(s.rng) rng=s.rng; if(s.sel&&G[s.sel]) sel=s.sel; }catch(e){}
const kaydet=()=>{ try{ localStorage.setItem("tcmb.d",JSON.stringify({cmp,rng,sel})); }catch(e){} };
const CMP={"1g":"Önceki gün","1h":"1 hafta","1a":"1 ay","yb":"Yılbaşı","1y":"1 yıl"}, CMPG={"1g":"önceki güne göre","1h":"bir hafta öncesine göre","1a":"bir ay öncesine göre","yb":"yılbaşına göre","1y":"bir yıl öncesine göre"};
function gunEkle(t,{d=0,m=0,y=0}){ const x=new Date(t+"T00:00:00Z"); x.setUTCFullYear(x.getUTCFullYear()+y); x.setUTCMonth(x.getUTCMonth()+m); x.setUTCDate(x.getUTCDate()+d); return x.toISOString().slice(0,10); }
// k serisinin i. güne kadarki son değeri: [indeks, değer]
function sonDeger(k,i){ const s=D.s[k]; for(let j=i;j>=0;j--) if(s[j]!=null) return [j,s[j]]; return [-1,null]; }
function refDeger(k,i,c){ const [j]=sonDeger(k,i); if(j<0) return [-1,null]; const t=D.t[j];
  let hedef; if(c==="1g") return sonDeger(k,j-1);
  if(c==="1h") hedef=gunEkle(t,{d:-7}); else if(c==="1a") hedef=gunEkle(t,{m:-1}); else if(c==="1y") hedef=gunEkle(t,{y:-1}); else hedef=(+t.slice(0,4)-1)+"-12-31";
  let m=-1; for(let q=0;q<D.t.length&&D.t[q]<=hedef;q++) m=q; if(m<0) return [-1,null]; const r=sonDeger(k,m); return r; }
function degis(g,v,r){ if(v==null||r==null) return null; return g.tur==="fiyat"?(v/r-1)*100:(v-r); }
function degFmt(g,x){ if(x==null) return "–"; const s=x>0?"+":x<0?"−":""; const a=Math.abs(x);
  return g.tur==="fiyat"?`${s}${nf2.format(a)}%`:g.tur==="oran"?`${s}${nf2.format(a)} puan`:`${s}${nfd(g.d).format(a)} ${g.b}`; }
function renk(g,x){ if(x==null||Math.abs(x)<1e-9||!g.yon) return ""; return (x*g.yon>0)?"pos":"neg"; }
function degerFmt(g,v){ return v==null?"–":nfd(g.d).format(v); }

function renderBar(){
  const box=document.getElementById("gunSec"); box.innerHTML="";
  const prev=el("button",{type:"button",class:"icon-btn","aria-label":"Önceki gün",title:"Önceki gün"},'<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M15.41 7.41 14 6l-6 6 6 6 1.41-1.41L10.83 12z"/></svg>');
  const next=el("button",{type:"button",class:"icon-btn","aria-label":"Sonraki gün",title:"Sonraki gün"},'<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M10 6 8.59 7.41 13.17 12l-4.58 4.59L10 18l6-6z"/></svg>');
  if(gi<=0) prev.disabled=true; if(gi>=D.t.length-1) next.disabled=true;
  prev.onclick=()=>{gi--;renderAll();}; next.onclick=()=>{gi++;renderAll();};
  const s=el("select",{"aria-label":"Gün"}); for(let i=D.t.length-1;i>=0;i--) s.appendChild(el("option",{value:i},gl(D.t[i]))); s.value=gi; s.onchange=()=>{gi=+s.value;renderAll();};
  box.append(prev,s,next);
  const cs=document.getElementById("cmpSeg"); cs.innerHTML="";
  Object.entries(CMP).forEach(([k,l])=>{ const b=el("button",{type:"button","aria-pressed":String(k===cmp)},l); b.onclick=()=>{cmp=k;kaydet();renderAll();}; cs.appendChild(b); });
  const [rj]=refDeger("sepet",gi,cmp);
  document.getElementById("cmpNot").textContent=rj>=0?`${gl(D.t[gi])} ↔ ${gl(D.t[rj])}`:"";
}
function renderGun(){
  const box=document.getElementById("gunKpi"); box.innerHTML="";
  T.gosterge.forEach(g=>{ const [j,v]=sonDeger(g.k,gi), [,r]=refDeger(g.k,gi,cmp), x=degis(g,v,r);
    const b=el("button",{type:"button",class:"kpi"+(g.k===sel?" sel":""),"aria-pressed":String(g.k===sel)},
      `<span class="k-l">${esc(g.kisa)}</span><span class="k-v">${degerFmt(g,v)}</span><span class="k-u">${esc(g.b)}${j>=0&&j!==gi?` · ${gk(D.t[j])}`:""}</span><span class="k-c ${renk(g,x)}">${degFmt(g,x)} <span class="muted" style="font-weight:400">${esc(CMPG[cmp])}</span></span>`);
    b.onclick=()=>{sel=g.k;kaydet();renderGun();}; box.appendChild(b); });
  const g=G[sel], gun=D.t[gi];
  const bas={"1a":gunEkle(gun,{m:-1}),"3a":gunEkle(gun,{m:-3}),"1y":gunEkle(gun,{y:-1}),"tum":"0000"}[rng];
  const idx=D.t.map((t,i)=>i).filter(i=>D.t[i]>=bas&&i<=gi);
  const grup=g.g==="Kurlar"&&g.k!=="eurusd"?["usd","eur","sepet"]:g.g==="Faizler"?["aofm","tlref"]:g.g==="Borsa"?[g.k]:[g.k];
  const cols={usd:"--b0",eur:"--b1",sepet:"--b4",aofm:"--b0",tlref:"--b3",bist:"--b0",mali:"--b2",netrez:"--b0",eurusd:"--b5"};
  const ds=grup.map(k=>({label:G[k].kisa,data:idx.map(i=>D.s[k][i]),borderColor:css(cols[k]),backgroundColor:css(cols[k]),borderWidth:k===sel?3:1.6,pointRadius:0,tension:.15,spanGaps:true}));
  mkChart("gunChart",{type:"line",data:{labels:idx.map(i=>gk(D.t[i])),datasets:ds},options:baseOpts(x=>nfd(Math.min(g.d,2)).format(x),{plugins:{legend:{display:ds.length>1,labels:{color:css("--ink-2"),boxWidth:10,boxHeight:10}}},scales:{x:{ticks:{color:css("--ink-2"),maxTicksLimit:10,autoSkip:true}},y:{ticks:{color:css("--ink-2")}}}})});
  document.getElementById("gunChTitle").textContent=`${g.ad} (${g.b||"oran"}), ${gl(D.t[idx[0]??gi])} – ${gl(gun)}`;
  document.getElementById("gunChNote").textContent=g.k==="netrez"?"Analitik bilançodan günlük: dış varlıklar − toplam döviz yükümlülükleri, günün dolar kuruyla. Swap düşülmemiştir.":g.g==="Kurlar"?"TCMB gösterge niteliğindeki döviz alış kurları (15:30). Sepet kur TL'nin dolar ve euro karşısındaki ortalama değeridir.":g.g==="Faizler"?"AOFM: TCMB'nin bankaları fonladığı ağırlıklı ortalama faiz; TLREF: piyasa gecelik referans faizi. İkisinin farkı likidite koşullarını gösterir.":"Kapanış değerleri; BIST Mali endeksi bankaları ve öteki finansal şirketleri kapsar.";
  const rs=document.getElementById("rngSeg"); rs.innerHTML="";
  [["1a","1 ay"],["3a","3 ay"],["1y","1 yıl"],["tum","Tümü"]].forEach(([k,l])=>{ const b=el("button",{type:"button","aria-pressed":String(k===rng)},l); b.onclick=()=>{rng=k;kaydet();renderGun();}; rs.appendChild(b); });
  const v=k=>sonDeger(k,gi)[1], d=k=>{ const g=G[k]; return degis(g,v(k),refDeger(k,gi,cmp)[1]); };
  document.getElementById("gunInsight").innerHTML=`${gl(gun)}: sepet kur <b>${degerFmt(G.sepet,v("sepet"))} TL</b> (${esc(CMPG[cmp])} ${degFmt(G.sepet,d("sepet"))}), dolar ${degerFmt(G.usd,v("usd"))}, euro ${degerFmt(G.eur,v("eur"))}. TCMB fonlama maliyeti <b>%${nf2.format(v("aofm")??0)}</b>, TLREF %${nf2.format(v("tlref")??0)}. BIST 100 ${degerFmt(G.bist,v("bist"))} (${degFmt(G.bist,d("bist"))}), BIST Mali ${degFmt(G.mali,d("mali"))}. TCMB net dış varlıkları <b>${degerFmt(G.netrez,v("netrez"))} mlr $</b>.`;
}
function renderDeg(){
  const C=["1g","1h","1a","yb","1y"];
  let h=`<thead><tr><th>Gösterge</th><th>${esc(gl(D.t[gi]))}</th>${C.map(c=>`<th>${CMP[c]}</th>`).join("")}</tr></thead><tbody>`;
  let grp=""; const csv=[["Gösterge","Birim","Tarih","Değer",...C.map(c=>CMP[c])]];
  T.gosterge.forEach(g=>{ if(g.g!==grp){ grp=g.g; h+=`<tr class="tc-grp"><td colspan="7">${esc(grp)}</td></tr>`; }
    const [j,v]=sonDeger(g.k,gi); const xs=C.map(c=>degis(g,v,refDeger(g.k,gi,c)[1]));
    h+=`<tr><td class="l">${esc(g.ad)} <span class="dim">${esc(g.b)}</span>${j>=0&&j!==gi?` <span class="dim">(${gk(D.t[j])})</span>`:""}</td><td>${degerFmt(g,v)}</td>${xs.map(x=>`<td class="${renk(g,x)}">${degFmt(g,x)}</td>`).join("")}</tr>`;
    csv.push([g.ad,g.b,j>=0?D.t[j]:"",v??"",...xs.map(x=>x==null?"":x.toFixed(4))]); });
  document.getElementById("degTab").innerHTML=h+"</tbody>";
  document.getElementById("degCsv").onclick=()=>indir(`tcmb_gunluk_${D.t[gi]}.csv`,csv);
  const x=k=>degis(G[k],sonDeger(k,gi)[1],refDeger(k,gi,"yb")[1]);
  document.getElementById("degInsight").innerHTML=`Yılbaşından bu yana sepet kur ${degFmt(G.sepet,x("sepet"))}, BIST 100 ${degFmt(G.bist,x("bist"))}, BIST Mali ${degFmt(G.mali,x("mali"))}; fonlama maliyeti ${degFmt(G.aofm,x("aofm"))}, net dış varlıklar ${degFmt(G.netrez,x("netrez"))}.`;
}
function renderHaf(){
  const gun=D.t[gi]; let n=-1; for(let i=0;i<W.t.length&&W.t[i]<=gun;i++) n=i;
  if(n<0){ document.getElementById("hafInsight").textContent="Seçilen güne kadar haftalık veri yok."; return; }
  const bas=gunEkle(W.t[n],{y:-1}), idx=W.t.map((t,i)=>i).filter(i=>W.t[i]>=bas&&i<=n), L=idx.map(i=>gk(W.t[i]));
  mkChart("hafRez",{type:"line",data:{labels:L,datasets:[{label:"Brüt rezerv",data:idx.map(i=>W.brut[i]),borderColor:css("--b0"),backgroundColor:css("--b0"),borderWidth:2.5,pointRadius:0,tension:.2}]},options:baseOpts(x=>nf1.format(x),{plugins:{legend:{display:false}}})});
  mkChart("hafAkim",{type:"bar",data:{labels:L,datasets:[{label:"Hisse",data:idx.map(i=>W.hisse[i]),backgroundColor:css("--b0"),borderRadius:2},{label:"DİBS",data:idx.map(i=>W.dibs[i]),backgroundColor:css("--b1"),borderRadius:2}]},options:baseOpts(x=>nf2.format(x),{scales:{x:{stacked:true,ticks:{color:css("--ink-2"),maxTicksLimit:12}},y:{stacked:true,ticks:{color:css("--ink-2")}}}})});
  const son=[...Array(Math.min(8,n+1)).keys()].map(k=>n-k);
  const top=(a,b)=>{ let h=0,d=0; for(let i=a;i<=b;i++){ h+=W.hisse[i]||0; d+=W.dibs[i]||0; } return [h,d]; };
  const yb=W.t.findIndex(t=>t>=W.t[n].slice(0,4)+"-01-01"); const [yh,yd]=top(yb,n), [fh,fd]=top(Math.max(0,n-3),n);
  const sg=x=>x==null?"–":(x>0?"+":x<0?"−":"")+nf2.format(Math.abs(x)), cl=x=>x>0?"pos":x<0?"neg":"";
  let h=`<thead><tr><th>Hafta</th><th>Brüt rezerv<br><span class="muted">mlr $</span></th><th>Haftalık değişim</th><th>Hisse</th><th>DİBS</th><th>Toplam akım</th></tr></thead><tbody>`;
  const csv=[["Hafta","Brüt rezerv (mlr $)","Haftalık değişim","Hisse (mlr $)","DİBS (mlr $)","Toplam"]];
  son.forEach(i=>{ const dg=i>0&&W.brut[i]!=null&&W.brut[i-1]!=null?W.brut[i]-W.brut[i-1]:null, t=(W.hisse[i]||0)+(W.dibs[i]||0);
    h+=`<tr><td>${gl(W.t[i])}</td><td>${W.brut[i]==null?"–":nf1.format(W.brut[i])}</td><td class="${cl(dg)}">${sg(dg)}</td><td class="${cl(W.hisse[i])}">${sg(W.hisse[i])}</td><td class="${cl(W.dibs[i])}">${sg(W.dibs[i])}</td><td class="${cl(t)}"><b>${sg(t)}</b></td></tr>`;
    csv.push([W.t[i],W.brut[i]??"",dg??"",W.hisse[i]??"",W.dibs[i]??"",t]); });
  h+=`<tr class="tc-grp"><td>Son 4 hafta</td><td></td><td></td><td>${sg(fh)}</td><td>${sg(fd)}</td><td>${sg(fh+fd)}</td></tr><tr class="tc-grp"><td>Yılbaşından</td><td></td><td></td><td>${sg(yh)}</td><td>${sg(yd)}</td><td>${sg(yh+yd)}</td></tr>`;
  document.getElementById("hafTab").innerHTML=h+"</tbody>";
  document.getElementById("hafCsv").onclick=()=>indir(`tcmb_haftalik_${W.t[n]}.csv`,csv);
  const d1=n>0?W.brut[n]-W.brut[n-1]:null;
  document.getElementById("hafInsight").innerHTML=`${gl(W.t[n])} haftasında brüt rezervler <b>${nf1.format(W.brut[n])} mlr $</b> (haftalık ${sg(d1)} mlr $). Yabancılar son 4 haftada hisse ${sg(fh)}, DİBS ${sg(fd)} mlr $; yılbaşından toplam <b>${sg(yh+yd)} mlr $</b>.`;
}
function renderAyl(){
  const MM=window.MAKRO, ay=D.t[gi].slice(0,7), A=MM.aylar, AS=T.aylik;
  let n=-1; for(let i=0;i<A.length&&A[i]<=ay;i++) n=i;
  const ser=k=>AS[k]||MM.seri[k]||[];
  const son=k=>{ const s=ser(k); for(let j=n;j>=0;j--) if(s[j]!=null) return j; return -1; };
  const dd=(t,k,j,m)=>{ const s=ser(k); const r=j-m; if(j<0||r<0||s[j]==null||s[r]==null) return null; return t==="fiyat"?(s[j]/s[r]-1)*100:s[j]-s[r]; };
  const ft=(t,b,x,d)=>x==null?"–":(x>0?"+":x<0?"−":"")+(t==="fiyat"?nf2.format(Math.abs(x))+"%":t==="oran"?nf2.format(Math.abs(x))+" puan":nfd(d).format(Math.abs(x))+" "+b);
  let h=`<thead><tr><th>Seri</th><th>Son değer</th><th>Ay</th><th>1 ay</th><th>3 ay</th><th>Yılbaşından</th><th>1 yıl</th></tr></thead><tbody>`;
  const csv=[["Seri","Birim","Ay","Değer","1 ay","3 ay","Yılbaşından","1 yıl"]];
  T.aylikTanim.forEach(([k,ad,b,t,d])=>{ const j=son(k); if(j<0) return; const s=ser(k), ybm=+A[j].slice(5,7);
    const xs=[dd(t,k,j,1),dd(t,k,j,3),dd(t,k,j,ybm),dd(t,k,j,12)];
    h+=`<tr><td class="l">${esc(ad)} <span class="dim">${esc(b)}</span></td><td>${nfd(d).format(s[j])}</td><td>${esc(MM.etiket[j])}</td>${xs.map(x=>`<td>${ft(t,b,x,d)}</td>`).join("")}</tr>`;
    csv.push([ad,b,A[j],s[j],...xs.map(x=>x==null?"":x.toFixed(4))]); });
  document.getElementById("aylTab").innerHTML=h+"</tbody>";
  document.getElementById("aylCsv").onclick=()=>indir(`tcmb_aylik_${ay}.csv`,csv);
  document.getElementById("aylInsight").textContent=`${n>=0?MM.etiket[n]:"–"} itibarıyla; her seri seçilen aya kadar yayımlanmış son değeriyle. Ödemeler dengesi ve dış borç gibi seriler 1–2 ay gecikmeyle gelir.`;
}
function indir(ad,rows){ const s="﻿"+rows.map(r=>r.map(x=>{ x=String(x); return /[;"\n]/.test(x)?`"${x.replace(/"/g,'""')}"`:x; }).join(";")).join("\n");
  const a=el("a",{href:URL.createObjectURL(new Blob([s],{type:"text/csv;charset=utf-8"})),download:ad}); document.body.appendChild(a); a.click(); a.remove(); }
__MAKROJS__
if(window.Chart){ Chart.defaults.font.family="Roboto"; Chart.defaults.color=css("--ink-2"); }
function renderAll(){
  curP=D.t[gi].slice(0,7);
  renderBar(); renderGun(); renderDeg(); renderHaf(); try{ renderMakro(); }catch(e){ console.error(e); } renderAyl();
}
document.getElementById("tcKapsam").textContent=`${gl(D.t[0])} – ${gl(D.t[D.t.length-1])} · ${D.t.length} iş günü`;
document.getElementById("tcGun").textContent=`Veri ${T.guncelleme} tarihinde alındı`;
document.getElementById("tcArsivNot").textContent=T.arsivNot;
renderAll();
</script>
</body></html>
"""


def uret(site: Path, arsiv: Path | None = None, makro_veri: Path | None = None):
    import makro_pano as mp
    arsiv = Path(arsiv or os.environ.get("TCMB_ARSIV") or BURADA / "tcmb_arsiv")
    makro_veri = Path(makro_veri or mp.VERI)
    V = json.loads(makro_veri.read_text(encoding="utf-8"))
    M = mp.hesapla(V, aylik=True)
    # aylık sepet / USD / EUR (ay sonu) — makro aylarına hizalı
    S = V["seriler"]
    us, eu = S.get("TP.DK.USD.A.YTL", {}), S.get("TP.DK.EUR.A.YTL", {})
    aylik = {"usd": [us.get(a) for a in M["aylar"]], "eur": [eu.get(a) for a in M["aylar"]],
             "sepet": [round((us[a] + eu[a]) / 2, 4) if a in us and a in eu else None for a in M["aylar"]]}
    gun = gunluk_veri(oku_csv(arsiv / "gunluk.csv"))
    haf = haftalik_veri(oku_csv(arsiv / "haftalik.csv"))
    if not gun["t"]:
        raise SystemExit(f"TCMB arşivi boş: {arsiv} (önce tcmb_veri.py çalıştırın)")
    guncel = (arsiv / "son_tarih.txt").read_text().strip() if (arsiv / "son_tarih.txt").exists() else gun["t"][-1]
    ilk = gun["t"][0]
    arsiv_not = (f"Günlük arşiv {ilk} tarihinden başlıyor. Tam geçmiş (2018'den) EVDS anahtarıyla tcmb_veri.py "
                 "ilk kez çalıştırıldığında gelir; sonraki çalıştırmalar yalnız yeni günleri ekler.") if ilk > "2018-12-31" else \
        f"Günlük arşiv {ilk} tarihinden başlıyor; her çalıştırmada yalnız yeni günler eklenir."
    T = {"gosterge": [dict(k=k, ad=a, kisa=ki, b=b, g=g, tur=t, d=d, yon=y) for k, a, ki, b, g, t, d, y in GOSTERGE],
         "gun": gun, "haf": haf, "aylik": aylik, "aylikTanim": AYLIK, "guncelleme": V.get("alindi") or guncel,
         "arsivNot": arsiv_not}

    html = (site / "index.html").read_text(encoding="utf-8")
    bas = html[:html.index("</head>")]
    bas = bas[bas.index("<link"):] if "<link" in bas else bas
    bas = re.sub(r"<script\b.*?</script>", "", bas, flags=re.S)
    i = html.index("function baseOpts")
    base = html[i:html.index("\n}\n", i) + 3]
    makro_html = mp.HTML.replace("<!--MKR-->", "").replace("<!--/MKR-->", "").replace(
        "seçili dönem sonundaki ya da ondan önceki son anket", "seçili ayın sonundaki ya da ondan önceki son anket").replace(
        "Göstergeler sayfanın üstünde seçilen raporlama döneminin sonu itibarıyladır (ör. Eylül 2026 raporunda Eylül 2026 sonu); grafikler 2018'den o aya kadar çizilir.",
        "Göstergeler seçilen günün ayı itibarıyladır; grafikler 2018'den o aya kadar çizilir.").replace(
        "<h2 id=\"h-mkr\">Makro pano</h2>", "<h2 id=\"h-mkr\">Makro pano (aylık)</h2>")
    makro_js = mp.JS.replace(" (raporun dönem sonu):", ":").replace(
        "bu fark aşağıdaki karnede yer alan swap düzeltmeli marjın öncüsüdür", "bu fark finansal raporlardaki swap düzeltmeli marjın öncüsüdür")
    js = lambda o: json.dumps(o, ensure_ascii=False, separators=(",", ":"))   # noqa: E731
    out = (HTML.replace("__HEAD__", bas).replace("__MAKRO__", makro_html).replace("__BASEOPTS__", base)
           .replace("__MAKROJS__", makro_js).replace("__TCMB__", js(T)).replace("__MAKROVERI__", js(M)))
    (site / "tcmb").mkdir(exist_ok=True)
    (site / "tcmb" / "index.html").write_text(out, encoding="utf-8")
    print(f"TCMB sayfası: {site / 'tcmb' / 'index.html'} ({len(gun['t'])} gün, {len(haf['t'])} hafta, {len(M['aylar'])} ay)")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--site", required=True)
    ap.add_argument("--arsiv")
    ap.add_argument("--makro-veri")
    a = ap.parse_args()
    uret(Path(a.site), a.arsiv, a.makro_veri)
