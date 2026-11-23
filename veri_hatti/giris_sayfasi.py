#!/usr/bin/env python3
"""
Giriş ekranı: üç veri kaynağı → (1) Finansal raporlar (TBB; tıklanınca mevcut dört seçenek açılır),
(2) TCMB verileri (tcmb/ sayfası, makro_pano.py üretir), (3) BDDK verileri (ayrı depo, bddk-veri).

makro_pano.uygula() önizlemeye uygularken bunu da çağırır; tek başına:
    python giris_sayfasi.py --site docs/onizleme
Tekrar çalıştırmak güvenlidir (önceki yama temizlenir).
"""
from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

BURADA = Path(__file__).resolve().parent
BDDK_ADRES = "https://resulkonur-gif.github.io/bddk-veri/"
OK = '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 4l-1.41 1.41L16.17 11H4v2h12.17l-5.58 5.59L12 20l8-8z"/></svg>'
DIS = '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M19 19H5V5h7V3H5a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14c1.1 0 2-.9 2-2v-7h-2v7zM14 3v2h3.59l-9.83 9.83 1.41 1.41L19 6.41V10h2V3h-7z"/></svg>'

UST_METIN = ("Türk bankacılık sistemini üç kaynaktan izleyin: TBB'nin yayımladığı banka banka finansal tablolar ve "
             "dipnotlar, TCMB'nin makro verileri (rezervler, faiz, enflasyon, dış finansman, kur) ve BDDK'nın sektör "
             "ve banka grubu bültenleri. Her kaynak kendi sayfasında; veriler birbirine karıştırılmaz.")


def kartlar(makro_ozet: str, bddk: tuple | None = None) -> str:
    bddk_href, bddk_hedef, bddk_go, bddk_cips = bddk or (BDDK_ADRES, ' target="_blank" rel="noopener"', DIS, '<span class="chip">Aylık bülten · 2018\'den</span><span class="chip">Haftalık bülten · 2014\'ten</span><span class="chip">10 banka grubu</span>')
    return f"""<!--GIRIS-->
    <div class="choices" id="chTop">
      <button type="button" class="choice" id="chFinBtn">
        <span class="ch-ico"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 10h3v7H4zm6.5 0h3v7h-3zM2 19h20v3H2zm15-9h3v7h-3zm-5-9L2 6v2h20V6z"/></svg></span>
        <span class="ch-t">Finansal raporlar</span>
        <span class="ch-d">TBB Veri Sistemi'ndeki banka finansal tabloları ve dipnotları: solo ve konsolide raporlar, banka bazlı karşılaştırma, tek banka görünümü.</span>
        <span class="chips" id="chipsFin"></span>
        <span class="ch-go">Raporları aç {OK}</span>
      </button>
      <a class="choice t" href="tcmb/">
        <span class="ch-ico"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M3.5 18.49 9.5 12.48l4 4L22 6.92 20.59 5.5l-7.09 7.97-4-4L2 16.99z"/></svg></span>
        <span class="ch-t">TCMB verileri</span>
        <span class="ch-d">Her gün kurlar ve sepet kur, TCMB fonlama maliyeti, TLREF, BIST 100 ve Mali endeks, TCMB net dış varlıkları; her hafta brüt rezervler ve yabancı portföy akımları; her ay makro pano ve aylık tablolar.</span>
        <span class="chips">{makro_ozet}</span>
        <span class="ch-go">TCMB verilerini aç {OK}</span>
      </a>
      <a class="choice k" href="{bddk_href}"{bddk_hedef}>
        <span class="ch-ico"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M19 3H5a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2V5a2 2 0 0 0-2-2zM9 17H7v-7h2v7zm4 0h-2V7h2v10zm4 0h-2v-4h2v4z"/></svg></span>
        <span class="ch-t">BDDK verileri</span>
        <span class="ch-d">BDDK aylık ve haftalık bültenleri: sektör ve banka grupları için bilanço, krediler, mevduat, menkul kıymetler, sermaye yeterliliği ve rasyolar.</span>
        <span class="chips">{bddk_cips}</span>
        <span class="ch-go">BDDK verilerini aç {bddk_go}</span>
      </a>
    </div>
<!--/GIRIS-->"""


FOOT_UST = ('<!--GIRIS--><span id="footTop">Kaynaklar: banka finansal tabloları ve dipnotları <b>TBB Veri Sistemi</b>; '
            'TBB\'nin henüz yayımlamadığı dönemin ön verisi, serbest karşılık ve TÜFE varsayımları bankaların <b>KAP</b> finansal raporları; '
            'makro, kur, faiz ve rezerv verileri <b>TCMB EVDS</b>; sektör ve banka grubu verileri <b>BDDK</b> aylık ve haftalık bültenleri. '
            'Tutarlar kaynağın yayımladığı hâliyle, nominaldir; kaynaklar birbirine karıştırılmaz.</span><!--/GIRIS-->')
FOOT_FIN = ('<!--GIRIS--><span id="footFin" hidden>Solo ya da konsolide seçimini her sayfanın üst çubuğundan istediğiniz an '
            'değiştirebilirsiniz. Tutarlar TBB\'nin yayımladığı hâliyle, nominaldir; TBB\'nin henüz yayımlamadığı dönem '
            'önizlemede bankaların KAP raporlarından gelir.</span><!--/GIRIS-->')

CSS = """<style id="girisCss">
#chTop{grid-template-columns:repeat(3,minmax(0,1fr))}
@media (max-width:900px){#chTop{grid-template-columns:minmax(0,1fr)}}
a.choice{color:inherit;text-decoration:none}
.land-back{all:unset;margin-bottom:-16px;cursor:pointer;display:inline-flex;align-items:center;gap:6px;font-size:14px;font-weight:500;color:var(--m-primary);padding:6px 10px 6px 4px;border-radius:20px;align-self:flex-start}
.land-back:hover{background:var(--m-scl)} .land-back svg{width:18px;height:18px;fill:currentColor}
#footTop[hidden],#footFin[hidden],.land-back[hidden],#chTop[hidden],#chFin[hidden],#subTop[hidden],#subFin[hidden]{display:none}
</style>"""

JS = """<script id="girisJs">
(function(){
  const g=id=>document.getElementById(id), top=g("chTop"), fin=g("chFin"), back=g("landBack"), sT=g("subTop"), sF=g("subFin"), q=g("landQ");
  if(!top||!fin) return;
  const kick=document.querySelector("#landing .land-kicker"), kickFin=kick?kick.textContent:"";
  try{ const d=DATA.solo, k=DATA.kons;
    g("chipsFin").innerHTML=[`${d.A.names.length} banka`,`${d.periods.length} çeyrek`,"Solo ve konsolide","Dipnotlar"].map(x=>`<span class="chip">${x}</span>`).join("");
  }catch(e){}
  const fT=g("footTop"), fF=g("footFin");
  function adim(f){ if(fT) fT.hidden=f; if(fF) fF.hidden=!f; if(kick) kick.textContent=f?kickFin:"TBB · TCMB · BDDK · Türk bankacılık sistemi verileri"; top.hidden=f; fin.hidden=!f; back.hidden=!f; sT.hidden=f; sF.hidden=!f;
    if(q) q.textContent=f?"Hangi görünümle başlayalım?":"Hangi veriye bakalım?";
    try{ history.replaceState(null,"",f?"#finansal":location.pathname+location.search); }catch(e){} window.scrollTo(0,0); }
  g("chFinBtn").onclick=()=>adim(true); back.onclick=()=>adim(false);
  adim(location.hash==="#finansal");
})();
</script>"""


def makro_ozeti(site: Path) -> str:
    """TCMB kartı için küçük özet (tcmb sayfasının verisinden): kapsam ve son gün."""
    AY = ["Oca", "Şub", "Mar", "Nis", "May", "Haz", "Tem", "Ağu", "Eyl", "Eki", "Kas", "Ara"]
    try:
        h = (site / "tcmb" / "index.html").read_text(encoding="utf-8")
        T = json.loads(re.search(r"window\.TCMB=(\{.*?\});window\.MAKRO=", h, re.S).group(1))
        son = T["gun"]["t"][-1]
        cips = ["Günlük · haftalık · aylık", f"{len(T['gosterge'])} günlük gösterge",
                f"Son gün {int(son[8:])} {AY[int(son[5:7]) - 1]} {son[:4]}"]
    except Exception:                                                   # noqa: BLE001
        cips = ["Günlük · haftalık · aylık", "Makro pano"]
    return "".join(f'<span class="chip">{c}</span>' for c in cips)


def bddk_ozeti(site: Path):
    """site/bddk/ varsa kart oraya bağlanır; çipler BDDK sayfalarının kapsamından."""
    k = site / "bddk"
    if not (k / "index.html").exists():
        return None
    def kapsam(ad):
        try:
            h = (k / ad).read_text(encoding="utf-8") if (k / ad).exists() else ""
            m = re.search(r'<span class="land-kicker">[^<]*?·\s*([^<]+)</span>', h)
            return m.group(1).strip() if m else None
        except Exception:                                               # noqa: BLE001
            return None
    cips = []
    a, w = kapsam("aylik.html") or kapsam("index.html"), kapsam("haftalik.html")
    if a: cips.append(f"Aylık · {a}")
    if w: cips.append(f"Haftalık · {w}")
    cips.append("10 banka grubu")
    return ("bddk/", "", OK, "".join(f'<span class="chip">{c}</span>' for c in cips))


def uygula(site: Path):
    yol = Path(site) / "index.html"
    h = yol.read_text(encoding="utf-8")
    # önceki yamayı temizle
    h = re.sub(r'<!--GIRIS--><span id="footTop">.*?<!--/GIRIS--><!--GIRIS--><span id="footFin" hidden>.*?<!--/GIRIS-->',
               "Solo ya da konsolide seçimini her sayfanın üst çubuğundan istediğiniz an değiştirebilirsiniz. Tutarlar TBB’nin yayımladığı haliyle, nominaldir.",
               h, flags=re.S)
    h = re.sub(r"\s*<!--GIRIS-->.*?<!--/GIRIS-->", "", h, flags=re.S)
    h = re.sub(r'<style id="girisCss">.*?</style>\s*', "", h, flags=re.S)
    h = re.sub(r'<script id="girisJs">.*?</script>\s*', "", h, flags=re.S)
    h = h.replace('<p class="land-sub" id="subFin" hidden>', '<p class="land-sub">')
    h = h.replace('<div class="choices" id="chFin" hidden>', '<div class="choices">')
    h = h.replace('<p class="land-q" id="landQ">', '<p class="land-q">')

    i = h.index('<div id="landing"')
    j = h.index('<div id="app"', i)
    land = h[i:j]
    assert land.count('<p class="land-sub">') == 1 and land.count('<div class="choices">') == 1, "giriş ekranı yapısı değişmiş"
    # sıra: başlık · subTop (genel) / subFin (TBB metni) · soru · geri düğmesi · üç kaynak kartı / dört rapor kartı
    land = land.replace('<p class="land-sub">', f'<!--GIRIS--><p class="land-sub" id="subTop">{UST_METIN}</p><!--/GIRIS-->'
                        '\n    <p class="land-sub" id="subFin" hidden>', 1)
    land = land.replace('<p class="land-q">', '<!--GIRIS--><button type="button" class="land-back" id="landBack" hidden><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M20 11H7.83l5.59-5.59L12 4l-8 8 8 8 1.41-1.41L7.83 13H20v-2z"/></svg>Veri kaynakları</button><!--/GIRIS-->\n    <p class="land-q" id="landQ">', 1)
    land = land.replace('<div class="choices">', kartlar(makro_ozeti(Path(site)), bddk_ozeti(Path(site))) + '\n    <div class="choices" id="chFin" hidden>', 1)
    eski_not = "Solo ya da konsolide seçimini her sayfanın üst çubuğundan istediğiniz an değiştirebilirsiniz. Tutarlar TBB’nin yayımladığı haliyle, nominaldir."
    if eski_not in land:
        land = land.replace(eski_not, FOOT_UST + FOOT_FIN, 1)
    h = h[:i] + land + h[j:]
    # landing'deki tıklama işleyicisi yalnız rapor kartlarına bağlansın (yeni kartlar bağlantı / adım düğmesi)
    eski = 'document.querySelectorAll(".choice").forEach(b=>b.onclick='
    if eski in h:
        h = h.replace(eski, 'document.querySelectorAll(".choice[data-mode],.choice[data-view]").forEach(b=>b.onclick=', 1)
    assert '.choice[data-mode],.choice[data-view]' in h, "landing tıklama işleyicisi bulunamadı"
    h = h.replace("</head>", CSS + "\n</head>", 1)
    k = h.rindex("</body>")
    h = h[:k] + JS + "\n" + h[k:]
    yol.write_text(h, encoding="utf-8")
    print(f"giriş ekranı: üç kaynak kartı eklendi ({yol})")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--site", required=True)
    uygula(Path(ap.parse_args().site))
