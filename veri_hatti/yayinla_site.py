#!/usr/bin/env python3
"""
veri_hatti.py çıktısını depodaki docs/ klasörüne yazar (GitHub Pages).

  ANA_KLASOR/7_site/site/      → docs/index.html, docs/dipnot/, docs/tcmb/, docs/bddk/
  ANA_KLASOR/7_site/onizleme/  → docs/onizleme/  (TBB'nin henüz yayımlamadığı dönem)

Önizleme kuralı: bu çalıştırmada önizleme üretildiyse docs/onizleme onunla değişir. Üretilmediyse
docs/onizleme yalnız TBB o dönemi yayımlamışsa kaldırılır; KAP'a ulaşılamadığı için üretilemediyse eskisi kalır.

    python yayinla_site.py --kok ANA_KLASOR --docs DEPO/docs
Dönen kod 0; değişiklik olup olmadığını git karar verir.
"""
from __future__ import annotations

import argparse
import json
import re
import shutil
from pathlib import Path


def son_donem(html: Path) -> str | None:
    """index.html içindeki DATA'dan solo dönemlerinin sonuncusu (YYYY-AA)."""
    try:
        h = html.read_text(encoding="utf-8")
        m = re.search(r'"periods":\[([^\]]*)\]', h)
        return max(json.loads("[" + m.group(1) + "]")) if m else None
    except Exception:                                                   # noqa: BLE001
        return None


def kopyala(kaynak: Path, hedef: Path):
    if hedef.exists():
        shutil.rmtree(hedef)
    shutil.copytree(kaynak, hedef)


NOT_STIL = ("background:#fff4cc;color:#4a3500;border-bottom:1px solid #e6c55a;padding:8px 16px;"
            "font:500 13.5px/1.45 Roboto,system-ui,sans-serif;text-align:center")


def ana_sayfa_html(oniz_html: str) -> str:
    """Önizleme sayfasını ana adrese uyarlar: dipnotlar tek klasörden, üstteki şerit sade bir not."""
    h = re.sub(r'const url=\(ONIZLEME_DN\.has\(key\)\?"dipnot/":"\.\./dipnot/"\)\+key\+"\.js";',
               'const url="dipnot/"+key+".js";', oniz_html, count=1)
    m = re.search(r'<body><div style="position:sticky[^"]*">ÖNİZLEME — (.*?)</div>', h, re.S)
    if m:
        metin = re.sub(r'\s*<a href="\.\./"[^>]*>Ana rapor</a>', "", m.group(1)).strip()
        metin = metin.replace("; TBB bu dönemi henüz yayımlamadı", "")
        metin = metin.replace(" dönemi KAP ön verisidir:", " dönemi TBB'de henüz yayımlanmadı; bu dönem bankaların KAP raporlarından okunan ön veridir:")
        h = h.replace(m.group(0), f'<body><div role="note" style="{NOT_STIL}"><b>Not:</b> {metin} '
                      "TBB dönemi yayımladığında bu veriler TBB verisiyle değiştirilir.</div>", 1)
    return h


def ana_siteye_yaz(site: Path, oniz: Path | None, docs: Path):
    """docs/ = tam site; önizleme varsa onun dönemi ana sayfaya katılır (dipnotlar üst üste yazılır)."""
    kaynak = oniz if oniz else site
    if oniz:
        (docs / "index.html").write_text(ana_sayfa_html((oniz / "index.html").read_text(encoding="utf-8")), encoding="utf-8")
    else:
        shutil.copy2(site / "index.html", docs / "index.html")
    if (site / "dipnot").exists():
        kopyala(site / "dipnot", docs / "dipnot")              # TBB dönemleri
    if oniz and (oniz / "dipnot").exists():
        for f in (oniz / "dipnot").glob("*.js"):               # KAP döneminin eklendiği paketler (bütün dönemleri içerir)
            shutil.copy2(f, docs / "dipnot" / f.name)
    for alt in ("tcmb", "bddk"):
        if (kaynak / alt).exists():
            kopyala(kaynak / alt, docs / alt)
        elif (site / alt).exists():
            kopyala(site / alt, docs / alt)
    if (docs / "onizleme").exists():                           # ayrı önizleme adresi artık yok
        shutil.rmtree(docs / "onizleme")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--kok", required=True)
    ap.add_argument("--docs", required=True)
    a = ap.parse_args()
    kok, docs = Path(a.kok), Path(a.docs)
    site, oniz = kok / "7_site" / "site", kok / "7_site" / "onizleme"

    if not (site / "index.html").exists():
        print("Tam site yok (7_site/site/index.html); docs/ değiştirilmedi.")
        return
    tbb_son, yayindaki = son_donem(site / "index.html"), son_donem(docs / "index.html")
    if (oniz / "index.html").exists():
        ana_siteye_yaz(site, oniz, docs)
        print(f"Ana site yazıldı: TBB dönemleri {tbb_son}'a kadar + {son_donem(oniz / 'index.html')} KAP ön verisi (notla).")
    elif tbb_son and yayindaki and yayindaki > tbb_son:
        # KAP dönemi yayında ama bu çalıştırmada üretilemedi (ör. KAP'a ulaşılamadı): ana sayfayı koru, yalnız TCMB/BDDK
        for alt in ("tcmb", "bddk"):
            if (site / alt).exists():
                kopyala(site / alt, docs / alt)
        print(f"Önizleme üretilemedi; yayındaki {yayindaki} KAP ön verili ana sayfa korunuyor.")
    else:
        ana_siteye_yaz(site, None, docs)
        print(f"Ana site yazıldı (TBB son dönem {tbb_son}).")
    import sys
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    import giris_sayfasi
    giris_sayfasi.uygula(docs)

if __name__ == "__main__":
    main()
