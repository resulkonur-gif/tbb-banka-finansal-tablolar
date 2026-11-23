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
    shutil.copy2(site / "index.html", docs / "index.html")
    for alt in ("dipnot", "tcmb", "bddk"):
        if (site / alt).exists():
            kopyala(site / alt, docs / alt)
    print(f"Ana site yazıldı (TBB son dönem {son_donem(site / 'index.html')}).")

    if (oniz / "index.html").exists():
        kopyala(oniz, docs / "onizleme")
        print(f"Önizleme yazıldı ({son_donem(oniz / 'index.html')}).")
    elif (docs / "onizleme" / "index.html").exists():
        tbb, onz = son_donem(site / "index.html"), son_donem(docs / "onizleme" / "index.html")
        if tbb and onz and tbb >= onz:
            shutil.rmtree(docs / "onizleme")
            print(f"TBB {onz} dönemini yayımladı; önizleme kaldırıldı.")
        else:
            print(f"Önizleme üretilemedi; mevcut önizleme ({onz}) korunuyor.")
    # giriş ekranı: önizleme duyurusu docs/onizleme'nin son hâline göre (varsa göster, yoksa kaldır)
    import sys
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    import giris_sayfasi
    giris_sayfasi.uygula(docs)


if __name__ == "__main__":
    main()
