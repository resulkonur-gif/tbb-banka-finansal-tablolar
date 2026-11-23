#!/usr/bin/env python3
"""
BDDK sayfalarını siteye taşır: site/bddk/index.html (aylık bülten) ve site/bddk/haftalik.html.

Kaynak, ayrı tutulan bddk-veri deposunun yayımlanmış raporlarıdır (her gün kendi iş akışıyla güncellenir).
Veriler TBB verisine karıştırılmaz; yalnız aynı sitede, kendi klasöründe gösterilir. Sayfaların başına
"← Başlangıç ekranı" bağlantısı eklenir, aylık ↔ haftalık bağlantıları olduğu gibi çalışır.

    python bddk_sayfa.py --site docs/onizleme [--kaynak /yol/bddk-veri/docs]
--kaynak verilmezse GitHub'dan (raw) indirilir. Tekrar çalıştırmak güvenlidir.
"""
from __future__ import annotations

import argparse
import urllib.request
from pathlib import Path

HAM = "https://raw.githubusercontent.com/resulkonur-gif/bddk-veri/main/docs/"
DOSYALAR = ("index.html", "haftalik.html")
GERI = ('<!--GERI--><div class="gut" style="max-width:920px;margin:0 auto;padding-top:18px">'
        '<a href="../" style="color:var(--m-primary,#2b5aa6);font:500 14px/1.4 Roboto,system-ui,sans-serif;'
        'text-decoration:none">← Başlangıç ekranı</a></div><!--/GERI-->')


def _oku(ad: str, kaynak: Path | None) -> str:
    if kaynak:
        return (Path(kaynak) / ad).read_text(encoding="utf-8")
    req = urllib.request.Request(HAM + ad, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=120) as r:
        return r.read().decode("utf-8")


def uygula(site: Path, kaynak: Path | None = None, log=print) -> Path | None:
    hedef = Path(site) / "bddk"
    hedef.mkdir(parents=True, exist_ok=True)
    for ad in DOSYALAR:
        try:
            h = _oku(ad, kaynak)
        except Exception as e:                                          # noqa: BLE001
            log(f"  ! BDDK {ad} alınamadı ({e}); varsa eski kopya kalır")
            continue
        if '<div id="landing"' not in h:
            log(f"  ! BDDK {ad}: beklenen yapı yok, kopyalanmadı")
            continue
        h = h.replace('<div id="landing"', GERI + '\n<div id="landing"', 1)
        (hedef / ad).write_text(h, encoding="utf-8")
        log(f"  BDDK {ad}: {len(h) / 1e6:.1f} MB → {hedef / ad}")
    return hedef if (hedef / "index.html").exists() else None


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--site", required=True)
    ap.add_argument("--kaynak", help="bddk-veri/docs klasörü (yoksa GitHub'dan indirilir)")
    a = ap.parse_args()
    uygula(Path(a.site), Path(a.kaynak) if a.kaynak else None)
