#!/usr/bin/env python3
"""
BDDK sayfalarını siteye taşır:
    site/bddk/index.html     seçim ekranı: Aylık bülten mi, Haftalık bülten mi?
    site/bddk/aylik.html     BDDK Aylık Bülten raporu
    site/bddk/haftalik.html  BDDK Haftalık Bülten raporu

Kaynak, ayrı tutulan bddk-veri deposunun yayımlanmış raporlarıdır (her gün kendi iş akışıyla güncellenir).
Veriler TBB verisine karıştırılmaz; yalnız aynı sitede, kendi klasöründe gösterilir. Rapor sayfalarının başına
"← Başlangıç ekranı · BDDK veri seçimi" bağlantısı eklenir; aylık ↔ haftalık bağlantıları çalışır.

    python bddk_sayfa.py --site docs/onizleme [--kaynak /yol/bddk-veri/docs]
--kaynak verilmezse GitHub'dan (raw) indirilir. Tekrar çalıştırmak güvenlidir.
"""
from __future__ import annotations

import argparse
import html as H
import re
import urllib.request
from pathlib import Path

HAM = "https://raw.githubusercontent.com/resulkonur-gif/bddk-veri/main/docs/"
KAYNAK_HEDEF = (("index.html", "aylik.html"), ("haftalik.html", "haftalik.html"))
GERI = ('<!--GERI--><div class="gut" style="max-width:920px;margin:0 auto;padding-top:18px;'
        'font:500 14px/1.4 Roboto,system-ui,sans-serif">'
        '<a href="../" style="color:var(--m-primary,#2b5aa6);text-decoration:none">← Başlangıç ekranı</a>'
        '<span style="color:var(--m-on-surface-variant,#777);margin:0 8px">·</span>'
        '<a href="./" style="color:var(--m-primary,#2b5aa6);text-decoration:none">BDDK: aylık / haftalık seçimi</a>'
        '</div><!--/GERI-->')
OK = '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 4l-1.41 1.41L16.17 11H4v2h12.17l-5.58 5.59L12 20l8-8z"/></svg>'


def _oku(ad: str, kaynak: Path | None) -> str:
    if kaynak:
        return (Path(kaynak) / ad).read_text(encoding="utf-8")
    req = urllib.request.Request(HAM + ad, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=120) as r:
        return r.read().decode("utf-8")


def _bilgi(h: str) -> dict:
    """Rapor giriş ekranından kapsam satırı ve tanıtım paragrafı."""
    k = re.search(r'<span class="land-kicker">[^<]*?·\s*([^<]+)</span>', h)
    i = h.find('<div class="land fade-in">')
    p = re.search(r"<p[^>]*>(.*?)</p>", h[i:i + 4000], re.S) if i >= 0 else None
    duz = lambda s: re.sub(r"\s+", " ", H.unescape(re.sub(r"<[^>]+>", "", s))).strip()   # noqa: E731
    return {"kapsam": k.group(1).strip() if k else "", "metin": duz(p.group(1)) if p else ""}


def secim_sayfasi(site: Path, aylik: dict, haftalik: dict) -> str:
    """Sitenin kendi görünümüyle (giriş ekranı sınıfları) iki kartlı seçim sayfası."""
    bas = ""
    ana = Path(site) / "index.html"
    if ana.exists():
        h = ana.read_text(encoding="utf-8")
        bas = h[:h.index("</head>")]
        bas = bas[bas.index("<link"):] if "<link" in bas else bas
        bas = re.sub(r"<script\b.*?</script>", "", bas, flags=re.S)
    kart = lambda cls, href, ikon, baslik, metin, cips, git: (                                   # noqa: E731
        f'<a class="choice{cls}" href="{href}"><span class="ch-ico">{ikon}</span><span class="ch-t">{baslik}</span>'
        f'<span class="ch-d">{H.escape(metin)}</span><span class="chips">'
        + "".join(f'<span class="chip">{H.escape(c)}</span>' for c in cips if c)
        + f'</span><span class="ch-go">{git} {OK}</span></a>')
    takvim = '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M19 4h-1V2h-2v2H8V2H6v2H5a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2V6a2 2 0 0 0-2-2zm0 16H5V9h14v11zM7 11h5v5H7z"/></svg>'
    hafta = '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M3.5 18.49 9.5 12.48l4 4L22 6.92 20.59 5.5l-7.09 7.97-4-4L2 16.99z"/></svg>'
    return f"""<!doctype html>
<html lang="tr"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>BDDK verileri · aylık ve haftalık bülten</title>
{bas}
<style>a.choice{{color:inherit;text-decoration:none}} .bd-top{{margin:-24px 0 8px}}
.bd-top a{{color:var(--m-primary);font-weight:500;font-size:14px;text-decoration:none}} .bd-top a:hover{{text-decoration:underline}}</style>
</head><body>
<div id="landing" class="gut" lang="tr">
  <div class="land fade-in">
    <div class="bd-top"><a href="../">← Başlangıç ekranı</a></div>
    <span class="land-kicker">BDDK · Bankacılık Düzenleme ve Denetleme Kurumu bültenleri</span>
    <h1>BDDK verileri<br><span class="land-h1b">Aylık mı, haftalık mı?</span></h1>
    <p class="land-sub">BDDK sektörü ve banka gruplarını iki bültenle yayımlar. Aylık bülten bilanço, kâr-zarar, sermaye yeterliliği ve rasyolar dahil 17 tablonun tamamıdır; haftalık bülten kredi, mevduat, takipteki alacaklar ve döviz pozisyonunu her cuma, daha az kalemle ama daha sık gösterir. Veriler TBB'nin banka banka finansal tablolarından ayrıdır.</p>
    <p class="land-q">Hangi bülteni açalım?</p>
    <div class="choices">
      {kart("", "aylik.html", takvim, "Aylık bülten", "17 tablonun tamamı: bilanço, kâr-zarar, krediler, mevduat, menkul kıymetler, likidite, sermaye yeterliliği ve rasyolar. Sektör ve 10 banka grubu, her ay; dönem karşılaştırması ve uzun dönem grafikleri.", ["Aylık · " + aylik["kapsam"] if aylik.get("kapsam") else "", "17 tablo", "10 banka grubu"], "Aylık bülteni aç")}
      {kart(" k", "haftalik.html", hafta, "Haftalık bülten", "9 tablo, her cuma: krediler, takipteki alacaklar, menkul değerler, mevduat ve döviz pozisyonu. TL ve yabancı para ayrımıyla haftalık, aylık ve yıllık değişimler.", ["Haftalık · " + haftalik["kapsam"] if haftalik.get("kapsam") else "", "9 tablo", "TL ve YP ayrımı"], "Haftalık bülteni aç")}
    </div>
    <p class="land-foot">Kaynak: BDDK Aylık ve Haftalık Bülten. Tutarlar BDDK'nın yayımladığı hâliyle, nominaldir.</p>
  </div>
</div>
</body></html>
"""


def uygula(site: Path, kaynak: Path | None = None, log=print) -> Path | None:
    hedef = Path(site) / "bddk"
    hedef.mkdir(parents=True, exist_ok=True)
    bilgi = {}
    for kaynak_ad, hedef_ad in KAYNAK_HEDEF:
        try:
            h = _oku(kaynak_ad, kaynak)
        except Exception as e:                                          # noqa: BLE001
            log(f"  ! BDDK {kaynak_ad} alınamadı ({e}); varsa eski kopya kalır")
            if (hedef / hedef_ad).exists():
                bilgi[hedef_ad] = _bilgi((hedef / hedef_ad).read_text(encoding="utf-8"))
            continue
        if '<div id="landing"' not in h:
            log(f"  ! BDDK {kaynak_ad}: beklenen yapı yok, kopyalanmadı")
            continue
        bilgi[hedef_ad] = _bilgi(h)
        h = h.replace('href="index.html"', 'href="aylik.html"')         # haftalık → aylık bağlantısı
        h = h.replace('<div id="landing"', GERI + '\n<div id="landing"', 1)
        (hedef / hedef_ad).write_text(h, encoding="utf-8")
        log(f"  BDDK {hedef_ad}: {len(h) / 1e6:.1f} MB")
    if not (hedef / "aylik.html").exists():
        return None
    (hedef / "index.html").write_text(secim_sayfasi(site, bilgi.get("aylik.html", {}), bilgi.get("haftalik.html", {})),
                                      encoding="utf-8")
    log(f"  BDDK seçim ekranı: {hedef / 'index.html'}")
    return hedef


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--site", required=True)
    ap.add_argument("--kaynak", help="bddk-veri/docs klasörü (yoksa GitHub'dan indirilir)")
    a = ap.parse_args()
    uygula(Path(a.site), Path(a.kaynak) if a.kaynak else None)
