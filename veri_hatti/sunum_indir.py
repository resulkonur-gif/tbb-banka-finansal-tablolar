#!/usr/bin/env python3
"""
Yatırımcı sunumu indirici — 7 bankanın çeyreklik sonuç (earnings) sunumları.

Her banka için iki yol denenir:
  1) "sablon": bilinen dosya adı kalıbı (ör. Garanti 3Q26_BRSA_Consolidated_Earnings_Presentation.pdf)
  2) "sayfa": bankanın yatırımcı ilişkileri sayfası taranır; dönem etiketi (3Q26, Q3 2026, 2026_q3,
     9M26, Eylül 2026 ...) ve "presentation / sunum" geçen PDF bağlantısı seçilir. Sayfa JavaScript ile
     doluyorsa playwright (chromium) ile açılır.
Kaynak adresleri kaynaklar.json'dadır (ilk çalıştırmada yazılır); banka sitesini değiştirirse kodu değil
o dosyayı düzeltin.

İndirilen her PDF'in yanına metni .txt olarak çıkarılır (değerlendirme bunları okur).

Kullanım (genelde veri_hatti.py çağırır):
    python sunum_indir.py --kok ANA_KLASOR [--donem 2026-09]
Klasör: ANA_KLASOR/3_yatirimci_sunumlari/<YYYY-AA>/<Banka>/<dosya>.pdf (+ .txt)
"""
from __future__ import annotations

import argparse
import html
import json
import re
import sys
from pathlib import Path
from urllib.parse import urljoin

def tarayici_ac(p, **kw):
    """Playwright'ın Chromium'u; yoksa yüklü Edge ya da Chrome. Her aday bir sekme açılarak sınanır.
    Şirket politikası olan Edge'de normal açılış sekme açamayabilir; o zaman ayrı (boş) profille açılır."""
    import tempfile
    ek = ["--no-first-run", "--no-default-browser-check", "--disable-extensions", "--disable-features=msEdgeStartupBoost"]
    hatalar = []

    class _Kalici:                                          # launch_persistent_context'i Browser gibi kullanmak için
        def __init__(self, ctx):
            self._ctx = ctx
        def new_page(self, **_):
            return self._ctx.new_page()
        def close(self):
            self._ctx.close()

    adaylar = [(None, False), ("msedge", False), ("msedge", True), ("chrome", False), ("chrome", True)]
    for kanal, kalici in adaylar:
        b = None
        try:
            ayar = dict(kw, args=ek, **({"channel": kanal} if kanal else {}))
            if kalici:
                ctx = p.chromium.launch_persistent_context(tempfile.mkdtemp(prefix="bv_tarayici_"), **ayar)
                for s in list(ctx.pages):
                    s.close()
                b = _Kalici(ctx)
            else:
                b = p.chromium.launch(**ayar)
            s = b.new_page()
            s.close()
            return b
        except Exception as e:                              # noqa: BLE001
            hatalar.append(f"{kanal or 'chromium'}{' (ayrı profil)' if kalici else ''}: {str(e).splitlines()[0][:100]}")
            try:
                if b:
                    b.close()
            except Exception:                               # noqa: BLE001
                pass
    raise RuntimeError("Tarayıcı açılamadı (" + "; ".join(hatalar) + ")")


BURADA = Path(__file__).resolve().parent
sys.path.insert(0, str(BURADA))
PDF_BASLANGIC = "2025-12"
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36",
      "Accept-Language": "tr-TR,tr;q=0.9,en;q=0.8"}

# {yil} 2026 · {yy} 26 · {c} çeyrek no (1-4) · {ay} 09
VARSAYILAN_KAYNAKLAR = {
    "_aciklama": "Banka başına 'sablon' (doğrudan PDF adresi kalıbı) ve 'sayfa' (taranacak IR sayfaları). "
                 "dogrulanmis=şablon adresi gerçek bir dosyada görüldü; sayfa adresleri ve diğer bankalar tahmindir, ilk çalıştırma günlüğüne bakın.",
    "Ziraat": {"sablon": ["https://www.ziraatbank.com.tr/en/Investor-Relations-ZB/Financials/Documents/{yil}_q{c}_financial_presentation.pdf"],
               "sayfa": ["https://www.ziraatbank.com.tr/en/investor-relations/financials/presentations/investor-presentation"],
               "dogrulanmis": True},
    "İş Bankası": {"sablon": ["https://www.isbank.com.tr/contentmanagement/IsbankInvestorDocuments/pdf/YatirimciSunumu{c}Q{yil}.pdf"],
                   "sayfa": ["https://www.isbank.com.tr/en/investor-relations"], "dogrulanmis": True},
    "Garanti BBVA": {"sablon": ["https://garantibbvainvestorrelations.com/en/images/pdf/{c}Q{yy}_BRSA_Consolidated_Earnings_Presentation.pdf"],
                     "sayfa": ["https://garantibbvainvestorrelations.com/en/financials/presentations"], "dogrulanmis": True},
    "Yapı Kredi": {"sablon": [], "sayfa": ["https://www.yapikredi.com.tr/en/investor-relations/financials"], "dogrulanmis": False},
    "Halkbank": {"sablon": [], "sayfa": [
        "https://www.halkbank.com.tr/en/investor-relations/financial-information/presentations/earnings-presentations"],
        "dogrulanmis": False},
    "Akbank": {"sablon": [], "sayfa": ["https://www.akbankinvestorrelations.com/en/presentations",
                                       "https://www.akbankinvestorrelations.com/en/"], "dogrulanmis": False},
    "VakıfBank": {"sablon": [], "sayfa": ["https://www.vakifbank.com.tr/en/investor-relations/presentations",
                                          "https://www.vakifbank.com.tr/en/investor-relations"], "dogrulanmis": False},
}

AY_EN = {3: "March", 6: "June", 9: "September", 12: "December"}
AY_TR = {3: "Mart", 6: "Haziran", 9: "Eyl[uü]l", 12: "Aral[ıi]k"}


def donem_kaliplari(yil: int, ay: int) -> list[str]:
    """Bağlantı adresinde/metninde dönemi tanıyan düzenli ifadeler (büyük-küçük harf duyarsız)."""
    c, yy = ay // 3, yil % 100
    ay_say = {3: "3M", 6: "6M|1H|H1", 9: "9M", 12: "12M|FY|4Q"}[ay]
    k = [rf"(?<![0-9]){c}Q[ _-]?(?:20)?{yy}(?![0-9])", rf"Q{c}[ _-]?(?:20)?{yy}(?![0-9])", rf"{yil}[ _-]?Q{c}(?![0-9])",
         rf"{yil}[ _/-]?{c}Q", rf"(?:{ay_say})[ _-]?(?:20)?{yy}(?![0-9])", rf"{AY_EN[ay]}[ _-]?{yil}", rf"{AY_TR[ay]}[ _-]?{yil}",
         rf"{yil}[ _-]?(?:0)?{ay}(?![0-9])", rf"(?<![0-9])(?:0)?{ay}[ _.-]{yil}"]
    if ay == 12:
        k += [rf"FY[ _-]?(?:20)?{yy}(?![0-9])", rf"{yil}[ _-]?(?:Annual|Year[ _-]?End)"]
    return k


def baglantilar(sayfa_html: str, taban: str) -> list[tuple[str, str]]:
    out = []
    for m in re.finditer(r"<a\b[^>]*?href\s*=\s*[\"']([^\"']+)[\"'][^>]*>(.*?)</a>", sayfa_html, re.I | re.S):
        href, metin = html.unescape(m.group(1)), re.sub(r"<[^>]+>", " ", m.group(2))
        out.append((urljoin(taban, href), re.sub(r"\s+", " ", html.unescape(metin)).strip()))
    for m in re.finditer(r"[\"']([^\"'\s]+\.pdf(?:\?[^\"']*)?)[\"']", sayfa_html, re.I):   # JSON / data-* içindekiler
        out.append((urljoin(taban, html.unescape(m.group(1))), ""))
    return out


def sunum_sec(baglar: list[tuple[str, str]], yil: int, ay: int) -> list[str]:
    kal = [re.compile(k, re.I) for k in donem_kaliplari(yil, ay)]
    puanli = []
    for url, metin in baglar:
        hedef = f"{url} {metin}"
        if ".pdf" not in url.lower() and "download" not in url.lower():
            continue
        if not any(k.search(hedef) for k in kal):
            continue
        p = 0
        if re.search(r"presentation|sunum|earnings|webcast|results", hedef, re.I):
            p += 2
        if re.search(r"bddk|brsa|consolidated|konsolide", hedef, re.I):
            p += 1
        if re.search(r"financial[ _-]?statements|faaliyet|annual[ _-]?report|bulletin|bülten|transcript|press", hedef, re.I):
            p -= 3
        if p > 0:
            puanli.append((p, url))
    gor, out = set(), []
    for _, u in sorted(puanli, key=lambda x: -x[0]):
        if u not in gor:
            gor.add(u)
            out.append(u)
    return out


class Getirici:
    def __init__(self, log=print):
        import requests
        self.s = requests.Session()
        self.s.headers.update(UA)
        self.log = log
        self._pw = None

    def metin(self, url: str) -> str:
        r = self.s.get(url, timeout=40)
        r.raise_for_status()
        return r.text

    def tarayici(self, url: str) -> str:
        try:
            from playwright.sync_api import sync_playwright
        except Exception:                                   # noqa: BLE001
            return ""
        with sync_playwright() as p:
            b = tarayici_ac(p)
            sf = b.new_page(user_agent=UA["User-Agent"])
            sf.goto(url, wait_until="networkidle", timeout=60000)
            h = sf.content()
            b.close()
            return h

    def pdf(self, url: str) -> bytes | None:
        try:
            r = self.s.get(url, timeout=90)
        except Exception:                                   # noqa: BLE001
            return None
        if r.status_code != 200 or r.content[:4] != b"%PDF":
            return None
        return r.content


def metne_cevir(pdf: Path):
    try:
        import pymupdf
        d = pymupdf.open(pdf)
        pdf.with_suffix(".txt").write_text("\n\f\n".join(f"[s.{i + 1}]\n{p.get_text()}" for i, p in enumerate(d)),
                                           encoding="utf-8")
        d.close()
    except Exception:                                       # noqa: BLE001
        pass


def kaynaklari_oku(klasor: Path) -> dict:
    yol = klasor / "kaynaklar.json"
    if not yol.exists():
        klasor.mkdir(parents=True, exist_ok=True)
        yol.write_text(json.dumps(VARSAYILAN_KAYNAKLAR, ensure_ascii=False, indent=2), encoding="utf-8")
        return VARSAYILAN_KAYNAKLAR
    return json.loads(yol.read_text(encoding="utf-8"))


def banka_indir(g: Getirici, banka: str, kay: dict, yil: int, ay: int, hedef: Path) -> Path | None:
    c, yy = ay // 3, yil % 100
    adaylar = [s.format(yil=yil, yy=f"{yy:02d}", c=c, ay=f"{ay:02d}") for s in kay.get("sablon", [])]
    for url in adaylar:
        icerik = g.pdf(url)
        if icerik:
            return _kaydet(hedef, url, icerik)
    for sayfa in kay.get("sayfa", []):
        for oku in (g.metin, g.tarayici):
            try:
                h = oku(sayfa)
            except Exception as e:                          # noqa: BLE001
                g.log(f"    {banka}: {sayfa} açılamadı ({type(e).__name__})")
                continue
            secilen = sunum_sec(baglantilar(h, sayfa), yil, ay)
            for url in secilen[:3]:
                icerik = g.pdf(url)
                if icerik:
                    return _kaydet(hedef, url, icerik)
            if secilen:
                break
    return None


def _kaydet(hedef: Path, url: str, icerik: bytes) -> Path:
    hedef.mkdir(parents=True, exist_ok=True)
    ad = re.sub(r'[\\/:*?"<>|]+', "_", url.split("?")[0].rstrip("/").split("/")[-1]) or "sunum.pdf"
    if not ad.lower().endswith(".pdf"):
        ad += ".pdf"
    yol = hedef / ad
    yol.write_bytes(icerik)
    (hedef / "kaynak.txt").write_text(url + "\n", encoding="utf-8")
    metne_cevir(yol)
    return yol


def ceyrekler(bas: str, son: str) -> list[str]:
    y, a = int(bas[:4]), int(bas[5:])
    out = []
    while f"{y}-{a:02d}" <= son:
        out.append(f"{y}-{a:02d}")
        a += 3
        if a > 12:
            y, a = y + 1, 3
    return out


def guncelle(kok: Path, son_donem: str, donemler: list[str] | None = None, log=print) -> dict:
    """Eksik (dönem, banka) sunumlarını indirir. Dönen: {dönem: {banka: dosya adı | None}}"""
    klasor = Path(kok) / "3_yatirimci_sunumlari"
    kaynak = kaynaklari_oku(klasor)
    g = Getirici(log)
    sonuc: dict = {}
    kayit = klasor / "bulunamadi.json"                  # art arda bulunamayanlar: eski dönemlerde tekrar denenmez
    try:
        bulunamadi = json.loads(kayit.read_text(encoding="utf-8"))
    except Exception:                                   # noqa: BLE001
        bulunamadi = {}
    tum = donemler or ceyrekler(PDF_BASLANGIC, son_donem)
    taze = set(tum[-2:])                                # son iki dönem her çalıştırmada denenir
    for donem in tum:
        yil, ay = int(donem[:4]), int(donem[5:])
        for banka, kay in kaynak.items():
            if banka.startswith("_"):
                continue
            hedef = klasor / donem / re.sub(r'[\\/:*?"<>|]+', "_", banka)
            var = sorted(hedef.glob("*.pdf")) if hedef.exists() else []
            if var:
                sonuc.setdefault(donem, {})[banka] = var[0].name
                continue
            anahtar = f"{donem}/{banka}"
            if donem not in taze and bulunamadi.get(anahtar, 0) >= 2:
                sonuc.setdefault(donem, {})[banka] = None
                continue
            try:
                p = banka_indir(g, banka, kay, yil, ay, hedef)
            except Exception as e:                          # noqa: BLE001
                log(f"  ! sunum {donem} {banka}: {e}")
                p = None
            sonuc.setdefault(donem, {})[banka] = p.name if p else None
            if p:
                bulunamadi.pop(anahtar, None)
            else:
                bulunamadi[anahtar] = bulunamadi.get(anahtar, 0) + 1
            log(f"  sunum {donem} {banka}: {p.name if p else 'bulunamadı (yayımlanmamış olabilir ya da adres değişmiş)'}")
    kayit.write_text(json.dumps(bulunamadi, ensure_ascii=False, indent=1), encoding="utf-8")
    return sonuc


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--kok", required=True)
    ap.add_argument("--donem", action="append", help="YYYY-AA; verilmezse 2025-12'den son çeyreğe kadar eksikler")
    ap.add_argument("--son", default="", help="son dönem YYYY-AA (varsayılan: bugünden son çeyrek)")
    a = ap.parse_args()
    if not a.son:
        from datetime import date
        t = date.today()
        q = ((t.month - 1) // 3) * 3
        a.son = f"{t.year - 1}-12" if q == 0 else f"{t.year}-{q:02d}"
    print(json.dumps(guncelle(Path(a.kok), a.son, a.donem), ensure_ascii=False, indent=1))
