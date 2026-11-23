#!/usr/bin/env python3
"""
PDF ek veri deposu — TBB'de olmayan, yalnız banka raporlarının (KAP PDF) metninde geçen kalemler.

Taksonomi (taksonomi.json) hangi kalemin hangi cümleden okunacağını tanımlar. Okunan her değer
ek_veri.xlsx'e bir satır olarak yazılır:

    Dönem | Banka | Kapsam | Kod | Kalem | Birim | OKUNAN DEĞER | Kaynak | Okunan cümle |
    GÜNCEL DEĞER (elle) | Açıklama (elle) | GEÇERLİ DEĞER | Durum | Okuma zamanı

- "Okunan değer" programın PDF'ten okuduğudur; program bu sütunu yalnız ilk okumada doldurur.
- "Güncel değer" ve "Açıklama" sizindir: program hiçbir zaman üzerine yazmaz.
- "Geçerli değer" = güncel değer doluysa o, değilse okunan değer. Site ve değerlendirme bunu kullanır.
- Her çalıştırmada elle girilen değişiklikler "Düzeltme günlüğü" sayfasına tarihle eklenir.
Ayrıntı: DUZELTME_PROSEDURU.md

Kullanım (genelde veri_hatti.py çağırır):
    python pdf_ek_veri.py --kok ANA_KLASOR [--donem 2026-09] [--yeniden-oku]
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from datetime import datetime
from pathlib import Path

BURADA = Path(__file__).resolve().parent
sys.path.insert(0, str(BURADA))

PDF_BASLANGIC = "2025-12"

TAKSONOMI = {
    "surum": 1,
    "aciklama": "TBB'de bulunmayan, banka raporlarının dipnot metninden okunan kalemler. Birimler milyon TL ya da %.",
    "kalemler": [
        {"kod": "SK_BAKIYE", "ad": "Serbest karşılık bakiyesi (dönem sonu)", "birim": "mn TL", "grup": "Serbest karşılık",
         "kaynak": "5. Bölüm, Diğer karşılıklar: 'Muhtemel riskler için ayrılan serbest karşılıklara ilişkin bilgiler'",
         "not": "'Bulunmamaktadır' → 0."},
        {"kod": "SK_ONCEKI_YS", "ad": "Serbest karşılık bakiyesi (önceki yıl sonu)", "birim": "mn TL", "grup": "Serbest karşılık",
         "kaynak": "Aynı cümlenin parantez içi '(31 Aralık YYYY: ...)'", "not": ""},
        {"kod": "SK_HAREKET", "ad": "Yıl başından beri serbest karşılık hareketi (+ayrılan / −iptal)", "birim": "mn TL",
         "grup": "Serbest karşılık", "kaynak": "Hesaplanır: SK_BAKIYE − SK_ONCEKI_YS",
         "not": "Gelir tablosuna etkisi; iptal kârı artırır (negatif hareket = gelir)."},
        {"kod": "SK_IPTAL_METIN", "ad": "Metinde belirtilen serbest karşılık iptali (cari dönem)", "birim": "mn TL",
         "grup": "Serbest karşılık", "kaynak": "Denetim raporu / diğer faaliyet gelirleri dipnotu: '... iptal edilen X TL ... serbest karşılık'",
         "not": "SK_HAREKET ile tutarlı olmalı; fark varsa elle kontrol edin."},
        {"kod": "TUFE_VARSAYIM", "ad": "TÜFE'ye endeksli kıymet değerlemesinde kullanılan yıllık enflasyon tahmini", "birim": "%",
         "grup": "TÜFE'ye endeksli kıymetler", "kaynak": "Faiz gelirleri dipnotu / muhasebe politikaları: 'yıllık %X enflasyon tahmini'",
         "not": "Halkbank eğri kullanır; tek oran vermeyebilir."},
        {"kod": "TUFE_DUY_ARTIS", "ad": "TÜFE tahmini +1 puan → vergi öncesi kâr etkisi", "birim": "mn TL",
         "grup": "TÜFE'ye endeksli kıymetler", "kaynak": "'TÜFE tahmininin %1 artması ... vergi öncesi dönem kârı ... artacak'",
         "not": "Halkbank faiz geliri etkisini verir (Kapsam notuna bakın)."},
        {"kod": "TUFE_DUY_AZALIS", "ad": "TÜFE tahmini −1 puan → vergi öncesi kâr etkisi (pozitif yazılır, azalış)", "birim": "mn TL",
         "grup": "TÜFE'ye endeksli kıymetler", "kaynak": "Aynı cümlenin azalış kısmı; tek tutar varsa simetrik kabul edilir", "not": ""},
    ],
}

SUTUNLAR = ["Dönem", "Banka", "Kapsam", "Kod", "Kalem", "Birim", "Okunan değer", "Kaynak", "Okunan cümle",
            "Güncel değer (elle)", "Açıklama (elle)", "Geçerli değer", "Durum", "Okuma zamanı"]
C = {ad: i + 1 for i, ad in enumerate(SUTUNLAR)}       # 1 tabanlı sütun numarası


# ----------------------------------------------------------------------------- sayı okuma
def sayi(s: str, milyar: bool = False) -> float | None:
    """Türkçe/İngilizce karışık biçimler: 1.705 · 797,91 · 8,000 · 2.700 · 29,51"""
    s = s.strip()
    if not re.search(r"\d", s):
        return None
    if milyar:                                            # '1.705 milyar' → 1,705 (ondalık)
        v = float(s.replace(".", "#").replace(",", ".").replace("#", "."))
        return v * 1000
    ayr = re.findall(r"[.,]", s)
    if not ayr:
        return float(s)
    son = max(s.rfind("."), s.rfind(","))
    kuyruk = s[son + 1:]
    if len(set(ayr)) == 1 and len(kuyruk) == 3:          # tek tür ayraç + 3 hane → binlik
        return float(re.sub(r"[.,]", "", s))
    tam = re.sub(r"[.,]", "", s[:son])
    return float(f"{tam}.{kuyruk}")


SAYI = r"(\d{1,3}(?:[.,]\d{3})*(?:[.,]\d{1,2})?|\d+(?:[.,]\d+)?)"


# ----------------------------------------------------------------------------- çıkarıcılar
def _metinler(pdf: Path):
    import pymupdf
    d = pymupdf.open(pdf)
    for i, pg in enumerate(d):
        yield i + 1, re.sub(r"\s+", " ", pg.get_text())
    d.close()


def oku(pdf: Path) -> dict:
    """{kod: (değer, 'sayfa N', cümle, not)}"""
    out: dict = {}
    for s, t in _metinler(pdf):
        # --- serbest karşılık bakiyesi
        if "SK_BAKIYE" not in out:
            m = re.search(r"serbest karşılıklara ilişkin bilgiler\W*(.{0,260})", t, re.I)
            if m:
                cumle = m.group(1)
                par = re.search(r"\(([^)]*?\d{4})\s*:\s*([^)]*)\)", cumle)
                ana = cumle[:par.start()] if par else cumle
                kayit = cumle[:(par.end() if par else 200)]
                if re.search(r"bulunmamaktadır|yoktur", ana, re.I):
                    out["SK_BAKIYE"] = (0.0, f"s.{s}", kayit, "")
                else:
                    mm = re.search(SAYI + r"\s*(?:Milyon\s*)?TL", ana)
                    if mm:
                        out["SK_BAKIYE"] = (sayi(mm.group(1)), f"s.{s}", kayit, "")
                if par:
                    o = par.group(2)
                    if re.search(r"bulunmamaktadır|yoktur", o, re.I):
                        out["SK_ONCEKI_YS"] = (0.0, f"s.{s}", kayit, par.group(1))
                    else:
                        mm = re.search(SAYI, o)
                        if mm:
                            out["SK_ONCEKI_YS"] = (sayi(mm.group(1)), f"s.{s}", kayit, par.group(1))
        # --- metindeki iptal
        if "SK_IPTAL_METIN" not in out:
            m = re.search(r"cari dönemde iptal edilen\s+" + SAYI + r"\s*(?:Milyon\s*)?TL[^.]{0,200}?serbest karşılık", t, re.I)
            if m:
                out["SK_IPTAL_METIN"] = (sayi(m.group(1)), f"s.{s}", m.group(0)[:300], "")
        # --- TÜFE varsayımı
        if "TUFE_VARSAYIM" not in out:
            m = (re.search(r"yıllık\s*%\s*(\d{1,2}(?:[.,]\d{1,2})?)\s*enflasyon tahmin", t, re.I)
                 or re.search(r"enflasyon oranı\s+yıllık\s*%\s*(\d{1,2}(?:[.,]\d{1,2})?)", t, re.I))
            if m:
                v = float(m.group(1).replace(",", "."))
                out["TUFE_VARSAYIM"] = (v, f"s.{s}", t[max(0, m.start() - 120):m.end() + 40], "")
        # --- TÜFE duyarlılığı
        if "TUFE_DUY_ARTIS" not in out:
            m = re.search(r"(TÜFE tahmin\w*|Enflasyon oranlarının|TÜFE’ye endeksli bu kıymetlere)[^%]{0,120}?"
                          r"(%\s*1|100 baz puan)(.{0,320}?(?:aktır|ecektir|maktadır|ilecektir)\.)", t, re.I)
            if m:
                cumle = m.group(0)
                kapsam = "faiz geliri" if re.search(r"faiz gelir", cumle, re.I) and not re.search(r"vergi öncesi", cumle, re.I) \
                    else "vergi öncesi kâr"
                tut = [(x.start(), sayi(x.group(1), milyar=bool(x.group(2))))
                       for x in re.finditer(SAYI + r"\s*(milyar)?\s*(?:\(tam tutar\)\s*)?(?:Milyon\s*)?TL", m.group(3), re.I)]
                if tut:
                    def yon(poz):
                        ar = m.group(3)[poz:poz + 60].lower()
                        a, z = ar.find("art"), ar.find("azal")
                        if a < 0 and z < 0:
                            return None
                        return "azal" if (z >= 0 and (a < 0 or z < a)) else "art"
                    artis = azalis = None
                    for poz, v in tut[:2]:
                        if yon(poz) == "azal" and azalis is None:
                            azalis = v
                        elif artis is None:
                            artis = v
                    if len(tut) == 1:
                        artis = azalis = tut[0][1]
                    not_ = kapsam + ("" if len(tut) > 1 else "; tek tutar, simetrik")
                    out["TUFE_DUY_ARTIS"] = (artis, f"s.{s}", cumle[:420], not_)
                    out["TUFE_DUY_AZALIS"] = (azalis, f"s.{s}", cumle[:420], not_)
    if "SK_BAKIYE" in out and "SK_ONCEKI_YS" in out and out["SK_BAKIYE"][0] is not None and out["SK_ONCEKI_YS"][0] is not None:
        out["SK_HAREKET"] = (out["SK_BAKIYE"][0] - out["SK_ONCEKI_YS"][0], out["SK_BAKIYE"][1], "hesaplandı: bakiye − önceki yıl sonu", "")
    return out


# ----------------------------------------------------------------------------- depo (Excel)
def _stil(ws):
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.worksheet.datavalidation import DataValidation  # noqa: F401
    bas = PatternFill("solid", fgColor="1F3A5F")
    elle = PatternFill("solid", fgColor="FFF4CC")
    for i, ad in enumerate(SUTUNLAR, 1):
        h = ws.cell(1, i)
        h.font = Font(bold=True, color="FFFFFF")
        h.fill = PatternFill("solid", fgColor="B7791F") if "elle" in ad else bas
        h.alignment = Alignment(wrap_text=True, vertical="center")
    for r in range(2, ws.max_row + 1):
        for k in ("Güncel değer (elle)", "Açıklama (elle)"):
            ws.cell(r, C[k]).fill = elle
    gen = {"Dönem": 9, "Banka": 13, "Kapsam": 8, "Kod": 17, "Kalem": 44, "Birim": 7, "Okunan değer": 13, "Kaynak": 18,
           "Okunan cümle": 60, "Güncel değer (elle)": 14, "Açıklama (elle)": 30, "Geçerli değer": 13, "Durum": 16,
           "Okuma zamanı": 17}
    for ad, g in gen.items():
        ws.column_dimensions[ws.cell(1, C[ad]).column_letter].width = g
    ws.freeze_panes = "G2"
    ws.auto_filter.ref = ws.dimensions


class Depo:
    def __init__(self, klasor: Path):
        self.k = Path(klasor)
        self.k.mkdir(parents=True, exist_ok=True)
        self.xlsx = self.k / "ek_veri.xlsx"
        self.json = self.k / "ek_veri.json"
        (self.k / "taksonomi.json").write_text(json.dumps(TAKSONOMI, ensure_ascii=False, indent=2), encoding="utf-8")
        self.satirlar: dict[tuple, dict] = {}
        self.gunluk: list[list] = []
        self._yukle()

    def _yukle(self):
        if not self.xlsx.exists():
            return
        from openpyxl import load_workbook
        wb = load_workbook(self.xlsx)            # formüller değil değerler gerekmez: geçerli değeri biz hesaplıyoruz
        ws = wb["Veri"]
        bas = [c.value for c in ws[1]]
        for row in ws.iter_rows(min_row=2, values_only=True):
            d = dict(zip(bas, row))
            if not d.get("Kod"):
                continue
            self.satirlar[(str(d["Dönem"]), d["Banka"], d["Kapsam"], d["Kod"])] = d
        if "Düzeltme günlüğü" in wb.sheetnames:
            self.gunluk = [list(r) for r in wb["Düzeltme günlüğü"].iter_rows(min_row=2, values_only=True) if any(r)]

    def var_mi(self, donem: str, banka: str, kapsam: str) -> bool:
        return any(k[:3] == (donem, banka, kapsam) for k in self.satirlar)

    def ekle(self, donem, banka, kapsam, okunan: dict, yeniden=False):
        zaman = datetime.now().strftime("%Y-%m-%d %H:%M")
        for k in TAKSONOMI["kalemler"]:
            anahtar = (donem, banka, kapsam, k["kod"])
            eski = self.satirlar.get(anahtar)
            v, kaynak, cumle, not_ = okunan.get(k["kod"], (None, "", "", ""))
            if eski and not yeniden:
                continue
            d = dict(eski or {})
            if eski and yeniden and eski.get("Okunan değer") not in (None, "") and v != eski.get("Okunan değer"):
                self.gunluk.append([zaman, donem, banka, kapsam, k["kod"], "yeniden okuma", eski.get("Okunan değer"), v, "program"])
            d.update({"Dönem": donem, "Banka": banka, "Kapsam": kapsam, "Kod": k["kod"], "Kalem": k["ad"],
                      "Birim": k["birim"], "Okunan değer": v, "Kaynak": kaynak, "Okunan cümle": (cumle or "")[:500],
                      "Okuma zamanı": zaman})
            if not_:
                d["Kaynak"] = f"{kaynak} · {not_}" if kaynak else not_
            d.setdefault("Güncel değer (elle)", None)
            d.setdefault("Açıklama (elle)", None)
            self.satirlar[anahtar] = d

    def kaydet(self):
        from openpyxl import Workbook
        # elle yapılan değişiklikleri günlüğe al (önceki JSON anlık görüntüsüyle karşılaştır)
        once = {}
        if self.json.exists():
            try:
                for r in json.loads(self.json.read_text(encoding="utf-8"))["satirlar"]:
                    once[(r["donem"], r["banka"], r["kapsam"], r["kod"])] = r.get("guncel")
            except Exception:                                   # noqa: BLE001
                pass
        zaman = datetime.now().strftime("%Y-%m-%d %H:%M")
        for a, d in self.satirlar.items():
            g = d.get("Güncel değer (elle)")
            if a in once and once[a] != g:
                self.gunluk.append([zaman, *a, "elle düzeltme", once[a], g, d.get("Açıklama (elle)") or ""])

        wb = Workbook()
        ws = wb.active
        ws.title = "Veri"
        ws.append(SUTUNLAR)
        sira = {k["kod"]: i for i, k in enumerate(TAKSONOMI["kalemler"])}
        for a in sorted(self.satirlar, key=lambda a: (a[0], a[1], a[2], sira.get(a[3], 99))):
            d = self.satirlar[a]
            r = ws.max_row + 1
            for ad in SUTUNLAR:
                if ad in ("Geçerli değer", "Durum"):
                    continue
                ws.cell(r, C[ad], d.get(ad))
            ok, gu = ws.cell(r, C["Okunan değer"]).coordinate, ws.cell(r, C["Güncel değer (elle)"]).coordinate
            ws.cell(r, C["Geçerli değer"], f'=IF({gu}<>"",{gu},{ok})')
            ws.cell(r, C["Durum"], f'=IF({gu}<>"","elle düzeltildi",IF({ok}="","okunamadı — elle girin","okundu"))')
        _stil(ws)
        t = wb.create_sheet("Taksonomi")
        t.append(["Kod", "Kalem", "Birim", "Grup", "Nereden okunur", "Not"])
        for k in TAKSONOMI["kalemler"]:
            t.append([k["kod"], k["ad"], k["birim"], k["grup"], k["kaynak"], k["not"]])
        for col, g in zip("ABCDEF", (17, 50, 8, 24, 60, 50)):
            t.column_dimensions[col].width = g
        g = wb.create_sheet("Düzeltme günlüğü")
        g.append(["Zaman", "Dönem", "Banka", "Kapsam", "Kod", "Tür", "Eski", "Yeni", "Açıklama / kim"])
        for r in self.gunluk:
            g.append(r)
        wb.save(self.xlsx)

        satir = []
        for a, d in self.satirlar.items():
            gu = d.get("Güncel değer (elle)")
            satir.append({"donem": a[0], "banka": a[1], "kapsam": a[2], "kod": a[3], "okunan": d.get("Okunan değer"),
                          "guncel": gu, "gecerli": gu if gu not in (None, "") else d.get("Okunan değer"),
                          "kaynak": d.get("Kaynak"), "aciklama": d.get("Açıklama (elle)")})
        self.json.write_text(json.dumps({"guncelleme": zaman, "birimler": "mn TL / %", "satirlar": satir},
                                        ensure_ascii=False, indent=1), encoding="utf-8")


def guncelle(kok: Path, donemler: list[str] | None = None, yeniden=False, log=print) -> Path:
    """kok/2_kap_raporlari/<dönem>/<banka>/*.pdf → kok/4_pdf_ek_veri/ek_veri.xlsx"""
    import banka_karsilastir as bk
    kap = Path(kok) / "2_kap_raporlari"
    depo = Depo(Path(kok) / "4_pdf_ek_veri")
    tum = sorted(p.name for p in kap.glob("20??-??") if p.is_dir() and p.name >= PDF_BASLANGIC) if kap.exists() else []
    for donem in (donemler or tum):
        for b in bk.BANKALAR:
            k = kap / donem / bk.dosya_adi(b[2])
            if not k.exists():
                continue
            gorulen = set()
            for pdf in sorted(k.glob("*.pdf")):
                tur = bk.pdf_turu(pdf)
                if not tur or tur in gorulen:
                    continue
                gorulen.add(tur)
                if depo.var_mi(donem, b[2], tur) and not yeniden:
                    continue
                try:
                    o = oku(pdf)
                    depo.ekle(donem, b[2], tur, o, yeniden)
                    log(f"  ek veri {donem} {b[2]} {tur}: {len(o)} kalem okundu")
                except Exception as e:                          # noqa: BLE001
                    log(f"  ! ek veri {donem} {b[2]} {tur}: {e}")
    depo.kaydet()
    return depo.xlsx


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--kok", required=True, help="ana veri klasörü (veri_hatti.py --kok ile aynı)")
    ap.add_argument("--donem", action="append", help="YYYY-AA (birden çok verilebilir); yoksa hepsi")
    ap.add_argument("--yeniden-oku", action="store_true", help="okunan değerleri PDF'ten yeniden oku (elle sütunlar korunur)")
    a = ap.parse_args()
    print(guncelle(Path(a.kok), a.donem, a.yeniden_oku))
