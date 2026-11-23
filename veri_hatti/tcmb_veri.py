#!/usr/bin/env python3
"""
TCMB EVDS arşivi — günlük ve haftalık piyasa verileri (aylık makro seriler makro_pano.py / makro_veri.json'da).

Arşiv klasörü (varsayılan: bu dosyanın yanında tcmb_arsiv/, veri_hatti.py ile ANA_KLASOR/5_makro/tcmb/):
    gunluk.csv      tarih,kod,deger         (iş günü; EVDS'nin yayımladığı hâliyle)
    haftalik.csv    tarih,kod,deger         (cuma)
    seriler.json    seri tanımları (kod, ad, birim, kaynak)
    son_tarih.txt   arşivdeki en son gün
    TCMB_Arsiv.xlsx Bilgi + Günlük + Haftalık + Aylık sayfaları (geniş biçim, Excel'de açmak için)

Çalışma
    İlk çalıştırma: 2018-01-01'den bugüne her seri (2 yıllık parçalar hâlinde).
    Sonrakiler: her seri için arşivdeki son tarihten 30 gün öncesinden bugüne (EVDS geriye dönük düzeltmeleri için).

    python tcmb_veri.py [--arsiv KLASOR] [--anahtar EVDS_API_KEY] [--bastan]
Anahtar yoksa EVDS_API_KEY ortam değişkeni kullanılır; o da yoksa arşiv olduğu gibi kalır.
"""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import json
import os
import re
import sys
import urllib.parse
import urllib.request
from pathlib import Path

BURADA = Path(__file__).resolve().parent
BASLANGIC = "2018-01-01"
EVDS_ADRES = ["https://evds3.tcmb.gov.tr/igmevdsms-dis/", "https://evds2.tcmb.gov.tr/service/evds/"]

# kod, ad, birim, grup, frekans parametresi (None = serinin kendi frekansı; 2 = iş günü)
GUNLUK = [
    ("TP.DK.USD.A.YTL", "ABD doları (döviz alış)", "TL", "Kurlar", None),
    ("TP.DK.EUR.A.YTL", "Euro (döviz alış)", "TL", "Kurlar", None),
    ("TP.APIFON4", "TCMB ağırlıklı ortalama fonlama maliyeti", "%", "Faizler", 2),
    ("TP.BISTTLREF.ORAN", "BIST TLREF (gecelik referans faiz)", "%", "Faizler", 2),
    ("TP.MK.F.BILESIK", "BIST 100 endeksi (kapanış)", "endeks", "Borsa", 2),
    ("TP.MK.F.MALI", "BIST Mali endeks (kapanış)", "endeks", "Borsa", 2),
    ("TP.AB.A02", "Analitik bilanço: dış varlıklar", "bin TL", "Rezervler", 2),
    ("TP.AB.A10", "Analitik bilanço: toplam döviz yükümlülükleri", "bin TL", "Rezervler", 2),
]
HAFTALIK = [
    ("TP.AB.TOPLAM", "TCMB brüt rezervleri (altın + döviz)", "milyon $", "Rezervler", None),
    ("TP.AB.C1", "TCMB altın rezervi", "milyon $", "Rezervler", None),
    ("TP.AB.C2", "TCMB döviz rezervi", "milyon $", "Rezervler", None),
    ("TP.MKNETHAR.M7", "Yurt dışı yerleşikler: hisse senedi net alımı", "milyon $", "Portföy akımları", None),
    ("TP.MKNETHAR.M8", "Yurt dışı yerleşikler: DİBS net alımı", "milyon $", "Portföy akımları", None),
]


def _tarih(t) -> str | None:
    t = str(t)
    m = re.match(r"^(\d{1,2})-(\d{1,2})-(\d{4})", t)
    if m:
        return f"{m.group(3)}-{int(m.group(2)):02d}-{int(m.group(1)):02d}"
    m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})", t)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    return None


def evds(anahtar: str, kod: str, bas: str, son: str, frek=None) -> dict:
    """{YYYY-AA-GG: değer} — bas/son YYYY-AA-GG."""
    gg = lambda s: f"{s[8:10]}-{s[5:7]}-{s[:4]}"                       # noqa: E731
    q = {"series": kod, "startDate": gg(bas), "endDate": gg(son), "type": "json"}
    if frek:
        q["frequency"] = str(frek)
    yol = "&".join(f"{k}={urllib.parse.quote(str(v), safe='.-_')}" for k, v in q.items())
    son_hata = None
    for kok in EVDS_ADRES:
        try:
            req = urllib.request.Request(kok + yol, headers={"key": anahtar, "User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=90) as r:
                j = json.loads(r.read().decode("utf-8"))
            sutun = kod.replace(".", "_")
            out = {}
            for it in j.get("items", []):
                t = _tarih(it.get("Tarih"))
                for k, v in it.items():
                    if k.startswith(sutun) and v not in (None, "", "null"):
                        if t:
                            out[t] = float(v)
                        break
            return out
        except Exception as e:                                          # noqa: BLE001
            son_hata = e
    raise RuntimeError(f"{kod}: {son_hata}")


class Arsiv:
    def __init__(self, klasor: Path):
        self.k = Path(klasor)
        self.k.mkdir(parents=True, exist_ok=True)
        self.v = {"gunluk": self._oku("gunluk"), "haftalik": self._oku("haftalik")}

    def _oku(self, ad) -> dict:
        yol = self.k / f"{ad}.csv"
        out: dict = {}
        if yol.exists():
            with yol.open(encoding="utf-8-sig") as f:
                for r in csv.DictReader(f):
                    out.setdefault(r["kod"], {})[r["tarih"]] = float(r["deger"])
        return out

    def son(self, frek, kod) -> str | None:
        s = self.v[frek].get(kod)
        return max(s) if s else None

    def ekle(self, frek, kod, veri: dict) -> int:
        s = self.v[frek].setdefault(kod, {})
        n = sum(1 for t, x in veri.items() if s.get(t) != x)
        s.update(veri)
        return n

    def kaydet(self):
        for frek in ("gunluk", "haftalik"):
            with (self.k / f"{frek}.csv").open("w", encoding="utf-8-sig", newline="") as f:
                w = csv.writer(f)
                w.writerow(["tarih", "kod", "deger"])
                for kod in sorted(self.v[frek]):
                    for t in sorted(self.v[frek][kod]):
                        w.writerow([t, kod, repr(self.v[frek][kod][t])])
        tanim = {"gunluk": [dict(kod=k, ad=a, birim=b, grup=g) for k, a, b, g, _ in GUNLUK],
                 "haftalik": [dict(kod=k, ad=a, birim=b, grup=g) for k, a, b, g, _ in HAFTALIK],
                 "kaynak": "TCMB EVDS (evds3.tcmb.gov.tr)"}
        (self.k / "seriler.json").write_text(json.dumps(tanim, ensure_ascii=False, indent=1), encoding="utf-8")
        tum = [t for s in self.v["gunluk"].values() for t in s]
        if tum:
            (self.k / "son_tarih.txt").write_text(max(tum) + "\n", encoding="utf-8")

    def excel(self, makro_veri: Path | None = None):
        """Geniş biçim Excel: her seri bir sütun (BDDK arşivindeki gibi bilgi sayfasıyla)."""
        try:
            from openpyxl import Workbook
            from openpyxl.styles import Font, PatternFill
        except Exception:                                               # noqa: BLE001
            return None
        wb = Workbook()
        b = wb.active
        b.title = "Bilgi"
        b.append(["TCMB EVDS arşivi"]); b["A1"].font = Font(bold=True, size=14)
        b.append([f"Güncelleme: {dt.datetime.now():%Y-%m-%d %H:%M}"])
        b.append([])
        b.append(["Sayfa", "Seri kodu", "Seri", "Birim", "Grup"])
        for c in b[4]:
            c.font = Font(bold=True)
        for frek, liste, sayfa in (("gunluk", GUNLUK, "Günlük"), ("haftalik", HAFTALIK, "Haftalık")):
            for k, a, bi, g, _ in liste:
                b.append([sayfa, k, a, bi, g])
        b.column_dimensions["B"].width = 26
        b.column_dimensions["C"].width = 52
        bas = PatternFill("solid", fgColor="1F3A5F")
        for frek, liste, sayfa in (("gunluk", GUNLUK, "Günlük"), ("haftalik", HAFTALIK, "Haftalık")):
            ws = wb.create_sheet(sayfa)
            kodlar = [k for k, *_ in liste]
            ws.append(["Tarih"] + [f"{a} ({bi})" for k, a, bi, g, _ in liste])
            for c in ws[1]:
                c.font = Font(bold=True, color="FFFFFF"); c.fill = bas
            tarihler = sorted({t for k in kodlar for t in self.v[frek].get(k, {})})
            for t in tarihler:
                ws.append([t] + [self.v[frek].get(k, {}).get(t) for k in kodlar])
            ws.freeze_panes = "B2"
            ws.column_dimensions["A"].width = 12
        if makro_veri and Path(makro_veri).exists():
            S = json.loads(Path(makro_veri).read_text(encoding="utf-8"))["seriler"]
            ws = wb.create_sheet("Aylık")
            kodlar = sorted(S)
            ws.append(["Ay"] + kodlar)
            for c in ws[1]:
                c.font = Font(bold=True, color="FFFFFF"); c.fill = bas
            for a in sorted({a for k in kodlar for a in S[k]}):
                ws.append([a] + [S[k].get(a) for k in kodlar])
            ws.freeze_panes = "B2"
        yol = self.k / "TCMB_Arsiv.xlsx"
        wb.save(yol)
        return yol


def guncelle(arsiv_klasoru: Path, anahtar: str | None, bastan=False, log=print, makro_veri: Path | None = None) -> int:
    A = Arsiv(arsiv_klasoru)
    toplam = 0
    if anahtar:
        bugun = dt.date.today().isoformat()
        for frek, liste in (("gunluk", GUNLUK), ("haftalik", HAFTALIK)):
            for kod, ad, *_rest in liste:
                frek_p = _rest[-1]
                son = None if bastan else A.son(frek, kod)
                ilk = min(A.v[frek].get(kod, {}) or ["9999"])
                if ilk > "2018-02-01":                          # geçmiş eksik (ör. önizleme tohumu): 2018'den tamamla
                    son = None
                bas = (dt.date.fromisoformat(son) - dt.timedelta(days=30)).isoformat() if son else BASLANGIC
                parca = bas
                try:
                    while parca <= bugun:                       # 2 yıllık parçalar (istek başına gözlem sınırı)
                        bit = min(bugun, (dt.date.fromisoformat(parca) + dt.timedelta(days=730)).isoformat())
                        toplam += A.ekle(frek, kod, evds(anahtar, kod, parca, bit, frek_p))
                        parca = (dt.date.fromisoformat(bit) + dt.timedelta(days=1)).isoformat()
                except Exception as e:                          # noqa: BLE001
                    log(f"  ! EVDS {kod}: {e}")
        log(f"  TCMB arşivi: {toplam} yeni/değişen gözlem")
    else:
        log("  TCMB arşivi: EVDS anahtarı yok, arşiv olduğu gibi kullanıldı")
    A.kaydet()
    A.excel(makro_veri)
    return toplam


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--arsiv", default=os.environ.get("TCMB_ARSIV") or str(BURADA / "tcmb_arsiv"))
    ap.add_argument("--anahtar", default=os.environ.get("EVDS_API_KEY"))
    ap.add_argument("--bastan", action="store_true", help="2018'den bütün geçmişi yeniden çek")
    ap.add_argument("--makro-veri", default=os.environ.get("MAKRO_VERI") or str(BURADA / "makro_veri.json"))
    a = ap.parse_args()
    sys.exit(0 if guncelle(Path(a.arsiv), a.anahtar, a.bastan, makro_veri=Path(a.makro_veri)) >= 0 else 1)
