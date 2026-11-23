#!/usr/bin/env python3
"""
TBB verisini artımlı günceller: arşivdeki (tbb_veri.csv) son dönemden sonraki dönemleri çeker ve birleştirir.
İlk çalıştırmada (arşiv yoksa) 2018-03'ten bütün dönemler çekilir. Site (rapor_olustur.py) yalnız veri
değiştiyse ya da site yoksa yeniden üretilir.

Depodaki tbb_cek.py ve rapor_olustur.py olduğu gibi kullanılır (onlara dokunulmaz).
"""
from __future__ import annotations

import datetime as dt
import json
import subprocess
import sys
from pathlib import Path


def _oku(yol: Path):
    banks, items, data, tur = [], {}, {}, None
    with open(yol, encoding="utf-8-sig") as f:
        for i, l in enumerate(f):
            l = l.rstrip("\n")
            if i == 0 or not l:
                continue
            if l.startswith("#BANKS "):
                banks = json.loads(l[7:])
            elif l.startswith("#ITEMS "):
                _, tur_, js = l.split(" ", 2)
                items[tur_] = {str(it[0]): it for it in json.loads(js)}
            elif l.startswith("#DATA "):
                tur = l.split(" ")[1]
                data.setdefault(tur, {})
            elif tur:
                c, b, o, tp, yp = l.split(",")
                data[tur][(c, b, o)] = f"{tp},{yp}"
    return banks, items, data


def donemler(yol: Path) -> list[str]:
    if not yol.exists():
        return []
    _, _, data = _oku(yol)
    return sorted({k[2] for d in data.values() for k in d})


def birlestir(eski: Path, yeni: Path, hedef: Path) -> list[str]:
    """yeni dosyanın değerleri eskiyi ezer; bankalar ve kalemler birleşir. Dönen: yeni gelen dönemler."""
    b1, i1, d1 = _oku(eski)
    b2, i2, d2 = _oku(yeni)
    bank = {str(b[0]): b for b in b1}
    bank.update({str(b[0]): b for b in b2})
    once = {k[2] for d in d1.values() for k in d}
    for tur, d in d2.items():
        d1.setdefault(tur, {}).update(d)
        i1.setdefault(tur, {}).update(i2.get(tur, {}))
    tum = sorted({k[2] for d in d1.values() for k in d})
    with open(hedef, "w", encoding="utf-8-sig", newline="\n") as f:
        f.write(f"# TBB TFRS9 {','.join(d1)} {tum[0]}..{tum[-1]} | alındı: {dt.date.today():%d.%m.%Y}\n")
        f.write("#BANKS " + json.dumps(list(bank.values()), ensure_ascii=False) + "\n")
        for tur in d1:
            f.write(f"#ITEMS {tur} " + json.dumps(list(i1.get(tur, {}).values()), ensure_ascii=False) + "\n")
            f.write(f"#DATA {tur}\n")
            f.write("\n".join(f"{c},{b},{o},{v}" for (c, b, o), v in sorted(d1[tur].items())) + "\n")
    return sorted({k[2] for d in d2.values() for k in d} - once)


def _sonraki(o: str) -> int:
    y, m = int(o[:4]), int(o[4:]) + 3
    if m > 12:
        y, m = y + 1, m - 12
    return y * 100 + m


def guncelle(depo: Path, klasor: Path, log=print, bastan=False) -> dict | None:
    """Dönen: {'periods','yeni','site','csv','html','dipnot'} ya da None (hata)."""
    klasor.mkdir(parents=True, exist_ok=True)
    csv, gecici = klasor / "tbb_veri.csv", klasor / "tbb_yeni.csv"
    mevcut = [] if bastan else donemler(csv)
    komut = [sys.executable, str(depo / "tbb_cek.py"), "--cikti", str(gecici if mevcut else csv)]
    if mevcut:
        komut += ["--baslangic", str(_sonraki(mevcut[-1]))]
        log(f"  TBB arşivi: {len(mevcut)} dönem ({mevcut[0]}–{mevcut[-1]}); yalnız {_sonraki(mevcut[-1])} ve sonrası sorgulanıyor")
    else:
        log("  TBB arşivi yok: 2018-03'ten bütün dönemler çekiliyor (ilk çalıştırma, uzun sürer)")
    log("$", " ".join(komut))
    p = subprocess.run(komut, cwd=depo, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=6 * 3600)
    cikti = (p.stdout or "") + (p.stderr or "")
    for s in cikti.splitlines()[-12:]:
        log("   ", s)
    yeni: list[str] = []
    if p.returncode != 0:
        if mevcut and "aralıkta dönem yok" in cikti:
            log("  TBB: yeni dönem yok")
        else:
            log("  ! TBB verisi çekilemedi")
            if not csv.exists():
                return None
    elif mevcut:
        yeni = birlestir(csv, gecici, csv)
        gecici.unlink(missing_ok=True)
        log(f"  TBB: yeni dönem(ler) eklendi: {', '.join(yeni) or '—'}")
    else:
        yeni = donemler(csv)
    site, js = klasor / "site", klasor / "tbb_rapor.json"
    if yeni or not (site / "index.html").exists() or not js.exists():
        site.mkdir(parents=True, exist_ok=True)
        r = subprocess.run([sys.executable, str(depo / "rapor_olustur.py"), "--girdi", str(csv), "--html", str(site / "index.html"),
                            "--json", str(js)], cwd=depo, capture_output=True, text=True, encoding="utf-8", errors="replace")
        if r.returncode != 0:
            log("  ! TBB raporu üretilemedi:", (r.stderr or "")[-500:])
            return None
        log(f"  TBB sitesi üretildi: {site / 'index.html'}")
    else:
        log("  TBB sitesi güncel (yeniden üretilmedi)")
    D = json.loads(js.read_text(encoding="utf-8"))
    return {"periods": D["solo"]["periods"], "yeni": yeni, "html": site / "index.html", "dipnot": site / "dipnot",
            "csv": csv, "site": site}
