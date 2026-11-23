#!/usr/bin/env python3
"""
TBB / KAP otomatik güncelleme – tek çalıştırma.

Ne yapar
  1. Bugünün tarihinden son tamamlanmış çeyreği bulur (ör. 18.11.2026 → 2026-09).
  2. TBB Veri Sistemi'nden veriyi çeker (depodaki tbb_cek.py) ve raporu üretir (rapor_olustur.py).
     - TBB bu çeyreği yayımladıysa: rapor TBB verisiyle güncellenir; daha önce KAP'tan önizleme
       üretildiyse önizleme kaldırılır (TBB verisi esas alınır).
     - TBB yayımlamadıysa: 7 büyük bankanın raporları KAP'tan indirilir, PDF'ler okunup TBB ile karşılaştırılır
       (Excel) ve sitenin "önizleme" sürümü üretilir. KAP'ta da henüz yoksa bir önceki çeyrek denenir.
  3. Bütün çıktıları tek klasöre toplar:  cikti/<çalıştırma zamanı>/
       site/            TBB verisiyle ana rapor (index.html + dipnot/)
       onizleme/        KAP verisiyle önizleme (yalnız TBB son çeyreği yayımlamadıysa)
       karsilastirma_YYYYAA.xlsx   KAP ↔ TBB karşılaştırması (eşleme haritası, banka iş kuralları dahil)
       kap_pdf/         indirilen KAP raporları
       gunluk.txt       çalıştırma günlüğü
  4. --yayinla verilirse depodaki docs/ klasörünü günceller ve GitHub'a gönderir
     (TBB verisi: docs/index.html + docs/dipnot, önizleme silinir; KAP verisi: yalnız docs/onizleme).
  5. Durumu durum.json'a yazar: hangi çeyrek hangi kaynaktan (tbb / kap) alındı.

Kullanım
  python otomatik_guncelle.py --depo C:\\yol\\tbb-banka-finansal-tablolar [--yayinla] [--donem 202609]
  Haftalık (pazartesi) çalıştırmak için: haftalik_gorev_kur.bat (Windows) ya da README'deki cron satırı.

Gerekenler: Python 3.10+, playwright (+ chromium), pymupdf, openpyxl, requests; git (yayınlamak için).
"""
from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
import traceback
from datetime import datetime
from pathlib import Path

BURADA = Path(__file__).resolve().parent
sys.path.insert(0, str(BURADA))


# ----------------------------------------------------------------------------- yardımcılar
class Gunluk:
    def __init__(self, yol: Path):
        self.yol = yol
        yol.parent.mkdir(parents=True, exist_ok=True)

    def __call__(self, *x):
        satir = f"[{datetime.now():%H:%M:%S}] " + " ".join(str(a) for a in x)
        print(satir, flush=True)
        with self.yol.open("a", encoding="utf-8") as f:
            f.write(satir + "\n")


def son_ceyrek(bugun: datetime | None = None) -> tuple[int, int]:
    """Bugünden önce biten son çeyrek (yıl, ay)."""
    n = bugun or datetime.now()
    ay = (n.month - 1) // 3 * 3
    return (n.year - 1, 12) if ay == 0 else (n.year, ay)


def onceki_ceyrek(yil: int, ay: int) -> tuple[int, int]:
    return (yil - 1, 12) if ay == 3 else (yil, ay - 3)


def ceyrek_str(yil: int, ay: int) -> str:
    return f"{yil}-{ay:02d}"


def calistir(komut: list[str], log: Gunluk, cwd: Path | None = None, zaman_asimi: int = 3 * 3600) -> bool:
    log("$", " ".join(str(k) for k in komut))
    try:
        p = subprocess.run([str(k) for k in komut], cwd=cwd, capture_output=True, text=True,
                           encoding="utf-8", errors="replace", timeout=zaman_asimi)
    except subprocess.TimeoutExpired:
        log("  ! zaman aşımı")
        return False
    for satir in (p.stdout or "").splitlines()[-25:]:
        log("  ", satir)
    if p.returncode != 0:
        for satir in (p.stderr or "").splitlines()[-25:]:
            log("  !", satir)
    return p.returncode == 0


def durum_oku(yol: Path) -> dict:
    try:
        return json.loads(yol.read_text(encoding="utf-8"))
    except Exception:
        return {}


def durum_yaz(yol: Path, d: dict):
    yol.write_text(json.dumps(d, ensure_ascii=False, indent=2), encoding="utf-8")


# ----------------------------------------------------------------------------- 1) TBB
def tbb_guncelle(depo: Path, is_klasoru: Path, log: Gunluk) -> dict | None:
    """tbb_cek.py + rapor_olustur.py; dönen: {'periods': [...], 'html': yol, 'dipnot': yol, 'csv': yol}"""
    csv = is_klasoru / "tbb_veri.csv"
    if not calistir([sys.executable, depo / "tbb_cek.py", "--cikti", csv], log, cwd=depo):
        log("  ! TBB verisi çekilemedi")
        return None
    site = is_klasoru / "site"
    site.mkdir(parents=True, exist_ok=True)
    js = is_klasoru / "tbb_rapor.json"
    if not calistir([sys.executable, depo / "rapor_olustur.py", "--girdi", csv, "--html", site / "index.html",
                     "--json", js], log, cwd=depo):
        log("  ! TBB raporu üretilemedi")
        return None
    try:
        D = json.loads(js.read_text(encoding="utf-8"))
        donemler = D["solo"]["periods"]
    except Exception as e:
        log(f"  ! rapor JSON okunamadı: {e}")
        return None
    return {"periods": donemler, "html": site / "index.html", "dipnot": site / "dipnot", "csv": csv, "site": site}


# ----------------------------------------------------------------------------- 2) KAP
def kap_indir_ve_karsilastir(yil: int, ay: int, is_klasoru: Path, log: Gunluk, eski_pdf: Path | None) -> Path | None:
    """7 bankanın KAP raporlarını indirir, TBB ile karşılaştırır (Excel). Dönen: PDF klasörü (rapor yoksa None)."""
    import banka_karsilastir as bk
    pdf_kok = is_klasoru / "kap_pdf" / f"banka_karsilastirma_{yil}{ay:02d}"
    pdf_kok.mkdir(parents=True, exist_ok=True)
    raporlar = []
    for b in bk.BANKALAR:
        hedef = pdf_kok / bk.dosya_adi(b[2])
        onceki = (eski_pdf / bk.dosya_adi(b[2])) if eski_pdf else None
        if onceki and onceki.exists() and any(onceki.glob("*.pdf")):          # daha önce indirilmiş: yeniden indirme
            shutil.copytree(onceki, hedef, dirs_exist_ok=True)
            pdfler = list(hedef.glob("*.pdf"))
            log(f"  • {b[2]}: önceki indirmeden ({len(pdfler)} PDF)")
        else:
            try:
                pdfler = bk.kap_paketi_indir(b, yil, ay, hedef)
                log(f"  • {b[2]}: KAP'tan {len(pdfler)} PDF")
            except Exception as e:
                log(f"  • {b[2]}: KAP'ta yok ya da indirilemedi ({e})")
                continue
        bulunan = {}
        for p in pdfler:
            t = bk.pdf_turu(p)
            if t and t not in bulunan:
                bulunan[t] = p
        raporlar += [(b, t, p) for t, p in bulunan.items()]
    if not raporlar:
        return None
    log(f"  {len(raporlar)} rapor bulundu; TBB ile karşılaştırılıyor")
    D = bk.tbb_verisi()
    rapor = bk.Rapor(yil, ay, D)
    for b, t, p in raporlar:
        try:
            ok = bk.Okuyucu(p).oku()
            rapor.ekle(b, t, ok, p.name)
            rapor.ekle_kur(b, t, ok)
            rapor.ekle_dipnot(b, t, ok)
            rapor.ekle_gosterge(b, t, ok)
            rapor.ekle_genel(b, t, ok)
        except Exception as e:
            log(f"  ! {b[2]} {t}: okunamadı ({e})")
    excel = is_klasoru / f"karsilastirma_{yil}{ay:02d}.xlsx"
    rapor.kaydet(excel)
    log(f"  → {excel.name}")
    return pdf_kok


def onizleme_uret(yil: int, ay: int, pdf_kok: Path, site_docs: Path, is_klasoru: Path, log: Gunluk) -> Path | None:
    hedef = is_klasoru / "onizleme"
    if hedef.exists():
        shutil.rmtree(hedef)
    ok = calistir([sys.executable, BURADA / "eylul_ekle.py", "--donem", f"{yil}{ay:02d}", "--pdf-klasoru", pdf_kok,
                   "--site", site_docs, "--cikti", hedef], log, cwd=BURADA)
    if not (ok and (hedef / "index.html").exists()):
        return None
    # yatırımcı özeti (karne, çeyreklik marj, aktif kalitesi, sermaye tamponu, banka notları)
    # makro pano (EVDS): anahtar EVDS_API_KEY ortam değişkeninde; yoksa makro_veri.json'daki son veriyle
    calistir([sys.executable, BURADA / "makro_pano.py", "--site", hedef], log, cwd=BURADA)
    calistir([sys.executable, BURADA / "yatirimci_ozeti.py", "--site", hedef, "--onizleme", f"{yil}-{ay:02d}",
              "--dipnot", Path(site_docs) / "dipnot", "--pdf-klasoru", pdf_kok], log, cwd=BURADA)
    return hedef


# ----------------------------------------------------------------------------- 3) yayın
def git(depo: Path, log: Gunluk, *args) -> bool:
    return calistir(["git", *args], log, cwd=depo, zaman_asimi=600)


def yayinla_tbb(depo: Path, tbb: dict, log: Gunluk, mesaj: str):
    docs = depo / "docs"
    shutil.copy2(tbb["html"], docs / "index.html")
    if tbb["dipnot"].exists():
        if (docs / "dipnot").exists():
            shutil.rmtree(docs / "dipnot")
        shutil.copytree(tbb["dipnot"], docs / "dipnot")
    if (docs / "onizleme").exists():                       # TBB yayımladı: KAP önizlemesi kaldırılır
        shutil.rmtree(docs / "onizleme")
        log("  docs/onizleme silindi (TBB verisi geldi)")
    git(depo, log, "add", "-A", "docs")
    if git(depo, log, "commit", "-m", mesaj):
        git(depo, log, "push", "origin", "HEAD")


def yayinla_onizleme(depo: Path, onizleme: Path, log: Gunluk, mesaj: str):
    hedef = depo / "docs" / "onizleme"
    if hedef.exists():
        shutil.rmtree(hedef)
    shutil.copytree(onizleme, hedef)
    git(depo, log, "add", "docs/onizleme")
    if git(depo, log, "commit", "-m", mesaj):
        git(depo, log, "push", "origin", "HEAD")


# ----------------------------------------------------------------------------- ana akış
def main():
    ap = argparse.ArgumentParser(description="TBB yayımladıysa TBB'den, yayımlamadıysa KAP'tan son çeyreği alır.")
    ap.add_argument("--depo", required=True, help="tbb-banka-finansal-tablolar deposunun yerel kopyası")
    ap.add_argument("--cikti", default=str(BURADA / "cikti"), help="çıktıların toplanacağı klasör")
    ap.add_argument("--donem", default="", help="YYYYAA; verilmezse bugünden son çeyrek bulunur")
    ap.add_argument("--yayinla", action="store_true", help="docs/ klasörünü güncelle ve GitHub'a gönder")
    ap.add_argument("--tbb-atla", action="store_true", help="TBB'den çekme (yalnız KAP; deneme için)")
    a = ap.parse_args()

    depo = Path(a.depo).resolve()
    if not (depo / "tbb_cek.py").exists():
        sys.exit(f"Depo bulunamadı ya da eksik: {depo}")
    is_klasoru = Path(a.cikti).resolve() / datetime.now().strftime("%Y%m%d_%H%M")
    is_klasoru.mkdir(parents=True, exist_ok=True)
    log = Gunluk(is_klasoru / "gunluk.txt")
    durum_yolu = Path(a.cikti).resolve() / "durum.json"
    durum = durum_oku(durum_yolu)

    yil, ay = (int(a.donem[:4]), int(a.donem[4:])) if a.donem else son_ceyrek()
    log(f"Hedef çeyrek: {ceyrek_str(yil, ay)}   (çıktı: {is_klasoru})")
    if a.yayinla:
        git(depo, log, "pull", "--ff-only")

    # 1) TBB
    tbb = None
    if not a.tbb_atla:
        log("[1] TBB Veri Sistemi")
        tbb = tbb_guncelle(depo, is_klasoru, log)
    tbb_son = tbb["periods"][-1] if tbb else None
    log(f"  TBB'deki son dönem: {tbb_son}")

    hedefler = [(yil, ay), onceki_ceyrek(yil, ay)] if not a.donem else [(yil, ay)]
    for y_, a_ in hedefler:
        p = ceyrek_str(y_, a_)
        if tbb and p in tbb["periods"]:
            log(f"[2] {p}: TBB yayımlamış → TBB verisi esas")
            eski = durum.get(p, {}).get("kaynak")
            if eski == "kap":
                log(f"  {p} daha önce KAP'tan alınmıştı; TBB verisiyle değiştiriliyor")
            if a.yayinla:
                yayinla_tbb(depo, tbb, log, f"TBB verisi: {p}" + (" (KAP önizlemesi kaldırıldı)" if eski == "kap" else ""))
            durum[p] = {"kaynak": "tbb", "tarih": datetime.now().isoformat(timespec="seconds")}
            break
        log(f"[2] {p}: TBB henüz yayımlamamış → KAP")
        eski_pdf = None
        if durum.get(p, {}).get("kaynak") == "kap" and durum[p].get("pdf"):
            eski_pdf = Path(durum[p]["pdf"])
        try:
            pdf_kok = kap_indir_ve_karsilastir(y_, a_, is_klasoru, log, eski_pdf)
        except Exception:
            log("  ! KAP adımı hata verdi:\n" + traceback.format_exc())
            pdf_kok = None
        if not pdf_kok:
            log(f"  {p} için KAP'ta rapor yok; bir önceki çeyrek deneniyor" if (y_, a_) == hedefler[0] and len(hedefler) > 1
                else f"  {p} için KAP'ta rapor yok")
            continue
        site_docs = tbb["site"] if tbb else depo / "docs"
        oniz = onizleme_uret(y_, a_, pdf_kok, site_docs, is_klasoru, log)
        if not oniz:
            log("  ! önizleme üretilemedi")
            break
        log(f"  → önizleme: {oniz}")
        if a.yayinla:
            if tbb:                                    # TBB'nin önceki dönemleri de güncel olsun
                yayinla_tbb(depo, tbb, log, f"TBB verisi: {tbb_son}")
            yayinla_onizleme(depo, oniz, log, f"Önizleme: {p} KAP verisi")
        durum[p] = {"kaynak": "kap", "tarih": datetime.now().isoformat(timespec="seconds"), "pdf": str(pdf_kok)}
        break
    else:
        log("Yeni çeyrek verisi bulunamadı; yalnız TBB raporu güncellendi." if tbb else "Hiçbir kaynaktan veri alınamadı.")

    durum_yaz(durum_yolu, durum)
    shutil.copy2(durum_yolu, is_klasoru / "durum.json")
    log(f"Bitti. Bütün çıktılar: {is_klasoru}")


if __name__ == "__main__":
    main()
