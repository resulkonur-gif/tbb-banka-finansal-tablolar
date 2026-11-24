"""Tüm veri hattını tek bir çalıştırılabilir .py dosyasına paketler (git ve depo kopyası gerekmez).

    python veri_hatti/tek_dosya_uret.py            ->  banka_veri.py (depo kökünde)

Üretilen dosya; TBB çekicisini, rapor şablonunu, veri hattı modüllerini ve TCMB arşivini
sıkıştırılmış olarak içinde taşır. Çalıştırıldığında kendini ana klasöre açar ve hattı başlatır.
"""
from __future__ import annotations

import base64
import datetime as dt
import hashlib
import io
import zipfile
from pathlib import Path

DEPO = Path(__file__).resolve().parent.parent
KOK_DOSYALAR = ["tbb_cek.py", "rapor_olustur.py", "rapor_sablonu.html", "tbb_yardimci.js", "LICENSE"]
ATLA_UZANTI = {".pyc", ".bat", ".sh"}

BASLATICI = r'''#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
BANKA VERİ — tek dosyalık çalıştırıcı (sürüm __SURUM__)
=======================================================

Ne yapar
  TBB Veri Sistemi, KAP banka raporları (PDF), yatırımcı sunumları, TCMB/EVDS ve BDDK
  verilerini indirir; Excel, HTML site ve özet dosyalarını ana klasöre yazar.
  İlk çalıştırmada her şeyi indirir; sonrakilerde yalnız yeni dönemleri getirir.

Nasıl çalıştırılır
  python banka_veri.py                      (ya da dosyaya çift tıklayın)
  python banka_veri.py --klasor C:\Veri\BankaVeri
  python banka_veri.py --evds-anahtar ANAHTAR
  python banka_veri.py --yardim             (bütün seçenekler)

Gerekenler
  Python 3.10 veya üstü. Eksik paketler (playwright, pymupdf, openpyxl, requests)
  ilk çalıştırmada kullanıcı hesabınıza kurulur (yönetici yetkisi gerekmez).
  Playwright'ın Chromium'u indirilemezse bilgisayardaki Edge ya da Chrome kullanılır.

Çıktılar (varsayılan ana klasör: Belgeler\BankaVeri)
  7_site\site\index.html   tam site (tarayıcıda açılır)
  OZET.md                  çalıştırma özeti
  1_tbb, 2_kap, 3_ek_veri, 4_karsilastirma, 5_makro, 6_sunum, 8_degerlendirme
  gunluk.log               ayrıntılı günlük
"""
import argparse
import base64
import hashlib
import importlib.util
import io
import os
import subprocess
import sys
import zipfile
from pathlib import Path

SURUM = "__SURUM__"
PAKET_OZET = "__OZET__"
PAKETLER = [("playwright", "playwright"), ("pymupdf", "pymupdf"), ("openpyxl", "openpyxl"), ("requests", "requests")]


def yaz(m=""):
    print(m, flush=True)


def varsayilan_klasor() -> Path:
    ev = Path.home()
    for ad in ("Documents", "Belgeler", "OneDrive/Documents", "OneDrive/Belgeler"):
        if (ev / ad).is_dir():
            return ev / ad / "BankaVeri"
    return ev / "BankaVeri"


def pip_kur(paketler, log):
    sanal = sys.prefix != getattr(sys, "base_prefix", sys.prefix)
    komut = [sys.executable, "-m", "pip", "install", "--disable-pip-version-check", *([] if sanal else ["--user"]), *paketler]
    log("  $ " + " ".join(komut[2:]))
    r = subprocess.run(komut)
    if r.returncode != 0 and not sanal:              # bazı kurulumlarda --user yasak (ör. Microsoft Store Python'u değil, venv)
        r = subprocess.run([c for c in komut if c != "--user"])
    return r.returncode == 0


def paketleri_hazirla(log, kurma: bool) -> bool:
    eksik = [pip for mod, pip in PAKETLER if importlib.util.find_spec(mod) is None
             and not (mod == "pymupdf" and importlib.util.find_spec("fitz"))]
    if eksik:
        if kurma:
            log("Eksik paketler: " + ", ".join(eksik) + "  (--kurma verildi, kurulmadı)")
            return False
        log("Eksik paketler kuruluyor: " + ", ".join(eksik))
        if not pip_kur(eksik, log):
            log("! Paketler kurulamadı. Şirket ağında vekil sunucu (proxy) varsa önce şunu deneyin:")
            log("    set HTTPS_PROXY=http://kullanici:sifre@proxy.sirket:8080   (Windows)")
            log("  ya da BT biriminden şu paketleri isteyin: " + " ".join(eksik))
            return False
    if importlib.util.find_spec("truststore") is None and not kurma:
        # Şirket ağları HTTPS'i kendi sertifikalarıyla açıyorsa Windows sertifika deposu gerekir
        subprocess.run([sys.executable, "-m", "pip", "install", "-q", "--disable-pip-version-check",
                        *([] if sys.prefix != getattr(sys, "base_prefix", sys.prefix) else ["--user"]), "truststore"],
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    return True


def tarayici_bul(program: Path) -> str:
    """tbb_cek.tarayici_ac ile gerçekten sekme açılabiliyor mu; açılabilen adayın adını döndürür."""
    kod = ("import sys; sys.path.insert(0, sys.argv[1])\n"
           "from playwright.sync_api import sync_playwright\n"
           "import tbb_cek\n"
           "with sync_playwright() as p:\n"
           "    b = tbb_cek.tarayici_ac(p, headless=True); b.close(); print('tamam')\n")
    try:
        r = subprocess.run([sys.executable, "-c", kod, str(program)], capture_output=True, text=True, timeout=240)
    except Exception as e:                              # noqa: BLE001
        return "", str(e)
    if "tamam" in (r.stdout or ""):
        return "tamam", ""
    hata = [x for x in (r.stderr or "").splitlines() if "Tarayıcı açılamadı" in x or "Error" in x]
    return "", (hata[-1] if hata else (r.stderr or "").strip()[-300:])


def chromium_hazirla(klasor: Path, program: Path, log, kurma: bool):
    isaret = klasor / "_program" / ".tarayici"
    if isaret.exists() and isaret.read_text(encoding="utf-8") == SURUM:
        return
    log("Tarayıcı sınanıyor…")
    tamam, hata = tarayici_bul(program)
    if not tamam and not kurma:
        log("  Açılamadı; Playwright güncelleniyor ve yeniden deneniyor…")
        pip_kur(["-q", "--upgrade", "playwright"], lambda m: None)
        tamam, hata = tarayici_bul(program)
    if not tamam and not kurma:
        log("  Playwright Chromium indiriliyor… (~150 MB)")
        r = subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"], capture_output=True, text=True)
        if r.returncode == 0:
            tamam, hata = tarayici_bul(program)
        else:
            log("  Chromium indirilemedi (şirket ağı engelliyor olabilir).")
    if tamam:
        log("  Tarayıcı hazır.")
        isaret.parent.mkdir(parents=True, exist_ok=True)
        isaret.write_text(SURUM, encoding="utf-8")
    else:
        log("! Tarayıcı açılamadı; TBB ve yatırımcı sunumu adımları atlanacak (KAP, TCMB, BDDK çalışır).")
        log("  Ayrıntı: " + hata[:400])
    return bool(tamam)


def programi_ac(klasor: Path, log) -> Path:
    hedef = klasor / "_program" / SURUM
    isaret = hedef / ".paket"
    if isaret.exists() and isaret.read_text(encoding="utf-8") == PAKET_OZET:
        return hedef
    log(f"Program dosyaları açılıyor: {hedef}")
    veri = base64.b64decode(PAKET)
    if hashlib.sha256(veri).hexdigest() != PAKET_OZET:
        sys.exit("! Dosya bozulmuş görünüyor (paket özeti tutmadı). banka_veri.py'yi yeniden indirin.")
    hedef.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(io.BytesIO(veri)) as z:
        z.extractall(hedef)
    # Şirket sertifikaları: truststore varsa bütün alt süreçlerde Windows/macOS sertifika deposu kullanılır
    (hedef / "sitecustomize.py").write_text(
        "try:\n    import truststore\n    truststore.inject_into_ssl()\nexcept Exception:\n    pass\n", encoding="utf-8")
    isaret.write_text(PAKET_OZET, encoding="utf-8")
    return hedef


def evds_anahtari(a, klasor: Path, log):
    dosya = klasor / "evds_anahtar.txt"
    if a.evds_anahtar:
        dosya.write_text(a.evds_anahtar.strip(), encoding="utf-8")
        return a.evds_anahtar.strip()
    if os.environ.get("EVDS_API_KEY"):
        return os.environ["EVDS_API_KEY"]
    if dosya.exists() and dosya.read_text(encoding="utf-8").strip():
        return dosya.read_text(encoding="utf-8").strip()
    if sys.stdin and sys.stdin.isatty() and not a.sor_ma:
        yaz("")
        yaz("TCMB/EVDS API anahtarı (evds2.tcmb.gov.tr → Profil → API Anahtarı).")
        yaz("Boş bırakırsanız TCMB verileri paketteki arşivle kalır, yenisi indirilmez.")
        try:
            k = input("EVDS anahtarı: ").strip()
        except EOFError:
            k = ""
        if k:
            dosya.write_text(k, encoding="utf-8")
            log(f"  Anahtar kaydedildi: {dosya} (bir dahaki sefere sorulmaz)")
            return k
    return ""


def main():
    ap = argparse.ArgumentParser(description="Banka verileri: TBB + KAP + sunumlar + TCMB + BDDK → Excel ve HTML",
                                 add_help=False)
    ap.add_argument("-h", "--yardim", "--help", action="help", help="bu yardımı göster")
    ap.add_argument("--klasor", default=str(varsayilan_klasor()), help="ana veri klasörü (varsayılan: Belgeler\\BankaVeri)")
    ap.add_argument("--evds-anahtar", default="", help="EVDS API anahtarı (bir kez verilir, klasöre kaydedilir)")
    ap.add_argument("--son", default="", help="son dönem YYYY-AA (varsayılan: son tamamlanan çeyrek)")
    ap.add_argument("--tbb-atla", action="store_true", help="TBB'yi sorgulama")
    ap.add_argument("--kap-atla", action="store_true", help="KAP'tan PDF indirme")
    ap.add_argument("--sunum-atla", action="store_true", help="yatırımcı sunumlarını indirme")
    ap.add_argument("--ek-veri-yeniden", action="store_true", help="ek veriyi PDF'lerden yeniden oku")
    ap.add_argument("--bastan", action="store_true", help="TBB ve EVDS'yi baştan indir (normalde gerekmez)")
    ap.add_argument("--kurma", action="store_true", help="paket/Chromium kurmayı deneme")
    ap.add_argument("--acma", action="store_true", help="bitince siteyi tarayıcıda açma")
    ap.add_argument("--sor-ma", action="store_true", help="hiçbir şey sorma (zamanlanmış görev için)")
    a = ap.parse_args()

    yaz(f"BANKA VERİ  (sürüm {SURUM})")
    if sys.version_info < (3, 10):
        sys.exit(f"! Python 3.10 veya üstü gerekli; bu Python {sys.version.split()[0]}. python.org'dan yeni sürüm kurun.")
    klasor = Path(a.klasor).expanduser().resolve()
    klasor.mkdir(parents=True, exist_ok=True)
    yaz(f"Ana klasör: {klasor}")

    if not paketleri_hazirla(yaz, a.kurma):
        sys.exit(1)
    program = programi_ac(klasor, yaz)
    tarayici_var = chromium_hazirla(klasor, program, yaz, a.kurma) is not False
    anahtar = evds_anahtari(a, klasor, yaz)

    env = dict(os.environ, PYTHONPATH=str(program) + os.pathsep + os.environ.get("PYTHONPATH", ""),
               PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
    if anahtar:
        env["EVDS_API_KEY"] = anahtar
    komut = [sys.executable, str(program / "veri_hatti" / "veri_hatti.py"), "--kok", str(klasor), "--depo", str(program)]
    if not tarayici_var:
        a.tbb_atla = a.sunum_atla = True
    for bayrak in ("tbb_atla", "kap_atla", "sunum_atla", "ek_veri_yeniden", "bastan"):
        if getattr(a, bayrak):
            komut.append("--" + bayrak.replace("_", "-"))
    if a.son:
        komut += ["--son", a.son]
    yaz("")
    r = subprocess.run(komut, cwd=str(program), env=env)

    site = klasor / "7_site" / "site" / "index.html"
    yaz("")
    if site.exists():
        yaz(f"Site:  {site}")
    yaz(f"Özet:  {klasor / 'OZET.md'}")
    if site.exists() and not a.acma and not a.sor_ma:
        try:
            import webbrowser
            webbrowser.open(site.as_uri())
        except Exception:
            pass
    if os.name == "nt" and not a.sor_ma and sys.stdin and sys.stdin.isatty():
        try:
            input("\nKapatmak için Enter'a basın…")
        except EOFError:
            pass
    sys.exit(r.returncode)


PAKET = (
__PAKET__
)

if __name__ == "__main__":
    main()
'''


def paketle() -> bytes:
    tampon = io.BytesIO()
    with zipfile.ZipFile(tampon, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        for ad in KOK_DOSYALAR:
            if (DEPO / ad).exists():
                z.write(DEPO / ad, ad)
        for f in sorted((DEPO / "veri_hatti").rglob("*")):
            if f.is_file() and "__pycache__" not in f.parts and f.suffix not in ATLA_UZANTI and f.name != Path(__file__).name:
                z.write(f, f.relative_to(DEPO).as_posix())
    return tampon.getvalue()


def main():
    veri = paketle()
    b64 = base64.b64encode(veri).decode()
    satirlar = "\n".join(f'    "{b64[i:i + 100]}"' for i in range(0, len(b64), 100))
    surum = dt.date.today().isoformat()
    metin = (BASLATICI.replace("__SURUM__", surum).replace("__OZET__", hashlib.sha256(veri).hexdigest())
             .replace("__PAKET__", satirlar))
    hedef = DEPO / "banka_veri.py"
    hedef.write_text(metin, encoding="utf-8", newline="\n")
    print(f"{hedef}  ({len(metin) // 1024} KB, paket {len(veri) // 1024} KB)")


if __name__ == "__main__":
    main()
