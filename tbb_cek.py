#!/usr/bin/env python3
"""
TBB Veri Sistemi'nden TFRS 9 solo ve konsolide finansal tabloları çeker.

Çekilen tablolar (her iki tür için):
  1. Varlıklar, 2. Yükümlülükler, 3. Nazım hesaplar, 4. Gelir-gider tablosu,
  5. Kâr veya zarar ve diğer kapsamlı gelir tablosu, 7. Nakit akış tablosu
  + likidite karşılama oranı (LKO) ve net istikrarlı fonlama oranı (NİFO) kalemleri
  + dipnotlardan mevduatın vade yapısı ile mali bünye ile ilgili dipnotların tamamı (özkaynak, risk profilleri, kur ve faiz riski,
    likidite, kalan vadeler, kaldıraç, gerçeğe uygun değer)

Yöntem: https://verisistemi.tbb.org.tr sayfası Playwright ile açılır, sitenin kendi
seçim ve "Rapor Al" fonksiyonları çağrılır, sonuç pivot tablosundan okunur.
Dönemler birkaç çeyreklik gruplar halinde çekilir.

Kurulum:
    pip install playwright
    playwright install chromium

Kullanım:
    python tbb_cek.py                                  # 2018-03'ten son döneme, solo + konsolide
    python tbb_cek.py --baslangic 202203 --bitis 202606
    python tbb_cek.py --tur solo --gorunur             # tarayıcıyı ekranda göster

Çıktı (varsayılan tbb_veri.csv), satır biçimi: kod,banka_kodu,donem,TP,YP
Tutarlar bin TL'dir; oran kalemleri (LKO, NİFO) yüzde olarak gelir.
"""
import argparse
import datetime as dt
import json
import sys
import time
from pathlib import Path

try:
    from playwright.sync_api import sync_playwright
except ImportError:
    sys.exit("Playwright kurulu değil:  pip install playwright  &&  playwright install chromium")

URL = "https://verisistemi.tbb.org.tr/index.php?/tbb/report_mali"
HERE = Path(__file__).resolve().parent

# Tablo kök kodları: alt kalemleriyle seçilenler ve yalnız kendisi seçilen likidite kalemleri
# Dipnotlar: mevduatın vade yapısı (cari dönem, tüm alt kalemleriyle) ve yasal özkaynak / sermaye yeterliliği
MEVDUAT_VADE = 114172
AKTIF_DIPNOT = 102774  # "2. Aktiflerle ilgili dipnotlar": krediler, menkul değerler, bankalar, iştirakler vb.
NAZIM_DIPNOT = 102780  # "4. Nazım hesaplarla ilgili dipnotlar"
GELIR_DIPNOT = 102778  # "5. Gelir tablosu ile ilgili dipnotlar"
RISK_GRUBU_DIPNOT = 102782  # "6. Risk grubuna ait dipnotlar"
PASIF_DIPNOT = 102776  # "3. Pasiflerle ilgili dipnotlar": mevduat vade yapısı, sigorta kapsamı, alınan krediler, sermaye benzeri borçlar vb.
MALI_BUNYE = 102772  # "1. Mali bünye ile ilgili dipnotlar": özkaynak kalemleri, risk profilleri, kur ve faiz riski,
                     # LKO, NİFO, kalan vadeye göre aktif-pasif, kaldıraç oranı, gerçeğe uygun değer (tüm alt kalemleriyle)

PARCA_BOYU = 600     # dipnotlar, en çok bu kadar kalemlik parçalar halinde çekilir (TBB raporu büyüyünce zaman aşımına düşüyor)

TABLOLAR = {
    "solo": {
        "alt_kalemleriyle": [102087, 102086, 102519, 102567, 16039, 102404, MEVDUAT_VADE],
        "yalniz": [109460, 109463, 109466, 109467, 17130, 17131, 17144],
    },
    "kons": {
        "alt_kalemleriyle": [352087, 352086, 352519, 352567, 266039, 352404, MEVDUAT_VADE + 250000],
        "yalniz": [359460, 359463, 359466, 359467, 487130, 487131, 487144],
    },
}


def log(*a):
    print(time.strftime("%H:%M:%S"), *a, flush=True)


def main():
    ap = argparse.ArgumentParser(description="TBB Veri Sistemi'nden solo/konsolide finansal tabloları çeker.")
    ap.add_argument("--baslangic", type=int, default=201803, help="İlk dönem, YYYYAA (varsayılan 201803)")
    ap.add_argument("--bitis", type=int, default=None, help="Son dönem, YYYYAA (varsayılan: sitedeki son dönem)")
    ap.add_argument("--tur", default="solo,kons", help="solo, kons ya da solo,kons")
    ap.add_argument("--parti", type=int, default=6, help="Bir seferde çekilecek dönem sayısı (varsayılan 6)")
    ap.add_argument("--cikti", default="tbb_veri.csv", help="Çıktı dosyası")
    ap.add_argument("--gorunur", action="store_true", help="Tarayıcıyı ekranda göster")
    ap.add_argument("--dipnotsuz", action="store_true", help="Mali bünye dipnotlarını çekme (yalnız finansal tablolar)")
    args = ap.parse_args()

    turler = [t.strip() for t in args.tur.split(",") if t.strip()]
    for t in turler:
        if t not in TABLOLAR:
            sys.exit(f"Bilinmeyen tür: {t} (solo ya da kons olmalı)")

    js = (HERE / "tbb_yardimci.js").read_text(encoding="utf-8")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not args.gorunur)
        page = browser.new_page(viewport={"width": 1400, "height": 900}, locale="tr-TR")
        page.set_default_timeout(240_000)
        log("TBB Veri Sistemi açılıyor…")
        page.goto(URL, wait_until="domcontentloaded")
        page.wait_for_function(
            """() => { try { return window.mlt_data_donemler && Object.keys(mlt_data_donemler).length
                 && window.mlt_data_bankalar && Object.keys(mlt_data_bankalar).length
                 && jQuery('#maliTablolarTree').dxTreeView('instance').option('items').length; } catch(e) { return false; } }""",
            timeout=120_000,
        )
        page.add_script_tag(content=js)
        page.evaluate("__tbb.init()")

        donemler = [d for d in page.evaluate("__tbb.periods()") if d >= args.baslangic]
        if args.bitis:
            donemler = [d for d in donemler if d <= args.bitis]
        if not donemler:
            sys.exit("Seçilen aralıkta dönem yok.")
        bankalar = page.evaluate("__tbb.banks()")
        grupsuz = [b[1] for b in bankalar if b[2] == "Grup bilgisi yok"]
        if grupsuz:
            log(f"TBB grubu bulunamayan bankalar: {grupsuz}")
            log("  grup kayıtları: " + json.dumps(page.evaluate("""(adlar) => Object.values(mlt_data_bankalar).filter(b => adlar.includes(b.TR_ADI.trim())).map(b => [b.TR_ADI.trim(), b.GRUPLAR])""", grupsuz), ensure_ascii=False)[:3000])
        log(f"{len(donemler)} dönem ({donemler[0]}–{donemler[-1]}), {len(bankalar)} banka")

        bolumler = {}
        for tur in turler:
            # 1. parça: finansal tablolar ve likidite kalemleri; sonra mali bünye dipnotları dallara bölünerek
            parcalar = [(TABLOLAR[tur]["alt_kalemleriyle"], TABLOLAR[tur]["yalniz"], args.parti)]
            if not args.dipnotsuz:
                dallar = []
                for kok, ad_ in ((MALI_BUNYE, "mali bünye"), (AKTIF_DIPNOT, "aktiflerle ilgili"), (PASIF_DIPNOT, "pasiflerle ilgili"),
                                 (NAZIM_DIPNOT, "nazım hesaplarla ilgili"), (GELIR_DIPNOT, "gelir tablosu ile ilgili"),
                                 (RISK_GRUBU_DIPNOT, "risk grubuna ait")):
                    d_ = page.evaluate("(k) => __tbb.children(k)", kok + (250000 if tur == "kons" else 0))
                    log(f"[{tur}] {ad_} dipnotları: {len(d_)} dal, {sum(n for _, n in d_)} kalem")
                    dallar += d_
                cur, cnt = [], 0
                for k, n in dallar:
                    if cur and cnt + n > PARCA_BOYU:
                        parcalar.append((cur, [], max(1, args.parti * PARCA_BOYU // max(cnt, 1) // 2) if cnt > PARCA_BOYU else args.parti)); cur, cnt = [], 0
                    cur.append(k); cnt += n
                if cur:
                    parcalar.append((cur, [], args.parti))
            satirlar, kalemler, toplam = [], {}, 0.0
            for pi, (alt, yalniz, parti) in enumerate(parcalar, 1):
                page.evaluate("__tbb.clearItems()")
                n = page.evaluate("([a, b]) => __tbb.selectItems(a, b)", [alt, yalniz])
                log(f"[{tur}] parça {pi}/{len(parcalar)}: {n} kalem seçildi")
                for i in range(0, len(donemler), parti):
                    grup = donemler[i:i + parti]
                    for deneme in range(3):
                        r = page.evaluate("(ovs) => __tbb.runBatch(ovs)", grup)
                        if not r.get("timeout"):
                            break
                        log(f"[{tur}] {grup[0]}–{grup[-1]} zaman aşımı, yeniden deneniyor ({deneme + 1}/3)")
                    else:
                        sys.exit(f"[{tur}] {grup} çekilemedi.")
                    beklenen = "KONSOL" if tur == "kons" else "SOLO"
                    if beklenen not in r["root"].upper():
                        sys.exit(f"[{tur}] Beklenmeyen tablo döndü: {r['root']}")
                    if r["unknown"]:
                        log(f"[{tur}] Eşleşmeyen banka adları: {r['unknown']}")
                    satirlar += r["lines"]
                    kalemler.update(r["items"])
                    toplam += r["sum"]
                    log(f"[{tur}] parça {pi}: {grup[0]}–{grup[-1]}: {len(r['lines']):,} değer")
            bolumler[tur] = {"satirlar": satirlar, "kalemler": kalemler, "toplam": toplam}
            log(f"[{tur}] tamam: {len(satirlar):,} değer, {len(kalemler)} kalem, kontrol toplamı {toplam:,.0f}")
        browser.close()

    tarih = dt.date.today().strftime("%d.%m.%Y")
    with open(args.cikti, "w", encoding="utf-8-sig", newline="\n") as f:
        f.write(f"# TBB TFRS9 {','.join(turler)} {donemler[0]}..{donemler[-1]} | alındı: {tarih}\n")
        f.write("#BANKS " + json.dumps(bankalar, ensure_ascii=False) + "\n")
        for tur, b in bolumler.items():
            items = [[c, v[0], v[1], v[2]] for c, v in b["kalemler"].items()]
            f.write(f"#ITEMS {tur} " + json.dumps(items, ensure_ascii=False) + "\n")
            f.write(f"#DATA {tur}\n")
            f.write("\n".join(b["satirlar"]) + "\n")
    log(f"Yazıldı: {args.cikti}")
    for tur, b in bolumler.items():
        log(f"  {tur}: kontrol toplamı (TP+YP) {b['toplam']:,.2f}")


if __name__ == "__main__":
    main()
