"""TBB banka grup kayıtlarını döker (grubu belirsiz bankalar için)."""
import json
from playwright.sync_api import sync_playwright
URL = "https://verisistemi.tbb.org.tr/index.php?/tbb/report_mali"
with sync_playwright() as p:
    b = p.chromium.launch(); pg = b.new_page(locale="tr-TR")
    pg.goto(URL, wait_until="domcontentloaded")
    pg.wait_for_function("() => window.mlt_data_bankalar && Object.keys(mlt_data_bankalar).length && window.mlt_data_bankaGruplari", timeout=120000)
    res = pg.evaluate("""() => ({gruplar: Object.values(mlt_data_bankaGruplari),
        bankalar: Object.values(mlt_data_bankalar).filter(b => /ING|Fiba|Ticaret|Pasha|Freedom|Pozitif|Vakıflar|Akbank/.test(b.TR_ADI)).map(b => ({ad: b.TR_ADI.trim(), kod: b.BANKA_KODU, gruplar: b.GRUPLAR}))})""")
    b.close()
json.dump(res, open("kesif/gruplar.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1)
