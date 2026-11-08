"""TBB mali tablo ağacından dipnot (9. tablo) ve mevduat ile ilgili kalemleri çıkarır."""
import json, sys
from playwright.sync_api import sync_playwright
URL = "https://verisistemi.tbb.org.tr/index.php?/tbb/report_mali"
with sync_playwright() as p:
    b = p.chromium.launch(); pg = b.new_page(locale="tr-TR")
    pg.goto(URL, wait_until="domcontentloaded")
    pg.wait_for_function("""() => { try { return jQuery('#maliTablolarTree').dxTreeView('instance').option('items').length > 0; } catch(e) { return false; } }""", timeout=120000)
    tree = pg.evaluate("""() => jQuery('#maliTablolarTree').dxTreeView('instance').option('items').map(x => [x.UNIQUE_KEY, x.TR_ADI, x.UST_UK, x.SORT_KEY, x.ROOT_KEY])""")
    b.close()
kids = {}
for k, n, u, s, r in tree: kids.setdefault(u, []).append((k, n, u, s, r))
def desc(root, depth=0, out=None):
    out = out if out is not None else []
    for k, n, u, s, r in sorted(kids.get(root, []), key=lambda x: str(x[3])):
        out.append({"kod": k, "ad": n, "ust": u, "derinlik": depth, "sira": s}); desc(k, depth + 1, out)
    return out
roots = [x for x in tree if x[2] in (None, 0)]
res = {"kokler": [[k, n] for k, n, *_ in roots],
       "solo_dipnot": desc(102561), "kons_dipnot": desc(352561),
       "mevduat_gecen": [{"kod": k, "ad": n, "ust": u, "kok": r} for k, n, u, s, r in tree if "evduat" in (n or "") and r in (102080, 352080)]}
json.dump(res, open("kesif/agac.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(len(tree), len(res["solo_dipnot"]), len(res["kons_dipnot"]), len(res["mevduat_gecen"]))
