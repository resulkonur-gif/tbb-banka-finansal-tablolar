// TBB Veri Sistemi sayfasına enjekte edilen yardımcı fonksiyonlar.
// Sayfanın kendi arayüz fonksiyonlarını (maliTablolar_select, showResults, pivot tablosu) kullanır.
window.__tbb = (function () {
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const dummy = () => document.createElement("div");
  const GROUP_FIX = { "Türkiye´de Kurulmuş Yabancı Sermayeli Bankalar": "Yabancı Sermayeli Bankalar" };

  function ready() {
    try {
      return !!(window.mlt_data_donemler && Object.keys(mlt_data_donemler).length &&
        window.mlt_data_bankalar && Object.keys(mlt_data_bankalar).length &&
        jQuery("#maliTablolarTree").dxTreeView("instance").option("items").length);
    } catch (e) { return false; }
  }

  function init() {
    selected.TP = true; selected.YP = true; selected.TOPLAM = true; selected.birim = 1000; // bin TL
    return true;
  }

  // Bankaları ve bugünkü (en ayrıntılı) TBB grubunu döndürür
  function banks() {
    const G = {}; Object.values(mlt_data_bankaGruplari).forEach(g => G[g.GRUP_NO] = g);
    const depth = no => { let d = 0, g = G[no]; while (g && g.UST_GRUP_NO && G[g.UST_GRUP_NO] && d < 10) { d++; g = G[g.UST_GRUP_NO]; } return d; };
    const now = new Date().toISOString();
    return Object.values(mlt_data_bankalar).map(b => {
      const cur = (b.GRUPLAR || []).filter(g => g.GIRIS_TARIHI <= now && g.CIKIS_TARIHI > now && G[+g.GRUP_NO] && +g.GRUP_NO !== 0);
      cur.sort((a, c) => depth(+c.GRUP_NO) - depth(+a.GRUP_NO));
      let name = cur.length ? G[+cur[0].GRUP_NO].TR_ADI.trim() : "";
      name = GROUP_FIX[name] || name || "Grup bilgisi yok";
      return [b.BANKA_KODU, b.TR_ADI.trim(), name];
    });
  }

  function periods() {
    return Object.values(mlt_data_donemler).map(x => x.ORDER_VAL).sort((a, b) => a - b);
  }

  async function clearItems() {
    const keys = selected.maliTablolarArr.map(x => x.UNIQUE_KEY);
    for (const k of keys) { try { removeFromBasket_maliTablo(k); } catch (e) {} }
    await sleep(500);
    return selected.maliTablolarArr.length;
  }

  // Sepette zaten olan kalem yeniden seçilmez (alt kalemleriyle seçilen bir dalın içindekiler)
  const inBasket = k => selected.maliTablolarArr.some(x => +x.UNIQUE_KEY === +k);
  async function settle() { let n = -1; for (let i = 0; i < 40 && n !== selected.maliTablolarArr.length; i++) { n = selected.maliTablolarArr.length; await sleep(500); } }
  async function selectItems(withChildren, onlyThis) {
    for (const k of withChildren) { if (inBasket(k)) continue; maliTablolar_select(1, k, dummy()); await sleep(400); await settle(); }
    for (const k of onlyThis) { if (inBasket(k)) continue; maliTablolar_select(0, k, dummy()); await sleep(250); }
    await settle();
    return selected.maliTablolarArr.length;
  }

  function setPeriods(ovs) {
    const keys = {}; Object.values(mlt_data_donemler).forEach(x => keys[x.ORDER_VAL] = x.KEY);
    const a = selected.donemlerIds; a.length = 0; ovs.forEach(o => a.push(keys[o]));
    const b = selected.bankalarIds; b.length = 0; Object.values(mlt_data_bankalar).forEach(x => b.push(x.BANKA_KODU));
  }

  function pivotData() {
    try { return jQuery("#pivotBanka1").dxPivotGrid("instance").getDataSource().getData(); } catch (e) { return null; }
  }

  // Bir dönem grubunu çalıştırır, pivot tablosunu okur.
  async function runBatch(ovs, timeoutMs) {
    init();
    const d0 = pivotData();
    setPeriods(ovs);
    showResults("pivot");
    const t0 = Date.now();
    while (Date.now() - t0 < (timeoutMs || 180000)) {
      await sleep(1500);
      const d = pivotData();
      if (!d || d === d0 || !d.columns || !d.rows || !d.rows.length) continue;
      const have = new Set(); d.columns.forEach(y => (y.children || []).forEach(m => have.add(y.value * 100 + m.value)));
      if (!(ovs.every(o => have.has(o)) && have.size === ovs.length)) continue;
      await sleep(1000);
      return read(pivotData(), ovs);
    }
    return { timeout: true };
  }

  function read(d, ovs) {
    const nameToCode = {}; Object.values(mlt_data_bankalar).forEach(b => nameToCode[b.TR_ADI.trim()] = b.BANKA_KODU);
    const tree = jQuery("#maliTablolarTree").dxTreeView("instance").option("items"); const byKey = {}; tree.forEach(x => byKey[x.UNIQUE_KEY] = x);
    const col = {}; d.columns.forEach(y => (y.children || []).forEach(m => col[y.value * 100 + m.value] = m.index));
    const lines = [], items = {}, unknown = new Set(); let sum = 0;
    d.rows[0].children.forEach(r => {
      const m = r.text.match(/\((\d+)\)\s*$/); if (!m) return;
      const code = m[1], t = byKey[+code] || {};
      items[code] = [r.text.replace(/^-+>\s*/, "").replace(/\s*\(\d+\)\s*$/, "").trim(), t.UST_UK || null, t.SORT_KEY || null];
      (r.children || []).forEach(b => {
        const bk = nameToCode[b.text.trim()]; if (bk == null) { unknown.add(b.text); return; }
        ovs.forEach(o => {
          const v = d.values[b.index] && d.values[b.index][col[o]]; if (!v) return;
          const tp = v[0] || 0, yp = v[1] || 0; if (!tp && !yp) return;
          lines.push(code + "," + bk + "," + o + "," + tp + "," + yp); sum += tp + yp;
        });
      });
    });
    // Değeri olmayan başlık kalemleri pivotta görünmez; ağaç yapısı bozulmasın diye seçili tüm kalemler eklenir
    selected.maliTablolarArr.forEach(x => { const k = +x.UNIQUE_KEY, t = byKey[k]; if (!t || items[k]) return;
      items[k] = [String(t.TR_ADI || "").replace(/\s*\(\d+\)\s*$/, "").trim(), t.UST_UK || null, t.SORT_KEY || null]; });
    return { lines, items, sum, root: d.rows[0].text.trim(), unknown: [...unknown] };
  }

  // Bir kökün doğrudan alt dalları ve her dalın (kendisi dahil) kalem sayısı
  function children(root) {
    const t = jQuery("#maliTablolarTree").dxTreeView("instance").option("items"), kids = {};
    t.forEach(x => { (kids[x.UST_UK] = kids[x.UST_UK] || []).push(x.UNIQUE_KEY); });
    const cnt = k => 1 + (kids[k] || []).reduce((s, c) => s + cnt(c), 0);
    return (kids[root] || []).map(k => [k, cnt(k)]);
  }

  return { ready, init, banks, periods, clearItems, selectItems, runBatch, children };
})();
