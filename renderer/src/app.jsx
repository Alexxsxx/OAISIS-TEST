import React, { useMemo, useState, useEffect, useRef } from "react";
import { jsPDF } from "jspdf";
import * as XLSX from "xlsx";

// Local storage keys
const STORAGE_KEY = "dot_drapery_projects_v8";
const CATALOG_KEY = "dot_vendor_catalog_v1";

// Small helpers
const fmt = (n) => (isFinite(n) ? n.toFixed(1) : "-");
const money = (n) => (isFinite(n) ? `$${n.toFixed(2)}` : "—");
const parseMoneyNum = (s) => {
  if (s === undefined || s === null || s === "") return NaN;
  const n = parseFloat(String(s).replace(/[^0-9.]/g, ""));
  return isFinite(n) ? n : NaN;
};
const ceilUp = (n, dec = 2) => {
  const f = Math.pow(10, dec);
  return Math.ceil((n + 1e-9) * f) / f;
};
const metersToYards = (m) => m * 1.0936133;

// Import helpers
function getFirst(obj, names) {
  for (const n of names) {
    if (obj[n] !== undefined) return obj[n];
    const k = Object.keys(obj).find(
      (key) => key.trim().toLowerCase() === String(n).trim().toLowerCase()
    );
    if (k) return obj[k];
  }
  return "";
}
function rowToEntryFromGenericRow(row) {
  const id = String(getFirst(row, ["id", "sku", "item #", "pattern #", "code"])) || "";
  const name = String(getFirst(row, ["name", "color", "colorway", "description"])) || "";
  const supplier = String(getFirst(row, ["supplier", "brand", "vendor"])) || "";
  const width_in = parseFloat(getFirst(row, ["width_in", "width (in)", "width inches", "usable width"])) || 0;
  const vertical_repeat_in = parseFloat(getFirst(row, ["vertical_repeat_in", "v repeat", "vertical repeat"])) || 0;
  const horizontal_repeat_in = parseFloat(getFirst(row, ["horizontal_repeat_in", "h repeat", "horizontal repeat"])) || 0;
  const railroaded = /^true|yes|1$/i.test(String(getFirst(row, ["railroaded"])));
  const sales_unit = String(getFirst(row, ["sales_unit", "unit"])) || "yard";
  const sales_min = parseFloat(getFirst(row, ["sales_min", "min", "minimum"])) || 0;
  const sales_increment = parseFloat(getFirst(row, ["sales_increment", "increment", "step"])) || 1;
  const pricing = [];
  for (let i = 1; i <= 5; i++) {
    const source = String(getFirst(row, [`pricing_source_${i}`, `source_${i}`, i === 1 ? "pricing_source" : ""])) || "";
    const type = (String(getFirst(row, [`pricing_type_${i}`, `type_${i}`])) || "other").toLowerCase();
    const yd = getFirst(row, [`usd_per_yd_${i}`, `price_per_yd_${i}`, i === 1 ? "usd_per_yd" : ""]).toString();
    const m = getFirst(row, [`usd_per_meter_${i}`, `price_per_meter_${i}`, i === 1 ? "usd_per_meter" : ""]).toString();
    const item = { source, type };
    if (yd && !isNaN(parseFloat(yd))) item.usd_per_yd = parseFloat(yd);
    if (m && !isNaN(parseFloat(m))) item.usd_per_meter = parseFloat(m);
    if (item.source || item.usd_per_yd || item.usd_per_meter) pricing.push(item);
  }
  return {
    id: id.trim(),
    name: name.trim(),
    supplier: supplier.trim(),
    width_in,
    vertical_repeat_in,
    horizontal_repeat_in,
    railroaded,
    sales: { unit: sales_unit.trim().toLowerCase() || "yard", min: sales_min, increment: sales_increment },
    pricing,
  };
}
function parseDelimited(text, delim) {
  const lines = text.replace(/\r/g, "\n").split(/\n+/).filter((l) => l.trim().length > 0);
  if (lines.length < 2) return [];
  const headers = lines[0].split(delim).map((h) => h.trim());
  const rows = [];
  for (let i = 1; i < lines.length; i++) {
    const parts = lines[i].split(delim);
    const obj = {};
    headers.forEach((h, idx) => (obj[h] = (parts[idx] !== undefined ? parts[idx] : "").trim()));
    rows.push(rowToEntryFromGenericRow(obj));
  }
  return rows;
}

export default function App() {
  // ---- Catalog ----
  const [catalog, setCatalog] = useState({ version: "user", updated: null, entries: [] });
  const [importNote, setImportNote] = useState("");
  const fileInputRef = useRef(null);

  // Load catalog: prefer localStorage; else Electron preload; else vendor_catalog.json; else empty
  useEffect(() => {
    (async () => {
      try {
        const raw = localStorage.getItem(CATALOG_KEY);
        if (raw) {
          const parsed = JSON.parse(raw);
          if (parsed && Array.isArray(parsed.entries)) {
            setCatalog(parsed);
            return;
          }
        }
      } catch {}
      try {
        if (globalThis.qt?.loadCatalog) {
          const arr = await globalThis.qt.loadCatalog();
          if (Array.isArray(arr)) {
            setCatalog({ version: "file", updated: new Date().toISOString(), entries: arr });
            return;
          }
        }
      } catch {}
      try {
        const r = await fetch("./vendor_catalog.json");
        if (r.ok) {
          const data = await r.json();
          const entries = Array.isArray(data) ? data : Array.isArray(data.entries) ? data.entries : [];
          setCatalog({ version: "file", updated: new Date().toISOString(), entries });
          return;
        }
      } catch {}
      setCatalog({ version: "empty", updated: new Date().toISOString(), entries: [] });
    })();
  }, []);
  function setCatalogAndSave(obj) {
    setCatalog(obj);
    try {
      localStorage.setItem(CATALOG_KEY, JSON.stringify(obj));
    } catch {}
  }

  // ---- Inputs ----
  const [projectName, setProjectName] = useState("Sample Project");
  const [fabricRef, setFabricRef] = useState("");

  // Fabric specifications
  const [fabricWidth, setFabricWidth] = useState(54);
  const [verticalRepeat, setVerticalRepeat] = useState(0);
  const [horizontalRepeat, setHorizontalRepeat] = useState(0);
  const [sideHem, setSideHem] = useState(3);
  const [topAllowance, setTopAllowance] = useState(5);
  const [bottomAllowance, setBottomAllowance] = useState(8);
  const [panels, setPanels] = useState(2);
  const [panelPercents, setPanelPercents] = useState([50, 50]);

  // Project dimensions
  const [rodWidth, setRodWidth] = useState(100);
  const [returnDepth, setReturnDepth] = useState(3.5);
  const [overlap, setOverlap] = useState(3.5);
  const [finishedLength, setFinishedLength] = useState(96);

  // Pleat specs
  const [headerStyle, setHeaderStyle] = useState("Double Pinch Pleat");
  const [pleatSpacing, setPleatSpacing] = useState(4);
  const [pleatSize, setPleatSize] = useState(4);

  // Pattern & waste (waste hidden; default 15%)
  const [patternMatching, setPatternMatching] = useState(false);
  const [wastePct, setWastePct] = useState(15);

  // Sales unit constraints
  const [salesUnit, setSalesUnit] = useState("yard");
  const [minQty, setMinQty] = useState(2);
  const [increment, setIncrement] = useState(1);
  const [rollSizes, setRollSizes] = useState("");

  // Costing inputs
  const [laborRetailPerWidth, setLaborRetailPerWidth] = useState(260);
  const [laborWholesalePerWidth, setLaborWholesalePerWidth] = useState(125);
  const [mainFabricCostYd, setMainFabricCostYd] = useState("");
  const [liningCostYd, setLiningCostYd] = useState("");

  // Saved projects
  const [saved, setSaved] = useState([]);
  const [status, setStatus] = useState("");

  // Load/Save projects to localStorage
  useEffect(() => {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) setSaved(JSON.parse(raw));
    } catch {}
  }, []);
  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(saved));
    } catch {}
  }, [saved]);

  // Catalog helpers
  function cheapestUsdPerYard(entry) {
    if (!entry || !Array.isArray(entry.pricing)) return NaN;
    const vals = entry.pricing
      .map((p) => {
        if (typeof p.usd_per_yd === "number") return p.usd_per_yd;
        if (typeof p.usd_per_meter === "number") return p.usd_per_meter * 1.0936133;
        return NaN;
      })
      .filter((x) => isFinite(x));
    return vals.length ? Math.min(...vals) : NaN;
  }
  function findCatalogEntry(ref) {
    const entries = catalog?.entries || [];
    const q = String(ref || "").trim().toLowerCase();
    if (!q) return null;
    let hit = entries.find((e) => (e.id || "").toLowerCase() === q);
    if (hit) return hit;
    hit = entries.find((e) => (e.name || "").toLowerCase() === q);
    if (hit) return hit;
    const pool = entries.filter(
      (e) => (e.id || "").toLowerCase().includes(q) || (e.name || "").toLowerCase().includes(q)
    );
    return pool.length ? pool[0] : null;
  }
  function applyCatalog(entry) {
    if (!entry) return;
    if (entry.width_in) setFabricWidth(entry.width_in);
    if (typeof entry.vertical_repeat_in === "number") setVerticalRepeat(entry.vertical_repeat_in);
    if (typeof entry.horizontal_repeat_in === "number") setHorizontalRepeat(entry.horizontal_repeat_in);
    if (entry.sales) {
      if (entry.sales.unit) setSalesUnit(entry.sales.unit);
      if (typeof entry.sales.min === "number") setMinQty(entry.sales.min);
      if (typeof entry.sales.increment === "number") setIncrement(entry.sales.increment);
    }
    const usd = cheapestUsdPerYard(entry);
    if (isFinite(usd)) setMainFabricCostYd(String(usd));
    setStatus(`Auto-filled from catalog: ${entry.name || entry.id} — ${entry.supplier || ""}`);
  }
  function handleAutofill() {
    const entry = findCatalogEntry(fabricRef);
    if (entry) applyCatalog(entry);
    else setStatus("No catalog match found.");
  }
  function onFabricBlur() {
    const entry = findCatalogEntry(fabricRef);
    if (entry) applyCatalog(entry);
  }

  // Importers
  function normalizeEntry(e) {
    return e && e.id && e.sales && Array.isArray(e.pricing) ? e : rowToEntryFromGenericRow(e || {});
  }
  async function importCatalogFiles(files) {
    let added = 0,
      updated = 0,
      skipped = 0;
    const warnings = [];
    const map = new Map((catalog.entries || []).map((e) => [(e.id || "").toLowerCase(), e]));
    for (const file of files) {
      const ext = (file.name.split(".").pop() || "").toLowerCase();
      try {
        if (ext === "json") {
          const txt = await file.text();
          const obj = JSON.parse(txt);
          const arr = Array.isArray(obj) ? obj : Array.isArray(obj.entries) ? obj.entries : [];
          for (const r of arr) {
            const e = normalizeEntry(r);
            const k = (e.id || "").toLowerCase();
            if (!k) {
              skipped++;
              continue;
            }
            if (map.has(k)) {
              map.set(k, { ...map.get(k), ...e });
              updated++;
            } else {
              map.set(k, e);
              added++;
            }
          }
        } else if (ext === "csv" || ext === "tsv") {
          const txt = await file.text();
          const rows = parseDelimited(txt, ext === "tsv" ? "\t" : ",");
          for (const e of rows) {
            const k = (e.id || "").toLowerCase();
            if (!k) {
              skipped++;
              continue;
            }
            if (map.has(k)) {
              map.set(k, { ...map.get(k), ...e });
              updated++;
            } else {
              map.set(k, e);
              added++;
            }
          }
        } else if (ext === "xlsx" || ext === "xls" || ext === "ods") {
          const buf = await file.arrayBuffer();
          const wb = XLSX.read(buf, { type: "array" });
          const ws = wb.Sheets[wb.SheetNames[0]];
          const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
          for (const row of rows) {
            const e = rowToEntryFromGenericRow(row);
            const k = (e.id || "").toLowerCase();
            if (!k) {
              skipped++;
              continue;
            }
            if (map.has(k)) {
              map.set(k, { ...map.get(k), ...e });
              updated++;
            } else {
              map.set(k, e);
              added++;
            }
          }
        } else {
          warnings.push(`${file.name}: Unsupported file type`);
        }
      } catch (err) {
        warnings.push(`${file.name}: ${err?.message || "parse error"}`);
      }
    }
    const merged = { version: "user", updated: new Date().toISOString(), entries: [...map.values()] };
    setCatalogAndSave(merged);
    setImportNote(
      `Imported +${added}, updated ${updated}, skipped ${skipped}. ${
        warnings.length ? "Warnings: " + warnings.join(" | ") : ""
      }`
    );
  }
  function openFilePicker() {
    fileInputRef.current?.click();
  }
  async function onFilesPicked(e) {
    const files = Array.from(e.target.files || []);
    if (files.length) {
      await importCatalogFiles(files);
      e.target.value = "";
    }
  }
  function downloadCatalog() {
    try {
      const blob = new Blob([JSON.stringify(catalog, null, 2)], { type: "application/json" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "vendor_catalog.json";
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
      setImportNote("Catalog downloaded.");
    } catch {}
  }

  // Calculations (pleat-driven + seam-aware)
  const faceToCover = useMemo(
    () => rodWidth + returnDepth + returnDepth + overlap,
    [rodWidth, returnDepth, overlap]
  );
  const baseCutLength = useMemo(
    () => finishedLength + topAllowance + bottomAllowance,
    [finishedLength, topAllowance, bottomAllowance]
  );
  const cutLength = useMemo(() => {
    if (!patternMatching || verticalRepeat <= 0) return baseCutLength;
    const multiples = Math.ceil(baseCutLength / verticalRepeat);
    return multiples * verticalRepeat;
  }, [patternMatching, verticalRepeat, baseCutLength]);

  // Panels % distribution
  useEffect(() => {
    const n = Math.max(1, Math.floor(panels || 1));
    if (panelPercents.length !== n) {
      const equal = Array.from({ length: n }, () => +(100 / n).toFixed(2));
      const sumFirst = equal.slice(0, n - 1).reduce((a, b) => a + b, 0);
      equal[n - 1] = +(100 - sumFirst).toFixed(2);
      setPanelPercents(equal);
    }
  }, [panels]);
  const percentSum = useMemo(
    () => panelPercents.reduce((a, b) => a + (parseFloat(b) || 0), 0),
    [panelPercents]
  );
  const percentsValid = panels <= 1 || Math.abs(percentSum - 100) < 1e-6;
  const shares = useMemo(() => {
    const n = Math.max(1, Math.floor(panels || 1));
    if (n === 1) return [1];
    if (percentsValid) return panelPercents.slice(0, n).map((p) => (parseFloat(p) || 0) / 100);
    return Array.from({ length: n }, () => 1 / n);
  }, [panels, panelPercents, percentsValid]);

  const perPanel = useMemo(() => {
    const STITCH_LOSS = 1.0; // 1" per interior join
    const n = shares.length;
    const arr = [];
    for (let i = 0; i < n; i++) {
      const finishedWidthForPanel = faceToCover * shares[i];
      const spacing = Math.max(pleatSpacing || 0, 0.0001);
      const pleats = Math.max(1, Math.round(finishedWidthForPanel / spacing));
      const takeup = Math.max(0, pleatSize || 0);
      const requiredFlat = finishedWidthForPanel + pleats * takeup;
      const denom = Math.max(0.0001, fabricWidth - STITCH_LOSS);
      const widths = Math.max(1, Math.ceil((requiredFlat + 2 * sideHem - STITCH_LOSS) / denom));
      const yards = (widths * cutLength) / 36;
      arr.push({ finishedWidthForPanel, pleatCount: pleats, requiredFlat, widths, yards });
    }
    return arr;
  }, [shares, faceToCover, pleatSpacing, pleatSize, fabricWidth, sideHem, cutLength]);

  const widthsCount = useMemo(() => perPanel.reduce((a, p) => a + p.widths, 0), [perPanel]);
  const totalYards = useMemo(() => perPanel.reduce((a, p) => a + p.yards, 0), [perPanel]);
  const totalWithWaste = useMemo(() => totalYards * (1 + (wastePct || 0) / 100), [totalYards, wastePct]);
  const avgYardsPerPanel = useMemo(() => totalYards / shares.length, [totalYards, shares.length]);
  const avgYardsPerPanelWithWaste = useMemo(
    () => totalWithWaste / shares.length,
    [totalWithWaste, shares.length]
  );

  function computeOrderedYards() {
    const needYards = totalWithWaste;
    if (salesUnit === "yard") {
      const inc = Math.max(0.001, increment);
      const min = Math.max(0, minQty);
      const qty = Math.max(min, Math.ceil(needYards / inc) * inc);
      return ceilUp(qty, 2);
    }
    if (salesUnit === "meter") {
      const inc = Math.max(0.001, increment);
      const min = Math.max(0, minQty);
      const needMeters = needYards / 1.0936133;
      const metersOrdered = Math.max(min, Math.ceil(needMeters / inc) * inc);
      const yardsOrdered = metersToYards(metersOrdered);
      return ceilUp(yardsOrdered, 2);
    }
    return ceilUp(needYards, 2);
  }
  const orderedYards = useMemo(() => computeOrderedYards(), [totalWithWaste, salesUnit, minQty, increment]);
  const liningYards = useMemo(() => ceilUp(totalYards, 2), [totalYards]);

  // Totals
  const retailLaborTotal = useMemo(
    () => widthsCount * (parseFloat(laborRetailPerWidth) || 0),
    [widthsCount, laborRetailPerWidth]
  );
  const wholesaleLaborTotal = useMemo(
    () => widthsCount * (parseFloat(laborWholesalePerWidth) || 0),
    [widthsCount, laborWholesalePerWidth]
  );
  const mainCostPerYd = useMemo(() => parseMoneyNum(mainFabricCostYd), [mainFabricCostYd]);
  const liningCostPerYd = useMemo(() => parseMoneyNum(liningCostYd), [liningCostYd]);
  const mainFabricTotal = useMemo(
    () => (isFinite(mainCostPerYd) ? ceilUp(orderedYards, 2) * mainCostPerYd : NaN),
    [orderedYards, mainCostPerYd]
  );
  const liningTotal = useMemo(
    () => (isFinite(liningCostPerYd) ? liningYards * liningCostPerYd : NaN),
    [liningYards, liningCostPerYd]
  );

  // Saved projects
  function snapshot() {
    return {
      id: Date.now(),
      timestamp: new Date().toISOString(),
      projectName,
      fabricRef,
      inputs: {
        fabricWidth,
        verticalRepeat,
        horizontalRepeat,
        sideHem,
        topAllowance,
        bottomAllowance,
        panels,
        panelPercents,
        rodWidth,
        returnDepth,
        overlap,
        finishedLength,
        headerStyle,
        pleatSpacing,
        pleatSize,
        patternMatching,
        wastePct,
        salesUnit,
        minQty,
        increment,
        rollSizes,
        laborRetailPerWidth,
        laborWholesalePerWidth,
        mainFabricCostYd,
        liningCostYd,
      },
      computed: {
        faceToCover,
        cutLength,
        widthsCount,
        perPanel,
        totalYards,
        totalWithWaste,
        avgYardsPerPanel,
        avgYardsPerPanelWithWaste,
        orderedYards,
        liningYards,
        retailLaborTotal,
        wholesaleLaborTotal,
        mainFabricTotal,
        liningTotal,
      },
    };
  }
  function saveProject() {
    setSaved((prev) => [snapshot(), ...prev].slice(0, 100));
    setStatus("Project saved.");
  }
  function copySummary(p) {
    const item = p || snapshot();
    const c = item.computed || {};
    const w = item.inputs?.wastePct ?? 0;
    const text = [
      `Project: ${item.projectName}`,
      `Fabric: ${item.fabricRef}`,
      `Panels: ${item.inputs?.panels}`,
      `Widths total: ${c.widthsCount}`,
      `Avg yards per panel: ${fmt(c.avgYardsPerPanel)}`,
      `Avg yards per panel (with waste ${w}%): ${fmt(c.avgYardsPerPanelWithWaste)}`,
      `Total: ${fmt(c.totalYards)} yards`,
      `Total with waste (${w}%): ${fmt(c.totalWithWaste)} yards`,
      `Ordered yards (face): ${isFinite(c.orderedYards) ? c.orderedYards.toFixed(2) : "-" } yards`,
      `Lining yards (0% waste): ${isFinite(c.liningYards) ? c.liningYards.toFixed(2) : "-" } yards`,
    ].join("\n");
    navigator.clipboard?.writeText(text).then(
      () => setStatus("Summary copied to clipboard."),
      () => setStatus(text)
    );
  }

  // PDF
  function buildQuoteLines() {
    const lines = [];
    lines.push(`Quote — ${projectName || "Untitled"}`);
    lines.push(`Fabric: ${fabricRef || "—"}`);
    lines.push("");
    lines.push("Fabric Details");
    lines.push(`• Fabric Width: ${fabricWidth}"`);
    lines.push(`• Vertical Repeat: ${verticalRepeat}"`);
    lines.push(`• Horizontal Repeat: ${horizontalRepeat}"`);
    lines.push(`• Allowances: Top ${topAllowance}" · Bottom ${bottomAllowance}" · Side hems ${sideHem}"`);
    lines.push(`• Panels: ${panels}`);
    lines.push("");
    lines.push("Yardage Calculation");
    lines.push(`• Face width total: ${faceToCover.toFixed(2)} in`);
    lines.push(`• Seam model: fabric width ${fabricWidth}" · side hems ${sideHem}" each (2 total) · interior join loss 1"`);
    lines.push(`• Cut length ${patternMatching && verticalRepeat > 0 ? "(pattern matching)" : "(no pattern matching)"}: ${cutLength.toFixed(2)} in`);
    lines.push(`• Widths per panel (sum): ${widthsCount}`);
    perPanel.forEach((p, i) =>
      lines.push(
        `  ◦ Panel ${i + 1}: finished width ${p.finishedWidthForPanel.toFixed(2)} in · pleats ${p.pleatCount} · required flat ${fmt(
          p.requiredFlat
        )} in · widths ${p.widths} · yards ${fmt(p.yards)}`
      )
    );
    lines.push("");
    lines.push("Yardage Summary");
    lines.push(`• Avg yards per panel: ${fmt(avgYardsPerPanel)}`);
    lines.push(`• Avg yards per panel with waste (${wastePct}%): ${fmt(avgYardsPerPanelWithWaste)}`);
    lines.push(`• Total: ${fmt(totalYards)} yd`);
    lines.push(`• Total with waste (${wastePct}%): ${fmt(totalWithWaste)} yd`);
    lines.push("");
    lines.push("Costing Summary");
    lines.push(`• Widths count: ${widthsCount}`);
    lines.push(`• Ordered yards (face, rounded up): ${orderedYards.toFixed(2)} yd`);
    lines.push(`• Lining yards (0% waste, rounded up): ${liningYards.toFixed(2)} yd`);
    lines.push(`• Labor total (Retail): ${money(retailLaborTotal)}`);
    lines.push(`• Labor total (Wholesale): ${money(wholesaleLaborTotal)}`);
    lines.push(`• Main fabric total: ${money(mainFabricTotal)}`);
    lines.push(`• Lining total: ${money(liningTotal)}`);
    lines.push("");
    lines.push("Generated by Dot — your virtual quoting assistant · Developed by Ottomatic AI Solutions");
    return lines;
  }
  function generatePdf() {
    const doc = new jsPDF({ unit: "pt", format: "letter" });
    const marginX = 40;
    const marginY = 48;
    const maxWidth = 612 - marginX * 2;
    doc.setFontSize(14);
    const heading = `Quote — ${projectName || "Untitled"}`;
    doc.text(heading, marginX, marginY);
    doc.setFontSize(10);
    const sub = `Fabric: ${fabricRef || "—"}  •  ${new Date().toLocaleString()}`;
    doc.text(sub, marginX, marginY + 16);
    doc.setFontSize(11);
    let y = marginY + 40;
    const lines = buildQuoteLines();
    for (const raw of lines) {
      const wrapped = doc.splitTextToSize(raw, maxWidth);
      for (const line of wrapped) {
        if (y > 760) {
          doc.addPage();
          y = marginY;
        }
        doc.text(line, marginX, y);
        y += 14;
      }
      y += 6;
    }
    const safeProject = (projectName || "quote").replace(/[^a-z0-9\-\_]+/gi, "_");
    const filename = `${safeProject}.pdf`;
    doc.save(filename);
    setStatus("PDF downloaded.");
  }

  // UI
  return (
    <div className="min-h-screen bg-neutral-50 text-neutral-900 p-6">
      <div className="max-w-6xl mx-auto grid md:grid-cols-2 gap-6">
        {/* LEFT */}
        <div className="bg-white rounded-2xl shadow p-5 space-y-5">
          <div className="flex items-center justify-between">
            <div>
              <h1 className="text-2xl font-semibold">Prototype Order Calculator</h1>
              <div className="text-xs text-neutral-500 mt-1">Developed by Ottomatic AI Solutions</div>
            </div>
            <span className="text-xs text-neutral-500">Dot • Live</span>
          </div>

          {/* Project Info */}
          <div>
            <h2 className="text-lg font-medium">Project Info</h2>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mt-2 items-end">
              <label className="block">
                <span className="text-sm">Project Name</span>
                <input className="mt-1 w-full rounded-xl border p-2" value={projectName} onChange={(e) => setProjectName(e.target.value)} />
              </label>
              <label className="block">
                <span className="text-sm">Fabric Used?</span>
                <div className="mt-1 flex gap-2">
                  <input
                    className={`w-full rounded-xl border p-2 ${fabricRef.trim() ? (findCatalogEntry(fabricRef) ? "border-green-500" : "border-red-500") : ""}`}
                    placeholder="ID or Name"
                    value={fabricRef}
                    onChange={(e) => setFabricRef(e.target.value)}
                    onBlur={onFabricBlur}
                  />
                  <button onClick={handleAutofill} className="px-3 py-2 rounded-xl border shadow-sm text-sm shrink-0">
                    Auto-fill
                  </button>
                </div>
                {fabricRef.trim() && !findCatalogEntry(fabricRef) ? (
                  <div className="text-xs text-red-600 italic mt-1">*Fabric Missing From Catalogue*</div>
                ) : null}
              </label>
            </div>
          </div>

          {/* Fabric Specifications */}
          <div className="pt-2">
            <h2 className="text-lg font-medium">Fabric Specifications</h2>
            <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mt-2 items-end">
              <label className="block">
                <span className="text-sm">Fabric Width (inches)</span>
                <input type="number" className="mt-1 w-full rounded-xl border p-2" value={fabricWidth} onChange={(e) => setFabricWidth(parseFloat(e.target.value) || 0)} />
              </label>
              <label className="block">
                <span className="text-sm">Vertical Repeat (inches)</span>
                <input type="number" className="mt-1 w-full rounded-xl border p-2" value={verticalRepeat} onChange={(e) => setVerticalRepeat(parseFloat(e.target.value) || 0)} />
              </label>
              <label className="block">
                <span className="text-sm">Horizontal Repeat (inches)</span>
                <input type="number" className="mt-1 w-full rounded-xl border p-2" value={horizontalRepeat} onChange={(e) => setHorizontalRepeat(parseFloat(e.target.value) || 0)} />
              </label>
              <label className="block">
                <span className="text-sm">Side Hem (inches per side)</span>
                <input type="number" className="mt-1 w-full rounded-xl border p-2" value={sideHem} onChange={(e) => setSideHem(parseFloat(e.target.value) || 0)} />
              </label>
              <label className="block">
                <span className="text-sm">Top Allowance (inches)</span>
                <input type="number" className="mt-1 w-full rounded-xl border p-2" value={topAllowance} onChange={(e) => setTopAllowance(parseFloat(e.target.value) || 0)} />
              </label>
              <label className="block">
                <span className="text-sm">Bottom Allowance (inches)</span>
                <input type="number" className="mt-1 w-full rounded-xl border p-2" value={bottomAllowance} onChange={(e) => setBottomAllowance(parseFloat(e.target.value) || 0)} />
              </label>
              <label className="block">
                <span className="text-sm">Panels (number)</span>
                <input
                  type="number"
                  min={1}
                  className="mt-1 w-full rounded-xl border p-2"
                  value={panels}
                  onChange={(e) => setPanels(Math.max(1, Math.floor(parseFloat(e.target.value) || 1)))}
                />
              </label>
            </div>
          </div>

          {/* Panel Specifications */}
          {panels > 1 && (
            <div className="pt-2">
              <div className="flex items-center gap-2">
                <h2 className="text-lg font-medium">Panel Specifications</h2>
                {!(panels <= 1 || Math.abs(panelPercents.reduce((a, b) => a + (parseFloat(b) || 0), 0) - 100) < 1e-6) && (
                  <span className="text-red-600 text-sm">Fields must equal 100%</span>
                )}
              </div>
              <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mt-2 items-end">
                {Array.from({ length: panels }).map((_, i) => (
                  <label key={i} className="block">
                    <span className="text-sm">Panel {i + 1} — Fabric Percentage</span>
                    <input
                      type="number"
                      className="mt-1 w-full rounded-xl border p-2"
                      value={panelPercents[i] ?? 0}
                      onChange={(e) => {
                        const v = parseFloat(e.target.value);
                        setPanelPercents((prev) => {
                          const next = prev.slice(0, panels);
                          next[i] = isFinite(v) ? v : 0;
                          return next;
                        });
                      }}
                    />
                  </label>
                ))}
              </div>
              <div className="text-xs text-neutral-500 mt-1">Percentages may be uneven, but must total 100%.</div>
            </div>
          )}

          {/* Project Dimensions */}
          <div className="pt-2">
            <h2 className="text-lg font-medium">Project Dimensions</h2>
            <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mt-2 items-end">
              <label className="block">
                <span className="text-sm">Rod or Track Width (inches)</span>
                <input type="number" className="mt-1 w-full rounded-xl border p-2" value={rodWidth} onChange={(e) => setRodWidth(parseFloat(e.target.value) || 0)} />
              </label>
              <label className="block">
                <span className="text-sm">Returns (inches, each side)</span>
                <input type="number" className="mt-1 w-full rounded-xl border p-2" value={returnDepth} onChange={(e) => setReturnDepth(parseFloat(e.target.value) || 0)} />
              </label>
              <label className="block">
                <span className="text-sm">Overlap (inches, center)</span>
                <input type="number" className="mt-1 w-full rounded-xl border p-2" value={overlap} onChange={(e) => setOverlap(parseFloat(e.target.value) || 0)} />
              </label>
              <label className="block">
                <span className="text-sm">Finished Length (inches)</span>
                <input type="number" className="mt-1 w-full rounded-xl border p-2" value={finishedLength} onChange={(e) => setFinishedLength(parseFloat(e.target.value) || 0)} />
              </label>
            </div>
          </div>

          {/* Pleat Specifications */}
          <div className="pt-2">
            <h2 className="text-lg font-medium">Pleat Specifications</h2>
            <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mt-2 items-end">
              <label className="block md:col-span-2">
                <span className="text-sm">Pleat Type</span>
                <select className="mt-1 w-full rounded-xl border p-2" value={headerStyle} onChange={(e) => setHeaderStyle(e.target.value)}>
                  <option>Double Pinch Pleat</option>
                  <option>Triple Pinch Pleat</option>
                  <option>Euro Pleat</option>
                  <option>Box Pleat</option>
                </select>
              </label>
              <label className="block">
                <span className="text-sm">Spacing (inches between pleats)</span>
                <input type="number" className="mt-1 w-full rounded-xl border p-2" value={pleatSpacing} onChange={(e) => setPleatSpacing(parseFloat(e.target.value) || 0)} />
              </label>
              <label className="block">
                <span className="text-sm">Pleat Size (inches of fabric per pleat)</span>
                <input type="number" className="mt-1 w-full rounded-xl border p-2" value={pleatSize} onChange={(e) => setPleatSize(parseFloat(e.target.value) || 0)} />
              </label>
            </div>
          </div>

          {/* Pattern */}
          <div className="pt-2 grid grid-cols-1 md:grid-cols-3 gap-3 items-end">
            <label className="inline-flex items-center gap-2">
              <input type="checkbox" className="rounded" checked={patternMatching} onChange={(e) => setPatternMatching(e.target.checked)} />
              <span className="text-sm">Pattern Matching (use Vertical Repeat)</span>
            </label>
          </div>

          {/* Sales Unit and Constraints */}
          <div className="pt-2">
            <h2 className="text-lg font-medium">Sales Unit and Constraints (for Ordered Size)</h2>
            <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mt-2 items-end">
              <label className="block">
                <span className="text-sm">Sales Unit</span>
                <select className="mt-1 w-full rounded-xl border p-2" value={salesUnit} onChange={(e) => setSalesUnit(e.target.value)}>
                  <option value="yard">Yard</option>
                  <option value="meter">Meter</option>
                  <option value="panel">Panel or Drop</option>
                  <option value="roll">Roll or Bolt</option>
                </select>
              </label>
              <label className="block">
                <span className="text-sm">Minimum Order (in selected unit)</span>
                <input type="number" className="mt-1 w-full rounded-xl border p-2" value={minQty} onChange={(e) => setMinQty(parseFloat(e.target.value) || 0)} />
              </label>
              <label className="block">
                <span className="text-sm">Increment (in selected unit)</span>
                <input type="number" className="mt-1 w-full rounded-xl border p-2" value={increment} onChange={(e) => setIncrement(parseFloat(e.target.value) || 0)} />
              </label>
              <label className="block">
                <span className="text-sm">Roll or Pack Sizes (comma separated)</span>
                <input className="mt-1 w-full rounded-xl border p-2" placeholder="for example, 10, 20" value={rollSizes} onChange={(e) => setRollSizes(e.target.value)} />
              </label>
            </div>
          </div>

          {/* Costing */}
          <div className="pt-2">
            <h2 className="text-lg font-medium">Costing</h2>
            <div className="rounded-2xl border p-4 mt-2">
              <div className="grid grid-cols-2 md:grid-cols-4 gap-3 items-end">
                <label className="block text-sm">
                  <span>Labor (Retail) — $/fabric width</span>
                  <input className="mt-1 w-full rounded-xl border p-2" value={laborRetailPerWidth} onChange={(e) => setLaborRetailPerWidth(parseFloat(e.target.value) || 0)} />
                </label>
                <label className="block text-sm">
                  <span>Labor (Wholesale) — $/fabric width</span>
                  <input className="mt-1 w-full rounded-xl border p-2" value={laborWholesalePerWidth} onChange={(e) => setLaborWholesalePerWidth(parseFloat(e.target.value) || 0)} />
                </label>
                <label className="block text-sm">
                  <span>Main Fabric Cost — $/yd</span>
                  <input className="mt-1 w-full rounded-xl border p-2" placeholder="for example, 45" value={mainFabricCostYd} onChange={(e) => setMainFabricCostYd(e.target.value)} />
                </label>
                <label className="block text-sm">
                  <span>Lining Cost — $/yd</span>
                  <input className="mt-1 w-full rounded-xl border p-2" placeholder="for example, 12" value={liningCostYd} onChange={(e) => setLiningCostYd(e.target.value)} />
                </label>
              </div>
            </div>
          </div>

          <div className="flex flex-wrap gap-2 pt-3">
            <button onClick={saveProject} className="px-3 py-2 rounded-xl border shadow-sm text-sm">Save Project</button>
            <button onClick={() => copySummary()} className="px-3 py-2 rounded-xl border shadow-sm text-sm">Copy Current Summary</button>
            <button onClick={generatePdf} className="px-3 py-2 rounded-xl border shadow-sm text-sm">Print Client</button>
          </div>
        </div>

        {/* RIGHT */}
        <div className="space-y-5">
          <div className="bg-white rounded-2xl shadow p-5">
            <h2 className="text-xl font-semibold">Fabric Details</h2>
            <div className="mt-3 grid grid-cols-2 md:grid-cols-3 gap-2 text-sm">
              <div><span className="text-neutral-500">Fabric Reference:</span> {fabricRef || "—"}</div>
              <div><span className="text-neutral-500">Fabric Width:</span> {fabricWidth}"</div>
              <div><span className="text-neutral-500">Vertical Repeat:</span> {verticalRepeat}"</div>
              <div><span className="text-neutral-500">Horizontal Repeat:</span> {horizontalRepeat}"</div>
              <div><span className="text-neutral-500">Allowances:</span> Top {topAllowance}" · Bottom {bottomAllowance}" · Side hems {sideHem}"</div>
              <div><span className="text-neutral-500">Panels:</span> {panels}</div>
            </div>
          </div>

          <div className="bg-white rounded-2xl shadow p-5">
            <h2 className="text-xl font-semibold">Yardage Calculation</h2>
            <div className="mt-3 text-sm grid gap-1">
              <div>1) Face width total: <span className="font-medium">{faceToCover.toFixed(2)} inches</span></div>
              <div>2) Seam model: <span className="font-medium">fabric width {fabricWidth}" · 2 side hems {sideHem}" each · 1" per interior join</span></div>
              <div>3) Cut length {patternMatching && verticalRepeat > 0 ? "(pattern matching)" : "(no pattern matching)"}: <span className="font-medium">{cutLength.toFixed(2)} inches</span></div>
              <div>4) Widths per panel (sum): <span className="font-medium">{widthsCount}</span></div>
              <div className="mt-2">
                <div className="font-medium">Per-panel breakdown:</div>
                <ul className="list-disc ml-5">
                  {perPanel.map((p, i) => (
                    <li key={i}>Panel {i + 1}: finished width {p.finishedWidthForPanel.toFixed(2)} in · pleats {p.pleatCount} · required flat {p.requiredFlat.toFixed(2)} in · widths {p.widths} · yards {fmt(p.yards)}</li>
                  ))}
                </ul>
              </div>
            </div>
          </div>

            <div className="bg-white rounded-2xl shadow p-5">
              <h2 className="text-xl font-semibold">Yardage Summary</h2>
              <div className="mt-3 text-sm grid gap-1">
                <div><span className="font-medium">Avg yards per panel:</span> {fmt(avgYardsPerPanel)} yards</div>
                <div><span className="font-medium">Avg yards per panel with waste ({wastePct}%):</span> {fmt(avgYardsPerPanelWithWaste)} yards</div>
                <div><span className="font-medium">Total:</span> {fmt(totalYards)} yards</div>
                <div><span className="font-medium">Total with waste ({wastePct}%):</span> {fmt(totalWithWaste)} yards</div>
              </div>
            </div>

          <div className="bg-white rounded-2xl shadow p-5">
            <h2 className="text-xl font-semibold">Costing Summary</h2>
            <div className="mt-3 text-sm grid gap-1">
              <div className="flex justify-between"><span className="text-neutral-500">Widths count:</span><span>{widthsCount}</span></div>
              <div className="flex justify-between"><span className="text-neutral-500">Ordered yards (face, rounded up):</span><span>{orderedYards.toFixed(2)} yd</span></div>
              <div className="flex justify-between"><span className="text-neutral-500">Lining yards (0% waste, rounded up):</span><span>{liningYards.toFixed(2)} yd</span></div>
              <div className="flex justify-between"><span className="text-neutral-500">Labor total (Retail):</span><span>{money(retailLaborTotal)}</span></div>
              <div className="flex justify-between"><span className="text-neutral-500">Labor total (Wholesale):</span><span>{money(wholesaleLaborTotal)}</span></div>
              <div className="flex justify-between"><span className="text-neutral-500">Main fabric total:</span><span>{money(mainFabricTotal)}</span></div>
              <div className="flex justify-between"><span className="text-neutral-500">Lining total:</span><span>{money(liningTotal)}</span></div>
            </div>
          </div>

          {/* Import Catalog */}
          <div className="flex justify-end">
            <input ref={fileInputRef} type="file" className="hidden" multiple accept=".csv,.tsv,.xlsx,.xls,.json,.ods" onChange={onFilesPicked} />
            <div className="text-right">
              <button onClick={openFilePicker} className="px-3 py-2 rounded-xl border shadow-sm text-sm">Import Catalog</button>
              {importNote && (
                <div className="mt-2 text-xs text-neutral-500">
                  {importNote} <button onClick={downloadCatalog} className="underline">Download catalog</button>
                </div>
              )}
            </div>
          </div>

          <div className="text-xs text-neutral-500">Calculated using Dot’s workroom defaults • Adjust any input and values update instantly.</div>
          {status && <div className="text-xs text-neutral-600 mt-2 whitespace-pre-wrap">{status}</div>}
        </div>
      </div>
    </div>
  );
}
