let lastRenderTs = 0;
const RENDER_COOLDOWN_MS = 150;

const HtmlMap = {
  default:  "./forms/default.html",
  settings: "./forms/settings.html",
  colorPalette: "./forms/colorPalette.html"
};

const SelectionRoutes = [
  { sheet: "settings", match: { address: "B3" }, form: "colorPalette" }
];

Office.onReady(async () => {
  await renderForActiveWorksheet();
  try {
    await Excel.run(async (ctx) => {
      const sheets = ctx.workbook.worksheets;
      if (sheets.onActivated && sheets.onActivated.add) {
        await sheets.onActivated.add(onWorksheetActivated);
      }
      sheets.load("items/name");
      await ctx.sync();
      for (const ws of sheets.items) {
        if (ws.onSelectionChanged && ws.onSelectionChanged.add) {
          await ws.onSelectionChanged.add(onSelectionChanged);
        }
      }
      if (sheets.onAdded && sheets.onAdded.add) {
        await sheets.onAdded.add(async (args) => {
          await Excel.run(async (inner) => {
            const newWs = inner.workbook.worksheets.getItem(args.worksheetId);
            if (newWs.onSelectionChanged && newWs.onSelectionChanged.add) {
              await newWs.onSelectionChanged.add(onSelectionChanged);
            }
            await inner.sync();
          });
        });
      }
      await ctx.sync();
    });
  } catch (e) {
    console.error("Event subscription error:", e);
  }
});

async function onWorksheetActivated() { await renderForActiveWorksheet(); }
async function onSelectionChanged() {
  const now = Date.now();
  if (now - lastRenderTs > RENDER_COOLDOWN_MS) await renderForActiveWorksheet();
}

async function renderForActiveWorksheet() {
  lastRenderTs = Date.now();
  await Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getActiveWorksheet();
    ws.load("name");
    const a1 = ws.getRange("A1"); a1.load(["text"]);
    const sel = ctx.workbook.getSelectedRange(); sel.load("address");
    await ctx.sync();

    const sheetName = (ws.name || "").trim();
    const hint = parseFormHintFromCell(a1.text);
    const baseFormId = pickFormId(hint, sheetName);

    const selectionAddress = localizeAddress(sel.address);
    const overrideFormId = pickSelectionOverride(sheetName, selectionAddress);
    const finalFormId = overrideFormId || baseFormId;

    updateSheetBadge(sheetName);
    await renderForm(finalFormId, { sheetName, hint, selectionAddress, override: !!overrideFormId });
  });
}

function parseFormHintFromCell(text2d) {
  try {
    const t = (text2d && text2d[0] && text2d[0][0]) ? String(text2d[0][0]) : "";
    const m = /form\s*:\s*([a-z0-9_-]+)/i.exec(t);
    return m ? m[1].toLowerCase() : null;
  } catch { return null; }
}

function pickFormId(hint, sheetName) {
  if (hint && HtmlMap[hint]) return hint;
  const key = (sheetName || "").toLowerCase().trim();
  if (HtmlMap[key]) return key;
  return "default";
}

function pickSelectionOverride(sheetName, selectionAddress) {
  const s = (sheetName || "").toLowerCase().trim();
  const sel = normalizeA1(selectionAddress);
  for (const rule of SelectionRoutes) {
    if ((rule.sheet || "").toLowerCase().trim() !== s) continue;
    const m = rule.match || {};
    if (m.address && containsAddress(sel, normalizeA1(m.address))) return rule.form;
  }
  return null;
}

function localizeAddress(fullAddress) {
  const parts = fullAddress.split("!");
  const local = parts.length > 1 ? parts.slice(1).join("!") : parts[0];
  return local.replace(/\$/g, "");
}
function normalizeA1(addr) { return String(addr || "").replace(/\$/g, "").toUpperCase(); }

function containsAddress(selection, target) {
  const selAreas = selection.split(",");
  const tgtAreas = target.split(",");
  for (const t of tgtAreas) {
    const tArea = parseArea(t.trim());
    let covered = false;
    for (const s of selAreas) {
      const sArea = parseArea(s.trim());
      if (areaContains(sArea, tArea)) { covered = true; break; }
    }
    if (!covered) return false;
  }
  return true;
}

function parseArea(a1) {
  const parts = a1.split(":");
  if (parts.length === 1) {
    const pt = parsePoint(parts[0]);
    return { c1: pt.c, r1: pt.r, c2: pt.c, r2: pt.r };
  }
  const p1 = parsePoint(parts[0]);
  const p2 = parsePoint(parts[1]);
  return {
    c1: Math.min(p1.c, p2.c),
    r1: Math.min(p1.r, p2.r),
    c2: Math.max(p1.c, p2.c),
    r2: Math.max(p1.r, p2.r)
  };
}

function parsePoint(p) {
  const m = /^([A-Z]+)(\d+)$/.exec(String(p).toUpperCase());
  if (!m) return { c: 1, r: 1 };
  return { c: colToNum(m[1]), r: parseInt(m[2], 10) };
}
function colToNum(col) { let n = 0; for (let i=0;i<col.length;i++) n = n*26 + (col.charCodeAt(i)-64); return n; }
function areaContains(a, b) { return a.c1 <= b.c1 && a.r1 <= b.r1 && a.c2 >= b.c2 && a.r2 >= b.r2; }

function updateSheetBadge(name) {
  const el = document.getElementById("sheetName");
  if (el) el.textContent = name || "(unknown)";
}

async function renderForm(formId, ctx) {
  const app = document.getElementById("app");
  if (!app) return;
  const url = HtmlMap[formId] || HtmlMap.default;
  app.innerHTML = `<div class="loading">Loading ${formId}…</div>`;
  const html = await (await fetch(url, { cache: "no-store" })).text();
  app.innerHTML = html;
  await wireBindings(app);
}

// ------------- Two-way bindings -------------
let sheetChangeUnsub = null;
let debounceTimer = null;

async function wireBindings(container) {
  await refreshBoundControls(container);
  container.querySelectorAll("[data-bind]").forEach((el) => {
    const handler = async () => { try { await writeBinding(el); } catch (e) { console.error(e); } };
    el.addEventListener("change", handler);
    if (el.type === "text" || el.tagName === "TEXTAREA") el.addEventListener("blur", handler);
  });

  try {
    await Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getActiveWorksheet();
      if (sheetChangeUnsub && sheetChangeUnsub.remove) {
        sheetChangeUnsub.remove(); sheetChangeUnsub = null;
      }
      if (ws.onChanged && ws.onChanged.add) {
        sheetChangeUnsub = await ws.onChanged.add(async () => {
          clearTimeout(debounceTimer);
          debounceTimer = setTimeout(() => refreshBoundControls(container), 120);
        });
      }
    });
  } catch (e) { console.warn("Worksheet.onChanged unavailable.", e); }
}

async function refreshBoundControls(container) {
  const els = [...container.querySelectorAll("[data-bind]")];
  if (els.length === 0) return;
  await Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getActiveWorksheet();
    const toLoad = [];
    for (const el of els) {
      const bind = (el.dataset.bind || "").trim();
      if (!bind) continue;
      const rng = await resolveRange(ctx, ws, bind);
      if (!rng) continue;
      rng.load("values"); toLoad.push({ el, rng });
    }
    await ctx.sync();
    for (const { el, rng } of toLoad) {
      const v = (rng.values && rng.values[0] && rng.values[0][0]) ?? "";
      setElValueFromCell(el, v);
    }
  });
}

async function writeBinding(el) {
  const bind = (el.dataset.bind || "").trim();
  if (!bind) return;
  await Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getActiveWorksheet();
    const rng = await resolveRange(ctx, ws, bind);
    if (!rng) return;
    rng.values = [[getCellValueFromEl(el)]];
    await ctx.sync();
  });
}

function setElValueFromCell(el, cellValue) {
  const t = (el.dataset.type || "").toLowerCase();
  if (el.type === "checkbox" || t === "boolean") {
    const valStr = String(cellValue).toLowerCase();
    el.checked = !!cellValue && valStr !== "false" && cellValue !== 0;
  } else if (t === "number") {
    el.value = (cellValue ?? "") === "" ? "" : Number(cellValue);
  } else {
    el.value = cellValue ?? "";
  }
}

function getCellValueFromEl(el) {
  const t = (el.dataset.type || "").toLowerCase();
  if (el.type === "checkbox" || t === "boolean") return el.checked ? true : false;
  if (t === "number") {
    const n = Number(el.value);
    return Number.isFinite(n) ? n : null;
  }
  return el.value ?? "";
}

async function resolveRange(ctx, ws, bind) {
  const nm = ctx.workbook.names.getItemOrNullObject(bind);
  nm.load("name"); await ctx.sync();
  if (!nm.isNullObject) return nm.getRange();
  try { return ws.getRange(bind); } catch { return null; }
}
