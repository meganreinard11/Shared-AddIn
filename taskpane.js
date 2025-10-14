// Dynamic Forms: selection wiring + address and NAMED RANGE routing
let lastRenderTs = 0;
const RENDER_COOLDOWN_MS = 120;
let lastRenderedFormId = null;

// Map form ids → html files
const HtmlMap = {
  default:  "./forms/default.html",
  settings: "./forms/settings.html",
  colorPalette: "./forms/colorPalette.html"
};

/**
 * Selection-based overrides per sheet.
 * Each rule supports either:
 *   - match.address: A1 address on the sheet (e.g., "B3" or "A2:C10")
 *   - match.name:    Workbook/worksheet named range (e.g., "CodeLink")
 *
 * Examples:
 *   { sheet: "settings", match: { address: "B3" }, form: "settingsCode" }
 *   { sheet: "settings", match: { name: "CodeLink" }, form: "settingsCode" }
 */
const SelectionRoutes = [
  { sheet: "settings", match: { address: "B3" }, form: "colorPalette" },
];

// Keep current selection subscription so we can add/remove dynamically
let selectionSub = null;
let selectionSubSheet = null;

Office.onReady(async () => {
  await renderForActiveWorksheet();
  await setupWorkbookEvents();
});

async function setupWorkbookEvents() {
  try {
    await Excel.run(async (ctx) => {
      const sheets = ctx.workbook.worksheets;

      // Re-render + adjust selection wiring on tab switch
      if (sheets.onActivated && sheets.onActivated.add) {
        await sheets.onActivated.add(async () => {
          await renderForActiveWorksheet();
          await manageSelectionSubscription();
        });
      }

      // Initial selection wiring for the current active sheet
      await manageSelectionSubscription();

      await ctx.sync();
    });
  } catch (e) {
    console.error("Workbook event setup failed:", e);
  }
}

// Add/remove selectionChanged depending on whether ACTIVE sheet has any SelectionRoutes
async function manageSelectionSubscription() {
  try {
    await Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getActiveWorksheet();
      ws.load("name");
      await ctx.sync();
      const sheetName = ws.name || "";

      const needsSelection = hasSelectionRoutesForSheet(sheetName);

      // Remove existing listener if not needed or sheet changed
      if (selectionSub && selectionSub.remove) {
        if (!needsSelection || (selectionSubSheet || "") !== sheetName) {
          await selectionSub.remove();
          selectionSub = null;
          selectionSubSheet = null;
        }
      }

      // Add when needed
      if (needsSelection && !selectionSub) {
        if (ws.onSelectionChanged && ws.onSelectionChanged.add) {
          selectionSub = await ws.onSelectionChanged.add(onSelectionChanged);
          selectionSubSheet = sheetName;
        }
      }
      await ctx.sync();
    });
  } catch (e) {
    console.warn("manageSelectionSubscription error:", e);
  }
}

// Handle selection changes only when needed; reload only if target form differs
async function onSelectionChanged() {
  const now = Date.now();
  if (now - lastRenderTs < RENDER_COOLDOWN_MS) return;

  try {
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
      const overrideFormId = await pickSelectionOverride(ctx, sheetName, selectionAddress);
      const nextFormId = overrideFormId || baseFormId;

      // Only reload if the form is changing
      if (nextFormId !== lastRenderedFormId) {
        lastRenderTs = Date.now();
        updateSheetBadge(sheetName);
        await renderForm(nextFormId, { sheetName, hint, selectionAddress, override: !!overrideFormId });
      }
    });
  } catch (e) {
    console.error("onSelectionChanged error:", e);
  }
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
    const overrideFormId = await pickSelectionOverride(ctx, sheetName, selectionAddress);
    const finalFormId = overrideFormId || baseFormId;

    updateSheetBadge(sheetName);

    if (finalFormId !== lastRenderedFormId) {
      await renderForm(finalFormId, { sheetName, hint, selectionAddress, override: !!overrideFormId });
    }
  });
}

// ---------- Routing helpers ----------

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

function hasSelectionRoutesForSheet(sheetName) {
  const s = (sheetName || "").toLowerCase().trim();
  return SelectionRoutes.some(r => (r.sheet || "").toLowerCase().trim() === s);
}

// Async: can resolve named ranges within rules
async function pickSelectionOverride(ctx, sheetName, selectionAddress) {
  const s = (sheetName || "").toLowerCase().trim();
  const sel = normalizeA1(selectionAddress);
  if (!hasSelectionRoutesForSheet(s)) return null;

  for (const rule of SelectionRoutes) {
    if ((rule.sheet || "").toLowerCase().trim() !== s) continue;
    const m = rule.match || {};

    // Address match (fast path)
    if (m.address && containsAddress(sel, normalizeA1(m.address))) {
      return rule.form;
    }

    // Named range match
    if (m.name) {
      // Try workbook-scoped name first
      const nm = ctx.workbook.names.getItemOrNullObject(m.name);
      nm.load(["name", "type"]);
      await ctx.sync();

      if (!nm.isNullObject) {
        try {
          const r = nm.getRange();
          r.load("address");
          await ctx.sync();
          // address like "Settings!$B$3" → ensure it's on the same sheet
          const parts = String(r.address || "").split("!");
          const sheetFromName = parts.length > 1 ? parts[0].replace(/^'/, "").replace(/'$/, "") : s;
          if (sheetFromName.toLowerCase().trim() !== s) {
            continue; // name exists but scoped to a different sheet
          }
          const localA1 = localizeAddress(r.address);
          if (containsAddress(sel, normalizeA1(localA1))) {
            return rule.form;
          }
        } catch (e) {
          // name exists but not a range; ignore
        }
      }

      // Optionally: look for a worksheet-scoped name with same identifier
      // (Excel JS exposes most names via workbook.names; skipping extra search for brevity)
    }
  }
  return null;
}

// ---------- Address utilities ----------

function localizeAddress(fullAddress) {
  const parts = String(fullAddress || "").split("!");
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

// ---------- UI & rendering ----------

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
  lastRenderedFormId = formId; // track current
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
