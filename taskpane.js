// Dynamic Forms: sheet → HTML + data bindings
// - Routes to an HTML form based on sheet name or a "form:..." tag in A1
// - Binds inputs with data-bind="NameOrA1" to worksheet values (two-way)

let lastRenderTs = 0;
const RENDER_COOLDOWN_MS = 300;

// Map form ids → html files
const HtmlMap = {
  default:  "./forms/default.html",
  settings: "./forms/settings.html",
  diagnostics: "./forms/diagnostics.html",
};

Office.onReady(async () => {
  // Initial render
  await renderForActiveWorksheet();

  // Subscribe to events for sheet/tab changes && selection (fallback)
  try {
    await Excel.run(async (ctx) => {
      const sheets = ctx.workbook.worksheets;

      // Preferred: worksheets.onActivated (fires on tab switch)
      if (sheets.onActivated && sheets.onActivated.add) {
        await sheets.onActivated.add(onWorksheetActivated);
        console.log("Subscribed: worksheets.onActivated");
      } else {
        console.log("onActivated not available; relying on selectionChanged");
      }

      // Fallback: per-sheet selection changed
      sheets.load("items/name");
      await ctx.sync();
      for (const ws of sheets.items) {
        if (ws.onSelectionChanged && ws.onSelectionChanged.add) {
          await ws.onSelectionChanged.add(onSelectionChanged);
        }
      }

      // For newly added sheets, attach selectionChanged too
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

async function onWorksheetActivated(event) {
  await renderForActiveWorksheet();
}

// Fallback: throttle re-render on selection change
async function onSelectionChanged(event) {
  const now = Date.now();
  if (now - lastRenderTs > RENDER_COOLDOWN_MS) {
    await renderForActiveWorksheet();
  }
}

async function renderForActiveWorksheet() {
  lastRenderTs = Date.now();
  await Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getActiveWorksheet();
    ws.load("name");
    const a1 = ws.getRange("A1");
    a1.load(["text"]);
    await ctx.sync();

    const sheetName = (ws.name || "").trim();
    const hint = parseFormHintFromCell(a1.text);
    const formId = pickFormId(hint, sheetName);

    updateSheetBadge(sheetName);
    await renderForm(formId, { sheetName, hint });
  });
}

// Detect "form:xyz" in A1
function parseFormHintFromCell(text2d) {
  try {
    const t = (text2d && text2d[0] && text2d[0][0]) ? String(text2d[0][0]) : "";
    const m = /form\s*:\s*([a-z0-9_-]+)/i.exec(t);
    return m ? m[1].toLowerCase() : null;
  } catch {
    return null;
  }
}

function pickFormId(hint, sheetName) {
  if (hint && HtmlMap[hint]) return hint;
  const key = (sheetName || "").toLowerCase().trim();
  if (HtmlMap[key]) return key;
  return "default";
}

function updateSheetBadge(name) {
  const el = document.getElementById("sheetName");
  if (el) el.textContent = name || "(unknown)";
}

// Load the HTML file && then wire data bindings
async function renderForm(formId, ctx) {
  const app = document.getElementById("app");
  if (!app) return;
  const url = HtmlMap[formId] || HtmlMap.default;

  app.innerHTML = `<div class="loading">Loading ${formId}…</div>`;
  const html = await (await fetch(url, { cache: "no-store" })).text();
  app.innerHTML = html;

  await wireBindings(app);
}

// ----------------- Two-way binding layer -----------------

let sheetChangeUnsub = null;
let debounceTimer = null;

async function wireBindings(container) {
  // 1) Initial pull from sheet → controls
  await refreshBoundControls(container);

  // 2) UI → sheet on change/blur
  container.querySelectorAll("[data-bind]").forEach((el) => {
    const handler = async () => { try { await writeBinding(el); } catch (e) { console.error(e); } };
    el.addEventListener("change", handler);
    if (el.type === "text" || el.tagName === "TEXTAREA") el.addEventListener("blur", handler);
  });

  // 3) Observe Excel sheet changes → UI (ExcelApi >= 1.9)
  try {
    await Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getActiveWorksheet();
      if (sheetChangeUnsub && sheetChangeUnsub.remove) {
        sheetChangeUnsub.remove();
        sheetChangeUnsub = null;
      }
      if (ws.onChanged && ws.onChanged.add) {
        sheetChangeUnsub = await ws.onChanged.add(async () => {
          clearTimeout(debounceTimer);
          debounceTimer = setTimeout(() => refreshBoundControls(container), 120);
        });
      }
    });
  } catch (e) {
    console.warn("Worksheet.onChanged unavailable; form will update on tab switch.", e);
  }
}

async function refreshBoundControls(container) {
  const els = [...container.querySelectorAll("[data-bind]")];
  if (els.length === 0) return;

  await Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getActiveWorksheet();
    for (const el of els) {
      const bind = (el.dataset.bind || "").trim();
      if (!bind) continue;
      const rng = await resolveRange(ctx, ws, bind);
      if (!rng) continue;
      rng.load("values");
    }
    await ctx.sync();

    // Second pass to assign values
    for (const el of els) {
      const bind = (el.dataset.bind || "").trim();
      if (!bind) continue;
      const rng = await resolveRange(ctx, ws, bind);
      if (!rng) continue;
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
    const toWrite = getCellValueFromEl(el);
    rng.values = [[toWrite]];
    await ctx.sync();
  });
}

// Helpers

function setElValueFromCell(el, cellValue) {
  const t = (el.dataset.type || "").toLowerCase();
  if (el.type === "checkbox" || t === "boolean") {
    el.checked = !!cellValue && String(cellValue).toLowerCase() !== "false" && cellValue !== 0 ? true : false;
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

// Try workbook-level named item first; else A1 address on active sheet
async function resolveRange(ctx, ws, bind) {
  const nm = ctx.workbook.names.getItemOrNullObject(bind);
  nm.load("name");
  await ctx.sync();
  if (!nm.isNullObject) return nm.getRange();
  try { return ws.getRange(bind); } catch { return null; }
}
