// Throttle to avoid overly-frequent re-renders on selection changes
let lastRenderTs = 0;
const RENDER_COOLDOWN_MS = 300;

Office.onReady(async () => {
  if (!Office.context.requirements.isSetSupported("ExcelApi", "1.13")) {
    console.warn("ExcelApi 1.13 not fully available. Falling back to selectionChanged heuristic.");
  }

  // Initial render
  await renderForActiveWorksheet();

  // Subscribe to workbook-level "worksheet activated" when available
  try {
    await Excel.run(async (ctx) => {
      const sheets = ctx.workbook.worksheets;

      // Workbook-level activation (preferred when available)
      if (sheets.onActivated && sheets.onActivated.add) {
        await sheets.onActivated.add(onWorksheetActivated);
        console.log("Subscribed to worksheets.onActivated.");
      } else {
        console.log("worksheets.onActivated not available; using per-sheet selectionChanged.");
      }

      // Also attach selectionChanged on each sheet as a robust fallback
      sheets.load("items/name");
      await ctx.sync();

      for (const ws of sheets.items) {
        if (ws.onSelectionChanged && ws.onSelectionChanged.add) {
          await ws.onSelectionChanged.add(onSelectionChanged);
        }
      }

      // If new sheets get added later, attach selectionChanged to them too
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

// Workbook-level: fires when user changes tabs (ideal)
async function onWorksheetActivated(event) {
  await renderForActiveWorksheet();
}

// Fallback: selection changes inside a sheet; we only re-render if cooldown passed
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
    renderForm(formId, { sheetName, hint });
  });
}

// If A1 contains something like "form:orders" (case-insensitive), use that.
// For Excel JS, range.text is a 2D array of strings.
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
  if (hint && Forms[hint]) return hint;
  const key = (sheetName || "").toLowerCase().trim();
  const mapped = SheetToForm[key];
  if (mapped && Forms[mapped]) return mapped;
  return "default";
}

function updateSheetBadge(name) {
  const el = document.getElementById("sheetName");
  if (el) el.textContent = name || "(unknown)";
}

function renderForm(formId, ctx) {
  const app = document.getElementById("app");
  if (!app) return;
  const renderFn = Forms[formId] || Forms.default;
  app.innerHTML = renderFn(ctx);
}
