/* Simple registry of forms keyed by a "form id".
   You can map by sheet name or by a tag in A1 like: "form:orders"
*/

const Forms = {
  default: (ctx) => `
    <div class="form-card">
      <h2>Welcome</h2>
      <p>This is the default form. Create a sheet named <span class="badge">Orders</span>, <span class="badge">Inventory</span>, or <span class="badge">Settings</span>, or put <code>form:&lt;id&gt;</code> in A1 to force a form.</p>
      <div class="row"><label>Note</label><input id="note" placeholder="Type something…"/></div>
      <div class="actions"><button class="btn primary" onclick="saveNote()">Save</button></div>
    </div>
  `,

  colorPalette: (ctx) => `
    <div class="form-card">
    <section class="controls">
      <div class="row">
        <label for="scheme">Scheme</label>
        <select id="scheme">
          <option value="auto">Auto</option>
          <option value="monochrome">Monochrome</option>
          <option value="analogous">Analogous</option>
          <option value="complementary">Complementary</option>
          <option value="split-complementary">Split‑Complementary</option>
          <option value="triadic">Triadic</option>
          <option value="tetradic">Tetradic</option>
          <option value="neutral">Neutral (Grays)</option>
          <option value="pastel">Pastel</option>
          <option value="vibrant">Vibrant</option>
          <option value="warm">Warm</option>
          <option value="cool">Cool</option>
        </select>
      </div>

      <div class="row">
        <label for="baseColor">Base color (optional)</label>
        <input id="baseColor" type="text" placeholder="#4a86e8 or hsl(220,60%,60%)" />
        <button id="randomBase" type="button">Random Base</button>
      </div>

      <div class="row">
        <label for="count">Count</label>
        <input id="count" type="number" min="3" max="32" value="10" />
      </div>

      <div class="row">
        <button id="generate" class="primary" type="button">Generate</button>
        <button id="insert" type="button" title="Insert into sheet as a table">Insert to Sheet</button>
        <button id="copyHex" type="button" title="Copy hex codes to clipboard">Copy HEX</button>
      </div>
    </section>

    <section class="palette" id="palette"></section>

    <section class="recents">
      <div class="recents-header">
        <h2>Recent Palettes</h2>
        <div class="row">
          <button id="reloadRecents" type="button">Reload Recents</button>
        </div>
      </div>
      <div id="recentList" class="recent-list"></div>
    </section>
    </div>
  `,

  orders: (ctx) => `
    <div class="form-card">
      <h2>Orders</h2>
      <div class="row"><label>Order #</label><input id="order-id" placeholder="e.g. SO-10023"/></div>
      <div class="row"><label>Customer</label><input id="customer" placeholder="Acme Co."/></div>
      <div class="row"><label>Status</label>
        <select id="status">
          <option>Pending</option><option>Processing</option><option>Shipped</option><option>Closed</option>
        </select>
      </div>
      <div class="actions">
        <button class="btn" onclick="writeOrderRow()">Insert Row</button>
        <button class="btn primary" onclick="syncOrderStatus()">Sync Status</button>
      </div>
    </div>
  `,

  inventory: (ctx) => `
    <div class="form-card">
      <h2>Inventory</h2>
      <div class="row"><label>SKU</label><input id="sku" placeholder="SKU-001"/></div>
      <div class="row"><label>On Hand</label><input id="onhand" type="number" min="0"/></div>
      <div class="row"><label>Reorder Point</label><input id="rop" type="number" min="0"/></div>
      <div class="actions"><button class="btn primary" onclick="upsertSku()">Upsert SKU</button></div>
    </div>
  `,

  settings: (ctx) => `
    <div class="form-card">
      <h2>Settings</h2>
      <div class="row"><label>Company</label><input id="company" placeholder="Contoso, Ltd."/></div>
      <div class="row"><label>Theme</label>
        <select id="theme"><option>Light</option><option>Dark</option></select>
      </div>
      <div class="actions"><button class="btn primary" onclick="saveSettings()">Save Settings</button></div>
    </div>
  `,
};

// Optional: map common sheet names → form ids
const SheetToForm = {
  orders: "orders",
  inventory: "inventory",
  settings: "settings",
};

// Example handlers used by the forms above.
// Keep them global so inline onclick works. In production, wire listeners programmatically.

async function saveNote() {
  await Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getActiveWorksheet();
    const tgt = ws.getRange("B1");
    tgt.values = [[document.getElementById("note").value || ""]];
    await ctx.sync();
  });
}

async function writeOrderRow() {
  await Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getActiveWorksheet();
    const tbl = ws.tables.getItemOrNullObject("OrdersTable");
    await ctx.sync();
    const orderId = document.getElementById("order-id").value || "";
    const cust = document.getElementById("customer").value || "";
    const status = document.getElementById("status").value || "";
    if (tbl.isNullObject) {
      // Append to A1:C1 if no table
      const range = ws.getRange("A1:C1");
      range.getOffsetRange(1, 0).insert(Excel.InsertShiftDirection.down);
      range.values = [["OrderId", "Customer", "Status"]];
      range.format.font.bold = true;
      ws.getRange("A2:C2").values = [[orderId, cust, status]];
    } else {
      tbl.rows.add(null, [[orderId, cust, status]]);
    }
    await ctx.sync();
  });
}

async function syncOrderStatus() {
  await Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getActiveWorksheet();
    const range = ws.getRange("C2:C1048576"); // Status column area
    range.load("values");
    await ctx.sync();
    // noop demo: you could iterate and push to an API, etc.
    console.log("Statuses found:", range.values.slice(0, 10));
  });
}

async function upsertSku() {
  await Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getActiveWorksheet();
    const sku = (document.getElementById("sku").value || "").trim();
    const onhand = Number(document.getElementById("onhand").value || 0);
    const rop = Number(document.getElementById("rop").value || 0);
    if (!sku) return;

    // Simple demo upsert into a 3-col table starting at A1.
    const header = ws.getRange("A1:C1");
    header.values = [["SKU", "OnHand", "ROP"]];
    header.format.font.bold = true;

    // Find first empty row in column A (naive)
    const colA = ws.getRange("A:A");
    colA.load("values");
    await ctx.sync();
    let r = 2;
    while (colA.values[r - 1] && colA.values[r - 1][0]) r++;
    ws.getRange(`A${r}:C${r}`).values = [[sku, onhand, rop]];
    await ctx.sync();
  });
}

async function saveSettings() {
  await Excel.run(async (ctx) => {
    const theme = document.getElementById("theme").value;
    const company = document.getElementById("company").value || "";
    // Store to a hidden sheet or named items in a real app
    const ws = ctx.workbook.worksheets.getActiveWorksheet();
    ws.getRange("E1:F1").values = [["Company", "Theme"]];
    ws.getRange("E2:F2").values = [[company, theme]];
    await ctx.sync();
  });
}
