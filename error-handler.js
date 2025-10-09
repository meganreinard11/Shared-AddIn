/* Central error/notify util for shared runtime */
(function () {
  const EH = {
    initialized: false,
    env: { dev: true, source: "_Settings!B4" }, // TRUE => dev; FALSE => prod

    async init() {
      if (this.initialized) return;
      try {
        await Excel.run(async (ctx) => {
          const ws = ctx.workbook.worksheets.getItemOrNullObject("_Settings");
          await ctx.sync();
          if (ws.isNullObject) throw new Error("_Settings not found");
          const rng = ws.getRange("B4");
          rng.load("values");
          await ctx.sync();
          const v = String(rng.values?.[0]?.[0] ?? "").trim().toLowerCase();
          this.env.dev = v === "true" || v === "1" || v === "yes";
        });
      } catch {
        // If settings sheet not present, keep default dev = true
        this.env.dev = true;
      }
      this.initialized = true;
    },

    notify(message, opts = {}) {
      const type = opts.type || "info"; // success | info | error
      const bg = type === "error" ? "#ef4444" : type === "success" ? "#22c55e" : "#3b82f6";
      const el = document.createElement("div");
      el.textContent = message;
      el.style.cssText =
        "position:fixed;right:16px;bottom:16px;z-index:9999;background:" +
        bg +
        ";color:#fff;padding:10px 12px;border-radius:10px;box-shadow:0 6px 18px rgba(0,0,0,.18);font-weight:600";
      document.body.appendChild(el);
      setTimeout(() => el.remove(), opts.timeout ?? 2400);
    },

    async handleError(err, context = "") {
      const msg = err?.message || String(err);
      if (this.env.dev) console.error("❌", context, err);

      this.notify((context ? context + ": " : "") + msg, { type: "error", timeout: 4000 });

      // Best-effort async log to a hidden _Logs sheet
      if (typeof Excel === "undefined" || !Office?.context?.host) return;
      setTimeout(() => {
        Excel.run(async (ctx) => {
          let ws = ctx.workbook.worksheets.getItemOrNullObject("_Logs");
          await ctx.sync();
          if (ws.isNullObject) ws = ctx.workbook.worksheets.add("_Logs");

          const header = ws.getRange("A1:D1");
          header.load("values");
          await ctx.sync();

          if (!header.values?.[0]?.[0]) {
            ws.getRange("A1:D1").values = [["Timestamp", "Context", "Message", "Stack"]];
          }
          const row = [[
            new Date().toISOString(),
            context || "",
            msg,
            (err && err.stack ? String(err.stack).slice(0, 8000) : "")
          ]];
          const used = ws.getUsedRangeOrNullObject();
          used.load("rowCount");
          await ctx.sync();
          const nextRow = (used.isNullObject ? 1 : used.rowCount) + 1;
          ws.getRange(`A${nextRow}:D${nextRow}`).values = row;

          try { ws.visibility = Excel.SheetVisibility.hidden; } catch {}
          await ctx.sync();
        }).catch(() => {});
      }, 0);
    },

    async tryWrap(label, fn) {
      try {
        await this.init();
        return await fn();
      } catch (e) {
        await this.handleError(e, label);
        return undefined;
      }
    },
  };

  window.addEventListener("unhandledrejection", (e) => EH.handleError(e.reason, "unhandledrejection"));
  window.addEventListener("error", (e) => EH.handleError(e.error || e.message, "window.onerror"));
  window.ErrorHandler = EH;
})();
