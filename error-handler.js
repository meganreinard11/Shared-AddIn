/* Robust central error/notify util for shared runtime
   - Guards against recursive error handling (stack overflows)
   - Dedupe repeated errors
   - Rate-limits notifications and workbook logging
   - Never rethrow/reject from handlers
*/
(function () {
  if (window.ErrorHandler && window.ErrorHandler.__BOUND_GLOBALS__) {
    // Already initialized; avoid double-binding listeners
    return;
  }

  const EH = {
    initialized: false,
    inHandle: false,           // recursion guard
    __BOUND_GLOBALS__: true,   // mark once
    env: { dev: true, source: "_Settings!B4" },
    _dedupe: { lastSig: "", lastAt: 0 },

    // --- rate limiter buckets (token bucket) ---
    _rate: {
      notify: { tokens: 6, max: 6, refillMs: 10000, last: Date.now(), suppressed: false }, // ~6 toasts / 10s
      log:    { tokens: 20, max: 20, refillMs: 60000, last: Date.now(), suppressed: false } // ~20 logs / min
    },
    _refill(bucket){
      const now = Date.now();
      const elapsed = now - bucket.last;
      const add = Math.floor(elapsed / bucket.refillMs);
      if (add > 0) {
        bucket.tokens = Math.min(bucket.max, bucket.tokens + add);
        bucket.last = now;
        // clear suppression when tokens refill
        if (bucket.tokens > 0) bucket.suppressed = false;
      }
    },
    _consume(bucket){
      this._refill(bucket);
      if (bucket.tokens > 0) { bucket.tokens--; return true; }
      return false;
    },

    safeToString(x) {
      try {
        if (!x) return String(x);
        if (x instanceof Error) return x.message || String(x);
        if (typeof x === "string") return x;
        if (typeof x === "object") {
          const seen = new WeakSet();
          return JSON.stringify(x, (k, v) => {
            if (typeof v === "object" && v !== null) {
              if (seen.has(v)) return "[Circular]";
              seen.add(v);
            }
            return v;
          });
        }
        return String(x);
      } catch {
        try { return String(x); } catch { return "[unstringifiable]"; }
      }
    },

    signature(err, context) {
      const m = (err && err.message) ? err.message : this.safeToString(err);
      const s = (err && err.stack) ? String(err.stack).slice(0, 120) : "";
      return `${context}|${m}|${s}`;
    },

    shouldDedupe(sig) {
      const now = Date.now();
      const isSame = sig === this._dedupe.lastSig;
      const within = now - this._dedupe.lastAt < 1200; // 1.2s window
      if (isSame && within) return true;
      this._dedupe.lastSig = sig;
      this._dedupe.lastAt = now;
      return false;
    },

    async init() {
      if (this.initialized) return;
      try {
        if (typeof Excel !== "undefined" && Office?.context?.host) {
          await Excel.run(async (ctx) => {
            const ws = ctx.workbook.worksheets.getItemOrNullObject("_Settings");
            await ctx.sync();
            if (!ws.isNullObject) {
              const rng = ws.getRange("B4");
              rng.load("values");
              await ctx.sync();
              const v = String(rng.values?.[0]?.[0] ?? "").trim().toLowerCase();
              this.env.dev = v === "true" || v === "1" || v === "yes";
            }
          });
        }
      } catch { /* default dev:true if missing */ }
      this.initialized = true;
    },

    notify(message, opts = {}) {
      try {
        // rate limit user-facing toasts
        if (!this._consume(this._rate.notify)) {
          if (!this._rate.notify.suppressed) {
            this._rate.notify.suppressed = true;
            // Show one compact suppression note in dev only
            if (this.env.dev) {
              const el = document.createElement("div");
              el.textContent = "Too many errors — notifications temporarily suppressed.";
              el.style.cssText =
                "position:fixed;right:16px;bottom:16px;z-index:9999;background:#6b7280;color:#fff;padding:8px 10px;border-radius:10px;box-shadow:0 6px 18px rgba(0,0,0,.18);font-weight:600;font-size:12px";
              document.body.appendChild(el);
              setTimeout(() => { try { el.remove(); } catch {} }, 1800);
            }
          }
          return;
        }

        const type = opts.type || "info"; // success | info | error
        const bg = type === "error" ? "#ef4444" : type === "success" ? "#22c55e" : "#3b82f6";
        const el = document.createElement("div");
        el.textContent = message;
        el.style.cssText =
          "position:fixed;right:16px;bottom:16px;z-index:9999;background:" +
          bg +
          ";color:#fff;padding:10px 12px;border-radius:10px;box-shadow:0 6px 18px rgba(0,0,0,.18);font-weight:600";
        document.body.appendChild(el);
        setTimeout(() => { try { el.remove(); } catch {} }, opts.timeout ?? 2400);
      } catch { /* never throw from notify */ }
    },

    async handleError(err, context = "") {
      // Re-entrancy guard to prevent stack overflow from recursive unhandledrejection
      if (this.inHandle) {
        try { console.error("Nested error ignored:", err); } catch {}
        return;
      }
      this.inHandle = true;
      try {
        await this.init();
        const msg = (err && err.message) ? err.message : this.safeToString(err);
        const stack = err && err.stack ? String(err.stack) : "";
        const sig = this.signature(err, context);
        if (this.shouldDedupe(sig)) {
          return; // Skip noisy duplicates
        }

        if (this.env.dev) {
          try { console.error("❌", context, err); } catch {}
        }

        this.notify((context ? context + ": " : "") + msg, { type: "error", timeout: 4000 });

        // Async log to workbook; rate-limited & failures swallowed
        if (typeof Excel !== "undefined" && Office?.context?.host) {
          if (!this._consume(this._rate.log)) {
            if (!this._rate.log.suppressed && this.env.dev) {
              this._rate.log.suppressed = true;
              this.notify("Too many errors — logging suppressed temporarily.", { type: "error" });
            }
          } else {
            setTimeout(() => {
              try {
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
                    stack ? stack.slice(0, 8000) : ""
                  ]];
                  const used = ws.getUsedRangeOrNullObject();
                  used.load("rowCount");
                  await ctx.sync();
                  const nextRow = (used.isNullObject ? 1 : used.rowCount) + 1;
                  ws.getRange(`A${nextRow}:D${nextRow}`).values = row;

                  try { ws.visibility = Excel.SheetVisibility.hidden; } catch {}
                  await ctx.sync();
                }).catch(() => {});
              } catch { /* ignore logging failures */ }
            }, 0);
          }
        }
      } catch (inner) {
        try { console.error("Error in handleError:", inner); } catch {}
      } finally {
        this.inHandle = false;
      }
    },

    bindGlobalHandlers() {
      if (window.__EH_BOUND__) return;
      window.__EH_BOUND__ = true;

      window.addEventListener("unhandledrejection", (e) => {
        try { e && typeof e.preventDefault === "function" && e.preventDefault(); } catch {}
        try { this.handleError(e?.reason, "unhandledrejection"); } catch {}
      });

      window.addEventListener("error", (e) => {
        try { e && e.preventDefault && e.preventDefault(); } catch {}
        const payload = e?.error || e?.message || "Unknown error";
        try { this.handleError(payload, "window.onerror"); } catch {}
      });
    },

    async tryWrap(label, fn) {
      try {
        await this.init();
        const out = await fn();
        return out;
      } catch (e) {
        try { await this.handleError(e, label); } catch {}
        return undefined;
      }
    },
  };

  EH.bindGlobalHandlers();
  window.ErrorHandler = EH;
})();
