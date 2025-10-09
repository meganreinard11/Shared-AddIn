/* Shared runtime entry (Option B: partial injection) with recursion guards */
let __BOOTSTRAPPED__ = false;
let __RENDER_BUSY__ = false;

Office.onReady(() => {
  if (__BOOTSTRAPPED__) return;
  __BOOTSTRAPPED__ = true;

  // Initialize error handler (reads _Settings!B4 if present)
  try { ErrorHandler.init(); } catch {}

  // Defer first render to next tick to avoid synchronous command re-entry
  setTimeout(() => { try { showGuest(); } catch (e) { /* swallow */ } }, 0);
});

/* Ribbon entry point (ExecuteFunction) */
window.showGuest = function (event) {
  // Complete the command immediately to avoid Office waiting & re-entering
  if (event && typeof event.completed === "function") {
    try { event.completed(); } catch {}
  }
  // Defer render to break any potential recursion
  setTimeout(() => { try { showGuest(); } catch (e) { /* swallowed by EH global */ } }, 0);
};

/* ----- Render Option B: inject fragment + hydrate ----- */
async function showGuest() {
  if (__RENDER_BUSY__) return; // prevent overlapping renders
  __RENDER_BUSY__ = true;
  try {
    const host = document.getElementById("app");
    if (!host) return;

    host.innerHTML = '<div style="padding:16px;color:#666">Loading…</div>';
    ensureGuestStyles();

    await ErrorHandler.tryWrap("Load guest form", async () => {
      const url = new URL("./guest-form.partial.html", window.location.href).toString();
      const html = await fetch(url, { cache: "no-store" }).then((r) => r.text());
      // Avoid infinite innerHTML loops by only replacing when different
      if (host.innerHTML !== html) host.innerHTML = html;
      hydrateGuestForm(host.querySelector("#guestForm"));
    });
  } finally {
    __RENDER_BUSY__ = false;
  }
}

/* ----- Attach behaviors and save to workbook ----- */
function hydrateGuestForm(formEl) {
  if (!formEl) return;

  // Remove any prior listener if rehydrating
  formEl.addEventListener("submit", onSubmit, { once: true });

  function onSubmit(e) {
    e.preventDefault();
    const data = Object.fromEntries(new FormData(formEl).entries());
    if (data.dob) {
      try { data.dob = new Date(`${data.dob}T00:00:00`); } catch {}
    }

    ErrorHandler.tryWrap("Save Guest", async () => {
      await Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.getItemOrNullObject("Guests");
        await ctx.sync();
        const sheet = ws.isNullObject ? ctx.workbook.worksheets.add("Guests") : ws;

        const tbl = sheet.tables.getItemOrNullObject("Guests");
        await ctx.sync();

        const headers = ["Full Name","Email","Phone","Birth Date","Gender","Address","Country","Postal"];
        const row = [[
          data.fullName || "", data.email || "", data.phone || "",
          data.dob || "", data.gender || "", data.address || "",
          data.country || "", data.postal || ""
        ]];

        if (tbl.isNullObject) {
          const t = sheet.tables.add("A1:H1", true);
          t.name = "Guests";
          t.getHeaderRowRange().values = [headers];
          t.rows.add(null, row);
        } else {
          tbl.rows.add(null, row);
        }
        sheet.activate();
      });

      formEl.reset();
      ErrorHandler.notify("Guest added to table.", { type: "success" });
    });
  }
}

/* ----- UI styles for the form (added once) ----- */
function ensureGuestStyles() {
  if (document.getElementById("guest-form-css")) return;
  const style = document.createElement("style");
  style.id = "guest-form-css";
  style.textContent = `
    :root{--bg:#7c68ff;--card:#fff;--text:#2c2c2c;--muted:#6b7280;--border:#e6e6ea;--focus:#7c68ff;--radius:12px}
    .wrap{display:grid;place-items:center;padding:24px;background:var(--bg);border-radius:14px;height:100%;overflow:auto}
    .card{width:100%;max-width:640px;background:#fff;border-radius:16px;box-shadow:0 8px 30px rgba(0,0,0,.12);padding:28px}
    .title{text-align:center;font-weight:700;font-size:28px;margin:4px 0 18px}
    .form-grid{display:grid;gap:16px}
    .two-col{display:grid;grid-template-columns:1fr 1fr;gap:14px}
    @media (max-width:560px){.two-col{grid-template-columns:1fr}}
    label{display:block;font-size:13px;color:var(--muted);margin-bottom:6px}
    input[type=text],input[type=email],input[type=tel],input[type=date],select,textarea{
      width:100%;height:44px;border:1px solid var(--border);border-radius:10px;padding:0 14px;background:#fff;outline:none}
    textarea{height:88px;padding-top:8px;resize:vertical}
    input:focus,select:focus,textarea:focus{border-color:var(--focus);box-shadow:0 0 0 3px rgba(124,104,255,.18)}
    fieldset{border:0;padding:0;margin:0}
    .radio-row{display:flex;gap:28px;align-items:center}
    .radio{display:flex;gap:8px;align-items:center}
    input[type=radio]{accent-color:var(--focus)}
    .btn{display:inline-flex;align-items:center;justify-content:center;width:100%;height:46px;border:0;border-radius:10px;background:var(--focus);color:#fff;font-weight:600;font-size:16px;cursor:pointer}
  `;
  document.head.appendChild(style);
}
