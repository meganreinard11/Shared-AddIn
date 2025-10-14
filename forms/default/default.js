document.getElementById("save-note").addEventListener("click", async () => {
  const el = document.getElementById("note");
  el.dispatchEvent(new Event("change"));
});
