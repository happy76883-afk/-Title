const NAMES = [
  "(표)(2023연장)투입산출표_생산자가격_기본부문.xlsx",
  "(표)(2023연장)투입산출표_생산자가격_통합대분류.xlsx",
  "(표)(2023연장)투입산출표_생산자가격_통합중분류.xlsx",
  "(표)(2023연장)투입산출표_생산자가격_통합소분류.xlsx"
];

window.addEventListener("DOMContentLoaded", async () => {
  const box = document.getElementById("auto-status");
  if (!box) return;
  const out = [];
  for (const name of NAMES) {
    try {
      const res = await fetch(`./data/${encodeURIComponent(name)}`);
      out.push(`${name.slice(-12)}: ${res.status} (${res.ok})`);
    } catch (e) {
      out.push(`${name.slice(-12)}: ERR (${e.message})`);
    }
    box.textContent = out.join(" / ");
  }
});
