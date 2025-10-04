// Expose helpers globally
(function () {
  function normalizeHex(v) {
    v = (v || "").trim();
    if (!v) return null;
    if (v[0] !== "#") v = "#" + v;
    if (/^#([0-9A-Fa-f]{3})$/.test(v)) {
      v = "#" + v.slice(1).split("").map(ch => ch + ch).join("");
    }
    if (!/^#([0-9A-Fa-f]{6}|[0-9A-Fa-f]{8})$/.test(v)) return null;
    return v.toUpperCase();
  }
  window.ColorUtils = { normalizeHex };
})();
