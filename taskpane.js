(function () {
  const standardPalette = [
    "#000000","#434343","#666666","#999999","#b7b7b7","#cccccc","#d9d9d9","#efefef","#f3f3f3","#ffffff",
    "#980000","#ff0000","#ff9900","#ffff00","#00ff00","#00ffff","#4a86e8","#0000ff","#9900ff","#ff00ff",
    "#e6b8af","#f4cccc","#fce5cd","#fff2cc","#d9ead3","#d0e0e3","#c9daf8","#cfe2f3","#d9d2e9","#ead1dc",
    "#dd7e6b","#ea9999","#f9cb9c","#ffe599","#b6d7a8","#a2c4c9","#a4c2f4","#9fc5e8","#b4a7d6","#d5a6bd"
  ];
	  
  const RECENT_START = 5;
  const RECENT_END = 23;
  const RECENT_RANGE = VALUE_COLUMN + RECENT_START + ":" + VALUE_COLUMN + RECENT_END;
  const RECENT_LIMIT = RECENT_END - RECENT_START;
  let recentCache = [];


  function normalizeHex(value) {
    if (!value) return null;
    let v = (""+value).trim();
    if (v.startsWith("rgb")) {
      const m = v.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/i);
      if (m) {
        const r = Number(m[1]).toString(16).padStart(2,"0");
        const g = Number(m[2]).toString(16).padStart(2,"0");
        const b = Number(m[3]).toString(16).padStart(2,"0");
        v = "#" + (r+g+b).toUpperCase();
      }
    }
    if (v[0] !== "#") v = "#"+v;
    if (/^#([0-9A-Fa-f]{3})$/.test(v)) v = "#"+v.slice(1).split("").map(ch => ch+ch).join("");
    if (/^#([0-9A-Fa-f]{6})$/.test(v)) return v.toUpperCase();
    if (/^#([0-9A-Fa-f]{8})$/.test(v)) return v.toUpperCase();
    return null;
  }

  function setColorInputs(hex) {
    const norm = normalizeHex(hex);
    if (!norm) return;
    document.getElementById("colorInput").value = norm.slice(0,7);
    document.getElementById("hexInput").value = norm;
  }

  async function ensureStoreSheet(context) {
    let sheet = context.workbook.worksheets.getItemOrNullObject(SETTINGS_SHEET);
    sheet.load("name,isNullObject,visibility");
    await context.sync();
    if (sheet.isNullObject) {
      sheet = context.workbook.worksheets.add(SETTINGS_SHEET);
      sheet.visibility = "Hidden";
      const r = sheet.getRange(RECENT_RANGE);
      r.numberFormat = "@";
    }
    return sheet;
  }

  async function getRecentColors() {
    return Excel.run(async (context) => {
      const sheet = await ensureStoreSheet(context);
      const rng = sheet.getRange(RECENT_RANGE);
      rng.load("values");
      await context.sync();
      const list = (rng.values || [])
        .map(row => (row[0] || "").toString().trim())
        .filter(v => !!normalizeHex(v));
      const seen = new Set();
      const uniq = [];
      list.forEach(v => {
        const u = normalizeHex(v);
        if (u && !seen.has(u)) { seen.add(u); uniq.push(u); }
      });
      recentCache = uniq;
      renderRecentsFromArray(uniq);
      return uniq;
    }).catch(e => {
      console.warn("getRecentColors failed:", e);
      recentCache = [];
      renderRecentsFromArray([]);
      return [];
    });
  }

  async function setRecentColors(newList) {
    return Excel.run(async (context) => {
      const sheet = await ensureStoreSheet(context);
      const seen = new Set();
      const uniq = [];
      newList.forEach(v => {
        const u = normalizeHex(v);
        if (u && !seen.has(u)) { seen.add(u); uniq.push(u); }
      });
      const trimmed = uniq.slice(0, RECENT_LIMIT);
      const rows = [];
      for (let i = 0; i < RECENT_LIMIT; i++) rows.push([trimmed[i] || ""]);
      const rng = sheet.getRange(RECENT_RANGE);
      rng.values = rows;
      rng.numberFormat = "@";
      await context.sync();
      recentCache = trimmed;
      renderRecentsFromArray(trimmed);
      return trimmed;
    }).catch(e => {
      console.error("setRecentColors failed:", e);
      return recentCache;
    });
  }

  async function pushRecentColor(color) {
    const hex = normalizeHex(color);
    if (!hex) return recentCache;
    const merged = [hex, ...recentCache.filter(c => c !== hex)];
    return setRecentColors(merged);
  }

  let themeColors = null;
  function readOfficeTheme() {
    try {
      const t = Office.context && Office.context.officeTheme;
      if (t) {
        const keys = ["background1","text1","background2","text2","accent1","accent2","accent3","accent4","accent5","accent6","hyperlink","followedHyperlink"];
        const list = [];
        keys.forEach(k => {
          const v = t[k];
          if (v && typeof v === "string") {
            const hex = normalizeHex(v);
            if (hex) list.push({ key:k, value:hex });
          }
        });
        if (list.length) return list;
      }
    } catch (e) { console.warn("Theme read failed:", e); }
    return null;
  }

  function renderTheme() {
    const el = document.getElementById("themeGrid");
    if (!el) return;
    el.innerHTML = "";
    const theme = themeColors || readOfficeTheme();
    if (theme) {
      themeColors = theme;
      theme.forEach(item => {
        const sw = document.createElement("button");
        sw.className = "swatch";
        sw.title = item.key + ": " + item.value;
        sw.style.background = item.value;
        sw.addEventListener("click", () => setColorInputs(item.value));
        el.appendChild(sw);
      });
    } else {
      const fallback = ["#000000","#FFFFFF","#1F497D","#4F81BD","#C0504D","#9BBB59","#8064A2","#4BACC6","#F79646"];
      fallback.forEach(c => {
        const sw = document.createElement("button");
        sw.className = "swatch";
        sw.title = c;
        sw.style.background = c;
        sw.addEventListener("click", () => setColorInputs(c));
        el.appendChild(sw);
      });
    }
  }

  if (Office && Office.onOfficeThemeChanged) {
    Office.onOfficeThemeChanged(() => { themeColors = null; renderTheme(); });
  }

  function renderRecentsFromArray(arr) {
    const el = document.getElementById("recentGrid");
    if (!el) return;
    el.innerHTML = "";
    (arr || []).forEach(c => {
      const sw = document.createElement("button");
      sw.className = "swatch";
      sw.title = c;
      sw.style.background = c;
      sw.addEventListener("click", () => setColorInputs(c));
      el.appendChild(sw);
    });
  }

  function renderStandard() {
    const el = document.getElementById("standardGrid");
    if (!el) return;
    el.innerHTML = "";
    standardPalette.forEach(c => {
      const sw = document.createElement("button");
      sw.className = "swatch";
      sw.title = c;
      sw.style.background = c;
      sw.addEventListener("click", () => setColorInputs(c));
      el.appendChild(sw);
    });
  }

  function getTargets() {
    return {
      fill: document.getElementById("targetFill").checked,
      font: document.getElementById("targetFont").checked,
      borders: document.getElementById("targetBorders").checked,
    };
  }

  async function applyToSelection(hex) {
    const color = normalizeHex(hex);
    if (!color) { 
		alert("Enter a valid hex color like #FFAA33");
		return;
	}
    const targets = getTargets();
    if (!targets.fill && !targets.font && !targets.borders) {
		alert("Choose at least one target (Fill, Font, Borders)."); 
		return; 
	}
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(["address"]);
      await context.sync();
      if (targets.fill)  range.format.fill.color = color;
      if (targets.font)  range.format.font.color = color;
      if (targets.borders) {
        const edges = ["EdgeTop","EdgeBottom","EdgeLeft","EdgeRight"];
        edges.forEach(e => {
          range.format.borders.getItem(e).color = color;
          range.format.borders.getItem(e).style = "Continuous";
          range.format.borders.getItem(e).weight = "Medium";
        });
      }
      await context.sync();
    }).catch(e => {
		alert("Excel API error: " + e));
    await pushRecentColor(color);
  }

  async function readFillFromSelection() {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("format/fill/color");
      await context.sync();
      const c = range.format.fill.color;
      if (c) setColorInputs(c);
      await pushRecentColor(c);
    }).catch(e => alert("Excel API error: " + e));
  }

  async function readFontFromSelection() {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("format/font/color");
      await context.sync();
      const c = range.format.font.color;
      if (c) setColorInputs(c);
      await pushRecentColor(c);
    }).catch(e => alert("Excel API error: " + e));
  }

  async function clearFillFont() {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.format.fill.clear();
      range.format.font.color = null;
      await context.sync();
    }).catch(e => alert("Excel API error: " + e));
  }

  async function eyedropperScreen() {
    const status = document.getElementById("eyedropperStatus");
    if (!("EyeDropper" in window)) { status.textContent = "Screen eyedropper not supported here. Use Eyedropper (Cell)."; return; }
    try {
      status.textContent = "Pick a pixel…";
      const dropper = new window.EyeDropper();
      const result = await dropper.open();
      if (result && result.sRGBHex) {
        setColorInputs(result.sRGBHex);
        await pushRecentColor(result.sRGBHex);
        status.textContent = "Picked " + result.sRGBHex.toUpperCase();
      } else { status.textContent = ""; }
    } catch (e) { status.textContent = ""; }
  }

  let cellSampleHandler = null;
  async function eyedropperCell() {
    const status = document.getElementById("eyedropperStatus");
    status.textContent = "Click a cell to sample its fill…";
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        if (cellSampleHandler) { try { sheet.onSelectionChanged.remove(cellSampleHandler); } catch {} cellSampleHandler = null; }
        cellSampleHandler = sheet.onSelectionChanged.add(async (_evt) => {
          try {
            await Excel.run(async (innerContext) => {
              const r = innerContext.workbook.getSelectedRange();
              r.load("format/fill/color");
              await innerContext.sync();
              const c = r.format.fill.color;
              if (c) {
                setColorInputs(c);
                await pushRecentColor(c);
                const statusEl = document.getElementById("eyedropperStatus");
                if (statusEl) statusEl.textContent = "Sampled " + c.toUpperCase();
              } else {
                const statusEl = document.getElementById("eyedropperStatus");
                if (statusEl) statusEl.textContent = "No fill on that cell.";
              }
            });
          } finally {
            try { sheet.onSelectionChanged.remove(cellSampleHandler); } catch {}
            cellSampleHandler = null;
          }
        });
      });
    } catch (e) {
      status.textContent = "Selection-change not available. Use 'Read Fill from Selection' instead.";
    }
  }

  function wireUi() {
    const colorInput = document.getElementById("colorInput");
    const hexInput = document.getElementById("hexInput");
    const applyBtn = document.getElementById("applyBtn");
    const getFillBtn = document.getElementById("getFillBtn");
    const getFontBtn = document.getElementById("getFontBtn");
    const clearBtn = document.getElementById("clearBtn");
    const eyedropperScreenBtn = document.getElementById("eyedropperScreenBtn");
    const eyedropperCellBtn = document.getElementById("eyedropperCellBtn");

    colorInput.addEventListener("input", () => { hexInput.value = colorInput.value.toUpperCase(); });
    hexInput.addEventListener("change", () => {
      const normalized = normalizeHex(hexInput.value);
      if (normalized) { colorInput.value = normalized.slice(0,7); hexInput.value = normalized; }
      else { alert("Invalid hex color. Use #RGB, #RRGGBB, or #RRGGBBAA."); hexInput.value = colorInput.value.toUpperCase(); }
    });
    applyBtn.addEventListener("click", () => applyToSelection(hexInput.value));
    getFillBtn.addEventListener("click", readFillFromSelection);
    getFontBtn.addEventListener("click", readFontFromSelection);
    clearBtn.addEventListener("click", clearFillFont);
    eyedropperScreenBtn.addEventListener("click", eyedropperScreen);
    eyedropperCellBtn.addEventListener("click", eyedropperCell);
  }

  Office.onReady(async () => {
    wireUi();
    renderTheme();
    renderStandard();
    await getRecentColors();
  });
})();
