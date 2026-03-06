

function escapeHtml(value){
  const s = String(value ?? "");
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}
// Magallanes Funeral — Local Prototype v12
// Adds: Filter system + Store-based rendering (Month headers + TOTAL + GRAND TOTAL), non-editable grid, dblclick edit, CSV/XLSX import
(function () {
  const $ = (sel, root = document) => root.querySelector(sel);
  const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));

  // ---------------------------
  // Themes
  // ---------------------------
  const THEMES = {
    midnight: { vars: {"--bg":"#0b0d10","--panel":"#12161c","--panel2":"#0f1318","--text":"#eef2f7","--muted":"#9aa7b5","--accent":"#4ea1ff","--danger":"#ff4d6d","--line":"rgba(255,255,255,.10)","--line2":"rgba(255,255,255,.14)"}, font:'ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial' },
    slate:    { vars: {"--bg":"#0e1116","--panel":"#141a22","--panel2":"#10161e","--text":"#eef2f7","--muted":"#9aa7b5","--accent":"#60a5fa","--danger":"#fb7185","--line":"rgba(255,255,255,.10)","--line2":"rgba(255,255,255,.16)"}, font:'ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial' },
    royal:    { vars: {"--bg":"#070b16","--panel":"#0f1730","--panel2":"#0b1226","--text":"#eef2ff","--muted":"#a7b0c9","--accent":"#2f7cff","--danger":"#ff5c7a","--line":"rgba(255,255,255,.10)","--line2":"rgba(255,255,255,.18)"}, font:'ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial' },
    emerald:  { vars: {"--bg":"#071113","--panel":"#0e1b1e","--panel2":"#0a1619","--text":"#eafff7","--muted":"#9bc6b8","--accent":"#22c55e","--danger":"#fb7185","--line":"rgba(255,255,255,.10)","--line2":"rgba(255,255,255,.16)"}, font:'ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial' },
    forest:   { vars: {"--bg":"#06110a","--panel":"#0d1d12","--panel2":"#0a1710","--text":"#eaf7ee","--muted":"#a1c1aa","--accent":"#34d399","--danger":"#f97316","--line":"rgba(255,255,255,.10)","--line2":"rgba(255,255,255,.16)"}, font:'ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial' },
    sand:     { vars: {"--bg":"#f6f1e8","--panel":"#ffffff","--panel2":"#fbf7f0","--text":"#1f2937","--muted":"#5b6777","--accent":"#b45309","--danger":"#dc2626","--line":"rgba(2,6,23,.10)","--line2":"rgba(2,6,23,.16)"}, font:'ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial' },
    graphite: { vars: {"--bg":"#0c0c0e","--panel":"#141417","--panel2":"#101013","--text":"#f1f5f9","--muted":"#a1a1aa","--accent":"#e5e7eb","--danger":"#fb7185","--line":"rgba(255,255,255,.10)","--line2":"rgba(255,255,255,.16)"}, font:'ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial' },
    lavender: { vars: {"--bg":"#0b0a12","--panel":"#151228","--panel2":"#100f20","--text":"#f6f3ff","--muted":"#b7acd8","--accent":"#a78bfa","--danger":"#fb7185","--line":"rgba(255,255,255,.10)","--line2":"rgba(255,255,255,.16)"}, font:'ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial' },
    rose:     { vars: {"--bg":"#0d0b10","--panel":"#17131c","--panel2":"#120f16","--text":"#fff1f2","--muted":"#c4b3b8","--accent":"#fb7185","--danger":"#f97316","--line":"rgba(255,255,255,.10)","--line2":"rgba(255,255,255,.16)"}, font:'ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial' },
    classic:  { vars: {"--bg":"#f4f6f9","--panel":"#ffffff","--panel2":"#f7f9fc","--text":"#111827","--muted":"#5b6777","--accent":"#2563eb","--danger":"#dc2626","--line":"rgba(2,6,23,.10)","--line2":"rgba(2,6,23,.16)"}, font:'Segoe UI, Tahoma, Arial, sans-serif' },

    sunset:   { vars: {"--bg":"#0b0a0f","--panel":"#15111e","--panel2":"#110d18","--text":"#fff7ed","--muted":"#f4c7b2","--accent":"#fb923c","--danger":"#f43f5e","--line":"rgba(255,255,255,.10)","--line2":"rgba(255,255,255,.18)"}, font:'ui-sans-serif, system-ui, -apple-system, Segoe UI, Inter, Roboto, Helvetica, Arial' },
    tropical: { vars: {"--bg":"#050c10","--panel":"#0c1a22","--panel2":"#07141b","--text":"#ecfeff","--muted":"#9de3ff","--accent":"#22d3ee","--danger":"#f59e0b","--line":"rgba(255,255,255,.10)","--line2":"rgba(255,255,255,.18)"}, font:'ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial' },
    aqua:     { vars: {"--bg":"#041014","--panel":"#082029","--panel2":"#061820","--text":"#e6fffb","--muted":"#9fe9e3","--accent":"#2dd4bf","--danger":"#fb7185","--line":"rgba(255,255,255,.10)","--line2":"rgba(255,255,255,.18)"}, font:'ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial' },
    citrus:   { vars: {"--bg":"#0a0f06","--panel":"#121b0a","--panel2":"#0f1608","--text":"#f7fee7","--muted":"#cfe6a1","--accent":"#a3e635","--danger":"#fb7185","--line":"rgba(255,255,255,.10)","--line2":"rgba(255,255,255,.18)"}, font:'ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial' },
    galaxy:   { vars: {"--bg":"#070612","--panel":"#120f2a","--panel2":"#0d0b20","--text":"#f5f3ff","--muted":"#c4b5fd","--accent":"#7c3aed","--danger":"#22d3ee","--line":"rgba(255,255,255,.10)","--line2":"rgba(255,255,255,.18)"}, font:'ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial' },

    coral:    { vars: {"--bg":"#0f0a0a","--panel":"#1b1213","--panel2":"#150e0f","--text":"#fff7f6","--muted":"#ffc2b8","--accent":"#fb7185","--danger":"#f59e0b","--line":"rgba(255,255,255,.10)","--line2":"rgba(255,255,255,.18)"}, font:'ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial' },
    sky:      { vars: {"--bg":"#060f16","--panel":"#0b1b2a","--panel2":"#081521","--text":"#ecfeff","--muted":"#b3e7ff","--accent":"#38bdf8","--danger":"#fb7185","--line":"rgba(255,255,255,.10)","--line2":"rgba(255,255,255,.18)"}, font:'ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial' },
    mint:     { vars: {"--bg":"#06110e","--panel":"#0b201a","--panel2":"#081a15","--text":"#ecfdf5","--muted":"#a7f3d0","--accent":"#34d399","--danger":"#fb7185","--line":"rgba(255,255,255,.10)","--line2":"rgba(255,255,255,.18)"}, font:'ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial' },
    sunflower:{ vars: {"--bg":"#0e0d06","--panel":"#1a170a","--panel2":"#141206","--text":"#fffbeb","--muted":"#fde68a","--accent":"#fbbf24","--danger":"#f97316","--line":"rgba(255,255,255,.10)","--line2":"rgba(255,255,255,.18)"}, font:'ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial' },
    candy:    { vars: {"--bg":"#0b0612","--panel":"#140a22","--panel2":"#0f071a","--text":"#faf5ff","--muted":"#e9d5ff","--accent":"#a855f7","--danger":"#22d3ee","--line":"rgba(255,255,255,.10)","--line2":"rgba(255,255,255,.18)"}, font:'ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial' },

    neonPulse:   { vars: {"--bg":"#050510","--panel":"#0f0f25","--panel2":"#0b0b1c","--text":"#f0f9ff","--muted":"#a5f3fc","--accent":"#00f5ff","--danger":"#ff00c8","--line":"rgba(255,255,255,.10)","--line2":"rgba(255,255,255,.18)"}, font:'ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial' },
    electricLime:{ vars: {"--bg":"#0a1406","--panel":"#14250b","--panel2":"#0f1c08","--text":"#f7fee7","--muted":"#d9f99d","--accent":"#84cc16","--danger":"#f43f5e","--line":"rgba(255,255,255,.10)","--line2":"rgba(255,255,255,.18)"}, font:'ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial' },
    aquaVibe:    { vars: {"--bg":"#061214","--panel":"#0c2226","--panel2":"#08191c","--text":"#ecfeff","--muted":"#99f6e4","--accent":"#14b8a6","--danger":"#fb7185","--line":"rgba(255,255,255,.10)","--line2":"rgba(255,255,255,.18)"}, font:'ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial' },
    pinkRush:    { vars: {"--bg":"#14060e","--panel":"#2a0d1e","--panel2":"#200a17","--text":"#fff1f7","--muted":"#fbcfe8","--accent":"#ec4899","--danger":"#22d3ee","--line":"rgba(255,255,255,.10)","--line2":"rgba(255,255,255,.18)"}, font:'ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial' },
    ultraViolet: { vars: {"--bg":"#0a0514","--panel":"#1a0f2e","--panel2":"#130b23","--text":"#f5f3ff","--muted":"#ddd6fe","--accent":"#8b5cf6","--danger":"#f472b6","--line":"rgba(255,255,255,.10)","--line2":"rgba(255,255,255,.18)"}, font:'ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial' },

    premiumGlass:{ vars: {"--bg":"#05070b","--panel":"rgba(20,24,34,.55)","--panel2":"rgba(12,16,24,.45)","--text":"#eef2ff","--muted":"#b6c1d8","--accent":"#22d3ee","--danger":"#fb7185","--line":"rgba(255,255,255,.12)","--line2":"rgba(255,255,255,.18)"}, font:'ui-sans-serif, system-ui, -apple-system, Segoe UI, Inter, Roboto, Helvetica, Arial' }
  };

  const themeSelect = $("#themeSelect");
  const THEME_KEY = "mf_theme_v12";
  function applyTheme(key) {
    const t = THEMES[key] || THEMES.midnight;
    document.documentElement.setAttribute("data-theme", key);
    Object.entries(t.vars).forEach(([k, v]) => document.documentElement.style.setProperty(k, v));
    document.documentElement.style.setProperty("--sans", t.font);
    localStorage.setItem(THEME_KEY, key);
    if (themeSelect) themeSelect.value = key;
  }
  applyTheme(localStorage.getItem(THEME_KEY) || "midnight");
  themeSelect?.addEventListener("change", () => applyTheme(themeSelect.value));

  // ================================================================
  //  DISPLAY CONTROLS — device detection + manual scaling
  // ================================================================
  const DISPLAY_KEY = "mf_display_v1";

  const btnDisplayPanel  = document.querySelector("#btnDisplayPanel");
  const displayPanel     = document.querySelector("#displayPanel");
  const sliderFontSize   = document.querySelector("#sliderFontSize");
  const sliderSpacing    = document.querySelector("#sliderSpacing");
  const sliderTableSize  = document.querySelector("#sliderTableSize");
  const valFontSize      = document.querySelector("#valFontSize");
  const valSpacing       = document.querySelector("#valSpacing");
  const valTableSize     = document.querySelector("#valTableSize");
  const btnDisplayReset  = document.querySelector("#btnDisplayReset");
  const displayDeviceLbl = document.querySelector("#displayDeviceLabel");
  const presetBtns       = document.querySelectorAll(".display-preset-btn");

  // Presets: [fontScale, spacingScale, tableScale]
  const PRESETS = {
    desktop: { font: 100, spacing: 100, table: 100 },
    tablet:  { font:  90, spacing:  80, table:  85 },
    mobile:  { font:  80, spacing:  65, table:  75 },
  };

  // ── Auto-detect device type on load ──────────────────────────
  function detectDevice() {
    const w = window.innerWidth;
    if (w <= 480)  return "mobile";
    if (w <= 900)  return "tablet";
    return "desktop";
  }

  // ── Apply CSS variables from slider values ─────────────────
  function applyScales(font, spacing, table) {
    const root = document.documentElement.style;
    root.setProperty("--scale-font",    (font    / 100).toFixed(2));
    root.setProperty("--scale-spacing", (spacing / 100).toFixed(2));
    root.setProperty("--scale-table",   (table   / 100).toFixed(2));

    if (valFontSize)  valFontSize.textContent  = font    + "%";
    if (valSpacing)   valSpacing.textContent   = spacing + "%";
    if (valTableSize) valTableSize.textContent = table   + "%";

    if (sliderFontSize)  sliderFontSize.value  = font;
    if (sliderSpacing)   sliderSpacing.value   = spacing;
    if (sliderTableSize) sliderTableSize.value = table;
  }

  // ── Apply device class to <html> ───────────────────────────
  function applyDeviceClass(device) {
    document.documentElement.classList.remove("display-desktop","display-tablet","display-mobile");
    document.documentElement.classList.add("display-" + device);
    presetBtns.forEach(b => b.classList.toggle("is-active", b.dataset.preset === device));
    if (displayDeviceLbl) {
      const labels = { desktop:"🖥 Desktop", tablet:"📱 Tablet", mobile:"📲 Mobile" };
      displayDeviceLbl.textContent = labels[device] || "";
    }
  }

  // ── Load saved or auto-detect ──────────────────────────────
  function loadDisplaySettings() {
    try {
      const saved = JSON.parse(localStorage.getItem(DISPLAY_KEY) || "null");
      if (saved) {
        applyDeviceClass(saved.device || "desktop");
        applyScales(saved.font ?? 100, saved.spacing ?? 100, saved.table ?? 100);
        return;
      }
    } catch {}
    // No saved — auto detect
    const device = detectDevice();
    const preset = PRESETS[device];
    applyDeviceClass(device);
    applyScales(preset.font, preset.spacing, preset.table);

    // If not desktop, show a one-time toast
    if (device !== "desktop") {
      setTimeout(() => {
        if (displayDeviceLbl) displayDeviceLbl.textContent = "Auto-detected: " + device;
      }, 500);
    }
  }

  // ── Save settings ──────────────────────────────────────────
  function saveDisplaySettings() {
    const current = document.documentElement.className.match(/display-(\w+)/);
    const device = current ? current[1] : "desktop";
    localStorage.setItem(DISPLAY_KEY, JSON.stringify({
      device,
      font:    Number(sliderFontSize?.value  ?? 100),
      spacing: Number(sliderSpacing?.value   ?? 100),
      table:   Number(sliderTableSize?.value ?? 100),
    }));
  }

  // ── Toggle panel open/close ────────────────────────────────
  btnDisplayPanel?.addEventListener("click", (e) => {
    e.stopPropagation();
    const isOpen = displayPanel?.style.display !== "none";
    if (displayPanel) displayPanel.style.display = isOpen ? "none" : "block";
  });

  // Close panel when clicking outside
  document.addEventListener("click", (e) => {
    if (displayPanel && !displayPanel.contains(e.target) && e.target !== btnDisplayPanel) {
      displayPanel.style.display = "none";
    }
  });

  // ── Preset buttons ─────────────────────────────────────────
  presetBtns.forEach(btn => {
    btn.addEventListener("click", () => {
      const preset = PRESETS[btn.dataset.preset];
      if (!preset) return;
      applyDeviceClass(btn.dataset.preset);
      applyScales(preset.font, preset.spacing, preset.table);
      saveDisplaySettings();
    });
  });

  // ── Sliders ────────────────────────────────────────────────
  sliderFontSize?.addEventListener("input", () => {
    applyScales(
      Number(sliderFontSize.value),
      Number(sliderSpacing?.value  ?? 100),
      Number(sliderTableSize?.value ?? 100)
    );
    saveDisplaySettings();
  });
  sliderSpacing?.addEventListener("input", () => {
    applyScales(
      Number(sliderFontSize?.value  ?? 100),
      Number(sliderSpacing.value),
      Number(sliderTableSize?.value ?? 100)
    );
    saveDisplaySettings();
  });
  sliderTableSize?.addEventListener("input", () => {
    applyScales(
      Number(sliderFontSize?.value  ?? 100),
      Number(sliderSpacing?.value   ?? 100),
      Number(sliderTableSize.value)
    );
    saveDisplaySettings();
  });

  // ── Reset ──────────────────────────────────────────────────
  btnDisplayReset?.addEventListener("click", () => {
    localStorage.removeItem(DISPLAY_KEY);
    const device = detectDevice();
    const preset = PRESETS[device];
    applyDeviceClass(device);
    applyScales(preset.font, preset.spacing, preset.table);
    saveDisplaySettings();
  });

  // ── Re-detect on window resize (debounced) ─────────────────
  let resizeTimer;
  window.addEventListener("resize", () => {
    clearTimeout(resizeTimer);
    resizeTimer = setTimeout(() => {
      // Only auto-adjust if user hasn't manually saved a preference
      const saved = localStorage.getItem(DISPLAY_KEY);
      if (!saved) {
        const device = detectDevice();
        const preset = PRESETS[device];
        applyDeviceClass(device);
        applyScales(preset.font, preset.spacing, preset.table);
      }
    }, 300);
  });

  // ── Init ───────────────────────────────────────────────────
  loadDisplaySettings();


  // ---------------------------
  // UI Effects toggles (Glow / Gradient Header / Glass UI)
  // ---------------------------
  const toggleGlow = $("#toggleGlow");
  const toggleGradient = $("#toggleGradient");
  const toggleGlass = $("#toggleGlass");

  const UI_TOGGLES_KEY = "mf_ui_toggles_v17";

  function applyToggles(state) {
    const s = state || { glow: false, gradient: false, glass: false };
    document.documentElement.classList.toggle("glow-on", !!s.glow);
    document.documentElement.classList.toggle("header-gradient", !!s.gradient);
    document.documentElement.classList.toggle("glass-on", !!s.glass);

    if (toggleGlow) toggleGlow.checked = !!s.glow;
    if (toggleGradient) toggleGradient.checked = !!s.gradient;
    if (toggleGlass) toggleGlass.checked = !!s.glass;
  }

  function loadToggles() {
    try { return JSON.parse(localStorage.getItem(UI_TOGGLES_KEY) || "null"); } catch { return null; }
  }
  function saveToggles(state) {
    localStorage.setItem(UI_TOGGLES_KEY, JSON.stringify(state));
  }

  let uiState = loadToggles() || { glow: true, gradient: false, glass: false };
  applyToggles(uiState);

  function onToggleChange() {
    uiState = {
      glow: !!toggleGlow?.checked,
      gradient: !!toggleGradient?.checked,
      glass: !!toggleGlass?.checked
    };
    applyToggles(uiState);
    saveToggles(uiState);
  }

  toggleGlow?.addEventListener("change", onToggleChange);
  toggleGradient?.addEventListener("change", onToggleChange);
  toggleGlass?.addEventListener("change", onToggleChange);


  // Tabs
  const tabs = $$(".tab");
  const panels = $$(".panel");
  const TAB_KEY = "mf_active_tab_v12";
  function activateTab(name) {
    tabs.forEach(t => {
      const active = t.dataset.tab === name;
      t.classList.toggle("is-active", active);
      t.setAttribute("aria-selected", active ? "true" : "false");
    });
    panels.forEach(p => p.classList.toggle("is-active", p.id === `panel-${name}`));
    localStorage.setItem(TAB_KEY, name);
  }
  tabs.forEach(t => t.addEventListener("click", () => activateTab(t.dataset.tab)));
  activateTab(localStorage.getItem(TAB_KEY) || "contracts");


  // ---------------------------
  // Transactions Sub-tabs
  // ---------------------------
  const subtabs = $$(".subtab");
  const subpanels = $$(".subpanel");
  const SUBTAB_KEY = "mf_transactions_subtab_v15";

  function activateSubtab(name) {
    subtabs.forEach(t => {
      const active = t.dataset.subtab === name;
      t.classList.toggle("is-active", active);
      t.setAttribute("aria-selected", active ? "true" : "false");
    });
    subpanels.forEach(p => p.classList.toggle("is-active", p.id === `subpanel-${name}`));
    localStorage.setItem(SUBTAB_KEY, name);
  }

  if (subtabs.length) {
    subtabs.forEach(t => t.addEventListener("click", () => activateSubtab(t.dataset.subtab)));
    activateSubtab(localStorage.getItem(SUBTAB_KEY) || "cashReceived");
  }


  // ---------------------------
  // Contracts Store + Rendering
  // ---------------------------
  const table = $("#contractsTable");
  const rowCountEl = $("#contractsRowCount");
  const selectedEl = $("#contractsSelected");

  const btnAdd = $("#btnAddContract");
  const btnEdit = $("#btnEditSelected");
  const btnDel = $("#btnDeleteSelected");
  const btnImport = $("#btnImportExcel");
  const btnExport = $("#btnExportExcel");
  const btnRefresh = $("#btnRefresh");

  // ── Top scrollbar mirror for contracts table ──
  const contractsGridWrap    = $("#contractsGridWrap");
  const contractsTopScroll   = $("#contractsTopScroll");
  const contractsTopInner    = $("#contractsTopScrollInner");

  function syncTopScrollWidth() {
    if (contractsTopInner && table) {
      contractsTopInner.style.width = table.offsetWidth + "px";
    }
  }

  if (contractsGridWrap && contractsTopScroll) {
    // Sync scroll positions both ways
    contractsTopScroll.addEventListener("scroll", () => {
      contractsGridWrap.scrollLeft = contractsTopScroll.scrollLeft;
    });
    contractsGridWrap.addEventListener("scroll", () => {
      contractsTopScroll.scrollLeft = contractsGridWrap.scrollLeft;
    });
    // Update inner width on render and resize
    new ResizeObserver(syncTopScrollWidth).observe(table);
    syncTopScrollWidth();
  }
  const btnSaveData = $("#btnSaveData");
  const btnLoadData = $("#btnLoadData");
  const fileLoadJson = $("#fileLoadJson");
  const fileImport = $("#fileImport");

  let selectedContractNo = null;

  // Filter controls
  const filterCategory = $("#filterCategory");
  const filterDateBox = $("#filterDateBox");
  const filterDateMode = $("#filterDateMode");
  const filterMonthField = $("#filterMonthField");
  const filterYearField = $("#filterYearField");
  const filterMonth = $("#filterMonth");
  const filterYear = $("#filterYear");

  const filterTextBox = $("#filterTextBox");
  const filterText = $("#filterText");

  const filterAmountBox = $("#filterAmountBox");
  const filterMin = $("#filterMin");
  const filterMax = $("#filterMax");

  const btnApplyFilter = $("#btnApplyFilter");
  const btnClearFilter = $("#btnClearFilter");

  const FILTER_KEY = "mf_contracts_filter_v12";
  const STORE_KEY = "mf_contracts_store_v13";

  function parseMoney(s) {
    const n = Number(String(s ?? "").replace(/[^0-9.-]/g, ""));
    return Number.isFinite(n) ? n : 0;
  }
  
  function computeRemaining(amount, discount, totalPaid){
    const a = Number(amount) || 0;
    const d = Number(discount) || 0;
    const p = Number(totalPaid) || 0;
    return a - d - p;
  }

function fmtMoney(n) {
    return (Number(n) || 0).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }
  function pad2(n){ return String(n).padStart(2, "0"); }
  function parseMMDDYYYY(s) {
    const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(String(s || "").trim());
    if (!m) return null;
    const mm = Number(m[1]), dd = Number(m[2]), yyyy = Number(m[3]);
    const d = new Date(yyyy, mm - 1, dd);
    if (!Number.isFinite(d.getTime())) return null;
    if (d.getFullYear() !== yyyy || d.getMonth() !== (mm - 1) || d.getDate() !== dd) return null;
    return d;
  }
  function formatMMDDYYYY(d) {
    const dt = (d instanceof Date) ? d : new Date(d);
    if (!Number.isFinite(dt.getTime())) return "";
    return `${pad2(dt.getMonth()+1)}/${pad2(dt.getDate())}/${dt.getFullYear()}`;
  }
  function yyyyMMddToMMDDYYYY(yyyyMMdd) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(yyyyMMdd || ""));
    return m ? `${m[2]}/${m[3]}/${m[1]}` : "";
  }
  function mmddyyyyToYyyyMMdd(mmddyyyy) {
    const m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(String(mmddyyyy || ""));
    return m ? `${m[3]}-${m[1]}-${m[2]}` : "";
  }

  function monthKeyFromDate(dateText) {
    const d = parseMMDDYYYY(dateText);
    if (!d) return "unknown";
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
  }
  function monthLabelFromKey(key) {
    if (key === "unknown") return "Unknown Date";
    const m = /^(\d{4})-(\d{2})$/.exec(key);
    if (!m) return key;
    const yyyy = Number(m[1]);
    const mm = Number(m[2]);
    const names = ["January","February","March","April","May","June","July","August","September","October","November","December"];
    return `${names[mm - 1]} ${yyyy}`;
  }

  // Master store (single source of truth)
  /** @type {Array<{
   *  date:string, contract:string, deceased:string, casket:string, address:string,
   *  amount:number, inhaus:number, bai:number, gl:number, gcash:number, cash:number,
   *  discount:number, lastPayment:string
   * }>} */
  let contractsStore = [];

  function saveStore() {
    // No-op: individual saves happen via DB.saveContract() per operation
  }

  function loadStore() {
    // No-op: data loaded async via initFromSupabase()
    return null;
  }

  function clearStoredData() {
    // No-op for Supabase version
  }

  function normalizeText(s){ return String(s ?? "").toLowerCase().trim(); }

  function calcComputed(c) {
    // totalPaid = only actual customer cash payments
    const totalPaid    = Math.max(0,
      (Number(c.inhaus)||0) + (Number(c.gl)||0) +
      (Number(c.gcash)||0)  + (Number(c.cash)||0)
    );
    // These reduce remaining balance but are NOT customer payments
    const baiAssist    = Number(c.baiAssist)    || 0;  // from BAI Tab
    const dswdAfterTax = Number(c.dswdAfterTax) || 0;  // from DSWD Tab
    const dswdDiscount = Number(c.dswdDiscount) || 0;  // from DSWD Tab
    const discount     = Number(c.discount)     || 0;
    const remaining    = (Number(c.amount) || 0) - totalPaid - discount - baiAssist - dswdAfterTax - dswdDiscount;
    return { totalPaid, remaining, baiAssist, dswdAfterTax, dswdDiscount };
  }

  function createGroupRow(label) {
    const tr = document.createElement("tr");
    tr.dataset.rowType = "monthHeader";
    tr.classList.add("group-row");
    const td = document.createElement("td");
    td.colSpan = 17;
    td.innerHTML = `<span class="group-chip"><span class="dot"></span><span>${label}</span></span>`;
    tr.appendChild(td);
    return tr;
  }

  function createSpacerRow() {
    const tr = document.createElement("tr");
    tr.dataset.rowType = "spacer";
    tr.classList.add("spacer-row");
    const td = document.createElement("td");
    td.colSpan = 17;
    tr.appendChild(td);
    return tr;
  }

  function createTotalRow(label, totals, rowType) {
    const tr = document.createElement("tr");
    tr.dataset.rowType = rowType;
    tr.classList.add(rowType === "grandTotal" ? "grand-total-row" : "total-row");

    // 0..3 blank, 4 label under Address column
    for (let i = 0; i < 4; i++) tr.appendChild(document.createElement("td"));

    const tdLabel = document.createElement("td");
    tdLabel.textContent = label;
    tdLabel.classList.add("total-label");
    tr.appendChild(tdLabel);

    function tdNum(val) {
      const td = document.createElement("td");
      td.classList.add("num");
      td.textContent = fmtMoney(val);
      return td;
    }

    tr.appendChild(tdNum(totals.amount));
    tr.appendChild(tdNum(totals.inhaus));
    tr.appendChild(tdNum(totals.gl));
    tr.appendChild(tdNum(totals.gcash));
    tr.appendChild(tdNum(totals.cash));
    tr.appendChild(tdNum(totals.dswdAfterTax || 0));
    tr.appendChild(tdNum(totals.discount));
    tr.appendChild(tdNum(totals.dswdDiscount || 0));
    tr.appendChild(tdNum(totals.baiAssist || 0));
    tr.appendChild(tdNum(totals.totalPaid));

    const tdLP = document.createElement("td");
    tdLP.textContent = "";
    tr.appendChild(tdLP);

    tr.appendChild(tdNum(totals.remaining));
    return tr;
  }

  function createDataRow(c) {
    const tr = document.createElement("tr");
    tr.dataset.rowType = "data";
    tr.dataset.contract = c.contract;

    const { totalPaid, remaining } = calcComputed(c);

    const cells = [
      { text: c.date },
      { text: c.contract },
      { text: c.deceased },
      { text: c.casket },
      { text: c.address },

      { text: fmtMoney(c.amount),                     num: true },
      { text: fmtMoney(c.inhaus),                     num: true },
      { text: fmtMoney(c.gl),                         num: true },
      { text: fmtMoney(c.gcash),                      num: true },
      { text: fmtMoney(c.cash),                       num: true },
      { text: fmtMoney(c.dswdAfterTax || 0),          num: true, computed: true },
      { text: fmtMoney(c.discount),                   num: true },
      { text: fmtMoney(c.dswdDiscount || 0),          num: true, computed: true },
      { text: fmtMoney(c.baiAssist || 0),             num: true, computed: true },

      { text: fmtMoney(totalPaid),                    num: true, computed: true },
      { text: c.lastPayment || "—" },
      { text: fmtMoney(remaining),                    num: true, computed: true },
    ];

    cells.forEach((cell, idx) => {
      const td = document.createElement("td");
      td.textContent = cell.text ?? "";
      if (cell.num) td.classList.add("num");
      if (cell.computed) td.classList.add("computed");

      const dataCol = ["","","","","",
        "amount","inhaus","gl","gcash","cash",
        "dswdAfterTax","discount","dswdDiscount","baiAssist",
        "totalPaid","lastPayment","remaining"][idx];
      if (dataCol) td.setAttribute("data-col", dataCol);

      tr.appendChild(td);
    });

    return tr;
  }

  // Current filter (null = no filter)
  let activeFilter = null;

  function applyFilterToRows(rows, filter) {
    if (!filter) return rows;
    const cat = filter.category;

    if (cat === "date") {
      const mode = filter.dateMode; // monthYear|month|year
      const m = filter.month; // 1..12 or null
      const y = filter.year; // number or null
      return rows.filter(r => {
        const d = parseMMDDYYYY(r.date);
        if (!d) return false;
        const mm = d.getMonth() + 1;
        const yy = d.getFullYear();
        if (mode === "monthYear") {
          if (m == null || y == null) return false;
          return mm === m && yy === y;
        }
        if (mode === "month") {
          if (m == null) return false;
          return mm === m;
        }
        if (mode === "year") {
          if (y == null) return false;
          return yy === y;
        }
        return true;
      });
    }

    if (cat === "amount") {
      const min = filter.min;
      const max = filter.max;
      return rows.filter(r => {
        const a = Number(r.amount) || 0;
        if (min != null && a < min) return false;
        if (max != null && a > max) return false;
        return true;
      });
    }

    // text categories
    const needle = normalizeText(filter.text || "");
    if (!needle) return rows;

    if (cat === "contract") return rows.filter(r => normalizeText(r.contract).includes(needle));
    if (cat === "deceased") return rows.filter(r => normalizeText(r.deceased).includes(needle));
    if (cat === "casket") return rows.filter(r => normalizeText(r.casket).includes(needle));
    if (cat === "address") return rows.filter(r => normalizeText(r.address).includes(needle));

    return rows;
  }

  function renderContracts() {
    // Derive visible rows from store + filter
    const visible = applyFilterToRows(contractsStore.slice(), activeFilter);

    // Sort by month, then date, then contract #
    visible.sort((a, b) => {
      const ak = monthKeyFromDate(a.date);
      const bk = monthKeyFromDate(b.date);
      if (ak !== bk) {
        if (ak === "unknown") return 1;
        if (bk === "unknown") return -1;
        return ak.localeCompare(bk);
      }
      const ad = parseMMDDYYYY(a.date)?.getTime() ?? Number.POSITIVE_INFINITY;
      const bd = parseMMDDYYYY(b.date)?.getTime() ?? Number.POSITIVE_INFINITY;
      if (ad !== bd) return ad - bd;
      return (a.contract || "").localeCompare(b.contract || "");
    });

    // Group
    const groups = new Map();
    for (const r of visible) {
      const k = monthKeyFromDate(r.date);
      if (!groups.has(k)) groups.set(k, []);
      groups.get(k).push(r);
    }

    const keys = Array.from(groups.keys()).sort((a, b) => {
      if (a === "unknown" && b !== "unknown") return 1;
      if (b === "unknown" && a !== "unknown") return -1;
      return a.localeCompare(b);
    });

    // Rebuild tbody
    table.tBodies[0].innerHTML = "";
    selectedContractNo = null;
    selectedEl.textContent = "Selected: —";

    const grand = { amount:0, inhaus:0, gl:0, gcash:0, cash:0, dswdAfterTax:0, discount:0, dswdDiscount:0, baiAssist:0, totalPaid:0, remaining:0 };
    let visibleDataCount = 0;

    for (let gi = 0; gi < keys.length; gi++) {
      const key = keys[gi];
      const rows = groups.get(key) || [];
      if (!rows.length) continue;

      table.tBodies[0].appendChild(createGroupRow(monthLabelFromKey(key)));

      const totals = { amount:0, inhaus:0, gl:0, gcash:0, cash:0, dswdAfterTax:0, discount:0, dswdDiscount:0, baiAssist:0, totalPaid:0, remaining:0 };

      for (const c of rows) {
        const tr = createDataRow(c);
        table.tBodies[0].appendChild(tr);
        visibleDataCount++;

        const computed = calcComputed(c);
        totals.amount      += c.amount      || 0;
        totals.inhaus      += c.inhaus      || 0;
        totals.gl          += c.gl          || 0;
        totals.gcash       += c.gcash       || 0;
        totals.cash        += c.cash        || 0;
        totals.dswdAfterTax+= c.dswdAfterTax|| 0;
        totals.discount    += c.discount    || 0;
        totals.dswdDiscount+= c.dswdDiscount|| 0;
        totals.baiAssist   += c.baiAssist   || 0;
        totals.totalPaid   += computed.totalPaid;
        totals.remaining   += computed.remaining;
      }

      table.tBodies[0].appendChild(createTotalRow("TOTAL", totals, "monthTotal"));
      // Extra space between months (after TOTAL row)
      if (gi < keys.length - 1) {
        table.tBodies[0].appendChild(createSpacerRow());
      }

      grand.amount       += totals.amount;
      grand.inhaus       += totals.inhaus;
      grand.gl           += totals.gl;
      grand.gcash        += totals.gcash;
      grand.cash         += totals.cash;
      grand.dswdAfterTax += totals.dswdAfterTax;
      grand.discount     += totals.discount;
      grand.dswdDiscount += totals.dswdDiscount;
      grand.baiAssist    += totals.baiAssist;
      grand.totalPaid    += totals.totalPaid;
      grand.remaining    += totals.remaining;
    }

    if (visibleDataCount > 0) {
      table.tBodies[0].appendChild(createSpacerRow());
      table.tBodies[0].appendChild(createTotalRow("GRAND TOTAL", grand, "grandTotal"));
    }

    rowCountEl.textContent = `Rows: ${visibleDataCount}`;
    syncTopScrollWidth();
  }

  function selectByContract(contractNo) {
    const key = String(contractNo || "").trim().toLowerCase();
    selectedContractNo = key || null;

    $$(".grid tbody tr").forEach(r => r.classList.remove("is-selected"));
    if (!key) {
      selectedEl.textContent = "Selected: —";
      return;
    }
    const tr = table.tBodies[0].querySelector(`tr[data-row-type="data"][data-contract="${CSS.escape(key)}"]`);
    if (tr) {
      tr.classList.add("is-selected");
      selectedEl.textContent = "Selected: " + (tr.cells[1]?.innerText || key);
    } else {
      selectedEl.textContent = "Selected: —";
    }
  }

  // Click row to select, double-click to edit
  table.addEventListener("click", (e) => {
    const tr = e.target.closest("tr");
    if (!tr || tr.parentElement?.tagName !== "TBODY") return;
    if ((tr.dataset.rowType || "data") !== "data") return;
    selectByContract(tr.dataset.contract);
  });

  table.addEventListener("dblclick", (e) => {
    const tr = e.target.closest("tr");
    if (!tr || tr.parentElement?.tagName !== "TBODY") return;
    if ((tr.dataset.rowType || "data") !== "data") return;
    selectByContract(tr.dataset.contract);
    openModal("edit", tr.dataset.contract);
  });

  // ---------------------------
  // Modal (Add/Edit)
  // ---------------------------
  const overlay = $("#contractOverlay");
  const modal = $("#contractModal");
  const titleEl = $("#contractTitle");
  const subtitleEl = $("#contractSubtitle");
  const form = $("#contractForm");
  const btnClose = $("#btnCloseContract");
  const btnCancel = $("#btnCancelContract");
  const submitBtn = $("#btnSubmitContract");

  const cDate = $("#cDate");
  const cContractNo = $("#cContractNo");
  const cDeceased = $("#cDeceased");
  const cCasket = $("#cCasket");
  const cAddress = $("#cAddress");
  const cAmount = $("#cAmount");
  const cDiscount = $("#cDiscount");

  let mode = "add";
  let editingContractKey = null;

  function ensureMoneyInput(el) {
    const n = parseMoney(el.value);
    el.value = (Number.isFinite(n) ? n : 0).toFixed(2);
  }
  cAmount?.addEventListener("blur", () => ensureMoneyInput(cAmount));
  cDiscount?.addEventListener("blur", () => ensureMoneyInput(cDiscount));

  function openModal(newMode, contractKeyOrNull = null) {
    mode = newMode;
    editingContractKey = null;

    if (mode === "add") {
      titleEl.textContent = "Add Contract";
      subtitleEl.textContent = "Enter contract details (local-only). Payments start at 0.";
      submitBtn.textContent = "Add Contract";
      if (btnPaymentSummary) btnPaymentSummary.style.display = "none";

      cDate.value = new Date().toISOString().slice(0,10);
      cContractNo.value = "";
      cDeceased.value = "";
      cCasket.value = "";
      cAddress.value = "";
      cAmount.value = "0.00";
      cDiscount.value = "0.00";
    } else {
      titleEl.textContent = "Edit Contract";
      subtitleEl.textContent = "Update contract details. Payments stay in the table (linked later).";
      submitBtn.textContent = "Save Changes";
      if (btnPaymentSummary) btnPaymentSummary.style.display = "inline-flex";

      // Ensure payment stores are loaded so Payment Summary has fresh data
      try{ const t = loadCashStore(); if(Array.isArray(t)) cashStore = t; }catch{}
      try{ const t = loadBankStore(); if(Array.isArray(t)) bankStore = t; }catch{}

      const key = String(contractKeyOrNull || "").toLowerCase().trim();
      const found = contractsStore.find(x => normalizeText(x.contract) === key);
      if (!found) {
        alert("Could not find the selected contract in memory.");
        return;
      }
      editingContractKey = key;

      cDate.value = mmddyyyyToYyyyMMdd(found.date) || new Date().toISOString().slice(0,10);
      cContractNo.value = found.contract || "";
      cDeceased.value = found.deceased || "";
      cCasket.value = found.casket || "";
      cAddress.value = found.address || "";
      cAmount.value = (Number(found.amount) || 0).toFixed(2);
      cDiscount.value = (Number(found.discount) || 0).toFixed(2);
    }

    overlay.classList.add("is-open");
    modal.classList.add("is-open");
    setTimeout(() => cContractNo.focus(), 0);
  }

  function closeModal() {
    overlay.classList.remove("is-open");
    modal.classList.remove("is-open");
    mode = "add";
    editingContractKey = null;
  }

  btnAdd?.addEventListener("click", () => openModal("add"));
  btnEdit?.addEventListener("click", () => {
    if (!selectedContractNo) return alert("Please select a row first.");
    openModal("edit", selectedContractNo);
  });

  btnClose?.addEventListener("click", closeModal);
  btnCancel?.addEventListener("click", closeModal);
  overlay?.addEventListener("click", closeModal);
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && modal.classList.contains("is-open")) closeModal();
  });

  // ---------------------------
  // Payment Summary Modal
  // ---------------------------
  const paymentSummaryOverlay = $("#paymentSummaryOverlay");
  const paymentSummaryModal   = $("#paymentSummaryModal");
  const btnClosePaymentSummary = $("#btnClosePaymentSummary");
  const paymentSummarySubtitle = $("#paymentSummarySubtitle");
  const paymentSummaryRows     = $("#paymentSummaryRows");
  const paymentSummaryTotal    = $("#paymentSummaryTotal");
  const paymentSummaryEmpty    = $("#paymentSummaryEmpty");
  const paymentSummaryTable    = $("#paymentSummaryTable");
  const btnPaymentSummary      = $("#btnPaymentSummary");

  function openPaymentSummary(contractNo) {
    const normKey = normalizeText(contractNo);

    // Gather all matching cash received entries
    const cashEntries = (cashStore || [])
      .filter(r => normalizeText(r.contract ?? r.contractNumber ?? "") === normKey)
      .map(r => ({
        date:   r.date || "",
        name:   r.client ?? r.clientName ?? r.nameOfClient ?? "",
        method: r.particular || "Cash",
        amount: Number(r.amount) || 0,
      }));

    // Gather all matching bank received entries
    const bankEntries = (bankStore || [])
      .filter(r => normalizeText(r.contract ?? r.contractNumber ?? "") === normKey)
      .map(r => ({
        date:   r.date || "",
        name:   r.client ?? r.clientName ?? r.nameOfClient ?? "",
        method: r.type || "Bank Deposit",
        amount: Number(r.amount) || 0,
      }));

    const allEntries = [...cashEntries, ...bankEntries]
      .sort((a, b) => (a.date > b.date ? 1 : a.date < b.date ? -1 : 0));

    const total = allEntries.reduce((s, r) => s + r.amount, 0);

    paymentSummarySubtitle.textContent = `Contract: ${contractNo}`;

    if (allEntries.length === 0) {
      paymentSummaryTable.style.display = "none";
      paymentSummaryEmpty.style.display = "block";
    } else {
      paymentSummaryTable.style.display = "table";
      paymentSummaryEmpty.style.display = "none";
      paymentSummaryRows.innerHTML = allEntries.map(r => `
        <tr style="border-bottom:1px solid var(--border,#333);">
          <td style="padding:9px 14px;">${escapeHtml(r.date)}</td>
          <td style="padding:9px 14px;">${escapeHtml(r.name)}</td>
          <td style="padding:9px 14px;">${escapeHtml(r.method)}</td>
          <td style="padding:9px 14px; text-align:right;">${fmtMoney(r.amount)}</td>
        </tr>
      `).join("");
      paymentSummaryTotal.textContent = fmtMoney(total);
    }

    paymentSummaryOverlay.classList.add("is-open");
    paymentSummaryModal.classList.add("is-open");
  }

  function closePaymentSummary() {
    paymentSummaryOverlay.classList.remove("is-open");
    paymentSummaryModal.classList.remove("is-open");
  }

  btnClosePaymentSummary?.addEventListener("click", closePaymentSummary);
  paymentSummaryOverlay?.addEventListener("click", closePaymentSummary);
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && paymentSummaryModal?.classList.contains("is-open")) closePaymentSummary();
  });

  // ---------------------------
  // Generate Payment Report PDF
  // ---------------------------
  const btnGeneratePaymentReport = $("#btnGeneratePaymentReport");

  btnGeneratePaymentReport?.addEventListener("click", function() {
    // Gather contract info
    const contractNo = editingContractKey
      ? (contractsStore.find(x => normalizeText(x.contract) === editingContractKey)?.contract || editingContractKey)
      : "";
    const contract = contractsStore.find(x => normalizeText(x.contract) === normalizeText(contractNo));
    if (!contract) return alert("Contract data not found.");

    const normKey = normalizeText(contractNo);

    // Gather payment rows
    const cashEntries = (cashStore || [])
      .filter(r => normalizeText(r.contract ?? r.contractNumber ?? "") === normKey)
      .map(r => ({ date: r.date || "", name: r.client ?? r.clientName ?? "", method: r.particular || "Cash", amount: Number(r.amount) || 0 }));
    const bankEntries = (bankStore || [])
      .filter(r => normalizeText(r.contract ?? r.contractNumber ?? "") === normKey)
      .map(r => ({ date: r.date || "", name: r.client ?? r.clientName ?? "", method: r.type || "Bank Deposit", amount: Number(r.amount) || 0 }));
    const allEntries = [...cashEntries, ...bankEntries].sort((a,b) => a.date > b.date ? 1 : a.date < b.date ? -1 : 0);
    const totalPaid  = allEntries.reduce((s,r) => s + r.amount, 0);
    const remaining  = (Number(contract.amount) || 0) - (Number(contract.discount) || 0) - totalPaid;

    const btn = btnGeneratePaymentReport;
    const origText = btn.textContent;
    btn.textContent = "⏳ Generating…";
    btn.disabled = true;

    function loadScript(src, cb) {
      const existing = document.querySelector(`script[src="${src}"]`);
      if (existing && existing._mfLoaded) { cb(); return; }
      if (existing) { existing.addEventListener("load", cb); return; }
      const s = document.createElement("script");
      s.src = src;
      s.onload = function() { s._mfLoaded = true; cb(); };
      s.onerror = function() { btn.textContent = origText; btn.disabled = false; alert("Failed to load PDF library. Check your internet connection."); };
      document.head.appendChild(s);
    }

    loadScript("https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js", function() {
      try {
        const { jsPDF } = window.jspdf;
        const doc = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });

        const PW = 210, PH = 297, ML = 18, MR = 18, MT = 18, MB = 18;
        const CW = PW - ML - MR;
        let y = MT;

        function checkNewPage(need) {
          if (y + need > PH - MB) { doc.addPage(); y = MT; }
        }

        // ── Header ──
        doc.setFont("helvetica", "bold");
        doc.setFontSize(16);
        doc.setTextColor(30, 30, 30);
        doc.text("MAGALLANES FUNERAL", PW / 2, y, { align: "center" });
        y += 7;
        doc.setFont("helvetica", "normal");
        doc.setFontSize(10);
        doc.setTextColor(80, 80, 80);
        doc.text("Contract Payment Report", PW / 2, y, { align: "center" });
        y += 5;
        doc.setDrawColor(180, 180, 180);
        doc.setLineWidth(0.4);
        doc.line(ML, y, PW - MR, y);
        y += 8;

        // ── Contract Details Section ──
        doc.setFillColor(240, 240, 240);
        doc.rect(ML, y, CW, 7, "F");
        doc.setFont("helvetica", "bold");
        doc.setFontSize(9);
        doc.setTextColor(0, 0, 0);
        doc.text("CONTRACT DETAILS", ML + 3, y + 4.8);
        y += 10;

        const labelX = ML;
        const valueX = ML + 48;
        const col2LabelX = ML + 95;
        const col2ValueX = ML + 143;

        function detailRow(label1, val1, label2, val2) {
          checkNewPage(7);
          doc.setFont("helvetica", "bold");
          doc.setFontSize(9);
          doc.setTextColor(80, 80, 80);
          doc.text(label1 + ":", labelX, y);
          doc.setFont("helvetica", "normal");
          doc.setTextColor(0, 0, 0);
          doc.text(String(val1 || "—"), valueX, y);
          if (label2 !== undefined) {
            doc.setFont("helvetica", "bold");
            doc.setTextColor(80, 80, 80);
            doc.text(label2 + ":", col2LabelX, y);
            doc.setFont("helvetica", "normal");
            doc.setTextColor(0, 0, 0);
            doc.text(String(val2 || "—"), col2ValueX, y);
          }
          y += 6;
        }

        detailRow("Contract #",    contract.contract,                   "Date",       contract.date);
        detailRow("Deceased",      contract.deceased,                   "Casket",      contract.casket);
        detailRow("Address",       contract.address,                    "",            undefined);
        detailRow("Contract Amt",  fmtMoney(Number(contract.amount)||0),"Discount",    fmtMoney(Number(contract.discount)||0));
        detailRow("Total Paid",    fmtMoney(totalPaid),                 "Remaining",   fmtMoney(remaining));

        y += 4;
        doc.setDrawColor(200, 200, 200);
        doc.setLineWidth(0.3);
        doc.line(ML, y, PW - MR, y);
        y += 8;

        // ── Payment History Section ──
        doc.setFillColor(240, 240, 240);
        doc.rect(ML, y, CW, 7, "F");
        doc.setFont("helvetica", "bold");
        doc.setFontSize(9);
        doc.setTextColor(0, 0, 0);
        doc.text("PAYMENT HISTORY", ML + 3, y + 4.8);
        y += 10;

        // Table columns: Date | Name | Payment Method | Amount
        const COLS = [28, 58, 62, 26];  // sum = 174 = CW

        function drawRow(cells, opts) {
          opts = opts || {};
          const bold    = opts.bold    || false;
          const fillRGB = opts.fill    || null;
          const fs      = opts.fs      || 8.5;
          const padH    = 2, padV = 1.8;

          doc.setFont("helvetica", bold ? "bold" : "normal");
          doc.setFontSize(fs);

          const wrapped = cells.map(function(cell, i) {
            return doc.splitTextToSize(String(cell || ""), COLS[i] - padH * 2);
          });
          const maxLines = wrapped.reduce(function(m, l) { return Math.max(m, l.length); }, 1);
          const lineH    = fs * 0.3528 * 1.35;
          const rowH     = Math.max(6.5, padV * 2 + maxLines * lineH);

          checkNewPage(rowH + 1);

          if (fillRGB) {
            doc.setFillColor(fillRGB[0], fillRGB[1], fillRGB[2]);
            doc.rect(ML, y, CW, rowH, "F");
          }
          doc.setDrawColor(160, 160, 160);
          doc.setLineWidth(0.2);
          doc.rect(ML, y, CW, rowH, "S");

          doc.setFont("helvetica", bold ? "bold" : "normal");
          doc.setFontSize(fs);
          doc.setTextColor(0, 0, 0);

          let cx = ML;
          cells.forEach(function(cell, i) {
            const cw = COLS[i];
            const isAmt = i === 3;
            if (i > 0) {
              doc.setDrawColor(160, 160, 160);
              doc.setLineWidth(0.2);
              doc.line(cx, y, cx, y + rowH);
            }
            const lines = wrapped[i];
            lines.forEach(function(line, li) {
              const tx = isAmt ? cx + cw - padH : cx + padH;
              const align = isAmt ? "right" : "left";
              doc.text(line, tx, y + padV + lineH * (li + 0.75), { align });
            });
            cx += cw;
          });
          y += rowH;
        }

        // Header row
        drawRow(["Date", "Name", "Payment Method", "Amount"], { bold: true, fill: [220, 220, 220], fs: 8.5 });

        if (allEntries.length === 0) {
          checkNewPage(10);
          doc.setFont("helvetica", "italic");
          doc.setFontSize(9);
          doc.setTextColor(150, 150, 150);
          doc.text("No payments recorded for this contract yet.", ML + 3, y + 6);
          y += 10;
        } else {
          allEntries.forEach(function(r, idx) {
            const fill = idx % 2 === 0 ? null : [248, 248, 248];
            drawRow([r.date, r.name, r.method, fmtMoney(r.amount)], { fill });
          });

          // Total row
          drawRow(["", "", "TOTAL PAYMENTS", fmtMoney(totalPaid)], { bold: true, fill: [230, 230, 230] });
        }

        y += 10;
        // Footer note
        checkNewPage(10);
        doc.setFont("helvetica", "italic");
        doc.setFontSize(8);
        doc.setTextColor(150, 150, 150);
        const now = new Date();
        doc.text("Generated: " + now.toLocaleDateString() + " " + now.toLocaleTimeString(), ML, y);

        // Save
        const safeName = (contract.contract || "contract").replace(/[^a-zA-Z0-9_\-]/g, "_");
        doc.save("PaymentReport_" + safeName + ".pdf");

      } catch(err) {
        alert("PDF generation failed: " + (err?.message || err));
      } finally {
        btn.textContent = origText;
        btn.disabled = false;
      }
    });
  });

  btnPaymentSummary?.addEventListener("click", () => {
    if (editingContractKey) openPaymentSummary(
      contractsStore.find(x => normalizeText(x.contract) === editingContractKey)?.contract || editingContractKey
    );
  });

  form?.addEventListener("submit", (e) => {
    e.preventDefault();

    const contractNo = (cContractNo.value || "").trim();
    const deceased = (cDeceased.value || "").trim();
    if (!contractNo) return alert("Contract Number is required.");
    if (!deceased) return alert("Deceased is required.");

    const date = yyyyMMddToMMDDYYYY(cDate.value) || "";
    const amount = parseMoney(cAmount.value);
    const discount = parseMoney(cDiscount.value);

    if (mode === "add") {
      const exists = contractsStore.some(x => normalizeText(x.contract) === normalizeText(contractNo));
      if (exists && !confirm("That Contract # already exists. Add anyway?")) return;

      const newRow = {
        date, contract: contractNo, deceased,
        casket: (cCasket.value || "").trim(),
        address: (cAddress.value || "").trim(),
        amount, inhaus: 0, bai: 0, gl: 0, gcash: 0, cash: 0,
        discount, lastPayment: "—"
      };
      DB.saveContract(newRow).then(saved => {
        if (saved) newRow.id = saved.id;
        contractsStore.push(newRow);
        closeModal();
        renderContracts();
        selectByContract(contractNo);
      });
      return;
    }

    // edit
    if (!editingContractKey) return;
    const idx = contractsStore.findIndex(x => normalizeText(x.contract) === editingContractKey);
    if (idx < 0) return;

    // If contract # changed, check duplicates
    const newKey = normalizeText(contractNo);
    if (newKey !== editingContractKey) {
      const dup = contractsStore.some((x, i) => i !== idx && normalizeText(x.contract) === newKey);
      if (dup && !confirm("That Contract # already exists. Save anyway?")) return;
    }

    const existing = contractsStore[idx];
    const updated = {
      ...existing, date, contract: contractNo, deceased,
      casket: (cCasket.value || "").trim(),
      address: (cAddress.value || "").trim(),
      amount, discount
    };
    contractsStore[idx] = updated;
    DB.saveContract(updated);

    closeModal();
    renderContracts();
    selectByContract(contractNo);
  });

  // Delete selected
  btnDel?.addEventListener("click", () => {
    if (!selectedContractNo) return alert("Please select a row first.");
    if (!confirm("Delete the selected contract?")) return;
    const key = selectedContractNo;
    const row = contractsStore.find(x => normalizeText(x.contract) === key);
    if (row?.id) DB.deleteContract(row.id);
    contractsStore = contractsStore.filter(x => normalizeText(x.contract) !== key);
    selectedContractNo = null;
    renderContracts();
  });

  // ---------------------------
  // Filter UI logic
  // ---------------------------
  function syncFilterUI() {
    const cat = filterCategory.value;
    filterDateBox.style.display = (cat === "date") ? "" : "none";
    filterTextBox.style.display = (cat === "contract" || cat === "deceased" || cat === "casket" || cat === "address") ? "" : "none";
    filterAmountBox.style.display = (cat === "amount") ? "" : "none";

    // date mode visibility
    const mode = filterDateMode.value;
    filterMonthField.style.display = (mode === "year") ? "none" : "";
    filterYearField.style.display = (mode === "month") ? "none" : "";
  }

  filterCategory?.addEventListener("change", syncFilterUI);
  filterDateMode?.addEventListener("change", syncFilterUI);

  function readFilterFromUI() {
    const cat = filterCategory.value;
    if (cat === "date") {
      const mode = filterDateMode.value;
      const m = Number(filterMonth.value);
      const y = Number(filterYear.value);
      const month = Number.isFinite(m) && m >= 1 && m <= 12 ? m : null;
      const year = Number.isFinite(y) && y >= 1900 && y <= 3000 ? y : null;
      return { category: "date", dateMode: mode, month, year };
    }
    if (cat === "amount") {
      const min = filterMin.value === "" ? null : parseMoney(filterMin.value);
      const max = filterMax.value === "" ? null : parseMoney(filterMax.value);
      return { category: "amount", min, max };
    }
    const txt = (filterText.value || "").trim();
    return { category: cat, text: txt };
  }

  function saveFilter(filter) {
    localStorage.setItem(FILTER_KEY, JSON.stringify(filter || null));
  }
  function loadFilter() {
    try { return JSON.parse(localStorage.getItem(FILTER_KEY) || "null"); } catch { return null; }
  }
  function setUIFromFilter(filter) {
    if (!filter) return;
    filterCategory.value = filter.category || "date";
    if (filter.category === "date") {
      filterDateMode.value = filter.dateMode || "monthYear";
      filterMonth.value = filter.month ? String(filter.month) : "";
      filterYear.value = filter.year ? String(filter.year) : "";
    } else if (filter.category === "amount") {
      filterMin.value = filter.min != null ? String(filter.min) : "";
      filterMax.value = filter.max != null ? String(filter.max) : "";
    } else {
      filterText.value = filter.text || "";
    }
    syncFilterUI();
  }

  btnApplyFilter?.addEventListener("click", () => {
    activeFilter = readFilterFromUI();
    saveFilter(activeFilter);
    renderContracts();
  });

  btnClearFilter?.addEventListener("click", () => {
    activeFilter = null;
    saveFilter(null);
    // reset UI
    filterCategory.value = "date";
    filterDateMode.value = "monthYear";
    filterMonth.value = "";
    filterYear.value = "";
    filterText.value = "";
    filterMin.value = "";
    filterMax.value = "";
    syncFilterUI();
    renderContracts();
  });

  // ---------------------------
  // Export (CSV) - exports only data rows in current view or full?
  // We'll export current visible view (after filter), like what you see.
  // ---------------------------
  function escapeCSV(value) {
    const s = String(value ?? "");
    if (/[",\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
    return s;
  }
  function exportVisibleToCSV() {
    const headers = Array.from(table.tHead.rows[0].cells).map(th => th.innerText.trim());
    const rows = $$('.grid tbody tr[data-row-type="data"]').map(tr => Array.from(tr.cells).map(td => td.innerText.trim()));
    return [headers, ...rows].map(r => r.map(escapeCSV).join(",")).join("\n");
  }
  function download(filename, text) {
    const blob = new Blob([text], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }
  btnExport?.addEventListener("click", () => {
    // Export a styled Excel-like sheet (HTML .xls) that matches the grouped layout:
    // - Month header rows
    // - TOTAL rows per month
    // - GRAND TOTAL row at the bottom
    //
    // This approach preserves colors/merged cells in Excel without requiring a paid XLSX styling library.
    const visible = applyFilterToRows(contractsStore.slice(), activeFilter);

    // sort same as render
    visible.sort((a, b) => {
      const ak = monthKeyFromDate(a.date);
      const bk = monthKeyFromDate(b.date);
      if (ak !== bk) {
        if (ak === "unknown") return 1;
        if (bk === "unknown") return -1;
        return ak.localeCompare(bk);
      }
      const ad = parseMMDDYYYY(a.date)?.getTime() ?? Number.POSITIVE_INFINITY;
      const bd = parseMMDDYYYY(b.date)?.getTime() ?? Number.POSITIVE_INFINITY;
      if (ad !== bd) return ad - bd;
      return (a.contract || "").localeCompare(b.contract || "");
    });

    // group
    const groups = new Map();
    for (const r of visible) {
      const k = monthKeyFromDate(r.date);
      if (!groups.has(k)) groups.set(k, []);
      groups.get(k).push(r);
    }
    const keys = Array.from(groups.keys()).sort((a, b) => {
      if (a === "unknown" && b !== "unknown") return 1;
      if (b === "unknown" && a !== "unknown") return -1;
      return a.localeCompare(b);
    });

    const cols = [
      "Date","Contract Number","Deceased","Casket","Address",
      "Contract Amount","INHAUS","BAI","GL","GCASH","CASH","Discount","Total Paid","Last Payment Date","Remaining"
    ];

    // Excel-friendly number format
    const msoMoney = 'mso-number-format:"\\#\\,\\#\\#0\\.00"';
    const msoText = 'mso-number-format:"\\@"';

    const css = `
      <style>
        table { border-collapse: collapse; font-family: Calibri, Arial, sans-serif; font-size: 11pt; }
        th, td { border: 1px solid #cfcfcf; padding: 4px 6px; white-space: nowrap; }
        th { font-weight: 700; background: #ffffff; }
        td.num { text-align: right; }
        tr.data-row td { background: #ffffff; }
        tr.month-total td { font-weight: 700; background: #fff2cc; border-top: 2px solid #000000; }
        tr.month-total td.label { text-align: center; letter-spacing: .10em; }
        tr.month-header td { font-weight: 800; text-align: center; background: #d9d9d9; border-top: 2px solid #000000; border-bottom: 2px solid #000000; }
        tr.spacer td { border: none; height: 10px; background: #ffffff; }
        tr.grand-total td { font-weight: 800; background: #92d050; border-top: 2px solid #000000; border-bottom: 2px solid #000000; }
        tr.grand-total td.label { text-align: center; letter-spacing: .10em; }
      </style>
    `;

    // Column widths similar to screenshot
    const colgroup = `
      <colgroup>
        <col style="width:90px" />
        <col style="width:110px" />
        <col style="width:190px" />
        <col style="width:110px" />
        <col style="width:140px" />
        <col style="width:120px" />
        <col style="width:90px" />
        <col style="width:80px" />
        <col style="width:80px" />
        <col style="width:95px" />
        <col style="width:90px" />
        <col style="width:90px" />
        <col style="width:110px" />
        <col style="width:140px" />
        <col style="width:110px" />
      </colgroup>
    `;

    const esc = (s) => String(s ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

    let html = `<!doctype html><html><head><meta charset="utf-8" />${css}</head><body>`;
    html += `<table>${colgroup}<thead><tr>`;
    for (const h of cols) html += `<th>${esc(h)}</th>`;
    html += `</tr></thead><tbody>`;

    const grand = { amount:0, inhaus:0, bai:0, gl:0, gcash:0, cash:0, discount:0, totalPaid:0, remaining:0 };
    let any = false;

    for (let gi = 0; gi < keys.length; gi++) {
      const key = keys[gi];
      const rows = groups.get(key) || [];
      if (!rows.length) continue;
      any = true;

      // Month header row (merged across all columns)
      html += `<tr class="month-header"><td colspan="15" style="${msoText}; font-size: 12pt;">${esc(monthLabelFromKey(key).toUpperCase())}</td></tr>`;

// Repeat column headers for this month group
html += `<tr>`;
for (const h of cols) {
  html += `<th>${esc(h)}</th>`;
}
html += `</tr>`;

      const totals = { amount:0, inhaus:0, bai:0, gl:0, gcash:0, cash:0, discount:0, totalPaid:0, remaining:0 };

      for (const c of rows) {
        const computed = calcComputed(c);
        totals.amount += c.amount || 0;
        totals.inhaus += c.inhaus || 0;
        totals.bai += c.bai || 0;
        totals.gl += c.gl || 0;
        totals.gcash += c.gcash || 0;
        totals.cash += c.cash || 0;
        totals.discount += c.discount || 0;
        totals.totalPaid += computed.totalPaid;
        totals.remaining += computed.remaining;

        html += `<tr class="data-row">`;
        html += `<td style="${msoText}">${esc(c.date)}</td>`;
        html += `<td style="${msoText}">${esc(c.contract)}</td>`;
        html += `<td style="${msoText}">${esc(c.deceased)}</td>`;
        html += `<td style="${msoText}">${esc(c.casket)}</td>`;
        html += `<td style="${msoText}">${esc(c.address)}</td>`;

        const numCell = (v) => `<td class="num" style="${msoMoney}">${fmtMoney(v)}</td>`;
        html += numCell(c.amount);
        html += numCell(c.inhaus);
        html += numCell(c.bai);
        html += numCell(c.gl);
        html += numCell(c.gcash);
        html += numCell(c.cash);
        html += numCell(c.discount);
        html += numCell(computed.totalPaid);
        html += `<td style="${msoText}">${esc(c.lastPayment || "")}</td>`;
        html += numCell(computed.remaining);
        html += `</tr>`;
      }

      // TOTAL row: label under Address column
      html += `<tr class="month-total">`;
      html += `<td colspan="4"></td>`;
      html += `<td class="label">TOTAL</td>`;
      const numCell = (v) => `<td class="num" style="${msoMoney}">${fmtMoney(v)}</td>`;
      html += numCell(totals.amount);
      html += numCell(totals.inhaus);
      html += numCell(totals.bai);
      html += numCell(totals.gl);
      html += numCell(totals.gcash);
      html += numCell(totals.cash);
      html += numCell(totals.discount);
      html += numCell(totals.totalPaid);
      html += `<td></td>`;
      html += numCell(totals.remaining);
      html += `</tr>`;

      // add to grand
      grand.amount += totals.amount;
      grand.inhaus += totals.inhaus;
      grand.bai += totals.bai;
      grand.gl += totals.gl;
      grand.gcash += totals.gcash;
      grand.cash += totals.cash;
      grand.discount += totals.discount;
      grand.totalPaid += totals.totalPaid;
      grand.remaining += totals.remaining;

      // Extra spacer line between months
      if (gi < keys.length - 1) {
        html += `<tr class="spacer"><td colspan="15"></td></tr>`;
      }
    }

    if (any) {
      // Spacer + GRAND TOTAL row (green)
      html += `<tr class="spacer"><td colspan="15"></td></tr>`;
      html += `<tr class="grand-total">`;
      html += `<td colspan="4"></td>`;
      html += `<td class="label">GRAND TOTAL</td>`;
      const numCell = (v) => `<td class="num" style="${msoMoney}">${fmtMoney(v)}</td>`;
      html += numCell(grand.amount);
      html += numCell(grand.inhaus);
      html += numCell(grand.bai);
      html += numCell(grand.gl);
      html += numCell(grand.gcash);
      html += numCell(grand.cash);
      html += numCell(grand.discount);
      html += numCell(grand.totalPaid);
      html += `<td></td>`;
      html += numCell(grand.remaining);
      html += `</tr>`;
    }

    html += `</tbody></table></body></html>`;

    const blob = new Blob([html], { type: "application/vnd.ms-excel;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `Magallanes_Contracts_${new Date().toISOString().slice(0,10)}.xls`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  });

  
  // ---------------------------
  // Save/Load Data (JSON file backup)
  // NOTE: Browsers cannot automatically write files into the project folder.
  // "Save Data" downloads a JSON file; you can keep it in the same folder as index.html.
  // "Load Data" lets you pick that JSON file to restore.
  // ---------------------------
  function exportStoreToJSON() {
    // Keep it simple: just the contracts array
    return JSON.stringify({ version: 1, exportedAt: new Date().toISOString(), contracts: contractsStore }, null, 2);
  }

  function importStoreFromJSON(text) {
    const parsed = JSON.parse(text);
    const arr = Array.isArray(parsed) ? parsed : parsed?.contracts;
    if (!Array.isArray(arr)) throw new Error("Invalid JSON format. Expected an array or { contracts: [...] }");
    // normalize
    return arr.map(x => ({
      date: String(x.date || "").trim(),
      contract: String(x.contract || "").trim(),
      deceased: String(x.deceased || "").trim(),
      casket: String(x.casket || "").trim(),
      address: String(x.address || "").trim(),
      amount: Number(x.amount) || 0,
      inhaus: Number(x.inhaus) || 0,
      bai: Number(x.bai) || 0,
      gl: Number(x.gl) || 0,
      gcash: Number(x.gcash) || 0,
      cash: Number(x.cash) || 0,
      discount: Number(x.discount) || 0,
      lastPayment: String(x.lastPayment || "—").trim() || "—"
    })).filter(r => r.contract || r.deceased || r.date);
  }

  btnSaveData?.addEventListener("click", () => {
    const jsonText = exportStoreToJSON();
    download("contracts_data.json", jsonText);
    alert("Saved! A file named contracts_data.json was downloaded.\n\nTip: Move it into the same folder as index.html for easy access.");
  });

  btnLoadData?.addEventListener("click", () => fileLoadJson?.click());

  fileLoadJson?.addEventListener("change", async () => {
    const file = fileLoadJson.files?.[0];
    if (!file) return;
    try {
      const text = await file.text();
      const imported = importStoreFromJSON(text);
      if (!confirm(`Load this file and replace your current data?\n\nRows in file: ${imported.length}`)) return;
      contractsStore = imported;
      // Bulk save all to Supabase
      Promise.all(contractsStore.map(r => DB.saveContract(r))).then(() => {
        renderContracts();
        alert("Loaded! Data restored from JSON file.");
      });
    } catch (err) {
      alert("Load failed: " + (err?.message || String(err)));
    } finally {
      fileLoadJson.value = "";
    }
  });


// ---------------------------
  // Import (CSV + XLSX) into store
  // ---------------------------
  btnImport?.addEventListener("click", () => fileImport.click());

  function normHeader(h) {
    return String(h ?? "")
      .trim()
      .toLowerCase()
      .replace(/\s+/g, " ")
      .replace(/[#]/g, "#")
      .replace(/[().,:;\-]/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

  const HEADER_ALIASES = {
    date: ["date", "contract date", "date of contract"],
    contract: ["contract #", "contract#", "contract no", "contract number", "contract", "contract id", "contract num", "contract no."],
    deceased: ["deceased", "name", "deceased name", "client", "customer", "customer name"],
    casket: ["casket", "casket type", "casket model"],
    address: ["address", "addr", "location"],
    amount: ["amount", "contract amount", "total amount", "contract price", "price"],
    inhaus: ["inhaus", "in house", "inhouse"],
    bai: ["bai"],
    gl: ["gl"],
    gcash: ["gcash", "g cash", "g-cash"],
    cash: ["cash"],
    discount: ["discount", "disc"],
    lastPayment: ["last payment", "lastpay", "last paid", "last payment date"]
  };

  function buildHeaderIndex(headers) {
    const map = {};
    const norms = headers.map(normHeader);
    function findKey(key) {
      const aliases = HEADER_ALIASES[key] || [];
      for (const a of aliases) {
        const ai = norms.indexOf(normHeader(a));
        if (ai >= 0) return ai;
      }
      return -1;
    }
    for (const key of Object.keys(HEADER_ALIASES)) {
      const ix = findKey(key);
      if (ix >= 0) map[key] = ix;
    }
    return { map, norms, raw: headers };
  }

  function excelSerialToDate(serial) {
    const n = Number(serial);
    if (!Number.isFinite(n)) return null;
    const utc_days = Math.floor(n - 25569);
    const utc_value = utc_days * 86400;
    const date_info = new Date(utc_value * 1000);
    return date_info;
  }

  function parseDateAny(v) {
    if (v == null || v === "") return "";
    if (v instanceof Date) return formatMMDDYYYY(v);
    if (typeof v === "number") {
      const d = excelSerialToDate(v);
      return d ? formatMMDDYYYY(d) : "";
    }
    const s = String(v).trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return yyyyMMddToMMDDYYYY(s);
    if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(s)) {
      const parts = s.split("/");
      return `${pad2(parts[0])}/${pad2(parts[1])}/${parts[2]}`;
    }
    const d = new Date(s);
    if (Number.isFinite(d.getTime())) return formatMMDDYYYY(d);
    return s;
  }

  function getCell(row, idx) {
    return (idx >= 0 && idx < row.length) ? row[idx] : "";
  }

  function rowToContract(row, headerMap) {
    const ix = (k) => (k in headerMap ? headerMap[k] : -1);
    // Only import the 6 core columns; payment columns are always zeroed out
    // and will be computed as transactions are added.
    const date = parseDateAny(getCell(row, ix("date")));
    const contract = String(getCell(row, ix("contract")) ?? "").trim();
    const deceased = String(getCell(row, ix("deceased")) ?? "").trim();
    const casket = String(getCell(row, ix("casket")) ?? "").trim();
    const address = String(getCell(row, ix("address")) ?? "").trim();
    const amount = parseMoney(getCell(row, ix("amount")));
    return { date, contract, deceased, casket, address, amount, inhaus: 0, bai: 0, gl: 0, gcash: 0, cash: 0, discount: 0, lastPayment: "—" };
  }

  function detectHeaderRow(rows) {
    for (let r = 0; r < Math.min(rows.length, 5); r++) {
      const row = rows[r];
      const headers = row.map(x => String(x ?? "").trim()).filter(Boolean);
      if (headers.length < 2) continue;
      const { map } = buildHeaderIndex(headers);
      const hits = Object.keys(map).length;
      if (hits >= 2) return r;
    }
    return -1;
  }

  function summarizeImport(result) {
    const lines = [];
    lines.push(`Imported: ${result.added}`);
    if (result.updated) lines.push(`Updated: ${result.updated}`);
    if (result.skipped) lines.push(`Skipped: ${result.skipped}`);
    if (result.errors.length) {
      lines.push("");
      lines.push("Errors (first 8):");
      result.errors.slice(0,8).forEach(e => lines.push(`- Row ${e.row}: ${e.message}`));
    }
    alert(lines.join("\n"));
  }

  function findIndexByContract(contractNo) {
    const key = normalizeText(contractNo);
    return contractsStore.findIndex(x => normalizeText(x.contract) === key);
  }

  async function importRows(rows2d) {
    const headerRowIndex = detectHeaderRow(rows2d);
    let headers = null;
    let dataRows = null;

    if (headerRowIndex >= 0) {
      headers = rows2d[headerRowIndex].map(x => String(x ?? "").trim());
      dataRows = rows2d.slice(headerRowIndex + 1);
    } else {
      headers = ["Date","Contract #","Deceased","Casket","Address","Amount","INHAUS","BAI","GL","GCASH","CASH","Discount","Total Paid","Last Payment","Remaining"];
      dataRows = rows2d;
    }

    dataRows = dataRows.filter(r => r && r.some(v => String(v ?? "").trim() !== ""));
    if (dataRows.length === 0) { alert("No data rows found in the selected file."); return; }

    const { map: headerMap } = buildHeaderIndex(headers);

    // Require Contract and Deceased when header exists
    const missing = [];
    if (headerRowIndex >= 0) {
      if (!("contract" in headerMap)) missing.push("Contract #");
      if (!("deceased" in headerMap)) missing.push("Deceased");
      if (missing.length) {
        alert(`Missing required column(s): ${missing.join(", ")}\n\nDetected headers:\n- ${headers.filter(Boolean).join("\n- ")}`);
        return;
      }
    }

    const replace = confirm("Replace current data with imported data?\nOK = Replace\nCancel = Append/Update");
    if (replace) {
      contractsStore = [];
      selectedContractNo = null;
      saveStore();
    }

    let dupMode = "skip";
    if (!replace) {
      const overwrite = confirm("When a Contract # already exists:\nOK = Update existing\nCancel = Skip duplicates");
      dupMode = overwrite ? "update" : "skip";
    }

    const result = { added: 0, updated: 0, skipped: 0, errors: [] };

    dataRows.forEach((row, i) => {
      try {
        const c = rowToContract(row, headerMap);
        if (!c.contract && !c.deceased) { result.skipped++; return; }
        if (headerRowIndex >= 0 && !c.contract) { result.skipped++; return; }

        const idx = c.contract ? findIndexByContract(c.contract) : -1;
        if (!replace && idx >= 0) {
          if (dupMode === "update") {
            const existing = contractsStore[idx];
            contractsStore[idx] = {
              ...existing,
              date: c.date || existing.date,
              contract: c.contract || existing.contract,
              deceased: c.deceased || existing.deceased,
              casket: c.casket || existing.casket,
              address: c.address || existing.address,
              amount: Number.isFinite(c.amount) ? c.amount : existing.amount,
              // Payment columns are intentionally preserved from existing data
              // and never overwritten by import — they come from transactions.
            };
            result.updated++;
          } else {
            result.skipped++;
          }
          return;
        }

        contractsStore.push({
          date: c.date || "",
          contract: c.contract || "",
          deceased: c.deceased || "",
          casket: c.casket || "",
          address: c.address || "",
          amount: c.amount || 0,
          inhaus: c.inhaus || 0,
          bai: c.bai || 0,
          gl: c.gl || 0,
          gcash: c.gcash || 0,
          cash: c.cash || 0,
          discount: c.discount || 0,
          lastPayment: c.lastPayment || "—"
        });
        result.added++;
      } catch (err) {
        result.errors.push({ row: i + 1, message: err?.message || String(err) });
      }
    });

    // Bulk save to Supabase
    Promise.all(contractsStore.map(r => DB.saveContract(r)));
    renderContracts();
    summarizeImport(result);
  }

  function parseCSVTo2D(text) {
    const rows = [];
    let row = [];
    let i = 0, field = "", inQuotes = false;
    while (i < text.length) {
      const c = text[i];
      if (inQuotes) {
        if (c === '"') {
          if (text[i + 1] === '"') { field += '"'; i += 2; continue; }
          inQuotes = false; i++; continue;
        }
        field += c; i++; continue;
      } else {
        if (c === '"') { inQuotes = true; i++; continue; }
        if (c === ",") { row.push(field); field = ""; i++; continue; }
        if (c === "\n") { row.push(field); rows.push(row); row = []; field = ""; i++; continue; }
        if (c === "\r") { i++; continue; }
        field += c; i++; continue;
      }
    }
    row.push(field);
    rows.push(row);
    return rows;
  }

  async function readXLSX(file) {
    if (typeof XLSX === "undefined") {
      alert("XLSX library not loaded. If you're offline, import a CSV instead.");
      return null;
    }
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array", cellDates: true });
    const sheetName = wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    return XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  }

  fileImport?.addEventListener("change", async () => {
    const file = fileImport.files?.[0];
    if (!file) return;
    try {
      const name = (file.name || "").toLowerCase();
      if (name.endsWith(".csv")) {
        const text = await file.text();
        await importRows(parseCSVTo2D(text));
      } else if (name.endsWith(".xlsx")) {
        const rows = await readXLSX(file);
        if (!rows) return;
        await importRows(rows);
      } else {
        alert("Unsupported file type. Please select a .csv or .xlsx file.");
      }
    } catch (err) {
      alert("Import failed: " + (err?.message || String(err)));
    } finally {
      fileImport.value = "";
    }
  });

  // Refresh (re-render + keep filter)
  btnRefresh?.addEventListener("click", () => {
    renderContracts();
    alert("Refreshed view.");
  });

  // ---------------------------
  // Initialize Store from existing sample row in HTML
  // ---------------------------
  function initStoreFromDOM() {
    const trs = Array.from(table.tBodies[0].rows).filter(r => (r.dataset.rowType || "data") === "data");
    contractsStore = trs.map(tr => {
      const get = (col) => tr.querySelector(`[data-col="${col}"]`)?.innerText ?? "";
      return {
        date: (tr.cells[0]?.innerText || "").trim(),
        contract: (tr.cells[1]?.innerText || "").trim(),
        deceased: (tr.cells[2]?.innerText || "").trim(),
        casket: (tr.cells[3]?.innerText || "").trim(),
        address: (tr.cells[4]?.innerText || "").trim(),
        amount: parseMoney(get("amount")),
        inhaus: parseMoney(get("inhaus")),
        bai: parseMoney(get("bai")),
        gl: parseMoney(get("gl")),
        gcash: parseMoney(get("gcash")),
        cash: parseMoney(get("cash")),
        discount: parseMoney(get("discount")),
        lastPayment: String(get("lastPayment") || "—").trim() || "—"
      };
    });
  }

  // Restore filter
  syncFilterUI();
  const restored = loadFilter();
  if (restored) {
    activeFilter = restored;
    setUIFromFilter(restored);
  }

  // Load contracts from Supabase
  Promise.all([
    DB.getContracts(),
    DB.getDswd(),
    DB.getBai(),
  ]).then(([contracts, dswd, bai]) => {
    contractsStore = contracts;
    dswdStore      = dswd;
    baiStore       = bai;
    // Sync assisted-payment amounts into contracts before first render
    // (guards reset in case prior run set them)
    _syncingDswd = false;
    _syncingBai  = false;
    // Write dswdAfterTax, dswdDiscount, baiAssist directly onto contractsStore
    for (const c of contractsStore) {
      const key = normalizeText(c.contract || "");
      // DSWD
      let dswdAfterTax = 0, dswdDiscount = 0;
      for (const r of dswdStore) {
        if (normalizeText(r.contract || "") === key) {
          dswdAfterTax += Number(r.afterTax)    || 0;
          dswdDiscount += Number(r.dswdDiscount) || 0;
        }
      }
      c.dswdAfterTax = dswdAfterTax;
      c.dswdDiscount = dswdDiscount;
      // BAI
      let baiAssist = 0;
      for (const r of baiStore) {
        if (normalizeText(r.contract || "") === key) {
          baiAssist += Number(r.amount) || 0;
        }
      }
      c.baiAssist = baiAssist;
    }
    renderContracts();
    // Render sub-tab tables (defined later in file, use setTimeout to ensure they exist)
    setTimeout(() => {
      if (typeof renderDswdTable === "function") renderDswdTable();
      if (typeof renderBaiTable  === "function") renderBaiTable();
    }, 0);
  });

  // ---------------------------
  // ---------------------------
  // ---------------------------
  // Cash Received (Transactions) — full functionality + monthly grouping (v28)
  // ---------------------------
  const cashTable = $("#cashReceivedTable");
  const cashRowCountEl = $("#cashReceivedRowCount");
  const cashSelectedEl = $("#cashReceivedSelected");

  const btnCashAdd = $("#btnCashAddEntry");
  const btnCashEdit = $("#btnCashEditSelected");
  const btnCashDel = $("#btnCashDeleteSelected");
  const btnCashImport = $("#btnCashImportExcel");
  const btnCashExport = $("#btnCashExportExcel");
  const btnCashRefresh = $("#btnCashRefresh");
  const fileCashImport = $("#fileCashImport");

  // Modal elements
  const cashOverlay = $("#cashOverlay");
  const cashModal = $("#cashModal");
  const cashTitle = $("#cashTitle");
  const cashSubtitle = $("#cashSubtitle");
  const cashForm = $("#cashForm");
  const btnCloseCash = $("#btnCloseCash");
  const btnCancelCash = $("#btnCancelCash");
  const btnSubmitCash = $("#btnSubmitCash");

  const crDate = $("#crDate");
  const crContract = $("#crContract");
  const crReceipt = $("#crReceipt");
  const crClient = $("#crClient");
  const crParticular = $("#crParticular");
  const crAmount = $("#crAmount");

  const CASH_STORE_KEY = "mf_cash_received_store_v28";

  /** @type {Array<{date:string, contract:string, receipt:string, client:string, particular:string, amount:number, _id?:string}>} */
  let cashStore = [];

  let cashSelectedKey = null; // internal key (receipt or _id)
  let cashMode = "add";
  let cashEditingKey = null;

  function cashKeyFor(r) {
    return normalizeText(r.receipt || r._id || "");
  }

  function ensureCashId(r) {
    if (!r._id) r._id = "__AUTO__" + Date.now() + "_" + Math.floor(Math.random()*100000);
    return r;
  }

  function saveCashStore() { /* no-op: saves happen per-row via DB */ }
  function loadCashStore() { return null; /* loaded async */ }

  function cashYyyyMMddToMMDDYYYY(yyyyMMdd) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(yyyyMMdd || ""));
    return m ? `${m[2]}/${m[3]}/${m[1]}` : "";
  }
  function cashMMDDYYYYToYyyyMMdd(mmddyyyy) {
    const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(String(mmddyyyy || "").trim());
    if (!m) return "";
    const mm = String(m[1]).padStart(2,"0");
    const dd = String(m[2]).padStart(2,"0");
    return `${m[3]}-${mm}-${dd}`;
  }
  function canonicalizeParticularText(txt) {
      const raw = String(txt || "").trim();
      const p = normalizeText(raw);
      if (!p) return "";
      if (/g[\s\-]?cash/.test(p) || p.includes("gcash")) return "GCASH";
      if (/in[\s\-]?house/.test(p) || p.includes("inhaus") || p.includes("inhouse")) return "INHAUS";
      if (/\bbai\b/.test(p)) return "BAI";
      if (/\bbank\b/.test(p) || p.includes("deposit bank") || p.includes("bank received")) return "BANK";
      if (/\bg\/l\b/.test(p) || /\bg\s*l\b/.test(p) || /\bgl\b/.test(p)) return "GL";
      if (/\bcash\b/.test(p)) return "CASH";
      return raw;
    }

  function parseMoney2(s) {
    const n = Number(String(s ?? "").replace(/[^0-9.-]/g, ""));
    return Number.isFinite(n) ? n : 0;
  }

  function openCashModal(mode, keyOrNull = null) {
    if (!cashOverlay || !cashModal) {
      alert("Cash Entry form is not available (modal not found in page).");
      return;
    }
    cashMode = mode;
    cashEditingKey = null;

    if (cashMode === "add") {
      cashTitle.textContent = "Add Entry";
      cashSubtitle.textContent = "Enter cash receipt details.";
      btnSubmitCash.textContent = "Add Entry";

      crDate.value = new Date().toISOString().slice(0,10);
      crContract.value = "";
      crReceipt.value = "";
      crClient.value = "";
      crParticular.value = "";
      crAmount.value = "0.00";
    } else {
      cashTitle.textContent = "Edit Entry";
      cashSubtitle.textContent = "Update the selected cash receipt entry.";
      btnSubmitCash.textContent = "Save Changes";

      const key = normalizeText(keyOrNull);
      const found = cashStore.find(x => cashKeyFor(x) === key);
      if (!found) { alert("Could not find the selected entry."); return; }
      cashEditingKey = key;

      crDate.value = cashMMDDYYYYToYyyyMMdd(found.date) || new Date().toISOString().slice(0,10);
      crContract.value = found.contract || "";
      crReceipt.value = found.receipt || "";
      crClient.value = found.client || "";
      crParticular.value = found.particular || "";
      crAmount.value = (Number(found.amount) || 0).toFixed(2);
    }

    cashOverlay.classList.add("is-open");
    cashModal.classList.add("is-open");
    setTimeout(() => crClient.focus(), 0);
  }

  function closeCashModal() {
    cashOverlay?.classList.remove("is-open");
    cashModal?.classList.remove("is-open");
    cashMode = "add";
    cashEditingKey = null;
  }

  btnCloseCash?.addEventListener("click", closeCashModal);
  btnCancelCash?.addEventListener("click", closeCashModal);
  cashOverlay?.addEventListener("click", closeCashModal);
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && cashModal?.classList.contains("is-open")) closeCashModal();
  });

  // ---------- Linking Cash Received -> Contracts (CASH column) ----------
  // We compute CASH for each contract based on Cash Received entries with matching Contract #.
  // This will overwrite the Contracts tab CASH column for those contracts (automatic sync).
  function syncCashReceivedToContracts() {
    if (!Array.isArray(contractsStore)) return;

    // contractNo -> categorized sums
    const sumsByContract = new Map(); // key -> { inhaus,bai,gl,gcash,cash }
    const latestByContract = new Map(); // key -> latest date string (MM/DD/YYYY)

    function bucketFor(particular) {
      const raw = String(particular || "").trim();
      const p = normalizeText(raw);

      if (!p) return "cash"; // default bucket when blank

      // GCASH: "g-cash", "g cash", etc.
      if (/g[\s\-]?cash/.test(p) || p.includes("gcash")) return "gcash";

      // INHAUS: "in house", "in-house", "inhouse"
      if (/in[\s\-]?house/.test(p) || p.includes("inhaus") || p.includes("inhouse")) return "inhaus";

      // BAI/BANK: allow "bank", "deposit bank", etc.
      if (/\bbai\b/.test(p) || /\bbank\b/.test(p) || p.includes("bank received") || p.includes("deposit bank")) return "bai";

      // GL: "g/l", "g l", "gl"
      if (/\bgl\b/.test(p) || /\bg\/l\b/.test(p) || /\bg\s*l\b/.test(p)) return "gl";

      // CASH: "cash", "paid cash"
      if (/\bcash\b/.test(p)) return "cash";

      return "cash";
    }

    for (const r of cashStore) {
      const cno = normalizeText(r.contract);
      if (!cno) continue;

      if (!sumsByContract.has(cno)) {
        sumsByContract.set(cno, { inhaus: 0, bai: 0, gl: 0, gcash: 0, cash: 0 });
      }
      const sums = sumsByContract.get(cno);
      const bucket = bucketFor(r.particular);
      sums[bucket] += Number(r.amount) || 0;

      const d = parseMMDDYYYY(r.date);
      if (!d) continue;
      const prev = latestByContract.get(cno);
      const prevD = prev ? parseMMDDYYYY(prev) : null;
      if (!prevD || d.getTime() > prevD.getTime()) latestByContract.set(cno, r.date);
    }

    let changed = false;
    for (const c of contractsStore) {
      const key = normalizeText(c.contract);
      if (!key) continue;
      const sums = sumsByContract.get(key);
      if (!sums) continue;

      // Overwrite these payment columns based on Cash Received ledger for this contract.
      // This keeps Contracts consistent with the Transactions ledger.
      const fields = ["inhaus","bai","gl","gcash","cash"];
      for (const f of fields) {
        const nv = Number(sums[f]) || 0;
        if ((Number(c[f]) || 0) !== nv) { c[f] = nv; changed = true; }
      }

      const latest = latestByContract.get(key);
      if (latest) {
        const cur = parseMMDDYYYY(String(c.lastPayment || "").trim());
        const lat = parseMMDDYYYY(latest);
        if (!cur || (lat && lat.getTime() > cur.getTime())) {
          c.lastPayment = latest;
          changed = true;
        }
      }
    }

    if (changed) {
      // Save updated contracts to Supabase
      contractsStore.forEach(c => { if (c.id) DB.saveContract(c); });
      renderContracts();
    }
  }

  // ---------- Render Cash table with Monthly Grouping ----------
  function cashGroupColumnsRow() {
    const tr = document.createElement("tr");
    tr.classList.add("cashGroupCols");
    tr.dataset.rowType = "cashGroupCols";
    const labels = ["Date","Contract #","Receipt #","Name of Client","Particular","Amount"];
    for (const lab of labels) {
      const td = document.createElement("td");
      td.textContent = lab;
      tr.appendChild(td);
    }
    return tr;
  }

  function cashMonthHeaderRow(label) {
    const tr = document.createElement("tr");
    tr.classList.add("cashMonthHeader");
    tr.dataset.rowType = "cashMonthHeader";
    const td = document.createElement("td");
    td.colSpan = 6;
    td.textContent = String(label || "").toUpperCase();
    tr.appendChild(td);
    return tr;
  }

  function cashMonthTotalRow(total) {
    const tr = document.createElement("tr");
    tr.classList.add("cashMonthTotal");
    tr.dataset.rowType = "cashMonthTotal";

    // Date, Contract, Receipt, Client empty; label under Particular; amount under Amount
    for (let i=0;i<4;i++){ tr.appendChild(document.createElement("td")); }
    const label = document.createElement("td");
    label.textContent = "TOTAL";
    label.classList.add("label");
    tr.appendChild(label);

    const amt = document.createElement("td");
    amt.classList.add("num");
    amt.textContent = fmtMoney(total);
    tr.appendChild(amt);
    return tr;
  }

  function cashGrandTotalRow(total) {
    const tr = document.createElement("tr");
    tr.classList.add("cashGrandTotal");
    tr.dataset.rowType = "cashGrandTotal";

    for (let i=0;i<4;i++){ tr.appendChild(document.createElement("td")); }
    const label = document.createElement("td");
    label.textContent = "GRAND TOTAL";
    label.classList.add("label");
    tr.appendChild(label);

    const amt = document.createElement("td");
    amt.classList.add("num");
    amt.textContent = fmtMoney(total);
    tr.appendChild(amt);
    return tr;
  }

  function cashSpacerRow() {
    const tr = document.createElement("tr");
    tr.classList.add("spacer");
    tr.dataset.rowType = "spacer";
    const td = document.createElement("td");
    td.colSpan = 6;
    td.textContent = "";
    tr.appendChild(td);
    return tr;
  }


  // ---------- Cash Received Filters ----------
  const cashFilterCategory = $("#cashFilterCategory");
  const cashFilterInputs = $("#cashFilterInputs");
  const btnCashApplyFilter = $("#btnCashApplyFilter");
  const btnCashClearFilter = $("#btnCashClearFilter");

  let cashActiveFilter = { category: "", value: null };

  function setCashFilterInputs(category) {
    if (!cashFilterInputs) return;
    cashFilterInputs.innerHTML = "";

    const makeText = (id, label, placeholder="") => {
      const wrap = document.createElement("label");
      wrap.className = "field inline";
      wrap.innerHTML = `<span>${label}</span><input class="input" id="${id}" type="text" placeholder="${placeholder}"/>`;
      cashFilterInputs.appendChild(wrap);
    };

    const makeNumber = (id, label, placeholder="") => {
      const wrap = document.createElement("label");
      wrap.className = "field inline";
      wrap.innerHTML = `<span>${label}</span><input class="input" id="${id}" type="number" step="0.01" placeholder="${placeholder}"/>`;
      cashFilterInputs.appendChild(wrap);
    };

    if (!category) return;

    if (category === "date") {
      makeText("cashFilterDate", "Date", "MM/DD/YYYY or MM/YYYY or YYYY");
      return;
    }
    if (category === "contract") {
      makeText("cashFilterContract", "Contract #", "e.g., 13566");
      return;
    }
    if (category === "receipt") {
      makeText("cashFilterReceipt", "Receipt #", "e.g., R-0001 or (blank)");
      return;
    }
    if (category === "client") {
      makeText("cashFilterClient", "Client", "contains...");
      return;
    }
    if (category === "particular") {
      makeText("cashFilterParticular", "Particular", "contains...");
      return;
    }
    if (category === "amount") {
      makeNumber("cashFilterMin", "Min", "0.00");
      makeNumber("cashFilterMax", "Max", "10000.00");
      return;
    }
  }

  cashFilterCategory?.addEventListener("change", () => setCashFilterInputs(cashFilterCategory.value));

  function parseFilterDatePattern(s) {
    const t = String(s || "").trim();
    if (!t) return null;
    if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(t)) {
      const d = parseMMDDYYYY(t);
      return d ? { type: "full", d: formatMMDDYYYY(d) } : null;
    }
    if (/^\d{1,2}\/\d{4}$/.test(t)) {
      const [mm, yyyy] = t.split("/");
      return { type: "monthYear", mm: pad2(mm), yyyy: yyyy };
    }
    if (/^\d{4}$/.test(t)) {
      return { type: "year", yyyy: t };
    }
    return { type: "contains", text: normalizeText(t) };
  }

  function rowMatchesDateFilter(rowDate, pattern) {
    if (!pattern) return true;
    const rd = String(rowDate || "").trim();
    if (!rd) return false;

    if (pattern.type === "full") return rd === pattern.d;
    if (pattern.type === "monthYear") {
      const m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(rd);
      if (!m) return false;
      return m[1] === pattern.mm && m[3] === pattern.yyyy;
    }
    if (pattern.type === "year") {
      const m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(rd);
      return m ? m[3] === pattern.yyyy : false;
    }
    if (pattern.type === "contains") {
      return normalizeText(rd).includes(pattern.text);
    }
    return true;
  }

  function getCashFilterFromUI() {
    const cat = cashFilterCategory?.value || "";
    if (!cat) return { category: "", value: null };

    if (cat === "date") {
      const v = $("#cashFilterDate")?.value || "";
      return { category: "date", value: parseFilterDatePattern(v) };
    }
    if (cat === "contract") {
      const v = normalizeText($("#cashFilterContract")?.value || "");
      return { category: "contract", value: v };
    }
    if (cat === "receipt") {
      const v = normalizeText($("#cashFilterReceipt")?.value || "");
      return { category: "receipt", value: v };
    }
    if (cat === "client") {
      const v = normalizeText($("#cashFilterClient")?.value || "");
      return { category: "client", value: v };
    }
    if (cat === "particular") {
      const v = normalizeText($("#cashFilterParticular")?.value || "");
      return { category: "particular", value: v };
    }
    if (cat === "amount") {
      const minv = Number($("#cashFilterMin")?.value || "");
      const maxv = Number($("#cashFilterMax")?.value || "");
      return {
        category: "amount",
        value: {
          min: Number.isFinite(minv) ? minv : null,
          max: Number.isFinite(maxv) ? maxv : null
        }
      };
    }
    return { category: "", value: null };
  }

  function cashRowMatchesFilter(r, f) {
    if (!f || !f.category) return true;
    const cat = f.category;
    const val = f.value;

    if (cat === "date") return rowMatchesDateFilter(r.date, val);
    if (cat === "contract") return !val || normalizeText(r.contract).includes(val);
    if (cat === "receipt") {
      if (!val) return true;
      if (val === "(blank)" || val === "blank") return !normalizeText(r.receipt);
      return normalizeText(r.receipt).includes(val);
    }
    if (cat === "client") return !val || normalizeText(r.client).includes(val);
    if (cat === "particular") return !val || normalizeText(r.particular).includes(val);
    if (cat === "amount") {
      const a = Number(r.amount) || 0;
      const minv = val?.min;
      const maxv = val?.max;
      if (minv != null && a < minv) return false;
      if (maxv != null && a > maxv) return false;
      return true;
    }
    return true;
  }

  function getFilteredCashRows() {
    return cashStore.filter(r => cashRowMatchesFilter(r, cashActiveFilter));
  }

  btnCashApplyFilter?.addEventListener("click", () => {
    cashActiveFilter = getCashFilterFromUI();
    renderCashTable();
  });

  btnCashClearFilter?.addEventListener("click", () => {
    cashActiveFilter = { category: "", value: null };
    if (cashFilterCategory) cashFilterCategory.value = "";
    setCashFilterInputs("");
    renderCashTable();
  });

  setCashFilterInputs(cashFilterCategory?.value || "");

  function renderCashTable() {
    if (!cashTable) return;

    // sort by date then contract then client
    const rows = getFilteredCashRows().sort((a, b) => {
      const ak = monthKeyFromDate(a.date);
      const bk = monthKeyFromDate(b.date);
      if (ak !== bk) {
        if (ak === "unknown") return 1;
        if (bk === "unknown") return -1;
        return ak.localeCompare(bk);
      }
      const ad = parseMMDDYYYY(a.date)?.getTime() ?? Number.POSITIVE_INFINITY;
      const bd = parseMMDDYYYY(b.date)?.getTime() ?? Number.POSITIVE_INFINITY;
      if (ad !== bd) return ad - bd;
      return (a.client || "").localeCompare(b.client || "");
    });

    // group by month
    const groups = new Map();
    for (const r of rows) {
      const k = monthKeyFromDate(r.date);
      if (!groups.has(k)) groups.set(k, []);
      groups.get(k).push(r);
    }
    const keys = Array.from(groups.keys()).sort((a,b) => {
      if (a === "unknown" && b !== "unknown") return 1;
      if (b === "unknown" && a !== "unknown") return -1;
      return a.localeCompare(b);
    });

    cashTable.tBodies[0].innerHTML = "";
    let grand = 0;

    for (let gi=0; gi<keys.length; gi++) {
      const key = keys[gi];
      const grp = groups.get(key) || [];
      if (!grp.length) continue;

      cashTable.tBodies[0].appendChild(cashMonthHeaderRow(monthLabelFromKey(key)));

      let monthTotal = 0;
      for (const r of grp) {
        ensureCashId(r);
        const tr = document.createElement("tr");
        tr.dataset.rowType = "data";
        tr.dataset.receipt = cashKeyFor(r);

        const tdDate = document.createElement("td"); tdDate.textContent = r.date;
        const tdContract = document.createElement("td"); tdContract.textContent = r.contract;
        const tdReceipt = document.createElement("td"); tdReceipt.textContent = r.receipt || "";
        const tdClient = document.createElement("td"); tdClient.textContent = r.client;
        const tdPart = document.createElement("td"); tdPart.textContent = r.particular;
        const tdAmt = document.createElement("td"); tdAmt.textContent = fmtMoney(r.amount); tdAmt.classList.add("num");

        tr.append(tdDate, tdContract, tdReceipt, tdClient, tdPart, tdAmt);
        cashTable.tBodies[0].appendChild(tr);

        monthTotal += Number(r.amount) || 0;
      }

      cashTable.tBodies[0].appendChild(cashMonthTotalRow(monthTotal));
      grand += monthTotal;

      if (gi < keys.length - 1) cashTable.tBodies[0].appendChild(cashSpacerRow());
    }

    // Bottom running total (grand total)
    if (rows.length) {
      cashTable.tBodies[0].appendChild(cashSpacerRow());
      cashTable.tBodies[0].appendChild(cashGrandTotalRow(grand));
    }

    cashSelectedKey = null;
    cashSelectedEl.textContent = "Selected: —";
    cashRowCountEl.textContent = `Rows: ${rows.length}`;
    // Keep BAI monthly totals in sync when Cash Received changes
    if (typeof renderBaiTable === "function") try { renderBaiTable(); } catch(e) {}
  }

  function selectCashKey(key) {
    if (!cashTable) return;
    const k = normalizeText(key);
    cashSelectedKey = k || null;
    Array.from(cashTable.tBodies[0].rows).forEach(r => r.classList.remove("is-selected"));
    if (!k) { cashSelectedEl.textContent = "Selected: —"; return; }
    const tr = cashTable.tBodies[0].querySelector(`tr[data-row-type="data"][data-receipt="${CSS.escape(k)}"]`);
    if (tr) {
      tr.classList.add("is-selected");
      cashSelectedEl.textContent = `Selected: ${tr.cells[2]?.innerText || "(no receipt)"}`;
    } else {
      cashSelectedEl.textContent = "Selected: —";
    }
  }

  cashTable?.addEventListener("click", (e) => {
    const tr = e.target.closest("tr");
    if (!tr || tr.parentElement?.tagName !== "TBODY") return;
    if ((tr.dataset.rowType || "data") !== "data") return;
    selectCashKey(tr.dataset.receipt);
  });

  cashTable?.addEventListener("dblclick", (e) => {
    const tr = e.target.closest("tr");
    if (!tr || tr.parentElement?.tagName !== "TBODY") return;
    if ((tr.dataset.rowType || "data") !== "data") return;
    selectCashKey(tr.dataset.receipt);
    openCashModal("edit", tr.dataset.receipt);
  });

  btnCashAdd?.addEventListener("click", () => openCashModal("add"));
  btnCashEdit?.addEventListener("click", () => {
    if (!cashSelectedKey) return alert("Please select a Cash Received row first.");
    openCashModal("edit", cashSelectedKey);
  });

  cashForm?.addEventListener("submit", (e) => {
    e.preventDefault();

    const date = cashYyyyMMddToMMDDYYYY(crDate.value) || "";
    const contract = (crContract.value || "").trim();
    const receipt = (crReceipt.value || "").trim(); // optional
    const client = (crClient.value || "").trim();
    const particular = canonicalizeParticularText(crParticular.value);
    const amount = parseMoney2(crAmount.value);

    if (!client) return alert("Name of Client is required.");

    if (cashMode === "add") {
      const entry = ensureCashId({ date, contract, receipt, client, particular, amount });
      if (receipt) {
        const exists = cashStore.some(x => normalizeText(x.receipt) === normalizeText(receipt));
        if (exists && !confirm("That Receipt # already exists. Add anyway?")) return;
      }
      DB.saveCashReceived(entry).then(saved => { if (saved) entry.id = saved.id; });
      cashStore.push(entry);
      syncCashReceivedToContracts();
      closeCashModal();
      renderCashTable();
      selectCashKey(cashKeyFor(entry));
      return;
    }

    // edit
    const key = cashEditingKey;
    if (!key) return;
    const idx = cashStore.findIndex(x => cashKeyFor(x) === key);
    if (idx < 0) return;
    const updatedEntry = ensureCashId({ date, contract, receipt, client, particular, amount, _id: cashStore[idx]._id, id: cashStore[idx].id });
    const newKey = cashKeyFor(updatedEntry);
    if (receipt) {
      const dup = cashStore.some((x, i) => i !== idx && normalizeText(x.receipt) === normalizeText(receipt));
      if (dup && !confirm("That Receipt # already exists. Save anyway?")) return;
    }
    cashStore[idx] = updatedEntry;
    DB.saveCashReceived(updatedEntry);
    syncCashReceivedToContracts();
    closeCashModal();
    renderCashTable();
    selectCashKey(newKey);
  });

  btnCashDel?.addEventListener("click", () => {
    if (!cashSelectedKey) return alert("Please select a Cash Received row first.");
    if (!confirm("Delete the selected Cash Received entry?")) return;
    const key = cashSelectedKey;
    const row = cashStore.find(x => cashKeyFor(x) === key);
    if (row?.id) DB.deleteCashReceived(row.id);
    cashStore = cashStore.filter(x => cashKeyFor(x) !== key);
    syncCashReceivedToContracts();
    renderCashTable();
  });

  // ---------- Import / Export ----------
  btnCashImport?.addEventListener("click", () => fileCashImport?.click());

  function normHeader(h) {
    return String(h ?? "").trim().toLowerCase().replace(/\s+/g, " ");
  }

  const CASH_HEADER_ALIASES = {
    date: ["date"],
    contract: ["contract #", "contract#", "contract number", "contract no", "contract"],
    receipt: ["receipt #", "receipt#", "receipt no", "receipt number", "or", "official receipt"],
    client: ["name of client", "client", "name", "customer", "customer name"],
    particular: ["particular", "description", "remarks", "memo"],
    amount: ["amount", "amt", "value"]
  };

  function buildCashHeaderIndex(headers) {
    const norms = headers.map(normHeader);
    const map = {};
    for (const k of Object.keys(CASH_HEADER_ALIASES)) {
      for (const alias of CASH_HEADER_ALIASES[k]) {
        const ix = norms.indexOf(normHeader(alias));
        if (ix >= 0) { map[k] = ix; break; }
      }
    }
    return map;
  }

  function detectHeaderRow(rows) {
    for (let r = 0; r < Math.min(rows.length, 6); r++) {
      const headers = rows[r].map(x => String(x ?? "").trim()).filter(Boolean);
      if (headers.length < 2) continue;
      const map = buildCashHeaderIndex(headers);
      if (Object.keys(map).length >= 2) return r;
    }
    return -1;
  }

  function parseCSVTo2D(text) {
    const rows = [];
    let row = [];
    let i = 0, field = "", inQuotes = false;
    while (i < text.length) {
      const c = text[i];
      if (inQuotes) {
        if (c === '"') {
          if (text[i + 1] === '"') { field += '"'; i += 2; continue; }
          inQuotes = false; i++; continue;
        }
        field += c; i++; continue;
      } else {
        if (c === '"') { inQuotes = true; i++; continue; }
        if (c === ",") { row.push(field); field = ""; i++; continue; }
        if (c === "\n") { row.push(field); rows.push(row); row = []; field = ""; i++; continue; }
        if (c === "\r") { i++; continue; }
        field += c; i++; continue;
      }
    }
    row.push(field);
    rows.push(row);
    return rows;
  }

  async function readXLSX(file) {
    if (typeof XLSX === "undefined") {
      alert("XLSX library not loaded. If you're offline, import a CSV instead.");
      return null;
    }
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array", cellDates: true });
    const sheetName = wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    return XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  }

  function excelSerialToDate(serial) {
    const n = Number(serial);
    if (!Number.isFinite(n)) return null;
    const utc_days = Math.floor(n - 25569);
    const utc_value = utc_days * 86400;
    return new Date(utc_value * 1000);
  }

  function parseDateAny(v) {
    if (v == null || v === "") return "";
    if (v instanceof Date) return formatMMDDYYYY(v);
    if (typeof v === "number") {
      const d = excelSerialToDate(v);
      return d ? formatMMDDYYYY(d) : "";
    }
    const s = String(v).trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return yyyyMMddToMMDDYYYY(s);
    if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(s)) {
      const parts = s.split("/");
      return `${pad2(parts[0])}/${pad2(parts[1])}/${parts[2]}`;
    }
    const d = new Date(s);
    if (Number.isFinite(d.getTime())) return formatMMDDYYYY(d);
    return s;
  }

  function getCell(row, idx) {
    return (idx >= 0 && idx < row.length) ? row[idx] : "";
  }

  async function importCashRows(rows2d) {
    const headerRowIndex = detectHeaderRow(rows2d);
    let headers = null;
    let dataRows = null;

    if (headerRowIndex >= 0) {
      headers = rows2d[headerRowIndex].map(x => String(x ?? "").trim());
      dataRows = rows2d.slice(headerRowIndex + 1);
    } else {
      headers = ["Date","Contract #","Receipt #","Name of Client","Particular","Amount"];
      dataRows = rows2d;
    }

    dataRows = dataRows.filter(r => r && r.some(v => String(v ?? "").trim() !== ""));
    if (!dataRows.length) { alert("No data rows found."); return; }

    const map = buildCashHeaderIndex(headers);

    const replace = confirm("Replace current Cash Received entries with imported data?\nOK = Replace\nCancel = Append/Update");
    if (replace) cashStore = [];

    let dupMode = "skip";
    if (!replace) {
      const overwrite = confirm("When a Receipt # already exists:\nOK = Update existing\nCancel = Skip duplicates");
      dupMode = overwrite ? "update" : "skip";
    }

    let added = 0, updated = 0, skipped = 0;

    for (const row of dataRows) {
      const entry = ensureCashId({
        date: parseDateAny(getCell(row, map.date ?? 0)),
        contract: String(getCell(row, map.contract ?? 1) ?? "").trim(),
        receipt: String(getCell(row, map.receipt ?? 2) ?? "").trim(),
        client: String(getCell(row, map.client ?? 3) ?? "").trim(),
        particular: String(getCell(row, map.particular ?? 4) ?? "").trim(),
        amount: parseMoney2(getCell(row, map.amount ?? 5))
      });

      if (!entry.client && !entry.receipt) { skipped++; continue; }

      if (entry.receipt) {
        const idx = cashStore.findIndex(x => normalizeText(x.receipt) === normalizeText(entry.receipt));
        if (idx >= 0 && !replace) {
          if (dupMode === "update") { entry._id = cashStore[idx]._id; cashStore[idx] = entry; updated++; }
          else skipped++;
        } else {
          cashStore.push(entry); added++;
        }
      } else {
        // no receipt => always append as unique entry
        cashStore.push(entry);
        added++;
      }
    }

    saveCashStore();
    syncCashReceivedToContracts();
    renderCashTable();
    alert(`Import complete.\nAdded: ${added}\nUpdated: ${updated}\nSkipped: ${skipped}`);
  }

  fileCashImport?.addEventListener("change", async () => {
    const file = fileCashImport.files?.[0];
    if (!file) return;
    try {
      const name = (file.name || "").toLowerCase();
      if (name.endsWith(".csv")) {
        const text = await file.text();
        await importCashRows(parseCSVTo2D(text));
      } else if (name.endsWith(".xlsx")) {
        const rows = await readXLSX(file);
        if (!rows) return;
        await importCashRows(rows);
      } else {
        alert("Unsupported file type. Please select a .csv or .xlsx file.");
      }
    } catch (err) {
      alert("Import failed: " + (err?.message || String(err)));
    } finally {
      fileCashImport.value = "";
    }
  });

  btnCashExport?.addEventListener("click", () => {
    // Styled Excel export with month grouping + totals
    const cols = ["Date","Contract #","Receipt #","Name of Client","Particular","Amount"];
    const msoMoney = 'mso-number-format:"\\#\\,\\#\\#0\\.00"';
    const msoText = 'mso-number-format:"\\@"';
    const esc = (s) => String(s ?? "").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");

    // sort + group like render
    const rows = getFilteredCashRows().sort((a, b) => {
      const ak = monthKeyFromDate(a.date);
      const bk = monthKeyFromDate(b.date);
      if (ak !== bk) {
        if (ak === "unknown") return 1;
        if (bk === "unknown") return -1;
        return ak.localeCompare(bk);
      }
      const ad = parseMMDDYYYY(a.date)?.getTime() ?? Number.POSITIVE_INFINITY;
      const bd = parseMMDDYYYY(b.date)?.getTime() ?? Number.POSITIVE_INFINITY;
      if (ad !== bd) return ad - bd;
      return (a.client || "").localeCompare(b.client || "");
    });

    const groups = new Map();
    for (const r of rows) {
      const k = monthKeyFromDate(r.date);
      if (!groups.has(k)) groups.set(k, []);
      groups.get(k).push(r);
    }
    const keys = Array.from(groups.keys()).sort((a,b) => {
      if (a === "unknown" && b !== "unknown") return 1;
      if (b === "unknown" && a !== "unknown") return -1;
      return a.localeCompare(b);
    });

    const css = `
      <style>
        table { border-collapse: collapse; font-family: Calibri, Arial, sans-serif; font-size: 11pt; }
        th, td { border: 1px solid #cfcfcf; padding: 4px 6px; white-space: nowrap; }
        th { font-weight: 700; background: #ffffff; text-align:center; }
        td.num { text-align: right; }
        tr.month-header td{ font-weight: 800; text-align:center; background:#d9d9d9; border-top:2px solid #000; border-bottom:2px solid #000; }
        tr.month-total td{ font-weight: 800; background:#fff2cc; border-top:2px solid #000; }
        tr.month-total td.label{ text-align:center; letter-spacing:.10em; }
        tr.grand-total td{ font-weight: 900; background:#92d050; border-top:2px solid #000; border-bottom:2px solid #000; }
        tr.grand-total td.label{ text-align:center; letter-spacing:.10em; }
        tr.spacer td{ border:none; height:10px; }
      </style>
    `;

    let html = `<!doctype html><html><head><meta charset="utf-8" />${css}</head><body>`;
    html += `<table><tbody>`;

    let grand = 0;

    for (let gi=0; gi<keys.length; gi++) {
      const key = keys[gi];
      const grp = groups.get(key) || [];
      if (!grp.length) continue;

      // Month header
      html += `<tr class="month-header"><td colspan="6" style="${msoText}">${esc(monthLabelFromKey(key).toUpperCase())}</td></tr>`;
      // Column headers for group
      html += `<tr>`;
      for (const h of cols) html += `<th>${esc(h)}</th>`;
      html += `</tr>`;

      let monthTotal = 0;
      for (const r of grp) {
        html += `<tr>`;
        html += `<td style="${msoText}">${esc(r.date)}</td>`;
        html += `<td style="${msoText}">${esc(r.contract)}</td>`;
        html += `<td style="${msoText}">${esc(r.receipt || "")}</td>`;
        html += `<td style="${msoText}">${esc(r.client)}</td>`;
        html += `<td style="${msoText}">${esc(r.particular)}</td>`;
        html += `<td class="num" style="${msoMoney}">${fmtMoney(r.amount)}</td>`;
        html += `</tr>`;
        monthTotal += Number(r.amount) || 0;
      }

      html += `<tr class="month-total">`;
      html += `<td colspan="4"></td>`;
      html += `<td class="label">TOTAL</td>`;
      html += `<td class="num" style="${msoMoney}">${fmtMoney(monthTotal)}</td>`;
      html += `</tr>`;

      grand += monthTotal;
      if (gi < keys.length - 1) html += `<tr class="spacer"><td colspan="6"></td></tr>`;
    }

    if (rows.length) {
      html += `<tr class="spacer"><td colspan="6"></td></tr>`;
      html += `<tr class="grand-total">`;
      html += `<td colspan="4"></td>`;
      html += `<td class="label">GRAND TOTAL</td>`;
      html += `<td class="num" style="${msoMoney}">${fmtMoney(grand)}</td>`;
      html += `</tr>`;
    }

    html += `</tbody></table></body></html>`;

    const blob = new Blob([html], { type: "application/vnd.ms-excel;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `Magallanes_CashReceived_${new Date().toISOString().slice(0,10)}.xls`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  });

  btnCashRefresh?.addEventListener("click", () => {
    renderCashTable();
    alert("Refreshed Cash Received view.");
  });

  // Initialize Cash store (load saved first, otherwise read sample row from DOM)
  function initCashFromDOM() {
    if (!cashTable) return [];
    const trs = Array.from(cashTable.tBodies[0].rows).filter(r => (r.dataset.rowType || "data") === "data");
    return trs.map(tr => ensureCashId({
      date: String(tr.cells[0]?.innerText || "").trim(),
      contract: String(tr.cells[1]?.innerText || "").trim(),
      receipt: String(tr.cells[2]?.innerText || "").trim(),
      client: String(tr.cells[3]?.innerText || "").trim(),
      particular: String(tr.cells[4]?.innerText || "").trim(),
      amount: parseMoney2(String(tr.cells[5]?.innerText || ""))
    })).filter(r => r.client || r.receipt || r.date);
  }

  // Load cash received from Supabase
  DB.getCashReceived().then(rows => {
    cashStore = rows;
    syncCashReceivedToContracts();
    renderCashTable();
  });


  // ---------------------------
  // Cash Expense (Transactions) — full functionality + monthly grouping (v33)
  // ---------------------------
  const cashExpTable = $("#cashExpenseTable");
  const cashExpRowCountEl = $("#cashExpenseRowCount");
  const cashExpSelectedEl = $("#cashExpenseSelected");

  const btnCashExpAdd = $("#btnCashExpAddEntry");
  const btnCashExpEdit = $("#btnCashExpEditSelected");
  const btnCashExpDel = $("#btnCashExpDeleteSelected");
  const btnCashExpImport = $("#btnCashExpImportExcel");
  const btnCashExpExport = $("#btnCashExpExportExcel");
  const btnCashExpRefresh = $("#btnCashExpRefresh");
  const fileCashExpImport = $("#fileCashExpImport");

  const cashExpOverlay = $("#cashExpOverlay");
  const cashExpModal = $("#cashExpModal");
  const cashExpTitle = $("#cashExpTitle");
  const cashExpSubtitle = $("#cashExpSubtitle");
  const cashExpForm = $("#cashExpForm");
  const btnCloseCashExp = $("#btnCloseCashExp");
  const btnCancelCashExp = $("#btnCancelCashExp");
  const btnSubmitCashExp = $("#btnSubmitCashExp");

  const ceDate = $("#ceDate");
  const ceParticular = $("#ceParticular");
  const ceAmount = $("#ceAmount");

  const CASH_EXP_STORE_KEY = "mf_cash_expense_store_v33";

  /** @type {Array<{date:string, particular:string, amount:number, _id?:string}>} */
  let cashExpStore = [];

  let cashExpSelectedKey = null;
  let cashExpMode = "add";
  let cashExpEditingKey = null;

  function ensureCashExpId(r) {
    if (!r._id) r._id = "__AUTO__" + Date.now() + "_" + Math.floor(Math.random()*100000);
    return r;
  }
  function cashExpKeyFor(r) { return normalizeText(r._id || ""); }

  function saveCashExpStore() { /* no-op */ }
  function loadCashExpStore() { return null; /* loaded async */ }

  function cashExpYyyyMMddToMMDDYYYY(yyyyMMdd) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(yyyyMMdd || ""));
    return m ? `${m[2]}/${m[3]}/${m[1]}` : "";
  }
  function cashExpMMDDYYYYToYyyyMMdd(mmddyyyy) {
    const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(String(mmddyyyy || "").trim());
    if (!m) return "";
    const mm = String(m[1]).padStart(2,"0");
    const dd = String(m[2]).padStart(2,"0");
    return `${m[3]}-${mm}-${dd}`;
  }
  function parseMoney3(s) {
    const n = Number(String(s ?? "").replace(/[^0-9.-]/g, ""));
    return Number.isFinite(n) ? n : 0;
  }

  // Filters
  const cashExpFilterCategory = $("#cashExpFilterCategory");
  const cashExpFilterInputs = $("#cashExpFilterInputs");
  const btnCashExpApplyFilter = $("#btnCashExpApplyFilter");
  const btnCashExpClearFilter = $("#btnCashExpClearFilter");
  let cashExpActiveFilter = { category: "", value: null };

  function setCashExpFilterInputs(category) {
    if (!cashExpFilterInputs) return;
    cashExpFilterInputs.innerHTML = "";
    const makeText = (id, label, placeholder="") => {
      const wrap = document.createElement("label");
      wrap.className = "field inline";
      wrap.innerHTML = `<span>${label}</span><input class="input" id="${id}" type="text" placeholder="${placeholder}"/>`;
      cashExpFilterInputs.appendChild(wrap);
    };
    const makeNumber = (id, label, placeholder="") => {
      const wrap = document.createElement("label");
      wrap.className = "field inline";
      wrap.innerHTML = `<span>${label}</span><input class="input" id="${id}" type="number" step="0.01" placeholder="${placeholder}"/>`;
      cashExpFilterInputs.appendChild(wrap);
    };
    if (!category) return;
    if (category === "date") return makeText("cashExpFilterDate", "Date", "MM/DD/YYYY or MM/YYYY or YYYY");
    if (category === "particular") return makeText("cashExpFilterParticular", "Particular", "contains...");
    if (category === "amount") { makeNumber("cashExpFilterMin","Min","0.00"); makeNumber("cashExpFilterMax","Max","10000.00"); }
  }
  cashExpFilterCategory?.addEventListener("change", () => setCashExpFilterInputs(cashExpFilterCategory.value));

  function parseFilterDatePattern2(s) {
    const t = String(s || "").trim();
    if (!t) return null;
    if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(t)) {
      const d = parseMMDDYYYY(t);
      return d ? { type:"full", d: formatMMDDYYYY(d) } : null;
    }
    if (/^\d{1,2}\/\d{4}$/.test(t)) {
      const [mm, yyyy] = t.split("/"); return { type:"monthYear", mm: pad2(mm), yyyy };
    }
    if (/^\d{4}$/.test(t)) return { type:"year", yyyy: t };
    return { type:"contains", text: normalizeText(t) };
  }
  function rowMatchesDateFilter2(rowDate, pattern) {
    if (!pattern) return true;
    const rd = String(rowDate || "").trim();
    if (!rd) return false;
    if (pattern.type === "full") return rd === pattern.d;
    if (pattern.type === "monthYear") {
      const m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(rd);
      return m ? (m[1]===pattern.mm && m[3]===pattern.yyyy) : false;
    }
    if (pattern.type === "year") {
      const m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(rd);
      return m ? (m[3]===pattern.yyyy) : false;
    }
    if (pattern.type === "contains") return normalizeText(rd).includes(pattern.text);
    return true;
  }
  function getCashExpFilterFromUI() {
    const cat = cashExpFilterCategory?.value || "";
    if (!cat) return { category:"", value:null };
    if (cat === "date") return { category:"date", value: parseFilterDatePattern2($("#cashExpFilterDate")?.value || "") };
    if (cat === "particular") return { category:"particular", value: normalizeText($("#cashExpFilterParticular")?.value || "") };
    if (cat === "amount") {
      const minv = Number($("#cashExpFilterMin")?.value || "");
      const maxv = Number($("#cashExpFilterMax")?.value || "");
      return { category:"amount", value:{ min: Number.isFinite(minv)?minv:null, max: Number.isFinite(maxv)?maxv:null } };
    }
    return { category:"", value:null };
  }
  function cashExpRowMatchesFilter(r, f) {
    if (!f || !f.category) return true;
    if (f.category === "date") return rowMatchesDateFilter2(r.date, f.value);
    if (f.category === "particular") return !f.value || normalizeText(r.particular).includes(f.value);
    if (f.category === "amount") {
      const a = Number(r.amount)||0;
      const minv = f.value?.min; const maxv = f.value?.max;
      if (minv != null && a < minv) return false;
      if (maxv != null && a > maxv) return false;
      return true;
    }
    return true;
  }
  function getFilteredCashExpRows() { return cashExpStore.filter(r => cashExpRowMatchesFilter(r, cashExpActiveFilter)); }
  btnCashExpApplyFilter?.addEventListener("click", () => { cashExpActiveFilter = getCashExpFilterFromUI(); renderCashExpTable(); });
  btnCashExpClearFilter?.addEventListener("click", () => {
    cashExpActiveFilter = { category:"", value:null };
    if (cashExpFilterCategory) cashExpFilterCategory.value = "";
    setCashExpFilterInputs("");
    renderCashExpTable();
  });
  setCashExpFilterInputs(cashExpFilterCategory?.value || "");

  function openCashExpModal(mode, keyOrNull=null) {
    if (!cashExpOverlay || !cashExpModal) { alert("Cash Expense form not available."); return; }
    cashExpMode = mode; cashExpEditingKey = null;
    if (mode === "add") {
      cashExpTitle.textContent = "Add Entry";
      cashExpSubtitle.textContent = "Enter cash expense details.";
      btnSubmitCashExp.textContent = "Add Entry";
      ceDate.value = new Date().toISOString().slice(0,10);
      ceParticular.value = "";
      ceAmount.value = "0.00";
    } else {
      cashExpTitle.textContent = "Edit Entry";
      cashExpSubtitle.textContent = "Update the selected cash expense entry.";
      btnSubmitCashExp.textContent = "Save Changes";
      const key = normalizeText(keyOrNull);
      const found = cashExpStore.find(x => cashExpKeyFor(x) === key);
      if (!found) { alert("Could not find the selected entry."); return; }
      cashExpEditingKey = key;
      ceDate.value = cashExpMMDDYYYYToYyyyMMdd(found.date) || new Date().toISOString().slice(0,10);
      ceParticular.value = found.particular || "";
      ceAmount.value = (Number(found.amount)||0).toFixed(2);
    }
    cashExpOverlay.classList.add("is-open");
    cashExpModal.classList.add("is-open");
    setTimeout(() => ceParticular.focus(), 0);
  }
  function closeCashExpModal() {
    cashExpOverlay?.classList.remove("is-open");
    cashExpModal?.classList.remove("is-open");
    cashExpMode = "add"; cashExpEditingKey = null;
  }
  btnCloseCashExp?.addEventListener("click", closeCashExpModal);
  btnCancelCashExp?.addEventListener("click", closeCashExpModal);
  cashExpOverlay?.addEventListener("click", closeCashExpModal);
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && cashExpModal?.classList.contains("is-open")) closeCashExpModal();
  });

  function cashExpMonthHeaderRow(label) {
    const tr = document.createElement("tr");
    tr.classList.add("cashExpMonthHeader");
    tr.dataset.rowType = "cashExpMonthHeader";
    const td = document.createElement("td");
    td.colSpan = 3;
    td.textContent = String(label||"").toUpperCase();
    tr.appendChild(td);
    return tr;
  }
  function cashExpMonthTotalRow(total) {
    const tr = document.createElement("tr");
    tr.classList.add("cashExpMonthTotal");
    tr.dataset.rowType = "cashExpMonthTotal";
    const td1 = document.createElement("td"); td1.textContent="";
    const td2 = document.createElement("td"); td2.textContent="TOTAL"; td2.classList.add("label");
    const td3 = document.createElement("td"); td3.textContent=fmtMoney(total); td3.classList.add("num");
    tr.append(td1, td2, td3);
    return tr;
  }
  function cashExpGrandTotalRow(total) {
    const tr = document.createElement("tr");
    tr.classList.add("cashExpGrandTotal");
    tr.dataset.rowType = "cashExpGrandTotal";
    const td1 = document.createElement("td"); td1.textContent="";
    const td2 = document.createElement("td"); td2.textContent="GRAND TOTAL"; td2.classList.add("label");
    const td3 = document.createElement("td"); td3.textContent=fmtMoney(total); td3.classList.add("num");
    tr.append(td1, td2, td3);
    return tr;
  }
  function cashExpSpacerRow() {
    const tr = document.createElement("tr");
    tr.classList.add("spacer");
    tr.dataset.rowType = "spacer";
    const td = document.createElement("td"); td.colSpan=3; td.textContent="";
    tr.appendChild(td);
    return tr;
  }

  function renderCashExpTable() {
    if (!cashExpTable) return;
    const rows = getFilteredCashExpRows().sort((a,b)=>{
      const ak = monthKeyFromDate(a.date);
      const bk = monthKeyFromDate(b.date);
      if (ak !== bk) {
        if (ak === "unknown") return 1;
        if (bk === "unknown") return -1;
        return ak.localeCompare(bk);
      }
      const ad = parseMMDDYYYY(a.date)?.getTime() ?? Number.POSITIVE_INFINITY;
      const bd = parseMMDDYYYY(b.date)?.getTime() ?? Number.POSITIVE_INFINITY;
      if (ad !== bd) return ad - bd;
      return (a.particular||"").localeCompare(b.particular||"");
    });

    const groups = new Map();
    for (const r of rows) {
      const k = monthKeyFromDate(r.date);
      if (!groups.has(k)) groups.set(k, []);
      groups.get(k).push(r);
    }
    const keys = Array.from(groups.keys()).sort((a,b)=>{
      if (a === "unknown" && b !== "unknown") return 1;
      if (b === "unknown" && a !== "unknown") return -1;
      return a.localeCompare(b);
    });

    cashExpTable.tBodies[0].innerHTML = "";
    let grand = 0;

    for (let gi=0; gi<keys.length; gi++) {
      const key = keys[gi];
      const grp = groups.get(key) || [];
      if (!grp.length) continue;

      cashExpTable.tBodies[0].appendChild(cashExpMonthHeaderRow(monthLabelFromKey(key)));
      let monthTotal = 0;

      for (const r of grp) {
        ensureCashExpId(r);
        const tr = document.createElement("tr");
        tr.dataset.rowType="data";
        tr.dataset.key=cashExpKeyFor(r);

        const tdDate = document.createElement("td"); tdDate.textContent=r.date;
        const tdPart = document.createElement("td"); tdPart.textContent=r.particular;
        const tdAmt = document.createElement("td"); tdAmt.textContent=fmtMoney(r.amount); tdAmt.classList.add("num");

        tr.append(tdDate, tdPart, tdAmt);
        cashExpTable.tBodies[0].appendChild(tr);
        monthTotal += Number(r.amount)||0;
      }

      cashExpTable.tBodies[0].appendChild(cashExpMonthTotalRow(monthTotal));
      grand += monthTotal;
      if (gi < keys.length - 1) cashExpTable.tBodies[0].appendChild(cashExpSpacerRow());
    }

    if (rows.length) {
      cashExpTable.tBodies[0].appendChild(cashExpSpacerRow());
      cashExpTable.tBodies[0].appendChild(cashExpGrandTotalRow(grand));
    }

    cashExpSelectedKey = null;
    cashExpSelectedEl.textContent="Selected: —";
    cashExpRowCountEl.textContent=`Rows: ${rows.length}`;
  }

  function selectCashExpKey(key) {
    if (!cashExpTable) return;
    const k = normalizeText(key);
    cashExpSelectedKey = k || null;
    Array.from(cashExpTable.tBodies[0].rows).forEach(r=>r.classList.remove("is-selected"));
    if (!k) { cashExpSelectedEl.textContent="Selected: —"; return; }
    const tr = cashExpTable.tBodies[0].querySelector(`tr[data-row-type="data"][data-key="${CSS.escape(k)}"]`);
    if (tr) { tr.classList.add("is-selected"); cashExpSelectedEl.textContent=`Selected: ${tr.cells[1]?.innerText || "—"}`; }
    else cashExpSelectedEl.textContent="Selected: —";
  }

  cashExpTable?.addEventListener("click",(e)=>{
    const tr = e.target.closest("tr");
    if (!tr || tr.parentElement?.tagName !== "TBODY") return;
    if ((tr.dataset.rowType || "data") !== "data") return;
    selectCashExpKey(tr.dataset.key);
  });
  cashExpTable?.addEventListener("dblclick",(e)=>{
    const tr = e.target.closest("tr");
    if (!tr || tr.parentElement?.tagName !== "TBODY") return;
    if ((tr.dataset.rowType || "data") !== "data") return;
    selectCashExpKey(tr.dataset.key);
    openCashExpModal("edit", tr.dataset.key);
  });

  btnCashExpAdd?.addEventListener("click",()=>openCashExpModal("add"));
  btnCashExpEdit?.addEventListener("click",()=>{
    if (!cashExpSelectedKey) return alert("Please select a Cash Expense row first.");
    openCashExpModal("edit", cashExpSelectedKey);
  });

  cashExpForm?.addEventListener("submit",(e)=>{
    e.preventDefault();
    const date = cashExpYyyyMMddToMMDDYYYY(ceDate.value) || "";
    const particular = String(ceParticular.value||"").trim();
    const amount = parseMoney3(ceAmount.value);
    if (!particular) return alert("Particular is required.");

    if (cashExpMode === "add") {
      const entry = ensureCashExpId({ date, particular, amount });
      DB.saveCashExpense(entry).then(saved => { if (saved) entry.id = saved.id; });
      cashExpStore.push(entry);
      closeCashExpModal();
      renderCashExpTable();
      selectCashExpKey(cashExpKeyFor(entry));
      return;
    }

    const key = cashExpEditingKey;
    if (!key) return;
    const idx = cashExpStore.findIndex(x => cashExpKeyFor(x) === key);
    if (idx < 0) return;

    const updatedCE = ensureCashExpId({ date, particular, amount, _id: cashExpStore[idx]._id, id: cashExpStore[idx].id });
    cashExpStore[idx] = updatedCE;
    DB.saveCashExpense(updatedCE);
    closeCashExpModal();
    renderCashExpTable();
    selectCashExpKey(cashExpKeyFor(updatedCE));
  });

  btnCashExpDel?.addEventListener("click",()=>{
    if (!cashExpSelectedKey) return alert("Please select a Cash Expense row first.");
    if (!confirm("Delete the selected Cash Expense entry?")) return;
    const rowCE = cashExpStore.find(x => cashExpKeyFor(x) === cashExpSelectedKey);
    if (rowCE?.id) DB.deleteCashExpense(rowCE.id);
    cashExpStore = cashExpStore.filter(x => cashExpKeyFor(x) !== cashExpSelectedKey);
    renderCashExpTable();
  });

  // Import / Export
  btnCashExpImport?.addEventListener("click",()=>fileCashExpImport?.click());

  async function importCashExpRows(rows2d) {
    // Try to detect header, else assume Date/Particular/Amount
    const headerRowIndex = (function(rows){
      for (let r=0;r<Math.min(rows.length,6);r++){
        const row = rows[r].map(x=>String(x??"").trim().toLowerCase());
        if (row.includes("date") && (row.includes("particular") || row.includes("description")) && (row.includes("amount") || row.includes("amt"))) return r;
      }
      return -1;
    })(rows2d);

    let dataRows = rows2d;
    if (headerRowIndex >= 0) dataRows = rows2d.slice(headerRowIndex+1);

    dataRows = dataRows.filter(r => r && r.some(v => String(v ?? "").trim() !== ""));
    if (!dataRows.length) { alert("No data rows found."); return; }

    const replace = confirm("Replace current Cash Expense entries with imported data?\nOK = Replace\nCancel = Append");
    if (replace) cashExpStore = [];

    let added=0, skipped=0;
    for (const row of dataRows) {
      const d = String(row[0] ?? "").trim();
      let date = d;
      if (/^\d{4}-\d{2}-\d{2}$/.test(d)) date = yyyyMMddToMMDDYYYY(d);
      if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(d)) date = `${pad2(d.split("/")[0])}/${pad2(d.split("/")[1])}/${d.split("/")[2]}`;

      const particular = String(row[1] ?? "").trim();
      const amount = parseMoney3(row[2] ?? "");
      if (!particular) { skipped++; continue; }
      cashExpStore.push(ensureCashExpId({ date, particular, amount }));
      added++;
    }
    saveCashExpStore();
    renderCashExpTable();
    alert(`Import complete.\nAdded: ${added}\nSkipped: ${skipped}`);
  }

  fileCashExpImport?.addEventListener("change", async ()=>{
    const file = fileCashExpImport.files?.[0];
    if (!file) return;
    try {
      const name = (file.name||"").toLowerCase();
      if (name.endsWith(".csv")) {
        const text = await file.text();
        await importCashExpRows(parseCSVTo2D(text));
      } else if (name.endsWith(".xlsx")) {
        if (typeof XLSX === "undefined") { alert("XLSX library not loaded. Import a CSV instead."); return; }
        const buf = await file.arrayBuffer();
        const wb = XLSX.read(buf, { type:"array", cellDates:true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { header:1, defval:"" });
        await importCashExpRows(rows);
      } else {
        alert("Unsupported file type. Please select a .csv or .xlsx file.");
      }
    } catch(err) {
      alert("Import failed: " + (err?.message || String(err)));
    } finally {
      fileCashExpImport.value = "";
    }
  });

  btnCashExpExport?.addEventListener("click",()=>{
    const msoMoney = 'mso-number-format:"\\#\\,\\#\\#0\\.00"';
    const msoText = 'mso-number-format:"\\@"';
    const esc = (s)=>String(s??"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");

    const rows = getFilteredCashExpRows().sort((a,b)=>{
      const ak = monthKeyFromDate(a.date);
      const bk = monthKeyFromDate(b.date);
      if (ak !== bk) {
        if (ak === "unknown") return 1;
        if (bk === "unknown") return -1;
        return ak.localeCompare(bk);
      }
      const ad = parseMMDDYYYY(a.date)?.getTime() ?? Number.POSITIVE_INFINITY;
      const bd = parseMMDDYYYY(b.date)?.getTime() ?? Number.POSITIVE_INFINITY;
      if (ad !== bd) return ad - bd;
      return (a.particular||"").localeCompare(b.particular||"");
    });

    const groups = new Map();
    for (const r of rows) {
      const k = monthKeyFromDate(r.date);
      if (!groups.has(k)) groups.set(k, []);
      groups.get(k).push(r);
    }
    const keys = Array.from(groups.keys()).sort((a,b)=>{
      if (a === "unknown" && b !== "unknown") return 1;
      if (b === "unknown" && a !== "unknown") return -1;
      return a.localeCompare(b);
    });

    const css = `
      <style>
        table{ border-collapse:collapse; font-family:Calibri,Arial,sans-serif; font-size:11pt; }
        th,td{ border:1px solid #cfcfcf; padding:4px 6px; white-space:nowrap; }
        th{ font-weight:700; text-align:center; background:#fff; }
        td.num{ text-align:right; }
        tr.month-header td{ font-weight:800; text-align:center; background:#d9d9d9; border-top:2px solid #000; border-bottom:2px solid #000; }
        tr.month-total td{ font-weight:800; background:#fff2cc; border-top:2px solid #000; }
        tr.month-total td.label{ text-align:center; letter-spacing:.10em; }
        tr.grand-total td{ font-weight:900; background:#92d050; border-top:2px solid #000; border-bottom:2px solid #000; }
        tr.grand-total td.label{ text-align:center; letter-spacing:.10em; }
        tr.spacer td{ border:none; height:10px; }
      </style>
    `;

    let html = `<!doctype html><html><head><meta charset="utf-8" />${css}</head><body><table><tbody>`;
    let grand=0;

    for (let gi=0; gi<keys.length; gi++) {
      const key = keys[gi];
      const grp = groups.get(key) || [];
      if (!grp.length) continue;

      html += `<tr class="month-header"><td colspan="3" style="${msoText}">${esc(monthLabelFromKey(key).toUpperCase())}</td></tr>`;
      html += `<tr><th>Date</th><th>Particular</th><th>Amount</th></tr>`;

      let monthTotal=0;
      for (const r of grp) {
        html += `<tr>`;
        html += `<td style="${msoText}">${esc(r.date)}</td>`;
        html += `<td style="${msoText}">${esc(r.particular)}</td>`;
        html += `<td class="num" style="${msoMoney}">${fmtMoney(r.amount)}</td>`;
        html += `</tr>`;
        monthTotal += Number(r.amount)||0;
      }

      html += `<tr class="month-total"><td></td><td class="label">TOTAL</td><td class="num" style="${msoMoney}">${fmtMoney(monthTotal)}</td></tr>`;
      grand += monthTotal;
      if (gi < keys.length - 1) html += `<tr class="spacer"><td colspan="3"></td></tr>`;
    }

    if (rows.length) {
      html += `<tr class="spacer"><td colspan="3"></td></tr>`;
      html += `<tr class="grand-total"><td></td><td class="label">GRAND TOTAL</td><td class="num" style="${msoMoney}">${fmtMoney(grand)}</td></tr>`;
    }

    html += `</tbody></table></body></html>`;
    const blob = new Blob([html], { type:"application/vnd.ms-excel;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href=url;
    a.download=`Magallanes_CashExpense_${new Date().toISOString().slice(0,10)}.xls`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  });

  btnCashExpRefresh?.addEventListener("click",()=>{ renderCashExpTable(); alert("Refreshed Cash Expense view."); });

  function initCashExpFromDOM() {
    if (!cashExpTable) return [];
    const trs = Array.from(cashExpTable.tBodies[0].rows).filter(r => (r.dataset.rowType || "data") === "data");
    return trs.map(tr => ensureCashExpId({
      date: String(tr.cells[0]?.innerText||"").trim(),
      particular: String(tr.cells[1]?.innerText||"").trim(),
      amount: parseMoney3(String(tr.cells[2]?.innerText||""))
    })).filter(r => r.particular || r.date);
  }

  // Load cash expense from Supabase
  DB.getCashExpense().then(rows => {
    cashExpStore = rows;
    renderCashExpTable();
  });

  renderCashExpTable();


  // ---------------------------
  // Bank Received (Transactions) — full functionality + monthly grouping (v35)
  // ---------------------------
  const bankTable = $("#bankReceivedTable");
  const bankRowCountEl = $("#bankReceivedRowCount");
  const bankSelectedEl = $("#bankReceivedSelected");

  const btnBankAdd = $("#btnBankAddEntry");
  const btnBankEdit = $("#btnBankEditSelected");
  const btnBankDel = $("#btnBankDeleteSelected");
  const btnBankImport = $("#btnBankImportExcel");
  const btnBankExport = $("#btnBankExportExcel");
  const btnBankRefresh = $("#btnBankRefresh");
  const fileBankImport = $("#fileBankImport");

  const bankOverlay = $("#bankOverlay");
  const bankModal = $("#bankModal");
  const bankTitle = $("#bankTitle");
  const bankSubtitle = $("#bankSubtitle");
  const bankForm = $("#bankForm");
  const btnCloseBank = $("#btnCloseBank");
  const btnCancelBank = $("#btnCancelBank");
  const btnSubmitBank = $("#btnSubmitBank");

  const brDate = $("#brDate");
  const brContract = $("#brContract");
  const brType = $("#brType");
  const brClient = $("#brClient");
  const brAmount = $("#brAmount");

  const BANK_STORE_KEY = "mf_bank_received_store_v35";

  /** @type {Array<{date:string, contract:string, type:string, client:string, amount:number, _id?:string}>} */
  let bankStore = [];
  let bankSelectedKey = null;
  let bankMode = "add";
  let bankEditingKey = null;

  function ensureBankId(r) {
    if (!r._id) r._id = "__AUTO__" + Date.now() + "_" + Math.floor(Math.random()*100000);
    return r;
  }
  function bankKeyFor(r) { return normalizeText(r._id || ""); }

  function saveBankStore() { /* no-op */ }
  function loadBankStore() { return null; /* loaded async */ }

  function brYyyyMMddToMMDDYYYY(yyyyMMdd) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(yyyyMMdd || ""));
    return m ? `${m[2]}/${m[3]}/${m[1]}` : "";
  }
  function brMMDDYYYYToYyyyMMdd(mmddyyyy) {
    const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(String(mmddyyyy || "").trim());
    if (!m) return "";
    const mm = String(m[1]).padStart(2,"0");
    const dd = String(m[2]).padStart(2,"0");
    return `${m[3]}-${mm}-${dd}`;
  }
  function parseMoneyBR(s) {
    const n = Number(String(s ?? "").replace(/[^0-9.-]/g, ""));
    return Number.isFinite(n) ? n : 0;
  }
  function normalizeBankType(t) {
    const p = normalizeText(t);
    if (!p) return "GCASH";
    if (/g[\s\-]?cash/.test(p) || p.includes("gcash")) return "GCASH";
    if (p.includes("bank") || p.includes("transfer")) return "BANK TRANSFER";
    return String(t).trim().toUpperCase();
  }

  // Filters
  const bankFilterCategory = $("#bankFilterCategory");
  const bankFilterInputs = $("#bankFilterInputs");
  const btnBankApplyFilter = $("#btnBankApplyFilter");
  const btnBankClearFilter = $("#btnBankClearFilter");
  let bankActiveFilter = { category:"", value:null };

  function setBankFilterInputs(category) {
    if (!bankFilterInputs) return;
    bankFilterInputs.innerHTML = "";

    const makeText = (id, label, placeholder="") => {
      const wrap = document.createElement("label");
      wrap.className = "field inline";
      wrap.innerHTML = `<span>${label}</span><input class="input" id="${id}" type="text" placeholder="${placeholder}"/>`;
      bankFilterInputs.appendChild(wrap);
    };
    const makeNumber = (id, label, placeholder="") => {
      const wrap = document.createElement("label");
      wrap.className = "field inline";
      wrap.innerHTML = `<span>${label}</span><input class="input" id="${id}" type="number" step="0.01" placeholder="${placeholder}"/>`;
      bankFilterInputs.appendChild(wrap);
    };
    const makeSelect = (id, label, opts) => {
      const wrap = document.createElement("label");
      wrap.className = "field inline";
      const options = opts.map(o => `<option value="${o.value}">${o.label}</option>`).join("");
      wrap.innerHTML = `<span>${label}</span><select class="select" id="${id}">${options}</select>`;
      bankFilterInputs.appendChild(wrap);
    };

    if (!category) return;
    if (category === "date") return makeText("bankFilterDate","Date","MM/DD/YYYY or MM/YYYY or YYYY");
    if (category === "contract") return makeText("bankFilterContract","Contract #","contains...");
    if (category === "client") return makeText("bankFilterClient","Client","contains...");
    if (category === "type") return makeSelect("bankFilterType","Type",[
      {value:"", label:"(any)"},
      {value:"GCASH", label:"GCASH"},
      {value:"BANK TRANSFER", label:"BANK TRANSFER"}
    ]);
    if (category === "amount") { makeNumber("bankFilterMin","Min","0.00"); makeNumber("bankFilterMax","Max","10000.00"); }
  }
  bankFilterCategory?.addEventListener("change", () => setBankFilterInputs(bankFilterCategory.value));

  function parseFilterDatePatternBR(s) {
    const t = String(s || "").trim();
    if (!t) return null;
    if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(t)) {
      const d = parseMMDDYYYY(t);
      return d ? { type:"full", d: formatMMDDYYYY(d) } : null;
    }
    if (/^\d{1,2}\/\d{4}$/.test(t)) {
      const [mm, yyyy] = t.split("/");
      return { type:"monthYear", mm: pad2(mm), yyyy };
    }
    if (/^\d{4}$/.test(t)) return { type:"year", yyyy:t };
    return { type:"contains", text: normalizeText(t) };
  }
  function rowMatchesDateFilterBR(rowDate, pattern) {
    if (!pattern) return true;
    const rd = String(rowDate || "").trim();
    if (!rd) return false;
    if (pattern.type === "full") return rd === pattern.d;
    if (pattern.type === "monthYear") {
      const m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(rd);
      return m ? (m[1]===pattern.mm && m[3]===pattern.yyyy) : false;
    }
    if (pattern.type === "year") {
      const m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(rd);
      return m ? (m[3]===pattern.yyyy) : false;
    }
    if (pattern.type === "contains") return normalizeText(rd).includes(pattern.text);
    return true;
  }
  function getBankFilterFromUI() {
    const cat = bankFilterCategory?.value || "";
    if (!cat) return { category:"", value:null };
    if (cat === "date") return { category:"date", value: parseFilterDatePatternBR($("#bankFilterDate")?.value || "") };
    if (cat === "contract") return { category:"contract", value: normalizeText($("#bankFilterContract")?.value || "") };
    if (cat === "client") return { category:"client", value: normalizeText($("#bankFilterClient")?.value || "") };
    if (cat === "type") return { category:"type", value: normalizeBankType($("#bankFilterType")?.value || "") };
    if (cat === "amount") {
      const minv = Number($("#bankFilterMin")?.value || "");
      const maxv = Number($("#bankFilterMax")?.value || "");
      return { category:"amount", value:{ min: Number.isFinite(minv)?minv:null, max: Number.isFinite(maxv)?maxv:null } };
    }
    return { category:"", value:null };
  }
  function bankRowMatchesFilter(r, f) {
    if (!f || !f.category) return true;
    if (f.category === "date") return rowMatchesDateFilterBR(r.date, f.value);
    if (f.category === "contract") return !f.value || normalizeText(r.contract).includes(f.value);
    if (f.category === "client") return !f.value || normalizeText(r.client).includes(f.value);
    if (f.category === "type") return !f.value || normalizeBankType(r.type) === f.value;
    if (f.category === "amount") {
      const a = Number(r.amount)||0;
      const minv = f.value?.min; const maxv = f.value?.max;
      if (minv != null && a < minv) return false;
      if (maxv != null && a > maxv) return false;
      return true;
    }
    return true;
  }
  function getFilteredBankRows() { return bankStore.filter(r => bankRowMatchesFilter(r, bankActiveFilter)); }
  btnBankApplyFilter?.addEventListener("click", () => { bankActiveFilter = getBankFilterFromUI(); renderBankTable(); });
  btnBankClearFilter?.addEventListener("click", () => {
    bankActiveFilter = { category:"", value:null };
    if (bankFilterCategory) bankFilterCategory.value = "";
    setBankFilterInputs("");
    renderBankTable();
  });
  setBankFilterInputs(bankFilterCategory?.value || "");

  function openBankModal(mode, keyOrNull=null) {
    if (!bankOverlay || !bankModal) { alert("Bank Received form not available."); return; }
    bankMode = mode; bankEditingKey = null;
    if (mode === "add") {
      bankTitle.textContent = "Add Entry";
      bankSubtitle.textContent = "Enter bank received details.";
      btnSubmitBank.textContent = "Add Entry";
      brDate.value = new Date().toISOString().slice(0,10);
      brContract.value = "";
      brType.value = "GCASH";
      brClient.value = "";
      brAmount.value = "0.00";
    } else {
      bankTitle.textContent = "Edit Entry";
      bankSubtitle.textContent = "Update the selected bank received entry.";
      btnSubmitBank.textContent = "Save Changes";
      const key = normalizeText(keyOrNull);
      const found = bankStore.find(x => bankKeyFor(x) === key);
      if (!found) { alert("Could not find the selected entry."); return; }
      bankEditingKey = key;
      brDate.value = brMMDDYYYYToYyyyMMdd(found.date) || new Date().toISOString().slice(0,10);
      brContract.value = found.contract || "";
      brType.value = normalizeBankType(found.type);
      brClient.value = found.client || "";
      brAmount.value = (Number(found.amount)||0).toFixed(2);
    }
    bankOverlay.classList.add("is-open");
    bankModal.classList.add("is-open");
    setTimeout(() => brClient.focus(), 0);
  }
  function closeBankModal() {
    bankOverlay?.classList.remove("is-open");
    bankModal?.classList.remove("is-open");
    bankMode = "add"; bankEditingKey = null;
  }
  btnCloseBank?.addEventListener("click", closeBankModal);
  btnCancelBank?.addEventListener("click", closeBankModal);
  bankOverlay?.addEventListener("click", closeBankModal);
  document.addEventListener("keydown",(e)=>{
    if (e.key === "Escape" && bankModal?.classList.contains("is-open")) closeBankModal();
  });

  function bankMonthHeaderRow(label) {
    const tr = document.createElement("tr");
    tr.classList.add("bankMonthHeader");
    tr.dataset.rowType = "bankMonthHeader";
    const td = document.createElement("td");
    td.colSpan = 5;
    td.textContent = String(label||"").toUpperCase();
    tr.appendChild(td);
    return tr;
  }
  function bankMonthTotalRow(total) {
    const tr = document.createElement("tr");
    tr.classList.add("bankMonthTotal");
    tr.dataset.rowType = "bankMonthTotal";
    for (let i=0;i<3;i++) tr.appendChild(document.createElement("td"));
    const label = document.createElement("td"); label.textContent="TOTAL"; label.classList.add("label");
    tr.appendChild(label);
    const amt = document.createElement("td"); amt.textContent=fmtMoney(total); amt.classList.add("num");
    tr.appendChild(amt);
    return tr;
  }
  function bankGrandTotalRow(total) {
    const tr = document.createElement("tr");
    tr.classList.add("bankGrandTotal");
    tr.dataset.rowType = "bankGrandTotal";
    for (let i=0;i<3;i++) tr.appendChild(document.createElement("td"));
    const label = document.createElement("td"); label.textContent="GRAND TOTAL"; label.classList.add("label");
    tr.appendChild(label);
    const amt = document.createElement("td"); amt.textContent=fmtMoney(total); amt.classList.add("num");
    tr.appendChild(amt);
    return tr;
  }
  function bankSpacerRow() {
    const tr = document.createElement("tr");
    tr.classList.add("spacer");
    tr.dataset.rowType = "spacer";
    const td = document.createElement("td"); td.colSpan=5; td.textContent="";
    tr.appendChild(td);
    return tr;
  }

  function renderBankTable() {
    if (!bankTable) return;
    const rows = getFilteredBankRows().sort((a,b)=>{
      const ak = monthKeyFromDate(a.date);
      const bk = monthKeyFromDate(b.date);
      if (ak !== bk) {
        if (ak === "unknown") return 1;
        if (bk === "unknown") return -1;
        return ak.localeCompare(bk);
      }
      const ad = parseMMDDYYYY(a.date)?.getTime() ?? Number.POSITIVE_INFINITY;
      const bd = parseMMDDYYYY(b.date)?.getTime() ?? Number.POSITIVE_INFINITY;
      if (ad !== bd) return ad - bd;
      return (a.client||"").localeCompare(b.client||"");
    });

    const groups = new Map();
    for (const r of rows) {
      const k = monthKeyFromDate(r.date);
      if (!groups.has(k)) groups.set(k, []);
      groups.get(k).push(r);
    }
    const keys = Array.from(groups.keys()).sort((a,b)=>{
      if (a === "unknown" && b !== "unknown") return 1;
      if (b === "unknown" && a !== "unknown") return -1;
      return a.localeCompare(b);
    });

    bankTable.tBodies[0].innerHTML = "";
    let grand = 0;

    for (let gi=0; gi<keys.length; gi++) {
      const key = keys[gi];
      const grp = groups.get(key) || [];
      if (!grp.length) continue;

      bankTable.tBodies[0].appendChild(bankMonthHeaderRow(monthLabelFromKey(key)));
      let monthTotal = 0;

      for (const r of grp) {
        ensureBankId(r);
        const tr = document.createElement("tr");
        tr.dataset.rowType="data";
        tr.dataset.key=bankKeyFor(r);

        const tdDate = document.createElement("td"); tdDate.textContent=r.date;
        const tdContract = document.createElement("td"); tdContract.textContent=r.contract || "";
        const tdType = document.createElement("td"); tdType.textContent=normalizeBankType(r.type);
        const tdClient = document.createElement("td"); tdClient.textContent=r.client || "";
        const tdAmt = document.createElement("td"); tdAmt.textContent=fmtMoney(r.amount); tdAmt.classList.add("num");

        tr.append(tdDate, tdContract, tdType, tdClient, tdAmt);
        bankTable.tBodies[0].appendChild(tr);

        monthTotal += Number(r.amount)||0;
      }

      bankTable.tBodies[0].appendChild(bankMonthTotalRow(monthTotal));
      grand += monthTotal;
      if (gi < keys.length - 1) bankTable.tBodies[0].appendChild(bankSpacerRow());
    }

    if (rows.length) {
      bankTable.tBodies[0].appendChild(bankSpacerRow());
      bankTable.tBodies[0].appendChild(bankGrandTotalRow(grand));
    }

    bankSelectedKey = null;
    bankSelectedEl.textContent = "Selected: —";
    bankRowCountEl.textContent = `Rows: ${rows.length}`;
  }

  function selectBankKey(key) {
    if (!bankTable) return;
    const k = normalizeText(key);
    bankSelectedKey = k || null;
    Array.from(bankTable.tBodies[0].rows).forEach(r=>r.classList.remove("is-selected"));
    if (!k) { bankSelectedEl.textContent="Selected: —"; return; }
    const tr = bankTable.tBodies[0].querySelector(`tr[data-row-type="data"][data-key="${CSS.escape(k)}"]`);
    if (tr) {
      tr.classList.add("is-selected");
      bankSelectedEl.textContent = `Selected: ${tr.cells[3]?.innerText || "—"}`;
    } else bankSelectedEl.textContent="Selected: —";
  }

  bankTable?.addEventListener("click",(e)=>{
    const tr = e.target.closest("tr");
    if (!tr || tr.parentElement?.tagName !== "TBODY") return;
    if ((tr.dataset.rowType || "data") !== "data") return;
    selectBankKey(tr.dataset.key);
  });
  bankTable?.addEventListener("dblclick",(e)=>{
    const tr = e.target.closest("tr");
    if (!tr || tr.parentElement?.tagName !== "TBODY") return;
    if ((tr.dataset.rowType || "data") !== "data") return;
    selectBankKey(tr.dataset.key);
    openBankModal("edit", tr.dataset.key);
  });

  btnBankAdd?.addEventListener("click",()=>openBankModal("add"));
  btnBankEdit?.addEventListener("click",()=>{
    if (!bankSelectedKey) return alert("Please select a Bank Received row first.");
    openBankModal("edit", bankSelectedKey);
  });

  bankForm?.addEventListener("submit",(e)=>{
    e.preventDefault();
    const date = brYyyyMMddToMMDDYYYY(brDate.value) || "";
    const contract = String(brContract.value || "").trim();
    const type = normalizeBankType(brType.value);
    const client = String(brClient.value || "").trim();
    const amount = parseMoneyBR(brAmount.value);
    if (!client) return alert("Name of Client is required.");

    if (bankMode === "add") {
      const entry = ensureBankId({ date, contract, type, client, amount });
      DB.saveBankReceived(entry).then(saved => { if (saved) entry.id = saved.id; });
      bankStore.push(entry);
      closeBankModal();
      renderBankTable();
      selectBankKey(bankKeyFor(entry));
      return;
    }

    const key = bankEditingKey;
    if (!key) return;
    const idx = bankStore.findIndex(x => bankKeyFor(x) === key);
    if (idx < 0) return;

    const updatedBR = ensureBankId({ date, contract, type, client, amount, _id: bankStore[idx]._id, id: bankStore[idx].id });
    bankStore[idx] = updatedBR;
    DB.saveBankReceived(updatedBR);
    closeBankModal();
    renderBankTable();
    selectBankKey(bankKeyFor(updatedBR));
  });

  btnBankDel?.addEventListener("click",()=>{
    if (!bankSelectedKey) return alert("Please select a Bank Received row first.");
    if (!confirm("Delete the selected Bank Received entry?")) return;
    const rowBR = bankStore.find(x => bankKeyFor(x) === bankSelectedKey);
    if (rowBR?.id) DB.deleteBankReceived(rowBR.id);
    bankStore = bankStore.filter(x => bankKeyFor(x) !== bankSelectedKey);
    renderBankTable();
  });

  // Import / Export
  btnBankImport?.addEventListener("click",()=>fileBankImport?.click());

  function buildBankHeaderIndex(headers) {
    const norms = headers.map(h => String(h ?? "").trim().toLowerCase().replace(/\s+/g," "));
    const find = (...aliases) => {
      for (const a of aliases) {
        const i = norms.indexOf(String(a).toLowerCase());
        if (i >= 0) return i;
      }
      return -1;
    };
    return {
      date: find("date"),
      contract: find("contract #","contract#","contract number","contract no","contract"),
      type: find("gcash/bank transfer","type","payment type","method"),
      client: find("name of client","client","name","customer","customer name"),
      amount: find("amount","amt","value")
    };
  }
  function detectHeaderRowBR(rows) {
    for (let r=0; r<Math.min(rows.length,6); r++) {
      const headers = rows[r].map(x => String(x ?? "").trim()).filter(Boolean);
      if (headers.length < 2) continue;
      const map = buildBankHeaderIndex(headers);
      let score = 0;
      if (map.date >= 0) score++;
      if (map.client >= 0) score++;
      if (map.amount >= 0) score++;
      if (score >= 2) return r;
    }
    return -1;
  }

  async function importBankRows(rows2d) {
    const headerRowIndex = detectHeaderRowBR(rows2d);
    let headers = null;
    let dataRows = null;
    if (headerRowIndex >= 0) {
      headers = rows2d[headerRowIndex].map(x => String(x ?? "").trim());
      dataRows = rows2d.slice(headerRowIndex + 1);
    } else {
      headers = ["Date","Contract #","Gcash/Bank Transfer","Name of Client","Amount"];
      dataRows = rows2d;
    }
    dataRows = dataRows.filter(r => r && r.some(v => String(v ?? "").trim() !== ""));
    if (!dataRows.length) { alert("No data rows found."); return; }

    const map = buildBankHeaderIndex(headers);
    const replace = confirm("Replace current Bank Received entries with imported data?\nOK = Replace\nCancel = Append");
    if (replace) bankStore = [];

    let added=0, skipped=0;
    for (const row of dataRows) {
      const date = (function(v){
        if (v==null || v==="") return "";
        if (v instanceof Date) return formatMMDDYYYY(v);
        if (typeof v === "number") {
          const utc_days = Math.floor(v - 25569);
          const d = new Date(utc_days * 86400 * 1000);
          return Number.isFinite(d.getTime()) ? formatMMDDYYYY(d) : "";
        }
        const s = String(v).trim();
        if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return yyyyMMddToMMDDYYYY(s);
        if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(s)) return `${pad2(s.split("/")[0])}/${pad2(s.split("/")[1])}/${s.split("/")[2]}`;
        const d = new Date(s);
        return Number.isFinite(d.getTime()) ? formatMMDDYYYY(d) : s;
      })( row[map.date >=0 ? map.date : 0] );

      const contract = String(row[map.contract>=0?map.contract:1] ?? "").trim();
      const type = normalizeBankType(String(row[map.type>=0?map.type:2] ?? ""));
      const client = String(row[map.client>=0?map.client:3] ?? "").trim();
      const amount = parseMoneyBR(row[map.amount>=0?map.amount:4]);

      if (!client && !amount) { skipped++; continue; }
      bankStore.push(ensureBankId({ date, contract, type, client, amount }));
      added++;
    }

    saveBankStore();
    renderBankTable();
    alert(`Import complete.\nAdded: ${added}\nSkipped: ${skipped}`);
  }

  fileBankImport?.addEventListener("change", async ()=>{
    const file = fileBankImport.files?.[0];
    if (!file) return;
    try {
      const name = (file.name||"").toLowerCase();
      if (name.endsWith(".csv")) {
        const text = await file.text();
        await importBankRows(parseCSVTo2D(text));
      } else if (name.endsWith(".xlsx")) {
        if (typeof XLSX === "undefined") { alert("XLSX library not loaded. Import a CSV instead."); return; }
        const buf = await file.arrayBuffer();
        const wb = XLSX.read(buf, { type:"array", cellDates:true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { header:1, defval:"" });
        await importBankRows(rows);
      } else {
        alert("Unsupported file type. Please select a .csv or .xlsx file.");
      }
    } catch(err) {
      alert("Import failed: " + (err?.message || String(err)));
    } finally {
      fileBankImport.value = "";
    }
  });

  btnBankExport?.addEventListener("click",()=>{
    const cols = ["Date","Contract #","Gcash/Bank Transfer","Name of Client","Amount"];
    const msoMoney = 'mso-number-format:"\\#\\,\\#\\#0\\.00"';
    const msoText = 'mso-number-format:"\\@"';
    const esc = (s)=>String(s??"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");

    const rows = getFilteredBankRows().sort((a,b)=>{
      const ak = monthKeyFromDate(a.date);
      const bk = monthKeyFromDate(b.date);
      if (ak !== bk) {
        if (ak === "unknown") return 1;
        if (bk === "unknown") return -1;
        return ak.localeCompare(bk);
      }
      const ad = parseMMDDYYYY(a.date)?.getTime() ?? Number.POSITIVE_INFINITY;
      const bd = parseMMDDYYYY(b.date)?.getTime() ?? Number.POSITIVE_INFINITY;
      if (ad !== bd) return ad - bd;
      return (a.client||"").localeCompare(b.client||"");
    });

    const groups = new Map();
    for (const r of rows) {
      const k = monthKeyFromDate(r.date);
      if (!groups.has(k)) groups.set(k, []);
      groups.get(k).push(r);
    }
    const keys = Array.from(groups.keys()).sort((a,b)=>{
      if (a === "unknown" && b !== "unknown") return 1;
      if (b === "unknown" && a !== "unknown") return -1;
      return a.localeCompare(b);
    });

    const css = `
      <style>
        table{ border-collapse:collapse; font-family:Calibri,Arial,sans-serif; font-size:11pt; }
        th,td{ border:1px solid #cfcfcf; padding:4px 6px; white-space:nowrap; }
        th{ font-weight:700; text-align:center; background:#fff; }
        td.num{ text-align:right; }
        tr.month-header td{ font-weight:800; text-align:center; background:#d9d9d9; border-top:2px solid #000; border-bottom:2px solid #000; }
        tr.month-total td{ font-weight:800; background:#fff2cc; border-top:2px solid #000; }
        tr.month-total td.label{ text-align:center; letter-spacing:.10em; }
        tr.grand-total td{ font-weight:900; background:#92d050; border-top:2px solid #000; border-bottom:2px solid #000; }
        tr.grand-total td.label{ text-align:center; letter-spacing:.10em; }
        tr.spacer td{ border:none; height:10px; }
      </style>
    `;

    let html = `<!doctype html><html><head><meta charset="utf-8" />${css}</head><body><table><tbody>`;
    let grand = 0;

    for (let gi=0; gi<keys.length; gi++) {
      const key = keys[gi];
      const grp = groups.get(key) || [];
      if (!grp.length) continue;

      html += `<tr class="month-header"><td colspan="5" style="${msoText}">${esc(monthLabelFromKey(key).toUpperCase())}</td></tr>`;
      html += `<tr>`;
      for (const h of cols) html += `<th>${esc(h)}</th>`;
      html += `</tr>`;

      let monthTotal = 0;
      for (const r of grp) {
        html += `<tr>`;
        html += `<td style="${msoText}">${esc(r.date)}</td>`;
        html += `<td style="${msoText}">${esc(r.contract||"")}</td>`;
        html += `<td style="${msoText}">${esc(normalizeBankType(r.type))}</td>`;
        html += `<td style="${msoText}">${esc(r.client||"")}</td>`;
        html += `<td class="num" style="${msoMoney}">${fmtMoney(r.amount)}</td>`;
        html += `</tr>`;
        monthTotal += Number(r.amount)||0;
      }

      html += `<tr class="month-total"><td colspan="3"></td><td class="label">TOTAL</td><td class="num" style="${msoMoney}">${fmtMoney(monthTotal)}</td></tr>`;
      grand += monthTotal;
      if (gi < keys.length - 1) html += `<tr class="spacer"><td colspan="5"></td></tr>`;
    }

    if (rows.length) {
      html += `<tr class="spacer"><td colspan="5"></td></tr>`;
      html += `<tr class="grand-total"><td colspan="3"></td><td class="label">GRAND TOTAL</td><td class="num" style="${msoMoney}">${fmtMoney(grand)}</td></tr>`;
    }

    html += `</tbody></table></body></html>`;
    const blob = new Blob([html], { type:"application/vnd.ms-excel;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href=url;
    a.download=`Magallanes_BankReceived_${new Date().toISOString().slice(0,10)}.xls`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  });

  btnBankRefresh?.addEventListener("click",()=>{ renderBankTable(); alert("Refreshed Bank Received view."); });

  function initBankFromDOM() {
    if (!bankTable) return [];
    const trs = Array.from(bankTable.tBodies[0].rows).filter(r => (r.dataset.rowType || "data") === "data");
    return trs.map(tr => ensureBankId({
      date: String(tr.cells[0]?.innerText||"").trim(),
      contract: String(tr.cells[1]?.innerText||"").trim(),
      type: normalizeBankType(String(tr.cells[2]?.innerText||"")),
      client: String(tr.cells[3]?.innerText||"").trim(),
      amount: parseMoneyBR(String(tr.cells[4]?.innerText||""))
    })).filter(r => r.client || r.amount || r.date);
  }

  // Load bank received from Supabase
  DB.getBankReceived().then(rows => {
    bankStore = rows;
    renderBankTable();
  });

  renderBankTable();


  // ---------------------------
  // Bank Expense (Transactions) — full functionality + monthly grouping (v36)
  // ---------------------------
  const bankExpTable = $("#bankExpenseTable");
  const bankExpRowCountEl = $("#bankExpenseRowCount");
  const bankExpSelectedEl = $("#bankExpenseSelected");

  const btnBankExpAdd = $("#btnBankExpAddEntry");
  const btnBankExpEdit = $("#btnBankExpEditSelected");
  const btnBankExpDel = $("#btnBankExpDeleteSelected");
  const btnBankExpImport = $("#btnBankExpImportExcel");
  const btnBankExpExport = $("#btnBankExpExportExcel");
  const btnBankExpRefresh = $("#btnBankExpRefresh");
  const fileBankExpImport = $("#fileBankExpImport");

  const bankExpOverlay = $("#bankExpOverlay");
  const bankExpModal = $("#bankExpModal");
  const bankExpTitle = $("#bankExpTitle");
  const bankExpSubtitle = $("#bankExpSubtitle");
  const bankExpForm = $("#bankExpForm");
  const btnCloseBankExp = $("#btnCloseBankExp");
  const btnCancelBankExp = $("#btnCancelBankExp");
  const btnSubmitBankExp = $("#btnSubmitBankExp");

  const beDate = $("#beDate");
  const beCV = $("#beCV");
  const beCheck = $("#beCheck");
  const beParticular = $("#beParticular");
  const beWithdraw = $("#beWithdraw");

  const BANK_EXP_STORE_KEY = "mf_bank_expense_store_v36";

  /** @type {Array<{date:string, cv:string, check:string, particular:string, withdraw:number, _id?:string}>} */
  let bankExpStore = [];
  let bankExpSelectedKey = null;
  let bankExpMode = "add";
  let bankExpEditingKey = null;

  function ensureBankExpId(r) {
    if (!r._id) r._id = "__AUTO__" + Date.now() + "_" + Math.floor(Math.random()*100000);
    return r;
  }
  function bankExpKeyFor(r) { return normalizeText(r._id || ""); }

  function saveBankExpStore() { /* no-op */ }
  function loadBankExpStore() { return null; /* loaded async */ }

  function beYyyyMMddToMMDDYYYY(yyyyMMdd) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(yyyyMMdd || ""));
    return m ? `${m[2]}/${m[3]}/${m[1]}` : "";
  }
  function beMMDDYYYYToYyyyMMdd(mmddyyyy) {
    const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(String(mmddyyyy || "").trim());
    if (!m) return "";
    const mm = String(m[1]).padStart(2,"0");
    const dd = String(m[2]).padStart(2,"0");
    return `${m[3]}-${mm}-${dd}`;
  }
  function parseMoneyBE(s) {
    const n = Number(String(s ?? "").replace(/[^0-9.-]/g, ""));
    return Number.isFinite(n) ? n : 0;
  }

  // Filters
  const bankExpFilterCategory = $("#bankExpFilterCategory");
  const bankExpFilterInputs = $("#bankExpFilterInputs");
  const btnBankExpApplyFilter = $("#btnBankExpApplyFilter");
  const btnBankExpClearFilter = $("#btnBankExpClearFilter");
  let bankExpActiveFilter = { category:"", value:null };

  function setBankExpFilterInputs(category) {
    if (!bankExpFilterInputs) return;
    bankExpFilterInputs.innerHTML = "";

    const makeText = (id, label, placeholder="") => {
      const wrap = document.createElement("label");
      wrap.className = "field inline";
      wrap.innerHTML = `<span>${label}</span><input class="input" id="${id}" type="text" placeholder="${placeholder}"/>`;
      bankExpFilterInputs.appendChild(wrap);
    };
    const makeNumber = (id, label, placeholder="") => {
      const wrap = document.createElement("label");
      wrap.className = "field inline";
      wrap.innerHTML = `<span>${label}</span><input class="input" id="${id}" type="number" step="0.01" placeholder="${placeholder}"/>`;
      bankExpFilterInputs.appendChild(wrap);
    };

    if (!category) return;
    if (category === "date") return makeText("bankExpFilterDate", "Date", "MM/DD/YYYY or MM/YYYY or YYYY");
    if (category === "cv") return makeText("bankExpFilterCV", "CV", "contains...");
    if (category === "check") return makeText("bankExpFilterCheck", "Check #", "contains...");
    if (category === "particular") return makeText("bankExpFilterParticular", "Particular", "contains...");
    if (category === "withdraw") { makeNumber("bankExpFilterMin","Min","0.00"); makeNumber("bankExpFilterMax","Max","10000.00"); }
  }
  bankExpFilterCategory?.addEventListener("change", () => setBankExpFilterInputs(bankExpFilterCategory.value));

  function parseFilterDatePatternBE(s) {
    const t = String(s || "").trim();
    if (!t) return null;
    if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(t)) {
      const d = parseMMDDYYYY(t);
      return d ? { type:"full", d: formatMMDDYYYY(d) } : null;
    }
    if (/^\d{1,2}\/\d{4}$/.test(t)) {
      const [mm, yyyy] = t.split("/");
      return { type:"monthYear", mm: pad2(mm), yyyy };
    }
    if (/^\d{4}$/.test(t)) return { type:"year", yyyy:t };
    return { type:"contains", text: normalizeText(t) };
  }
  function rowMatchesDateFilterBE(rowDate, pattern) {
    if (!pattern) return true;
    const rd = String(rowDate || "").trim();
    if (!rd) return false;
    if (pattern.type === "full") return rd === pattern.d;
    if (pattern.type === "monthYear") {
      const m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(rd);
      return m ? (m[1]===pattern.mm && m[3]===pattern.yyyy) : false;
    }
    if (pattern.type === "year") {
      const m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(rd);
      return m ? (m[3]===pattern.yyyy) : false;
    }
    if (pattern.type === "contains") return normalizeText(rd).includes(pattern.text);
    return true;
  }
  function getBankExpFilterFromUI() {
    const cat = bankExpFilterCategory?.value || "";
    if (!cat) return { category:"", value:null };
    if (cat === "date") return { category:"date", value: parseFilterDatePatternBE($("#bankExpFilterDate")?.value || "") };
    if (cat === "cv") return { category:"cv", value: normalizeText($("#bankExpFilterCV")?.value || "") };
    if (cat === "check") return { category:"check", value: normalizeText($("#bankExpFilterCheck")?.value || "") };
    if (cat === "particular") return { category:"particular", value: normalizeText($("#bankExpFilterParticular")?.value || "") };
    if (cat === "withdraw") {
      const minv = Number($("#bankExpFilterMin")?.value || "");
      const maxv = Number($("#bankExpFilterMax")?.value || "");
      return { category:"withdraw", value:{ min: Number.isFinite(minv)?minv:null, max: Number.isFinite(maxv)?maxv:null } };
    }
    return { category:"", value:null };
  }
  function bankExpRowMatchesFilter(r, f) {
    if (!f || !f.category) return true;
    if (f.category === "date") return rowMatchesDateFilterBE(r.date, f.value);
    if (f.category === "cv") return !f.value || normalizeText(r.cv).includes(f.value);
    if (f.category === "check") return !f.value || normalizeText(r.check).includes(f.value);
    if (f.category === "particular") return !f.value || normalizeText(r.particular).includes(f.value);
    if (f.category === "withdraw") {
      const a = Number(r.withdraw)||0;
      const minv = f.value?.min; const maxv = f.value?.max;
      if (minv != null && a < minv) return false;
      if (maxv != null && a > maxv) return false;
      return true;
    }
    return true;
  }
  function getFilteredBankExpRows() { return bankExpStore.filter(r => bankExpRowMatchesFilter(r, bankExpActiveFilter)); }
  btnBankExpApplyFilter?.addEventListener("click", () => { bankExpActiveFilter = getBankExpFilterFromUI(); renderBankExpTable(); });
  btnBankExpClearFilter?.addEventListener("click", () => {
    bankExpActiveFilter = { category:"", value:null };
    if (bankExpFilterCategory) bankExpFilterCategory.value = "";
    setBankExpFilterInputs("");
    renderBankExpTable();
  });
  setBankExpFilterInputs(bankExpFilterCategory?.value || "");

  function openBankExpModal(mode, keyOrNull=null) {
    if (!bankExpOverlay || !bankExpModal) { alert("Bank Expense form not available."); return; }
    bankExpMode = mode; bankExpEditingKey = null;
    if (mode === "add") {
      bankExpTitle.textContent = "Add Entry";
      bankExpSubtitle.textContent = "Enter bank expense details.";
      btnSubmitBankExp.textContent = "Add Entry";
      beDate.value = new Date().toISOString().slice(0,10);
      beCV.value = "";
      beCheck.value = "";
      beParticular.value = "";
      beWithdraw.value = "0.00";
    } else {
      bankExpTitle.textContent = "Edit Entry";
      bankExpSubtitle.textContent = "Update the selected bank expense entry.";
      btnSubmitBankExp.textContent = "Save Changes";
      const key = normalizeText(keyOrNull);
      const found = bankExpStore.find(x => bankExpKeyFor(x) === key);
      if (!found) { alert("Could not find the selected entry."); return; }
      bankExpEditingKey = key;
      beDate.value = beMMDDYYYYToYyyyMMdd(found.date) || new Date().toISOString().slice(0,10);
      beCV.value = found.cv || "";
      beCheck.value = found.check || "";
      beParticular.value = found.particular || "";
      beWithdraw.value = (Number(found.withdraw)||0).toFixed(2);
    }
    bankExpOverlay.classList.add("is-open");
    bankExpModal.classList.add("is-open");
    setTimeout(() => beParticular.focus(), 0);
  }
  function closeBankExpModal() {
    bankExpOverlay?.classList.remove("is-open");
    bankExpModal?.classList.remove("is-open");
    bankExpMode = "add"; bankExpEditingKey = null;
  }
  btnCloseBankExp?.addEventListener("click", closeBankExpModal);
  btnCancelBankExp?.addEventListener("click", closeBankExpModal);
  bankExpOverlay?.addEventListener("click", closeBankExpModal);
  document.addEventListener("keydown",(e)=>{
    if (e.key === "Escape" && bankExpModal?.classList.contains("is-open")) closeBankExpModal();
  });

  function bankExpMonthHeaderRow(label) {
    const tr = document.createElement("tr");
    tr.classList.add("bankExpMonthHeader");
    tr.dataset.rowType = "bankExpMonthHeader";
    const td = document.createElement("td");
    td.colSpan = 5;
    td.textContent = String(label||"").toUpperCase();
    tr.appendChild(td);
    return tr;
  }
  function bankExpMonthTotalRow(total) {
    const tr = document.createElement("tr");
    tr.classList.add("bankExpMonthTotal");
    tr.dataset.rowType = "bankExpMonthTotal";
    // Date blank, CV blank, Check blank, label under Particular, withdraw at end
    for (let i=0;i<3;i++) tr.appendChild(document.createElement("td"));
    const label = document.createElement("td"); label.textContent="TOTAL"; label.classList.add("label");
    tr.appendChild(label);
    const amt = document.createElement("td"); amt.textContent=fmtMoney(total); amt.classList.add("num");
    tr.appendChild(amt);
    return tr;
  }
  function bankExpGrandTotalRow(total) {
    const tr = document.createElement("tr");
    tr.classList.add("bankExpGrandTotal");
    tr.dataset.rowType = "bankExpGrandTotal";
    for (let i=0;i<3;i++) tr.appendChild(document.createElement("td"));
    const label = document.createElement("td"); label.textContent="GRAND TOTAL"; label.classList.add("label");
    tr.appendChild(label);
    const amt = document.createElement("td"); amt.textContent=fmtMoney(total); amt.classList.add("num");
    tr.appendChild(amt);
    return tr;
  }
  function bankExpSpacerRow() {
    const tr = document.createElement("tr");
    tr.classList.add("spacer");
    tr.dataset.rowType = "spacer";
    const td = document.createElement("td"); td.colSpan=5; td.textContent="";
    tr.appendChild(td);
    return tr;
  }

  function renderBankExpTable() {
    if (!bankExpTable) return;
    const rows = getFilteredBankExpRows().sort((a,b)=>{
      const ak = monthKeyFromDate(a.date);
      const bk = monthKeyFromDate(b.date);
      if (ak !== bk) {
        if (ak === "unknown") return 1;
        if (bk === "unknown") return -1;
        return ak.localeCompare(bk);
      }
      const ad = parseMMDDYYYY(a.date)?.getTime() ?? Number.POSITIVE_INFINITY;
      const bd = parseMMDDYYYY(b.date)?.getTime() ?? Number.POSITIVE_INFINITY;
      if (ad !== bd) return ad - bd;
      return (a.particular||"").localeCompare(b.particular||"");
    });

    const groups = new Map();
    for (const r of rows) {
      const k = monthKeyFromDate(r.date);
      if (!groups.has(k)) groups.set(k, []);
      groups.get(k).push(r);
    }
    const keys = Array.from(groups.keys()).sort((a,b)=>{
      if (a === "unknown" && b !== "unknown") return 1;
      if (b === "unknown" && a !== "unknown") return -1;
      return a.localeCompare(b);
    });

    bankExpTable.tBodies[0].innerHTML = "";
    let grand = 0;

    for (let gi=0; gi<keys.length; gi++) {
      const key = keys[gi];
      const grp = groups.get(key) || [];
      if (!grp.length) continue;

      bankExpTable.tBodies[0].appendChild(bankExpMonthHeaderRow(monthLabelFromKey(key)));
      let monthTotal = 0;

      for (const r of grp) {
        ensureBankExpId(r);
        const tr = document.createElement("tr");
        tr.dataset.rowType="data";
        tr.dataset.key=bankExpKeyFor(r);

        const tdDate = document.createElement("td"); tdDate.textContent=r.date;
        const tdCV = document.createElement("td"); tdCV.textContent=r.cv || "";
        const tdCheck = document.createElement("td"); tdCheck.textContent=r.check || "";
        const tdPart = document.createElement("td"); tdPart.textContent=r.particular || "";
        const tdAmt = document.createElement("td"); tdAmt.textContent=fmtMoney(r.withdraw); tdAmt.classList.add("num");

        tr.append(tdDate, tdCV, tdCheck, tdPart, tdAmt);
        bankExpTable.tBodies[0].appendChild(tr);

        monthTotal += Number(r.withdraw)||0;
      }

      bankExpTable.tBodies[0].appendChild(bankExpMonthTotalRow(monthTotal));
      grand += monthTotal;
      if (gi < keys.length - 1) bankExpTable.tBodies[0].appendChild(bankExpSpacerRow());
    }

    if (rows.length) {
      bankExpTable.tBodies[0].appendChild(bankExpSpacerRow());
      bankExpTable.tBodies[0].appendChild(bankExpGrandTotalRow(grand));
    }

    bankExpSelectedKey = null;
    bankExpSelectedEl.textContent = "Selected: —";
    bankExpRowCountEl.textContent = `Rows: ${rows.length}`;
  }

  function selectBankExpKey(key) {
    if (!bankExpTable) return;
    const k = normalizeText(key);
    bankExpSelectedKey = k || null;
    Array.from(bankExpTable.tBodies[0].rows).forEach(r=>r.classList.remove("is-selected"));
    if (!k) { bankExpSelectedEl.textContent="Selected: —"; return; }
    const tr = bankExpTable.tBodies[0].querySelector(`tr[data-row-type="data"][data-key="${CSS.escape(k)}"]`);
    if (tr) {
      tr.classList.add("is-selected");
      bankExpSelectedEl.textContent = `Selected: ${tr.cells[3]?.innerText || "—"}`;
    } else bankExpSelectedEl.textContent="Selected: —";
  }

  bankExpTable?.addEventListener("click",(e)=>{
    const tr = e.target.closest("tr");
    if (!tr || tr.parentElement?.tagName !== "TBODY") return;
    if ((tr.dataset.rowType || "data") !== "data") return;
    selectBankExpKey(tr.dataset.key);
  });
  bankExpTable?.addEventListener("dblclick",(e)=>{
    const tr = e.target.closest("tr");
    if (!tr || tr.parentElement?.tagName !== "TBODY") return;
    if ((tr.dataset.rowType || "data") !== "data") return;
    selectBankExpKey(tr.dataset.key);
    openBankExpModal("edit", tr.dataset.key);
  });

  btnBankExpAdd?.addEventListener("click",()=>openBankExpModal("add"));
  btnBankExpEdit?.addEventListener("click",()=>{
    if (!bankExpSelectedKey) return alert("Please select a Bank Expense row first.");
    openBankExpModal("edit", bankExpSelectedKey);
  });

  bankExpForm?.addEventListener("submit",(e)=>{
    e.preventDefault();
    const date = beYyyyMMddToMMDDYYYY(beDate.value) || "";
    const cv = String(beCV.value || "").trim();
    const check = String(beCheck.value || "").trim();
    const particular = String(beParticular.value || "").trim();
    const withdraw = parseMoneyBE(beWithdraw.value);

    if (!particular) return alert("Particular is required.");

    if (bankExpMode === "add") {
      const entry = ensureBankExpId({ date, cv, check, particular, withdraw });
      DB.saveBankExpense(entry).then(saved => { if (saved) entry.id = saved.id; });
      bankExpStore.push(entry);
      closeBankExpModal();
      renderBankExpTable();
      selectBankExpKey(bankExpKeyFor(entry));
      return;
    }

    const key = bankExpEditingKey;
    if (!key) return;
    const idx = bankExpStore.findIndex(x => bankExpKeyFor(x) === key);
    if (idx < 0) return;

    const updatedBE = ensureBankExpId({ date, cv, check, particular, withdraw, _id: bankExpStore[idx]._id, id: bankExpStore[idx].id });
    bankExpStore[idx] = updatedBE;
    DB.saveBankExpense(updatedBE);
    closeBankExpModal();
    renderBankExpTable();
    selectBankExpKey(bankExpKeyFor(updatedBE));
  });

  btnBankExpDel?.addEventListener("click",()=>{
    if (!bankExpSelectedKey) return alert("Please select a Bank Expense row first.");
    if (!confirm("Delete the selected Bank Expense entry?")) return;
    const rowBE = bankExpStore.find(x => bankExpKeyFor(x) === bankExpSelectedKey);
    if (rowBE?.id) DB.deleteBankExpense(rowBE.id);
    bankExpStore = bankExpStore.filter(x => bankExpKeyFor(x) !== bankExpSelectedKey);
    renderBankExpTable();
  });

  btnBankExpImport?.addEventListener("click",()=>fileBankExpImport?.click());

  function buildBankExpHeaderIndex(headers) {
    const norms = headers.map(h => String(h ?? "").trim().toLowerCase().replace(/\s+/g," "));
    const find = (...aliases) => {
      for (const a of aliases) {
        const i = norms.indexOf(String(a).toLowerCase());
        if (i >= 0) return i;
      }
      return -1;
    };
    return {
      date: find("date"),
      cv: find("cv","cv #","cv#","voucher","voucher #"),
      check: find("check #","check#","cheque","cheque #"),
      particular: find("particular","description","remarks","memo"),
      withdraw: find("withdraw","withdrawal","amount","amt","value")
    };
  }
  function detectHeaderRowBE(rows) {
    for (let r=0; r<Math.min(rows.length,6); r++) {
      const headers = rows[r].map(x => String(x ?? "").trim()).filter(Boolean);
      if (headers.length < 2) continue;
      const map = buildBankExpHeaderIndex(headers);
      let score = 0;
      if (map.date >= 0) score++;
      if (map.particular >= 0) score++;
      if (map.withdraw >= 0) score++;
      if (score >= 2) return r;
    }
    return -1;
  }

  async function importBankExpRows(rows2d) {
    const headerRowIndex = detectHeaderRowBE(rows2d);
    let headers = null;
    let dataRows = null;

    if (headerRowIndex >= 0) {
      headers = rows2d[headerRowIndex].map(x => String(x ?? "").trim());
      dataRows = rows2d.slice(headerRowIndex + 1);
    } else {
      headers = ["Date","CV","Check #","Particular","Withdraw"];
      dataRows = rows2d;
    }
    dataRows = dataRows.filter(r => r && r.some(v => String(v ?? "").trim() !== ""));
    if (!dataRows.length) { alert("No data rows found."); return; }

    const map = buildBankExpHeaderIndex(headers);
    const replace = confirm("Replace current Bank Expense entries with imported data?\nOK = Replace\nCancel = Append");
    if (replace) bankExpStore = [];

    let added=0, skipped=0;
    for (const row of dataRows) {
      const date = (function(v){
        if (v==null || v==="") return "";
        if (v instanceof Date) return formatMMDDYYYY(v);
        if (typeof v === "number") {
          const utc_days = Math.floor(v - 25569);
          const d = new Date(utc_days * 86400 * 1000);
          return Number.isFinite(d.getTime()) ? formatMMDDYYYY(d) : "";
        }
        const s = String(v).trim();
        if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return yyyyMMddToMMDDYYYY(s);
        if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(s)) return `${pad2(s.split("/")[0])}/${pad2(s.split("/")[1])}/${s.split("/")[2]}`;
        const d = new Date(s);
        return Number.isFinite(d.getTime()) ? formatMMDDYYYY(d) : s;
      })( row[map.date >= 0 ? map.date : 0] );

      const cv = String(row[map.cv>=0?map.cv:1] ?? "").trim();
      const check = String(row[map.check>=0?map.check:2] ?? "").trim();
      const particular = String(row[map.particular>=0?map.particular:3] ?? "").trim();
      const withdraw = parseMoneyBE(row[map.withdraw>=0?map.withdraw:4]);

      if (!particular && !withdraw && !date) { skipped++; continue; }
      if (!particular) { skipped++; continue; }

      bankExpStore.push(ensureBankExpId({ date, cv, check, particular, withdraw }));
      added++;
    }

    saveBankExpStore();
    renderBankExpTable();
    alert(`Import complete.\nAdded: ${added}\nSkipped: ${skipped}`);
  }

  fileBankExpImport?.addEventListener("change", async ()=>{
    const file = fileBankExpImport.files?.[0];
    if (!file) return;
    try {
      const name = (file.name||"").toLowerCase();
      if (name.endsWith(".csv")) {
        const text = await file.text();
        await importBankExpRows(parseCSVTo2D(text));
      } else if (name.endsWith(".xlsx")) {
        if (typeof XLSX === "undefined") { alert("XLSX library not loaded. Import a CSV instead."); return; }
        const buf = await file.arrayBuffer();
        const wb = XLSX.read(buf, { type:"array", cellDates:true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { header:1, defval:"" });
        await importBankExpRows(rows);
      } else {
        alert("Unsupported file type. Please select a .csv or .xlsx file.");
      }
    } catch(err) {
      alert("Import failed: " + (err?.message || String(err)));
    } finally {
      fileBankExpImport.value = "";
    }
  });

  btnBankExpExport?.addEventListener("click",()=>{
    const cols = ["Date","CV","Check #","Particular","Withdraw"];
    const msoMoney = 'mso-number-format:"\\#\\,\\#\\#0\\.00"';
    const msoText = 'mso-number-format:"\\@"';
    const esc = (s)=>String(s??"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");

    const rows = getFilteredBankExpRows().sort((a,b)=>{
      const ak = monthKeyFromDate(a.date);
      const bk = monthKeyFromDate(b.date);
      if (ak !== bk) {
        if (ak === "unknown") return 1;
        if (bk === "unknown") return -1;
        return ak.localeCompare(bk);
      }
      const ad = parseMMDDYYYY(a.date)?.getTime() ?? Number.POSITIVE_INFINITY;
      const bd = parseMMDDYYYY(b.date)?.getTime() ?? Number.POSITIVE_INFINITY;
      if (ad !== bd) return ad - bd;
      return (a.particular||"").localeCompare(b.particular||"");
    });

    const groups = new Map();
    for (const r of rows) {
      const k = monthKeyFromDate(r.date);
      if (!groups.has(k)) groups.set(k, []);
      groups.get(k).push(r);
    }
    const keys = Array.from(groups.keys()).sort((a,b)=>{
      if (a === "unknown" && b !== "unknown") return 1;
      if (b === "unknown" && a !== "unknown") return -1;
      return a.localeCompare(b);
    });

    const css = `
      <style>
        table{ border-collapse:collapse; font-family:Calibri,Arial,sans-serif; font-size:11pt; }
        th,td{ border:1px solid #cfcfcf; padding:4px 6px; white-space:nowrap; }
        th{ font-weight:700; text-align:center; background:#fff; }
        td.num{ text-align:right; }
        tr.month-header td{ font-weight:800; text-align:center; background:#d9d9d9; border-top:2px solid #000; border-bottom:2px solid #000; }
        tr.month-total td{ font-weight:800; background:#fff2cc; border-top:2px solid #000; }
        tr.month-total td.label{ text-align:center; letter-spacing:.10em; }
        tr.grand-total td{ font-weight:900; background:#92d050; border-top:2px solid #000; border-bottom:2px solid #000; }
        tr.grand-total td.label{ text-align:center; letter-spacing:.10em; }
        tr.spacer td{ border:none; height:10px; }
      </style>
    `;

    let html = `<!doctype html><html><head><meta charset="utf-8" />${css}</head><body><table><tbody>`;
    let grand = 0;

    for (let gi=0; gi<keys.length; gi++) {
      const key = keys[gi];
      const grp = groups.get(key) || [];
      if (!grp.length) continue;

      html += `<tr class="month-header"><td colspan="5" style="${msoText}">${esc(monthLabelFromKey(key).toUpperCase())}</td></tr>`;
      html += `<tr>`;
      for (const h of cols) html += `<th>${esc(h)}</th>`;
      html += `</tr>`;

      let monthTotal = 0;
      for (const r of grp) {
        html += `<tr>`;
        html += `<td style="${msoText}">${esc(r.date)}</td>`;
        html += `<td style="${msoText}">${esc(r.cv||"")}</td>`;
        html += `<td style="${msoText}">${esc(r.check||"")}</td>`;
        html += `<td style="${msoText}">${esc(r.particular||"")}</td>`;
        html += `<td class="num" style="${msoMoney}">${fmtMoney(r.withdraw)}</td>`;
        html += `</tr>`;
        monthTotal += Number(r.withdraw)||0;
      }

      html += `<tr class="month-total"><td colspan="3"></td><td class="label">TOTAL</td><td class="num" style="${msoMoney}">${fmtMoney(monthTotal)}</td></tr>`;
      grand += monthTotal;
      if (gi < keys.length - 1) html += `<tr class="spacer"><td colspan="5"></td></tr>`;
    }

    if (rows.length) {
      html += `<tr class="spacer"><td colspan="5"></td></tr>`;
      html += `<tr class="grand-total"><td colspan="3"></td><td class="label">GRAND TOTAL</td><td class="num" style="${msoMoney}">${fmtMoney(grand)}</td></tr>`;
    }

    html += `</tbody></table></body></html>`;
    const blob = new Blob([html], { type:"application/vnd.ms-excel;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href=url;
    a.download=`Magallanes_BankExpense_${new Date().toISOString().slice(0,10)}.xls`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  });

  btnBankExpRefresh?.addEventListener("click",()=>{ renderBankExpTable(); alert("Refreshed Bank Expense view."); });

  function initBankExpFromDOM() {
    if (!bankExpTable) return [];
    const trs = Array.from(bankExpTable.tBodies[0].rows).filter(r => (r.dataset.rowType || "data") === "data");
    return trs.map(tr => ensureBankExpId({
      date: String(tr.cells[0]?.innerText||"").trim(),
      cv: String(tr.cells[1]?.innerText||"").trim(),
      check: String(tr.cells[2]?.innerText||"").trim(),
      particular: String(tr.cells[3]?.innerText||"").trim(),
      withdraw: parseMoneyBE(String(tr.cells[4]?.innerText||""))
    })).filter(r => r.particular || r.withdraw || r.date);
  }

  // Load bank expense from Supabase
  DB.getBankExpense().then(rows => {
    bankExpStore = rows;
    renderBankExpTable();
  });

  renderBankExpTable();


  // ---------------------------
  // PNB Deposit (Transactions) — full functionality + monthly grouping (v38)
  // ---------------------------
  const pnbTable = $("#pnbDepositTable");
  const pnbRowCountEl = $("#pnbDepositRowCount");
  const pnbSelectedEl = $("#pnbDepositSelected");

  const btnPnbAdd = $("#btnPnbAddEntry");
  const btnPnbEdit = $("#btnPnbEditSelected");
  const btnPnbDel = $("#btnPnbDeleteSelected");
  const btnPnbImport = $("#btnPnbImportExcel");
  const btnPnbExport = $("#btnPnbExportExcel");
  const btnPnbRefresh = $("#btnPnbRefresh");
  const filePnbImport = $("#filePnbImport");

  const pnbOverlay = $("#pnbOverlay");
  const pnbModal = $("#pnbModal");
  const pnbTitle = $("#pnbTitle");
  const pnbSubtitle = $("#pnbSubtitle");
  const pnbForm = $("#pnbForm");
  const btnClosePnb = $("#btnClosePnb");
  const btnCancelPnb = $("#btnCancelPnb");
  const btnSubmitPnb = $("#btnSubmitPnb");

  const pdDate = $("#pdDate");
  const pdAmount = $("#pdAmount");

  const PNB_STORE_KEY = "mf_pnb_deposit_store_v38";

  /** @type {Array<{date:string, amount:number, _id?:string}>} */
  let pnbStore = [];
  let pnbSelectedKey = null;
  let pnbMode = "add";
  let pnbEditingKey = null;

  function ensurePnbId(r) {
    if (!r._id) r._id = "__AUTO__" + Date.now() + "_" + Math.floor(Math.random()*100000);
    return r;
  }
  function pnbKeyFor(r) { return normalizeText(r._id || ""); }

  function savePnbStore() { /* no-op */ }
  function loadPnbStore() { return null; /* loaded async */ }

  function pdYyyyMMddToMMDDYYYY(yyyyMMdd) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(yyyyMMdd || ""));
    return m ? `${m[2]}/${m[3]}/${m[1]}` : "";
  }
  function pdMMDDYYYYToYyyyMMdd(mmddyyyy) {
    const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(String(mmddyyyy || "").trim());
    if (!m) return "";
    const mm = String(m[1]).padStart(2,"0");
    const dd = String(m[2]).padStart(2,"0");
    return `${m[3]}-${mm}-${dd}`;
  }
  function parseMoneyPD(s) {
    const n = Number(String(s ?? "").replace(/[^0-9.-]/g, ""));
    return Number.isFinite(n) ? n : 0;
  }

  // Filters
  const pnbFilterCategory = $("#pnbFilterCategory");
  const pnbFilterInputs = $("#pnbFilterInputs");
  const btnPnbApplyFilter = $("#btnPnbApplyFilter");
  const btnPnbClearFilter = $("#btnPnbClearFilter");
  let pnbActiveFilter = { category:"", value:null };

  function setPnbFilterInputs(category) {
    if (!pnbFilterInputs) return;
    pnbFilterInputs.innerHTML = "";
    const makeText = (id, label, placeholder="") => {
      const wrap = document.createElement("label");
      wrap.className = "field inline";
      wrap.innerHTML = `<span>${label}</span><input class="input" id="${id}" type="text" placeholder="${placeholder}"/>`;
      pnbFilterInputs.appendChild(wrap);
    };
    const makeNumber = (id, label, placeholder="") => {
      const wrap = document.createElement("label");
      wrap.className = "field inline";
      wrap.innerHTML = `<span>${label}</span><input class="input" id="${id}" type="number" step="0.01" placeholder="${placeholder}"/>`;
      pnbFilterInputs.appendChild(wrap);
    };
    if (!category) return;
    if (category === "date") return makeText("pnbFilterDate","Date","MM/DD/YYYY or MM/YYYY or YYYY");
    if (category === "amount") { makeNumber("pnbFilterMin","Min","0.00"); makeNumber("pnbFilterMax","Max","10000.00"); }
  }
  pnbFilterCategory?.addEventListener("change", () => setPnbFilterInputs(pnbFilterCategory.value));

  function parseFilterDatePatternPD(s) {
    const t = String(s || "").trim();
    if (!t) return null;
    if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(t)) {
      const d = parseMMDDYYYY(t);
      return d ? { type:"full", d: formatMMDDYYYY(d) } : null;
    }
    if (/^\d{1,2}\/\d{4}$/.test(t)) {
      const [mm, yyyy] = t.split("/");
      return { type:"monthYear", mm: pad2(mm), yyyy };
    }
    if (/^\d{4}$/.test(t)) return { type:"year", yyyy:t };
    return { type:"contains", text: normalizeText(t) };
  }
  function rowMatchesDateFilterPD(rowDate, pattern) {
    if (!pattern) return true;
    const rd = String(rowDate || "").trim();
    if (!rd) return false;
    if (pattern.type === "full") return rd === pattern.d;
    if (pattern.type === "monthYear") {
      const m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(rd);
      return m ? (m[1]===pattern.mm && m[3]===pattern.yyyy) : false;
    }
    if (pattern.type === "year") {
      const m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(rd);
      return m ? (m[3]===pattern.yyyy) : false;
    }
    if (pattern.type === "contains") return normalizeText(rd).includes(pattern.text);
    return true;
  }
  function getPnbFilterFromUI() {
    const cat = pnbFilterCategory?.value || "";
    if (!cat) return { category:"", value:null };
    if (cat === "date") return { category:"date", value: parseFilterDatePatternPD($("#pnbFilterDate")?.value || "") };
    if (cat === "amount") {
      const minv = Number($("#pnbFilterMin")?.value || "");
      const maxv = Number($("#pnbFilterMax")?.value || "");
      return { category:"amount", value:{ min: Number.isFinite(minv)?minv:null, max: Number.isFinite(maxv)?maxv:null } };
    }
    return { category:"", value:null };
  }
  function pnbRowMatchesFilter(r, f) {
    if (!f || !f.category) return true;
    if (f.category === "date") return rowMatchesDateFilterPD(r.date, f.value);
    if (f.category === "amount") {
      const a = Number(r.amount)||0;
      const minv = f.value?.min; const maxv = f.value?.max;
      if (minv != null && a < minv) return false;
      if (maxv != null && a > maxv) return false;
      return true;
    }
    return true;
  }
  function getFilteredPnbRows() { return pnbStore.filter(r => pnbRowMatchesFilter(r, pnbActiveFilter)); }

  btnPnbApplyFilter?.addEventListener("click", () => { pnbActiveFilter = getPnbFilterFromUI(); renderPnbTable(); });
  btnPnbClearFilter?.addEventListener("click", () => {
    pnbActiveFilter = { category:"", value:null };
    if (pnbFilterCategory) pnbFilterCategory.value = "";
    setPnbFilterInputs("");
    renderPnbTable();
  });
  setPnbFilterInputs(pnbFilterCategory?.value || "");

  function openPnbModal(mode, keyOrNull=null) {
    if (!pnbOverlay || !pnbModal) { alert("PNB Deposit form not available."); return; }
    pnbMode = mode; pnbEditingKey = null;
    if (mode === "add") {
      pnbTitle.textContent = "Add Entry";
      pnbSubtitle.textContent = "Enter PNB deposit details.";
      btnSubmitPnb.textContent = "Add Entry";
      pdDate.value = new Date().toISOString().slice(0,10);
      pdAmount.value = "0.00";
    } else {
      pnbTitle.textContent = "Edit Entry";
      pnbSubtitle.textContent = "Update the selected PNB deposit entry.";
      btnSubmitPnb.textContent = "Save Changes";
      const key = normalizeText(keyOrNull);
      const found = pnbStore.find(x => pnbKeyFor(x) === key);
      if (!found) { alert("Could not find the selected entry."); return; }
      pnbEditingKey = key;
      pdDate.value = pdMMDDYYYYToYyyyMMdd(found.date) || new Date().toISOString().slice(0,10);
      pdAmount.value = (Number(found.amount)||0).toFixed(2);
    }
    pnbOverlay.classList.add("is-open");
    pnbModal.classList.add("is-open");
    setTimeout(() => pdAmount.focus(), 0);
  }
  function closePnbModal() {
    pnbOverlay?.classList.remove("is-open");
    pnbModal?.classList.remove("is-open");
    pnbMode = "add"; pnbEditingKey = null;
  }
  $("#btnClosePnb")?.addEventListener("click", closePnbModal);
  $("#btnCancelPnb")?.addEventListener("click", closePnbModal);
  pnbOverlay?.addEventListener("click", closePnbModal);
  document.addEventListener("keydown",(e)=>{
    if (e.key === "Escape" && pnbModal?.classList.contains("is-open")) closePnbModal();
  });

  function pnbMonthHeaderRow(label) {
    const tr = document.createElement("tr");
    tr.classList.add("pnbMonthHeader");
    tr.dataset.rowType = "pnbMonthHeader";
    const td = document.createElement("td");
    td.colSpan = 2;
    td.textContent = String(label||"").toUpperCase();
    tr.appendChild(td);
    return tr;
  }
  function pnbMonthTotalRow(total) {
    const tr = document.createElement("tr");
    tr.classList.add("pnbMonthTotal");
    tr.dataset.rowType = "pnbMonthTotal";
    const td1 = document.createElement("td"); td1.textContent="TOTAL"; td1.classList.add("label");
    const td2 = document.createElement("td"); td2.textContent=fmtMoney(total); td2.classList.add("num");
    tr.append(td1, td2);
    return tr;
  }
  function pnbGrandTotalRow(total) {
    const tr = document.createElement("tr");
    tr.classList.add("pnbGrandTotal");
    tr.dataset.rowType = "pnbGrandTotal";
    const td1 = document.createElement("td"); td1.textContent="GRAND TOTAL"; td1.classList.add("label");
    const td2 = document.createElement("td"); td2.textContent=fmtMoney(total); td2.classList.add("num");
    tr.append(td1, td2);
    return tr;
  }
  function pnbSpacerRow() {
    const tr = document.createElement("tr");
    tr.classList.add("spacer");
    tr.dataset.rowType = "spacer";
    const td = document.createElement("td"); td.colSpan=2; td.textContent="";
    tr.appendChild(td);
    return tr;
  }

  function renderPnbTable() {
    if (!pnbTable) return;
    const rows = getFilteredPnbRows().sort((a,b)=>{
      const ak = monthKeyFromDate(a.date);
      const bk = monthKeyFromDate(b.date);
      if (ak !== bk) {
        if (ak === "unknown") return 1;
        if (bk === "unknown") return -1;
        return ak.localeCompare(bk);
      }
      const ad = parseMMDDYYYY(a.date)?.getTime() ?? Number.POSITIVE_INFINITY;
      const bd = parseMMDDYYYY(b.date)?.getTime() ?? Number.POSITIVE_INFINITY;
      if (ad !== bd) return ad - bd;
      return (Number(a.amount)||0) - (Number(b.amount)||0);
    });

    const groups = new Map();
    for (const r of rows) {
      const k = monthKeyFromDate(r.date);
      if (!groups.has(k)) groups.set(k, []);
      groups.get(k).push(r);
    }
    const keys = Array.from(groups.keys()).sort((a,b)=>{
      if (a === "unknown" && b !== "unknown") return 1;
      if (b === "unknown" && a !== "unknown") return -1;
      return a.localeCompare(b);
    });

    pnbTable.tBodies[0].innerHTML = "";
    let grand = 0;

    for (let gi=0; gi<keys.length; gi++) {
      const key = keys[gi];
      const grp = groups.get(key) || [];
      if (!grp.length) continue;

      pnbTable.tBodies[0].appendChild(pnbMonthHeaderRow(monthLabelFromKey(key)));
      let monthTotal = 0;

      for (const r of grp) {
        ensurePnbId(r);
        const tr = document.createElement("tr");
        tr.dataset.rowType = "data";
        tr.dataset.key = pnbKeyFor(r);

        const tdDate = document.createElement("td"); tdDate.textContent = r.date;
        const tdAmt = document.createElement("td"); tdAmt.textContent = fmtMoney(r.amount); tdAmt.classList.add("num");
        tr.append(tdDate, tdAmt);
        pnbTable.tBodies[0].appendChild(tr);

        monthTotal += Number(r.amount)||0;
      }

      pnbTable.tBodies[0].appendChild(pnbMonthTotalRow(monthTotal));
      grand += monthTotal;
      if (gi < keys.length - 1) pnbTable.tBodies[0].appendChild(pnbSpacerRow());
    }

    if (rows.length) {
      pnbTable.tBodies[0].appendChild(pnbSpacerRow());
      pnbTable.tBodies[0].appendChild(pnbGrandTotalRow(grand));
    }

    pnbSelectedKey = null;
    pnbSelectedEl.textContent = "Selected: —";
    pnbRowCountEl.textContent = `Rows: ${rows.length}`;
  }

  function selectPnbKey(key) {
    if (!pnbTable) return;
    const k = normalizeText(key);
    pnbSelectedKey = k || null;
    Array.from(pnbTable.tBodies[0].rows).forEach(r=>r.classList.remove("is-selected"));
    if (!k) { pnbSelectedEl.textContent="Selected: —"; return; }
    const tr = pnbTable.tBodies[0].querySelector(`tr[data-row-type="data"][data-key="${CSS.escape(k)}"]`);
    if (tr) { tr.classList.add("is-selected"); pnbSelectedEl.textContent = `Selected: ${tr.cells[0]?.innerText || "—"}`; }
    else pnbSelectedEl.textContent="Selected: —";
  }

  pnbTable?.addEventListener("click",(e)=>{
    const tr = e.target.closest("tr");
    if (!tr || tr.parentElement?.tagName !== "TBODY") return;
    if ((tr.dataset.rowType || "data") !== "data") return;
    selectPnbKey(tr.dataset.key);
  });
  pnbTable?.addEventListener("dblclick",(e)=>{
    const tr = e.target.closest("tr");
    if (!tr || tr.parentElement?.tagName !== "TBODY") return;
    if ((tr.dataset.rowType || "data") !== "data") return;
    selectPnbKey(tr.dataset.key);
    openPnbModal("edit", tr.dataset.key);
  });

  btnPnbAdd?.addEventListener("click",()=>openPnbModal("add"));
  btnPnbEdit?.addEventListener("click",()=>{
    if (!pnbSelectedKey) return alert("Please select a PNB Deposit row first.");
    openPnbModal("edit", pnbSelectedKey);
  });

  pnbForm?.addEventListener("submit",(e)=>{
    e.preventDefault();
    const date = pdYyyyMMddToMMDDYYYY(pdDate.value) || "";
    const amount = parseMoneyPD(pdAmount.value);

    if (pnbMode === "add") {
      const entry = ensurePnbId({ date, amount });
      DB.savePnbDeposit(entry).then(saved => { if (saved) entry.id = saved.id; });
      pnbStore.push(entry);
      closePnbModal();
      renderPnbTable();
      selectPnbKey(pnbKeyFor(entry));
      return;
    }

    const key = pnbEditingKey;
    if (!key) return;
    const idx = pnbStore.findIndex(x => pnbKeyFor(x) === key);
    if (idx < 0) return;

    const updatedPNB = ensurePnbId({ date, amount, _id: pnbStore[idx]._id, id: pnbStore[idx].id });
    pnbStore[idx] = updatedPNB;
    DB.savePnbDeposit(updatedPNB);
    closePnbModal();
    renderPnbTable();
    selectPnbKey(pnbKeyFor(updatedPNB));
  });

  btnPnbDel?.addEventListener("click",()=>{
    if (!pnbSelectedKey) return alert("Please select a PNB Deposit row first.");
    if (!confirm("Delete the selected PNB Deposit entry?")) return;
    const rowPNB = pnbStore.find(x => pnbKeyFor(x) === pnbSelectedKey);
    if (rowPNB?.id) DB.deletePnbDeposit(rowPNB.id);
    pnbStore = pnbStore.filter(x => pnbKeyFor(x) !== pnbSelectedKey);
    renderPnbTable();
  });

  btnPnbImport?.addEventListener("click",()=>filePnbImport?.click());

  async function importPnbRows(rows2d) {
    // try detect header: Date/Amount
    let headerRowIndex = -1;
    for (let r=0;r<Math.min(rows2d.length,6);r++){
      const row = rows2d[r].map(x=>String(x??"").trim().toLowerCase());
      if (row.includes("date") && (row.includes("amount") || row.includes("amt"))) { headerRowIndex = r; break; }
    }
    let dataRows = headerRowIndex >= 0 ? rows2d.slice(headerRowIndex+1) : rows2d;
    dataRows = dataRows.filter(r => r && r.some(v => String(v ?? "").trim() !== ""));
    if (!dataRows.length) { alert("No data rows found."); return; }

    const replace = confirm("Replace current PNB Deposit entries with imported data?\nOK = Replace\nCancel = Append");
    if (replace) pnbStore = [];

    let added=0, skipped=0;
    for (const row of dataRows) {
      const d0 = row[0];
      let date = "";
      if (d0 instanceof Date) date = formatMMDDYYYY(d0);
      else {
        const s = String(d0 ?? "").trim();
        if (/^\d{4}-\d{2}-\d{2}$/.test(s)) date = yyyyMMddToMMDDYYYY(s);
        else if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(s)) date = `${pad2(s.split("/")[0])}/${pad2(s.split("/")[1])}/${s.split("/")[2]}`;
        else {
          const d = new Date(s);
          date = Number.isFinite(d.getTime()) ? formatMMDDYYYY(d) : s;
        }
      }
      const amount = parseMoneyPD(row[1]);
      if (!date && !amount) { skipped++; continue; }
      pnbStore.push(ensurePnbId({ date, amount }));
      added++;
    }

    savePnbStore();
    renderPnbTable();
    alert(`Import complete.\nAdded: ${added}\nSkipped: ${skipped}`);
  }

  filePnbImport?.addEventListener("change", async ()=>{
    const file = filePnbImport.files?.[0];
    if (!file) return;
    try {
      const name = (file.name||"").toLowerCase();
      if (name.endsWith(".csv")) {
        const text = await file.text();
        await importPnbRows(parseCSVTo2D(text));
      } else if (name.endsWith(".xlsx")) {
        if (typeof XLSX === "undefined") { alert("XLSX library not loaded. Import a CSV instead."); return; }
        const buf = await file.arrayBuffer();
        const wb = XLSX.read(buf, { type:"array", cellDates:true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { header:1, defval:"" });
        await importPnbRows(rows);
      } else {
        alert("Unsupported file type. Please select a .csv or .xlsx file.");
      }
    } catch(err) {
      alert("Import failed: " + (err?.message || String(err)));
    } finally {
      filePnbImport.value = "";
    }
  });

  btnPnbExport?.addEventListener("click",()=>{
    const msoMoney = 'mso-number-format:"\\#\\,\\#\\#0\\.00"';
    const msoText = 'mso-number-format:"\\@"';
    const esc = (s)=>String(s??"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");

    const rows = getFilteredPnbRows().sort((a,b)=>{
      const ak = monthKeyFromDate(a.date);
      const bk = monthKeyFromDate(b.date);
      if (ak !== bk) {
        if (ak === "unknown") return 1;
        if (bk === "unknown") return -1;
        return ak.localeCompare(bk);
      }
      const ad = parseMMDDYYYY(a.date)?.getTime() ?? Number.POSITIVE_INFINITY;
      const bd = parseMMDDYYYY(b.date)?.getTime() ?? Number.POSITIVE_INFINITY;
      return ad - bd;
    });

    const groups = new Map();
    for (const r of rows) {
      const k = monthKeyFromDate(r.date);
      if (!groups.has(k)) groups.set(k, []);
      groups.get(k).push(r);
    }
    const keys = Array.from(groups.keys()).sort((a,b)=>{
      if (a === "unknown" && b !== "unknown") return 1;
      if (b === "unknown" && a !== "unknown") return -1;
      return a.localeCompare(b);
    });

    const css = `
      <style>
        table{ border-collapse:collapse; font-family:Calibri,Arial,sans-serif; font-size:11pt; }
        th,td{ border:1px solid #cfcfcf; padding:4px 6px; white-space:nowrap; }
        th{ font-weight:700; text-align:center; background:#fff; }
        td.num{ text-align:right; }
        tr.month-header td{ font-weight:800; text-align:center; background:#d9d9d9; border-top:2px solid #000; border-bottom:2px solid #000; }
        tr.month-total td{ font-weight:800; background:#fff2cc; border-top:2px solid #000; }
        tr.grand-total td{ font-weight:900; background:#92d050; border-top:2px solid #000; border-bottom:2px solid #000; }
        tr.spacer td{ border:none; height:10px; }
      </style>
    `;

    let html = `<!doctype html><html><head><meta charset="utf-8" />${css}</head><body><table><tbody>`;
    let grand = 0;

    for (let gi=0; gi<keys.length; gi++) {
      const key = keys[gi];
      const grp = groups.get(key) || [];
      if (!grp.length) continue;

      html += `<tr class="month-header"><td colspan="2" style="${msoText}">${esc(monthLabelFromKey(key).toUpperCase())}</td></tr>`;
      html += `<tr><th>Date</th><th>Amount</th></tr>`;

      let monthTotal = 0;
      for (const r of grp) {
        html += `<tr>`;
        html += `<td style="${msoText}">${esc(r.date)}</td>`;
        html += `<td class="num" style="${msoMoney}">${fmtMoney(r.amount)}</td>`;
        html += `</tr>`;
        monthTotal += Number(r.amount)||0;
      }

      html += `<tr class="month-total"><td style="${msoText}">TOTAL</td><td class="num" style="${msoMoney}">${fmtMoney(monthTotal)}</td></tr>`;
      grand += monthTotal;
      if (gi < keys.length - 1) html += `<tr class="spacer"><td colspan="2"></td></tr>`;
    }

    if (rows.length) {
      html += `<tr class="spacer"><td colspan="2"></td></tr>`;
      html += `<tr class="grand-total"><td style="${msoText}">GRAND TOTAL</td><td class="num" style="${msoMoney}">${fmtMoney(grand)}</td></tr>`;
    }

    html += `</tbody></table></body></html>`;

    const blob = new Blob([html], { type:"application/vnd.ms-excel;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href=url;
    a.download=`Magallanes_PNBDeposit_${new Date().toISOString().slice(0,10)}.xls`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  });

  btnPnbRefresh?.addEventListener("click",()=>{ renderPnbTable(); alert("Refreshed PNB Deposit view."); });

  function initPnbFromDOM() {
    if (!pnbTable) return [];
    const trs = Array.from(pnbTable.tBodies[0].rows).filter(r => (r.dataset.rowType || "data") === "data");
    return trs.map(tr => ensurePnbId({
      date: String(tr.cells[0]?.innerText||"").trim(),
      amount: parseMoneyPD(String(tr.cells[1]?.innerText||""))
    })).filter(r => r.date || r.amount);
  }

  // Load PNB deposit from Supabase
  DB.getPnbDeposit().then(rows => {
    pnbStore = rows;
    renderPnbTable();
  });

  renderPnbTable();


  // ---------------------------
  // DSWD Tab Logic
  // ---------------------------
  const dswdTable      = $("#dswdTable");
  const dswdRowCountEl = $("#dswdRowCount");
  const dswdSelectedEl = $("#dswdSelected");

  const btnDswdAdd    = $("#btnDswdAddEntry");
  const btnDswdEdit   = $("#btnDswdEditSelected");
  const btnDswdDel    = $("#btnDswdDeleteSelected");
  const btnDswdExport = $("#btnDswdExportExcel");
  const btnDswdRefresh= $("#btnDswdRefresh");

  const dswdOverlay   = $("#dswdOverlay");
  const dswdModal     = $("#dswdModal");
  const dswdTitle     = $("#dswdTitle");
  const dswdSubtitle  = $("#dswdSubtitle");
  const dswdForm      = $("#dswdForm");
  const btnCloseDswd  = $("#btnCloseDswd");
  const btnCancelDswd = $("#btnCancelDswd");
  const btnSubmitDswd = $("#btnSubmitDswd");

  const dwContract     = $("#dwContract");
  const dwDate         = $("#dwDate");
  const dwDeceased     = $("#dwDeceased");
  const dwContractAmt  = $("#dwContractAmt");
  const dwPayment      = $("#dwPayment");
  const dwBalance      = $("#dwBalance");
  const dwDswdRefund   = $("#dwDswdRefund");
  const dwAfterTax     = $("#dwAfterTax");
  const dwDateReceived = $("#dwDateReceived");
  const dwPayable      = $("#dwPayable");
  const dwDateRelease  = $("#dwDateRelease");
  const dwBeneficiary  = $("#dwBeneficiary");
  const dwDswdDiscount = $("#dwDswdDiscount");
  const dwStatus       = $("#dwStatus");

  let dswdStore = [];
  let dswdSelectedKey = null;
  let dswdMode = "add";
  let dswdEditingKey = null;

  function dswdKeyFor(r) { return r.id || r._id || ""; }

  function ensureDswdId(r) {
    if (!r._id) r._id = "__DSWD__" + Date.now() + "_" + Math.floor(Math.random()*100000);
    return r;
  }

  // ── Autofill from Contracts ──
  function dswdAutofill(contractNo) {
    const key = normalizeText(contractNo);
    const c = contractsStore.find(x => normalizeText(x.contract) === key);
    if (!c) {
      dwDeceased.value = "";
      dwContractAmt.value = "0.00";
      dwPayment.value = "0.00";
      dwBalance.value = "0.00";
      return;
    }
    // Compute totals same way as contracts tab
    const totalPaid = Math.max(0, (Number(c.inhaus)||0) + (Number(c.bai)||0) + (Number(c.gl)||0) + (Number(c.gcash)||0) + (Number(c.cash)||0));
    const remaining = (Number(c.amount)||0) - totalPaid - (Number(c.discount)||0);
    dwDeceased.value     = c.deceased || "";
    dwContractAmt.value  = (Number(c.amount)||0).toFixed(2);
    dwPayment.value      = totalPaid.toFixed(2);
    dwBalance.value      = remaining.toFixed(2);
  }

  dwContract?.addEventListener("blur", () => {
    if (dswdMode === "add") dswdAutofill(dwContract.value.trim());
  });
  dwContract?.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === "Tab") {
      dswdAutofill(dwContract.value.trim());
    }
  });

  // Auto-calculate After Tax and DSWD Discount from DSWD Refund
  function recalcDswdRefund() {
    const refund   = Number(dwDswdRefund?.value) || 0;
    const discount = refund * 0.0625;
    const afterTax = refund - discount;
    if (dwAfterTax)     dwAfterTax.value     = afterTax.toFixed(2);
    if (dwDswdDiscount) dwDswdDiscount.value = discount.toFixed(2);
  }
  dwDswdRefund?.addEventListener("input", recalcDswdRefund);
  dwDswdRefund?.addEventListener("change", recalcDswdRefund);

  // ── Filter ──
  const dswdFilterCategory = $("#dswdFilterCategory");
  const dswdFilterInputs   = $("#dswdFilterInputs");
  const btnDswdApplyFilter = $("#btnDswdApplyFilter");
  const btnDswdClearFilter = $("#btnDswdClearFilter");
  let dswdActiveFilter = { category: "", value: null };

  function setDswdFilterInputs(category) {
    if (!dswdFilterInputs) return;
    dswdFilterInputs.innerHTML = "";
    function makeText(id, label, ph) {
      const wrap = document.createElement("label");
      wrap.className = "field inline";
      wrap.innerHTML = `<span>${label}</span><input class="input" id="${id}" type="text" placeholder="${ph}" style="width:160px;" />`;
      dswdFilterInputs.appendChild(wrap);
    }
    if (category === "date")     makeText("dswdFilterDate",     "Date",     "MM/DD/YYYY or MM/YYYY");
    if (category === "contract") makeText("dswdFilterContract", "Contract #","e.g. MC-0105");
    if (category === "deceased") makeText("dswdFilterDeceased", "Deceased", "name...");
  }

  dswdFilterCategory?.addEventListener("change", () => setDswdFilterInputs(dswdFilterCategory.value));

  function getDswdFilterFromUI() {
    const cat = dswdFilterCategory?.value || "";
    if (cat === "date")     return { category:"date",     value: ($("#dswdFilterDate")?.value||"").trim() };
    if (cat === "contract") return { category:"contract", value: normalizeText($("#dswdFilterContract")?.value||"") };
    if (cat === "deceased") return { category:"deceased", value: normalizeText($("#dswdFilterDeceased")?.value||"") };
    return { category:"", value: null };
  }

  function dswdRowMatchesFilter(r, f) {
    if (!f || !f.category) return true;
    if (f.category === "date" && f.value) {
      const v = String(f.value);
      return String(r.date||"").startsWith(v.slice(0,2)) || String(r.date||"").includes(v);
    }
    if (f.category === "contract" && f.value) return normalizeText(r.contract||"").includes(f.value);
    if (f.category === "deceased" && f.value) return normalizeText(r.deceased||"").includes(f.value);
    return true;
  }

  btnDswdApplyFilter?.addEventListener("click", () => { dswdActiveFilter = getDswdFilterFromUI(); renderDswdTable(); });
  btnDswdClearFilter?.addEventListener("click", () => {
    dswdActiveFilter = { category:"", value:null };
    if (dswdFilterCategory) dswdFilterCategory.value = "";
    if (dswdFilterInputs) dswdFilterInputs.innerHTML = "";
    renderDswdTable();
  });

  // ── Render ──
  function fmtNum(n) { return (Number(n)||0).toLocaleString("en-US",{minimumFractionDigits:2,maximumFractionDigits:2}); }

  function renderDswdTable() {
    if (!dswdTable) return;
    const filtered = dswdStore.filter(r => dswdRowMatchesFilter(r, dswdActiveFilter));
    const tbody = dswdTable.tBodies[0];
    tbody.innerHTML = "";

    if (filtered.length === 0) {
      const tr = document.createElement("tr");
      tr.innerHTML = `<td colspan="13" style="text-align:center;opacity:0.5;padding:18px;">(no entries)</td>`;
      tbody.appendChild(tr);
    } else {
      filtered.forEach(r => {
        // Live balance — look up current Remaining from Contracts Tab
        const contractKey = normalizeText(r.contract || "");
        const matchedContract = contractsStore.find(c => normalizeText(c.contract || "") === contractKey);
        const liveBalance = matchedContract ? calcComputed(matchedContract).remaining : (r.balance || 0);

        const tr = document.createElement("tr");
        tr.dataset.rowType = "data";
        tr.dataset.id = dswdKeyFor(r);
        if (dswdKeyFor(r) === dswdSelectedKey) tr.classList.add("is-selected");
        if (r.status === "Processed")       tr.classList.add("dswd-processed");
        else if (r.status === "Scheduled")  tr.classList.add("dswd-scheduled");
        tr.innerHTML = `
          <td>${r.date||""}</td>
          <td>${r.contract||""}</td>
          <td>${r.deceased||""}</td>
          <td class="num">${fmtNum(r.contractAmt)}</td>
          <td class="num">${fmtNum(r.payment)}</td>
          <td class="num">${fmtNum(liveBalance)}</td>
          <td class="num">${fmtNum(r.dswdRefund)}</td>
          <td class="num">${fmtNum(r.afterTax)}</td>
          <td>${r.dateReceived||""}</td>
          <td class="num">${fmtNum(r.payable)}</td>
          <td>${r.dateRelease||""}</td>
          <td>${r.beneficiary||""}</td>
          <td class="num">${fmtNum(r.dswdDiscount)}</td>
          <td>${r.status||"Waiting"}</td>
        `;
        tbody.appendChild(tr);
      });
    }

    if (dswdRowCountEl) dswdRowCountEl.textContent = `Rows: ${filtered.length}`;
    if (dswdSelectedEl) dswdSelectedEl.textContent = dswdSelectedKey ? `Selected: ${dswdSelectedKey}` : "Selected: —";

    syncDswdToContracts();
  }

  // ── Sync DSWD amounts into contracts ──
  // dswdAfterTax → "DSWD" column (reduces remaining, not totalPaid)
  // dswdDiscount → "DSWD Discount" column (also reduces remaining)
  let _syncingDswd = false;
  function syncDswdToContracts() {
    if (_syncingDswd || !Array.isArray(contractsStore)) return;
    _syncingDswd = true;
    const afterTaxByContract = new Map();
    const discountByContract = new Map();
    for (const r of dswdStore) {
      const key = normalizeText(r.contract || "");
      if (!key) continue;
      afterTaxByContract.set(key, (afterTaxByContract.get(key) || 0) + (Number(r.afterTax)    || 0));
      discountByContract.set(key, (discountByContract.get(key) || 0) + (Number(r.dswdDiscount) || 0));
    }
    let changed = false;
    for (const c of contractsStore) {
      const key = normalizeText(c.contract || "");
      const newAfterTax = afterTaxByContract.get(key) || 0;
      const newDiscount = discountByContract.get(key) || 0;
      if ((Number(c.dswdAfterTax) || 0) !== newAfterTax) { c.dswdAfterTax = newAfterTax; changed = true; }
      if ((Number(c.dswdDiscount) || 0) !== newDiscount) { c.dswdDiscount = newDiscount; changed = true; }
    }
    _syncingDswd = false;
    if (changed) renderContracts();
  }

  dswdTable?.addEventListener("click", (e) => {
    const tr = e.target.closest("tr");
    if (!tr || (tr.dataset.rowType||"") !== "data") return;
    dswdSelectedKey = tr.dataset.id || null;
    // Update selection highlight without re-rendering — preserve status classes
    Array.from(dswdTable.tBodies[0].rows).forEach(r => {
      r.classList.toggle("is-selected", r.dataset.id === dswdSelectedKey);
    });
    if (dswdSelectedEl) dswdSelectedEl.textContent = dswdSelectedKey ? `Selected: ${dswdSelectedKey}` : "Selected: —";
  });

  dswdTable?.addEventListener("dblclick", (e) => {
    const tr = e.target.closest("tr");
    if (!tr || (tr.dataset.rowType||"") !== "data") return;
    dswdSelectedKey = tr.dataset.id || null;
    openDswdModal("edit", dswdSelectedKey);
  });

  // ── Modal ──
  function mmddyyyyFromDateInput(v) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(v||""));
    return m ? `${m[2]}/${m[3]}/${m[1]}` : "";
  }
  function dateInputFromMmddyyyy(v) {
    const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(String(v||"").trim());
    if (!m) return "";
    return `${m[3]}-${String(m[1]).padStart(2,"0")}-${String(m[2]).padStart(2,"0")}`;
  }

  function openDswdModal(mode, keyOrNull = null) {
    if (!dswdOverlay || !dswdModal) { alert("DSWD form not available."); return; }
    dswdMode = mode;
    dswdEditingKey = null;

    // Set readonly state on autofill fields
    [dwDeceased, dwContractAmt, dwPayment, dwBalance].forEach(el => {
      if (el) { el.readOnly = (mode === "add"); el.style.opacity = mode === "add" ? "0.75" : "1"; }
    });

    if (mode === "add") {
      dswdTitle.textContent = "Add DSWD Entry";
      dswdSubtitle.textContent = "Enter Contract # to auto-fill details, then complete remaining fields.";
      btnSubmitDswd.textContent = "Add Entry";

      dwContract.value = ""; dwDate.value = new Date().toISOString().slice(0,10);
      dwDeceased.value = ""; dwContractAmt.value = "0.00";
      dwPayment.value = "0.00"; dwBalance.value = "0.00";
      dwDswdRefund.value = "0.00"; dwAfterTax.value = "0.00";
      dwDateReceived.value = ""; dwPayable.value = "0.00";
      dwDateRelease.value = ""; dwBeneficiary.value = ""; dwDswdDiscount.value = "0.00";
      if (dwStatus) dwStatus.value = "Waiting";
    } else {
      dswdTitle.textContent = "Edit DSWD Entry";
      dswdSubtitle.textContent = "Update this DSWD entry.";
      btnSubmitDswd.textContent = "Save Changes";

      const found = dswdStore.find(x => dswdKeyFor(x) === keyOrNull);
      if (!found) { alert("Could not find selected entry."); return; }
      dswdEditingKey = keyOrNull;

      dwContract.value     = found.contract || "";
      dwDate.value         = dateInputFromMmddyyyy(found.date) || new Date().toISOString().slice(0,10);
      dwDeceased.value     = found.deceased || "";
      dwContractAmt.value  = (Number(found.contractAmt)||0).toFixed(2);
      dwPayment.value      = (Number(found.payment)||0).toFixed(2);
      dwBalance.value      = (Number(found.balance)||0).toFixed(2);
      dwDswdRefund.value   = (Number(found.dswdRefund)||0).toFixed(2);
      dwAfterTax.value     = (Number(found.afterTax)||0).toFixed(2);
      dwDateReceived.value = dateInputFromMmddyyyy(found.dateReceived) || "";
      dwPayable.value      = (Number(found.payable)||0).toFixed(2);
      dwDateRelease.value  = dateInputFromMmddyyyy(found.dateRelease) || "";
      dwBeneficiary.value  = found.beneficiary || "";
      dwDswdDiscount.value = (Number(found.dswdDiscount)||0).toFixed(2);
      if (dwStatus) dwStatus.value = found.status || "Waiting";
    }

    dswdOverlay.classList.add("is-open");
    dswdModal.classList.add("is-open");
    setTimeout(() => dwContract.focus(), 0);
  }

  function closeDswdModal() {
    dswdOverlay?.classList.remove("is-open");
    dswdModal?.classList.remove("is-open");
    dswdMode = "add"; dswdEditingKey = null;
  }

  btnCloseDswd?.addEventListener("click", closeDswdModal);
  btnCancelDswd?.addEventListener("click", closeDswdModal);
  dswdOverlay?.addEventListener("click", closeDswdModal);
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && dswdModal?.classList.contains("is-open")) closeDswdModal();
  });

  btnDswdAdd?.addEventListener("click", () => openDswdModal("add"));
  btnDswdEdit?.addEventListener("click", () => {
    if (!dswdSelectedKey) return alert("Please select a row first.");
    openDswdModal("edit", dswdSelectedKey);
  });

  dswdForm?.addEventListener("submit", (e) => {
    e.preventDefault();
    const contractNo = (dwContract.value || "").trim();
    if (!contractNo) return alert("Contract # is required.");

    const entry = {
      date:         mmddyyyyFromDateInput(dwDate.value) || "",
      contract:     contractNo,
      deceased:     dwDeceased.value.trim(),
      contractAmt:  Number(dwContractAmt.value) || 0,
      payment:      Number(dwPayment.value) || 0,
      balance:      Number(dwBalance.value) || 0,
      dswdRefund:   Number(dwDswdRefund.value) || 0,
      afterTax:     Number(dwAfterTax.value) || 0,
      dateReceived: mmddyyyyFromDateInput(dwDateReceived.value) || "",
      payable:      Number(dwPayable.value) || 0,
      dateRelease:  mmddyyyyFromDateInput(dwDateRelease.value) || "",
      beneficiary:  dwBeneficiary.value.trim(),
      dswdDiscount: Number(dwDswdDiscount.value) || 0,
      status:       dwStatus?.value || "Waiting",
    };

    if (dswdMode === "add") {
      ensureDswdId(entry);
      DB.saveDswd(entry).then(saved => {
        if (saved) {
          entry.id = saved.id;
          renderDswdTable();
        }
      });
      dswdStore.push(entry);
    } else {
      const idx = dswdStore.findIndex(x => dswdKeyFor(x) === dswdEditingKey);
      if (idx < 0) return alert("Could not find entry to update.");
      entry.id = dswdStore[idx].id;
      entry._id = dswdStore[idx]._id;
      dswdStore[idx] = entry;
      DB.saveDswd(entry);
    }

    closeDswdModal();
    renderDswdTable();
  });

  btnDswdDel?.addEventListener("click", () => {
    if (!dswdSelectedKey) return alert("Please select a row first.");
    if (!confirm("Delete the selected DSWD entry?")) return;
    const row = dswdStore.find(x => dswdKeyFor(x) === dswdSelectedKey);
    if (row?.id) DB.deleteDswd(row.id);
    dswdStore = dswdStore.filter(x => dswdKeyFor(x) !== dswdSelectedKey);
    dswdSelectedKey = null;
    renderDswdTable();
  });

  btnDswdRefresh?.addEventListener("click", () => {
    DB.getDswd().then(rows => { dswdStore = rows; renderDswdTable(); });
    alert("Refreshed DSWD view.");
  });

  // ── Export to Excel ──
  btnDswdExport?.addEventListener("click", () => {
    if (!dswdStore.length) return alert("No DSWD entries to export.");
    const headers = ["Date","Contract #","Name of Deceased","Contract","Payment","Balance","DSWD Refund","After Tax","Date Received from DSWD","Payable","Date/Release","Beneficiary","DSWD Discount","Status"];
    const rows = dswdStore.map(r => [
      r.date, r.contract, r.deceased,
      r.contractAmt, r.payment, r.balance,
      r.dswdRefund, r.afterTax, r.dateReceived,
      r.payable, r.dateRelease, r.beneficiary, r.dswdDiscount, r.status||"Waiting"
    ]);
    let csv = headers.join(",") + "\n";
    rows.forEach(row => { csv += row.map(v => `"${String(v||"").replace(/"/g,'""')}"`).join(",") + "\n"; });
    const blob = new Blob([csv], { type:"text/csv" });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement("a");
    a.href = url; a.download = `MagallanesDSWD_${new Date().toISOString().slice(0,10)}.csv`;
    document.body.appendChild(a); a.click(); a.remove();
    URL.revokeObjectURL(url);
  });

  // ---------------------------
  // BAI Tab Logic
  // ---------------------------
  const baiTable      = $("#baiTable");
  const baiRowCountEl = $("#baiRowCount");
  const baiSelectedEl = $("#baiSelected");
  const btnBaiAdd     = $("#btnBaiAddEntry");
  const btnBaiEdit    = $("#btnBaiEditSelected");
  const btnBaiDel     = $("#btnBaiDeleteSelected");
  const btnBaiExport  = $("#btnBaiExportExcel");
  const btnBaiRefresh = $("#btnBaiRefresh");
  const baiOverlay    = $("#baiOverlay");
  const baiModal      = $("#baiModal");
  const baiTitleEl    = $("#baiTitle");
  const baiSubtitleEl = $("#baiSubtitle");
  const baiForm       = $("#baiForm");
  const btnCloseBai   = $("#btnCloseBai");
  const btnCancelBai  = $("#btnCancelBai");
  const btnSubmitBai  = $("#btnSubmitBai");
  const baDateApplied   = $("#baDateApplied");
  const baContract      = $("#baContract");
  const baAmount        = $("#baAmount");
  const baDateCompleted = $("#baDateCompleted");
  const baStatus        = $("#baStatus");

  let baiStore = [];
  let baiSelectedKey = null;
  let baiMode = "add";
  let baiEditingKey = null;

  function baiKeyFor(r) { return r.id || r._id || ""; }
  function ensureBaiId(r) {
    if (!r._id) r._id = "__BAI__" + Date.now() + "_" + Math.floor(Math.random()*100000);
    return r;
  }
  function baiComputeStatus(dateCompleted) {
    return (dateCompleted && String(dateCompleted).trim()) ? "Completed" : "Pending";
  }
  function baiMmddyyyyFromInput(v) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(v||""));
    return m ? `${m[2]}/${m[3]}/${m[1]}` : "";
  }
  function baiInputFromMmddyyyy(v) {
    const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(String(v||"").trim());
    if (!m) return "";
    return `${m[3]}-${String(m[1]).padStart(2,"0")}-${String(m[2]).padStart(2,"0")}`;
  }

  // Auto-set status when Date Completed changes
  baDateCompleted?.addEventListener("input",  () => { if (baStatus) baStatus.value = baiComputeStatus(baDateCompleted.value); });
  baDateCompleted?.addEventListener("change", () => { if (baStatus) baStatus.value = baiComputeStatus(baDateCompleted.value); });

  // ── Filter ──
  const baiFilterCategory = $("#baiFilterCategory");
  const baiFilterInputs   = $("#baiFilterInputs");
  const btnBaiApplyFilter = $("#btnBaiApplyFilter");
  const btnBaiClearFilter = $("#btnBaiClearFilter");
  let baiActiveFilter = { category: "", value: null };

  function setBaiFilterInputs(cat) {
    if (!baiFilterInputs) return;
    baiFilterInputs.innerHTML = "";
    const mk = (id, label, ph) => {
      const w = document.createElement("label");
      w.className = "field inline";
      w.innerHTML = `<span>${label}</span><input class="input" id="${id}" type="text" placeholder="${ph}" style="width:160px;" />`;
      baiFilterInputs.appendChild(w);
    };
    if (cat === "date")     mk("baiFilterDate",     "Date Applied", "MM/DD/YYYY");
    if (cat === "contract") mk("baiFilterContract", "Contract #",   "e.g. MC-0105");
    if (cat === "status")   mk("baiFilterStatus",   "Status",       "Pending or Completed");
  }
  baiFilterCategory?.addEventListener("change", () => setBaiFilterInputs(baiFilterCategory.value));

  function getBaiFilter() {
    const cat = baiFilterCategory?.value || "";
    if (cat === "date")     return { category:"date",     value: ($("#baiFilterDate")?.value||"").trim() };
    if (cat === "contract") return { category:"contract", value: normalizeText($("#baiFilterContract")?.value||"") };
    if (cat === "status")   return { category:"status",   value: normalizeText($("#baiFilterStatus")?.value||"") };
    return { category:"", value:null };
  }
  function baiRowMatches(r, f) {
    if (!f || !f.category) return true;
    if (f.category === "date"     && f.value) return String(r.dateApplied||"").includes(f.value);
    if (f.category === "contract" && f.value) return normalizeText(r.contract||"").includes(f.value);
    if (f.category === "status"   && f.value) return normalizeText(r.status||"").includes(f.value);
    return true;
  }
  btnBaiApplyFilter?.addEventListener("click", () => { baiActiveFilter = getBaiFilter(); renderBaiTable(); });
  btnBaiClearFilter?.addEventListener("click", () => {
    baiActiveFilter = { category:"", value:null };
    if (baiFilterCategory) baiFilterCategory.value = "";
    if (baiFilterInputs)   baiFilterInputs.innerHTML = "";
    renderBaiTable();
  });

  // ── Sync BAI amounts into contracts (reduces remaining, NOT totalPaid) ──
  let _syncingBai = false;
  function syncBaiToContracts() {
    if (_syncingBai || !Array.isArray(contractsStore)) return;
    _syncingBai = true;
    const baiByContract = new Map();
    for (const r of baiStore) {
      const key = normalizeText(r.contract || "");
      if (!key) continue;
      baiByContract.set(key, (baiByContract.get(key) || 0) + (Number(r.amount) || 0));
    }
    let changed = false;
    for (const c of contractsStore) {
      const key = normalizeText(c.contract || "");
      const baiAmt = baiByContract.get(key) || 0;
      if ((Number(c.baiAssist) || 0) !== baiAmt) {
        c.baiAssist = baiAmt;
        changed = true;
      }
    }
    _syncingBai = false;
    if (changed) renderContracts();
  }

  // ── Render ──
  function renderBaiTable() {
    if (!baiTable) return;
    const filtered = baiStore.filter(r => baiRowMatches(r, baiActiveFilter));
    const tbody = baiTable.tBodies[0];
    tbody.innerHTML = "";

    if (filtered.length === 0) {
      const tr = document.createElement("tr");
      tr.innerHTML = `<td colspan="5" style="text-align:center;opacity:0.5;padding:18px;">(no entries)</td>`;
      tbody.appendChild(tr);
    } else {
      // ── Group by month using dateApplied ──
      const groups = new Map();
      const keyOrder = [];
      filtered.forEach(r => {
        const key = monthKeyFromDate(r.dateApplied || "");
        if (!groups.has(key)) { groups.set(key, []); keyOrder.push(key); }
        groups.get(key).push(r);
      });
      keyOrder.sort();

      // ── Pre-compute Cash Received BAI totals by month ──
      // cashStore entries where particular maps to "bai" bucket
      const cashBaiByMonth = new Map();
      for (const r of (cashStore || [])) {
        const p = normalizeText(r.particular || "");
        const isBai = /\bbai\b/.test(p) || /\bbank\b/.test(p) || p.includes("bank received") || p.includes("deposit bank");
        if (!isBai) continue;
        const key = monthKeyFromDate(r.date || "");
        cashBaiByMonth.set(key, (cashBaiByMonth.get(key) || 0) + (Number(r.amount) || 0));
      }

      // ── Render each month group ──
      keyOrder.forEach(key => {
        const rows = groups.get(key);

        // Month header row
        const thdr = document.createElement("tr");
        thdr.dataset.rowType = "monthHeader";
        thdr.innerHTML = `<td colspan="5" style="padding:6px 10px;font-weight:900;opacity:0.7;font-size:11px;letter-spacing:.06em;text-transform:uppercase;border-bottom:1px solid var(--line2);">${monthLabelFromKey(key)}</td>`;
        tbody.appendChild(thdr);

        // Data rows
        let monthTotal = 0;
        rows.forEach(r => {
          monthTotal += Number(r.amount) || 0;
          const tr = document.createElement("tr");
          tr.dataset.rowType = "data";
          tr.dataset.id = baiKeyFor(r);
          if (baiKeyFor(r) === baiSelectedKey) tr.classList.add("is-selected");
          if (r.status === "Completed") tr.classList.add("dswd-processed");
          tr.innerHTML = `
            <td>${r.dateApplied||""}</td>
            <td>${r.contract||""}</td>
            <td class="num">${fmtNum(r.amount)}</td>
            <td>${r.dateCompleted||""}</td>
            <td>${r.status||"Pending"}</td>
          `;
          tbody.appendChild(tr);
        });

        // Monthly Total row
        const tTotal = document.createElement("tr");
        tTotal.dataset.rowType = "monthTotal";
        tTotal.innerHTML = `
          <td colspan="2" style="font-weight:700;padding:5px 10px;border-top:1px solid var(--line2);">Total</td>
          <td class="num" style="font-weight:700;border-top:1px solid var(--line2);">${fmtNum(monthTotal)}</td>
          <td colspan="2" style="border-top:1px solid var(--line2);"></td>
        `;
        tbody.appendChild(tTotal);

        // Total BAI Collected row (from Cash Received)
        const cashBaiTotal = cashBaiByMonth.get(key) || 0;
        const tCollected = document.createElement("tr");
        tCollected.dataset.rowType = "baiCollected";
        tCollected.innerHTML = `
          <td colspan="2" style="padding:4px 10px;opacity:0.75;font-style:italic;">Total BAI Collected</td>
          <td class="num" style="opacity:0.75;font-style:italic;">${fmtNum(cashBaiTotal)}</td>
          <td colspan="2"></td>
        `;
        tbody.appendChild(tCollected);

        // Balance row
        const balance = monthTotal - cashBaiTotal;
        const balanceColor = balance > 0 ? "color:var(--danger,#e55);" : balance < 0 ? "color:var(--accent,#4c8);" : "";
        const tBalance = document.createElement("tr");
        tBalance.dataset.rowType = "baiBalance";
        tBalance.innerHTML = `
          <td colspan="2" style="font-weight:700;padding:4px 10px 10px;${balanceColor}">Balance</td>
          <td class="num" style="font-weight:700;${balanceColor}">${fmtNum(balance)}</td>
          <td colspan="2" style="padding-bottom:10px;"></td>
        `;
        tbody.appendChild(tBalance);
      });
    }

    if (baiRowCountEl) baiRowCountEl.textContent = `Rows: ${filtered.length}`;
    if (baiSelectedEl) baiSelectedEl.textContent = baiSelectedKey ? `Selected: ${baiSelectedKey}` : "Selected: —";
    syncBaiToContracts();
  }

  // ── Click / dblclick ──
  baiTable?.addEventListener("click", (e) => {
    const tr = e.target.closest("tr");
    if (!tr || (tr.dataset.rowType||"") !== "data") return;
    baiSelectedKey = tr.dataset.id || null;
    Array.from(baiTable.tBodies[0].rows).forEach(r => r.classList.toggle("is-selected", r.dataset.id === baiSelectedKey));
    if (baiSelectedEl) baiSelectedEl.textContent = baiSelectedKey ? `Selected: ${baiSelectedKey}` : "Selected: —";
  });
  baiTable?.addEventListener("dblclick", (e) => {
    const tr = e.target.closest("tr");
    if (!tr || (tr.dataset.rowType||"") !== "data") return;
    baiSelectedKey = tr.dataset.id || null;
    openBaiModal("edit", baiSelectedKey);
  });

  // ── Modal ──
  function openBaiModal(mode, keyOrNull = null) {
    if (!baiOverlay || !baiModal) { alert("BAI form not available."); return; }
    baiMode = mode;
    baiEditingKey = null;
    if (mode === "add") {
      baiTitleEl.textContent    = "Add BAI Entry";
      baiSubtitleEl.textContent = "Leave Date Completed blank if still pending.";
      btnSubmitBai.textContent  = "Add Entry";
      baDateApplied.value   = new Date().toISOString().slice(0,10);
      baContract.value      = "";
      baAmount.value        = "0.00";
      baDateCompleted.value = "";
      baStatus.value        = "Pending";
    } else {
      baiTitleEl.textContent    = "Edit BAI Entry";
      baiSubtitleEl.textContent = "Update this BAI entry.";
      btnSubmitBai.textContent  = "Save Changes";
      const found = baiStore.find(x => baiKeyFor(x) === keyOrNull);
      if (!found) { alert("Could not find selected entry."); return; }
      baiEditingKey = keyOrNull;
      baDateApplied.value   = baiInputFromMmddyyyy(found.dateApplied)   || new Date().toISOString().slice(0,10);
      baContract.value      = found.contract      || "";
      baAmount.value        = (Number(found.amount)||0).toFixed(2);
      baDateCompleted.value = baiInputFromMmddyyyy(found.dateCompleted) || "";
      baStatus.value        = found.status || "Pending";
    }
    baiOverlay.classList.add("is-open");
    baiModal.classList.add("is-open");
    setTimeout(() => baContract.focus(), 0);
  }

  function closeBaiModal() {
    baiOverlay?.classList.remove("is-open");
    baiModal?.classList.remove("is-open");
    baiMode = "add"; baiEditingKey = null;
  }

  btnCloseBai?.addEventListener("click",  closeBaiModal);
  btnCancelBai?.addEventListener("click", closeBaiModal);
  baiOverlay?.addEventListener("click",   closeBaiModal);
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && baiModal?.classList.contains("is-open")) closeBaiModal();
  });

  btnBaiAdd?.addEventListener("click",  () => openBaiModal("add"));
  btnBaiEdit?.addEventListener("click", () => {
    if (!baiSelectedKey) return alert("Please select a row first.");
    openBaiModal("edit", baiSelectedKey);
  });

  baiForm?.addEventListener("submit", (e) => {
    e.preventDefault();
    const contractNo    = (baContract.value||"").trim();
    if (!contractNo) return alert("Contract # is required.");
    const dateCompleted = baiMmddyyyyFromInput(baDateCompleted.value) || "";
    const entry = {
      dateApplied:   baiMmddyyyyFromInput(baDateApplied.value) || "",
      contract:      contractNo,
      amount:        Number(baAmount.value) || 0,
      dateCompleted,
      status:        baiComputeStatus(dateCompleted),
    };
    if (baiMode === "add") {
      ensureBaiId(entry);
      DB.saveBai(entry).then(saved => { if (saved) { entry.id = saved.id; renderBaiTable(); } });
      baiStore.push(entry);
    } else {
      const idx = baiStore.findIndex(x => baiKeyFor(x) === baiEditingKey);
      if (idx < 0) return alert("Could not find entry to update.");
      entry.id  = baiStore[idx].id;
      entry._id = baiStore[idx]._id;
      baiStore[idx] = entry;
      DB.saveBai(entry);
    }
    closeBaiModal();
    renderBaiTable();
  });

  btnBaiDel?.addEventListener("click", () => {
    if (!baiSelectedKey) return alert("Please select a row first.");
    if (!confirm("Delete the selected BAI entry?")) return;
    const row = baiStore.find(x => baiKeyFor(x) === baiSelectedKey);
    if (row?.id) DB.deleteBai(row.id);
    baiStore = baiStore.filter(x => baiKeyFor(x) !== baiSelectedKey);
    baiSelectedKey = null;
    renderBaiTable();
  });

  btnBaiRefresh?.addEventListener("click", () => {
    DB.getBai().then(rows => { baiStore = rows; renderBaiTable(); });
  });

  btnBaiExport?.addEventListener("click", () => {
    if (!baiStore.length) return alert("No BAI entries to export.");
    const headers = ["Date Applied","Contract #","Amount","Date Completed","Status"];
    const rows = baiStore.map(r => [r.dateApplied, r.contract, r.amount, r.dateCompleted, r.status]);
    let csv = headers.join(",") + "\n";
    rows.forEach(row => { csv += row.map(v => `"${String(v||"").replace(/"/g,'""')}"`).join(",") + "\n"; });
    const blob = new Blob([csv], {type:"text/csv"});
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement("a"); a.href=url; a.download=`MagallanesBai_${new Date().toISOString().slice(0,10)}.csv`;
    document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
  });


  // ---------------------------
  // Settings Tab Logic (v115)
  // ---------------------------
  const SETTINGS_STORE_KEY = "mf_settings_store";

  const settingsForm       = $("#settingsForm");
  const setCashBalance     = $("#setCashBalance");
  const setBankBalance     = $("#setBankBalance");
  const setSigFinanceClerk   = $("#setSigFinanceClerk");
  const setSigAccountant     = $("#setSigAccountant");
  const setSigFinanceManager = $("#setSigFinanceManager");

  // Default personnel names (used when nothing is saved yet)
  const DEFAULT_SIG = {
    financeClerk:   "Jennifer F. Landicho",
    accountant:     "Ranni V. Dalisay",
    financeManager: "June Lizette M. Quizon"
  };

  function loadSettings(){
    DB.getSettings().then(data => {
      if (!data) return;
      if(setCashBalance)     setCashBalance.value     = Number(data.cash_balance || 0).toFixed(2);
      if(setBankBalance)     setBankBalance.value     = Number(data.bank_balance || 0).toFixed(2);
      if(setSigFinanceClerk)   setSigFinanceClerk.value   = data.finance_clerk   || DEFAULT_SIG.financeClerk;
      if(setSigAccountant)     setSigAccountant.value     = data.accountant      || DEFAULT_SIG.accountant;
      if(setSigFinanceManager) setSigFinanceManager.value = data.finance_manager || DEFAULT_SIG.financeManager;
    }).catch(()=>{});
  }

  function getSettingsData(){
    return {
      cashBalance:    Number(setCashBalance?.value)     || 0,
      bankBalance:    Number(setBankBalance?.value)     || 0,
      financeClerk:   setSigFinanceClerk?.value.trim()   || DEFAULT_SIG.financeClerk,
      accountant:     setSigAccountant?.value.trim()     || DEFAULT_SIG.accountant,
      financeManager: setSigFinanceManager?.value.trim() || DEFAULT_SIG.financeManager
    };
  }

  function saveSettings(data){
    DB.saveSettings(data).catch(e => console.error("saveSettings", e));
  }

  // Helper — returns current personnel from settings inputs (already loaded into DOM)
  function getSignaturePersonnel(){
    return {
      financeClerk:   (setSigFinanceClerk?.value  || DEFAULT_SIG.financeClerk).trim(),
      accountant:     (setSigAccountant?.value    || DEFAULT_SIG.accountant).trim(),
      financeManager: (setSigFinanceManager?.value|| DEFAULT_SIG.financeManager).trim(),
    };
  }

  let settingsSaveTimer = null;
  function queueSaveSettings(){
    clearTimeout(settingsSaveTimer);
    settingsSaveTimer = setTimeout(()=>{ saveSettings(getSettingsData()); }, 150);
  }

  setCashBalance?.addEventListener("input",       queueSaveSettings);
  setBankBalance?.addEventListener("input",       queueSaveSettings);
  setSigFinanceClerk?.addEventListener("input",   queueSaveSettings);
  setSigAccountant?.addEventListener("input",     queueSaveSettings);
  setSigFinanceManager?.addEventListener("input", queueSaveSettings);

  settingsForm?.addEventListener("submit",(e)=>{
    e.preventDefault();
    saveSettings(getSettingsData());
    alert("Settings saved successfully.");
  });

  loadSettings();

  // Reload settings whenever Settings sub-tab is opened
  document.addEventListener("click", (ev)=>{
    const btn = ev.target?.closest?.("[data-subtab='settings']");
    if(btn){ setTimeout(loadSettings, 0); }
  });

  // ── Backup & Restore ─────────────────────────────────────────────────────

  const btnCreateBackup  = document.querySelector("#btnCreateBackup");
  const btnRestoreBackup = document.querySelector("#btnRestoreBackup");
  const fileRestoreBackup = document.querySelector("#fileRestoreBackup");
  const backupStatus     = document.querySelector("#backupStatus");

  function showBackupStatus(msg, isError = false) {
    if (!backupStatus) return;
    backupStatus.textContent = msg;
    backupStatus.style.display = "block";
    backupStatus.style.background = isError ? "rgba(255,60,60,0.15)" : "rgba(60,200,100,0.15)";
    backupStatus.style.color = isError ? "#ff6b6b" : "#5edb8a";
    backupStatus.style.border = isError ? "1px solid rgba(255,60,60,0.3)" : "1px solid rgba(60,200,100,0.3)";
    clearTimeout(backupStatus._timer);
    backupStatus._timer = setTimeout(() => { backupStatus.style.display = "none"; }, 6000);
  }

  // ── Create Backup ──────────────────────────────────────────────────────
  btnCreateBackup?.addEventListener("click", async () => {
    btnCreateBackup.disabled = true;
    btnCreateBackup.textContent = "⏳ Fetching data…";
    try {
      // Pull everything fresh from Supabase
      const [contracts, cashReceived, cashExpense, bankReceived, bankExpense, pnbDeposit, settings] = await Promise.all([
        DB.getContracts(),
        DB.getCashReceived(),
        DB.getCashExpense(),
        DB.getBankReceived(),
        DB.getBankExpense(),
        DB.getPnbDeposit(),
        DB.getSettings(),
      ]);

      const backup = {
        _version:    3,
        _exportedAt: new Date().toISOString(),
        _source:     "MagallanesFuneral-Supabase",
        contracts,
        cashReceived,
        cashExpense,
        bankReceived,
        bankExpense,
        pnbDeposit,
        settings,
      };

      const json     = JSON.stringify(backup, null, 2);
      const date     = new Date().toISOString().slice(0, 10);
      const filename = `MagallanesBackup_${date}.json`;

      const blob = new Blob([json], { type: "application/json;charset=utf-8" });
      const url  = URL.createObjectURL(blob);
      const a    = document.createElement("a");
      a.href = url; a.download = filename;
      document.body.appendChild(a); a.click(); a.remove();
      URL.revokeObjectURL(url);

      showBackupStatus(
        `✓ Backup saved as "${filename}"  •  ` +
        `Contracts: ${contracts.length}  •  ` +
        `Cash Received: ${cashReceived.length}  •  ` +
        `Cash Expense: ${cashExpense.length}  •  ` +
        `Bank Received: ${bankReceived.length}  •  ` +
        `Bank Expense: ${bankExpense.length}  •  ` +
        `PNB Deposit: ${pnbDeposit.length}`
      );
    } catch (err) {
      showBackupStatus("✗ Backup failed: " + (err?.message || err), true);
    } finally {
      btnCreateBackup.disabled = false;
      btnCreateBackup.textContent = "⬇ Create Backup";
    }
  });

  // ── Restore from Backup ───────────────────────────────────────────────
  btnRestoreBackup?.addEventListener("click", () => fileRestoreBackup?.click());

  fileRestoreBackup?.addEventListener("change", async () => {
    const file = fileRestoreBackup?.files?.[0];
    if (!file) return;

    try {
      const text = await file.text();
      const data = JSON.parse(text);

      // Validate
      if (!data._source || !data._exportedAt) {
        if (!confirm("This file doesn't look like a Magallanes backup.\nRestore anyway? This will overwrite all current data.")) return;
      }

      const counts = {
        Contracts:       (data.contracts    || []).length,
        "Cash Received": (data.cashReceived || []).length,
        "Cash Expense":  (data.cashExpense  || []).length,
        "Bank Received": (data.bankReceived || []).length,
        "Bank Expense":  (data.bankExpense  || []).length,
        "PNB Deposit":   (data.pnbDeposit   || []).length,
      };
      const summary = Object.entries(counts).map(([k,v]) => `  • ${k}: ${v} rows`).join("\n");

      if (!confirm(
        `⚠ Restore from Backup?\n\n` +
        `This will DELETE all current data in the database and replace it with:\n\n${summary}\n\n` +
        `Backup date: ${data._exportedAt ? new Date(data._exportedAt).toLocaleString() : "unknown"}\n\n` +
        `This cannot be undone. Continue?`
      )) return;

      btnRestoreBackup.disabled = true;
      btnRestoreBackup.textContent = "⏳ Restoring…";
      showBackupStatus("⏳ Restoring data — please wait, do not close this page…");

      // Helper: delete all rows in a table then insert fresh
      async function replaceTable(tableName, rows, saveRow) {
        // Delete all existing rows
        await _sb.from(tableName).delete().neq("id", "00000000-0000-0000-0000-000000000000");
        // Insert new rows in batches of 100
        for (let i = 0; i < rows.length; i += 100) {
          await Promise.all(rows.slice(i, i + 100).map(r => {
            const clean = { ...r };
            delete clean.id; // let Supabase generate new IDs
            return saveRow(clean);
          }));
        }
      }

      await replaceTable("contracts",    data.contracts    || [], r => DB.saveContract(r));
      await replaceTable("cash_received",data.cashReceived || [], r => DB.saveCashReceived(r));
      await replaceTable("cash_expense", data.cashExpense  || [], r => DB.saveCashExpense(r));
      await replaceTable("bank_received",data.bankReceived || [], r => DB.saveBankReceived(r));
      await replaceTable("bank_expense", data.bankExpense  || [], r => DB.saveBankExpense(r));
      await replaceTable("pnb_deposit",  data.pnbDeposit   || [], r => DB.savePnbDeposit(r));

      // Restore settings if present
      if (data.settings) {
        await DB.saveSettings({
          cashBalance:    data.settings.cash_balance    || data.settings.cashBalance    || 0,
          bankBalance:    data.settings.bank_balance    || data.settings.bankBalance    || 0,
          financeClerk:   data.settings.finance_clerk   || data.settings.financeClerk   || "",
          accountant:     data.settings.accountant      || "",
          financeManager: data.settings.finance_manager || data.settings.financeManager || "",
        });
      }

      showBackupStatus("✓ Restore complete! Reloading data…");

      // Reload all stores from Supabase
      const [newContracts, newCash, newCashExp, newBank, newBankExp, newPnb] = await Promise.all([
        DB.getContracts(), DB.getCashReceived(), DB.getCashExpense(),
        DB.getBankReceived(), DB.getBankExpense(), DB.getPnbDeposit(),
      ]);
      contractsStore = newContracts;
      cashStore      = newCash;
      cashExpStore   = newCashExp;
      bankStore      = newBank;
      bankExpStore   = newBankExp;
      pnbStore       = newPnb;

      renderContracts();
      renderCashTable();
      renderCashExpTable();
      renderBankTable();
      renderBankExpTable();
      renderPnbTable();
      loadSettings();

      showBackupStatus(
        `✓ Restore complete!  •  ` +
        `Contracts: ${newContracts.length}  •  ` +
        `Cash Received: ${newCash.length}  •  ` +
        `Cash Expense: ${newCashExp.length}  •  ` +
        `Bank Received: ${newBank.length}  •  ` +
        `Bank Expense: ${newBankExp.length}  •  ` +
        `PNB Deposit: ${newPnb.length}`
      );

    } catch (err) {
      showBackupStatus("✗ Restore failed: " + (err?.message || err), true);
    } finally {
      btnRestoreBackup.disabled = false;
      btnRestoreBackup.textContent = "⬆ Restore from Backup";
      if (fileRestoreBackup) fileRestoreBackup.value = "";
    }
  });

  // ── Fresh Restart ─────────────────────────────────────────────────────
  const btnFreshRestart   = document.querySelector("#btnFreshRestart");
  const freshRestartStatus = document.querySelector("#freshRestartStatus");

  function showFreshRestartStatus(msg, isError = false) {
    if (!freshRestartStatus) return;
    freshRestartStatus.textContent = msg;
    freshRestartStatus.style.display = "block";
    freshRestartStatus.style.background = isError ? "rgba(255,60,60,0.15)" : "rgba(60,200,100,0.15)";
    freshRestartStatus.style.color = isError ? "#ff6b6b" : "#5edb8a";
    freshRestartStatus.style.border = isError ? "1px solid rgba(255,60,60,0.3)" : "1px solid rgba(60,200,100,0.3)";
    clearTimeout(freshRestartStatus._timer);
    freshRestartStatus._timer = setTimeout(() => { freshRestartStatus.style.display = "none"; }, 8000);
  }

  btnFreshRestart?.addEventListener("click", async () => {
    const first = confirm(
      "⚠️ FRESH RESTART\n\n" +
      "This will permanently delete ALL data:\n" +
      "• Contracts\n• Cash Received\n• Cash Expenses\n• Bank Received\n• Bank Expenses\n• PNB Deposits\n• Settings\n\n" +
      "This CANNOT be undone.\n\nAre you absolutely sure you want to continue?"
    );
    if (!first) return;

    const second = confirm(
      "Last chance — are you sure?\n\nClick OK to delete everything and start fresh."
    );
    if (!second) return;

    btnFreshRestart.disabled = true;
    btnFreshRestart.textContent = "⏳ Deleting all data…";

    try {
      await Promise.all([
        DB.deleteAllContracts?.(),
        DB.deleteAllCashReceived?.(),
        DB.deleteAllCashExpense?.(),
        DB.deleteAllBankReceived?.(),
        DB.deleteAllBankExpense?.(),
        DB.deleteAllPnbDeposit?.(),
        DB.deleteAllDswd?.(),
        DB.deleteAllBai?.(),
        DB.deleteAllSettings?.(),
      ]);

      // Clear in-memory stores
      contractsStore = [];
      selectedContractNo = null;

      // Re-render all views
      renderContracts?.();

      showFreshRestartStatus("✓ All data deleted. The program has been reset to a fresh start.");
    } catch (err) {
      showFreshRestartStatus("✗ Fresh Restart failed: " + (err?.message || err), true);
    } finally {
      btnFreshRestart.disabled = false;
      btnFreshRestart.textContent = "🗑 Fresh Restart the Program";
    }
  });


  
    function drFmtBlankMoneyCR(){ return ''; }


// ---------------------------
  // Daily Report (v115) — Date selector + header
  // ---------------------------
const reportDatePicker = $("#reportDatePicker");

  reportDatePicker?.addEventListener("click", ()=>{
    try{ reportDatePicker.showPicker?.(); }catch{}
  });
  const drAsOfDate = $("#drAsOfDate");

  function yyyyMMddToMMDDYYYY_local(v){
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(v||""));
    return m ? `${m[2]}/${m[3]}/${m[1]}` : "";
  }

  function setDailyReportDate(yyyyMMdd){
    const nice = yyyyMMddToMMDDYYYY_local(yyyyMMdd);
    if (drAsOfDate) drAsOfDate.textContent = nice || "—";
    try{ localStorage.setItem("mf_daily_report_date_v115", String(yyyyMMdd||"")); }catch{}
  }

  function loadDailyReportDate(){
    try{
      const saved = localStorage.getItem("mf_daily_report_date_v115") || "";
      if (saved && reportDatePicker) reportDatePicker.value = saved;
      if (saved) setDailyReportDate(saved);
      else {
        // default to today
        const today = new Date().toISOString().slice(0,10);
        if (reportDatePicker) reportDatePicker.value = today;
        setDailyReportDate(today);
      }
    }catch{}
  }

  reportDatePicker?.addEventListener("change", ()=>{
    setDailyReportDate(reportDatePicker.value);
  });
  // also refresh report tables
  reportDatePicker?.addEventListener("change", ()=>{ try{ renderDailyReportCashReceived(); }catch{} });

  // init
  loadDailyReportDate();


  
  // ---------------------------
  // Daily Report (v115) — Cash Received table (step 1)
  // ---------------------------
  function drParseYmd(v){
    if(!v) return null;
    const s = String(v).trim();
    // YYYY-MM-DD
    if(/^\d{4}-\d{1,2}-\d{1,2}/.test(s)){
      const m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
      if(m){
        const mm = String(m[2]).padStart(2,'0');
        const dd = String(m[3]).padStart(2,'0');
        return `${m[1]}-${mm}-${dd}`;
      }
    }
    // MM/DD/YYYY
    const m2 = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if(m2){
      const mm = String(m2[1]).padStart(2,'0');
      const dd = String(m2[2]).padStart(2,'0');
      return `${m2[3]}-${mm}-${dd}`;
    }
    const d = new Date(s);
    if(Number.isFinite(d.getTime())) return d.toISOString().slice(0,10);
    return null;
  }

  function ensureDailyReportCashReceivedSection(){
    const panel = document.querySelector("#panel-dailyreport");
    const content = panel?.querySelector("#dailyReportContent");
    if(!content) return null;

    let el = content.querySelector("#drCashReceivedSection");
    if(el) return el;

    el = document.createElement("div");
    el.id = "drCashReceivedSection";

    // insert BEFORE signature section if it exists
    const sig = content.querySelector("#drSignatureSection");
    if(sig) content.insertBefore(el, sig);
    else content.appendChild(el);

    return el;
  }

  function getSelectedReportYmd(){
    const panel = document.querySelector("#panel-dailyreport");
    const picker = panel?.querySelector("#reportDatePicker");
    const ymd = drParseYmd(picker?.value) || new Date().toISOString().slice(0,10);
    return ymd;
  }

  function getSettingsCashStart(){
    const v = Number((document.querySelector("#setCashBalance")||{}).value || 0);
    return Number.isFinite(v) ? v : 0;
  }

  // Computes last known Cash On Hand BEFORE selected date using live transactions.
  // This is needed for "Balance Forwarded".
  function computeCashOnHandBefore(selectedYmd){
    // load stores (RETURN arrays)
    try{ const t = loadCashStore(); if(Array.isArray(t)) cashStore = t; }catch{}
    try{ const t = loadCashExpStore(); if(Array.isArray(t)) cashExpStore = t; }catch{}
    try{ const t = loadPnbStore(); if(Array.isArray(t)) pnbStore = t; }catch{}

    const dates = new Set();
    for(const r of (cashStore||[])){ const d=drParseYmd(r.date); if(d) dates.add(d); }
    for(const r of (cashExpStore||[])){ const d=drParseYmd(r.date); if(d) dates.add(d); }
    for(const r of (pnbStore||[])){ const d=drParseYmd(r.date); if(d) dates.add(d); }
    const sorted = Array.from(dates).sort();

    let cashOnHand = getSettingsCashStart();
    for(const d of sorted){
      if(d >= selectedYmd) break;
      const cashIn = (cashStore||[]).filter(r=>drParseYmd(r.date)===d).reduce((a,r)=>a+(Number(r.amount)||0),0);
      const pnb = (pnbStore||[]).filter(r=>drParseYmd(r.date)===d).reduce((a,r)=>a+(Number(r.amount)||0),0);
      const cashOut = (cashExpStore||[]).filter(r=>drParseYmd(r.date)===d).reduce((a,r)=>a+(Number(r.amount)||0),0);
      cashOnHand = (cashOnHand + cashIn - pnb) - cashOut;
    }
    return cashOnHand;
  }

  function renderDailyReportCashReceived(){
    const panel = document.querySelector("#panel-dailyreport");
    const content = panel?.querySelector("#dailyReportContent");
    if(!content) return;

    const selectedYmd = getSelectedReportYmd();
    setDailyReportDate(selectedYmd);

    // Load latest data (loaders RETURN arrays)
    try{ const t = loadCashStore(); if(Array.isArray(t)) cashStore = t; }catch(e){}
    try{ const t = loadCashExpStore(); if(Array.isArray(t)) cashExpStore = t; }catch(e){}
    try{ const t = loadBankStore(); if(Array.isArray(t)) bankStore = t; }catch(e){}
    try{ const t = loadBankExpStore(); if(Array.isArray(t)) bankExpStore = t; }catch(e){}
    try{ const t = loadPnbStore(); if(Array.isArray(t)) pnbStore = t; }catch(e){}

    // ---------------- Cash Received ----------------
    const cashBF = computeCashOnHandBefore(selectedYmd);
    const cashRows = (cashStore||[]).filter(r=>drParseYmd(r.date)===selectedYmd);
    const cashCollection = cashRows.reduce((a,r)=>a+(Number(r.amount)||0),0);
    const pnb = (pnbStore||[]).filter(r=>drParseYmd(r.date)===selectedYmd).reduce((a,r)=>a+(Number(r.amount)||0),0);
    const cashTotal = cashBF + cashCollection;
    const totalCashReceived = cashTotal - pnb;

    const cashBody = cashRows.map(r=>`
      <tr>
        <td>${escapeHtml(r.contract ?? r.contractNumber ?? "")}</td>
        <td>${escapeHtml(r.receipt ?? r.receiptNumber ?? "")}</td>
        <td>${escapeHtml(r.client ?? r.clientName ?? r.nameOfClient ?? "")}</td>
        <td>${escapeHtml(r.particular ?? "")}</td>
        <td style="text-align:right;">${fmtMoney(Number(r.amount)||0)}</td>
      </tr>
    `).join("");

    // ---------------- Cash Expense ----------------
    const cashExpRows = (cashExpStore||[]).filter(r=>drParseYmd(r.date)===selectedYmd);
    const totalCashExpenses = cashExpRows.reduce((a,r)=>a+(Number(r.amount)||0),0);
    const cashExpBody = cashExpRows.map(r=>`
      <tr>
        <td></td><td></td>
        <td>${escapeHtml(r.particular ?? "")}</td>
        <td></td>
        <td style="text-align:right;">${fmtMoney(Number(r.amount)||0)}</td>
      </tr>
    `).join("");

    // ---------------- Bank Received ----------------
    const bankBF = computeCashInBankBefore(selectedYmd);
    const bankRows = (bankStore||[]).filter(r=>drParseYmd(r.date)===selectedYmd);
    // Also include PNB deposit entries for the day — cash collected was deposited to bank
    const pnbRows = (pnbStore||[]).filter(r=>drParseYmd(r.date)===selectedYmd);
    // BAI completed entries — appear in Bank Received on their Date Completed
    const baiBankRows = (baiStore||[]).filter(r => r.status === "Completed" && drParseYmd(r.dateCompleted) === selectedYmd);
    const bankIn    = bankRows.reduce((a,r)=>a+(Number(r.amount)||0),0);
    const pnbBankIn = pnbRows.reduce((a,r)=>a+(Number(r.amount)||0),0);
    const baiBankIn = baiBankRows.reduce((a,r)=>a+(Number(r.amount)||0),0);
    const totalBankReceived = bankBF + bankIn + pnbBankIn + baiBankIn;

    const bankBody = [
      ...bankRows.map(r=>`
        <tr>
          <td>${escapeHtml(r.contract ?? r.contractNumber ?? "")}</td>
          <td>${escapeHtml(r.type ?? r.transfer ?? r.gcashOrBankTransfer ?? "")}</td>
          <td>${escapeHtml(r.client ?? r.clientName ?? r.nameOfClient ?? "")}</td>
          <td>${escapeHtml(r.particular ?? "")}</td>
          <td style="text-align:right;">${fmtMoney(Number(r.amount)||0)}</td>
        </tr>
      `),
      ...pnbRows.map(r=>`
        <tr>
          <td></td>
          <td>PNB Deposit</td>
          <td></td>
          <td>Cash deposit to PNB</td>
          <td style="text-align:right;">${fmtMoney(Number(r.amount)||0)}</td>
        </tr>
      `),
      ...baiBankRows.map(r=>`
        <tr>
          <td>${escapeHtml(r.contract||"")}</td>
          <td>BAI</td>
          <td></td>
          <td>BAI Assistance — Completed</td>
          <td style="text-align:right;">${fmtMoney(Number(r.amount)||0)}</td>
        </tr>
      `)
    ].join("");

    // ---------------- Bank Expense ----------------
    const bankExpRows = (bankExpStore||[]).filter(r=>drParseYmd(r.date)===selectedYmd);
    const totalBankExpense = bankExpRows.reduce((a,r)=>a+(Number(r.withdraw)||0),0);
    const bankExpBody = bankExpRows.map(r=>`
      <tr>
        <td>${escapeHtml(r.cv ?? "")}</td>
        <td>${escapeHtml(r.check ?? r.checkNumber ?? "")}</td>
        <td>${escapeHtml(r.particular ?? "")}</td>
        <td style="text-align:right;">${fmtMoney(Number(r.withdraw)||0)}</td>
        <td>${escapeHtml(r.others ?? "")}</td>
      </tr>
    `).join("");

    // ---------------- Summary ----------------
    const cashOnHand = totalCashReceived - totalCashExpenses;
    const cashInBank = totalBankReceived - totalBankExpense;

    content.innerHTML = `
      <div class="dr-section-spacer"></div>
      <div class="dr-section-spacer"></div>

      <div class="dr-section-title">Cash Received</div>
      <div class="dr-table-wrap">
        <table class="dr-table" id="drCashReceivedTable">
          <colgroup>
            <col class="dr-col1"><col class="dr-col2"><col class="dr-col3"><col class="dr-col4"><col class="dr-col5">
          </colgroup>
          <tbody>
            <tr>
              <td colspan="3"></td>
              <td style="text-align:right;font-weight:800;">Balance Forwarded :</td>
              <td style="text-align:right;font-weight:800;">${fmtMoney(cashBF)}</td>
            </tr>
            <tr class="dr-header-row">
              <td>Contract #</td><td>Receipt #</td><td>Name of Client</td><td>Particular</td><td>Amount</td>
            </tr>
            ${cashBody || `<tr><td colspan="5" class="dr-muted">(no matching Cash Received entries)</td></tr>`}
            <tr class="dr-total-row"><td>Cash collection for the day</td><td></td><td></td><td></td><td style="text-align:right;">${fmtMoney(cashCollection)}</td></tr>
            <tr class="dr-total-row"><td>Total</td><td></td><td></td><td></td><td style="text-align:right;">${fmtMoney(cashTotal)}</td></tr>
            <tr class="dr-total-row"><td>Less: PNB Deposit</td><td></td><td></td><td></td><td style="text-align:right;">${fmtMoney(pnb)}</td></tr>
            <tr class="dr-grand-row"><td>Total Cash Received</td><td></td><td></td><td></td><td style="text-align:right;">${fmtMoney(totalCashReceived)}</td></tr>
          </tbody>
        </table>
      </div>

      <div class="dr-section-spacer"></div>
      <div class="dr-section-title">Cash Expense</div>
      <div class="dr-table-wrap">
        <table class="dr-table" id="drCashExpenseTable">
          <colgroup>
            <col class="dr-col1"><col class="dr-col2"><col class="dr-col3"><col class="dr-col4"><col class="dr-col5">
          </colgroup>
          <tbody>
            <tr class="dr-header-row"><td></td><td></td><td>Particular</td><td></td><td>Amount</td></tr>
            ${cashExpBody || `<tr><td colspan="5" class="dr-muted">(no matching Cash Expense entries)</td></tr>`}
            <tr class="dr-grand-row"><td></td><td></td><td>Total Cash Expenses</td><td></td><td style="text-align:right;">${fmtMoney(totalCashExpenses)}</td></tr>
          </tbody>
        </table>
      </div>

      <div class="dr-section-spacer"></div>
      <div class="dr-section-spacer"></div>

      <div class="dr-section-title">Bank Received</div>
      <div class="dr-table-wrap">
        <table class="dr-table" id="drBankReceivedTable">
          <colgroup>
            <col class="dr-col1"><col class="dr-col2"><col class="dr-col3"><col class="dr-col4"><col class="dr-col5">
          </colgroup>
          <tbody>
            <tr>
              <td colspan="3"></td>
              <td style="text-align:right;font-weight:800;">Balance Forwarded :</td>
              <td style="text-align:right;font-weight:800;">${fmtMoney(bankBF)}</td>
            </tr>
            <tr class="dr-header-row">
              <td>Contract #</td><td>Gcash/Bank Transfer</td><td>Name of Client</td><td>Particular</td><td>Amount</td>
            </tr>
            ${bankBody || `<tr><td colspan="5" class="dr-muted">(no matching Bank Received entries)</td></tr>`}
            <tr class="dr-grand-row"><td></td><td></td><td></td><td>Total Bank Received</td><td style="text-align:right;">${fmtMoney(totalBankReceived)}</td></tr>
          </tbody>
        </table>
      </div>

      <div class="dr-section-spacer"></div>
      <div class="dr-section-title">Bank Expense</div>
      <div class="dr-table-wrap">
        <table class="dr-table" id="drBankExpenseTable">
          <colgroup>
            <col class="dr-col1"><col class="dr-col2"><col class="dr-col3"><col class="dr-col4"><col class="dr-col5">
          </colgroup>
          <tbody>
            <tr class="dr-header-row"><td>CV</td><td>Check #</td><td>Particular</td><td>Withdraw</td><td>Others</td></tr>
            ${bankExpBody || `<tr><td colspan="5" class="dr-muted">(no matching Bank Expense entries)</td></tr>`}
            <tr class="dr-grand-row"><td></td><td></td><td>Total Bank Expense</td><td style="text-align:right;">${fmtMoney(totalBankExpense)}</td><td></td></tr>
          </tbody>
        </table>
      </div>

      <div class="dr-section-spacer"></div>
      <div class="dr-section-spacer"></div>
      <div class="dr-table-wrap">
        <table class="dr-table" id="drSummaryTable">
          <colgroup>
            <col class="dr-col1"><col class="dr-col2"><col class="dr-col3"><col class="dr-col4"><col class="dr-col5">
          </colgroup>
          <tbody>
            <tr class="dr-grand-row"><td></td><td></td><td>Cash On Hand</td><td></td><td style="text-align:right;">${fmtMoney(cashOnHand)}</td></tr>
            <tr><td colspan="5" style="height:10px; border:none; background:transparent;"></td></tr>
            <tr class="dr-grand-row"><td></td><td></td><td>Cash in Bank</td><td></td><td style="text-align:right;">${fmtMoney(cashInBank)}</td></tr>
          </tbody>
        </table>
      </div>
    `;

    // Always re-render signature after report content is written
    try{ renderDailyReportSignatureBlock(); }catch(e){}
  }

  // Safe wiring: ensure DOM exists before binding
  (function wireDailyReportCashReceived(){
    const tryWire = ()=>{
      const panel = document.querySelector("#panel-dailyreport");
      if(!panel){ setTimeout(tryWire, 50); return; }

      const picker = panel.querySelector("#reportDatePicker");
      const refresh = panel.querySelector("#btnDailyReportRefresh");
      const rerender = ()=>{
        try{
          renderDailyReportCashReceived();
        }catch(e){
          console.warn(e);
          const panel = document.querySelector("#panel-dailyreport");
          const content = panel?.querySelector("#dailyReportContent");
          if(content){
            content.innerHTML = `<div class="dr-section-title">Cash Received</div>
              <div class="hint" style="margin:8px 0; color: var(--danger); font-weight:700;">
                Daily Report error: ${escapeHtml(e?.message || e)}
              </div>`;
          }
        }
      };

      picker?.addEventListener("change", rerender);
      picker?.addEventListener("input", rerender);
      refresh?.addEventListener("click", rerender);

      document.addEventListener("click", (ev)=>{
        const t = ev.target;
        if(t && t.classList && t.classList.contains("tab") && t.dataset.tab === "dailyreport"){
          setTimeout(rerender, 0);
        }
      });

      setTimeout(rerender, 0);
    };
    tryWire();
  })();

// ---------------------------
  // Daily Report (v115) — Signature Block
  // ---------------------------
  function ensureDailyReportSignatureSection(){
    const panel = document.querySelector("#panel-dailyreport");
    const content = panel?.querySelector("#dailyReportContent");
    if(!content) return null;
    let el = content.querySelector("#drSignatureSection");
    if(el) return el;
    el = document.createElement("div");
    el.id = "drSignatureSection";
    content.appendChild(el);
    return el;
  }

  function renderDailyReportSignatureBlock(){
    const sec = ensureDailyReportSignatureSection();
    if(!sec) return;

    const sig = (typeof getSignaturePersonnel === "function")
      ? getSignaturePersonnel()
      : { financeClerk:"Jennifer F. Landicho", accountant:"Ranni V. Dalisay", financeManager:"June Lizette M. Quizon" };

    sec.innerHTML = `
      <div class="dr-section-spacer"></div>
      <div class="dr-section-spacer"></div>

      <div class="dr-table-wrap dr-signature-block">
        <table class="dr-table dr-signature-table">
          <colgroup>
            <col style="width:18%">
            <col style="width:18%">
            <col style="width:26%">
            <col style="width:22%">
            <col style="width:16%">
          </colgroup>
          <tbody>
            <tr class="dr-signature-title">
              <td>Prepared by:</td>
              <td></td>
              <td>Check by</td>
              <td></td>
              <td>Noted by:</td>
            </tr>
            <tr><td colspan="5" style="height:18px; border:none; background:transparent;"></td></tr>
            <tr>
              <td>${escapeHtml(sig.financeClerk)}</td>
              <td></td>
              <td>${escapeHtml(sig.accountant)}</td>
              <td></td>
              <td>${escapeHtml(sig.financeManager)}</td>
            </tr>
            <tr>
              <td>Finance Clerk</td>
              <td></td>
              <td>Accountant</td>
              <td></td>
              <td>Finance Manager</td>
            </tr>
          </tbody>
        </table>
      </div>
    `;
  }

  (function wireDailyReportSignature(){
    const panel = document.querySelector("#panel-dailyreport");
    if(!panel) return;
    const picker     = panel.querySelector("#reportDatePicker");
    const refreshBtn = panel.querySelector("#btnDailyReportRefresh");

    // After any report render, re-append the signature (content.innerHTML wipe removes it)
    const renderAll = () => { try{ renderDailyReportSignatureBlock(); }catch(e){} };

    picker?.addEventListener("change",    () => setTimeout(renderAll, 80));
    refreshBtn?.addEventListener("click", () => setTimeout(renderAll, 80));

    // Show signature immediately when the Daily Report tab is clicked
    const tabBtn = document.querySelector(".tab[data-tab='dailyreport']");
    tabBtn?.addEventListener("click", () => setTimeout(renderAll, 80));

    // Render on page load — try a few times to ensure panel is ready
    setTimeout(renderAll, 150);
    setTimeout(renderAll, 500);
  })();


  

  function computeCashInBankBefore(selectedYmd){
    // Ensure latest stores are loaded
    try{ const t = loadBankStore(); if(Array.isArray(t)) bankStore = t; }catch{}
    try{ const t = loadBankExpStore(); if(Array.isArray(t)) bankExpStore = t; }catch{}
    try{ const t = loadPnbStore(); if(Array.isArray(t)) pnbStore = t; }catch{}

    // Starting value from Settings (Bank Received Balance)
    const bankEl = document.querySelector("#setBankBalance");
    let cashInBank = Number(bankEl?.value || 0);
    if(!Number.isFinite(cashInBank)) cashInBank = 0;

    const dates = new Set();
    for(const r of (bankStore||[])){ const d=drParseYmd(r.date); if(d) dates.add(d); }
    for(const r of (bankExpStore||[])){ const d=drParseYmd(r.date); if(d) dates.add(d); }
    for(const r of (pnbStore||[])){ const d=drParseYmd(r.date); if(d) dates.add(d); }
    const sorted = Array.from(dates).sort();

    for(const d of sorted){
      if(d >= selectedYmd) break;
      const bankIn  = (bankStore||[]).filter(r=>drParseYmd(r.date)===d).reduce((a,r)=>a+(Number(r.amount)||0),0);
      const pnbIn   = (pnbStore||[]).filter(r=>drParseYmd(r.date)===d).reduce((a,r)=>a+(Number(r.amount)||0),0);
      const bankOut = (bankExpStore||[]).filter(r=>drParseYmd(r.date)===d).reduce((a,r)=>a+(Number(r.withdraw)||0),0);
      cashInBank = (cashInBank + bankIn + pnbIn) - bankOut;
    }
    return cashInBank;
  }

  // ── Deselect row when clicking blank space outside any data row ──
  document.addEventListener("click", (e) => {
    const tables = [
      { el: $("#contractsTable"),  clearFn: () => { selectedContractNo = null; $$(".grid tbody tr").forEach(r => r.classList.remove("is-selected")); } },
      { el: cashTable,             clearFn: () => { selectCashKey(null); } },
      { el: cashExpTable,          clearFn: () => { cashExpSelectedKey = null; Array.from(cashExpTable?.tBodies[0]?.rows||[]).forEach(r=>r.classList.remove("is-selected")); if(cashExpSelectedEl) cashExpSelectedEl.textContent="Selected: —"; } },
      { el: bankTable,             clearFn: () => { selectBankKey(null); } },
      { el: bankExpTable,          clearFn: () => { bankExpSelectedKey = null; Array.from(bankExpTable?.tBodies[0]?.rows||[]).forEach(r=>r.classList.remove("is-selected")); if(bankExpSelectedEl) bankExpSelectedEl.textContent="Selected: —"; } },
      { el: pnbTable,              clearFn: () => { selectPnbKey(null); } },
      { el: dswdTable,             clearFn: () => { dswdSelectedKey = null; Array.from(dswdTable?.tBodies[0]?.rows||[]).forEach(r=>r.classList.remove("is-selected")); if(dswdSelectedEl) dswdSelectedEl.textContent="Selected: —"; } },
      { el: baiTable,              clearFn: () => { baiSelectedKey  = null; Array.from(baiTable?.tBodies[0]?.rows||[]).forEach(r=>r.classList.remove("is-selected")); if(baiSelectedEl)  baiSelectedEl.textContent ="Selected: —"; } },
    ];
    for (const { el, clearFn } of tables) {
      if (!el) continue;
      const wrap = el.closest(".grid-wrap") || el.closest(".grid-scroller") || el.parentElement;
      if (wrap && wrap.contains(e.target)) {
        const tr = e.target.closest("tr");
        if (!tr || (tr.dataset.rowType || "data") !== "data") clearFn();
        return;
      }
    }
  });

setTimeout(()=>{ try{ dr_recomputeDailyBalances(); }catch{} }, 0);

  function exportDailyReportToPdf(){
    const panel   = document.querySelector("#panel-dailyreport");
    const header  = panel?.querySelector("#dailyReportHeader");
    const content = panel?.querySelector("#dailyReportContent");

    if(!header || !content || !content.querySelector("table")){
      alert("Daily Report is not ready.\nPlease select a date and click Refresh first.");
      return;
    }

    // ── Build filename: MagallanesDailyReport_MM-DD-YYYY.pdf ──
    const picker = panel.querySelector("#reportDatePicker");
    let datePart = "NoDate";
    if(picker && picker.value){
      const [y,m,d] = picker.value.split("-");
      if(y && m && d) datePart = `${m}-${d}-${y}`;
    }
    const filename = `MagallanesDailyReport_${datePart}.pdf`;

    // ── Collect report text by walking the live DOM ──
    // We read text directly from the rendered elements so we never
    // touch any CSS color values at all — pure text + numbers only.

    function cellText(td){ return (td ? td.innerText.trim() : ""); }

    function tableToRows(tableEl){
      if(!tableEl) return [];
      const rows = [];
      tableEl.querySelectorAll("tr").forEach(function(tr){
        const cells = Array.from(tr.querySelectorAll("td,th")).map(function(c){ return c.innerText.trim(); });
        const cls = tr.className || "";
        rows.push({ cells: cells, cls: cls });
      });
      return rows;
    }

    // Header
    const titleText    = (header.querySelector(".dr-title")    || {}).innerText || "Magallanes Funeral Services";
    const subtitleText = (header.querySelector(".dr-subtitle") || {}).innerText || "Daily Collection / Expense Report";
    const asofText     = (header.querySelector(".dr-asof")     || {}).innerText || "";

    // Section titles
    function getSectionTitle(id){
      // Walk siblings — find .dr-section-title before the table with that id
      const tbl = content.querySelector("#" + id);
      if(!tbl) return "";
      let el = tbl.closest(".dr-table-wrap");
      if(!el) return "";
      let prev = el.previousElementSibling;
      while(prev){
        if(prev.classList && prev.classList.contains("dr-section-title")) return prev.innerText.trim();
        prev = prev.previousElementSibling;
      }
      return "";
    }

    const crTitle  = getSectionTitle("drCashReceivedTable")  || "Cash Received";
    const ceTitle  = getSectionTitle("drCashExpenseTable")   || "Cash Expense";
    const brTitle  = getSectionTitle("drBankReceivedTable")  || "Bank Received";
    const beTitle  = getSectionTitle("drBankExpenseTable")   || "Bank Expense";

    const crRows = tableToRows(content.querySelector("#drCashReceivedTable"));
    const ceRows = tableToRows(content.querySelector("#drCashExpenseTable"));
    const brRows = tableToRows(content.querySelector("#drBankReceivedTable"));
    const beRows = tableToRows(content.querySelector("#drBankExpenseTable"));
    const sumRows= tableToRows(content.querySelector("#drSummaryTable"));

    // Signature block
    const sigRows = tableToRows(content.querySelector(".dr-signature-table"));

    // ── Load jsPDF from CDN then draw ──
    const btn = document.querySelector("#btnExportDailyReportPdf");
    const origText = btn ? btn.textContent : "";
    if(btn){ btn.textContent = "⏳ Generating…"; btn.disabled = true; }

    function loadScript(src, cb){
      const existing = document.querySelector(`script[src="${src}"]`);
      if(existing && existing._mfLoaded){ cb(); return; }
      if(existing){ existing.addEventListener("load", cb); return; }
      const s = document.createElement("script");
      s.src = src;
      s.onload = function(){ s._mfLoaded = true; cb(); };
      s.onerror = function(){
        if(btn){ btn.textContent = origText; btn.disabled = false; }
        alert("Failed to load PDF library. Please check your internet connection.");
      };
      document.head.appendChild(s);
    }

    loadScript("https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js", function(){
      try{
        const { jsPDF } = window.jspdf;
        const doc = new jsPDF({ orientation:"portrait", unit:"mm", format:"a4" });

        const PW     = 210;   // A4 width mm
        const PH     = 297;   // A4 height mm
        const ML     = 12;    // left margin
        const MR     = 12;    // right margin
        const MT     = 14;    // top margin
        const MB     = 14;    // bottom margin
        const CW     = PW - ML - MR;  // content width = 186mm
        let   y      = MT;

        // ── helpers ──
        function checkNewPage(need){
          if(y + need > PH - MB){
            doc.addPage();
            y = MT;
          }
        }

        function drawText(text, x, yy, opts){
          doc.text(String(text || ""), x, yy, opts || {});
        }

        function sectionTitle(label){
          checkNewPage(10);
          doc.setFillColor(220,220,220);
          doc.rect(ML, y, CW, 6, "F");
          doc.setFont("helvetica","bold");
          doc.setFontSize(9);
          doc.setTextColor(0,0,0);
          drawText(label.toUpperCase(), ML + 2, y + 4.2);
          y += 8;
        }

        // col widths for the 5-col tables (sum = CW = 186)
        const COL5 = [30, 30, 56, 40, 30];
        // col widths for summary (blank | blank | label | blank | amount)
        const COLSUM = [20, 20, 66, 50, 30];

        function drawTableRow(cells, colWidths, opts){
          opts = opts || {};
          const bold     = opts.bold      || false;
          const fillRGB  = opts.fill      || null;
          const fontSize = opts.fontSize  || 8;
          const padH     = 1.8;  // horizontal padding mm inside each cell
          const padV     = 1.8;  // vertical padding mm top & bottom

          doc.setFont("helvetica", bold ? "bold" : "normal");
          doc.setFontSize(fontSize);

          const lineH = fontSize * 0.3528 * 1.25; // pt → mm × line-height factor

          // ── Step 1: measure how many lines each cell needs ──
          const cellLines = cells.map(function(cell, i){
            const cw   = (colWidths[i] || 20) - padH * 2;
            const txt  = String(cell || "");
            if(!txt) return [""];
            // splitTextToSize wraps text to fit cw mm
            return doc.splitTextToSize(txt, cw);
          });

          // ── Step 2: row height = tallest cell ──
          const maxLines = cellLines.reduce(function(m, lines){ return Math.max(m, lines.length); }, 1);
          const rowH = Math.max(6, padV * 2 + maxLines * lineH);

          checkNewPage(rowH + 1);

          // ── Step 3: fill background ──
          if(fillRGB){
            doc.setFillColor(fillRGB[0], fillRGB[1], fillRGB[2]);
            doc.rect(ML, y, CW, rowH, "F");
          }

          // ── Step 4: outer border rect ──
          doc.setDrawColor(100,100,100);
          doc.setLineWidth(0.2);
          doc.rect(ML, y, CW, rowH, "S");

          // ── Step 5: draw cell text + vertical dividers ──
          doc.setFont("helvetica", bold ? "bold" : "normal");
          doc.setFontSize(fontSize);
          doc.setTextColor(0,0,0);

          let cx = ML;
          cells.forEach(function(cell, i){
            const cw    = colWidths[i] || 20;

            // vertical divider (skip first — outer rect covers it)
            if(i > 0){
              doc.setDrawColor(100,100,100);
              doc.setLineWidth(0.2);
              doc.line(cx, y, cx, y + rowH);
            }

            const lines = cellLines[i];
            const isRight = (i === cells.length - 1) && /[\d,.]/.test(String(cell));
            const textX   = isRight ? cx + cw - padH : cx + padH;
            const align   = isRight ? "right" : "left";

            // Vertically centre the text block in the cell
            const blockH  = lines.length * lineH;
            let   textY   = y + (rowH - blockH) / 2 + lineH * 0.75;

            lines.forEach(function(line){
              doc.text(line, textX, textY, { align: align });
              textY += lineH;
            });

            cx += cw;
          });

          y += rowH;
        }

        function drawSection(titleLabel, rows, colWidths){
          if(!rows || rows.length === 0) return;
          sectionTitle(titleLabel);
          rows.forEach(function(row){
            const cls = row.cls;
            let fill = null;
            let bold = false;
            if(cls.indexOf("dr-header-row") >= 0){ fill=[200,200,200]; bold=true; }
            else if(cls.indexOf("dr-grand-row") >= 0){ fill=[185,215,185]; bold=true; }
            else if(cls.indexOf("dr-total-row") >= 0){ fill=[220,230,245]; bold=false; }
            // skip pure spacer rows (all empty cells, no border)
            const isEmpty = row.cells.every(function(c){ return c.trim()===""; });
            if(isEmpty && cls.indexOf("dr-grand-row") < 0 && cls.indexOf("dr-total-row") < 0) return;
            drawTableRow(row.cells, colWidths, { fill:fill, bold:bold });
          });
          y += 3; // small gap after section
        }

        // ── PAGE: Header ──
        doc.setFont("helvetica","bold");
        doc.setFontSize(16);
        doc.setTextColor(0,0,0);
        drawText(titleText, PW/2, y + 6, { align:"center" });
        y += 9;

        doc.setFont("helvetica","bold");
        doc.setFontSize(10);
        drawText(subtitleText, PW/2, y, { align:"center" });
        y += 6;

        doc.setFont("helvetica","normal");
        doc.setFontSize(9);
        drawText(asofText, PW/2, y, { align:"center" });
        y += 5;

        // Divider line
        doc.setDrawColor(0,0,0);
        doc.setLineWidth(0.5);
        doc.line(ML, y, ML + CW, y);
        y += 6;

        // ── Sections ──
        drawSection(crTitle, crRows, COL5);
        drawSection(ceTitle, ceRows, COL5);
        drawSection(brTitle, brRows, COL5);
        drawSection(beTitle, beRows, COL5);

        // Summary (no section title, just the grand rows)
        if(sumRows && sumRows.length > 0){
          y += 2;
          sumRows.forEach(function(row){
            const isEmpty = row.cells.every(function(c){ return c.trim()===""; });
            if(isEmpty) return;
            const cls = row.cls;
            const fill = (cls.indexOf("dr-grand-row") >= 0) ? [185,215,185] : null;
            drawTableRow(row.cells, COLSUM, { fill:fill, bold:true });
          });
          y += 4;
        }

        // ── Signature block ──
        if(sigRows && sigRows.length > 0){
          checkNewPage(30);
          y += 4;
          doc.setDrawColor(0,0,0);
          doc.setLineWidth(0.3);
          doc.line(ML, y, ML + CW, y);
          y += 6;

          const COLSIG = [CW/5, CW/5, CW/5, CW/5, CW/5];
          sigRows.forEach(function(row){
            const isEmpty = row.cells.every(function(c){ return c.trim()===""; });
            if(isEmpty){ y += 5; return; }
            const bold = row.cls.indexOf("dr-signature-title") >= 0;
            doc.setFont("helvetica", bold ? "bold" : "normal");
            doc.setFontSize(8);
            doc.setTextColor(0,0,0);
            let cx = ML;
            row.cells.forEach(function(cell, i){
              if(cell.trim()) drawText(cell.trim(), cx + 1, y);
              cx += COLSIG[i];
            });
            y += 5;
          });
        }

        doc.save(filename);

      } catch(err){
        alert("PDF generation failed: " + (err.message || err));
      } finally {
        if(btn){ btn.textContent = origText; btn.disabled = false; }
      }
    }); // jsPDF loaded
  }

  // Wire the Export PDF button
  (function wireDailyReportExportPdf(){
    const btn = document.querySelector("#btnExportDailyReportPdf");
    if(btn) btn.addEventListener("click", function(){ try{ exportDailyReportToPdf(); }catch(e){ alert("Export error: " + e.message); } });
  })();

})();
