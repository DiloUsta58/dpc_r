document.addEventListener("DOMContentLoaded", () => {
  const parseNumber = (value) => {
    const normalized = String(value).trim().replace(",", ".");
    const parsed = Number.parseFloat(normalized);
    return Number.isFinite(parsed) ? parsed : null;
  };

  const formatNumber = (value) => {
    if (!Number.isFinite(value)) {
      return "";
    }
    const rounded = Math.round(value * 100) / 100;
    if (Number.isInteger(rounded)) {
      return String(rounded);
    }
    return String(rounded).replace(".", ",");
  };

  const calculateBereitst = (istMaxRaw) => {
    const raw = String(istMaxRaw || "").trim();
    const parts = raw.split("/");
    if (parts.length !== 2) {
      return "";
    }

    const istPart = parts[0].trim();
    const maxPart = parts[1].trim();
    const max = parseNumber(maxPart);
    if (max === null) {
      return "";
    }

    if (istPart === "") {
      return "";
    }

    const ist = parseNumber(istPart);
    if (ist === null) {
      return "";
    }

    return formatNumber(max - ist);
  };

  const calculateOrDash = (istMaxRaw) => {
    const result = calculateBereitst(istMaxRaw);
    return result === "" ? "-" : result;
  };

  const clampIstToRange = (istRaw, maxValue) => {
    const normalized = String(istRaw || "").trim().replace(",", ".");
    if (normalized === "") {
      return "";
    }

    const istParsed = Number.parseFloat(normalized);
    if (!Number.isFinite(istParsed)) {
      return "";
    }

    const clamped = Math.min(maxValue, Math.max(0, istParsed));
    return formatNumber(clamped);
  };

  const deDateFormatter = new Intl.DateTimeFormat("de-DE", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric"
  });
  const weekdayFormatter = new Intl.DateTimeFormat("de-DE", { weekday: "short" });

  const toIsoLocal = (date) => {
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, "0");
    const d = String(date.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  };

  const fromIsoLocal = (iso) => {
    const [y, m, d] = iso.split("-").map(Number);
    return new Date(y, m - 1, d);
  };

  const getWeekMonday = (date) => {
    const d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
    const day = d.getDay();
    const diff = day === 0 ? -6 : 1 - day;
    d.setDate(d.getDate() + diff);
    return d;
  };

  const getIsoWeekNumber = (date) => {
    const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
    const dayNum = d.getUTCDay() || 7;
    d.setUTCDate(d.getUTCDate() + 4 - dayNum);
    const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  };

  const getIsoWeeksInYear = (year) => {
    const dec28 = new Date(year, 11, 28);
    return getIsoWeekNumber(dec28);
  };

  const getDateFromIsoWeek = (year, week, weekday) => {
    const jan4 = new Date(year, 0, 4);
    const jan4Day = jan4.getDay() || 7;
    const mondayWeek1 = new Date(year, 0, 4 - (jan4Day - 1));
    const target = new Date(mondayWeek1);
    target.setDate(mondayWeek1.getDate() + (week - 1) * 7 + (weekday - 1));
    return target;
  };

  const getWorkWeekByIso = (year, week) => {
    return [1, 2, 3, 4, 5].map((weekday) => getDateFromIsoWeek(year, week, weekday));
  };

  const todayDate = new Date();
  const todayIso = toIsoLocal(todayDate);
  const appVersion = "1.0.15";
  const appVersionFile = "app-version.json";
  const selectedDateStateKey = "dpc:selectedDate";
  let buildInfoCache = null;
  let selectedIso = todayIso;
  try {
    const storedSelectedIso = localStorage.getItem(selectedDateStateKey);
    if (/^\d{4}-\d{2}-\d{2}$/.test(String(storedSelectedIso || ""))) {
      selectedIso = String(storedSelectedIso);
    }
  } catch (error) {
    selectedIso = todayIso;
  }

  const dateTargets = document.querySelectorAll(".current-date");
  const renderSelectedDate = () => {
    const label = deDateFormatter.format(fromIsoLocal(selectedIso));
    dateTargets.forEach((el) => {
      el.textContent = label;
    });
  };

  const getBuildInfo = async () => {
    if (buildInfoCache) {
      return buildInfoCache;
    }

    try {
      const response = await fetch(`${appVersionFile}?t=${Date.now()}`, { cache: "no-store" });
      if (!response.ok) {
        return null;
      }
      const info = await response.json();
      if (!info || typeof info !== "object") {
        return null;
      }
      buildInfoCache = info;
      return buildInfoCache;
    } catch (error) {
      return null;
    }
  };

  const initFooter = async () => {
    const lastUpdateEl = document.getElementById("lastUpdate");
    const footerVersionEl = document.getElementById("footerVersion");
    const info = await getBuildInfo();

    if (lastUpdateEl) {
      const rawUpdatedAt = info && typeof info.updatedAt === "string" ? info.updatedAt : "";
      const parsedDate = rawUpdatedAt ? new Date(rawUpdatedAt) : null;
      if (parsedDate && Number.isFinite(parsedDate.getTime())) {
        const stamp = new Intl.DateTimeFormat("de-DE", {
          day: "2-digit",
          month: "2-digit",
          year: "numeric",
          hour: "2-digit",
          minute: "2-digit"
        }).format(parsedDate);
        lastUpdateEl.textContent = `Letztes Update: ${stamp}`;
      } else {
        lastUpdateEl.textContent = "Letztes Update: unbekannt";
      }
    }
    if (footerVersionEl) {
      const version = info && typeof info.version === "string" ? info.version : appVersion;
      footerVersionEl.textContent = `Version: ${version}`;
    }
  };

  const compareVersions = (a, b) => {
    const parse = (v) => String(v).split(".").map((part) => Number.parseInt(part, 10) || 0);
    const va = parse(a);
    const vb = parse(b);
    const maxLen = Math.max(va.length, vb.length);
    for (let i = 0; i < maxLen; i += 1) {
      const da = va[i] || 0;
      const db = vb[i] || 0;
      if (da > db) {
        return 1;
      }
      if (da < db) {
        return -1;
      }
    }
    return 0;
  };

  const checkForAppUpdate = async () => {
    try {
      const info = await getBuildInfo();
      if (!info || typeof info.version !== "string") {
        return;
      }

      if (compareVersions(info.version, appVersion) > 0) {
        const onceKey = `dpc:update-applied:${info.version}`;
        if (sessionStorage.getItem(onceKey) === "1") {
          return;
        }

        sessionStorage.setItem(onceKey, "1");
        const url = new URL(window.location.href);
        url.searchParams.set("appv", info.version);
        url.searchParams.set("upd", String(Date.now()));
        window.location.replace(url.toString());
      }
    } catch (error) {
      // Ignore update-check failures and continue app startup
    }
  };

  const persistSelectedDate = () => {
    try {
      localStorage.setItem(selectedDateStateKey, selectedIso);
    } catch (error) {
      // ignore storage errors in private mode / quota
    }
  };

  const readSelectedDateFromState = () => {
    try {
      const stored = localStorage.getItem(selectedDateStateKey);
      if (/^\d{4}-\d{2}-\d{2}$/.test(String(stored || ""))) {
        return String(stored);
      }
    } catch (error) {
      // ignore
    }
    return selectedIso;
  };

  checkForAppUpdate();

  let saveAutoWvorbeRows = () => {};
  const stockRowsData = [];
  const stockRows = document.querySelectorAll("table.tight tbody tr");
  stockRows.forEach((row) => {
    const cells = row.querySelectorAll("td");
    if (cells.length < 5) {
      return;
    }

    const istMaxInput = cells[1].querySelector("input[type='text']");
    const bereitstInput = cells[2].querySelector("input[type='text']");
    if (!istMaxInput || !bereitstInput) {
      return;
    }

    const initialValue = String(istMaxInput.value || "").trim();
    const fixedMaxMatch = initialValue.match(/^\/\s*([0-9]+(?:[.,][0-9]+)?)$/);
    const fixedMax = fixedMaxMatch ? fixedMaxMatch[1].replace(",", ".") : null;
    const table = row.closest("table");
    const tableTitleRaw = table ? (table.querySelector("thead th.section")?.textContent || "") : "";
    const tableTitle = tableTitleRaw.replace(/\s+/g, " ").trim().toUpperCase();
    const rowLabel = (cells[0].textContent || "").replace(/\s+/g, " ").trim();
    const rowName = rowLabel.toUpperCase();

    if (fixedMax) {
      istMaxInput.dataset.baseMax = fixedMax;
      istMaxInput.dataset.currentMax = fixedMax;
    }

    const istCell = cells[1];
    if (istCell && fixedMax) {
      const wrap = document.createElement("div");
      wrap.className = "istmax-wrap";

      const maxSuffix = document.createElement("span");
      maxSuffix.className = "max-suffix";
      maxSuffix.textContent = `/${fixedMax.replace(".", ",")}`;

      istMaxInput.classList.add("ist-only-input");
      istMaxInput.value = "";

      istCell.insertBefore(wrap, istMaxInput);
      wrap.appendChild(istMaxInput);
      wrap.appendChild(maxSuffix);
    }

    const lockMaxPart = () => {
      const currentMaxStr = istMaxInput.dataset.currentMax || istMaxInput.dataset.baseMax || "";
      const currentMaxValue = Number.parseFloat(String(currentMaxStr).replace(",", "."));
      if (!Number.isFinite(currentMaxValue)) {
        return;
      }

      const rawIst = String(istMaxInput.value || "");
      const safeIst = clampIstToRange(rawIst, currentMaxValue);
      istMaxInput.value = safeIst;
    };

    istMaxInput.addEventListener("input", lockMaxPart);
    istMaxInput.addEventListener("change", lockMaxPart);
    lockMaxPart();

    bereitstInput.readOnly = true;
    bereitstInput.tabIndex = -1;
    bereitstInput.classList.add("auto-bereitst");
    bereitstInput.value = "-";

    const syncBereitst = () => {
      const currentMaxStr = istMaxInput.dataset.currentMax || istMaxInput.dataset.baseMax || "";
      const combined = `${String(istMaxInput.value || "").trim()}/${String(currentMaxStr).replace(".", ",")}`;
      bereitstInput.value = calculateOrDash(combined);
      saveAutoWvorbeRows();
    };

    istMaxInput.addEventListener("input", syncBereitst);
    istMaxInput.addEventListener("change", syncBereitst);

    const bemerkInput = cells[4].querySelector("input[type='text']");
    const statusCheckInRow = row.querySelector(".status-check");
    if (bemerkInput) {
      const normalizeBemerkValue = (value) => {
        const raw = String(value || "").trim().toLowerCase();
        if (raw === "") {
          return "";
        }

        const compact = raw.replace(/\s+/g, "");
        if (compact === "nv" || compact === "n.v." || compact === "n.v") {
          return "n.v.";
        }

        return "";
      };

      const syncRemarkLockState = (normalizedValue) => {
        if (!statusCheckInRow) {
          return;
        }
        const isNotAvailable = normalizedValue === "n.v.";
        if (isNotAvailable && statusCheckInRow.checked) {
          statusCheckInRow.checked = false;
          statusCheckInRow.dispatchEvent(new Event("change", { bubbles: true }));
        }
        statusCheckInRow.disabled = isNotAvailable;
      };

      const syncBemerkState = () => {
        const value = String(bemerkInput.value || "").trim().toLowerCase();
        const isNotAvailable = value === "n.v.";
        row.classList.toggle("row-not-available", isNotAvailable);
        syncRemarkLockState(isNotAvailable ? "n.v." : "");
      };

      const validateBemerkValue = () => {
        const normalized = normalizeBemerkValue(bemerkInput.value);
        bemerkInput.value = normalized;
        row.classList.toggle("row-not-available", normalized === "n.v.");
        syncRemarkLockState(normalized);
        saveAutoWvorbeRows();
      };

      bemerkInput.addEventListener("input", syncBemerkState);
      bemerkInput.addEventListener("change", validateBemerkValue);
      bemerkInput.addEventListener("blur", validateBemerkValue);
      validateBemerkValue();
    }

    const statusCheck = row.querySelector(".status-check");

    stockRowsData.push({
      istMaxInput,
      bereitstInput,
      row,
      tableTitle,
      rowName,
      rowLabel,
      statusCheck
    });
  });

  const extractDeptCode = (tableTitle) => {
    const normalized = String(tableTitle || "").replace(/\s+/g, " ").trim().toUpperCase();
    const match = normalized.match(/ABT\.\s*:\s*([A-Z]+)/);
    return match ? match[1] : "";
  };

  const isExcludedForAutoWvorbe = (tableTitle) => {
    const t = String(tableTitle || "").toUpperCase();
    return t.includes("POWERGR") || t.includes("STRAHLHAUS");
  };

  const isFmRegalTable = (tableTitle) => {
    const t = String(tableTitle || "").toUpperCase();
    return /FM-\s*REGAL/.test(t);
  };

  const packSizeByPosition = {
    1: 250,
    2: 250,
    3: 250,
    4: 250,
    5: 50,
    6: 100,
    7: 100,
    8: 10,
    9: 20,
    10: 20,
    11: 100,
    12: 50,
    13: 50,
    14: 20,
    15: 20,
    16: 50,
    17: 250,
    18: 100,
    19: 20,
    20: 20,
    25: 5
  };

  const extractPositionNumber = (rowLabel) => {
    const match = String(rowLabel || "").toUpperCase().match(/POS\.?\s*(\d+)/);
    if (!match) {
      return null;
    }
    const parsed = Number.parseInt(match[1], 10);
    return Number.isFinite(parsed) ? parsed : null;
  };

  const buildAutoWvorbeRows = () => {
    const rows = [];
    const fmRegalTotals = new Map();
    stockRowsData.forEach((entry) => {
      if (isExcludedForAutoWvorbe(entry.tableTitle)) {
        return;
      }

      if (!entry.statusCheck || !entry.statusCheck.checked) {
        return;
      }

      const dept = extractDeptCode(entry.tableTitle);
      const material = String(entry.rowLabel || "").replace(/\s+/g, " ").trim();
      if (!dept || !material) {
        return;
      }

      const maxRaw = entry.istMaxInput.dataset.currentMax || entry.istMaxInput.dataset.baseMax || "";
      const maxValue = parseNumber(maxRaw);
      const istValue = parseNumber(entry.istMaxInput.value);
      if (maxValue === null || istValue === null) {
        return;
      }

      const diff = Math.max(0, maxValue - istValue);
      if (diff <= 0) {
        return;
      }

      const normalizedMaterial = material.toUpperCase();
      const isRhosealMaterial = normalizedMaterial === "RHOSEAL" || normalizedMaterial === "RHOSEAL HT";
      if (isRhosealMaterial) {
        const totalKg = diff * 25;
        rows.push([
          dept,
          material,
          "",
          "",
          `${formatNumber(totalKg)}kg`
        ]);
        return;
      }

      if (isFmRegalTable(entry.tableTitle)) {
        const posNo = extractPositionNumber(entry.rowLabel);
        if (posNo === null) {
          return;
        }
        const previous = fmRegalTotals.get(posNo) || 0;
        fmRegalTotals.set(posNo, previous + diff);
        return;
      }

      const integerDiff = Math.round(diff);
      if (Math.abs(diff - integerDiff) < 1e-9) {
        for (let i = 0; i < integerDiff; i += 1) {
          rows.push([
            dept,
            material,
            "",
            "",
            "1000kg"
          ]);
        }
        return;
      }

      rows.push([
        dept,
        material,
        "",
        "",
        `${formatNumber(diff)}`
      ]);
    });

    Array.from(fmRegalTotals.keys())
      .sort((a, b) => a - b)
      .forEach((posNo) => {
        const totalDiff = fmRegalTotals.get(posNo) || 0;
        if (totalDiff <= 0) {
          return;
        }

        const packSize = packSizeByPosition[posNo];
        if (!packSize) {
          return;
        }

        const totalPieces = totalDiff * packSize;
        rows.push([
          "FM",
          `Pos. ${posNo}`,
          "",
          "",
          `${formatNumber(totalPieces)} Stk.`
        ]);
      });

    return rows;
  };

  saveAutoWvorbeRows = () => {
    if (stockRowsData.length === 0) {
      return;
    }
    try {
      const rows = buildAutoWvorbeRows();
      const payload = {
        date: deDateFormatter.format(fromIsoLocal(selectedIso)),
        savedAt: new Date().toISOString(),
        source: "index-auto",
        rows
      };
      const autoKey = `dpc:auto:wvorbe:${selectedIso}`;
      localStorage.setItem(autoKey, JSON.stringify(payload));
    } catch (error) {
      // ignore storage errors
    }
  };

  const applyDateBasedDefaults = (isoDate, resetValues) => {
    const date = fromIsoLocal(isoDate);
    const isFriday = date.getDay() === 5;

    stockRowsData.forEach((entry) => {
      const baseMax = entry.istMaxInput.dataset.baseMax;
      if (!baseMax) {
        return;
      }

      let targetMax = baseMax;
      if (isFriday && entry.tableTitle.includes("ABT.: KE") && (entry.rowName === "RHOSEAL" || entry.rowName === "RHOSEAL HT")) {
        targetMax = "15";
      }
      if (isFriday && entry.tableTitle.includes("STRAHLHAUS")) {
        targetMax = "2";
      }

      entry.istMaxInput.dataset.currentMax = targetMax;
      const suffix = entry.istMaxInput.closest(".istmax-wrap")?.querySelector(".max-suffix");
      if (suffix) {
        suffix.textContent = `/${targetMax.replace(".", ",")}`;
      }
      if (resetValues) {
        entry.istMaxInput.value = "";
      }
      entry.istMaxInput.dispatchEvent(new Event("input", { bubbles: true }));
    });
    saveAutoWvorbeRows();
  };

  const statusChecks = document.querySelectorAll(".status-check");
  statusChecks.forEach((check) => {
    const row = check.closest("tr");
    if (!row) {
      return;
    }

    const syncRowState = () => {
      row.classList.toggle("row-done", check.checked);
      saveAutoWvorbeRows();
    };

    check.addEventListener("change", syncRowState);
    syncRowState();
  });

  const saveBtn = document.getElementById("saveBtn");
  const loadBtn = document.getElementById("loadBtn");
  const yearDownBtn = document.getElementById("yearDownBtn");
  const yearUpBtn = document.getElementById("yearUpBtn");
  const yearValue = document.getElementById("yearValue");
  const kwSelect = document.getElementById("kwSelect");
  const daySelect = document.getElementById("daySelect");

  if (kwSelect || daySelect) {
    selectedIso = todayIso;
  }
  const saveStatus = document.getElementById("saveStatus");
  const menuBtn = document.getElementById("menuBtn");
  const menuDropdown = document.getElementById("menuDropdown");
  const exportLocalFileBtn = document.getElementById("exportLocalFileBtn");
  const importLocalFileBtn = document.getElementById("importLocalFileBtn");
  const importFileInput = document.getElementById("importFileInput");
  const notesCells = Array.from(document.querySelectorAll(".notes-table td[contenteditable='true']"));
  const istInputs = Array.from(document.querySelectorAll("table.tight tbody tr td:nth-child(2) input[type='text']"));
  const checks = Array.from(document.querySelectorAll(".status-check"));
  const remarkInputs = Array.from(document.querySelectorAll("table.tight tbody tr td:nth-child(5) input[type='text']"));

  const getSnapshot = () => {
    const istMaxValues = istInputs.map((el) => {
      const max = String(el.dataset.currentMax || el.dataset.baseMax || "").replace(".", ",");
      const ist = String(el.value || "").trim();
      return `${ist}/${max}`;
    });
    const checksValues = checks.map((el) => el.checked);
    const remarkValues = remarkInputs.map((el) => el.value);
    const notes = notesCells.map((el) => el.textContent || "");

    return {
      date: deDateFormatter.format(fromIsoLocal(selectedIso)),
      savedAt: new Date().toISOString(),
      istMaxValues,
      checks: checksValues,
      remarkValues,
      notes,
      autoWvorbeRows: buildAutoWvorbeRows()
    };
  };

  if (menuBtn && menuDropdown) {
    const closeMenu = () => {
      menuDropdown.classList.remove("open");
      menuDropdown.setAttribute("aria-hidden", "true");
      menuBtn.setAttribute("aria-expanded", "false");
    };

    menuBtn.addEventListener("click", (event) => {
      event.stopPropagation();
      const isOpen = menuDropdown.classList.toggle("open");
      menuDropdown.setAttribute("aria-hidden", isOpen ? "false" : "true");
      menuBtn.setAttribute("aria-expanded", isOpen ? "true" : "false");
    });

    document.addEventListener("click", (event) => {
      if (!menuDropdown.contains(event.target) && event.target !== menuBtn) {
        closeMenu();
      }
    });
  }

  const applySnapshot = (data) => {
    if (!data || typeof data !== "object") {
      return;
    }

    if (Array.isArray(data.istMaxValues)) {
      istInputs.forEach((input, index) => {
        if (typeof data.istMaxValues[index] !== "string") {
          return;
        }
        const raw = data.istMaxValues[index];
        const parts = raw.split("/");
        if (parts.length === 2) {
          const nextIst = parts[0].trim();
          const nextMax = parts[1].trim().replace(",", ".");
          if (nextMax !== "") {
            input.dataset.currentMax = nextMax;
            const suffix = input.closest(".istmax-wrap")?.querySelector(".max-suffix");
            if (suffix) {
              suffix.textContent = `/${nextMax.replace(".", ",")}`;
            }
          }
          input.value = nextIst;
        } else {
          input.value = raw;
        }
        input.dispatchEvent(new Event("input", { bubbles: true }));
      });
    }

    if (Array.isArray(data.checks)) {
      checks.forEach((check, index) => {
        if (typeof data.checks[index] !== "boolean") {
          return;
        }
        check.checked = data.checks[index];
        check.dispatchEvent(new Event("change", { bubbles: true }));
      });
    }

    if (Array.isArray(data.remarkValues)) {
      remarkInputs.forEach((input, index) => {
        if (typeof data.remarkValues[index] !== "string") {
          return;
        }
        input.value = data.remarkValues[index];
        input.dispatchEvent(new Event("input", { bubbles: true }));
      });
    }

    if (Array.isArray(data.notes)) {
      notesCells.forEach((cell, index) => {
        if (typeof data.notes[index] !== "string") {
          return;
        }
        cell.textContent = data.notes[index];
      });
    }
    saveAutoWvorbeRows();
  };

  const setStatus = (text, isError) => {
    if (!saveStatus) {
      return;
    }
    saveStatus.textContent = text;
    saveStatus.classList.toggle("error", Boolean(isError));
  };

  const listDpcKeys = () => {
    const keys = [];
    for (let i = 0; i < localStorage.length; i += 1) {
      const key = localStorage.key(i);
      if (typeof key === "string" && key.startsWith("dpc:")) {
        keys.push(key);
      }
    }
    return keys;
  };

  const buildBackupPayload = () => {
    const storage = {};
    listDpcKeys().forEach((key) => {
      const value = localStorage.getItem(key);
      if (value !== null) {
        storage[key] = value;
      }
    });
    return {
      app: "dpc",
      version: 1,
      exportedAt: new Date().toISOString(),
      storage
    };
  };

  const applyBackupPayload = (payload) => {
    if (!payload || typeof payload !== "object" || !payload.storage || typeof payload.storage !== "object") {
      throw new Error("Ungültiges Backup-Format");
    }

    listDpcKeys().forEach((key) => localStorage.removeItem(key));

    Object.entries(payload.storage).forEach(([key, value]) => {
      if (!String(key).startsWith("dpc:")) {
        return;
      }
      localStorage.setItem(key, typeof value === "string" ? value : JSON.stringify(value));
    });
  };

  if (exportLocalFileBtn) {
    exportLocalFileBtn.addEventListener("click", () => {
      try {
        const payload = buildBackupPayload();
        const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
        const url = URL.createObjectURL(blob);
        const stamp = new Date().toISOString().replace(/[:T]/g, "-").slice(0, 16);
        const a = document.createElement("a");
        a.href = url;
        a.download = `dpc-backup-${stamp}.json`;
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(url);
        setStatus("Lokale JSON-Datei gespeichert", false);
      } catch (error) {
        setStatus("Datei speichern fehlgeschlagen", true);
      }
    });
  }

  if (importLocalFileBtn && importFileInput) {
    importLocalFileBtn.addEventListener("click", () => {
      importFileInput.click();
    });

    importFileInput.addEventListener("change", async () => {
      const file = importFileInput.files && importFileInput.files[0];
      if (!file) {
        return;
      }

      try {
        const text = await file.text();
        const payload = JSON.parse(text);
        applyBackupPayload(payload);
        setStatus("Lokale JSON-Datei geladen", false);
        window.location.reload();
      } catch (error) {
        setStatus("Datei laden fehlgeschlagen", true);
      } finally {
        importFileInput.value = "";
      }
    });
  }

  const setReadonlyForPastDays = () => {
    const isPastDay = selectedIso < todayIso;

    istInputs.forEach((input) => {
      input.readOnly = isPastDay;
      input.classList.toggle("locked-field", isPastDay);
    });

    checks.forEach((check) => {
      check.disabled = isPastDay;
      check.classList.toggle("locked-field", isPastDay);
    });

    remarkInputs.forEach((input) => {
      input.readOnly = isPastDay;
      input.classList.toggle("locked-field", isPastDay);
    });

    notesCells.forEach((cell) => {
      cell.contentEditable = isPastDay ? "false" : "true";
      cell.classList.toggle("locked-field", isPastDay);
    });

    if (saveBtn) {
      saveBtn.disabled = isPastDay;
    }
  };

  const getStorageKey = (iso) => `dpc:index:${iso}`;
  const defaultSnapshot = getSnapshot();

  const loadFromStorage = () => {
    const storageKey = getStorageKey(selectedIso);
    applySnapshot(defaultSnapshot);
    applyDateBasedDefaults(selectedIso, true);

    const raw = localStorage.getItem(storageKey);
    const selectedLabel = deDateFormatter.format(fromIsoLocal(selectedIso));
    if (raw) {
      applySnapshot(JSON.parse(raw));
      applyDateBasedDefaults(selectedIso, false);
      setStatus(`Daten von ${selectedLabel} geladen`, false);
      return true;
    }
    setStatus(`Keine Daten für ${selectedLabel}`, true);
    return false;
  };

  const minYear = 2026;
  const selectedDateInitial = fromIsoLocal(selectedIso);
  let selectedYear = Math.max(minYear, selectedDateInitial.getFullYear());
  const currentIsoWeek = getIsoWeekNumber(todayDate);
  let selectedWeek = getIsoWeekNumber(selectedDateInitial);

  const fillDaysForKw = (year, week) => {
    if (!daySelect) {
      return;
    }

    const weekDays = getWorkWeekByIso(year, week);
    daySelect.innerHTML = "";

    weekDays.forEach((date) => {
      const iso = toIsoLocal(date);
      const weekday = weekdayFormatter.format(date).replace(".", "");
      const label = `${weekday} ${deDateFormatter.format(date)}`;
      const option = document.createElement("option");
      option.value = iso;
      option.textContent = label;
      daySelect.appendChild(option);
    });

    const hasSelected = weekDays.some((d) => toIsoLocal(d) === selectedIso);
    const hasToday = weekDays.some((d) => toIsoLocal(d) === todayIso);
    if (hasSelected) {
      selectedIso = selectedIso;
    } else if (hasToday && year === todayDate.getFullYear()) {
      selectedIso = todayIso;
    } else {
      selectedIso = toIsoLocal(weekDays[0]);
    }
    daySelect.value = selectedIso;
  };

  const fillKwForYear = (year) => {
    if (!kwSelect) {
      return;
    }

    const weeksInYear = getIsoWeeksInYear(year);
    kwSelect.innerHTML = "";

    for (let week = 1; week <= weeksInYear; week += 1) {
      const option = document.createElement("option");
      option.value = String(week);
      option.textContent = `KW ${String(week).padStart(2, "0")}`;
      kwSelect.appendChild(option);
    }

    if (selectedWeek > weeksInYear) {
      selectedWeek = weeksInYear;
    }
    if (selectedWeek < 1) {
      selectedWeek = 1;
    }

    kwSelect.value = String(selectedWeek);
    fillDaysForKw(year, selectedWeek);
  };

  const renderYear = () => {
    if (yearValue) {
      yearValue.textContent = String(selectedYear);
    }
    if (yearDownBtn) {
      yearDownBtn.disabled = selectedYear <= minYear;
    }
  };

  if (kwSelect) {
    if (selectedYear === todayDate.getFullYear()) {
      selectedWeek = currentIsoWeek;
    } else {
      selectedWeek = 1;
    }
    fillKwForYear(selectedYear);
  } else if (daySelect) {
    fillDaysForKw(selectedYear, selectedWeek);
  }

  renderYear();

  renderSelectedDate();
  initFooter();
  persistSelectedDate();
  setReadonlyForPastDays();

  try {
    loadFromStorage();
  } catch (error) {
    setStatus("Laden fehlgeschlagen", true);
  }

  if (saveBtn) {
    saveBtn.addEventListener("click", () => {
      try {
        const storageKey = getStorageKey(selectedIso);
        localStorage.setItem(storageKey, JSON.stringify(getSnapshot()));
        saveAutoWvorbeRows();
        setStatus(`Gespeichert (${deDateFormatter.format(fromIsoLocal(selectedIso))})`, false);
      } catch (error) {
        setStatus("Speichern fehlgeschlagen", true);
      }
    });
  }

  if (loadBtn) {
    loadBtn.addEventListener("click", () => {
      try {
        loadFromStorage();
      } catch (error) {
        setStatus("Laden fehlgeschlagen", true);
      }
    });
  }

  if (daySelect) {
    daySelect.addEventListener("change", () => {
      selectedIso = daySelect.value;
      renderSelectedDate();
      persistSelectedDate();
      setReadonlyForPastDays();
      try {
        loadFromStorage();
      } catch (error) {
        setStatus("Laden fehlgeschlagen", true);
      }
    });
  }

  if (kwSelect) {
    kwSelect.addEventListener("change", () => {
      const parsedWeek = Number.parseInt(kwSelect.value, 10);
      if (!Number.isFinite(parsedWeek) || parsedWeek < 1) {
        return;
      }

      selectedWeek = parsedWeek;
      fillDaysForKw(selectedYear, selectedWeek);
      renderSelectedDate();
      persistSelectedDate();
      setReadonlyForPastDays();
      try {
        loadFromStorage();
      } catch (error) {
        setStatus("Laden fehlgeschlagen", true);
      }
    });
  }

  if (yearDownBtn) {
    yearDownBtn.addEventListener("click", () => {
      if (selectedYear <= minYear) {
        return;
      }
      selectedYear -= 1;
      selectedWeek = 1;
      renderYear();
      fillKwForYear(selectedYear);
      renderSelectedDate();
      persistSelectedDate();
      setReadonlyForPastDays();
      try {
        loadFromStorage();
      } catch (error) {
        setStatus("Laden fehlgeschlagen", true);
      }
    });
  }

  if (yearUpBtn) {
    yearUpBtn.addEventListener("click", () => {
      selectedYear += 1;
      if (selectedYear === todayDate.getFullYear()) {
        selectedWeek = currentIsoWeek;
      } else {
        selectedWeek = 1;
      }
      renderYear();
      fillKwForYear(selectedYear);
      renderSelectedDate();
      persistSelectedDate();
      setReadonlyForPastDays();
      try {
        loadFromStorage();
      } catch (error) {
        setStatus("Laden fehlgeschlagen", true);
      }
    });
  }

  const initSimpleTablePage = (config) => {
    const table = document.getElementById(config.tableId);
    if (!table) {
      return;
    }

    const saveBtnLocal = document.getElementById(config.saveBtnId);
    const loadBtnLocal = document.getElementById(config.loadBtnId);
    const statusEl = document.getElementById(config.statusId);
    const addRowBtn = document.getElementById(config.addRowBtnId);
    const backBtn = document.getElementById(config.backBtnId);
    const addRowLine = table.querySelector(".add-row-line");
    let hasUnsavedChanges = false;
    const initialRows = Array.from(table.querySelectorAll("tbody tr:not(.add-row-line)")).map((row) =>
      Array.from(row.querySelectorAll("td")).map((cell) => cell.textContent || "")
    );

    const setStatusLocal = (text, isError) => {
      if (!statusEl) {
        return;
      }
      statusEl.textContent = text;
      statusEl.classList.toggle("error", Boolean(isError));
    };

    const setDirty = (dirty) => {
      hasUnsavedChanges = dirty;
    };

    const createRow = (values, isAutoGenerated = false) => {
      const row = document.createElement("tr");
      if (isAutoGenerated) {
        row.dataset.autoGenerated = "1";
      }
      for (let i = 0; i < 5; i += 1) {
        const td = document.createElement("td");
        td.textContent = Array.isArray(values) && typeof values[i] === "string" ? values[i] : "";
        row.appendChild(td);
      }
      return row;
    };

    const setReadonlyLocal = () => {
      const isPastDay = selectedIso < todayIso;
      const editableRows = table.querySelectorAll("tbody tr:not(.add-row-line)");
      editableRows.forEach((row) => {
        row.querySelectorAll("td").forEach((cell) => {
          cell.contentEditable = isPastDay ? "false" : "true";
          cell.classList.toggle("locked-field", isPastDay);
        });
      });

      if (addRowBtn) {
        addRowBtn.disabled = isPastDay;
      }
      if (saveBtnLocal) {
        saveBtnLocal.disabled = isPastDay;
      }
    };

    const applyRows = (rows, autoRows = []) => {
      const existingRows = Array.from(table.querySelectorAll("tbody tr:not(.add-row-line)"));
      existingRows.forEach((row) => row.remove());

      const hasManual = Array.isArray(rows) && rows.length > 0;
      const hasAuto = Array.isArray(autoRows) && autoRows.length > 0;
      const sourceRows = hasManual ? rows : (hasAuto ? [] : initialRows);
      sourceRows.forEach((values) => {
        const row = createRow(values, false);
        if (addRowLine && addRowLine.parentNode) {
          addRowLine.parentNode.insertBefore(row, addRowLine);
        }
      });

      if (hasAuto) {
        autoRows.forEach((values) => {
          const row = createRow(values, true);
          if (addRowLine && addRowLine.parentNode) {
            addRowLine.parentNode.insertBefore(row, addRowLine);
          }
        });
      }

      setReadonlyLocal();
    };

    const getRows = (excludeAutoGenerated) => {
      const rows = Array.from(table.querySelectorAll("tbody tr:not(.add-row-line)"));
      const filteredRows = excludeAutoGenerated
        ? rows.filter((row) => row.dataset.autoGenerated !== "1")
        : rows;
      return filteredRows.map((row) => Array.from(row.querySelectorAll("td")).map((cell) => cell.textContent || ""));
    };

    const normalizeRows = (rows) => {
      if (!Array.isArray(rows)) {
        return [];
      }
      return rows.filter((row) => Array.isArray(row)).map((row) => row.map((cell) => String(cell || "")));
    };

    const getRowKey = (row) => {
      if (!Array.isArray(row)) {
        return "";
      }
      return row.map((cell) => String(cell || "").trim().toLowerCase()).join("|");
    };

    const getMaterialKey = (row) => {
      if (!Array.isArray(row)) {
        return "";
      }
      const abt = String(row[0] || "").trim().toLowerCase();
      const material = String(row[1] || "").trim().toLowerCase();
      if (!abt || !material) {
        return "";
      }
      return `${abt}|${material}`;
    };

    const getStorageKeyLocal = () => `dpc:${config.storagePrefix}:${selectedIso}`;
    const getAutoStorageKeyLocal = () => config.autoStoragePrefix ? `dpc:auto:${config.autoStoragePrefix}:${selectedIso}` : "";
    const syncSelectedDateLocal = () => {
      selectedIso = readSelectedDateFromState();
      renderSelectedDate();
    };

    const hasMeaningfulRows = (rows) => {
      if (!Array.isArray(rows)) {
        return false;
      }
      return rows.some((row) => Array.isArray(row) && row.some((cell) => String(cell || "").trim() !== ""));
    };

    const loadLocal = () => {
      syncSelectedDateLocal();
      const key = getStorageKeyLocal();
      const autoKey = getAutoStorageKeyLocal();
      const indexKey = `dpc:index:${selectedIso}`;
      const selectedLabel = deDateFormatter.format(fromIsoLocal(selectedIso));
      const raw = localStorage.getItem(key);
      let autoRows = null;
      if (autoKey) {
        const autoRaw = localStorage.getItem(autoKey);
        if (autoRaw) {
          const autoData = JSON.parse(autoRaw);
          autoRows = Array.isArray(autoData.rows) ? autoData.rows : [];
        }
      }

      if (config.storagePrefix === "wvorbe" && !hasMeaningfulRows(autoRows)) {
        const indexRaw = localStorage.getItem(indexKey);
        if (indexRaw) {
          const indexData = JSON.parse(indexRaw);
          const indexAutoRows = Array.isArray(indexData.autoWvorbeRows) ? indexData.autoWvorbeRows : [];
          if (hasMeaningfulRows(indexAutoRows)) {
            autoRows = indexAutoRows;
          }
        }
      }

      const normalizedAutoRows = normalizeRows(autoRows);

      if (raw) {
        const data = JSON.parse(raw);
        const manualRows = Array.isArray(data.rows) ? data.rows : [];
        if (hasMeaningfulRows(manualRows)) {
          if (config.storagePrefix === "wvorbe" && hasMeaningfulRows(normalizedAutoRows)) {
            const normalizedManualRows = normalizeRows(manualRows);
            const manualKeys = new Set(normalizedManualRows.map((row) => getRowKey(row)).filter((key) => key !== ""));
            const lockedMaterialKeys = new Set(
              normalizedManualRows
                .filter((row) => String(row[2] || "").trim() !== "" || String(row[3] || "").trim() !== "")
                .map((row) => getMaterialKey(row))
                .filter((key) => key !== "")
            );
            const filteredAutoRows = normalizedAutoRows.filter((row) => {
              const rowKey = getRowKey(row);
              const materialKey = getMaterialKey(row);
              if (manualKeys.has(rowKey)) {
                return false;
              }
              if (materialKey && lockedMaterialKeys.has(materialKey)) {
                return false;
              }
              return true;
            });
            applyRows(manualRows, filteredAutoRows);
            setStatusLocal(`Daten + Auto-Daten von ${selectedLabel} geladen`, false);
            return true;
          }
          applyRows(manualRows);
          setStatusLocal(`Daten von ${selectedLabel} geladen`, false);
          return true;
        }

        if (hasMeaningfulRows(normalizedAutoRows)) {
          applyRows(normalizedAutoRows);
          setStatusLocal(`Auto-Daten von ${selectedLabel} geladen`, false);
          return true;
        }
      }

      if (hasMeaningfulRows(normalizedAutoRows)) {
        applyRows(normalizedAutoRows);
        setStatusLocal(`Auto-Daten von ${selectedLabel} geladen`, false);
        return true;
      }

      applyRows([]);
      setStatusLocal(`Keine Daten für ${selectedLabel}`, true);
      return false;
    };

    const saveLocal = () => {
      syncSelectedDateLocal();
      const key = getStorageKeyLocal();
      const payload = {
        date: deDateFormatter.format(fromIsoLocal(selectedIso)),
        savedAt: new Date().toISOString(),
        rows: getRows(config.storagePrefix === "wvorbe")
      };
      localStorage.setItem(key, JSON.stringify(payload));
      setStatusLocal(`Gespeichert (${payload.date})`, false);
      setDirty(false);
    };

    applyRows(initialRows);
    try {
      loadLocal();
      setDirty(false);
    } catch (error) {
      setStatusLocal("Laden fehlgeschlagen", true);
    }

    if (addRowBtn) {
      addRowBtn.addEventListener("click", () => {
        const newRow = createRow([], false);
        if (addRowLine && addRowLine.parentNode) {
          addRowLine.parentNode.insertBefore(newRow, addRowLine);
          const isPastDay = selectedIso < todayIso;
          newRow.querySelectorAll("td").forEach((cell) => {
            cell.contentEditable = isPastDay ? "false" : "true";
            cell.classList.toggle("locked-field", isPastDay);
          });
          const firstCell = newRow.querySelector("td");
          if (firstCell && !isPastDay) {
            firstCell.focus();
          }
        }
        setDirty(true);
      });
    }

    table.addEventListener("input", (event) => {
      if (event.target instanceof HTMLElement && event.target.closest("td[contenteditable='true']")) {
        const row = event.target.closest("tr");
        if (row && row.dataset.autoGenerated === "1") {
          delete row.dataset.autoGenerated;
        }
        setDirty(true);
      }
    });

    if (saveBtnLocal) {
      saveBtnLocal.addEventListener("click", () => {
        try {
          saveLocal();
        } catch (error) {
          setStatusLocal("Speichern fehlgeschlagen", true);
        }
      });
    }

    if (loadBtnLocal) {
      loadBtnLocal.addEventListener("click", () => {
        try {
          loadLocal();
          setDirty(false);
        } catch (error) {
          setStatusLocal("Laden fehlgeschlagen", true);
        }
      });
    }

    const tryLeavePage = () => {
      window.location.href = "index.html";
    };

    if (backBtn) {
      backBtn.addEventListener("click", () => {
        syncSelectedDateLocal();
        if (!hasUnsavedChanges) {
          tryLeavePage();
          return;
        }

        const saveNow = window.confirm("Ungespeicherte Änderungen gefunden. Jetzt speichern?");
        if (saveNow) {
          try {
            saveLocal();
            tryLeavePage();
          } catch (error) {
            setStatusLocal("Speichern fehlgeschlagen", true);
          }
          return;
        }

        const leaveWithoutSave = window.confirm("Ohne Speichern zurück zur Startseite?");
        if (leaveWithoutSave) {
          tryLeavePage();
        }
      });
    }

    window.addEventListener("beforeunload", (event) => {
      if (!hasUnsavedChanges) {
        return;
      }
      event.preventDefault();
      event.returnValue = "";
    });

    window.addEventListener("storage", (event) => {
      if (event.key !== selectedDateStateKey) {
        return;
      }
      try {
        loadLocal();
      } catch (error) {
        setStatusLocal("Laden fehlgeschlagen", true);
      }
    });
  };

  initSimpleTablePage({
    tableId: "wvorbeTable",
    saveBtnId: "wvorbeSaveBtn",
    loadBtnId: "wvorbeLoadBtn",
    statusId: "wvorbeStatus",
    addRowBtnId: "addWvorbeRowBtn",
    backBtnId: "wvorbeBackBtn",
    storagePrefix: "wvorbe",
    autoStoragePrefix: "wvorbe"
  });

  initSimpleTablePage({
    tableId: "weingTable",
    saveBtnId: "weingSaveBtn",
    loadBtnId: "weingLoadBtn",
    statusId: "weingStatus",
    addRowBtnId: "addWeingRowBtn",
    backBtnId: "weingBackBtn",
    storagePrefix: "weing"
  });
});
