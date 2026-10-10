/**
 * Precision Injection Monitor - Core Application Engine
 * Daily Parameter Tracking, 7-Zone Thermal Modeling (Setting vs Actual) & Excel Integration
 */

// Global App State
const state = {
  rawDataset: [],        // All loaded records from Excel / Storage
  filteredDataset: [],   // Currently active filtered view
  selectedMachine: 'ALL',
  selectedModel: 'ALL',
  selectedRange: '30',   // '30', '60', '120', '360', 'ALL'
  activeMetric: 'pressure', // 'pressure', 'cushion', 'fillTime', 'cycleTime'
  searchQuery: ''
};

// Global Chart References
let trendChart = null;
const thermalCharts = {
  nozzle: null,
  front: null,
  middle: null,
  rear: null,
  hopper: null,
  core: null,
  cavity: null
};

// Initial Dataset populated directly from 20261007_injection monitor_R1.xlsx
const INITIAL_DEMO_DATA = [
  {
    machine: 'B01',
    model: 'AA551',
    date: '2026-08-20',
    dateDisplay: '8/20/2026',
    setSpeed1: 100, setSpeed2: 60, setSpeed3: null,
    setShotSize: 32, setPos1: 23, setPos2: null, setPos3: null,
    vpChange: 6.0,
    setPeakPressure: 200,
    setCoolingTime: 25,
    // Setting Temps
    setTempNozzle: 320, setTempFront: 310, setTempMiddle: 300, setTempRear: 280,
    setTempHopper: 100, setTempCore: 150, setTempCavity: 150,
    // Actual Process
    fillTime: 0.86, cushion: 4.20, recoveryTime: 12.58, peakPressure: 186, cycleTime: 33.0,
    // Actual Temps
    actTempNozzle: 320.4, actTempFront: 310.2, actTempMiddle: 299.9, actTempRear: 280.2,
    actTempHopper: 99.9, actTempCore: 152.3, actTempCavity: 151.2
  },
  {
    machine: 'V01',
    model: 'AA541',
    date: '2026-08-20',
    dateDisplay: '8/20/2026',
    setSpeed1: 120, setSpeed2: 80, setSpeed3: null,
    setShotSize: 45, setPos1: 35, setPos2: null, setPos3: null,
    vpChange: 6.0,
    setPeakPressure: 240,
    setCoolingTime: 32,
    // Setting Temps
    setTempNozzle: 325, setTempFront: 310, setTempMiddle: 300, setTempRear: 280,
    setTempHopper: 100, setTempCore: 160, setTempCavity: 160,
    // Actual Process
    fillTime: 1.20, cushion: 4.55, recoveryTime: 10.80, peakPressure: 215, cycleTime: 55.0,
    // Actual Temps
    actTempNozzle: 325.1, actTempFront: 310.7, actTempMiddle: 299.7, actTempRear: 280.3,
    actTempHopper: 99.7, actTempCore: 160.4, actTempCavity: 158.9
  },
  {
    machine: 'B01',
    model: 'AA551',
    date: '2026-08-21',
    dateDisplay: '8/21/2026',
    setSpeed1: 100, setSpeed2: 60, setSpeed3: null,
    setShotSize: 32, setPos1: 23, setPos2: null, setPos3: null,
    vpChange: 6.0,
    setPeakPressure: 200,
    setCoolingTime: 25,
    // Setting Temps
    setTempNozzle: 320, setTempFront: 310, setTempMiddle: 300, setTempRear: 280,
    setTempHopper: 100, setTempCore: 150, setTempCavity: 150,
    // Actual Process
    fillTime: 0.87, cushion: 4.27, recoveryTime: 14.00, peakPressure: 175, cycleTime: 33.0,
    // Actual Temps
    actTempNozzle: 319.9, actTempFront: 309.8, actTempMiddle: 300.2, actTempRear: 280.1,
    actTempHopper: 100.2, actTempCore: 149.2, actTempCavity: 148.9
  },
  {
    machine: 'V01',
    model: 'AA541',
    date: '2026-08-21',
    dateDisplay: '8/21/2026',
    setSpeed1: 120, setSpeed2: 80, setSpeed3: null,
    setShotSize: 45, setPos1: 35, setPos2: null, setPos3: null,
    vpChange: 6.0,
    setPeakPressure: 240,
    setCoolingTime: 32,
    // Setting Temps
    setTempNozzle: 325, setTempFront: 310, setTempMiddle: 300, setTempRear: 280,
    setTempHopper: 100, setTempCore: 160, setTempCavity: 160,
    // Actual Process
    fillTime: 1.20, cushion: 4.70, recoveryTime: 13.60, peakPressure: 207, cycleTime: 57.0,
    // Actual Temps
    actTempNozzle: 325.3, actTempFront: 309.4, actTempMiddle: 300.4, actTempRear: 280.4,
    actTempHopper: 100.4, actTempCore: 158.4, actTempCavity: 162.3
  }
];

// Security Passphrase
const AUTH_PASSPHRASE = 'monitor9710T1';
const AUTH_STORAGE_KEY = 'pim_security_auth_v1';

// App Initialization
document.addEventListener('DOMContentLoaded', async () => {
  checkAuthentication();
  initEventHandlers();
  await loadStoredData();
  populateMachineDropdown();
  applyFiltersAndRender();
});

/**
 * Find the latest Excel filename in the repository or local folder
 */
async function resolveLatestExcelFileName() {
  const candidates = [
    '20261007_injection monitor_R2.xlsx',
    '20261007_injection monitor_R3.xlsx',
    '20261007_injection monitor_R4.xlsx',
    '20261007_injection monitor_R1.xlsx',
    'injection_monitor.xlsx',
    'data.xlsx'
  ];

  // If running on GitHub Pages, try fetching the repo file tree to find the newest .xlsx automatically
  if (window.location.hostname.includes('github.io')) {
    try {
      const res = await fetch('https://api.github.com/repos/worachot33-boop/injection-monitor/contents/?_t=' + Date.now());
      if (res.ok) {
        const files = await res.json();
        const xlsxFiles = files
          .filter(f => f.name && f.name.toLowerCase().endsWith('.xlsx'))
          .map(f => f.name)
          .sort()
          .reverse(); // Newest revision first (e.g. R3 > R2 > R1)
        if (xlsxFiles.length > 0) {
          return xlsxFiles[0];
        }
      }
    } catch (e) {
      console.warn('GitHub API check failed, falling back to candidates list', e);
    }
  }

  // Probe candidate list
  for (const name of candidates) {
    try {
      const testRes = await fetch(encodeURI(name) + '?_t=' + Date.now(), { method: 'HEAD' });
      if (testRes.ok) return name;
    } catch (e) {}
  }

  return '20261007_injection monitor_R2.xlsx';
}

/**
 * Load Stored Data (Instant Cache + Always Background Sync)
 */
async function loadStoredData() {
  const cached = localStorage.getItem('precision_injection_data_r1');
  let hasCache = false;
  if (cached) {
    try {
      const parsed = JSON.parse(cached);
      if (Array.isArray(parsed) && parsed.length > 0) {
        state.rawDataset = parsed;
        hasCache = true;
        updateSyncStatus(`Loaded ${state.rawDataset.length} records (from cache)`);
      }
    } catch (e) {
      console.warn('Error reading local cache', e);
    }
  }

  if (!hasCache) {
    state.rawDataset = [...INITIAL_DEMO_DATA];
    updateSyncStatus('กำลังค้นหาข้อมูลล่าสุดจาก Excel...');
  }

  // Always sync latest file from GitHub/Server in background
  setTimeout(() => {
    syncLatestExcelData(false);
  }, 200);
}

/**
 * Sync Latest Excel Data (Bypasses browser cache with timestamp buster)
 */
async function syncLatestExcelData(isManual = false) {
  const syncSvg = document.getElementById('syncIconSvg');
  if (syncSvg) syncSvg.classList.add('spin-animation');
  if (isManual) updateSyncStatus('กำลังซิงค์ข้อมูลล่าสุดจาก GitHub...');

  try {
    const fileName = await resolveLatestExcelFileName();
    const cacheBuster = Date.now();
    const res = await fetch(encodeURI(fileName) + '?_t=' + cacheBuster, {
      cache: 'no-store'
    });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);

    const buffer = await res.arrayBuffer();
    const data = new Uint8Array(buffer);
    if (typeof XLSX === 'undefined') {
      console.warn('SheetJS library not yet loaded');
      return false;
    }

    const workbook = XLSX.read(data, { type: 'array' });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const jsonRows = XLSX.utils.sheet_to_json(worksheet, { defval: null });
    const parsedRecords = parseExcelRows(jsonRows);

    if (parsedRecords && parsedRecords.length > 0) {
      const isDifferent = state.rawDataset.length !== parsedRecords.length ||
                          JSON.stringify(state.rawDataset[0]) !== JSON.stringify(parsedRecords[0]);

      state.rawDataset = parsedRecords;
      saveToLocalStorage();
      populateMachineDropdown();
      applyFiltersAndRender();
      updateSyncStatus(`Data Source: ${fileName} (${parsedRecords.length} records)`);

      if (isManual) {
        showToast(`ซิงค์ข้อมูลสำเร็จ! โหลด ${parsedRecords.length} แถวจาก ${fileName}`, 'success');
      } else if (isDifferent) {
        showToast(`🔔 ตรวจพบไฟล์ใหม่ (${fileName}): อัปเดตข้อมูล ${parsedRecords.length} แถวเรียบร้อย`, 'success');
      }
      return true;
    }
  } catch (err) {
    console.warn('Could not sync latest Excel file:', err);
    if (isManual) {
      showToast(`ไม่สามารถซิงค์ไฟล์ได้: ${err.message}`, 'error');
    }
  } finally {
    if (syncSvg) {
      setTimeout(() => syncSvg.classList.remove('spin-animation'), 600);
    }
  }
  return false;
}

/**
 * Save to localStorage
 */
function saveToLocalStorage() {
  try {
    localStorage.setItem('precision_injection_data_r1', JSON.stringify(state.rawDataset));
  } catch (e) {
    console.error('Failed to save to localStorage', e);
  }
}

/**
 * Populate Machine Dropdown dynamically
 */
function populateMachineDropdown() {
  const selectMachine = document.getElementById('selectMachine');
  const machines = Array.from(new Set(state.rawDataset.map(d => d.machine))).filter(Boolean).sort();

  selectMachine.innerHTML = '<option value="ALL">All Machines (ทั้งหมด)</option>';
  machines.forEach(m => {
    const opt = document.createElement('option');
    opt.value = m;
    opt.textContent = `Machine: ${m}`;
    selectMachine.appendChild(opt);
  });

  selectMachine.value = state.selectedMachine;
  populateProductModelDropdown();
}

/**
 * Populate Product Model Dropdown (Cascades based on selected machine)
 */
function populateProductModelDropdown() {
  const selectModel = document.getElementById('selectProductModel');
  let filtered = state.rawDataset;
  if (state.selectedMachine !== 'ALL') {
    filtered = filtered.filter(d => d.machine === state.selectedMachine);
  }

  const models = Array.from(new Set(filtered.map(d => d.model))).filter(Boolean).sort();

  selectModel.innerHTML = '<option value="ALL">All Models (ทุกรุ่นโมเดล)</option>';
  models.forEach(m => {
    const opt = document.createElement('option');
    opt.value = m;
    opt.textContent = `Model: ${m}`;
    selectModel.appendChild(opt);
  });

  if (models.includes(state.selectedModel)) {
    selectModel.value = state.selectedModel;
  } else {
    state.selectedModel = 'ALL';
    selectModel.value = 'ALL';
  }
}

/**
 * Parse any date format reliably (Excel serial, ISO, M/D/YYYY, D/M/YYYY)
 */
function parseStandardDate(raw) {
  if (!raw) return null;
  if (raw instanceof Date && !isNaN(raw.getTime())) {
    const y = raw.getFullYear();
    const m = String(raw.getMonth() + 1).padStart(2, '0');
    const d = String(raw.getDate()).padStart(2, '0');
    return { dateStr: `${y}-${m}-${d}`, dateDisplay: `${raw.getMonth()+1}/${raw.getDate()}/${y}`, time: new Date(y, raw.getMonth(), raw.getDate()).getTime() };
  }
  if (typeof raw === 'number') {
    const utcDays = raw - 25569;
    const dateObj = new Date(Math.round(utcDays * 86400 * 1000));
    let y = dateObj.getUTCFullYear();
    const m = dateObj.getUTCMonth();
    const d = dateObj.getUTCDate();
    if (y > 2400) y -= 543;
    const cleanDate = new Date(y, m, d);
    return {
      dateStr: `${y}-${String(m + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`,
      dateDisplay: `${m + 1}/${d}/${y}`,
      time: cleanDate.getTime()
    };
  }
  const s = String(raw).trim();
  const isoMatch = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (isoMatch) {
    let y = parseInt(isoMatch[1], 10);
    if (y > 2400) y -= 543;
    const m = parseInt(isoMatch[2], 10) - 1;
    const d = parseInt(isoMatch[3], 10);
    const cleanDate = new Date(y, m, d);
    return {
      dateStr: `${y}-${String(m + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`,
      dateDisplay: `${m + 1}/${d}/${y}`,
      time: cleanDate.getTime()
    };
  }
  const slashMatch = s.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})/);
  if (slashMatch) {
    let p1 = parseInt(slashMatch[1], 10);
    let p2 = parseInt(slashMatch[2], 10);
    let y = parseInt(slashMatch[3], 10);
    if (y < 100) y += 2000;
    if (y > 2400) y -= 543;
    let m, d;
    if (p1 > 12) { d = p1; m = p2 - 1; }
    else { m = p1 - 1; d = p2; }
    const cleanDate = new Date(y, m, d);
    return {
      dateStr: `${y}-${String(m + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`,
      dateDisplay: `${m + 1}/${d}/${y}`,
      time: cleanDate.getTime()
    };
  }
  const parsed = new Date(s);
  if (!isNaN(parsed.getTime())) {
    let y = parsed.getFullYear();
    if (y > 2400) y -= 543;
    const m = parsed.getMonth();
    const d = parsed.getDate();
    return {
      dateStr: `${y}-${String(m + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`,
      dateDisplay: `${m + 1}/${d}/${y}`,
      time: new Date(y, m, d).getTime()
    };
  }
  return null;
}

/**
 * Filter Dataset based on Machine, Model, Calendar Day Range, and Search Query
 */
function applyFiltersAndRender() {
  let list = [...state.rawDataset];

  // 1. Filter by Machine
  if (state.selectedMachine !== 'ALL') {
    list = list.filter(d => d.machine === state.selectedMachine);
  }

  // 2. Filter by Product Model
  if (state.selectedModel !== 'ALL') {
    list = list.filter(d => d.model === state.selectedModel);
  }

  // Helper to get reliable epoch timestamp
  const getRecordTime = (d) => {
    if (d.dateTime) return d.dateTime;
    const p = parseStandardDate(d.date);
    return p ? p.time : 0;
  };

  // Sort chronologically ascending for trend chart analysis
  list.sort((a, b) => getRecordTime(a) - getRecordTime(b));

  // 3. Filter by True Calendar Days Range (30, 60, 120, 360 วันย้อนหลัง)
  if (state.selectedRange !== 'ALL' && list.length > 0) {
    const days = parseInt(state.selectedRange, 10);
    if (!isNaN(days) && days > 0) {
      const latestTime = getRecordTime(list[list.length - 1]);
      if (latestTime > 0) {
        // Go back 'days' full calendar days from the latest recorded date
        const cutoffTime = latestTime - ((days - 1) * 24 * 60 * 60 * 1000);
        list = list.filter(d => getRecordTime(d) >= cutoffTime);
      }
    }
  }

  // 4. Text Search (if any)
  if (state.searchQuery.trim() !== '') {
    const q = state.searchQuery.toLowerCase().trim();
    list = list.filter(d => 
      (d.machine && d.machine.toLowerCase().includes(q)) ||
      (d.model && d.model.toLowerCase().includes(q)) ||
      (d.dateDisplay && d.dateDisplay.toLowerCase().includes(q)) ||
      (d.date && d.date.toLowerCase().includes(q))
    );
  }

  state.filteredDataset = list;

  // Update Summary tags
  const rangeLabel = state.selectedRange === 'ALL' ? 'ทั้งหมด' : `${state.selectedRange} วัน`;
  document.getElementById('tagRecordCount').textContent = `Showing ${list.length} records (${rangeLabel})`;
  if (list.length > 0) {
    const firstDate = list[0].dateDisplay || list[0].date;
    const lastDate = list[list.length - 1].dateDisplay || list[list.length - 1].date;
    document.getElementById('tagDateSpan').textContent = `Range: ${firstDate} ➔ ${lastDate}`;
  } else {
    document.getElementById('tagDateSpan').textContent = 'No records in range';
  }

  renderKPICards();
  renderTrendChart();
  renderThermalZones();
  renderSettingProfiles();
  renderDataTable();
}

/**
 * Evaluate Cushion Tolerance
 * 1.1 Cushion 3 - 6 mm -> Green (Optimal)
 * 1.2 Cushion 2 - 3 mm -> Yellow (Warning)
 * 1.3 Cushion < 2 mm   -> Red (Critical Alert)
 * Cushion > 6 mm       -> Yellow (High Cushion Warning)
 */
function evaluateCushionStatus(cushion) {
  if (cushion === null || cushion === undefined) {
    return { status: 'normal', color: 'text-muted', badge: 'badge-info', label: 'No Data', level: 0 };
  }
  if (cushion < 2.0) {
    return { status: 'alert', color: 'text-rose', badge: 'badge-danger', label: 'Alert (<2.0mm)', reason: 'Cushion < 2mm', level: 2 };
  } else if (cushion >= 2.0 && cushion < 3.0) {
    return { status: 'warning', color: 'text-amber', badge: 'badge-warning', label: 'Warning (2.0-3.0mm)', reason: 'Cushion 2-3mm', level: 1 };
  } else if (cushion >= 3.0 && cushion <= 6.0) {
    return { status: 'normal', color: 'text-emerald', badge: 'badge-success', label: 'Normal (3.0-6.0mm)', reason: 'In Spec', level: 0 };
  } else {
    return { status: 'warning', color: 'text-amber', badge: 'badge-warning', label: 'High (>6.0mm)', reason: 'Cushion > 6mm', level: 1 };
  }
}

/**
 * Evaluate Peak Injection Pressure Tolerance
 * 2.1 Actual > Setting -> Red
 * 2.2 Deviation > 15% compared to average of 30 previous records -> Red
 * 2.3 Otherwise -> Green
 */
function evaluatePressureStatus(actP, setP, avgPrev30) {
  if (actP === null || actP === undefined) {
    return { status: 'normal', color: 'text-muted', badge: 'badge-info', label: 'No Data', level: 0 };
  }

  // 2.1: Act > Setting -> Red
  if (setP && actP > setP) {
    const diff = actP - setP;
    return {
      status: 'alert',
      color: 'text-rose',
      badge: 'badge-danger',
      label: `Exceeded Setting (+${diff.toFixed(1)} Mpa)`,
      reason: 'P > Setting',
      level: 2
    };
  }

  // 2.2: Deviation > 15% vs previous 30 data avg -> Red
  if (avgPrev30 && avgPrev30 > 0) {
    const devPct = (Math.abs(actP - avgPrev30) / avgPrev30) * 100;
    if (devPct > 15.0) {
      const diffPrev = actP - avgPrev30;
      return {
        status: 'alert',
        color: 'text-rose',
        badge: 'badge-danger',
        label: `Dev >15% vs 30-data avg (${diffPrev > 0 ? '+' : ''}${diffPrev.toFixed(1)} Mpa, ${devPct.toFixed(1)}%)`,
        reason: 'Δ30 > 15%',
        level: 2
      };
    }
  }

  // 2.3: Green
  return {
    status: 'normal',
    color: 'text-emerald',
    badge: 'badge-success',
    label: 'Normal (In Spec)',
    reason: 'Normal',
    level: 0
  };
}

/**
 * Render KPI Metric Summary Cards
 */
function renderKPICards() {
  const list = state.filteredDataset;
  if (list.length === 0) {
    document.getElementById('kpiActualPressure').textContent = '-';
    document.getElementById('kpiSettingPressure').textContent = 'Set: - Mpa';
    document.getElementById('kpiDeltaPressure').textContent = 'No Data';
    document.getElementById('kpiActualCushion').textContent = '-';
    document.getElementById('kpiActualFillTime').textContent = '-';
    document.getElementById('kpiActualCycleTime').textContent = '-';
    document.getElementById('kpiNozzleTemp').textContent = '-';
    return;
  }

  const latest = list[list.length - 1];

  // Calculate 30-data rolling average preceding the latest record
  const prev30 = list.slice(Math.max(0, list.length - 1 - 30), list.length - 1);
  const avgPrev30 = prev30.length > 0 
    ? (prev30.reduce((sum, d) => sum + (d.peakPressure || 0), 0) / prev30.length)
    : null;

  // 1. Peak Pressure
  const actP = latest.peakPressure || 0;
  const setP = latest.setPeakPressure || 0;
  const pressEval = evaluatePressureStatus(actP, setP, avgPrev30);

  document.getElementById('kpiActualPressure').textContent = actP.toFixed(1);
  document.getElementById('kpiSettingPressure').textContent = `Set: ${setP.toFixed(0)} Mpa`;

  const deltaEl = document.getElementById('kpiDeltaPressure');
  const badgePressure = document.getElementById('kpiBadgePressure');
  
  deltaEl.textContent = pressEval.label;
  deltaEl.className = `kpi-delta ${pressEval.color}`;
  badgePressure.className = `badge ${pressEval.badge}`;
  badgePressure.textContent = pressEval.level === 2 ? 'ALERT (RED)' : 'IN SPEC (GREEN)';

  // Avg Pressure in range
  const avgP = list.reduce((sum, d) => sum + (d.peakPressure || 0), 0) / list.length;
  document.getElementById('kpiAvgPressure').textContent = `${avgP.toFixed(1)} Mpa`;

  // 2. Cushion
  const actC = latest.cushion || 0;
  const cushionEval = evaluateCushionStatus(actC);

  document.getElementById('kpiActualCushion').textContent = actC.toFixed(2);
  const avgC = list.reduce((sum, d) => sum + (d.cushion || 0), 0) / list.length;
  document.getElementById('kpiAvgCushion').textContent = `${avgC.toFixed(2)} mm`;

  const deltaCushion = document.getElementById('kpiDeltaCushion');
  const badgeCushion = document.getElementById('kpiBadgeCushion');

  deltaCushion.textContent = cushionEval.label;
  deltaCushion.className = `kpi-delta ${cushionEval.color}`;
  badgeCushion.className = `badge ${cushionEval.badge}`;
  badgeCushion.textContent = cushionEval.level === 2 ? 'ALERT (<2mm)' : (cushionEval.level === 1 ? 'WARNING (2-3mm)' : 'OPTIMAL (3-6mm)');

  // 3. Fill Time
  const actFT = latest.fillTime || 0;
  document.getElementById('kpiActualFillTime').textContent = actFT.toFixed(2);
  const avgFT = list.reduce((sum, d) => sum + (d.fillTime || 0), 0) / list.length;
  document.getElementById('kpiAvgFillTime').textContent = `${avgFT.toFixed(2)} s`;

  // 4. Cycle Time & Cooling
  const actCT = latest.cycleTime || 0;
  const setCool = latest.setCoolingTime || 0;
  document.getElementById('kpiActualCycleTime').textContent = actCT.toFixed(1);
  document.getElementById('kpiSettingCooling').textContent = `Cooling: ${setCool.toFixed(0)} s`;
  document.getElementById('kpiActualRecovery').textContent = `${(latest.recoveryTime || 0).toFixed(1)} s`;

  // 5. Temperatures Summary
  const actNozzle = latest.actTempNozzle || latest.setTempNozzle || '-';
  const actCavity = latest.actTempCavity || latest.setTempCavity || '-';
  const actCore = latest.actTempCore || latest.setTempCore || '-';
  document.getElementById('kpiNozzleTemp').textContent = actNozzle;
  document.getElementById('kpiMoldTemp').textContent = `Mold: ${actCavity}°C`;
  document.getElementById('kpiCavityCore').textContent = `${actCavity} / ${actCore}°C`;
}

/**
 * Render Main Process Trend Chart (Peak Pressure, Cushion, Fill Time, Cycle Time)
 */
function renderTrendChart() {
  const ctx = document.getElementById('trendChart').getContext('2d');
  const list = state.filteredDataset;

  const labels = list.map(d => `${d.dateDisplay || d.date} (${d.machine})`);
  let datasets = [];
  let yTitle = '';

  if (state.activeMetric === 'pressure') {
    yTitle = 'Pressure (Mpa)';
    const actualPressures = list.map(d => d.peakPressure);
    const settingPressures = list.map(d => d.setPeakPressure);

    datasets = [
      {
        label: 'Actual Peak Pressure (Mpa)',
        data: actualPressures,
        borderColor: '#00f2fe',
        backgroundColor: 'rgba(0, 242, 254, 0.12)',
        fill: true,
        borderWidth: 2.5,
        pointRadius: 4,
        pointHoverRadius: 7,
        tension: 0.2
      },
      {
        label: 'Setting Setpoint (Mpa)',
        data: settingPressures,
        borderColor: '#f59e0b',
        borderDash: [6, 4],
        borderWidth: 2,
        pointRadius: 3,
        tension: 0.1
      }
    ];

    document.getElementById('chartSubtitle').textContent = 
      'เปรียบเทียบ Peak Pressure จริง (เส้นสีฟ้า) กับ Set Fill peak ที่ตั้งไว้ (เส้นประสีส้ม) [แดงถ้า > Setting หรือ เบี่ยงเบน > 15% เทียบ 30 ข้อมูลก่อนหน้า]';
    document.getElementById('insightSummary').textContent = 
      `จำนวนวันที่แสดง: ${list.length} จุด | พิกัดเป้าหมาย: ${list[list.length - 1]?.setPeakPressure || '-'} Mpa`;

  } else if (state.activeMetric === 'cushion') {
    yTitle = 'Cushion Stroke (mm)';
    const cushions = list.map(d => d.cushion);
    datasets = [
      {
        label: 'Actual Cushion (mm)',
        data: cushions,
        borderColor: '#10b981',
        backgroundColor: 'rgba(16, 185, 129, 0.12)',
        fill: true,
        borderWidth: 2.5,
        pointRadius: 4,
        tension: 0.2
      },
      {
        label: 'Upper Limit (6.0 mm - Green Limit)',
        data: Array(list.length).fill(6.0),
        borderColor: '#10b981',
        borderDash: [5, 4],
        borderWidth: 1.5,
        pointRadius: 0
      },
      {
        label: 'Lower Optimal (3.0 mm - Yellow Threshold)',
        data: Array(list.length).fill(3.0),
        borderColor: '#f59e0b',
        borderDash: [5, 4],
        borderWidth: 1.5,
        pointRadius: 0
      },
      {
        label: 'Critical Limit (2.0 mm - Red Threshold)',
        data: Array(list.length).fill(2.0),
        borderColor: '#f43f5e',
        borderDash: [4, 3],
        borderWidth: 2,
        pointRadius: 0
      }
    ];
    document.getElementById('chartSubtitle').textContent = 
      'ระยะคลายสกรูที่เหลือ (Cushion mm): เขียว = 3-6 มม., เหลือง = 2-3 มม., แดง = < 2 มม.';
  } else if (state.activeMetric === 'fillTime') {
    yTitle = 'Fill Time (seconds)';
    datasets = [
      {
        label: 'Actual Filling Time (s)',
        data: list.map(d => d.fillTime),
        borderColor: '#6366f1',
        backgroundColor: 'rgba(99, 102, 241, 0.15)',
        fill: true,
        borderWidth: 2.5,
        pointRadius: 4,
        tension: 0.2
      }
    ];
    document.getElementById('chartSubtitle').textContent = 
      'ความเร็วการไหลเข้าแม่พิมพ์ (Filling Time sec) แสดงความหนืดและการต้านทานการไหล';
  } else if (state.activeMetric === 'cycleTime') {
    yTitle = 'Cycle Time (seconds)';
    const cycleTimes = list.map(d => d.cycleTime);
    datasets = [
      {
        label: 'Actual Cycle Time (s)',
        data: cycleTimes,
        borderColor: '#00f2fe',
        backgroundColor: 'rgba(0, 242, 254, 0.12)',
        fill: true,
        borderWidth: 2.5,
        pointRadius: 4,
        pointHoverRadius: 7,
        tension: 0.2
      }
    ];
    document.getElementById('chartSubtitle').textContent = 
      'เวลาต่อรอบการทำงานรวม (Actual Cycle Time sec) ของแต่ละรอบการฉีด';
    document.getElementById('insightSummary').textContent = 
      `จำนวนวันที่แสดง: ${list.length} จุด | รอบล่าสุด: ${list[list.length - 1]?.cycleTime || '-'} s`;

  } else if (state.activeMetric === 'coolingRecovery') {
    yTitle = 'Time (seconds)';
    const coolingTimes = list.map(d => d.setCoolingTime);
    const recoveryTimes = list.map(d => d.recoveryTime);
    datasets = [
      {
        label: 'Set Cooling Time (s)',
        data: coolingTimes,
        borderColor: '#10b981',
        borderDash: [5, 4],
        borderWidth: 2,
        pointRadius: 3.5,
        tension: 0.1
      },
      {
        label: 'Screw Recovery Time (s)',
        data: recoveryTimes,
        borderColor: '#f59e0b',
        backgroundColor: 'rgba(245, 158, 11, 0.12)',
        fill: true,
        borderWidth: 2.5,
        pointRadius: 4,
        pointHoverRadius: 6,
        tension: 0.2
      }
    ];
    document.getElementById('chartSubtitle').textContent = 
      'เปรียบเทียบเวลาหล่อเย็นที่ตั้งไว้ (Set Cooling Time เส้นประสีเขียว) กับเวลาหมุนสกรูป้อนเม็ดจริง (Screw Recovery Time เส้นสีส้ม)';
    document.getElementById('insightSummary').textContent = 
      `เวลาหล่อเย็น: ${list[list.length - 1]?.setCoolingTime || '-'} s | เวลาหมุนสกรูล่าสุด: ${list[list.length - 1]?.recoveryTime || '-'} s`;
  }

  if (trendChart) {
    trendChart.destroy();
  }

  trendChart = new Chart(ctx, {
    type: 'line',
    data: { labels, datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: {
        mode: 'index',
        intersect: false
      },
      plugins: {
        legend: {
          display: true,
          position: 'top',
          labels: {
            color: '#94a3b8',
            font: { family: 'Inter', size: 11 },
            boxWidth: 14,
            padding: 12
          }
        },
        tooltip: {
          backgroundColor: '#0f172a',
          titleColor: '#00f2fe',
          bodyColor: '#e2e8f0',
          borderColor: '#334155',
          borderWidth: 1,
          padding: 10
        }
      },
      scales: {
        x: {
          grid: { color: 'rgba(255, 255, 255, 0.05)' },
          ticks: { color: '#94a3b8', maxTicksLimit: 15, font: { family: 'JetBrains Mono', size: 10 } }
        },
        y: {
          title: {
            display: true,
            text: yTitle,
            color: '#64748b',
            font: { size: 11 }
          },
          grid: { color: 'rgba(255, 255, 255, 0.05)' },
          ticks: { color: '#94a3b8', font: { family: 'JetBrains Mono', size: 10 } }
        }
      }
    }
  });
}

/**
 * Render the 7 Dedicated Temperature Zone Charts (Setting vs Actual)
 */
function renderThermalZones() {
  const list = state.filteredDataset;
  if (list.length === 0) return;

  const labels = list.map(d => d.dateDisplay || d.date);

  const zonesConfig = [
    { key: 'nozzle', name: 'Nozzle', setField: 'setTempNozzle', actField: 'actTempNozzle', color: '#f43f5e' },
    { key: 'front',  name: 'Front',  setField: 'setTempFront',  actField: 'actTempFront',  color: '#f59e0b' },
    { key: 'middle', name: 'Middle', setField: 'setTempMiddle', actField: 'actTempMiddle', color: '#fbbf24' },
    { key: 'rear',   name: 'Rear',   setField: 'setTempRear',   actField: 'actTempRear',   color: '#a855f7' },
    { key: 'hopper', name: 'Hopper', setField: 'setTempHopper', actField: 'actTempHopper', color: '#38bdf8' },
    { key: 'core',   name: 'Core',   setField: 'setTempCore',   actField: 'actTempCore',   color: '#00f2fe' },
    { key: 'cavity', name: 'Cavity', setField: 'setTempCavity', actField: 'actTempCavity', color: '#10b981' }
  ];

  const latest = list[list.length - 1];

  zonesConfig.forEach(zone => {
    // 1. Update Header Readouts
    const actVal = latest[zone.actField] !== null && latest[zone.actField] !== undefined ? latest[zone.actField] : '-';
    const setVal = latest[zone.setField] !== null && latest[zone.setField] !== undefined ? latest[zone.setField] : '-';

    const capKey = zone.key.charAt(0).toUpperCase() + zone.key.slice(1);
    const actEl = document.getElementById(`readoutAct${capKey}`);
    const setEl = document.getElementById(`readoutSet${capKey}`);
    const deltaEl = document.getElementById(`delta${capKey}`);

    if (actEl) actEl.textContent = `${actVal !== '-' ? Number(actVal).toFixed(1) : '-'} °C`;
    if (setEl) setEl.textContent = `Set: ${setVal !== '-' ? Number(setVal).toFixed(0) : '-'}°C`;

    if (deltaEl && actVal !== '-' && setVal !== '-') {
      const delta = actVal - setVal;
      deltaEl.textContent = `Δ ${delta > 0 ? '+' : ''}${delta.toFixed(1)}°`;
      if (Math.abs(delta) <= 1.5) {
        deltaEl.className = 't-delta-tag t-delta-good';
      } else if (Math.abs(delta) <= 4.0) {
        deltaEl.className = 't-delta-tag t-delta-warn';
      } else {
        deltaEl.className = 't-delta-tag t-delta-bad';
      }
    }

    // 2. Build Mini Chart
    const canvasId = `chartTemp${capKey}`;
    const canvas = document.getElementById(canvasId);
    if (!canvas) return;

    const ctx = canvas.getContext('2d');
    const actData = list.map(d => d[zone.actField]);
    const setData = list.map(d => d[zone.setField]);

    if (thermalCharts[zone.key]) {
      thermalCharts[zone.key].destroy();
    }

    thermalCharts[zone.key] = new Chart(ctx, {
      type: 'line',
      data: {
        labels: labels,
        datasets: [
          {
            label: 'Actual (°C)',
            data: actData,
            borderColor: zone.color,
            backgroundColor: `${zone.color}22`,
            fill: true,
            borderWidth: 2,
            pointRadius: 3,
            tension: 0.2
          },
          {
            label: 'Setting (°C)',
            data: setData,
            borderColor: '#f59e0b',
            borderDash: [5, 4],
            borderWidth: 1.5,
            pointRadius: 2,
            tension: 0.1
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            backgroundColor: '#0f172a',
            titleColor: zone.color,
            bodyColor: '#e2e8f0',
            borderColor: '#334155',
            borderWidth: 1,
            callbacks: {
              afterBody: (context) => {
                if (context.length >= 2) {
                  const act = context[0].parsed.y;
                  const set = context[1].parsed.y;
                  const diff = act - set;
                  return `Δ Drift: ${diff > 0 ? '+' : ''}${diff.toFixed(1)} °C`;
                }
              }
            }
          }
        },
        scales: {
          x: {
            grid: { color: 'rgba(255, 255, 255, 0.04)' },
            ticks: { color: '#64748b', maxTicksLimit: 6, font: { size: 9, family: 'JetBrains Mono' } }
          },
          y: {
            grid: { color: 'rgba(255, 255, 255, 0.04)' },
            ticks: { color: '#94a3b8', font: { size: 9, family: 'JetBrains Mono' } }
          }
        }
      }
    });
  });
}

/**
 * Render Setting Profile Box
 */
function renderSettingProfiles() {
  const list = state.filteredDataset;
  if (list.length === 0) return;
  const latest = list[list.length - 1];

  document.getElementById('badgeCurrentModel').textContent = `Model: ${latest.model} (${latest.machine})`;
  document.getElementById('valShotSize').textContent = `${latest.setShotSize || '-'} mm`;
  document.getElementById('valVpChange').textContent = `${latest.vpChange || '-'} mm`;
  document.getElementById('valSpeed1').textContent = `${latest.setSpeed1 || '-'} mm/s`;
  document.getElementById('valSpeed2').textContent = `${latest.setSpeed2 || '-'} mm/s`;
  document.getElementById('valPos1').textContent = `${latest.setPos1 || '-'} mm`;
  document.getElementById('valCooling').textContent = `${latest.setCoolingTime || '-'} s`;
}

/**
 * Render Daily Data Table with User-specified Tolerance Alerts:
 * 1. Cushion: 3-6mm (Green), 2-3mm (Yellow), <2mm (Red)
 * 2. Peak Pressure: >Setting (Red), >15% vs 30-data avg (Red), else Green
 */
function renderDataTable() {
  const tbody = document.getElementById('dataTableBody');
  tbody.innerHTML = '';
  const list = [...state.filteredDataset].reverse(); // newest on top

  if (list.length === 0) {
    tbody.innerHTML = `<tr><td colspan="12" style="text-align:center; padding: 20px; color: #64748b;">ไม่พบข้อมูลตามเงื่อนไขที่เลือก</td></tr>`;
    return;
  }

  list.forEach(row => {
    const tr = document.createElement('tr');

    const actP = row.peakPressure || 0;
    const setP = row.setPeakPressure || 0;
    const deltaP = actP - setP;
    const deltaPct = setP > 0 ? (deltaP / setP) * 100 : 0;

    // Calculate rolling 30 previous data points before this row in chronological order
    const chronoIndex = state.filteredDataset.indexOf(row);
    const prev30 = state.filteredDataset.slice(Math.max(0, chronoIndex - 30), chronoIndex);
    const avgPrev30 = prev30.length > 0 
      ? (prev30.reduce((sum, d) => sum + (d.peakPressure || 0), 0) / prev30.length)
      : null;

    // Evaluate Tolerance Conditions
    const pressEval = evaluatePressureStatus(actP, setP, avgPrev30);
    const cushionEval = evaluateCushionStatus(row.cushion);

    // Determine Overall Row Status
    let badgeClass = 'badge-success';
    let badgeText = 'Normal';
    let rowClass = '';

    if (pressEval.level === 2 || cushionEval.level === 2) {
      rowClass = 'row-danger';
      badgeClass = 'badge-danger';
      const reasons = [];
      if (pressEval.level === 2) reasons.push(pressEval.reason);
      if (cushionEval.level === 2) reasons.push(cushionEval.reason);
      badgeText = `Alert (${reasons.join(', ')})`;
    } else if (cushionEval.level === 1) {
      rowClass = 'row-warn';
      badgeClass = 'badge-warning';
      badgeText = `Warning (${cushionEval.reason})`;
    } else {
      badgeClass = 'badge-success';
      badgeText = 'Normal';
    }

    tr.className = rowClass;

    const deltaSign = deltaP > 0 ? '+' : '';
    let deltaInfoHtml = `<span class="${pressEval.color}"><strong>${deltaSign}${deltaP.toFixed(1)}</strong> Mpa (${deltaPct.toFixed(1)}%)</span>`;
    if (avgPrev30) {
      const diffPrev30 = actP - avgPrev30;
      const dev30Pct = (diffPrev30 / avgPrev30) * 100;
      deltaInfoHtml += `<br><small style="color: ${pressEval.level === 2 && Math.abs(dev30Pct) > 15 ? '#f43f5e' : '#94a3b8'};">Δ30: ${diffPrev30 > 0 ? '+' : ''}${diffPrev30.toFixed(1)} (${dev30Pct.toFixed(1)}%)</small>`;
    }

    const nozzleCell = `${row.actTempNozzle ?? '-'} / ${row.setTempNozzle ?? '-'}`;
    const cavityCell = `${row.actTempCavity ?? '-'} / ${row.setTempCavity ?? '-'}`;

    tr.innerHTML = `
      <td><strong>${row.dateDisplay || row.date}</strong></td>
      <td><span class="badge badge-info">${row.machine}</span></td>
      <td>${row.model}</td>
      <td><span class="${pressEval.color}"><strong>${actP.toFixed(1)}</strong></span> / ${setP.toFixed(0)} Mpa</td>
      <td>${deltaInfoHtml}</td>
      <td><span class="${cushionEval.color}"><strong>${(row.cushion || 0).toFixed(2)}</strong></span></td>
      <td>${(row.fillTime || 0).toFixed(2)}</td>
      <td>${(row.cycleTime || 0).toFixed(1)}</td>
      <td>${(row.recoveryTime || 0).toFixed(1)}</td>
      <td>${nozzleCell}</td>
      <td>${cavityCell}</td>
      <td><span class="badge ${badgeClass}">${badgeText}</span></td>
    `;
    tbody.appendChild(tr);
  });
}

/**
 * Handle Excel File Parsing via SheetJS (Supports 20261007_injection monitor_R1.xlsx)
 */
function handleExcelUpload(file) {
  if (!file) return;

  updateSyncStatus(`Reading file: ${file.name}...`);
  const reader = new FileReader();

  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });

      // First sheet
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const jsonRows = XLSX.utils.sheet_to_json(worksheet, { defval: null });

      if (!jsonRows || jsonRows.length === 0) {
        showToast('ไฟล์ Excel ว่างเปล่าหรือไม่พบข้อมูล', 'error');
        updateSyncStatus('Upload Error: Empty Sheet');
        return;
      }

      const parsedRecords = parseExcelRows(jsonRows);

      if (parsedRecords.length === 0) {
        showToast('ไม่สามารถจับคู่หัวคอลัมน์ของไฟล์ได้ กรุณาตรวจสอบรูปแบบ', 'error');
        updateSyncStatus('Upload Error: Header mismatch');
        return;
      }

      state.rawDataset = parsedRecords;
      saveToLocalStorage();

      populateMachineDropdown();
      applyFiltersAndRender();

      updateSyncStatus(`Updated from ${file.name} (${parsedRecords.length} records)`);
      showToast(`นำเข้าสำเร็จ! โหลดข้อมูล ${parsedRecords.length} แถวเรียบร้อย`, 'success');

    } catch (err) {
      console.error('Error parsing Excel', err);
      showToast('เกิดข้อผิดพลาดในการอ่านไฟล์ Excel: ' + err.message, 'error');
      updateSyncStatus('Error parsing Excel');
    }
  };

  reader.readAsArrayBuffer(file);
}

/**
 * Parse Rows from SheetJS JSON
 * Maps column headers flexibly to support 20261007_injection monitor_R1.xlsx (32 columns)
 */
function parseExcelRows(rows) {
  const parsed = [];

  rows.forEach((r, idx) => {
    const getVal = (...keys) => {
      for (const k of keys) {
        if (r[k] !== undefined && r[k] !== null && r[k] !== '-') return r[k];
        // clean matching
        const cleanK = k.replace(/[\s\(\)_]/g, '').toLowerCase();
        const found = Object.keys(r).find(actualKey => 
          actualKey.replace(/[\s\(\)_]/g, '').toLowerCase() === cleanK
        );
        if (found && r[found] !== undefined && r[found] !== null && r[found] !== '-') {
          return r[found];
        }
      }
      return null;
    };

    const machine = getVal('Machine number', 'Machine', 'Machine No');
    const model = getVal('Product Model', 'Model', 'Product');
    let dateRaw = getVal('Injection date', 'Date');

    if (!machine && !model) return;

    // Parse Excel date
    let dateStr = '';
    let dateDisplay = '';
    if (typeof dateRaw === 'number') {
      const dateObj = new Date(Math.round((dateRaw - 25569) * 86400 * 1000));
      dateStr = dateObj.toISOString().split('T')[0];
      dateDisplay = `${dateObj.getMonth() + 1}/${dateObj.getDate()}/${dateObj.getFullYear()}`;
    } else if (dateRaw) {
      dateDisplay = String(dateRaw).trim();
      const d = new Date(dateDisplay);
      if (!isNaN(d.getTime())) {
        dateStr = d.toISOString().split('T')[0];
      } else {
        dateStr = dateDisplay;
      }
    } else {
      dateDisplay = `Entry #${idx + 1}`;
      dateStr = `2026-08-${String(idx + 1).padStart(2, '0')}`;
    }

    const num = (v) => {
      if (v === null || v === undefined || v === '') return null;
      const n = parseFloat(v);
      return isNaN(n) ? null : n;
    };

    parsed.push({
      machine: String(machine || 'M01').trim(),
      model: String(model || 'Standard').trim(),
      date: dateStr,
      dateDisplay: dateDisplay,

      // Settings
      setSpeed1: num(getVal('Set speed step 1(mm/sec)', 'Speed 1')),
      setSpeed2: num(getVal('Set speed step 2(mm/sec)', 'Speed 2')),
      setSpeed3: num(getVal('Set speed step 3(mm/sec)', 'Speed 3')),
      setShotSize: num(getVal('Set short size(mm)', 'Set shot size(mm)', 'Shot size')),
      setPos1: num(getVal('Set position step1(mm)', 'Pos 1')),
      setPos2: num(getVal('Set speed step 2(mm)', 'Set position step 2(mm)')),
      setPos3: num(getVal('Set speed step 3(mm)', 'Set position step 3(mm)')),
      vpChange: num(getVal('VP change(mm)', 'V/P change')),
      setPeakPressure: num(getVal('Set Fill peak(Mpa)', 'Set Peak Pressure')),
      setCoolingTime: num(getVal('Set cooling time(sec)', 'Set cooling time')),

      // Temperature Settings (From R1: Set nozzle (C), Set front(C), etc.)
      setTempNozzle: num(getVal('Set nozzle (C)', 'Nozzle (C)', 'Set Nozzle')),
      setTempFront: num(getVal('Set front(C)', 'Front(C)', 'Set Front')),
      setTempMiddle: num(getVal('Set middle(C)', 'Middle(C)', 'Set Middle')),
      setTempRear: num(getVal('Set real (C)', 'Set rear(C)', 'Real (C)', 'Rear(C)', 'Set Real')),
      setTempHopper: num(getVal('Set hopper (C)', 'Hopper', 'Set Hopper')),
      setTempCore: num(getVal('Set core temp (C)', 'Core temp', 'Set Core')),
      setTempCavity: num(getVal('Set cavity temp (C)', 'Cavity temp', 'Set Cavity')),

      // Process Actuals
      fillTime: num(getVal('Fill time', 'Filling time')),
      cushion: num(getVal('cushion', 'Cushion')),
      recoveryTime: num(getVal('Recovery time', 'Plasticizing time')),
      peakPressure: num(getVal('Peak pressure', 'Actual Peak pressure')),
      cycleTime: num(getVal('cycle time', 'Cycle time')),

      // Temperature Actuals (From R1: Actual nozzle (C)2, Actual front(C)3, etc.)
      actTempNozzle: num(getVal('Actual nozzle (C)2', 'Actual nozzle (C)', 'Actual Nozzle', 'Act Nozzle')),
      actTempFront: num(getVal('Actual front(C)3', 'Actual front(C)', 'Actual Front', 'Act Front')),
      actTempMiddle: num(getVal('Actual middle(C)4', 'Actual middle(C)', 'Actual Middle', 'Act Middle')),
      actTempRear: num(getVal('Actual real (C)5', 'Actual rear (C)', 'Actual Real', 'Act Real')),
      actTempHopper: num(getVal('Actual hopper (C)6', 'Actual hopper (C)', 'Actual Hopper', 'Act Hopper')),
      actTempCore: num(getVal('Actual core temp (C)7', 'Actual core temp (C)', 'Actual Core', 'Act Core')),
      actTempCavity: num(getVal('Actual cavity temp (C)8', 'Actual cavity temp (C)', 'Actual Cavity', 'Act Cavity'))
    });
  });

  return parsed;
}

/**
 * Event Handlers
 */
function initEventHandlers() {
  // Machine Selector
  document.getElementById('selectMachine').addEventListener('change', (e) => {
    state.selectedMachine = e.target.value;
    populateProductModelDropdown();
    applyFiltersAndRender();
  });

  // Product Model Selector
  document.getElementById('selectProductModel').addEventListener('change', (e) => {
    state.selectedModel = e.target.value;
    applyFiltersAndRender();
  });

  // Range Buttons (30, 60, 120, 360, ALL)
  const rangeBtns = document.querySelectorAll('.range-buttons .range-btn');
  rangeBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      rangeBtns.forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      state.selectedRange = btn.dataset.days;
      applyFiltersAndRender();
    });
  });

  // Metric Switch Tabs
  const metricTabs = document.querySelectorAll('.chart-tab-controls .metric-tab');
  metricTabs.forEach(tab => {
    tab.addEventListener('click', () => {
      metricTabs.forEach(t => t.classList.remove('active'));
      tab.classList.add('active');
      state.activeMetric = tab.dataset.metric;
      renderTrendChart();
    });
  });

  // Table Search Input
  document.getElementById('tableSearchInput').addEventListener('input', (e) => {
    state.searchQuery = e.target.value;
    applyFiltersAndRender();
  });

  // Import Excel Button & File Input
  const fileInput = document.getElementById('excelFileInput');
  document.getElementById('btnImportExcel').addEventListener('click', () => {
    fileInput.click();
  });
  fileInput.addEventListener('change', (e) => {
    if (e.target.files && e.target.files[0]) {
      handleExcelUpload(e.target.files[0]);
      e.target.value = '';
    }
  });

  // Drag & Drop onto Window
  const dropOverlay = document.getElementById('dropOverlay');
  window.addEventListener('dragenter', (e) => {
    e.preventDefault();
    dropOverlay.classList.add('drag-active');
  });
  window.addEventListener('dragover', (e) => {
    e.preventDefault();
  });
  dropOverlay.addEventListener('dragleave', (e) => {
    if (e.relatedTarget === null) {
      dropOverlay.classList.remove('drag-active');
    }
  });
  dropOverlay.addEventListener('drop', (e) => {
    e.preventDefault();
    dropOverlay.classList.remove('drag-active');
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      handleExcelUpload(e.dataTransfer.files[0]);
    }
  });

  // Download Excel Template Button
  document.getElementById('btnDownloadTemplate').addEventListener('click', () => {
    downloadStandardTemplate();
  });

  // Export Filtered CSV Button
  document.getElementById('btnExportFiltered').addEventListener('click', () => {
    exportFilteredDataCSV();
  });

  // Security Authentication Form & Buttons
  const authForm = document.getElementById('authForm');
  if (authForm) {
    authForm.addEventListener('submit', (e) => {
      e.preventDefault();
      handleUnlockAttempt();
    });
  }

  const btnTogglePw = document.getElementById('btnTogglePw');
  if (btnTogglePw) {
    btnTogglePw.addEventListener('click', () => {
      const pwInput = document.getElementById('authPasswordInput');
      if (pwInput) {
        pwInput.type = pwInput.type === 'password' ? 'text' : 'password';
      }
    });
  }

  const btnLockApp = document.getElementById('btnLockApp');
  if (btnLockApp) {
    btnLockApp.addEventListener('click', () => {
      lockApplication();
    });
  }

  // Fullscreen Toggle
  const btnToggleFullscreen = document.getElementById('btnToggleFullscreen');
  if (btnToggleFullscreen) {
    btnToggleFullscreen.addEventListener('click', () => {
      if (!document.fullscreenElement) {
        if (document.documentElement.requestFullscreen) {
          document.documentElement.requestFullscreen().catch(err => {
            console.warn('Fullscreen error:', err);
          });
        }
      } else {
        if (document.exitFullscreen) {
          document.exitFullscreen();
        }
      }
    });
  }

  // Force Sync Latest Excel Data
  const btnSyncData = document.getElementById('btnSyncData');
  if (btnSyncData) {
    btnSyncData.addEventListener('click', () => {
      syncLatestExcelData(true);
    });
  }
}

/**
 * Security & Access Control
 */
function checkAuthentication() {
  const isAuth = localStorage.getItem(AUTH_STORAGE_KEY) === 'granted' || 
                 sessionStorage.getItem(AUTH_STORAGE_KEY) === 'granted';
  const overlay = document.getElementById('authOverlay');
  if (!overlay) return;

  if (isAuth) {
    overlay.classList.add('unlocked');
  } else {
    overlay.classList.remove('unlocked');
    const input = document.getElementById('authPasswordInput');
    if (input) setTimeout(() => input.focus(), 150);
  }
}

function handleUnlockAttempt() {
  const input = document.getElementById('authPasswordInput');
  const errorMsg = document.getElementById('authErrorMsg');
  const card = document.querySelector('.auth-card');
  const remember = document.getElementById('authRememberCheckbox')?.checked;
  const overlay = document.getElementById('authOverlay');

  if (!input) return;
  const entered = input.value.trim();

  if (entered === AUTH_PASSPHRASE) {
    if (remember) {
      localStorage.setItem(AUTH_STORAGE_KEY, 'granted');
    } else {
      sessionStorage.setItem(AUTH_STORAGE_KEY, 'granted');
    }
    if (errorMsg) errorMsg.textContent = '';
    if (overlay) overlay.classList.add('unlocked');
    showToast('ยินดีต้อนรับ! ยืนยันรหัสผ่านสำเร็จ', 'success');
  } else {
    if (errorMsg) errorMsg.textContent = '❌ รหัสผ่านไม่ถูกต้อง กรุณาลองใหม่อีกครั้ง';
    if (card) {
      card.classList.add('shake');
      setTimeout(() => card.classList.remove('shake'), 500);
    }
    input.value = '';
    input.focus();
  }
}

function lockApplication() {
  localStorage.removeItem(AUTH_STORAGE_KEY);
  sessionStorage.removeItem(AUTH_STORAGE_KEY);
  const overlay = document.getElementById('authOverlay');
  if (overlay) {
    overlay.classList.remove('unlocked');
    const input = document.getElementById('authPasswordInput');
    if (input) {
      input.value = '';
      input.focus();
    }
  }
  showToast('ล็อคหน้าจอเรียบร้อยแล้ว', 'info');
}

/**
 * Generate & Download Standard Excel Template matching 20261007_injection monitor_R1.xlsx
 */
function downloadStandardTemplate() {
  const templateHeaders = [
    'Machine number', 'Product Model', 'Injection date',
    'Set speed step 1(mm/sec)', 'Set speed step 2(mm/sec)', 'Set speed step 3(mm/sec)',
    'Set short size(mm)', 'Set position step1(mm)', 'Set speed step 2(mm)', 'Set speed step 3(mm)',
    'VP change(mm)', 'Set Fill peak(Mpa)', 'Set cooling time(sec)',
    'Set nozzle (C)', 'Set front(C)', 'Set middle(C)', 'Set real (C)', 'Set hopper (C)', 'Set core temp (C)', 'Set cavity temp (C)',
    'Fill time', 'cushion', 'Recovery time', 'Peak pressure', 'cycle time',
    'Actual nozzle (C)2', 'Actual front(C)3', 'Actual middle(C)4', 'Actual real (C)5', 'Actual hopper (C)6', 'Actual core temp (C)7', 'Actual cavity temp (C)8'
  ];

  const sampleRow = [
    'B01', 'AA551', '8/25/2026',
    100, 60, '',
    32, 23, '', '',
    6, 200, 25,
    320, 310, 300, 280, 100, 150, 150,
    0.86, 4.25, 12.6, 186, 33,
    320.2, 310.1, 299.8, 280.2, 100.1, 151.0, 150.5
  ];

  const wsData = [templateHeaders, sampleRow];
  const ws = XLSX.utils.aoa_to_sheet(wsData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'InjectionMonitor');

  XLSX.writeFile(wb, 'Precision_Injection_Template_R1.xlsx');
  showToast('ดาวน์โหลดไฟล์เทมเพลต Precision_Injection_Template_R1.xlsx สำเร็จ', 'success');
}

/**
 * Export Current Filtered Dataset as CSV
 */
function exportFilteredDataCSV() {
  const list = state.filteredDataset;
  if (list.length === 0) {
    showToast('ไม่มีข้อมูลสำหรับส่งออก', 'warning');
    return;
  }

  let csv = 'Machine,ProductModel,Date,SetPeakPressure_Mpa,ActualPeakPressure_Mpa,Delta_Mpa,Cushion_mm,FillTime_s,CycleTime_s,SetNozzle_C,ActNozzle_C,SetCavity_C,ActCavity_C\n';
  list.forEach(r => {
    const actP = r.peakPressure || 0;
    const setP = r.setPeakPressure || 0;
    const deltaP = actP - setP;
    csv += `${r.machine},${r.model},${r.dateDisplay || r.date},${setP},${actP},${deltaP.toFixed(1)},${r.cushion || ''},${r.fillTime || ''},${r.cycleTime || ''},${r.setTempNozzle || ''},${r.actTempNozzle || ''},${r.setTempCavity || ''},${r.actTempCavity || ''}\n`;
  });

  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `Precision_Injection_${state.selectedMachine}_${state.selectedModel}_${Date.now()}.csv`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);

  showToast('ส่งออกข้อมูล CSV เรียบร้อยแล้ว', 'success');
}

/**
 * Status Sync Box Text
 */
function updateSyncStatus(text) {
  const el = document.getElementById('syncStatusText');
  if (el) el.textContent = text;
}

/**
 * Toast Notifications
 */
function showToast(msg, type = 'info') {
  const container = document.getElementById('toastContainer');
  if (!container) return;

  const toast = document.createElement('div');
  toast.className = `toast ${type === 'error' ? 'toast-error' : type === 'success' ? 'toast-success' : type === 'warning' ? 'toast-warning' : ''}`;
  toast.innerHTML = `<span>${msg}</span>`;

  container.appendChild(toast);
  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateY(10px)';
    setTimeout(() => toast.remove(), 300);
  }, 4000);
}
