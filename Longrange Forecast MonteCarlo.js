/** ======================================================================
 * DonorsChoose Monte Carlo for Google Sheets — Refactored Version
 * ----------------------------------------------------------------------
 * WHAT THIS DOES (executive summary):
 * - Reads your planning inputs from several sheets (Params, Forecasts, etc.)
 * - Builds scenario-adjusted truncated normal distributions per Channel×Year×Quarter
 * - Uses 5 binary factors (Enabled/Disabled) with fiscal-year-specific multipliers
 * - Simulates many trials with seeded randomness and simplified global correlation
 * - Applies YoY bounds AT THE ANNUAL LEVEL (with scenario-based adjustments)
 * - Aggregates streaming means and percentiles + plan hit-rates
 * - Writes summary tables to Results_Quarterly / Results_Annually / Results_PI
 *
 * WHY IT'S SAFE:
 * - Only writes to dedicated Results_* tabs (clears data rows, keeps headers)
 * - Uses deterministic seed for reproducibility
 * - Streaming quantiles => fast & memory-safe (no storing trial arrays)
 *
 * YOU CONFIGURE:
 * - Params tab (N_Sims, Seed, 5 scenario factors, Global_Channel_Correlation, SD_inflation_factor, PI levels)
 * - Forecasts tab (per Channel×Year×Quarter mean & predictive SD)
 * - Scenario_Mapping tab (FY-specific multipliers for 5 factors: Weak/Strong pairs)
 * - Bounds_Scenario_Adjustments tab (Floor/Cap shifts per channel×factor when Enabled)
 * - Annual_YoY_Bounds tab (Base YoY_Min, YoY_Max per channel/year)
 * - Params_Plans tab (annual plan per channel + Total_Plan)
 *
 * KEY MODEL RULES:
 * - Truncated normal distribution (can model YoY declines, respects Floor_Zero param)
 * - Global correlation creates co-movement across channels (default 0.3)
 * - SD inflates proportionally: each enabled factor adds ~10% uncertainty (default 50% max = 0.50)
 * - YoY bounds: FY26 exempt; FY27+ bounded versus prior simulated year
 *   Bounds shift based on active scenario factors (Enabled selections)
 *   Special case: if prior=0, we ignore the upper bound to allow recovery
 *
 * REFACTOR CHANGES FROM PREVIOUS VERSION:
 * - Removed: Complex channel-specific correlation weights (w_Economy, w_Government)
 * - Removed: Correlation_Config sheet entirely
 * - Removed: Three-level scenarios (Headwind/Neutral/Tailwind)
 * - Removed: Lognormal distribution (switched to truncated normal)
 * - Added: 5 binary scenario factors (Enabled/Disabled)
 * - Added: Fiscal-year dimension to Scenario_Mapping
 * - Added: Bounds_Scenario_Adjustments sheet for dynamic bounds
 * - Added: Proportional SD inflation (scales with number of factors enabled)
 * - Simplified: Single global correlation parameter
 * - Simplified: SD inflation based on Forecast_SD, not separate SD_base column
 * - Hardcoded: WithinYear_Rho=0.5
 * - User-configurable: SD_inflation_factor from Params (default 0.50 = 50%)
 *
 * DEPLOYMENT:
 * - Paste this file into Apps Script, save, refresh the sheet
 * - Use the "Monte Carlo" custom menu to run/validate/clear
 * ====================================================================== */

// HARDCODED TECHNICAL CONSTANTS (no longer user-configurable)
const WITHIN_YEAR_RHO = 0.5;         // Correlation between quarters within channel-year

/* =============================== MENU =============================== */
/** Installs a custom menu for users to run/validate/clear without opening Apps Script. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Monte Carlo')
    .addItem('Run Simulation', 'runMonteCarlo')
    .addItem('Validate Inputs', 'validateInputs')
    .addItem('Clear Results', 'clearResults')
    .addToUi();
}

/* ============================= ENTRY POINTS ========================= */
/** The top-level function users call to perform a full simulation run. */
function runMonteCarlo() {
  const ss = SpreadsheetApp.getActive();
  const log = makeLogger(ss);                // Append-only logger to a "Log" sheet
  const p = readParams(ss, log);             // Controls (named ranges or Params table)
  const data = readAllInputs(ss, p, log);    // Reads sheets; precomputes per-row (μ,σ)
  simulateAndWrite(ss, p, data, log);        // Runs trials, applies rules, writes results
  log.flush();                                // No-op now; placeholder for buffered logging
}

/** Runs validation only (no simulation, no writes to Results_*). */
function validateInputs() {
  const ss = SpreadsheetApp.getActive();
  const log = makeLogger(ss);
  readParams(ss, log);         // Triggers warnings if named ranges missing
  readAllInputs(ss, null, log);// Validates tables & structure without simulating
  log.flush();
}

/** Clears data rows (keeps headers) in results tabs so outputs are fresh/clean. */
function clearResults() {
  const ss = SpreadsheetApp.getActive();
  const sheets = ['Results_Quarterly', 'Results_Annually', 'Results_PI'];
  for (const name of sheets) {
    const sh = ss.getSheetByName(name);
    if (!sh) continue;
    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow > 1 && lastCol > 0) {
      sh.getRange(2, 1, lastRow - 1, lastCol).clearContent(); // preserve headings
    }
  }
}

/* ============================== PARAMS ============================== */
/**
 * Reads run-time parameters.
 * 1) Try named ranges (N_Sims, Seed, etc.) — BEST for stability against sheet moves.
 * 2) If a named range is missing, fall back to the Params sheet table (Param/Value).
 * Logs a WARN when a value isn't found and a default is used.
 */
function readParams(ss, log) {
  const fromNamed = n => {
    const r = ss.getRangeByName(n);
    if (!r) return null;
    const v = r.getValue();
    return (v === '' || v === null || v === undefined) ? null : v;
  };

  // Table fallback: reads sheet "Params" with columns Param / Value.
  const tbl = readParamsTable(ss);

  // Helper to pick named-range value; else table; else default with warning.
  const pick = (name, def) => {
    const nv = fromNamed(name);
    if (nv !== null) return nv;
    if (Object.prototype.hasOwnProperty.call(tbl, name) && tbl[name] !== '' && tbl[name] !== null && tbl[name] !== undefined) {
      return tbl[name];
    }
    log.warn(name, `Named range missing; using default: ${def}`);
    return def;
  };

  const asBool = v => {
    if (typeof v === 'boolean') return v;
    const s = String(v).trim().toLowerCase();
    return s === 'true' || s === '1' || s === 'yes';
  };

  // Parse each param with sensible defaults
  const N_Sims = Math.max(1, Math.floor(Number(pick('N_Sims', 10000))));
  const Seed = Math.floor(Number(pick('Seed', 42)));
  const RoundK = asBool(pick('Round_To_Thousands', true));
  const PI_L = Number(pick('PI_Lower_Level', 0.05));
  const PI_U = Number(pick('PI_Upper_Level', 0.95));
  
  // DIAGNOSTIC: Log what was actually read
  log.info('Params', `PI_Lower_Level read as: ${PI_L}`);
  log.info('Params', `PI_Upper_Level read as: ${PI_U}`);
  
  if (!(PI_L > 0 && PI_L < 1 && PI_U > 0 && PI_U < 1 && PI_L < PI_U)) {
    log.warn('Params', `Invalid PI levels (L=${PI_L}, U=${PI_U}); defaulting to 0.05 and 0.95.`);
    // NOTE: This warning doesn't actually reset the values! That's intentional - values pass through.
  }

  // Read 5 scenario factors (Enabled/Disabled)
  const factors = {
    mng_giving: String(pick('internal_mng_giving', 'Disabled')).trim(),
    product_marketing: String(pick('internal_product_marketing', 'Disabled')).trim(),
    org_stability: String(pick('internal_org_stability', 'Disabled')).trim(),
    govt_policy: String(pick('external_govt_ed_policy', 'Disabled')).trim(),
    macroeconomic: String(pick('external_macroeconomic', 'Disabled')).trim()
  };

  // Validate each factor is "Enabled" or "Disabled"
  for (const [key, val] of Object.entries(factors)) {
    const normalized = val.toLowerCase();
    if (normalized !== 'enabled' && normalized !== 'disabled') {
      log.warn('Params', `${key} must be "Enabled" or "Disabled", got "${val}". Defaulting to "Disabled".`);
      factors[key] = 'Disabled';
    } else {
      factors[key] = normalized === 'enabled' ? 'Enabled' : 'Disabled'; // Normalize casing
    }
  }

  // Read Global_Channel_Correlation
  let globalCorr = Number(pick('Global_Channel_Correlation', 0.3));
  if (!isFinite(globalCorr) || globalCorr < 0 || globalCorr > 1) {
    log.warn('Params', `Global_Channel_Correlation must be 0-1, got ${globalCorr}. Using 0.3.`);
    globalCorr = 0.3;
  }

  // NEW: Read SD_inflation_factor from Params
  // This can be entered as: 0.30 (decimal), 30 (percentage), or 1.30 (multiplier)
  // We interpret intelligently and return the MULTIPLIER value
  const sdInflationRaw = Number(pick('SD_inflation_factor', 0.30));
  let sdInflationFactor = 1.3; // default if parsing fails
  
  if (isFinite(sdInflationRaw) && sdInflationRaw >= 0) {
    if (sdInflationRaw < 1) {
      // Interpret as decimal percentage (e.g., 0.30 means 30% increase → 1.30x multiplier)
      sdInflationFactor = 1 + sdInflationRaw;
      log.info('Params', `SD_inflation_factor interpreted as ${(sdInflationRaw * 100).toFixed(0)}% increase (multiplier: ${sdInflationFactor.toFixed(2)})`);
    } else if (sdInflationRaw >= 1 && sdInflationRaw < 10) {
      // Interpret as direct multiplier (e.g., 1.30 means 1.30x)
      sdInflationFactor = sdInflationRaw;
      log.info('Params', `SD_inflation_factor interpreted as ${sdInflationFactor.toFixed(2)}x multiplier`);
    } else if (sdInflationRaw >= 10) {
      // Interpret as percentage number (e.g., 30 means 30% increase → 1.30x multiplier)
      sdInflationFactor = 1 + (sdInflationRaw / 100);
      log.info('Params', `SD_inflation_factor interpreted as ${sdInflationRaw.toFixed(0)}% increase (multiplier: ${sdInflationFactor.toFixed(2)})`);
    }
  } else {
    log.warn('Params', `SD_inflation_factor invalid (${sdInflationRaw}), using default 1.30x (30% increase)`);
  }

  // Sanity check: inflation factor should be reasonable (between 1.0 and 3.0)
  if (sdInflationFactor < 1.0 || sdInflationFactor > 3.0) {
    log.warn('Params', `SD_inflation_factor out of reasonable range [1.0, 3.0]: ${sdInflationFactor.toFixed(2)}. Clamping.`);
    sdInflationFactor = Math.max(1.0, Math.min(3.0, sdInflationFactor));
  }

  return { N_Sims, Seed, RoundK, PI_L, PI_U, factors, globalCorr, sdInflationFactor };
}

/** Reads the Params sheet as a 2-column table: Param | Value. Returns {ParamName: Value}. */
function readParamsTable(ss) {
  const sh = ss.getSheetByName('Params');
  const out = {};
  if (!sh) return out;
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return out;

  const hdr = values[0].map(x => String(x).trim());
  const iParam = hdr.indexOf('Param');
  const iValue = hdr.indexOf('Value');
  if (iParam < 0 || iValue < 0) return out;

  for (let i = 1; i < values.length; i++) {
    const key = String(values[i][iParam] ?? '').trim();
    if (!key) continue;
    out[key] = values[i][iValue];
  }
  return out;
}

/* ============================ INPUT READING ========================= */
/**
 * Reads all input tabs, validates structure, and precomputes per-row normal distribution parameters.
 * Returns a compact "data contract" object used by the simulation.
 */
function readAllInputs(ss, p, log) {
  // Fetch the required tabs
  const shF = mustSheet(ss, 'Forecasts', log);
  const shSM = mustSheet(ss, 'Scenario_Mapping', log);
  const shPlans = mustSheet(ss, 'Params_Plans', log);

  // Accept either "Annual_YoY_Bounds" or legacy "Annual_Caps_Floors"
  const shBounds = ss.getSheetByName('Annual_YoY_Bounds') || ss.getSheetByName('Annual_Caps_Floors');
  if (!shBounds) {
    const msg = 'Missing sheet: Annual_YoY_Bounds (or Annual_Caps_Floors)';
    if (log) log.error('Sheets', msg);
    throw new Error(msg);
  }

  // Read Bounds_Scenario_Adjustments (optional)
  const shBoundsAdj = ss.getSheetByName('Bounds_Scenario_Adjustments');
  const boundsAdjustments = shBoundsAdj ? readBoundsAdjustments(shBoundsAdj, log) : {};
  if (!shBoundsAdj) {
    log.warn('Sheets', 'Bounds_Scenario_Adjustments sheet missing. No scenario adjustments applied to bounds.');
  }

  // Read each sheet as a headered table
  const forecasts = readTable(shF);
  const scen = readTable(shSM);
  const plansT = readTable(shPlans);
  const boundsT = readTable(shBounds);

  // Build the scenario multipliers dictionary: per Channel×FY, with 5 binary factors
  // Now returns only meanMult (no SD logic here)
  const multipliers = buildScenarioMultipliers(scen, p ? p.factors : null, log);

  // Load plans keyed by year
  const planByYear = {};
  for (const r of plansT.rows) {
    const y = Number(r.Fiscal_Year);
    if (!y) continue;
    planByYear[y] = {
      Corporate: numOrZero(r.Corporate_Plan),
      Government: numOrZero(r.Government_Plan),
      Major: numOrZero(r.Major_Plan),
      Marketplace: numOrZero(r.Marketplace_Plan),
      Total: numOrZero(r.Total_Plan)
    };
    const sum = planByYear[y].Corporate + planByYear[y].Government + planByYear[y].Major + planByYear[y].Marketplace;
    if (Math.abs(sum - planByYear[y].Total) > 1e-6) {
      log.warn(`Plans FY${y}`, `Channel sum ${sum} != Total_Plan ${planByYear[y].Total}`);
    }
  }

  // Build YoY bounds map keyed by (channel, year)
  const bounds = {};
  for (const r of boundsT.rows) {
    const y = Number(r.Fiscal_Year);
    const c = cleanChannel(r.Channel);
    if (!y || !c) continue;

    // Accept multiple header variants
    let yoyMin = safeNum(r.YoY_Min, null);
    let yoyMax = safeNum(r.YoY_Max, null);
    if (yoyMin === null) yoyMin = safeNum(r.Annual_Floor_Pct, null);
    if (yoyMax === null) yoyMax = safeNum(r.Annual_Cap_Pct, null);

    const floorMaybePct = safeNum(r.Annual_Floor, null);
    const capMaybePct   = safeNum(r.Annual_Cap, null);
    if (yoyMin === null && floorMaybePct !== null && Math.abs(floorMaybePct) <= 1.5) yoyMin = floorMaybePct;
    if (yoyMax === null && capMaybePct   !== null && Math.abs(capMaybePct)   <= 1.5) yoyMax = capMaybePct;

    bounds[ck(c, y)] = { min: yoyMin, max: yoyMax };
  }

  // Enumerate rows from Forecasts; precompute per-row normal (mean, sd)
  const QUARTERS = ['Q1', 'Q2', 'Q3', 'Q4'];
  const chSet = new Set();
  const yearSet = new Set();
  const spec = {}; // key c|y|q -> {mean, sd, degenerate}

  // Count enabled factors for proportional SD inflation
  const enabledFactors = p ? Object.values(p.factors).filter(v => v === 'Enabled') : [];
  const numEnabled = enabledFactors.length;
  const totalFactors = p ? Object.values(p.factors).length : 5;

  // Proportional inflation: scales from 1.0 to sdInflationFactor based on % of factors enabled
  const baseInflationRate = p ? (p.sdInflationFactor - 1.0) : 0.50;
  const sdInflation = numEnabled > 0 ? (1.0 + (numEnabled / totalFactors) * baseInflationRate) : 1.0;

  if (numEnabled > 0 && log) {
    log.info('SD Inflation', `${numEnabled}/${totalFactors} factors enabled. SD multiplier: ${sdInflation.toFixed(2)}x`);
  }

  for (const r of forecasts.rows) {
    const c = cleanChannel(r.Channel);
    const y = Number(r.Fiscal_Year);
    const q = String(r.Quarter || '').toUpperCase();
    if (!c || !y || QUARTERS.indexOf(q) < 0) continue;

    chSet.add(c); yearSet.add(y);

    // Read base forecast values from Forecasts sheet
    const mean = Math.max(0, Number(r.Forecast_Mean || 0));
    const sd = Math.max(0, Number(r.Forecast_SD || 0));

    // Fetch scenario multipliers for this channel×year (MEAN only, no SD here)
    const cyKey = `${c}|${y}`;
    const mult = multipliers[cyKey];
    if (!mult) {
      log.warn('Multipliers', `No scenario multipliers for ${cyKey}; using mean as-is.`);
    }
    const meanMult = mult ? mult.meanMult : 1.0;

    // Apply mean multiplier to forecast mean
    const mAdj = mean * meanMult;
    
    // Apply SD inflation (proportional to number of factors enabled)
    const sAdj = sd * sdInflation;

    // Store normal distribution parameters OR mark as degenerate
    let deg = false;
    const tiny = 1e-12;

    if (mAdj <= 0 && sAdj <= tiny) {
      deg = true;  // Zero mean and zero SD => deterministic zero
    } else if (sAdj <= tiny) {
      if (log) log.info(`${c} ${y} ${q}`, `SD=0 → deterministic at mean ${mAdj}.`);
    }

    spec[ck(c, y, q)] = { mean: mAdj, sd: sAdj, degenerate: deg };
  }

  // Finalize channel and year sets
  const channels = Array.from(chSet).sort();
  const years = Array.from(yearSet).sort((a, b) => a - b);

  // Gentle heads-up if user provided FY26 bounds (ignored by design)
  if (years.length) {
    const firstY = years[0];
    for (const c of channels) {
      const b = bounds[ck(c, firstY)];
      if (b && (b.min !== null || b.max !== null)) {
        log.info(`Bounds ${c} FY${firstY}`, `FY${firstY} YoY bounds present but ignored (first modeled year).`);
      }
    }
  }

  // Return the compact structure the simulator needs
  return { spec, channels, years, QUARTERS, planByYear, bounds, boundsAdjustments };
}

/** Helper: returns a sheet or throws an Error (and logs) if missing. */
function mustSheet(ss, name, log) {
  const sh = ss.getSheetByName(name);
  if (!sh) {
    const msg = `Missing sheet: ${name}`;
    if (log) log.error('Sheets', msg);
    throw new Error(msg);
  }
  return sh;
}

/** Reads a sheet as a simple table {headers:[], rows:[{col:value,...},...]}. Empty rows are dropped. */
function readTable(sh) {
  const rng = sh.getDataRange();
  const values = rng.getValues();
  if (values.length < 2) return { headers: [], rows: [] };

  const headers = values[0].map(h => String(h || '').trim());
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const row = {};
    let any = false;
    for (let j = 0; j < headers.length; j++) {
      row[headers[j]] = values[i][j];
      if (values[i][j] !== '' && values[i][j] !== null) any = true;
    }
    if (any) rows.push(row);
  }
  return { headers, rows };
}

/**
 * Builds scenario multipliers for the 5-factor system with FY dimension.
 * Returns: { "Channel|FY": { meanMult } }
 * 
 * IMPORTANT: This ONLY handles MEAN multipliers now.
 * SD inflation is handled separately in readAllInputs based on Forecast_SD.
 */
function buildScenarioMultipliers(scen, factors, log) {
  const out = {};
  
  // Default factors if not provided (validation mode)
  const f = factors || { 
    mng_giving: 'Disabled', 
    product_marketing: 'Disabled', 
    org_stability: 'Disabled', 
    govt_policy: 'Disabled', 
    macroeconomic: 'Disabled' 
  };
  
  // Map Enabled/Disabled to Strong/Weak to match spreadsheet columns
  const mapState = (state) => (state === 'Enabled' ? 'Strong' : 'Weak');
  
  for (const r of scen.rows) {
    const c = cleanChannel(r.Channel);
    const fy = Number(r.Fiscal_Year);
    if (!c || !fy) continue;
    
    const key = `${c}|${fy}`;
    
    // Calculate mean multiplier by multiplying selected factor values
    // Note: Using "people" not "org_stability" to match sheet column naming
    let meanMult = 1.0;
    meanMult *= Number(r[`mng_giving_${mapState(f.mng_giving)}`] || 1);
    meanMult *= Number(r[`product_marketing_${mapState(f.product_marketing)}`] || 1);
    meanMult *= Number(r[`people_${mapState(f.org_stability)}`] || 1);  // "people" in sheet, "org_stability" in params
    meanMult *= Number(r[`govt_policy_${mapState(f.govt_policy)}`] || 1);
    meanMult *= Number(r[`macro_${mapState(f.macroeconomic)}`] || 1);
    
    // Optional: Add diagnostic logging to verify multipliers are being applied
    if (log && meanMult !== 1.0) {
      log.info(`Multipliers ${c} FY${fy}`, `Combined multiplier: ${meanMult.toFixed(4)}`);
    }
    
    // Return only meanMult (SD handling removed from here)
    out[key] = { meanMult };
  }
  
  return out;
}

/**
 * Reads Bounds_Scenario_Adjustments sheet.
 * Returns: { "Channel|Factor": { floorShift, capShift } }
 */
function readBoundsAdjustments(sh, log) {
  const table = readTable(sh);
  const adjustments = {};
  
  for (const row of table.rows) {
    const ch = cleanChannel(row.Channel);
    const factor = String(row.Factor || '').trim();
    if (!ch || !factor) continue;
    
    const key = `${ch}|${factor}`;
    adjustments[key] = {
      floorShift: numOrZero(row.Floor_Shift),
      capShift: numOrZero(row.Cap_Shift)
    };
  }
  
  return adjustments;
}

/**
 * Applies scenario-based adjustments to annual YoY bounds.
 * Returns adjusted bounds: { floor, cap }
 * 
 * IMPORTANT: This does NOT modify the Annual_YoY_Bounds sheet.
 * It calculates adjusted bounds in-memory for each trial.
 */
function applyScenarioAdjustments(baseBounds, channel, factors, adjustments) {
  let floorAdj = 0;
  let capAdj = 0;
  
  // Map user factor selections to internal names matching sheet
  const factorMap = {
    'internal_mng_giving': factors.mng_giving,
    'internal_product_marketing': factors.product_marketing,
    'internal_org_stability': factors.org_stability,
    'external_govt_ed_policy': factors.govt_policy,
    'external_macroeconomic': factors.macroeconomic
  };
  
  // Sum adjustments for all factors set to "Enabled"
  for (const [factorName, factorValue] of Object.entries(factorMap)) {
    if (factorValue === 'Enabled') {
      const key = `${channel}|${factorName}`;
      const adj = adjustments[key];
      if (adj) {
        floorAdj += adj.floorShift;
        capAdj += adj.capShift;
      }
    }
  }
  
  // Apply to base bounds (does not modify sheet values)
  return {
    floor: baseBounds.floor + floorAdj,
    cap: baseBounds.cap + capAdj
  };
}

/* ============================ SIMULATION ============================ */
/**
 * The main simulation loop:
 * - Loops over trials with deterministic PRNG
 * - For each year & quarter, draws one common shock for global correlation
 * - For each channel, combines common + idiosyncratic shocks using globalCorr
 * - Applies within-year correlation (WITHIN_YEAR_RHO) between quarters
 * - After 4 quarters, applies YoY bounds (with scenario adjustments) by scaling
 * - Updates streaming statistics for quarterly, annual (channel), and annual (total)
 * - Writes outputs at the end (rounded on write if enabled)
 */
function simulateAndWrite(ss, p, d, log) {
  const { spec, channels, years, QUARTERS, planByYear, bounds, boundsAdjustments } = d;
  if (!years.length || !channels.length) {
    log.error('Inputs', 'No years or channels found. Nothing to simulate.');
    return;
  }

  // Prepare streaming aggregators
  const pLow = p.PI_L > 0 && p.PI_L < 1 ? p.PI_L : 0.05;
  const pUp  = p.PI_U > 0 && p.PI_U < 1 ? p.PI_U  : 0.95;
  const ps = [pLow, 0.25, 0.50, 0.75, pUp];
  
  // DIAGNOSTIC: Log percentiles being used
  log.info('Simulation', `Percentiles array: [${ps.join(', ')}]`);
  log.info('Simulation', `Header names will be: ${ps.map(p => 'P' + String(Math.round(p * 100)).padStart(2, '0')).join(', ')}`);

  const qAgg = {};   // quarterly per (c,y,q)
  const aAgg = {};   // annual channel per (c,y)
  const tAgg = {};   // annual total per y
  const piAgg = {};  // PI per (c,y)
  const piTot = {};  // PI per y

  // Initialize aggregators
  for (const y of years) {
    for (const c of channels) {
      for (const q of QUARTERS) qAgg[ck(c, y, q)] = makeAgg(ps);
      aAgg[ck(c, y)] = makeAgg(ps, true);
      piAgg[ck(c, y)] = makePI(pLow, pUp);
    }
    tAgg[y] = makeAgg(ps, true);
    piTot[y] = makePI(pLow, pUp);
  }

  // Deterministic PRNG
  const rng = mulberry32(p.Seed);
  const nTrials = p.N_Sims;

  // Storage for within-year correlation (previous quarter's draws per channel)
  const channelOrder = channels.slice().sort(); // Consistent ordering
  
  // MAIN TRIAL LOOP
  for (let t = 0; t < nTrials; t++) {
    const prevAnnual = {}; // Track prior year's annual per channel
    const prevQuarterZ = {}; // Track previous quarter's Z per channel for within-year correlation

    for (const y of years) {
      // Reset quarterly correlation storage for new year
      for (const c of channelOrder) prevQuarterZ[c] = null;

      // Generate raw quarterly values (pre-bounds) for this year
      const cqValue = {}; // key c|q => value
      
      for (const q of QUARTERS) {
        // Single common shock for global correlation
        const Z_common = randn(rng);
        
        for (const c of channelOrder) {
          // Generate idiosyncratic shock
          const Z_idio = randn(rng);

          // Quarter-specific correlated innovation (keeps global cross-channel correlation every quarter)
          const Z_innov = Math.sqrt(p.globalCorr) * Z_common + Math.sqrt(1 - p.globalCorr) * Z_idio;

          // Apply within-year correlation (AR(1) on the full shock with correlated innovation)
          const Z = (prevQuarterZ[c] === null)
            ? Z_innov
            : Math.sqrt(WITHIN_YEAR_RHO) * prevQuarterZ[c] + Math.sqrt(1 - WITHIN_YEAR_RHO) * Z_innov;

          prevQuarterZ[c] = Z; // Store for next quarter
          
          // Sample from truncated normal (apply floor from Floor_Zero param)
          const sp = spec[ck(c, y, q)];
          if (!sp || sp.degenerate) {
            cqValue[ck(c, q)] = 0;
          } else {
            const rawSample = sp.mean + sp.sd * Z;
            const floor = p.FloorZero ? 0 : -Infinity; // Respect Floor_Zero parameter
            cqValue[ck(c, q)] = Math.max(floor, rawSample);
          }
        }
      }

      // Apply annual YoY bounds with scenario adjustments
      const isFirstYear = (y === years[0]);
      let totalAnnual = 0;

      for (const c of channelOrder) {
        const qVals = QUARTERS.map(q => cqValue[ck(c, q)] || 0);
        let annual = qVals.reduce((a, b) => a + b, 0);

        // Enforce bounds for FY27+ if present
        const b = bounds[ck(c, y)];
        if (!isFirstYear && b && (b.min !== null || b.max !== null)) {
          const prev = prevAnnual[c] || 0;
          
          // Apply scenario adjustments to base bounds (in-memory only, does not modify sheet)
          const baseBounds = { floor: b.min, cap: b.max };
          const adjustedBounds = applyScenarioAdjustments(baseBounds, c, p.factors, boundsAdjustments);
          
          let minBound = (adjustedBounds.floor !== null && isFinite(adjustedBounds.floor)) ? prev * (1 + adjustedBounds.floor) : -Infinity;
          let maxBound = (adjustedBounds.cap !== null && isFinite(adjustedBounds.cap)) ? prev * (1 + adjustedBounds.cap) : +Infinity;

          // Special rule: if prior year is zero, allow recovery
          if (prev === 0) {
            minBound = Math.max(minBound, 0);
            maxBound = +Infinity;
          }

          // Guardrail: inconsistent bounds after adjustments (min > max)
          if (isFinite(minBound) && isFinite(maxBound) && minBound > maxBound) {
            log.error(
              'Bounds',
              `Inconsistent adjusted bounds for ${c} ${y}: minBound (${minBound}) > maxBound (${maxBound}). ` +
              `Base[min=${b.min}, max=${b.max}] Adjusted[floor=${adjustedBounds.floor}, cap=${adjustedBounds.cap}] prev=${prev}. ` +
              `Skipping bounds enforcement for this channel/year.`
            );
            minBound = -Infinity;
            maxBound = +Infinity;
          }

          // Apply floor
          if (annual < minBound && isFinite(minBound) && minBound > -Infinity) {
            if (annual <= 0) {
              if (minBound > 0) {
                const eq = minBound / 4;
                for (let i = 0; i < QUARTERS.length; i++) qVals[i] = eq;
                annual = minBound;
              }
            } else {
              const scale = minBound / annual;
              for (let i = 0; i < QUARTERS.length; i++) qVals[i] *= scale;
              annual = minBound;
            }
          }

          // Apply cap
          if (annual > maxBound && isFinite(maxBound) && maxBound < +Infinity) {
            if (annual > 0) {
              const scale = maxBound / annual;
              for (let i = 0; i < QUARTERS.length; i++) qVals[i] *= scale;
              annual = maxBound;
            }
          }
        }

        prevAnnual[c] = annual;
        totalAnnual += annual;

        // Update aggregators
        const AA = aAgg[ck(c, y)];
        AA.mean.push(annual);
        for (const p2 of AA.p2s) p2.push(annual);

        const plan = (planByYear[y] && planByYear[y][c]) || 0;
        if (annual >= plan) AA.hits++;
        if (annual >= 0.9 * plan) AA.hits10++;
        AA.n++;

        const PI = piAgg[ck(c, y)];
        PI.low.push(annual);
        PI.up.push(annual);

        // Update quarterly aggregators
        for (let i = 0; i < QUARTERS.length; i++) {
          const qk = ck(c, y, QUARTERS[i]);
          const v = qVals[i];
          const QA = qAgg[qk];
          QA.mean.push(v);
          for (const p2 of QA.p2s) p2.push(v);
        }
      }

      // Update total aggregators
      const TA = tAgg[y];
      TA.mean.push(totalAnnual);
      for (const p2 of TA.p2s) p2.push(totalAnnual);

      const tPlan = (planByYear[y] && planByYear[y].Total) || 0;
      if (totalAnnual >= tPlan) TA.hits++;
      if (totalAnnual >= 0.9 * tPlan) TA.hits10++;
      TA.n++;

      piTot[y].low.push(totalAnnual);
      piTot[y].up.push(totalAnnual);
    }
  }

  // Write all outputs
  writeQuarterly(ss, qAgg, d, p.RoundK, ps);
  writeAnnually(ss, aAgg, tAgg, d, p.RoundK, ps);
  writePI(ss, piAgg, piTot, d, p.RoundK);
}

/* ===================== AGGREGATORS & QUANTILES ====================== */
/** A simple streaming running mean (no raw storage). */
function RunningMean() { this.n = 0; this.mu = 0; }
RunningMean.prototype.push = function (x) { this.n++; this.mu += (x - this.mu) / this.n; };
RunningMean.prototype.value = function () { return this.mu; };

/**
 * P² quantile estimator for a single quantile p in (0,1).
 * Incrementally maintains 5 marker heights and positions using parabolic interpolation.
 */
function P2Quantile(p) {
  this.p = p;
  this.count = 0;
  this.q = [0, 0, 0, 0, 0];        // marker heights
  this.n = [1, 2, 3, 4, 5];        // marker positions (1-indexed conceptually)
  this.nPrime = [1, 1 + 2 * p, 1 + 4 * p, 3 + 2 * p, 5]; // desired positions
  this.dn = [0, p / 2, p, (1 + p) / 2, 1]; // increments
}

P2Quantile.prototype.push = function (x) {
  if (this.count < 5) {
    this.q[this.count] = x;
    this.count++;
    if (this.count === 5) this.q.sort((a, b) => a - b);
    return;
  }

  // Locate cell k: k is the index of the marker BELOW x (0..3)
  let k = 0;

  if (x < this.q[0]) {
    this.q[0] = x;
    k = 0;
  } else if (x < this.q[1]) {
    k = 0;
  } else if (x < this.q[2]) {
    k = 1;
  } else if (x < this.q[3]) {
    k = 2;
  } else if (x < this.q[4]) {
    k = 3;
  } else {
    // x >= current max
    this.q[4] = x;
    k = 3;
  }

  // Increment positions for markers k+1..4 (this ALWAYS increments n[4])
  for (let i = k + 1; i < 5; i++) this.n[i]++;

  // Update desired positions
  for (let i = 0; i < 5; i++) this.nPrime[i] += this.dn[i];

  // Adjust marker heights with P² algorithm
  for (let i = 1; i < 4; i++) {
    const d = this.nPrime[i] - this.n[i];
    if ((d >= 1 && (this.n[i + 1] - this.n[i]) > 1) || (d <= -1 && (this.n[i - 1] - this.n[i]) < -1)) {
      const dInt = (d >= 0) ? 1 : -1;
      const qNew = this.parabolic(i, dInt);
      if (this.q[i - 1] < qNew && qNew < this.q[i + 1]) {
        this.q[i] = qNew;
      } else {
        this.q[i] = this.linear(i, dInt);
      }
      this.n[i] += dInt;
    }
  }
};

P2Quantile.prototype.parabolic = function (i, d) {
  const qi = this.q[i], qPrev = this.q[i - 1], qNext = this.q[i + 1];
  const ni = this.n[i], nPrev = this.n[i - 1], nNext = this.n[i + 1];
  return qi + (d / (nNext - nPrev)) * (
    (ni - nPrev + d) * (qNext - qi) / (nNext - ni) +
    (nNext - ni - d) * (qi - qPrev) / (ni - nPrev)
  );
};

P2Quantile.prototype.linear = function (i, d) {
  return this.q[i] + d * (this.q[i + d] - this.q[i]) / (this.n[i + d] - this.n[i]);
};

P2Quantile.prototype.value = function () {
  return (this.count < 5) ? this.q[Math.floor(this.count / 2)] : this.q[2];
};

/** Factory for an aggregator bundle with mean + percentiles (optionally plan hit-rates). */
function makeAgg(ps, withHits) {
  const a = {
    mean: new RunningMean(),
    p2s: ps.map(p => new P2Quantile(p))
  };
  if (withHits) { a.hits = 0; a.hits10 = 0; a.n = 0; }
  return a;
}

/** Factory for PI aggregators (just two P² quantiles). */
function makePI(pLow, pUp) {
  return { low: new P2Quantile(pLow), up: new P2Quantile(pUp) };
}

/* ============================ OUTPUT WRITING ======================== */
/** Writes Results_Quarterly with correct column order and dynamic percentile headers. */
function writeQuarterly(ss, qAgg, d, doRound, percentiles) {
  const sh = ss.getSheetByName('Results_Quarterly');
  if (!sh) return;

  // Generate percentile column names from actual percentiles used
  const pNames = percentiles.map(p => 'P' + String(Math.round(p * 100)).padStart(2, '0'));
  
  // Define expected headers
  const expectedHeaders = ['Fiscal_Year', 'Channel', 'Quarter', 'Mean'].concat(pNames);
  
  // Update header row
  sh.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);

  // Clear only data rows (keep headers)
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow > 1 && lastCol > 0) sh.getRange(2, 1, lastRow - 1, lastCol).clearContent();

  const out = [];
  const keys = Object.keys(qAgg).sort(function (a, b) {
    const [ca, ya, qa] = a.split('|'); const [cb, yb, qb] = b.split('|');
    const ycmp = Number(ya) - Number(yb); if (ycmp) return ycmp;
    const ccmp = ca.localeCompare(cb); if (ccmp) return ccmp;
    return qa.localeCompare(qb);
  });

  for (const k of keys) {
    const [c, ys, q] = k.split('|');
    const y = Number(ys);
    const A = qAgg[k];
    out.push([
      y, c, q,
      roundK(A.mean.value(), doRound),
      roundK(A.p2s[0].value(), doRound),
      roundK(A.p2s[1].value(), doRound),
      roundK(A.p2s[2].value(), doRound),
      roundK(A.p2s[3].value(), doRound),
      roundK(A.p2s[4].value(), doRound)
    ]);
  }
  if (out.length) sh.getRange(2, 1, out.length, 9).setValues(out);
}

/** Writes Results_Annually with correct column order. */
function writeAnnually(ss, aAgg, tAgg, d, doRound, percentiles) {
  const sh = ss.getSheetByName('Results_Annually');
  if (!sh) return;

  const pNames = percentiles.map(p => 'P' + String(Math.round(p * 100)).padStart(2, '0'));
  const expectedHeaders = ['Fiscal_Year', 'Channel', 'Mean'].concat(pNames).concat(['Plan', 'Plan_Likelihood', 'Plan_Likelihood_within10%']);
  sh.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);

  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow > 1 && lastCol > 0) sh.getRange(2, 1, lastRow - 1, lastCol).clearContent();

  const channelRows = [];
  const keys = Object.keys(aAgg).sort(function(a, b) {
    const [ca, ya] = a.split('|'); const [cb, yb] = b.split('|');
    const ccmp = ca.localeCompare(cb); if (ccmp) return ccmp;
    return Number(ya) - Number(yb);
  });

  for (const k of keys) {
    const [c, ys] = k.split('|');
    const y = Number(ys);
    const A = aAgg[k];
    const planYear = d.planByYear[y] || {};
    const plan = (planYear && Object.prototype.hasOwnProperty.call(planYear, c)) ? planYear[c] : '';

    channelRows.push([
      y, c,
      roundK(A.mean.value(), doRound),
      roundK(A.p2s[0].value(), doRound),
      roundK(A.p2s[1].value(), doRound),
      roundK(A.p2s[2].value(), doRound),
      roundK(A.p2s[3].value(), doRound),
      roundK(A.p2s[4].value(), doRound),
      plan,
      (A.n && A.n > 0) ? (A.hits / A.n) : 0,
      (A.n && A.n > 0) ? (A.hits10 / A.n) : 0
    ]);
  }

  const totalRows = [];
  const tYears = Object.keys(tAgg).map(Number).sort((a, b) => a - b);
  for (const y of tYears) {
    const A = tAgg[y];
    const planYear = d.planByYear[y] || {};
    const totalPlan = (planYear && Object.prototype.hasOwnProperty.call(planYear, 'Total')) ? planYear['Total'] : '';

    totalRows.push([
      y, 'Total',
      roundK(A.mean.value(), doRound),
      roundK(A.p2s[0].value(), doRound),
      roundK(A.p2s[1].value(), doRound),
      roundK(A.p2s[2].value(), doRound),
      roundK(A.p2s[3].value(), doRound),
      roundK(A.p2s[4].value(), doRound),
      totalPlan,
      (A.n && A.n > 0) ? (A.hits / A.n) : 0,
      (A.n && A.n > 0) ? (A.hits10 / A.n) : 0
    ]);
  }

  const out = channelRows.concat(totalRows);
  if (out.length) sh.getRange(2, 1, out.length, 11).setValues(out);
}

/** Writes Results_PI with correct column order. */
function writePI(ss, piAgg, piTot, d, doRound) {
  const sh = ss.getSheetByName('Results_PI');
  if (!sh) return;

  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, lastCol).clearContent();

  const rows = [];
  const keys = Object.keys(piAgg).sort(function(a, b) {
    const [ca, ya] = a.split('|'); const [cb, yb] = b.split('|');
    const ycmp = Number(ya) - Number(yb); if (ycmp) return ycmp;
    return ca.localeCompare(cb);
  });

  for (const k of keys) {
    const [c, ys] = k.split('|');
    const y = Number(ys);
    const A = piAgg[k];
    rows.push([
      y, c,
      roundK(A.low.value(), doRound),
      roundK(A.up.value(), doRound)
    ]);
  }

  const tYears = Object.keys(piTot).map(Number).sort((a, b) => a - b);
  for (const y of tYears) {
    const A = piTot[y];
    rows.push([
      y, 'Total',
      roundK(A.low.value(), doRound),
      roundK(A.up.value(), doRound)
    ]);
  }

  if (rows.length) sh.getRange(2, 1, rows.length, 4).setValues(rows);
}

/* =============================== LOGGER ============================= */
/** Simple append-only logger. */
function makeLogger(ss) {
  const sh = ss.getSheetByName('Log') || ss.insertSheet('Log');
  const write = (lvl, key, msg) => {
    const ts = new Date();
    sh.appendRow([ts.toISOString(), String(key || ''), `[${lvl}] ${String(msg || '')}`]);
  };
  return {
    info: (k, m) => write('INFO', k, m),
    warn: (k, m) => write('WARN', k, m),
    error: (k, m) => write('ERROR', k, m),
    flush: () => {}
  };
}

/* ============================== UTILITIES =========================== */
/** Compact key builder. */
function ck(c, y, q) {
  if (q !== undefined) return `${c}|${y}|${q}`;
  if (y !== undefined) return `${c}|${y}`;
  return String(c);
}

/** Normalizes a channel label. */
function cleanChannel(v) {
  const s = String(v || '').trim();
  if (!s) return null;
  const x = s.toLowerCase();
  if (x.startsWith('corp')) return 'Corporate';
  if (x.startsWith('gov')) return 'Government';
  if (x.startsWith('maj')) return 'Major';
  if (x.startsWith('market')) return 'Marketplace';
  if (['Corporate','Government','Major','Marketplace'].indexOf(s) >= 0) return s;
  return s;
}

/** Numeric helpers */
function numOrZero(v) { return (isFinite(Number(v)) ? Number(v) : 0); }
function safeNum(v, dflt) { const n = Number(v); return isFinite(n) ? n : dflt; }

/** Display rounding. */
function roundK(x, on) { return on ? (Math.floor(x / 1000) * 1000) : x; }

/** Deterministic PRNG (Mulberry32). */
function mulberry32(a) {
  return function () {
    a |= 0; a = a + 0x6D2B79F5 | 0;
    let t = Math.imul(a ^ a >>> 15, 1 | a);
    t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t;
    return ((t ^ t >>> 14) >>> 0) / 4294967296;
  }
}

/** Standard normal via Box–Muller transform. */
function randn(rng) {
  let u = 0, v = 0;
  while (u === 0) u = rng();
  while (v === 0) v = rng();
  return Math.sqrt(-2.0 * Math.log(u)) * Math.cos(2.0 * Math.PI * v);
}