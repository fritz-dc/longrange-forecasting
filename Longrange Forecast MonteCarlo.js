/** ======================================================================
 * DonorsChoose Monte Carlo for Google Sheets — Fully Annotated Version
 * ----------------------------------------------------------------------
 * WHAT THIS DOES (executive summary):
 * - Reads your planning inputs from several sheets (Params, Forecasts, etc.)
 * - Builds scenario-adjusted lognormal distributions per Channel×Year×Quarter
 * - Simulates many trials with seeded randomness, including cross-channel
 *   co-movement via common shocks (economy/government factors)
 * - Applies YoY bounds AT THE ANNUAL LEVEL by proportionally scaling quarters
 * - Aggregates streaming means and percentiles + plan hit-rates
 * - Writes summary tables to Results_Quarterly / Results_Annually / Results_PI
 *
 * WHY IT'S SAFE:
 * - Only writes to dedicated Results_* tabs (clears data rows, keeps headers)
 * - Uses deterministic seed for reproducibility
 * - Streaming quantiles => fast & memory-safe (no storing trial arrays)
 *
 * YOU CONFIGURE:
 * - Params tab (N_Sims, Seed, rounding, scenario, PI levels, Sigma_Log_Max)
 * - Forecasts tab (per Channel×Year×Quarter mean & predictive SD)
 * - Scenario_Mapping tab (Headwind/Neutral/Tailwind multipliers for mean/SD)
 * - Correlation_Config tab (w_Economy, w_Government per channel)
 * - Annual_YoY_Bounds tab (YoY_Min, YoY_Max per channel/year)
 * - Params_Plans tab (annual plan per channel + Total_Plan)
 *
 * KEY MODEL RULES:
 * - Lognormal only (non-negative, right-skewed; realistic for revenue)
 * - σ_log is capped (default 2.5) to tame extreme right tails
 * - Common shocks reallocate variance into shared/idiosyncratic pieces
 *   WITHOUT changing the marginal SD you provided (Var(Z)=1)
 * - YoY bounds: FY26 exempt; FY27+ bounded versus prior simulated year
 *   Special case: if prior=0, we ignore the upper bound to allow recovery
 *
 * DEPLOYMENT:
 * - Paste this file into Apps Script, save, refresh the sheet
 * - Use the "Monte Carlo" custom menu to run/validate/clear
 * ====================================================================== */

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

  // Parse each param with sensible defaults (these are used if neither named range nor table supplies them)
  const N_Sims = Math.max(1, Math.floor(Number(pick('N_Sims', 10000))));
  const Seed = Math.floor(Number(pick('Seed', 42)));
  const Dist = String(pick('Distribution_Type', 'Lognormal'));
  if (Dist.toLowerCase() !== 'lognormal') log.warn('Params', 'Distribution_Type forced to Lognormal (engine is lognormal-only).');
  const RoundK = asBool(pick('Round_To_Thousands', true));
  const Econ = String(pick('Economy_Factor', 'Neutral'));
  const Govt = String(pick('Government_Factor', 'Neutral'));
  const PI_L = Number(pick('PI_Lower_Level', 0.20));
  const PI_U = Number(pick('PI_Upper_Level', 0.80));
  if (!(PI_L > 0 && PI_L < 1 && PI_U > 0 && PI_U < 1 && PI_L < PI_U)) {
    log.warn('Params', 'Invalid PI levels; defaulting to 0.20 and 0.80.');
  }
  const SIG_CAP = Number(pick('Sigma_Log_Max', 2.5));

  return { N_Sims, Seed, RoundK, Econ, Govt, PI_L, PI_U, SIG_CAP };
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
 * Reads all input tabs, validates structure, and precomputes per-row lognormal parameters.
 * Returns a compact "data contract" object used by the simulation.
 */
function readAllInputs(ss, p, log) {
  // Fetch the required tabs (raising if genuinely missing)
  const shF = mustSheet(ss, 'Forecasts', log);
  const shSM = mustSheet(ss, 'Scenario_Mapping', log);
  const shCC = mustSheet(ss, 'Correlation_Config', log);
  const shPlans = mustSheet(ss, 'Params_Plans', log);

  // Accept either "Annual_YoY_Bounds" or legacy "Annual_Caps_Floors"
  const shBounds = ss.getSheetByName('Annual_YoY_Bounds') || ss.getSheetByName('Annual_Caps_Floors');
  if (!shBounds) {
    const msg = 'Missing sheet: Annual_YoY_Bounds (or Annual_Caps_Floors)';
    if (log) log.error('Sheets', msg);
    throw new Error(msg);
  }

  // Read each sheet as a headered table (rows[] items keyed by header names)
  const forecasts = readTable(shF);
  const scen = readTable(shSM);
  const weights = parseWeights(shCC, log);            // Robust header detection for w_Economy/w_Government
  const plansT = readTable(shPlans);
  const boundsT = readTable(shBounds);

  // Build the scenario multipliers dictionary: per channel, mean/SD multipliers per Econ/Govt × Headwind/Neutral/Tailwind
  const multipliers = buildScenarioMultipliers(scen, log);

  // Load plans keyed by year; also sanity-check sum(channel) vs Total_Plan
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

  // Build YoY bounds map keyed by (channel, year), accepting multiple header variants
  const bounds = {};
  for (const r of boundsT.rows) {
    const y = Number(r.Fiscal_Year);
    const c = cleanChannel(r.Channel);
    if (!y || !c) continue;

    // Primary expected headers
    let yoyMin = safeNum(r.YoY_Min, null);
    let yoyMax = safeNum(r.YoY_Max, null);

    // Accept alternatives often seen in legacy tabs
    if (yoyMin === null) yoyMin = safeNum(r.Annual_Floor_Pct, null);
    if (yoyMax === null) yoyMax = safeNum(r.Annual_Cap_Pct, null);

    // Back-compat: if Annual_Floor / Annual_Cap look like percentages (-1..+1), accept as YoY %
    const floorMaybePct = safeNum(r.Annual_Floor, null);
    const capMaybePct   = safeNum(r.Annual_Cap, null);
    if (yoyMin === null && floorMaybePct !== null && Math.abs(floorMaybePct) <= 1.5) yoyMin = floorMaybePct;
    if (yoyMax === null && capMaybePct   !== null && Math.abs(capMaybePct)   <= 1.5) yoyMax = capMaybePct;

    bounds[ck(c, y)] = { min: yoyMin, max: yoyMax };
  }

  // Enumerate rows from Forecasts; precompute per-row lognormal (mu, sigma) post-scenario
  const QUARTERS = ['Q1', 'Q2', 'Q3', 'Q4'];
  const chSet = new Set();
  const yearSet = new Set();
  const spec = {}; // key c|y|q -> {mu, sigma, degenerate}

  for (const r of forecasts.rows) {
    const c = cleanChannel(r.Channel);
    const y = Number(r.Fiscal_Year);
    const q = String(r.Quarter || '').toUpperCase();
    if (!c || !y || QUARTERS.indexOf(q) < 0) continue;

    chSet.add(c); yearSet.add(y);

    // Input constraints: means and predictive SDs must be non-negative
    const mean = Math.max(0, Number(r.Forecast_Mean || 0));
    const sd = Math.max(0, Number(r.Forecast_SD || 0));

    // Fetch scenario multipliers for this channel; default to 1.0 if not present
    const mul = (multipliers[c] || { mean: { Econ: oneSet(), Govt: oneSet() }, sd: { Econ: oneSet(), Govt: oneSet() } });
    const meanMult = (mul.mean.Econ[p ? p.Econ : 'Neutral'] || 1) * (mul.mean.Govt[p ? p.Govt : 'Neutral'] || 1);
    const sdMult   = (mul.sd.Econ[p ? p.Econ : 'Neutral'] || 1)   * (mul.sd.Govt[p ? p.Govt : 'Neutral'] || 1);

    // Apply scenario multipliers
    const mAdj = mean * meanMult;
    const sAdj = sd * sdMult;

    // Parameterize lognormal (mu, sigma) OR mark as degenerate (deterministic 0) if the row can't form a sensible distribution
    let mu = 0, sigma = 0, deg = false;
    const tiny = 1e-12;

    if (mAdj <= 0) {
      // Truly zero mean => zero out deterministically
      deg = true;
    } else {
      // If SD is zero (or blank → 0), treat as deterministic at the mean:
      // sigma=0 => exp(mu + 0*Z) = exp(mu) = mAdj
      let sigmaLog = 0;

      if (sAdj > tiny) {
        const cv = sAdj / Math.max(mAdj, tiny);
        sigmaLog = Math.sqrt(Math.log(1 + cv * cv));
        if (p && sigmaLog > p.SIG_CAP) {
          log.info(`${c} ${y} ${q}`, `sigma_log capped from ${sigmaLog.toFixed(4)} to ${p.SIG_CAP}`);
          sigmaLog = p.SIG_CAP;
        }
      } else {
        // SD==0 → deterministic; helpful breadcrumb in the Log
        if (log) log.info(`${c} ${y} ${q}`, `SD=0 → deterministic at mean ${mAdj}.`);
      }

      sigma = sigmaLog;
      // With sigma=0, this simplifies to mu = ln(mAdj)
      mu = Math.log(Math.max(mAdj, tiny)) - 0.5 * sigmaLog * sigmaLog;
    }

    spec[ck(c, y, q)] = { mu, sigma, degenerate: deg };
  }

  // Finalize channel and year sets (sorted for stable output ordering)
  const channels = Array.from(chSet).sort();
  const years = Array.from(yearSet).sort((a, b) => a - b);

  // Normalize weights per channel and precompute residual variance share
  const wByC = {};
  for (const c of channels) {
    const wE = (weights[c] && isFinite(weights[c].wE)) ? Number(weights[c].wE) : 0;
    const wG = (weights[c] && isFinite(weights[c].wG)) ? Number(weights[c].wG) : 0;
    const s2 = wE * wE + wG * wG;

    let wEn = wE, wGn = wG;
    if (s2 > 1) {
      // Auto-shrink proportionally so that w_E^2 + w_G^2 == 1 (preserves direction, fixes magnitude)
      const k = 1 / Math.sqrt(s2);
      wEn = wE * k; wGn = wG * k;
      log.warn(`Weights ${c}`, `w_E^2 + w_G^2 = ${s2.toFixed(4)} > 1; auto-shrunk to (${wEn.toFixed(3)}, ${wGn.toFixed(3)})`);
    }
    // Residual weight ensures Var(Z)=1
    wByC[c] = { wE: wEn, wG: wGn, wR: Math.sqrt(Math.max(0, 1 - (wEn * wEn + wGn * wGn))) };
  }

  // Gentle heads-up if user mistakenly provided FY26 bounds (ignored by design)
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
  return { spec, channels, years, QUARTERS, wByC, planByYear, bounds };
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

/** Builds the Scenario_Mapping structure: for each channel, mean/sd multipliers by Econ/Govt × (Headwind,Neutral,Tailwind). */
function buildScenarioMultipliers(scen, log) {
  const out = {};
  for (const r of scen.rows) {
    const c = cleanChannel(r.Channel);
    if (!c) continue;
    const get = (k, d) => (isFinite(Number(r[k])) ? Number(r[k]) : d);
    out[c] = {
      mean: {
        Econ: { Headwind: get('Econ_Mean_Headwind', 1), Neutral: get('Econ_Mean_Neutral', 1), Tailwind: get('Econ_Mean_Tailwind', 1) },
        Govt: { Headwind: get('Govt_Mean_Headwind', 1), Neutral: get('Govt_Mean_Neutral', 1), Tailwind: get('Govt_Mean_Tailwind', 1) }
      },
      sd: {
        Econ: { Headwind: get('Econ_SD_Headwind', 1), Neutral: get('Econ_SD_Neutral', 1), Tailwind: get('Econ_SD_Tailwind', 1) },
        Govt: { Headwind: get('Govt_SD_Headwind', 1), Neutral: get('Govt_SD_Neutral', 1), Tailwind: get('Govt_SD_Tailwind', 1) }
      }
    };
  }
  return out;
}

/**
 * Parses Correlation_Config to extract per-channel w_Economy and w_Government.
 * Robust against having an older "WithinYear_Rho" table stacked above — we scan
 * for the header row that contains w_Economy and w_Government.
 */
function parseWeights(shCC, log) {
  const values = shCC.getDataRange().getValues();
  // Find the header row which contains both w_Economy and w_Government
  let headerRow = -1;
  for (let i = 0; i < values.length; i++) {
    const row = values[i].map(x => String(x || '').trim().toLowerCase());
    if (row.indexOf('w_economy') >= 0 && row.indexOf('w_government') >= 0 && row.indexOf('channel') >= 0) {
      headerRow = i;
      break;
    }
  }
  const out = {};
  if (headerRow < 0 || headerRow + 1 >= values.length) {
    log.warn('Correlation_Config', 'Weights table not found; defaulting all weights to 0.');
    return out; // means all channels default to independence
  }

  // Build column indices
  const hdr = values[headerRow].map(x => String(x || '').trim());
  const iC = hdr.findIndex(h => h.toLowerCase() === 'channel');
  const iE = hdr.findIndex(h => h.toLowerCase() === 'w_economy');
  const iG = hdr.findIndex(h => h.toLowerCase() === 'w_government');

  // Read subsequent rows until blank
  for (let i = headerRow + 1; i < values.length; i++) {
    const c = cleanChannel(values[i][iC]);
    if (!c) continue;
    out[c] = { wE: safeNum(values[i][iE], 0), wG: safeNum(values[i][iG], 0) };
  }
  return out;
}

/* ============================ SIMULATION ============================ */
/**
 * The main simulation loop:
 * - Loops over trials with deterministic PRNG
 * - For each year & quarter, draws common shocks E,G
 * - For each channel, builds standardized shock Z with Var=1, then draws lognormal value
 * - After 4 quarters, applies YoY bounds by scaling quarters if needed
 * - Updates streaming statistics for quarterly, annual (channel), and annual (total)
 * - Writes outputs at the end (rounded on write if enabled)
 */
function simulateAndWrite(ss, p, d, log) {
  const { spec, channels, years, QUARTERS, wByC, planByYear, bounds } = d;
  if (!years.length || !channels.length) {
    log.error('Inputs', 'No years or channels found. Nothing to simulate.');
    return;
  }

  // Prepare streaming aggregators (means & P² quantiles) and hit-rate counters
  const pLow = p.PI_L > 0 && p.PI_L < 1 ? p.PI_L : 0.20;
  const pUp  = p.PI_U > 0 && p.PI_U < 1 ? p.PI_U  : 0.80;
  const ps = [pLow, 0.25, 0.50, 0.75, pUp];                // required percentiles (now using dynamic values)

  const qAgg = {};   // quarterly per (c,y,q) -> { mean, p2s[], hits?, n? }
  const aAgg = {};   // annual channel per (c,y)
  const tAgg = {};   // annual total per y
  const piAgg = {};  // PI per (c,y)
  const piTot = {};  // PI per y

  // Initialize aggregators keyed by all combinations we will produce
  for (const y of years) {
    for (const c of channels) {
      for (const q of QUARTERS) qAgg[ck(c, y, q)] = makeAgg(ps);
      aAgg[ck(c, y)] = makeAgg(ps, true);
      piAgg[ck(c, y)] = makePI(pLow, pUp);
    }
    tAgg[y] = makeAgg(ps, true);
    piTot[y] = makePI(pLow, pUp);
  }

  // Deterministic PRNG seeded from Params.Seed
  const rng = mulberry32(p.Seed);
  const nTrials = p.N_Sims;

  // MAIN TRIAL LOOP
  for (let t = 0; t < nTrials; t++) {
    // Track prior year's annual per channel to enforce YoY bounds (pathwise)
    const prevAnnual = {}; // c -> last year's annual in this trial

    // Iterate years in ascending order so FY26 precedes FY27, etc.
    for (const y of years) {

      // Draw the two common shocks for each quarter and hold them for reuse across channels
      const EG = {}; // q -> {E, G}
      for (const q of QUARTERS) EG[q] = { E: randn(rng), G: randn(rng) };

      // Generate raw quarterly values (pre-bounds) for this year
      const cqValue = {}; // key c|q => value
      for (const q of QUARTERS) {
        const E = EG[q].E, G = EG[q].G;
        for (const c of channels) {
          const w = wByC[c] || { wE: 0, wG: 0, wR: 1 };
          const eps = randn(rng);                              // idiosyncratic piece per channel-quarter
          const Z = w.wE * E + w.wG * G + w.wR * eps;          // standardized shock with Var=1
          const sp = spec[ck(c, y, q)];
          const v = (!sp || sp.degenerate) ? 0 : Math.exp(sp.mu + sp.sigma * Z); // lognormal draw or deterministic 0
          cqValue[ck(c, q)] = v; // one value per (c,q) for this year
        }
      }

      // Apply annual YoY bounds by proportionally scaling a channel's four quarters if needed
      const isFirstYear = (y === years[0]);
      let totalAnnual = 0;

      for (const c of channels) {
        // Pull the four quarter values for this channel-year
        const qVals = d.QUARTERS.map(q => cqValue[ck(c, q)] || 0);

        // Aggregate to an annual figure BEFORE bounds
        let annual = qVals.reduce((a, b) => a + b, 0);

        // Enforce bounds for FY27+ if present
        const b = bounds[ck(c, y)];
        if (!isFirstYear && (b && (b.min !== null || b.max !== null))) {
          const prev = prevAnnual[c] || 0;
          // Compute level bounds relative to last year's simulated annual
          let minBound = (b.min !== null && isFinite(b.min)) ? prev * (1 + b.min) : -Infinity;
          let maxBound = (b.max !== null && isFinite(b.max)) ? prev * (1 + b.max) : +Infinity;

          // Special rule: if prior year is zero, allow recovery by IGNORING upper bound; lower bound is at least 0
          if (prev === 0) {
            minBound = Math.max(minBound, 0);
            maxBound = +Infinity;
          }

          // If annual is below the lower bound, raise it by scaling all four quarters equally
          if (annual < minBound && isFinite(minBound) && minBound > -Infinity) {
            if (annual <= 0) {
              // Avoid divide-by-zero: if we need positive annual but drew ~0, split the minimum evenly across quarters
              if (minBound > 0) {
                const eq = minBound / 4;
                for (let i = 0; i < d.QUARTERS.length; i++) qVals[i] = eq;
                annual = minBound;
              }
            } else {
              const scale = minBound / annual;
              for (let i = 0; i < d.QUARTERS.length; i++) qVals[i] *= scale;
              annual = minBound;
            }
          }

          // If annual is above the upper bound, reduce it by scaling all four quarters equally
          if (annual > maxBound && isFinite(maxBound) && maxBound < +Infinity) {
            if (annual > 0) {
              const scale = maxBound / annual;
              for (let i = 0; i < d.QUARTERS.length; i++) qVals[i] *= scale;
              annual = maxBound;
            }
            // If annual==0 and maxBound>=0, nothing to do
          }
        }

        // Update prevAnnual for the next year's bounding and add to total
        prevAnnual[c] = annual;
        totalAnnual += annual;

        // Update ANNUAL (channel) aggregators with post-bounds value
        const AA = aAgg[ck(c, y)];
        AA.mean.push(annual);
        for (const p2 of AA.p2s) p2.push(annual);

        // Plan hit-rate (channel): count successes against planByYear
        const plan = (planByYear[y] && planByYear[y][c]) || 0;
        if (annual >= plan) AA.hits++;
        if (annual >= 0.9 * plan) AA.hits10++;
        AA.n++;

        // PI aggregators (channel): feed both lower/upper percentiles
        const PI = piAgg[ck(c, y)];
        PI.low.push(annual);
        PI.up.push(annual);

        // Update QUARTERLY aggregators after scaling (the values visible in Results_Quarterly reflect post-rules numbers)
        for (let i = 0; i < d.QUARTERS.length; i++) {
          const qk = ck(c, d.QUARTERS[i], undefined); // not used; we key by (c,y,q) below
        }
        for (let i = 0; i < d.QUARTERS.length; i++) {
          const q = d.QUARTERS[i];
          const v = qVals[i];
          const QA = qAgg[ck(c, y, q)];
          QA.mean.push(v);
          for (const p2 of QA.p2s) p2.push(v);
        }
      }

      // Update TOTAL (annual) aggregators for this year
      const TA = tAgg[y];
      TA.mean.push(totalAnnual);
      for (const p2 of TA.p2s) p2.push(totalAnnual);

      // Plan hit-rate (total)
      const tPlan = (planByYear[y] && planByYear[y].Total) || 0;
      if (totalAnnual >= tPlan) TA.hits++;
      if (totalAnnual >= 0.9 * tPlan) TA.hits10++;
      TA.n++;

      // PI aggregators (total)
      piTot[y].low.push(totalAnnual);
      piTot[y].up.push(totalAnnual);
    } // end year loop
  } // end trials loop

  // Write all outputs (apply display rounding to $1k if requested)
  writeQuarterly(ss, qAgg, d, p.RoundK);
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
 * - Keeps 5 markers (positions & heights) and updates them on each push
 * - Very small memory footprint and fast; great for Apps Script limits
 * - Accuracy is excellent for the typical N (10k–50k)
 */
function P2Quantile(p) {
  this.p = p;
  this.initial = []; // we need 5 seeds to initialize the markers
  this.n = [0, 0, 0, 0, 0];  // marker positions
  this.q = [0, 0, 0, 0, 0];  // marker heights
  this.np = [0, 0, 0, 0, 0]; // desired positions
}
P2Quantile.prototype.push = function (x) {
  const p = this.p;
  // Collect first 5 points exactly, then seed markers
  if (this.initial.length < 5) {
    this.initial.push(x);
    if (this.initial.length === 5) {
      this.initial.sort(function (a, b) { return a - b; });
      this.q = this.initial.slice();
      this.n = [1, 2, 3, 4, 5];
      this.np = [1, 1 + 2 * p, 1 + 4 * p, 3 + 2 * p, 5];
    }
    return;
  }
  // Determine k: which marker interval x falls into
  let k;
  if (x < this.q[0]) { this.q[0] = x; k = 0; }
  else if (x < this.q[1]) { k = 0; }
  else if (x < this.q[2]) { k = 1; }
  else if (x < this.q[3]) { k = 2; }
  else if (x <= this.q[4]) { k = 3; }
  else { this.q[4] = x; k = 3; }

  // Increment positions for markers above k
  for (let i = k + 1; i < 5; i++) this.n[i]++;

  // Update desired positions based on p
  for (let i = 0; i < 5; i++) this.np[i] += [0, p / 2, p, (1 + p) / 2, 1][i];

  // Adjust interior markers using parabolic interpolation when possible
  for (let i = 1; i <= 3; i++) {
    const d = this.np[i] - this.n[i];
    if ((d >= 1 && this.n[i + 1] - this.n[i] > 1) || (d <= -1 && this.n[i - 1] - this.n[i] < -1)) {
      const dsgn = Math.sign(d);
      const qn = this.parabolic(i, dsgn);
      if (this.q[i - 1] < qn && qn < this.q[i + 1]) this.q[i] = qn;
      else this.q[i] = this.linear(i, dsgn);
      this.n[i] += dsgn;
    }
  }
};
P2Quantile.prototype.parabolic = function (i, d) {
  const q = this.q, n = this.n;
  const a = (d * (n[i] - n[i - 1] + d) * (q[i + 1] - q[i])) / (n[i + 1] - n[i]);
  const b = (d * (n[i + 1] - n[i] - d) * (q[i] - q[i - 1])) / (n[i] - n[i - 1]);
  return q[i] + (a + b) / (n[i + 1] - n[i - 1]);
};
P2Quantile.prototype.linear = function (i, d) {
  return this.q[i] + d * (this.q[i + d] - this.q[i]) / (this.n[i + d] - this.n[i]);
};
P2Quantile.prototype.value = function () {
  if (this.initial.length && this.initial.length < 5) {
    // For the very first few pushes, return an exact quantile from the tiny set
    const arr = this.initial.slice().sort(function (a, b) { return a - b; });
    const idx = Math.max(0, Math.min(arr.length - 1, Math.floor(this.p * (arr.length - 1))));
    return arr[idx];
  }
  // After initialization, the target quantile is the middle marker
  return this.q[2];
};

/** Factory: builds an aggregator with running mean + a set of P² trackers + optional hit-rate counters. */
function makeAgg(ps, withHits) {
  return {
    mean: new RunningMean(),
    p2s: ps.map(q => new P2Quantile(q)),
    qs: ps,
    hits: withHits ? 0 : undefined,     // plan hit-rate (>= 100% of plan)
    hits10: withHits ? 0 : undefined,   // within-10% hit-rate (>= 90% of plan)
    n: withHits ? 0 : undefined
  };
}

/** Factory: builds a PI aggregator holding two P² trackers for custom lower/upper percentiles. */
function makePI(pLow, pUp) {
  return { low: new P2Quantile(pLow), up: new P2Quantile(pUp), pLow, pUp };
}

/* ============================== WRITERS ============================= */
/** Writes the Results_Quarterly sheet: Channel | Fiscal_Year | Quarter | Mean | P10 | P25 | P50 | P75 | P90 */
function writeQuarterly(ss, qAgg, d, doRound) {
  const sh = ss.getSheetByName('Results_Quarterly');
  if (!sh) return;

  // Clear only data rows (keep headers)
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, lastCol).clearContent();

  // Sort keys by Channel, Year, Quarter
  const keys = Object.keys(qAgg).sort(function (a, b) {
    const [ca, ya, qa] = a.split('|'); const [cb, yb, qb] = b.split('|');
    const ccmp = ca.localeCompare(cb); if (ccmp) return ccmp;
    const ycmp = Number(ya) - Number(yb); if (ycmp) return ycmp;
    const order = { Q1: 1, Q2: 2, Q3: 3, Q4: 4 };
    return (order[qa] || 0) - (order[qb] || 0);
  });

  const out = [];
  for (const k of keys) {
    const [c, y, q] = k.split('|');
    const A = qAgg[k];
    out.push([
      c,
      Number(y),
      q,
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

/**
 * Writes Results_Annually with the correct column order:
 * Fiscal_Year | Channel | Mean | P10 | P25 | P50 | P75 | P90 | Plan | Plan_Likelihood | Plan_Likelihood_within10%
 * This also populates Plan and Total_Plan from Params_Plans (not left blank).
 */
function writeAnnually(ss, aAgg, tAgg, d, doRound, percentiles) {
  const sh = ss.getSheetByName('Results_Annually');
  if (!sh) return;

  // Format percentile names (e.g., 0.05 -> "P05", 0.95 -> "P95")
  const pNames = percentiles.map(p => 'P' + String(Math.round(p * 100)).padStart(2, '0'));
  
  // Define expected headers
  const expectedHeaders = ['Fiscal_Year', 'Channel', 'Mean'].concat(pNames).concat(['Plan', 'Plan_Likelihood', 'Plan_Likelihood_within10%']);
  
  // Update header row
  sh.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);

  // Clear only data rows (keep headers)
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow > 1 && lastCol > 0) sh.getRange(2, 1, lastRow - 1, lastCol).clearContent();

  const channelRows = [];
  // Sort by Channel FIRST, then Year
  const keys = Object.keys(aAgg).sort(function(a, b) {
    const [ca, ya] = a.split('|'); const [cb, yb] = b.split('|');
    const ccmp = ca.localeCompare(cb); if (ccmp) return ccmp;  // Channel first
    return Number(ya) - Number(yb);  // Then year
  });

  for (const k of keys) {
    const [c, ys] = k.split('|');
    const y = Number(ys);
    const A = aAgg[k];

    // Pull this channel's plan for the year
    const planYear = d.planByYear[y] || {};
    const plan = (planYear && Object.prototype.hasOwnProperty.call(planYear, c)) ? planYear[c] : '';

    channelRows.push([
      y,                                 // Fiscal_Year
      c,                                 // Channel
      roundK(A.mean.value(), doRound),   // Mean
      roundK(A.p2s[0].value(), doRound), // Dynamic lower percentile
      roundK(A.p2s[1].value(), doRound), // P25
      roundK(A.p2s[2].value(), doRound), // P50
      roundK(A.p2s[3].value(), doRound), // P75
      roundK(A.p2s[4].value(), doRound), // Dynamic upper percentile
      plan,                              // Plan
      (A.n && A.n > 0) ? (A.hits / A.n) : 0,   // Plan_Likelihood
      (A.n && A.n > 0) ? (A.hits10 / A.n) : 0  // Plan_Likelihood_within10%
    ]);
  }

  const totalRows = [];
  const tYears = Object.keys(tAgg).map(Number).sort((a, b) => a - b);
  for (const y of tYears) {
    const A = tAgg[y];
    const planYear = d.planByYear[y] || {};
    const totalPlan = (planYear && Object.prototype.hasOwnProperty.call(planYear, 'Total')) ? planYear['Total'] : '';

    totalRows.push([
      y,
      'Total',
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

/** Writes Results_PI: Fiscal_Year | Channel | PI_Lower_Value | PI_Upper_Value, including a Total row per year. */
function writePI(ss, piAgg, piTot, d, doRound) {
  const sh = ss.getSheetByName('Results_PI');
  if (!sh) return;

  // Clear only data rows (keep headers)
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
      y,
      c,
      roundK(A.low.value(), doRound),
      roundK(A.up.value(), doRound)
    ]);
  }

  const tYears = Object.keys(piTot).map(Number).sort((a, b) => a - b);
  for (const y of tYears) {
    const A = piTot[y];
    rows.push([
      y,
      'Total',
      roundK(A.low.value(), doRound),
      roundK(A.up.value(), doRound)
    ]);
  }

  if (rows.length) sh.getRange(2, 1, rows.length, 4).setValues(rows);
}

/* =============================== LOGGER ============================= */
/**
 * Simple append-only logger:
 * - Writes to (or creates) a "Log" sheet
 * - Each call adds one row: ISO timestamp | key | [LEVEL] message
 * - Keep messages short; this is operational telemetry, not a novel
 */
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
    flush: () => {} // placeholder if you later buffer logs and want to flush once
  };
}

/* ============================== UTILITIES =========================== */
/** Compact key builder: c|y or c|y|q depending on args. */
function ck(c, y, q) {
  if (q !== undefined) return `${c}|${y}|${q}`;
  if (y !== undefined) return `${c}|${y}`;
  return String(c);
}

/** Normalizes a channel label to one of the expected names where possible. */
function cleanChannel(v) {
  const s = String(v || '').trim();
  if (!s) return null;
  const x = s.toLowerCase();
  if (x.startsWith('corp')) return 'Corporate';
  if (x.startsWith('gov')) return 'Government';
  if (x.startsWith('maj')) return 'Major';
  if (x.startsWith('market')) return 'Marketplace';
  if (['Corporate','Government','Major','Marketplace'].indexOf(s) >= 0) return s;
  // If it's something else, return as-is; it just may not be picked up later
  return s;
}

/** A ready-to-use set of neutral multipliers (all 1.0). */
function oneSet() { return { Headwind: 1, Neutral: 1, Tailwind: 1 }; }

/** Numeric helpers */
function numOrZero(v) { return (isFinite(Number(v)) ? Number(v) : 0); }
function safeNum(v, dflt) { const n = Number(v); return isFinite(n) ? n : dflt; }

/** Display rounding: floor to nearest $1,000 only if enabled (never in-core). */
function roundK(x, on) { return on ? (Math.floor(x / 1000) * 1000) : x; }

/** Deterministic, fast PRNG (Mulberry32) returning U[0,1). */
function mulberry32(a) {
  return function () {
    a |= 0; a = a + 0x6D2B79F5 | 0;
    let t = Math.imul(a ^ a >>> 15, 1 | a);
    t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t;
    return ((t ^ t >>> 14) >>> 0) / 4294967296;
  }
}

/** Standard normal via Box–Muller transform using the PRNG above. */
function randn(rng) {
  let u = 0, v = 0;
  while (u === 0) u = rng(); // avoid log(0)
  while (v === 0) v = rng();
  return Math.sqrt(-2.0 * Math.log(u)) * Math.cos(2.0 * Math.PI * v);
}
