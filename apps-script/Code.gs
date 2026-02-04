const CONFIG = {
  MENU_NAME: 'Startup Scouting AI',
  SHEETS: {
    ACC: 'accelerators',
    ST: 'startups',
    LOG: 'logs'
  },
  
  PORTFOLIO_PATHS: ['/portfolio', '/companies', '/startups', '/alumni', '/portfolio-companies'],
  
  BLOCKLIST_DOMAINS: [
    'twitter.com', 'x.com', 'linkedin.com', 'facebook.com', 'instagram.com', 'youtube.com',
    'tiktok.com', 'medium.com', 'wikipedia.org', 'crunchbase.com', 'angel.co', 'wellfound.com'
  ],
  FETCH_TIMEOUT_MS: 25000,
  MAX_HTML_CHARS: 250000,
  TEXT_MAX_CHARS: 3500,
  GEMINI_DEFAULT_MODEL: 'gemini-1.5-flash',
  GEMINI_TEMPERATURE: 0.2
};


const SEED_ACCELERATORS = [
  { name: 'Seedcamp', website: 'https://seedcamp.com', country: 'UK', source_url: 'https://seedcamp.com' },
  { name: 'Antler', website: 'https://antler.co', country: 'Global', source_url: 'https://antler.co' },
  { name: 'Entrepreneur First', website: 'https://joinef.com', country: 'UK', source_url: 'https://joinef.com' },
  { name: 'Station F', website: 'https://stationf.co', country: 'France', source_url: 'https://stationf.co' }
];


function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu(CONFIG.MENU_NAME)
    .addItem('1) Scouting accelerators (seed batch)', 'scoutAccelerators')
    .addItem('2) Update startups from accelerators', 'updateStartupsFromAccelerators')
    .addItem('3) Generate missing value_proposition', 'generateMissingValuePropositions')
    .addSeparator()
    .addItem('Normalize URLs (all tabs)', 'normalizeAllUrls')
    .addToUi();

  ensureSchema_();
}


function ensureSchema_() {
  const ss = SpreadsheetApp.getActive();

  // accelerators headers
  ensureHeaders_(ss, CONFIG.SHEETS.ACC, [
    'website', 'name', 'country', 'source_url', 'created_at'
  ]);

  // startups headers (minimum + extras for traceability)
  ensureHeaders_(ss, CONFIG.SHEETS.ST, [
    'website', 'name', 'country', 'accelerator', 'source_url', 'discovered_at', 'value_proposition', 'vp_generated_at', 'vp_model'
  ]);

  // logs sheet
  ensureHeaders_(ss, CONFIG.SHEETS.LOG, [
    'ts', 'level', 'action', 'message', 'context'
  ]);
}

function ensureHeaders_(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);

  const firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const isEmpty = firstRow.every(c => !c);

  if (isEmpty) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
}

function isoNow_() {
  return new Date().toISOString();
}

/** Log to both console and logs sheet (skip + log philosophy) */
function log_(level, action, message, contextObj) {
  const context = contextObj ? JSON.stringify(contextObj).slice(0, 2000) : '';
  console.log(`[AI-SCOUTING][${level}][${action}] ${message} ${context}`);

  try {
    const ss = SpreadsheetApp.getActive();
    const logSheet = ss.getSheetByName(CONFIG.SHEETS.LOG);
    logSheet.appendRow([isoNow_(), level, action, message, context]);
  } catch (e) {
    console.log('Logging-to-sheet failed:', e);
  }
}

/**
 * Normalize website for dedup:
 * - keep only hostname (no path/query/hash)
 * - lowercase
 * - remove leading www.
 * Works with "https://x.com/a", "x.com", "www.x.com"
 */
function normalizeWebsite_(input) {
  if (!input) return '';
  const s = String(input).trim();

  // If user stored already normalized (like seedcamp.com), ensure parse works:
  const candidate = s.match(/^https?:\/\//i) ? s : `https://${s}`;

  try {
    const u = new URL(candidate);
    return u.hostname.toLowerCase().replace(/^www\./, '');
  } catch (e) {
    // fallback: strip protocol, www, path
    return s.toLowerCase()
      .replace(/^https?:\/\//, '')
      .replace(/^www\./, '')
      .split('/')[0]
      .split('?')[0]
      .split('#')[0]
      .trim();
  }
}

function normalizeAllUrls() {
  const ss = SpreadsheetApp.getActive();
  [CONFIG.SHEETS.ACC, CONFIG.SHEETS.ST].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;

    const lastRow = sh.getLastRow();
    if (lastRow < 2) return;

    const range = sh.getRange(2, 1, lastRow - 1, 1);
    const values = range.getValues().map(r => [normalizeWebsite_(r[0])]);
    range.setValues(values);
  });

  SpreadsheetApp.getUi().alert('URL normalization done (hostname-based).');
}

/** Read existing websites (dedup set) from a sheet column A (website) */
function getExistingWebsiteSet_(sheetName) {
  const sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  const lastRow = sh.getLastRow();
  const set = new Set();

  if (lastRow < 2) return set;

  const values = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  values.forEach(r => {
    const n = normalizeWebsite_(r[0]);
    if (n) set.add(n);
  });
  return set;
}

/* =========================
 * 1) SCOUT ACCELERATORS
 * ========================= */
function scoutAccelerators() {
  const action = 'scoutAccelerators';
  ensureSchema_();

  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEETS.ACC);
  const existing = getExistingWebsiteSet_(CONFIG.SHEETS.ACC);

  let inserted = 0;
  SEED_ACCELERATORS.forEach(acc => {
    try {
      const website = normalizeWebsite_(acc.website);
      if (!website) return;

      if (existing.has(website)) {
        log_('INFO', action, 'Skipping existing accelerator', { website });
        return;
      }

      sh.appendRow([website, acc.name || '', acc.country || '', acc.source_url || '', isoNow_()]);
      existing.add(website);
      inserted++;
    } catch (e) {
      log_('ERROR', action, 'Failed inserting accelerator', { err: String(e), acc });
    }
  });

  SpreadsheetApp.getUi().alert(`${inserted} accelerators added (seed batch).`);
}

/* =========================
 * 2) UPDATE STARTUPS
 * ========================= */

/** Robust fetch: returns html string or null (skip + log) */
function safeFetchHtml_(url, action, context) {
  try {
    const res = UrlFetchApp.fetch(url, {
      method: 'get',
      followRedirects: true,
      muteHttpExceptions: true,
      validateHttpsCertificates: true,
      timeout: CONFIG.FETCH_TIMEOUT_MS
    });

    const code = res.getResponseCode();
    if (code < 200 || code >= 400) {
      log_('WARN', action, 'Fetch non-OK response', { url, code, context });
      return null;
    }

    let html = res.getContentText() || '';
    if (!html) return null;

    if (html.length > CONFIG.MAX_HTML_CHARS) {
      html = html.slice(0, CONFIG.MAX_HTML_CHARS);
    }

    return html;
  } catch (e) {
    log_('WARN', action, 'Fetch failed', { url, err: String(e), context });
    return null;
  }
}

function extractExternalLinks_(html, acceleratorHost) {
  if (!html) return [];

  // capture href="..." or href='...'
  const regex = /href\s*=\s*["']([^"']+)["']/gi;
  const out = new Set();

  let match;
  while ((match = regex.exec(html)) !== null) {
    let href = match[1];
    if (!href) continue;

    // ignore anchors, mailto, tel, js
    if (href.startsWith('#') || href.startsWith('mailto:') || href.startsWith('tel:') || href.startsWith('javascript:')) continue;

    // keep only http(s)
    if (href.startsWith('//')) href = 'https:' + href;
    if (!href.match(/^https?:\/\//i)) continue;

    const host = normalizeWebsite_(href);

    // ignore same domain (accelerator)
    if (!host || host === acceleratorHost) continue;

    // ignore blocklisted domains
    if (CONFIG.BLOCKLIST_DOMAINS.some(d => host.endsWith(d))) continue;

    out.add(host);
  }

  return [...out];
}

function updateStartupsFromAccelerators() {
  const action = 'updateStartupsFromAccelerators';
  ensureSchema_();

  const ss = SpreadsheetApp.getActive();
  const accSh = ss.getSheetByName(CONFIG.SHEETS.ACC);
  const stSh = ss.getSheetByName(CONFIG.SHEETS.ST);

  const accLast = accSh.getLastRow();
  if (accLast < 2) {
    SpreadsheetApp.getUi().alert('No accelerators found. Run “Scouting accelerators” first.');
    return;
  }

  const existingStartups = getExistingWebsiteSet_(CONFIG.SHEETS.ST);

  // accelerators columns: website, name, country, source_url, created_at
  const accRows = accSh.getRange(2, 1, accLast - 1, 5).getValues();

  let inserted = 0;

  accRows.forEach(row => {
    const accWebsite = normalizeWebsite_(row[0]);
    const accName = row[1] || '';
    const accCountry = row[2] || '';
    const accSource = row[3] || '';

    if (!accWebsite) return;

    CONFIG.PORTFOLIO_PATHS.forEach(path => {
      const portfolioUrl = `https://${accWebsite}${path}`;
      const html = safeFetchHtml_(portfolioUrl, action, { accWebsite, path });
      if (!html) return;

      const startupHosts = extractExternalLinks_(html, accWebsite);

      startupHosts.forEach(stHost => {
        if (!stHost) return;
        if (existingStartups.has(stHost)) return;

        try {
          // startups headers:
          // website, name, country, accelerator, source_url, discovered_at, value_proposition, vp_generated_at, vp_model
          stSh.appendRow([
            stHost,
            '',              // name unknown for now
            '',              // country unknown for now
            accWebsite,      // accelerator reference (website-based)
            portfolioUrl,    // source_url where discovered
            isoNow_(),
            '', '', ''       // VP fields empty
          ]);

          existingStartups.add(stHost);
          inserted++;
        } catch (e) {
          log_('ERROR', action, 'Failed inserting startup', { err: String(e), stHost, accWebsite, portfolioUrl });
        }
      });

      // be polite with websites (avoid hammering)
      Utilities.sleep(250);
    });
  });

  SpreadsheetApp.getUi().alert(`${inserted} startups discovered from accelerators.`);
}

/* =========================
 * 3) GENERATE VALUE PROPOSITIONS (Gemini)
 * ========================= */

function extractVisibleText_(html) {
  if (!html) return '';

  // remove scripts/styles
  let text = html
    .replace(/<script[\s\S]*?>[\s\S]*?<\/script>/gi, ' ')
    .replace(/<style[\s\S]*?>[\s\S]*?<\/style>/gi, ' ')
    .replace(/<noscript[\s\S]*?>[\s\S]*?<\/noscript>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  if (text.length > CONFIG.TEXT_MAX_CHARS) {
    text = text.slice(0, CONFIG.TEXT_MAX_CHARS);
  }
  return text;
}

function getGeminiConfig_() {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('GEMINI_API_KEY');
  const model = props.getProperty('GEMINI_MODEL') || CONFIG.GEMINI_DEFAULT_MODEL;

  if (!apiKey) {
    throw new Error('GEMINI_API_KEY is not configured in Script Properties.');
  }
  return { apiKey, model };
}

/**
 * Gemini generateContent call
 * Docs-style endpoint:
 * https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key=API_KEY
 */
function geminiGenerateOneSentence_(startupHost, pageText) {
  const { apiKey, model } = getGeminiConfig_();

  const prompt = [
    `You are generating a startup value proposition.`,
    `Using ONLY the provided text, write EXACTLY ONE sentence in English with this schema:`,
    `Startup <X> helps <Target Y> do <What W> so that <Benefit Z>.`,
    ``,
    `Rules:`,
    `- Max 25 words`,
    `- Do NOT invent features`,
    `- If unclear, keep it generic and conservative`,
    `- No hype/buzzwords`,
    `- If the startup name is unclear, use "${startupHost}" as <X>`,
    ``,
    `TEXT:`,
    pageText
  ].join('\n');

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent?key=${encodeURIComponent(apiKey)}`;

  const payload = {
    contents: [
      {
        role: 'user',
        parts: [{ text: prompt }]
      }
    ],
    generationConfig: {
      temperature: CONFIG.GEMINI_TEMPERATURE
    }
  };

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  const body = res.getContentText();

  if (code < 200 || code >= 300) {
    throw new Error(`Gemini error ${code}: ${body}`);
  }

  const data = JSON.parse(body);
  const out = data?.candidates?.[0]?.content?.parts?.[0]?.text || '';

  // defensive cleanup
  return String(out).replace(/\s+/g, ' ').trim();
}

function generateMissingValuePropositions() {
  const action = 'generateMissingValuePropositions';
  ensureSchema_();

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CONFIG.SHEETS.ST);

  const lastRow = sh.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('No startups found. Run “Update startups from accelerators” first.');
    return;
  }

  // columns:
  // A website
  // G value_proposition
  // H vp_generated_at
  // I vp_model
  const data = sh.getRange(2, 1, lastRow - 1, 9).getValues();

  let updated = 0;

  data.forEach((row, idx) => {
    const rowNumber = idx + 2;
    const stHost = normalizeWebsite_(row[0]);
    const currentVP = row[6]; // value_proposition (col G)

    if (!stHost) return;
    if (currentVP && String(currentVP).trim().length > 0) return; // already done

    try {
      const html = safeFetchHtml_(`https://${stHost}`, action, { stHost });
      if (!html) {
        log_('WARN', action, 'No HTML for startup site (skipping)', { stHost });
        return;
      }

      const text = extractVisibleText_(html);
      if (!text || text.length < 80) {
        log_('WARN', action, 'Text too short/empty (skipping)', { stHost, textLen: text.length });
        return;
      }

      const vp = geminiGenerateOneSentence_(stHost, text);

      if (!vp) {
        log_('WARN', action, 'Gemini returned empty VP (skipping)', { stHost });
        return;
      }

      // write results
      sh.getRange(rowNumber, 7).setValue(vp);        // G value_proposition
      sh.getRange(rowNumber, 8).setValue(isoNow_()); // H vp_generated_at
      sh.getRange(rowNumber, 9).setValue(getGeminiConfig_().model); // I vp_model

      updated++;
      Utilities.sleep(300); // pacing for API
    } catch (e) {
      log_('ERROR', action, 'VP generation failed (skip + log)', { stHost, err: String(e) });
      // do not stop the whole process
    }
  });

  SpreadsheetApp.getUi().alert(`${updated} value propositions generated (missing only).`);
}


