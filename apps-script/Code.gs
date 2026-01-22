function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('AI Scouting')
    .addItem('Update accelerators', 'updateAccelerators')
    .addItem('Update startups', 'updateStartups')
    .addSeparator()
    .addItem('Normalize URLs', 'normalizeAllUrls')
    .addToUi();
}

function getISOTimestamp() {
  return new Date().toISOString();
}

function log(message, data = '') {
  console.log(`[AI-SCOUTING] ${message}`, data);
}

function normalizeUrl(url) {
  if (!url) return '';

  return url
    .toLowerCase()
    .trim()
    .replace(/^https?:\/\//, '')
    .replace(/^www\./, '')
    .replace(/\/$/, '');
}

function getExistingUrls(sheetName, columnIndex = 1) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  const values = sheet.getRange(2, columnIndex, sheet.getLastRow()).getValues();

  const urlSet = new Set();
  values.forEach(row => {
    const normalized = normalizeUrl(row[0]);
    if (normalized) urlSet.add(normalized);
  });

  return urlSet;
}

function normalizeAllUrls() {
  const ss = SpreadsheetApp.getActive();
  const sheets = ['accelerators', 'startups'];

  sheets.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const range = sheet.getRange(2, 1, sheet.getLastRow());
    const values = range.getValues();

    const updated = values.map(row => {
      row[0] = normalizeUrl(row[0]);
      return row;
    });

    range.setValues(updated);
  });

  SpreadsheetApp.getUi().alert('All URLs normalized');
}

const SEED_ACCELERATORS = [
  {
    name: 'Seedcamp',
    website: 'https://seedcamp.com',
    country: 'UK',
    source_url: 'https://seedcamp.com'
  },
  {
    name: 'Antler',
    website: 'https://antler.co',
    country: 'Global',
    source_url: 'https://antler.co'
  },
  {
    name: 'Entrepreneur First',
    website: 'https://joinef.com',
    country: 'UK',
    source_url: 'https://joinef.com'
  },
  {
    name: 'Station F',
    website: 'https://stationf.co',
    country: 'France',
    source_url: 'https://stationf.co'
  }
];

function updateAccelerators() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('accelerators');
  const existingUrls = getExistingUrls('accelerators');

  let inserted = 0;

  SEED_ACCELERATORS.forEach(acc => {
    const normalizedWebsite = normalizeUrl(acc.website);

    if (existingUrls.has(normalizedWebsite)) {
      log('Skipping existing accelerator', normalizedWebsite);
      return;
    }

    sheet.appendRow([
      normalizedWebsite,
      acc.name,
      acc.country,
      acc.source_url,
      getISOTimestamp()
    ]);

    existingUrls.add(normalizedWebsite);
    inserted++;
  });

  SpreadsheetApp.getUi().alert(`${inserted} accelerators added`);
}

const PORTFOLIO_PATHS = [
  '/portfolio',
  '/companies',
  '/startups',
  '/alumni'
];

function extractLinks(html, baseDomain) {
  if (!html) return [];

  const links = [];
  const regex = /href="(https?:\/\/[^"]+)"/g;
  let match;

  while ((match = regex.exec(html)) !== null) {
    const url = match[1];
    const normalized = normalizeUrl(url);

    
    if (normalized && !normalized.includes(baseDomain)) {
      links.push(normalized);
    }
  }

  return [...new Set(links)];
}

function updateStartups() {
  const ss = SpreadsheetApp.getActive();
  const accSheet = ss.getSheetByName('accelerators');
  const stSheet = ss.getSheetByName('startups');

  const existingStartups = getExistingUrls('startups');
  const accData = accSheet.getRange(2, 1, accSheet.getLastRow() - 1, 5).getValues();

  let inserted = 0;

  accData.forEach(row => {
    const accWebsite = row[0];
    const baseDomain = accWebsite;

    PORTFOLIO_PATHS.forEach(path => {
      const url = `https://${baseDomain}${path}`;
      const html = safeFetch(url);
      if (!html) return;

      const links = extractLinks(html, baseDomain);

      links.forEach(link => {
        if (existingStartups.has(link)) return;

        stSheet.appendRow([
          link,                
          '',                  
          '',                  
          baseDomain,           
          url,                  
          getISOTimestamp()     
        ]);

        existingStartups.add(link);
        inserted++;
      });
    });
  });

  SpreadsheetApp.getUi().alert(`${inserted} startups discovered`);
}

function extractVisibleText(html) {
  if (!html) return '';

  return html
    .replace(/<script[\s\S]*?>[\s\S]*?<\/script>/gi, '')
    .replace(/<style[\s\S]*?>[\s\S]*?<\/style>/gi, '')
    .replace(/<[^>]*>/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 3000);
}

function generateValueProposition(text) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) return 'API key not configured';

  const payload = {
    model: 'gpt-4o-mini',
    messages: [
      {
        role: 'system',
        content: 'You generate conservative, factual startup value propositions.'
      },
      {
        role: 'user',
        content: `
Using ONLY the text below, write ONE sentence (max 25 words):

Rules:
- Do not invent features
- If unclear, be generic
- No buzzwords

Text:
${text}
        `
      }
    ],
    temperature: 0.2
  };

  const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
    method: 'post',
    headers: {
      Authorization: `Bearer ${apiKey}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const data = JSON.parse(response.getContentText());
  return data?.choices?.[0]?.message?.content || 'No value proposition generated';
}

function generateValueProps() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('startups');
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();

  let updated = 0;

  data.forEach((row, i) => {
    const website = row[0];
    const valueProp = row[5];

    if (valueProp) return; 

    const html = safeFetch(`https://${website}`);
    if (!html) return;

    const text = extractVisibleText(html);
    if (!text) return;

    const vp = generateValueProposition(text);

    sheet.getRange(i + 2, 7).setValue(vp); // value_proposition
    sheet.getRange(i + 2, 8).setValue(getISOTimestamp());
    sheet.getRange(i + 2, 9).setValue('gpt-4o-mini');

    updated++;
  });

  SpreadsheetApp.getUi().alert(`${updated} value propositions generated`);
}
