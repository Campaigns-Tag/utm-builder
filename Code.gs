// @OnlyCurrentDoc
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('UTM and URL Shortener Builder')
    .addItem('Open Sidebar', 'showSidebar')
    .addToUi();
}
function showSidebar() {
  SpreadsheetApp.getUi()
    .showSidebar(HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('UTM URL Generator'));
}

// Global configurations for batching
const MAX_EXECUTION_TIME_MS = 4 * 60 * 1000; // 4 minutes to be safe (limit is 6)
const CHUNK_SIZE = 50;

function isValidUrl(string) {
  try {
    new URL(string);
    return true;
  } catch (_) {
    return false;  
  }
}

function generateUrlsChunked(options, startRow) {
  const startTime = Date.now();
  const { source, medium, campaign, term, content, colIdx } = options;
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 2) {
    return { status: 'error', message: 'No data found in the sheet.' };
  }

  // Ensure startRow doesn't exceed lastRow
  if (startRow > lastRow) {
    return { status: 'success', message: 'Generation complete!' };
  }

  // Calculate the end row for this chunk
  let endRow = startRow + CHUNK_SIZE - 1;
  if (endRow > lastRow) endRow = lastRow;
  
  const numRowsToProcess = endRow - startRow + 1;
  const baseVals = sheet.getRange(startRow, colIdx, numRowsToProcess, 1).getValues();
  
  const utmUrls = [];
  const shortUrls = [];
  
  // Use CacheService for ephemeral caching of shortened URLs
  const cache = CacheService.getScriptCache();

  function md5Hex(s) {
    return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, s)
      .map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
  }

  function tryCleanURI(longUrl) {
    const key = 'c_' + md5Hex(longUrl);
    const cached = cache.get(key);
    if (cached) return cached;
    try {
      const resp = UrlFetchApp.fetch('https://cleanuri.com/api/v1/shorten', {
        method: 'POST', contentType: 'application/json',
        payload: JSON.stringify({ url: longUrl }),
        muteHttpExceptions: true
      });
      if (resp.getResponseCode() === 200) {
        const d = JSON.parse(resp.getContentText());
        if (d.result_url) { cache.put(key, d.result_url, 21600); return d.result_url; } // 6 hours
      }
    } catch (e) {}
    return null;
  }

  function tryUlvis(longUrl) {
    const key = 'u_' + md5Hex(longUrl);
    const cached = cache.get(key);
    if (cached) return cached;
    try {
      const resp = UrlFetchApp.fetch('https://ulvis.net/api/v1/shorten', {
        method: 'POST', contentType: 'application/json',
        payload: JSON.stringify({ url: longUrl }),
        muteHttpExceptions: true
      });
      if (resp.getResponseCode() === 200) {
        const d = JSON.parse(resp.getContentText());
        if (d.shortUrl) { cache.put(key, d.shortUrl, 21600); return d.shortUrl; }
      }
    } catch (e) {}
    return null;
  }

  function tryManyApis(longUrl) {
    const key = 'm_' + md5Hex(longUrl);
    const cached = cache.get(key);
    if (cached) return cached;
    try {
      const resp = UrlFetchApp.fetch('https://manyapis.com/api/shorten-url', {
        method: 'POST', contentType: 'application/json',
        payload: JSON.stringify({ url: longUrl }),
        muteHttpExceptions: true
      });
      if (resp.getResponseCode() === 200) {
        const d = JSON.parse(resp.getContentText());
        if (d.shortened_url) { cache.put(key, d.shortened_url, 21600); return d.shortened_url; }
      }
    } catch (e) {}
    return null;
  }

  function tryIsGd(longUrl) {
    const key = 'g_' + md5Hex(longUrl);
    const cached = cache.get(key);
    if (cached) return cached;
    try {
      const resp = UrlFetchApp.fetch('https://is.gd/create.php?format=simple&url=' + encodeURIComponent(longUrl), { muteHttpExceptions: true });
      const text = resp.getContentText().trim();
      if (resp.getResponseCode() === 200 && text && !text.startsWith('Error:')) {
        cache.put(key, text, 21600);
        return text;
      }
    } catch (e) {}
    return null;
  }

  for (let i = 0; i < baseVals.length; i++) {
    // Basic timeout check within the chunk loop just in case URL fetching is very slow
    if (Date.now() - startTime > MAX_EXECUTION_TIME_MS) {
       // Save what we have so far
       if (utmUrls.length > 0) {
         sheet.getRange(startRow, colIdx + 1, utmUrls.length, 1).setValues(utmUrls);
         sheet.getRange(startRow, colIdx + 2, shortUrls.length, 1).setValues(shortUrls);
       }
       // Return partial status with the row we stopped at
       return { 
         status: 'partial', 
         nextRow: startRow + i, 
         currentRow: startRow + i - 1, 
         totalRows: lastRow 
       };
    }

    const raw = (baseVals[i][0] || '').toString().trim();
    
    // Skip empty cells or invalid URLs
    if (!raw || !isValidUrl(raw)) {
      utmUrls.push(['']); 
      shortUrls.push(['']); 
      continue;
    }

    let utm = raw;
    const urlObj = new URL(raw);
    
    urlObj.searchParams.set('utm_source', source);
    urlObj.searchParams.set('utm_medium', medium);
    urlObj.searchParams.set('utm_campaign', campaign);
    
    if (term) urlObj.searchParams.set('utm_term', term);
    if (content) urlObj.searchParams.set('utm_content', content);
    
    utm = urlObj.toString();
    
    // Decode the URL-encoded braces back to raw braces for ad network compatibility (Google/Meta dynamic tracking)
    utm = utm.replace(/%7B/g, '{').replace(/%7D/g, '}');
    
    utmUrls.push([utm]);

    let short = tryCleanURI(utm);
    if (!short) { Utilities.sleep(500); short = tryUlvis(utm); }
    if (!short) { Utilities.sleep(500); short = tryManyApis(utm); }
    if (!short) { Utilities.sleep(500); short = tryIsGd(utm); }
    shortUrls.push([short || 'SHORT_ERR']);
  }

  // Write this chunk's results to the sheet
  if (utmUrls.length > 0) {
    sheet.getRange(startRow, colIdx + 1, utmUrls.length, 1).setValues(utmUrls);
    sheet.getRange(startRow, colIdx + 2, shortUrls.length, 1).setValues(shortUrls);
  }

  // Check if we are done
  if (endRow >= lastRow) {
    return { status: 'success', message: 'Generation complete!' };
  }

  // Otherwise, signal that more chunks are needed
  return { 
    status: 'partial', 
    nextRow: endRow + 1, 
    currentRow: endRow, 
    totalRows: lastRow 
  };
}
