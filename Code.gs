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
  // Simple regex to check if it's a valid http or https URL
  return /^https?:\/\//i.test(string);
}

function generateUrlsChunked(options, startRow) {
  const startTime = Date.now();
  const { source, medium, campaign, term, content, mode } = options;
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let rangeInfo = options.rangeInfo;
  
  if (!rangeInfo) {
    const activeRange = sheet.getActiveRange();
    if (!activeRange) {
      return { status: 'error', message: 'Please select a range of URLs first.' };
    }
    rangeInfo = {
      startRow: activeRange.getRow(),
      endRow: activeRange.getLastRow(),
      startCol: activeRange.getColumn()
    };
    if (!startRow || startRow < rangeInfo.startRow) {
       startRow = rangeInfo.startRow;
    }
  }

  const lastRow = rangeInfo.endRow;
  
  if (startRow > lastRow) {
    return { status: 'success', message: 'Generation complete!' };
  }

  // Calculate the end row for this chunk
  let endRow = startRow + CHUNK_SIZE - 1;
  if (endRow > lastRow) endRow = lastRow;
  
  const numRowsToProcess = endRow - startRow + 1;
  const baseVals = sheet.getRange(startRow, rangeInfo.startCol, numRowsToProcess, 1).getValues();
  
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
       if (mode === 'both' && utmUrls.length > 0) {
         sheet.getRange(startRow, rangeInfo.startCol + 1, utmUrls.length, 1).setValues(utmUrls);
         sheet.getRange(startRow, rangeInfo.startCol + 2, shortUrls.length, 1).setValues(shortUrls);
       } else if (mode === 'utm' && utmUrls.length > 0) {
         sheet.getRange(startRow, rangeInfo.startCol + 1, utmUrls.length, 1).setValues(utmUrls);
       } else if (mode === 'short' && shortUrls.length > 0) {
         sheet.getRange(startRow, rangeInfo.startCol + 1, shortUrls.length, 1).setValues(shortUrls);
       }
       options.rangeInfo = rangeInfo;
       // Return partial status with the row we stopped at
       return { 
         status: 'partial', 
         nextRow: startRow + i, 
         currentRow: startRow + i - 1, 
         options: options
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
    let params = [];
    
    // URL-encode standard strings, but allow { and } to remain raw later
    if (source) params.push('utm_source=' + encodeURIComponent(source));
    if (medium) params.push('utm_medium=' + encodeURIComponent(medium));
    if (campaign) params.push('utm_campaign=' + encodeURIComponent(campaign));
    if (term) params.push('utm_term=' + encodeURIComponent(term));
    if (content) params.push('utm_content=' + encodeURIComponent(content));
    
    if (params.length > 0) {
      // Append either with ? or &, depending on if the URL already has parameters
      const separator = utm.indexOf('?') !== -1 ? '&' : '?';
      utm += separator + params.join('&');
    }

    // Now handle template vars for specific platforms
    if (options.templateVars) {
      let tParams = [];
      for (const [key, val] of Object.entries(options.templateVars)) {
        if (val) {
          tParams.push(`${key}=${encodeURIComponent(val)}`);
        }
      }
      if (tParams.length > 0) {
        const separator = utm.indexOf('?') !== -1 ? '&' : '?';
        utm += separator + tParams.join('&');
      }
    }
    
    // Decode the URL-encoded braces back to raw braces for ad network compatibility (Google/Meta dynamic tracking)
    utm = utm.replace(/%7B/gi, '{').replace(/%7D/gi, '}');
    
    if (mode === 'both' || mode === 'utm') {
      utmUrls.push([utm]);
    }

    if (mode === 'both' || mode === 'short') {
      let short = tryCleanURI(utm);
      if (!short) { Utilities.sleep(500); short = tryUlvis(utm); }
      if (!short) { Utilities.sleep(500); short = tryManyApis(utm); }
      if (!short) { Utilities.sleep(500); short = tryIsGd(utm); }
      shortUrls.push([short || 'SHORT_ERR']);
    }
  }

  // Write this chunk's results to the sheet
  if (mode === 'both' && utmUrls.length > 0) {
    sheet.getRange(startRow, rangeInfo.startCol + 1, utmUrls.length, 1).setValues(utmUrls);
    sheet.getRange(startRow, rangeInfo.startCol + 2, shortUrls.length, 1).setValues(shortUrls);
  } else if (mode === 'utm' && utmUrls.length > 0) {
    sheet.getRange(startRow, rangeInfo.startCol + 1, utmUrls.length, 1).setValues(utmUrls);
  } else if (mode === 'short' && shortUrls.length > 0) {
    sheet.getRange(startRow, rangeInfo.startCol + 1, shortUrls.length, 1).setValues(shortUrls);
  }

  // Check if we are done
  if (endRow >= lastRow) {
    return { status: 'success', message: 'Generation complete!' };
  }

  // Otherwise, signal that more chunks are needed
  options.rangeInfo = rangeInfo;
  return { 
    status: 'partial', 
    nextRow: endRow + 1, 
    currentRow: endRow, 
    options: options
  };
}
