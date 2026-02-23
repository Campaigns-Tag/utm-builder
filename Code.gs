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

function generateUrls(utmSource, utmMedium, utmCampaign) {
  if (!utmSource || !utmMedium || !utmCampaign) {
    return { status: 'error', message: 'Please fill all three UTM fields.' };
  }
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { status: 'error', message: 'No base URLs found in column A.' };
  }

  const baseVals = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const utmUrls = [], shortUrls = [];
  const props = PropertiesService.getScriptProperties();

  function md5Hex(s) {
    return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, s)
      .map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
  }
  function fromCache(key) { return props.getProperty(key); }
  function toCache(key, val) { try { props.setProperty(key, val); } catch (e) {} }

  function tryCleanURI(longUrl) {
    const key = 'c_' + md5Hex(longUrl);
    const cached = fromCache(key);
    if (cached) return cached;
    try {
      const resp = UrlFetchApp.fetch('https://cleanuri.com/api/v1/shorten', {
        method: 'POST', contentType: 'application/json',
        payload: JSON.stringify({ url: longUrl }),
        muteHttpExceptions: true
      });
      if (resp.getResponseCode() === 200) {
        const d = JSON.parse(resp.getContentText());
        if (d.result_url) { toCache(key, d.result_url); return d.result_url; }
      }
    } catch (e) {}
    return null;
  }

  function tryUlvis(longUrl) {
    const key = 'u_' + md5Hex(longUrl);
    const cached = fromCache(key);
    if (cached) return cached;
    try {
      const resp = UrlFetchApp.fetch('https://ulvis.net/api/v1/shorten', {
        method: 'POST', contentType: 'application/json',
        payload: JSON.stringify({ url: longUrl }),
        muteHttpExceptions: true
      });
      if (resp.getResponseCode() === 200) {
        const d = JSON.parse(resp.getContentText());
        if (d.shortUrl) { toCache(key, d.shortUrl); return d.shortUrl; }
      }
    } catch (e) {}
    return null;
  }

  function tryManyApis(longUrl) {
    const key = 'm_' + md5Hex(longUrl);
    const cached = fromCache(key);
    if (cached) return cached;
    try {
      const resp = UrlFetchApp.fetch('https://manyapis.com/api/shorten-url', {
        method: 'POST', contentType: 'application/json',
        payload: JSON.stringify({ url: longUrl }),
        muteHttpExceptions: true
      });
      if (resp.getResponseCode() === 200) {
        const d = JSON.parse(resp.getContentText());
        if (d.shortened_url) { toCache(key, d.shortened_url); return d.shortened_url; }
      }
    } catch (e) {}
    return null;
  }

  function tryIsGd(longUrl) {
    const key = 'g_' + md5Hex(longUrl);
    const cached = fromCache(key);
    if (cached) return cached;
    try {
      const resp = UrlFetchApp.fetch('https://is.gd/create.php?format=simple&url=' + encodeURIComponent(longUrl), { muteHttpExceptions: true });
      const text = resp.getContentText().trim();
      if (resp.getResponseCode() === 200 && text && !text.startsWith('Error:')) {
        toCache(key, text);
        return text;
      }
    } catch (e) {}
    return null;
  }

  baseVals.forEach((row, idx) => {
    const raw = (row[0] || '').toString().trim();
    if (!raw) {
      utmUrls.push(['']); shortUrls.push(['']); return;
    }
    const sep = raw.includes('?') ? '&' : '?';
    const utm = `${raw}${sep}utm_source=${encodeURIComponent(utmSource)}&utm_medium=${encodeURIComponent(utmMedium)}&utm_campaign=${encodeURIComponent(utmCampaign)}`;
    utmUrls.push([utm]);

    let short = tryCleanURI(utm);
    if (!short) { Utilities.sleep(500); short = tryUlvis(utm); }
    if (!short) { Utilities.sleep(500); short = tryManyApis(utm); }
    if (!short) { Utilities.sleep(500); short = tryIsGd(utm); }
    shortUrls.push([short || 'SHORT_ERR']);
  });

  sheet.getRange(2, 2, utmUrls.length, 1).setValues(utmUrls);
  sheet.getRange(2, 3, shortUrls.length, 1).setValues(shortUrls);

  return { status: 'success', message: 'Generation complete (watch for SHORT_ERR where all providers failed).' };
}
