# UTM & URL Shortener Builder (Google Sheets Extension)

A powerful, native Google Sheets Add-on designed for media buyers and marketers to generate UTM-tagged URLs in bulk, encode dynamic ad network parameters, and seamlessly generate short links. 

It runs entirely within Google Apps Script using a dark "Lone Wolf" aesthetic sidebar UI, preventing the need for external software or expensive URL shortener subscriptions.

---

## ⚡ Core Features

### 1. Bulk UTM Generation
- **Active Selection Processing:** Simply highlight any range of cells containing your base URLs directly in the spreadsheet before clicking generate. No more guessing column letters.
- Standard inputs for all UTM parameters (`utm_source`, `utm_medium`, `utm_campaign`, `utm_term`, `utm_content`).
- **Performance Optimized:** Uses an asynchronous, chunk-based processing model (handling 50 rows at a time) to completely bypass Google’s strict execution time limits. It allows generation of thousands of URLs without crashing.

### 2. Advanced Dynamic URL Tracking (Ad Networks)
- Comes fully equipped with a built-in helper for **Google Ads (ValueTrack)** and **Meta Ads (Facebook/Instagram)** tracking variables.
- One-click token insertion (e.g., `{campaignid}`, `{{adset.name}}`) directly into UTM fields.
- **Smart Encoding Logic:** The backend automatically ensures that curly brackets `{ }` and `{{ }}` remain un-encoded in the final URL outputs, making it strictly compliant with how Meta and Google servers inject live tracking data.

### 3. Waterfall URL Shortening
Includes an automated, fallback-enabled URL shortener engine that generates custom short links for every created UTM.
- **API Fallback Sequence:** Attempts short generation sequentially using multiple open APIs to guarantee reliability: `CleanURI` &rarr; `Ulvis` &rarr; `ManyAPIs` &rarr; `Is.gd`.
- **Intelligent Caching:** Uses Google Apps Script's `CacheService` to cache already generated short links for 6 hours. This drastically reduces API strain and speeds up script execution if you happen to run the same base URLs multiple times.

### 4. Custom Dark Mode UI
- Built with a sleek, custom "Lone Wolf" UI theme using deep dark tones, crisp squared borders, and sharp teal highlights. 
- Integrated real-time progress bar UI tracking chunked processing completions.

---

## 🛠 Project Structure & Logic

*   **`appsscript.json`:** The Apps Script manifest file. Demands the required OAuth scopes to edit spreadsheets, ping external APIs (for URL shortening), and display custom UI sidebars inside Google Sheets.
*   **`Code.gs`:** The backend Apps Script logic. Handles triggering the sidebar menu (`onOpen()`) and executing the hefty data-manipulation task (`generateUrlsChunked()`). Hosts all server-side API calls for the waterfall shortener network.
*   **`Sidebar.html`:** The HTML/CSS/JS frontend driving the Google Sheets Sidebar tab. Features the complete form structure, the Advanced Tracking token logic, dynamic DOM updates for the loading bar, and connection to the backend via `google.script.run`.

---

## 🚀 Setup & Installation

1. Open your target Google Sheet.
2. Go to **Extensions > Apps Script**.
3. Copy the contents of `Code.gs`, `Sidebar.html`, and `appsscript.json` (making sure to reveal the manifest via the gear icon if hidden) into the script editor.
4. Save and restart your Google Sheet.
5. Grant the necessary permissions on the first run. 
6. Access the tool anywhere by selecting **UTM and URL Shortener Builder > Open Sidebar** in the top menu.
