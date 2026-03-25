/**
 * Saks Filed Claims Scraper + Excel Generator (v3 — fully autonomous)
 *
 * Usage: Open https://cases.stretto.com/Saks/filed-claims/ in Chrome,
 *        then paste this entire script into the DevTools Console (F12).
 *        It runs fully autonomously — no re-pasting needed.
 *
 * When the WAF token expires, it auto-refreshes via a hidden iframe
 * and continues scraping without any manual intervention.
 */
(async function SaksScraper() {
  'use strict';

  // ═══════════════════════ CASE CONFIG (change these per case) ═══════════════════════
  const CASE_NAME     = 'iPic';
  const CASE_URL      = 'https://cases.stretto.com/iPicTheaters/filed-claims/';
  const EXPECTED_TOTAL = 500;   // rough estimate for progress bar — update if known

  // ═══════════════════════ CONFIG ═══════════════════════
  const BATCH_SIZE = 2;
  const BATCH_DELAY_MS = 2000;
  const REST_EVERY_N = 50;
  const REST_DURATION_MS = 30000;     // 30s rest every 50 claims
  const PAGE_WAIT_MS = 2500;
  const MAX_RETRIES = 2;
  const _now = new Date();
  const _dt = _now.getFullYear().toString() +
    String(_now.getMonth() + 1).padStart(2, '0') +
    String(_now.getDate()).padStart(2, '0') + '_' +
    String(_now.getHours()).padStart(2, '0') +
    String(_now.getMinutes()).padStart(2, '0') +
    String(_now.getSeconds()).padStart(2, '0');
  const XLSX_FILENAME = `CR_${CASE_NAME}_stretto-parser_${_dt}.xlsx`;
  const JSON_FILENAME = `CR_${CASE_NAME}_stretto-parser_${_dt}.json`;
  const STORAGE_KEY = `${CASE_NAME.toUpperCase()}_SCRAPER_STATE`;
  const REFRESH_COOLDOWN_MS = 60000;  // wait 60s after WAF refresh before resuming

  // ═══════════════════════ STATE (localStorage-backed) ═══════════════════════

  function loadState() {
    try {
      const saved = localStorage.getItem(STORAGE_KEY);
      if (saved) {
        const parsed = JSON.parse(saved);
        // Use console.log directly here — log() depends on S which isn't initialized yet
        console.log('%c[Scraper] Restored from localStorage: ' +
          parsed.basicData.length + ' basic, ' +
          Object.keys(parsed.childData).length + ' child rows',
          'color: #FF9800; font-weight: bold');
        return parsed;
      }
    } catch (e) {
      console.error('[Scraper] Failed to restore state:', e);
    }
    return null;
  }

  function saveState() {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(S));
    } catch (e) {
      log('Warning: could not save state: ' + e.message);
    }
  }

  function clearState() {
    localStorage.removeItem(STORAGE_KEY);
  }

  const restored = loadState();
  if (!window.SCRAPER && restored) {
    window.SCRAPER = restored;
    window.SCRAPER.startTime = Date.now();
  } else if (!window.SCRAPER) {
    window.SCRAPER = {
      basicData: [],
      childData: {},
      errors: [],
      phase: 'init',
      startTime: Date.now()
    };
  }
  const S = window.SCRAPER;

  // ═══════════════════════ LOGGING ═══════════════════════

  function log(msg) {
    const elapsed = ((Date.now() - S.startTime) / 1000).toFixed(0);
    console.log(`%c[Scraper ${elapsed}s] ${msg}`, 'color: #2196F3; font-weight: bold');
  }

  function logProgress(current, total, label) {
    const pct = ((current / total) * 100).toFixed(1);
    const bar = '\u2588'.repeat(Math.floor(current / total * 30)) +
                '\u2591'.repeat(30 - Math.floor(current / total * 30));
    console.log(`%c[${bar}] ${pct}% (${current}/${total}) ${label}`, 'color: #4CAF50');
  }

  // ═══════════════════════ UTILITIES ═══════════════════════

  async function waitForAjax(timeout = 15000) {
    return new Promise(resolve => {
      const start = Date.now();
      const check = setInterval(() => {
        if (jQuery.active === 0 || Date.now() - start > timeout) {
          clearInterval(check);
          setTimeout(resolve, 500);
        }
      }, 100);
    });
  }

  async function sleep(ms) {
    return new Promise(r => setTimeout(r, ms));
  }

  function parseCurrency(str) {
    if (!str || typeof str !== 'string') return null;
    const cleaned = str.replace(/[$,]/g, '').trim();
    if (!cleaned) return null;
    const num = parseFloat(cleaned);
    return isNaN(num) ? null : num;
  }

  // ═══════════════════════ WAF TOKEN REFRESH (hidden iframe) ═══════════════════════

  async function refreshWafToken() {
    log('WAF token expired. Auto-refreshing via hidden iframe...');
    log('(This takes ~30-60 seconds — the iframe loads the page to get fresh cookies)');
    saveState();

    return new Promise((resolve, reject) => {
      // Remove any previous iframe
      const old = document.getElementById('__waf_refresh_iframe');
      if (old) old.remove();

      const iframe = document.createElement('iframe');
      iframe.id = '__waf_refresh_iframe';
      iframe.style.cssText = 'position:fixed;top:-9999px;left:-9999px;width:1px;height:1px;opacity:0;';
      iframe.src = CASE_URL;

      let resolved = false;
      const timeout = setTimeout(() => {
        if (!resolved) {
          resolved = true;
          iframe.remove();
          reject(new Error('Iframe WAF refresh timed out after 90s'));
        }
      }, 90000);

      iframe.onload = () => {
        // The page loaded — WAF cookies are now refreshed.
        // Also extract new signed URLs from the iframe.
        try {
          const iframeWindow = iframe.contentWindow;
          if (iframeWindow.claimsLinks) {
            // Update our signed URLs with the fresh ones
            window.claimsLinks = { ...iframeWindow.claimsLinks };
            log('Refreshed signed URLs from iframe');
          }
        } catch (e) {
          log('Could not extract URLs from iframe (cross-origin?), but cookies should be refreshed');
        }

        clearTimeout(timeout);
        iframe.remove();
        if (!resolved) {
          resolved = true;
          resolve();
        }
      };

      iframe.onerror = () => {
        clearTimeout(timeout);
        iframe.remove();
        if (!resolved) {
          resolved = true;
          reject(new Error('Iframe failed to load'));
        }
      };

      document.body.appendChild(iframe);
    });
  }

  // (No external libraries needed — Excel generation uses native HTML table format)

  // ═══════════════════════ PHASE A: PAGINATION ═══════════════════════

  function readCurrentPageData() {
    const rows = document.querySelectorAll(
      '#claim-filed-and-scheduled-table tbody tr:not(.child-row)');
    return [...rows].map(row => {
      const cells = row.querySelectorAll('td');
      if (cells.length < 6) return null;
      const claimIdInput = row.querySelector('input[name="c_id"]');
      return {
        claimId: claimIdInput ? claimIdInput.value : '',
        creditorName: cells[0] ? cells[0].textContent.trim() : '',
        scheduleNo: cells[1] ? cells[1].textContent.trim() : '',
        claimNo: cells[2] ? cells[2].textContent.trim() : '',
        dateFiled: cells[3] ? cells[3].textContent.trim() : '',
        currentAmount: cells[4] ? cells[4].textContent.trim() : '',
        debtorName: cells[5] ? cells[5].textContent.trim() : ''
      };
    }).filter(r => r && r.claimId);
  }

  async function collectBasicData() {
    S.phase = 'pagination';
    log('Phase A: Collecting basic claim data via pagination...');

    const firstPageBtn = document.querySelector('.paginate_button[data-dt-idx="1"]');
    if (firstPageBtn && !firstPageBtn.classList.contains('current')) {
      firstPageBtn.click();
      await waitForAjax();
    }

    document.querySelectorAll('tr.child-row').forEach(r => r.remove());

    let pageNum = 0;
    while (true) {
      pageNum++;
      const pageData = readCurrentPageData();

      const existingIds = new Set(S.basicData.map(d => d.claimId));
      const newData = pageData.filter(d => !existingIds.has(d.claimId));
      S.basicData.push(...newData);

      logProgress(S.basicData.length, EXPECTED_TOTAL, `Page ${pageNum}`);

      const nextBtn = document.querySelector('.paginate_button.next');
      if (!nextBtn) break;
      if (window.getComputedStyle(nextBtn).visibility === 'hidden') break;
      if (pageData.length === 0) break;
      if (pageNum >= 100) break;

      nextBtn.click();
      await waitForAjax();
      await sleep(PAGE_WAIT_MS);
    }

    saveState();
    log(`Phase A complete: ${S.basicData.length} claims across ${pageNum} pages`);
  }

  // ═══════════════════════ PHASE B: CHILD ROWS ═══════════════════════

  let consecutive403 = 0;

  async function fetchChildRow(claimId) {
    for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
      try {
        const token = await runRecaptcha('claimsChild');
        const html = await jQuery.ajax({
          url: claimsLinks.claimsChildRow + '&claim_id=' + claimId,
          type: 'post',
          data: { 'g-recaptcha-response': token }
        });

        if (typeof html === 'object' && html.success === false) {
          throw new Error('API returned success:false');
        }
        if (typeof html === 'string' && html.length > 100) {
          consecutive403 = 0;
          return html;
        }
        return typeof html === 'string' ? html : '';

      } catch (err) {
        const is403 = err && (err.status === 403 ||
          String(err.statusText || '').toLowerCase().includes('forbidden'));

        if (is403) {
          consecutive403++;
          if (consecutive403 >= 2) {
            throw new Error('WAF_EXPIRED');
          }
          await sleep(5000);
        }

        if (attempt === MAX_RETRIES) throw err;
        const errMsg = is403 ? '403 Forbidden (WAF)'
          : (err.message || err.statusText || String(err));
        log(`Retry ${attempt}/${MAX_RETRIES} for claim ${claimId}: ${errMsg}`);
        await sleep(2000 * attempt);
      }
    }
  }

  function parseChildRowHTML(html) {
    const temp = document.createElement('div');
    temp.innerHTML = html;
    temp.querySelectorAll('script, style').forEach(el => el.remove());

    // ── Creditor Address ──
    const creditorBox = temp.querySelector('.claim-creditor-details .creditor-detail-box');
    let creditorAddress = '';
    if (creditorBox) {
      creditorAddress = creditorBox.innerHTML
        .replace(/<br\s*\/?>/gi, '\n')
        .replace(/<[^>]+>/g, '')
        .split('\n').map(l => l.trim()).filter(l => l).join(', ');
    }

    // ── Amount Breakdowns ──
    function parseAmountCard(card) {
      if (!card) return {
        generalUnsecured: '', priority: '', secured: '', adminPriority: '', total: ''
      };
      const titles = [...card.querySelectorAll('.creditor-inner-table-title')]
        .map(el => el.textContent.trim());
      const values = [...card.querySelectorAll('.creditor-inner-table-value')]
        .map(el => el.textContent.trim());
      const r = { generalUnsecured: '', priority: '', secured: '', adminPriority: '', total: '' };
      titles.forEach((t, i) => {
        const v = values[i] || '';
        if (t === 'General Unsecured') r.generalUnsecured = v;
        else if (t === 'Priority') r.priority = v;
        else if (t === 'Secured') r.secured = v;
        else if (t === 'Admin Priority') r.adminPriority = v;
        else if (t === 'TOTAL') r.total = v;
      });
      return r;
    }

    const cardBoxes = [...temp.querySelectorAll('.card-box')];
    let scheduleCard = null, filedCard = null, currentCard = null;
    cardBoxes.forEach(card => {
      const text = card.textContent;
      if (text.includes('Schedule Amount') && !scheduleCard) scheduleCard = card;
      else if (text.includes('Filed Claim Amount') && !filedCard) filedCard = card;
      else if (text.includes('Current Claim Amount') && !currentCard) currentCard = card;
    });

    // ── Claim Status ──
    let claimStatus = '';
    const allTitles = [...temp.querySelectorAll('.creditor-inner-table-title')];
    const allValues = [...temp.querySelectorAll('.creditor-inner-table-value')];
    allTitles.forEach((el, i) => {
      if (el.textContent.trim() === 'Claim Status')
        claimStatus = allValues[i] ? allValues[i].textContent.trim() : '';
    });

    // ── PDF Download Link ──
    const dlLink = [...temp.querySelectorAll('a')].find(a =>
      a.textContent.trim() === 'Download' || (a.getAttribute('href') || '').includes('pdf'));
    const pdfUrl = dlLink ? dlLink.getAttribute('href') || '' : '';

    // ── Notice Parties ──
    let noticeParties = '';
    for (const div of temp.querySelectorAll('div')) {
      const h = div.querySelector('h3, h4');
      if (h && h.textContent.includes('Notice Parties')) {
        noticeParties = div.textContent.replace('Notice Parties', '').trim() || 'None';
        break;
      }
    }

    // ── History Sections ──
    temp.querySelectorAll('.hide').forEach(el => el.classList.remove('hide'));
    function parseHist(name) {
      const div = [...temp.querySelectorAll('.claim-history')].find(d => d.textContent.includes(name));
      if (!div) return '';
      const tables = div.querySelectorAll('table');
      if (!tables.length) return '';
      const results = [];
      tables.forEach(tbl => {
        const headers = [...tbl.querySelectorAll('th')].map(th => th.textContent.trim());
        [...tbl.querySelectorAll('tr')].filter(tr => tr.querySelector('td')).forEach(tr => {
          const cells = [...tr.querySelectorAll('td')].map(td => td.textContent.trim());
          const entry = {};
          headers.forEach((h, i) => { entry[h] = cells[i] || ''; });
          results.push(entry);
        });
      });
      return results.length ? JSON.stringify(results) : '';
    }

    return {
      creditorAddress,
      scheduleAmount: parseAmountCard(scheduleCard),
      filedClaimAmount: parseAmountCard(filedCard),
      currentClaimAmount: parseAmountCard(currentCard),
      claimStatus, pdfUrl, noticeParties,
      objectionHistory: parseHist('Claim Objection History'),
      transferHistory: parseHist('Claim Transfer History'),
      withdrawalHistory: parseHist('Claim Withdrawal History'),
      stipulationHistory: parseHist('Stipulation History')
    };
  }

  async function collectChildData() {
    S.phase = 'childRows';

    while (true) {
      const remaining = S.basicData.map(d => d.claimId).filter(id => !S.childData[id]);
      const total = S.basicData.length;
      const done = total - remaining.length;

      if (remaining.length === 0) {
        log(`All ${total} child rows collected!`);
        return;
      }

      log(`Phase B: ${remaining.length} child rows remaining (${done} done)...`);
      let fetchedThisRound = 0;
      let wafExpired = false;

      for (let i = 0; i < remaining.length; i += BATCH_SIZE) {
        const batch = remaining.slice(i, i + BATCH_SIZE);

        const results = await Promise.allSettled(batch.map(async (claimId) => {
          const html = await fetchChildRow(claimId);
          S.childData[claimId] = parseChildRowHTML(html);
          fetchedThisRound++;
        }));

        // Check for WAF expiry
        for (const r of results) {
          if (r.status === 'rejected') {
            if (r.reason && r.reason.message === 'WAF_EXPIRED') {
              wafExpired = true;
            } else {
              const claimId = batch[results.indexOf(r)];
              S.errors.push({ claimId, error: r.reason?.message || String(r.reason) });
            }
          }
        }

        if (wafExpired) break;

        const nowDone = done + fetchedThisRound;
        if (nowDone % 10 === 0 || nowDone >= total) {
          logProgress(nowDone, total, 'child rows');
        }
        if (fetchedThisRound % 10 === 0) saveState();

        // ── Rest break ──
        if (fetchedThisRound > 0 && fetchedThisRound % REST_EVERY_N === 0) {
          saveState();
          log(`Rest break (30s) after ${fetchedThisRound} fetches...`);
          await sleep(REST_DURATION_MS);
          log('Resuming...');
          consecutive403 = 0; // reset after rest
        }

        if (i + BATCH_SIZE < remaining.length) await sleep(BATCH_DELAY_MS);
      }

      saveState();

      if (wafExpired) {
        // ── Auto-refresh WAF token via iframe ──
        try {
          await refreshWafToken();
          consecutive403 = 0;
          log(`WAF refreshed. Cooling down ${REFRESH_COOLDOWN_MS / 1000}s before resuming...`);
          await sleep(REFRESH_COOLDOWN_MS);
          log('Resuming scraping with fresh WAF token...');
          // Loop continues — goes back to top of while(true)
        } catch (refreshErr) {
          log('Iframe WAF refresh failed: ' + refreshErr.message);
          log('Waiting 2 minutes then trying again...');
          await sleep(120000);
          // Loop continues — will retry
        }
      } else {
        // No WAF issue — we're either done or had other errors
        break;
      }
    }

    const finalDone = Object.keys(S.childData).length;
    log(`Phase B complete: ${finalDone}/${S.basicData.length} child rows, ${S.errors.length} errors`);
  }

  // ═══════════════════════ BUILD COMBINED DATA ═══════════════════════

  function buildCombinedData() {
    return S.basicData.map(basic => {
      const c = S.childData[basic.claimId] || {};
      return {
        claimNo: basic.claimNo,
        creditorName: basic.creditorName,
        creditorAddress: c.creditorAddress || '',
        debtorName: basic.debtorName,
        dateFiled: basic.dateFiled,
        scheduleNo: basic.scheduleNo,
        claimStatus: c.claimStatus || '',
        currentAmountTotal: basic.currentAmount,
        currentGeneralUnsecured: c.currentClaimAmount?.generalUnsecured || '',
        currentPriority: c.currentClaimAmount?.priority || '',
        currentSecured: c.currentClaimAmount?.secured || '',
        currentAdminPriority: c.currentClaimAmount?.adminPriority || '',
        filedAmountTotal: c.filedClaimAmount?.total || '',
        filedGeneralUnsecured: c.filedClaimAmount?.generalUnsecured || '',
        filedPriority: c.filedClaimAmount?.priority || '',
        filedSecured: c.filedClaimAmount?.secured || '',
        filedAdminPriority: c.filedClaimAmount?.adminPriority || '',
        scheduleAmountTotal: c.scheduleAmount?.total || '',
        scheduleGeneralUnsecured: c.scheduleAmount?.generalUnsecured || '',
        schedulePriority: c.scheduleAmount?.priority || '',
        scheduleSecured: c.scheduleAmount?.secured || '',
        scheduleAdminPriority: c.scheduleAmount?.adminPriority || '',
        proofOfClaim: c.pdfUrl || '',
        noticeParties: c.noticeParties || '',
        objectionHistory: c.objectionHistory || '',
        transferHistory: c.transferHistory || '',
        withdrawalHistory: c.withdrawalHistory || '',
        stipulationHistory: c.stipulationHistory || ''
      };
    });
  }

  // ═══════════════════════ PHASE C: GENERATE EXCEL ═══════════════════════

  const COL_DEFS = [
    { key: 'claimNo',                header: 'Claim No.',                 width: 12,  type: 'number' },
    { key: 'creditorName',           header: 'Creditor Name',            width: 35,  type: 'text' },
    { key: 'creditorAddress',        header: 'Creditor Address',         width: 45,  type: 'text' },
    { key: 'debtorName',             header: 'Debtor Name',              width: 30,  type: 'text' },
    { key: 'dateFiled',              header: 'Date Filed',               width: 14,  type: 'date' },
    { key: 'claimStatus',            header: 'Claim Status',             width: 16,  type: 'text' },
    { key: 'scheduleNo',             header: 'Schedule No.',             width: 14,  type: 'text' },
    { key: 'currentAmountTotal',     header: 'Current Amount (Total)',   width: 22,  type: 'currency' },
    { key: 'currentGeneralUnsecured',header: 'Current - Gen. Unsecured', width: 24, type: 'currency' },
    { key: 'currentPriority',        header: 'Current - Priority',       width: 18,  type: 'currency' },
    { key: 'currentSecured',         header: 'Current - Secured',        width: 18,  type: 'currency' },
    { key: 'currentAdminPriority',   header: 'Current - Admin Priority', width: 22,  type: 'currency' },
    { key: 'filedAmountTotal',       header: 'Filed Amount (Total)',     width: 22,  type: 'currency' },
    { key: 'filedGeneralUnsecured',  header: 'Filed - Gen. Unsecured',   width: 24, type: 'currency' },
    { key: 'filedPriority',          header: 'Filed - Priority',         width: 18,  type: 'currency' },
    { key: 'filedSecured',           header: 'Filed - Secured',          width: 18,  type: 'currency' },
    { key: 'filedAdminPriority',     header: 'Filed - Admin Priority',   width: 22,  type: 'currency' },
    { key: 'scheduleAmountTotal',    header: 'Schedule Amount (Total)',  width: 22,  type: 'currency' },
    { key: 'scheduleGeneralUnsecured',header:'Schedule - Gen. Unsecured',width: 24, type: 'currency' },
    { key: 'schedulePriority',       header: 'Schedule - Priority',      width: 18,  type: 'currency' },
    { key: 'scheduleSecured',        header: 'Schedule - Secured',       width: 18,  type: 'currency' },
    { key: 'scheduleAdminPriority',  header: 'Schedule - Admin Priority',width: 22,  type: 'currency' },
    { key: 'proofOfClaim',           header: 'Proof of Claim (PDF)',     width: 30,  type: 'link' },
    { key: 'noticeParties',          header: 'Notice Parties',           width: 30,  type: 'text' },
    { key: 'objectionHistory',       header: 'Objection History',        width: 30,  type: 'text' },
    { key: 'transferHistory',        header: 'Transfer History',         width: 30,  type: 'text' },
    { key: 'withdrawalHistory',      header: 'Withdrawal History',       width: 30,  type: 'text' },
    { key: 'stipulationHistory',     header: 'Stipulation History',      width: 30,  type: 'text' },
  ];

  function formatHistory(val) {
    if (!val) return '';
    try {
      const entries = JSON.parse(val);
      if (!Array.isArray(entries) || !entries.length) return '';
      return entries.map(e =>
        Object.entries(e).filter(([,v]) => v).map(([k,v]) => `${k}: ${v}`).join('; ')
      ).join('\n');
    } catch { return String(val); }
  }

  async function generateExcel(data) {
    S.phase = 'excel';
    log('Phase C: Generating Excel file (no external libraries needed)...');

    data.sort((a, b) => (parseInt(a.claimNo) || 99999) - (parseInt(b.claimNo) || 99999));

    function esc(val) {
      if (val == null) return '';
      return String(val).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    }

    // Build Excel-compatible HTML with Microsoft Office XML namespaces
    let html = `<html xmlns:o="urn:schemas-microsoft-com:office:office"
      xmlns:x="urn:schemas-microsoft-com:office:excel"
      xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta charset="utf-8">
<!--[if gte mso 9]><xml>
<x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet>
<x:Name>Filed Claims</x:Name>
<x:WorksheetOptions>
<x:FreezePanes/><x:FrozenNoSplit/><x:SplitHorizontal>1</x:SplitHorizontal>
<x:TopRowBottomPane>1</x:TopRowBottomPane>
<x:ActivePane>2</x:ActivePane>
<x:Selected/>
</x:WorksheetOptions>
<x:AutoFilter x:Range="A1:${String.fromCharCode(64 + COL_DEFS.length)}1"/>
</x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook>
</xml><![endif]-->
<style>
  table { border-collapse: collapse; }
  th {
    background: #1F4E79; color: #FFFFFF; font-family: Calibri; font-size: 11pt;
    font-weight: bold; text-align: center; padding: 6px 8px;
    border-bottom: 2px solid #000000;
  }
  td {
    font-family: Calibri; font-size: 10pt; padding: 4px 6px;
    border-bottom: 1px solid #E0E0E0; vertical-align: top;
  }
  tr.even td { background: #F2F6FA; }
  .num { text-align: center; }
  .cur { text-align: right; mso-number-format: "#\\,##0\\.00"; }
  .dt  { text-align: center; }
  .lnk { color: #0563C1; text-decoration: underline; }
</style>
</head>
<body>
<table>`;

    // Column widths
    html += '\n';
    COL_DEFS.forEach(def => {
      html += `<col width="${def.width * 8}">`;
    });

    // Header row
    html += '\n<tr>';
    COL_DEFS.forEach(def => {
      html += `<th>${esc(def.header)}</th>`;
    });
    html += '</tr>\n';

    // Data rows
    data.forEach((record, idx) => {
      const rowClass = idx % 2 === 0 ? ' class="even"' : '';
      html += `<tr${rowClass}>`;

      COL_DEFS.forEach(def => {
        let val = record[def.key] || '';
        if (def.key.endsWith('History')) val = formatHistory(val);

        if (def.type === 'currency') {
          const num = parseCurrency(val);
          if (num !== null) {
            html += `<td class="cur">${num.toFixed(2)}</td>`;
          } else {
            html += `<td class="cur">${esc(val)}</td>`;
          }
        } else if (def.type === 'number') {
          html += `<td class="num">${esc(val)}</td>`;
        } else if (def.type === 'date') {
          html += `<td class="dt">${esc(val)}</td>`;
        } else if (def.type === 'link' && val) {
          html += `<td class="lnk">${esc(val)}</td>`;
        } else {
          html += `<td>${esc(val)}</td>`;
        }
      });

      html += '</tr>\n';
    });

    html += '</table></body></html>';

    // Download as .xls (Excel opens HTML tables natively)
    const blob = new Blob([html], { type: 'application/vnd.ms-excel;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = XLSX_FILENAME.replace('.xlsx', '.xls');
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    log(`Downloaded ${XLSX_FILENAME.replace('.xlsx', '.xls')} (${(html.length / 1024 / 1024).toFixed(1)} MB)`);
    log('Open in Excel for full formatting (frozen header, filters, currency format).');
  }

  function downloadJSON(data) {
    const json = JSON.stringify(data, null, 2);
    const blob = new Blob([json], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = JSON_FILENAME;
    document.body.appendChild(a); a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    log(`Backup JSON downloaded: ${JSON_FILENAME}`);
  }

  // ═══════════════════════ MAIN ═══════════════════════

  try {
    if (!window.claimsLinks || !window.runRecaptcha) {
      throw new Error(
        'Not on the Stretto claims page, or page not fully loaded.\n' +
        `Navigate to ${CASE_URL} ` +
        'and wait for the table to appear.');
    }

    log('');
    log(`=== ${CASE_NAME} Claims Scraper v3 (fully autonomous) ===`);
    log('Paste once. Walk away. It handles WAF refreshes automatically.');
    log('');

    // Phase A
    if (S.basicData.length === 0) {
      await collectBasicData();
    } else {
      log(`Basic data already collected: ${S.basicData.length} claims`);
    }

    // Phase B
    await collectChildData();

    // Check completeness
    const finalDone = Object.keys(S.childData).length;
    const total = S.basicData.length;
    if (finalDone < total) {
      log(`Scraped ${finalDone}/${total}. Some may have failed — check errors.`);
    }

    // Phase C
    const combined = buildCombinedData();
    downloadJSON(combined);
    await sleep(1000);
    await generateExcel(combined);

    clearState();

    const elapsed = ((Date.now() - S.startTime) / 1000 / 60).toFixed(1);
    log('');
    log('==========================================================');
    log(`  COMPLETE in ${elapsed} min`);
    log(`  Claims: ${total} | Child rows: ${finalDone} | Errors: ${S.errors.length}`);
    log('==========================================================');

    if (S.errors.length > 0) {
      console.warn('Failed claim IDs:', S.errors.map(e => e.claimId));
    }

  } catch (err) {
    log('FATAL ERROR: ' + err.message);
    console.error(err);
    saveState();
    log('State saved to localStorage. Fix the issue and re-paste to resume.');
  }
})();
