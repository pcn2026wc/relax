const SPREADSHEET_ID = '1U4Q-6syclM8Qe2eSp4jBAfCqQFyiXbFUpApA8R7s5gQ';
const STATE_KEY = 'wc2026_state_json';
const TZ = 'Asia/Bangkok';
const FINE_AMOUNT = 20000;
const PREDICTION_CUTOFF_MINUTES = 5;

function doGet(e) {
  const action = String(e.parameter.action || 'load');
  const callback = e.parameter.callback;
  let data;
  try {
    if (action === 'ping') data = { ok: true, updatedAt: new Date().toISOString() };
    else if (action === 'syncdays') data = importDailySheets_();
    else if (action === 'syncresults') data = importResultsSheet_();
    else if (action === 'clearresultsday') data = clearResultsDay_(String(e.parameter.day || ''));
    else data = loadState_();
  } catch (err) {
    data = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  return output_(data, callback);
}

function doPost(e) {
  const action = String((e.parameter && e.parameter.action) || '');
  let data;
  try {
    const payload = parsePayload_(e);
    const realAction = action || payload.action || 'save';
    if (realAction === 'exportdays') data = exportDailySheets_(payload.state || payload);
    else data = saveState_(payload.state || payload);
  } catch (err) {
    data = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  return output_(data);
}

function parsePayload_(e) {
  const raw = (e.parameter && e.parameter.payload) || (e.postData && e.postData.contents) || '{}';
  return JSON.parse(raw || '{}');
}

function output_(data, callback) {
  const text = callback ? callback + '(' + JSON.stringify(data) + ');' : JSON.stringify(data);
  return ContentService.createTextOutput(text).setMimeType(callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}

function loadState_() {
  const raw = PropertiesService.getScriptProperties().getProperty(STATE_KEY);
  return raw ? JSON.parse(raw) : { ok: true, state: { predictions: {}, results: {}, settings: {}, audit: [] } };
}

function saveState_(state) {
  const data = { ok: true, type: 'wc2026-predict-state', savedAt: new Date().toISOString(), state: state || {} };
  PropertiesService.getScriptProperties().setProperty(STATE_KEY, JSON.stringify(data));
  ensureResultsSheet_(state || {});
  writeSummarySheets_(state || {});
  return data;
}

function exportDailySheets_(state) {
  if (!state || !state.matches || !state.players) throw new Error('Payload không có state.matches/state.players');
  saveState_(state);
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const byDay = {};
  state.matches.forEach(function(match) {
    const key = match.day || Utilities.formatDate(new Date(match.kickoff), TZ, 'yyyy-MM-dd');
    if (!byDay[key]) byDay[key] = [];
    byDay[key].push(match);
  });
  const report = Object.keys(byDay).sort().map(function(dayKey) {
    const matches = byDay[dayKey];
    const sheetName = dayName_(dayKey);
    const sh = getOrCreateSheet_(ss, sheetName);
    const header = ['STT', 'Người chơi'];
    matches.forEach(function(m) {
      header.push(m.id + ' ' + m.home + ' vs ' + m.away);
      header.push('Gio du doan ' + m.id);
    });
    const rows = [header];
    state.players.forEach(function(player, index) {
      const row = [index + 1, player.name];
      matches.forEach(function(match) {
        const rec = state.predictions && state.predictions[match.id] && state.predictions[match.id][player.name];
        row.push(rec ? choiceLabel_(rec.choice, match) : '');
        row.push(rec ? formatDateTime_(predictionSubmittedAt_(rec)) : '');
      });
      rows.push(row);
    });
    sh.clearContents();
    sh.getRange(1, 1, rows.length, header.length).setValues(rows);
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, header.length);
    const imported = matches.reduce(function(sum, match) {
      const byPlayer = state.predictions && state.predictions[match.id] ? state.predictions[match.id] : {};
      return sum + Object.keys(byPlayer).length;
    }, 0);
    return { sheet: sheetName, matches: matches.length, exported: imported };
  });
  ensureResultsSheet_(state);
  writeSummarySheets_(state);
  return { ok: true, savedAt: new Date().toISOString(), report: report, state: state };
}

function importResultsSheet_() {
  const saved = loadState_();
  const state = saved.state || saved;
  if (!state.matches) return { ok: true, state: state, imported: 0, report: [] };
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ensureResultsSheet_(state);
  const values = sh.getDataRange().getValues();
  const report = [];
  let imported = 0;
  state.results = state.results || {};
  values.slice(1).forEach(function(row) {
    const matchId = String(row[0] || '').trim();
    if (matchId) delete state.results[matchId];
  });
  values.slice(1).forEach(function(row) {
    const matchId = String(row[0] || '').trim();
    const match = state.matches.find(function(m) { return m.id === matchId; });
    if (!match) return;
    const homeScore = parseScore_(row[5]);
    const awayScore = parseScore_(row[6]);
    const status = String(row[7] || '').trim().toLowerCase();
    if (homeScore === null || awayScore === null) return;
    const result = resultOfScore_(homeScore, awayScore);
    state.results[matchId] = {
      homeScore: homeScore,
      awayScore: awayScore,
      result: result,
      status: status === 'xoa' || status === 'clear' ? '' : 'done',
      updatedAt: new Date().toISOString()
    };
    if (!state.results[matchId].status) delete state.results[matchId];
    else imported++;
    report.push({ match: matchId, score: homeScore + '-' + awayScore, result: result });
  });
  saveState_(state);
  return { ok: true, state: state, imported: imported, report: report };
}

function clearResultsDay_(dayKey) {
  if (!dayKey) return { ok: false, error: 'Missing day parameter' };
  const saved = loadState_();
  const state = saved.state || saved;
  if (!state.matches) return { ok: true, state: state, cleared: 0, report: [] };
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ensureResultsSheet_(state);
  const values = sh.getDataRange().getValues();
  const report = [];
  let cleared = 0;
  state.results = state.results || {};
  values.slice(1).forEach(function(row, idx) {
    const matchId = String(row[0] || '').trim();
    const day = String(row[1] || '').trim();
    if (day !== dayKey) return;
    const hadSheetResult = row[5] !== '' || row[6] !== '' || row[7] !== '' || row[8] !== '';
    sh.getRange(idx + 2, 6, 1, 4).clearContent();
    if (state.results[matchId]) {
      delete state.results[matchId];
    }
    if (hadSheetResult) cleared++;
    report.push({ match: matchId, day: dayKey, cleared: true });
  });
  saveState_(state);
  return { ok: true, state: state, cleared: cleared, report: report };
}

function ensureResultsSheet_(state) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = getOrCreateSheet_(ss, 'KQ_TranDau');
  const header = ['Ma tran', 'Ngay', 'Doi 1', 'Doi 2', 'Gio VN', 'Ban doi 1', 'Ban doi 2', 'Trang thai', 'Ket qua'];
  const existing = {};
  const values = sh.getDataRange().getValues();
  if (values.length > 1) {
    values.slice(1).forEach(function(row) {
      const id = String(row[0] || '').trim();
      if (id) existing[id] = row;
    });
  }
  const rows = [header];
  (state.matches || []).forEach(function(m) {
    const old = existing[m.id] || [];
    const homeScore = old[5] !== undefined ? old[5] : '';
    const awayScore = old[6] !== undefined ? old[6] : '';
    const status = old[7] || '';
    const resultText = homeScore !== '' && awayScore !== '' && !isNaN(Number(homeScore)) && !isNaN(Number(awayScore))
      ? choiceLabel_(resultOfScore_(Number(homeScore), Number(awayScore)), m)
      : '';
    rows.push([m.id, m.day || dateKey_(m.kickoff), m.home, m.away, Utilities.formatDate(new Date(m.kickoff), TZ, 'HH:mm dd/MM/yyyy'), homeScore, awayScore, status, resultText]);
  });
  clearSheet_(sh);
  sh.getRange(1, 1, rows.length, header.length).setValues(rows);
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, header.length);
  return sh;
}

function importDailySheets_() {
  const saved = loadState_();
  const state = saved.state || saved;
  if (!state.matches || !state.players) return saved;
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const report = [];
  state.matches.forEach(function(match) {
    const dayKey = match.day || Utilities.formatDate(new Date(match.kickoff), TZ, 'yyyy-MM-dd');
    const sh = ss.getSheetByName(dayName_(dayKey));
    if (!sh) return;
    const values = sh.getDataRange().getValues();
    if (values.length < 2) return;
    const header = values[0].map(String);
    const col = header.findIndex(function(h) { return h.indexOf(match.id) === 0; });
    if (col < 0) return;
    const timeCol = header.findIndex(function(h) { return h.toLowerCase().indexOf('gio du doan ' + match.id.toLowerCase()) >= 0; });
    let imported = 0;
    values.slice(1).forEach(function(row) {
      const playerName = String(row[1] || '').trim();
      const choice = parseChoice_(String(row[col] || ''), match);
      if (!playerName || !choice) return;
      const submittedAt = parseDateTime_(timeCol >= 0 ? row[timeCol] : '') || new Date().toISOString();
      state.predictions = state.predictions || {};
      state.predictions[match.id] = state.predictions[match.id] || {};
      state.predictions[match.id][playerName] = {
        choice: choice,
        updatedAt: submittedAt,
        submittedAt: submittedAt,
        submitCount: 1,
        history: [{ choice: choice, updatedAt: submittedAt, submittedAt: submittedAt }]
      };
      imported++;
    });
    report.push({ sheet: dayName_(dayKey), match: match.id, imported: imported });
  });
  writeSummarySheets_(state);
  saveState_(state);
  return { ok: true, state: state, report: report };
}

function writeSummarySheets_(state) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const summary = buildSummaryData_(state || {});
  const resultSheet = getOrCreateSheet_(ss, 'SheetKetQua');
  clearSheet_(resultSheet);
  const resultHeader = ['Ma tran', 'Ngay', 'Tran', 'Ty so', 'Ket qua', 'Dung', 'Sai', 'Thieu', 'Tong tien phat'];
  resultSheet.getRange(1, 1, 1, resultHeader.length).setValues([resultHeader]);
  if (summary.results.length) {
    const values = summary.results.map(function(r) {
      return [r.matchId, r.date, r.home + ' vs ' + r.away, r.homeScore + '-' + r.awayScore, r.result, r.correct, r.wrong, r.missing, r.totalPenalty];
    });
    resultSheet.getRange(2, 1, values.length, resultHeader.length).setValues(values);
  }
  resultSheet.getRange(summary.results.length + 2, 1, 1, resultHeader.length).setValues([['Tổng tiền phạt', '', '', '', '', '', '', '', summary.totalPenalty]]);
  resultSheet.setFrozenRows(1);
  resultSheet.autoResizeColumns(1, resultHeader.length);

  const thongSheet = getOrCreateSheet_(ss, 'thongkedudoan');
  clearSheet_(thongSheet);
  const daySummaries = summary.thongKeDuDoan || [];
  const thongHeader = daySummaries.map(function(daySummary) {
    return 'Chua du doan ' + daySummary.dayLabel;
  });
  if (thongHeader.length) {
    thongSheet.getRange(1, 1, 1, thongHeader.length).setValues([thongHeader]);
    const maxMissing = Math.max.apply(null, [1].concat(daySummaries.map(function(daySummary) {
      return daySummary.missingPlayers.length;
    })));
    const values = [];
    for (let i = 0; i < maxMissing; i++) {
      values.push(daySummaries.map(function(daySummary) {
        return daySummary.missingPlayers[i] || '';
      }));
    }
    thongSheet.getRange(2, 1, values.length, thongHeader.length).setValues(values);
  }
  thongSheet.setFrozenRows(1);
  if (thongHeader.length) thongSheet.autoResizeColumns(1, thongHeader.length);
}

function buildSummaryData_(state) {
  const matches = state && state.matches ? state.matches : [];
  const players = state && state.players ? state.players : [];
  const results = state && state.results ? state.results : {};
  const predictions = state && state.predictions ? state.predictions : {};
  const resultRows = [];
  let totalPenalty = 0;

  matches.filter(function(m) {
    return results[m.id] && results[m.id].status === 'done';
  }).forEach(function(m) {
    const res = results[m.id];
    let correct = 0;
    let wrong = 0;
    let missing = 0;
    players.forEach(function(p) {
      const rec = predictions[m.id] && predictions[m.id][p.name] ? predictions[m.id][p.name] : null;
      const pred = isPredictionValidForScoring_(m, rec) ? rec.choice : '';
      if (!pred) missing++;
      else if (pred === res.result) correct++;
      else wrong++;
    });
    const penalty = (wrong + missing) * FINE_AMOUNT;
    totalPenalty += penalty;
    resultRows.push({
      matchId: m.id,
      date: m.day || dateKey_(m.kickoff),
      home: m.home,
      away: m.away,
      homeScore: res.homeScore,
      awayScore: res.awayScore,
      result: choiceLabel_(res.result, m),
      correct: correct,
      wrong: wrong,
      missing: missing,
      totalPenalty: penalty
    });
  });

  const matchDays = Array.from(new Set(matches.map(function(m) { return m.day || dateKey_(m.kickoff); }))).sort();
  const baseDay = currentMatchDayKey_(matchDays);
  const baseIndex = Math.max(0, matchDays.indexOf(baseDay));
  const recentDays = matchDays.slice(baseIndex, baseIndex + 5);
  const missingRows = [];
  recentDays.forEach(function(day) {
    const dayMatches = matches.filter(function(m) { return (m.day || dateKey_(m.kickoff)) === day; });
    const missingPlayers = players.filter(function(player) {
      return dayMatches.some(function(m) { return !(predictions[m.id] && predictions[m.id][player.name]); });
    }).map(function(player) { return player.name; });
    missingRows.push({ day: day, dayLabel: formatDayLabel_(day), missingPlayers: missingPlayers });
  });

  return { results: resultRows, totalPenalty: totalPenalty, thongKeDuDoan: missingRows };
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function clearSheet_(sheet) {
  const maxRows = sheet.getMaxRows();
  const maxCols = sheet.getMaxColumns();
  if (maxRows > 1 || maxCols > 1) {
    sheet.getRange(1, 1, maxRows, maxCols).clear();
  }
}

function dayName_(dayKey) {
  const parts = String(dayKey).split('-');
  return parts[2] + '/' + parts[1];
}

function dateKey_(iso) {
  return Utilities.formatDate(new Date(iso), TZ, 'yyyy-MM-dd');
}

function formatDayLabel_(dayKey) {
  if (!dayKey) return '';
  const parts = String(dayKey).split('-');
  return parts.length === 3 ? parts[2] + '/' + parts[1] : String(dayKey);
}

function currentMatchDayKey_(days) {
  if (!days || !days.length) return '';
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  return days.filter(function(day) { return day >= today; })[0] || days[0];
}

function predictionSubmittedAt_(record) {
  if (!record) return '';
  if (record.submittedAt) return record.submittedAt;
  if (record.updatedAt) return record.updatedAt;
  if (record.history && record.history.length) return record.history[record.history.length - 1].submittedAt || record.history[record.history.length - 1].updatedAt || '';
  return '';
}

function isPredictionValidForScoring_(match, record) {
  const submittedAt = predictionSubmittedAt_(record);
  if (!record || !record.choice || !submittedAt) return false;
  return new Date(submittedAt).getTime() <= new Date(match.kickoff).getTime() - PREDICTION_CUTOFF_MINUTES * 60000;
}

function formatDateTime_(iso) {
  if (!iso) return '';
  return Utilities.formatDate(new Date(iso), TZ, 'dd/MM/yyyy HH:mm:ss');
}

function parseDateTime_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) return value.toISOString();
  const s = String(value).trim();
  if (!s) return '';
  const direct = new Date(s);
  if (!isNaN(direct.getTime())) return direct.toISOString();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?$/);
  if (!m) return '';
  return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]), Number(m[4]), Number(m[5]), Number(m[6] || 0)).toISOString();
}

function parseScore_(value) {
  if (value === '' || value === null || value === undefined) return null;
  const n = Number(value);
  return Number.isFinite(n) && n >= 0 ? n : null;
}

function resultOfScore_(homeScore, awayScore) {
  if (homeScore > awayScore) return 'H';
  if (homeScore < awayScore) return 'A';
  return 'D';
}

function choiceLabel_(choice, match) {
  if (choice === 'H') return match.home + ' thắng';
  if (choice === 'D') return 'Hòa';
  if (choice === 'A') return match.away + ' thắng';
  return '';
}

function parseChoice_(text, match) {
  const s = text.toLowerCase();
  if (!s) return '';
  if (s.indexOf('hòa') >= 0 || s.indexOf('hoa') >= 0) return 'D';
  if (s.indexOf(String(match.home).toLowerCase()) >= 0) return 'H';
  if (s.indexOf(String(match.away).toLowerCase()) >= 0) return 'A';
  if (s === 'h') return 'H';
  if (s === 'd') return 'D';
  if (s === 'a') return 'A';
  return '';
}
