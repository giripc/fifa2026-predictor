function fetchMatches() {
  const url = FIFA.BASE + '/calendar/matches'
    + '?idCompetition=' + FIFA.COMPETITION
    + '&idSeason=' + FIFA.SEASON
    + '&count=500&language=en&timeZoneOffset=0';

  var resp;
  try {
    resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  } catch (e) {
    SpreadsheetApp.getUi().alert('Network error: ' + e.message);
    return;
  }

  if (resp.getResponseCode() !== 200) {
    SpreadsheetApp.getUi().alert(
      'FIFA API error ' + resp.getResponseCode() + ':\n' +
      resp.getContentText().substring(0, 300)
    );
    return;
  }

  var matches = JSON.parse(resp.getContentText()).Results || [];
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MATCHES);
  if (!sheet) { SpreadsheetApp.getUi().alert('Run Setup first.'); return; }

  if (sheet.getLastRow() > 1)
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();

  var now   = new Date();
  var nowTs = now.toISOString();

  var allRows = matches.map(function(m) {
    var home = m.Home || m.HomeTeam || {};
    var away = m.Away || m.AwayTeam || {};
    return [
      m.IdMatch  || '',
      _desc(m.StageName),
      m.MatchDay || '',
      _desc(m.GroupName),
      m.Date     || m.LocalDate || '',
      _teamName(home),
      _teamName(away),
      home.Score !== undefined ? home.Score : '',
      away.Score !== undefined ? away.Score : '',
      _status(m.MatchStatus),
      nowTs,
    ];
  });

  // Keep all valid matches (past and future) so completed scores are available for leaderboard
  var rows = allRows.filter(function(r) { return r[4]; });

  if (rows.length)
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);

  var completed = rows.filter(function(r) { return r[9] === 'Completed'; }).length;
  var upcoming  = rows.filter(function(r) { return new Date(r[4]) > now; }).length;
  SpreadsheetApp.getUi().alert(
    '✅ Loaded ' + rows.length + ' matches\n' +
    '   • ' + upcoming  + ' upcoming\n' +
    '   • ' + completed + ' completed (used for leaderboard)'
  );
}

function debugFifaApi() {
  var url = FIFA.BASE + '/calendar/matches'
    + '?idCompetition=' + FIFA.COMPETITION
    + '&idSeason=' + FIFA.SEASON
    + '&count=5&language=en&timeZoneOffset=0';
  var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  Logger.log('Status: ' + resp.getResponseCode());
  Logger.log('Body: ' + resp.getContentText().substring(0, 1000));
}

function fetchSeasons() {
  var url = FIFA.BASE + '/seasons?idCompetition=' + FIFA.COMPETITION + '&language=en';
  var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  Logger.log('Status: ' + resp.getResponseCode());
  Logger.log('Body: ' + resp.getContentText().substring(0, 2000));
}

function _teamName(t) {
  if (!t) return 'TBD';
  return (t.TeamName && t.TeamName[0] && t.TeamName[0].Description)
      || t.Abbreviation || t.Name || 'TBD';
}

function _desc(field) {
  if (!field) return '';
  return Array.isArray(field) ? (field[0] && field[0].Description) || '' : field;
}

function _status(s) {
  return { 0:'Upcoming', 1:'Live', 3:'Completed', 5:'Postponed' }[s] || String(s || '');
}
