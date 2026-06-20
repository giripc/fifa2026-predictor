function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  _ensureSheet(ss, SHEETS.CONFIG, [
    ['Key', 'Value'],
    ['Pool Name',            'FIFA 2026 Predictor'],
    ['Prediction Cutoff',    '60'],   // minutes before kickoff
    ['Scoring:Result',       SCORING.CORRECT_RESULT],
    ['Scoring:Score',        SCORING.CORRECT_SCORE],
    ['Scoring:KnockoutMult', SCORING.KNOCKOUT_MULT],
  ]);

  _ensureSheet(ss, SHEETS.MATCHES, [[
    'MatchId','Stage','MatchDay','Group','Date',
    'HomeTeam','AwayTeam','HomeScore','AwayScore','Status','FetchedAt'
  ]]);

  _ensureSheet(ss, SHEETS.PARTICIPANTS, [[
    'ParticipantId','Name','Email','JoinDate'
  ]]);

  _ensureSheet(ss, SHEETS.MEMBERSHIPS, [[
    'MembershipId','ParticipantId','GroupName','JoinDate'
  ]]);

  _ensureSheet(ss, SHEETS.PREDICTIONS, [[
    'PredictionId','ParticipantId','MatchId',
    'PredHome','PredAway','SubmittedAt'
  ]]);

  _ensureSheet(ss, SHEETS.GROUPS, [[
    'InviteCode', 'GroupName', 'CreatedAt'
  ]]);

  SpreadsheetApp.getUi().alert('✅ Sheets initialized!');
}

function _ensureSheet(ss, name, rows) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
    sh.getRange(1, 1, 1, rows[0].length)
      .setBackground('#1a4e8c').setFontColor('#ffffff').setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
