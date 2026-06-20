function doGet() {
  return HtmlService.createHtmlOutputFromFile('Predictor')
    .setTitle('⚽ FIFA 2026 Predictor')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function lookupParticipant(email) {
  const data = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(SHEETS.PARTICIPANTS).getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][2].toString().toLowerCase() === email.toLowerCase())
      return { id: data[i][0], name: data[i][1], group: data[i][3] };
  }
  return null;
}

function registerParticipant(name, email, inviteCode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const groups = ss.getSheetByName(SHEETS.GROUPS).getDataRange().getValues();
  let groupName = null;
  for (let i = 1; i < groups.length; i++) {
    if (groups[i][0].toString().toUpperCase() === inviteCode.toUpperCase()) {
      groupName = groups[i][1];
      break;
    }
  }
  if (!groupName) return { error: 'Invalid invite code. Please check with your group admin.' };

  if (lookupParticipant(email)) return { error: 'Email already registered. Go back and enter your email to predict.' };

  ss.getSheetByName(SHEETS.PARTICIPANTS)
    .appendRow([Utilities.getUuid(), name, email, groupName, new Date().toISOString()]);

  return { success: true, name, group: groupName };
}

function getUpcomingMatches() {
  var data = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(SHEETS.MATCHES).getDataRange().getValues();
  var now = new Date();
  var cutoffMs = 60 * 60 * 1000;

  return data.slice(1)
    .filter(function(r) {
      var kickoff = new Date(r[4]);
      return !isNaN(kickoff) && (kickoff - now) > cutoffMs;
    })
    .map(function(r) {
      return {
        id:    r[0],
        stage: r[1],
        date:  r[4],
        group: r[3],
        home:  r[5],
        away:  r[6],
      };
    });
}

function getMyPredictions(participantId) {
  const data = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(SHEETS.PREDICTIONS).getDataRange().getValues();
  const preds = {};
  data.slice(1).forEach(r => {
    if (r[1] === participantId) preds[r[2]] = { home: r[3], away: r[4] };
  });
  return preds;
}

function submitPrediction(participantId, matchId, predHome, predAway) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(SHEETS.PREDICTIONS);
  const data  = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === participantId && data[i][2] === matchId) {
      sheet.getRange(i + 1, 4, 1, 3)
        .setValues([[predHome, predAway, new Date().toISOString()]]);
      return { success: true };
    }
  }
  sheet.appendRow([
    Utilities.getUuid(), participantId, matchId,
    predHome, predAway, new Date().toISOString()
  ]);
  return { success: true };
}

function debugMatches() {
  const data = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(SHEETS.MATCHES).getDataRange().getValues();
  const now = new Date();
  Logger.log('Total rows: ' + (data.length - 1));
  Logger.log('Now: ' + now.toISOString());
  data.slice(1, 6).forEach(function(r) {
    Logger.log('Status: [' + r[9] + '] Date: [' + r[4] + '] Kickoff: ' + new Date(r[4]).toISOString());
  });
}
