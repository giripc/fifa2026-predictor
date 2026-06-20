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

/**
 * One-time migration: reads PoolGroup from the old PARTICIPANTS column 4
 * and writes a MEMBERSHIPS row for each participant that doesn't already have one.
 * Safe to run multiple times — skips participants already migrated.
 * After running, manually delete the PoolGroup column from PARTICIPANTS.
 */
function migratePoolGroupToMemberships() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Create MEMBERSHIPS sheet if it doesn't exist yet
  _ensureSheet(ss, SHEETS.MEMBERSHIPS, [[
    'MembershipId','ParticipantId','GroupName','JoinDate'
  ]]);

  const partSheet = ss.getSheetByName(SHEETS.PARTICIPANTS);
  const memberSheet = ss.getSheetByName(SHEETS.MEMBERSHIPS);

  const partData = partSheet.getDataRange().getValues();
  const memberData = memberSheet.getDataRange().getValues();

  // Build set of participantIds already in MEMBERSHIPS
  const alreadyMigrated = new Set(memberData.slice(1).map(r => r[1]));

  // Find which column index PoolGroup is in (header row)
  const headers = partData[0].map(h => h.toString().toLowerCase());
  const poolGroupCol = headers.indexOf('poolgroup');
  if (poolGroupCol === -1) {
    SpreadsheetApp.getUi().alert('PoolGroup column not found — may already be removed. Nothing to migrate.');
    return;
  }

  // Also find JoinDate column for use as membership date
  const joinDateCol = headers.indexOf('joindate');

  let migrated = 0;
  const newRows = [];
  partData.slice(1).forEach(function(r) {
    const participantId = r[0];
    const groupName = r[poolGroupCol] ? r[poolGroupCol].toString().trim() : '';
    if (!groupName || alreadyMigrated.has(participantId)) return;
    const joinDate = joinDateCol !== -1 && r[joinDateCol] ? r[joinDateCol] : new Date().toISOString();
    newRows.push([Utilities.getUuid(), participantId, groupName, joinDate]);
    migrated++;
  });

  if (newRows.length) {
    memberSheet.getRange(memberSheet.getLastRow() + 1, 1, newRows.length, 4).setValues(newRows);
  }

  SpreadsheetApp.getUi().alert(
    '✅ Migration complete: ' + migrated + ' participant(s) moved to MEMBERSHIPS.\n\n' +
    'You can now delete the PoolGroup column from the PARTICIPANTS sheet.'
  );
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
