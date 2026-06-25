function doGet() {
  return HtmlService.createHtmlOutputFromFile('Predictor')
    .setTitle('⚽ FIFA 2026 Predictor')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── Auth ─────────────────────────────────────────────────────

function lookupParticipant(email) {
  const data = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(SHEETS.PARTICIPANTS).getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][2].toString().toLowerCase() === email.toLowerCase()) {
      const groups = _getGroupsForParticipant(data[i][0]);
      return { id: data[i][0], name: data[i][1], groups: groups };
    }
  }
  return null;
}

function registerParticipant(name, email, inviteCode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const groupName = _resolveInviteCode(inviteCode);
  if (!groupName) return { error: 'Invalid invite code. Please check with your group admin.' };

  if (lookupParticipant(email)) return { error: 'Email already registered. Go back and enter your email to predict.' };

  const participantId = Utilities.getUuid();
  ss.getSheetByName(SHEETS.PARTICIPANTS)
    .appendRow([participantId, name, email, new Date().toISOString()]);
  ss.getSheetByName(SHEETS.MEMBERSHIPS)
    .appendRow([Utilities.getUuid(), participantId, groupName, new Date().toISOString()]);

  return { success: true, name, groups: [groupName] };
}

function joinGroup(participantId, inviteCode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const groupName = _resolveInviteCode(inviteCode);
  if (!groupName) return { error: 'Invalid invite code. Please check with your group admin.' };

  // Prevent duplicate membership
  const existing = _getGroupsForParticipant(participantId);
  if (existing.indexOf(groupName) !== -1)
    return { error: 'You are already in the group "' + groupName + '".' };

  ss.getSheetByName(SHEETS.MEMBERSHIPS)
    .appendRow([Utilities.getUuid(), participantId, groupName, new Date().toISOString()]);

  return { success: true, groupName: groupName };
}

// ── Predictions ──────────────────────────────────────────────

function getUpcomingMatches() {
  var data = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(SHEETS.MATCHES).getDataRange().getValues();
  var now = new Date();
  var cutoffMs = 60 * 60 * 1000;

  return data.slice(1)
    .filter(function(r) { return r[4] && !isNaN(new Date(r[4])); })
    .map(function(r) {
      var kickoff = new Date(r[4]);
      var status = r[9] ? String(r[9]) : '';
      var isCompleted = status === 'Completed';
      var isLive = status === 'Live';
      var isUpcoming = (kickoff - now) > cutoffMs;
      return {
        id: r[0], stage: r[1], date: r[4], group: r[3], home: r[5], away: r[6],
        homeScore: isCompleted || isLive ? r[7] : null,
        awayScore: isCompleted || isLive ? r[8] : null,
        status: isCompleted ? 'completed' : isLive ? 'live' : isUpcoming ? 'upcoming' : 'started'
      };
    })
    .filter(function(m) { return m.status !== 'started'; });
}

function getMyPredictions(participantId) {
  const data = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(SHEETS.PREDICTIONS).getDataRange().getValues();
  const preds = {};
  data.slice(1).forEach(r => {
    if (r[1] == participantId) preds[r[2]] = { home: r[3], away: r[4] };
  });
  return preds;
}

function submitPrediction(participantId, matchId, predHome, predAway) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(SHEETS.PREDICTIONS);
  const data  = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] == participantId && data[i][2] == matchId) {
      sheet.getRange(i + 1, 4, 1, 3)
        .setValues([[predHome, predAway, new Date().toISOString()]]);
      return { success: true };
    }
  }
  sheet.appendRow([Utilities.getUuid(), participantId, matchId,
    predHome, predAway, new Date().toISOString()]);
  return { success: true };
}

// ── Leaderboard ──────────────────────────────────────────────

const STAGE_ORDER = [
  'First Stage', 'Round of 32', 'Round of 16',
  'Quarter-final', 'Semi-final', 'Play-off for third place', 'Final'
];

function _stageIndex(s) {
  const i = STAGE_ORDER.indexOf(s);
  return i === -1 ? 99 : i;
}

/**
 * Returns { rows, stages } where rows are ranked leaderboard entries
 * (each with per-stage stageScores map) and stages is the ordered list
 * of stages that have completed matches.
 * groupName: null → global (all groups), string → filter to that group.
 */
function getLeaderboard(groupName, participantId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Build participantId → name map (and optional group filter)
  const partData = ss.getSheetByName(SHEETS.PARTICIPANTS).getDataRange().getValues().slice(1);
  const memberData = ss.getSheetByName(SHEETS.MEMBERSHIPS).getDataRange().getValues().slice(1);

  // Determine which participantIds are in scope
  let scopedIds = null;
  if (groupName) {
    scopedIds = new Set(
      memberData.filter(r => r[2] === groupName).map(r => r[1])
    );
  } else if (participantId) {
    // Global view: limit to participants sharing at least one group with the current user
    const userGroups = new Set(_getGroupsForParticipant(participantId));
    scopedIds = new Set(
      memberData.filter(r => userGroups.has(r[2])).map(r => r[1])
    );
  }

  const nameMap = {};
  partData.forEach(r => {
    if (!scopedIds || scopedIds.has(r[0])) nameMap[r[0]] = r[1];
  });

  // Load completed matches with scores
  const matchData = ss.getSheetByName(SHEETS.MATCHES).getDataRange().getValues().slice(1);
  const matchMap = {};
  matchData.forEach(r => {
    if (r[9] === 'Completed' && r[7] !== '' && r[8] !== '') {
      matchMap[r[0]] = {
        homeScore: Number(r[7]),
        awayScore: Number(r[8]),
        stage:     r[1] || 'Other',
        group:     r[3],
      };
    }
  });

  // Collect stages present in completed matches, sorted canonically
  const stageSet = new Set();
  Object.values(matchMap).forEach(m => stageSet.add(m.stage));
  const stages = [...stageSet].sort((a, b) => _stageIndex(a) - _stageIndex(b));

  // Build stage → [matchId] map for missed-prediction counting
  const stageMatchIds = {};
  Object.keys(matchMap).forEach(mid => {
    const stage = matchMap[mid].stage;
    if (!stageMatchIds[stage]) stageMatchIds[stage] = [];
    stageMatchIds[stage].push(mid);
  });

  // Tally scores per participant with overall and per-stage breakdown
  const scores = {};
  const breakdown = {};
  const stageScores = {};
  Object.keys(nameMap).forEach(id => {
    scores[id] = 0;
    breakdown[id] = { exactScore: 0, correctResult: 0, incorrect: 0, noPrediction: 0 };
    stageScores[id] = {};
  });

  const completedMatchIds = Object.keys(matchMap);
  const predData = ss.getSheetByName(SHEETS.PREDICTIONS).getDataRange().getValues().slice(1);

  function _ensureStage(pid, stage) {
    if (!stageScores[pid][stage]) {
      stageScores[pid][stage] = { score: 0, exactScore: 0, correctResult: 0, incorrect: 0, noPrediction: 0 };
    }
  }

  // Track which matches each participant predicted
  const predictedMatches = {};
  predData.forEach(r => {
    const pid = r[1], mid = String(r[2]);
    if (!nameMap[pid]) return;
    const match = matchMap[mid];
    if (!match) return;
    if (!predictedMatches[pid]) predictedMatches[pid] = new Set();
    predictedMatches[pid].add(mid);

    const predHome = Number(r[3]), predAway = Number(r[4]);
    // Group-stage matches have a non-empty GroupName (e.g. "Group A"); knockout matches don't
    const isKnockout = !match.group;
    const mult = isKnockout ? SCORING.KNOCKOUT_MULT : 1;
    const stage = match.stage;
    _ensureStage(pid, stage);

    if (predHome === match.homeScore && predAway === match.awayScore) {
      scores[pid] += SCORING.CORRECT_SCORE * mult;
      breakdown[pid].exactScore++;
      stageScores[pid][stage].score += SCORING.CORRECT_SCORE * mult;
      stageScores[pid][stage].exactScore++;
    } else {
      const predResult = Math.sign(predHome - predAway);
      const realResult = Math.sign(match.homeScore - match.awayScore);
      if (predResult === realResult) {
        scores[pid] += SCORING.CORRECT_RESULT * mult;
        breakdown[pid].correctResult++;
        stageScores[pid][stage].score += SCORING.CORRECT_RESULT * mult;
        stageScores[pid][stage].correctResult++;
      } else {
        breakdown[pid].incorrect++;
        stageScores[pid][stage].incorrect++;
      }
    }
  });

  // Count missed predictions overall and per stage
  Object.keys(nameMap).forEach(pid => {
    const predicted = predictedMatches[pid] || new Set();
    completedMatchIds.forEach(mid => {
      if (!predicted.has(mid)) {
        breakdown[pid].noPrediction++;
        const stage = matchMap[mid].stage;
        _ensureStage(pid, stage);
        stageScores[pid][stage].noPrediction++;
      }
    });
  });

  // Build participantId → groups map
  const groupsMap = {};
  memberData.forEach(r => {
    if (!groupsMap[r[1]]) groupsMap[r[1]] = [];
    groupsMap[r[1]].push(r[2]);
  });

  // Sort and rank
  const ranked = Object.keys(scores)
    .map(id => ({
      id, name: nameMap[id], score: scores[id],
      groups: groupsMap[id] || [],
      breakdown: breakdown[id],
      stageScores: stageScores[id],
    }))
    .sort((a, b) => b.score - a.score);

  let rank = 1;
  ranked.forEach((entry, i) => {
    if (i > 0 && ranked[i - 1].score > entry.score) rank = i + 1;
    entry.rank = rank;
  });

  return { rows: ranked, stages: stages };
}

function getGroupNames() {
  const data = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(SHEETS.GROUPS).getDataRange().getValues().slice(1);
  return data.map(r => r[1]);
}

// ── Helpers ───────────────────────────────────────────────────

function _resolveInviteCode(inviteCode) {
  const groups = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(SHEETS.GROUPS).getDataRange().getValues();
  for (let i = 1; i < groups.length; i++) {
    if (groups[i][0].toString().toUpperCase() === inviteCode.toUpperCase())
      return groups[i][1];
  }
  return null;
}

function _getGroupsForParticipant(participantId) {
  const data = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(SHEETS.MEMBERSHIPS).getDataRange().getValues();
  const groups = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === participantId) groups.push(data[i][2]);
  }
  return groups;
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
