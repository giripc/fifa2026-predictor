function doGet() {
  return HtmlService.createHtmlOutputFromFile('Predictor')
    .setTitle('⚽ FIFA 2026 Predictor')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── Auth ─────────────────────────────────────────────────────

function _getVerifiedEmail() {
  const email = Session.getActiveUser().getEmail();
  if (!email) throw new Error('Could not verify your identity. Please ensure you are signed in with your Google account.');
  return email.toLowerCase();
}

function lookupParticipant() {
  const email = _getVerifiedEmail();
  const data = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(SHEETS.PARTICIPANTS).getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][2].toString().toLowerCase() === email) {
      const groups = _getGroupsForParticipant(data[i][0]);
      return { id: data[i][0], name: data[i][1], email: email, groups: groups };
    }
  }
  return null;
}

function registerParticipant(name, inviteCode) {
  const email = _getVerifiedEmail();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const groupName = _resolveInviteCode(inviteCode);
  if (!groupName) return { error: 'Invalid invite code. Please check with your group admin.' };

  if (lookupParticipant()) return { error: 'Email already registered. Go back to sign in.' };

  const lock = LockService.getScriptLock();
  lock.tryLock(5000);
  try {
    if (lookupParticipant()) return { error: 'Email already registered. Go back to sign in.' };
    const participantId = Utilities.getUuid();
    ss.getSheetByName(SHEETS.PARTICIPANTS)
      .appendRow([participantId, name, email, new Date().toISOString()]);
    ss.getSheetByName(SHEETS.MEMBERSHIPS)
      .appendRow([Utilities.getUuid(), participantId, groupName, new Date().toISOString()]);
    return { success: true, name, groups: [groupName] };
  } finally {
    lock.releaseLock();
  }
}

function joinGroup(inviteCode) {
  const email = _getVerifiedEmail();
  const participant = lookupParticipant();
  if (!participant) return { error: 'Participant not found. Please register first.' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const groupName = _resolveInviteCode(inviteCode);
  if (!groupName) return { error: 'Invalid invite code. Please check with your group admin.' };

  const existing = _getGroupsForParticipant(participant.id);
  if (existing.indexOf(groupName) !== -1)
    return { error: 'You are already in the group "' + groupName + '".' };

  ss.getSheetByName(SHEETS.MEMBERSHIPS)
    .appendRow([Utilities.getUuid(), participant.id, groupName, new Date().toISOString()]);

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
    });
  // 'started' matches (past kickoff but not yet marked Live/Completed) are shown locked
}

function getMyPredictions() {
  const participant = lookupParticipant();
  if (!participant) return {};
  const data = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(SHEETS.PREDICTIONS).getDataRange().getValues();
  const preds = {};
  data.slice(1).forEach(r => {
    if (r[1] === participant.id) preds[String(r[2])] = { home: r[3], away: r[4] };
  });
  return preds;
}

function submitPrediction(matchId, predHome, predAway) {
  const participant = lookupParticipant();
  if (!participant) return { error: 'Unauthorized. Please sign in.' };
  const participantId = participant.id;

  // Validate score values
  const h = parseInt(predHome), a = parseInt(predAway);
  if (isNaN(h) || isNaN(a) || h < 0 || h > 20 || a < 0 || a > 20)
    return { error: 'Invalid score values.' };

  // Server-side cutoff check — reject if match is within 1 hour of kickoff or already started
  const matchData = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(SHEETS.MATCHES).getDataRange().getValues();
  const matchRow = matchData.slice(1).find(r => String(r[0]) === String(matchId));
  if (!matchRow) return { error: 'Match not found.' };

  const kickoff = new Date(matchRow[4]);
  const status  = matchRow[9] ? String(matchRow[9]) : '';
  const now     = new Date();
  const cutoffMs = 60 * 60 * 1000;

  if (status === 'Completed' || status === 'Live') {
    return { error: 'Predictions are locked — match has already started or finished.' };
  }
  if ((kickoff - now) <= cutoffMs) {
    return { error: 'Predictions are locked — less than 1 hour to kickoff.' };
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(SHEETS.PREDICTIONS);
  const data  = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === participantId && String(data[i][2]) === String(matchId)) {
      sheet.getRange(i + 1, 4, 1, 3)
        .setValues([[h, a, new Date().toISOString()]]);
      return { success: true };
    }
  }
  sheet.appendRow([Utilities.getUuid(), participantId, matchId,
    h, a, new Date().toISOString()]);
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
function getLeaderboard(groupName) {
  const participant = lookupParticipant();
  const participantId = participant ? participant.id : null;
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
  const knockoutScores = {};  // proximity scoring, Round of 32+ only
  const breakdown = {};
  const stageScores = {};
  Object.keys(nameMap).forEach(id => {
    scores[id] = 0;
    knockoutScores[id] = 0;
    breakdown[id] = { exactScore: 0, correctResult: 0, incorrect: 0, noPrediction: 0 };
    stageScores[id] = {};
  });

  const completedMatchIds = Object.keys(matchMap);
  const predData = ss.getSheetByName(SHEETS.PREDICTIONS).getDataRange().getValues().slice(1);

  function _ensureStage(pid, stage) {
    if (!stageScores[pid][stage]) {
      stageScores[pid][stage] = { score: 0, koScore: 0, exactScore: 0, correctResult: 0, incorrect: 0, noPrediction: 0 };
    }
  }

  // Returns proximity score for a knockout match prediction
  function _koScore(predHome, predAway, actualHome, actualAway) {
    const predResult = Math.sign(predHome - predAway);
    const realResult = Math.sign(actualHome - actualAway);
    if (predResult !== realResult) return SCORING.KO_WRONG;
    if (predHome === actualHome && predAway === actualAway) return SCORING.KO_EXACT;
    const diff = Math.abs(predHome - actualHome) + Math.abs(predAway - actualAway);
    return Math.max(SCORING.KO_FLOOR, SCORING.KO_EXACT - diff);
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

    if (isKnockout) {
      // Proximity scoring for knockout rounds
      const pts = _koScore(predHome, predAway, match.homeScore, match.awayScore);
      knockoutScores[pid] += pts;
      stageScores[pid][stage].koScore += pts;

      // Also track exact/correct/incorrect for breakdown chips
      if (predHome === match.homeScore && predAway === match.awayScore) {
        breakdown[pid].exactScore++;
        stageScores[pid][stage].exactScore++;
      } else if (Math.sign(predHome - predAway) === Math.sign(match.homeScore - match.awayScore)) {
        breakdown[pid].correctResult++;
        stageScores[pid][stage].correctResult++;
      } else {
        breakdown[pid].incorrect++;
        stageScores[pid][stage].incorrect++;
      }
    }

    // Cumulative score (original logic, unchanged — covers all stages)
    if (predHome === match.homeScore && predAway === match.awayScore) {
      scores[pid] += SCORING.CORRECT_SCORE * mult;
      stageScores[pid][stage].score += SCORING.CORRECT_SCORE * mult;
      if (!isKnockout) breakdown[pid].exactScore++;
      if (!isKnockout) stageScores[pid][stage].exactScore++;
    } else {
      const predResult = Math.sign(predHome - predAway);
      const realResult = Math.sign(match.homeScore - match.awayScore);
      if (predResult === realResult) {
        scores[pid] += SCORING.CORRECT_RESULT * mult;
        stageScores[pid][stage].score += SCORING.CORRECT_RESULT * mult;
        if (!isKnockout) breakdown[pid].correctResult++;
        if (!isKnockout) stageScores[pid][stage].correctResult++;
      } else {
        if (!isKnockout) breakdown[pid].incorrect++;
        if (!isKnockout) stageScores[pid][stage].incorrect++;
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
      knockoutScore: knockoutScores[id],
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
