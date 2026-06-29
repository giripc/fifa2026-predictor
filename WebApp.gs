function doGet() {
  return HtmlService.createHtmlOutputFromFile('Predictor')
    .setTitle('⚽ FIFA 2026 Predictor')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── OTP Auth ─────────────────────────────────────────────────

var OTP_TTL     = 10 * 60;  // 10 minutes
var SESSION_TTL = 40 * 60;  // 40 minutes (matches client inactivity timeout)
var OTP_RATE    =  5 * 60;  // minimum seconds between OTP requests

function sendOtp(email) {
  if (!email || !email.match(/^[^\s@]+@[^\s@]+\.[^\s@]+$/))
    return { error: 'Please enter a valid email address.' };

  const key = email.toLowerCase();
  const cache = CacheService.getScriptCache();

  // Rate-limit: block re-requests within OTP_RATE seconds
  if (cache.get('otp_rl_' + key))
    return { error: 'OTP already sent. Please wait a few minutes before requesting again.' };

  const code = String(Math.floor(100000 + Math.random() * 900000));

  try {
    MailApp.sendEmail({
      to: email,
      subject: '⚽ FIFA 2026 Predictor — your sign-in code',
      body: 'Your one-time sign-in code is: ' + code + '\n\nThis code expires in 10 minutes.'
    });
  } catch (e) {
    return { error: 'Mail error: ' + e.message };
  }

  // Only set rate-limit and store code after successful send
  cache.put('otp_' + key, code, OTP_TTL);
  cache.put('otp_rl_' + key, '1', OTP_RATE);
  return { success: true };
}

function verifyOtp(email, code) {
  if (!email || !code) return { error: 'Email and code are required.' };
  const key = email.toLowerCase();
  const cache = CacheService.getScriptCache();
  const stored = cache.get('otp_' + key);
  if (!stored || stored !== code.trim())
    return { error: 'Invalid or expired code. Please request a new one.' };

  cache.remove('otp_' + key);
  const token = Utilities.getUuid();
  cache.put('sess_' + token, key, SESSION_TTL);
  return { success: true, token: token };
}

function _getSessionEmail(token) {
  if (!token) throw new Error('Session expired. Please sign in again.');
  const email = CacheService.getScriptCache().get('sess_' + token);
  if (!email) throw new Error('Session expired. Please sign in again.');
  return email;
}

function refreshSession(token) {
  if (!token) return { error: 'No session.' };
  const cache = CacheService.getScriptCache();
  const email = cache.get('sess_' + token);
  if (!email) return { expired: true };
  cache.put('sess_' + token, email, SESSION_TTL);
  return { success: true };
}

function lookupParticipant(token) {
  const email = _getSessionEmail(token);
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

function registerParticipant(name, inviteCode, token) {
  const email = _getSessionEmail(token);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const groupName = _resolveInviteCode(inviteCode);
  if (!groupName) return { error: 'Invalid invite code. Please check with your group admin.' };

  if (lookupParticipant(token)) return { error: 'Email already registered. Go back to sign in.' };

  const lock = LockService.getScriptLock();
  lock.tryLock(5000);
  try {
    if (lookupParticipant(token)) return { error: 'Email already registered. Go back to sign in.' };
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

function joinGroup(inviteCode, token) {
  const participant = lookupParticipant(token);
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
      const homePen = r[11] !== '' && r[11] !== undefined ? Number(r[11]) : null;
      const awayPen = r[12] !== '' && r[12] !== undefined ? Number(r[12]) : null;
      return {
        id: r[0], stage: r[1], date: r[4], group: r[3], home: r[5], away: r[6],
        homeScore: isCompleted || isLive ? r[7] : null,
        awayScore: isCompleted || isLive ? r[8] : null,
        homePenScore: isCompleted && homePen !== null && homePen > 0 ? homePen : null,
        awayPenScore: isCompleted && awayPen !== null && awayPen > 0 ? awayPen : null,
        status: isCompleted ? 'completed' : isLive ? 'live' : isUpcoming ? 'upcoming' : 'started'
      };
    });
  // 'started' matches (past kickoff but not yet marked Live/Completed) are shown locked
}

function getMyPredictions(token) {
  const participant = lookupParticipant(token);
  if (!participant) return {};
  const data = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(SHEETS.PREDICTIONS).getDataRange().getValues();
  const preds = {};
  data.slice(1).forEach(r => {
    if (r[1] === participant.id) preds[String(r[2])] = { home: r[3], away: r[4], penaltyWinner: r[6] || null };
  });
  return preds;
}

function submitPrediction(matchId, predHome, predAway, penaltyWinner, token) {
  const participant = lookupParticipant(token);
  if (!participant) return { error: 'Unauthorized. Please sign in.' };
  const participantId = participant.id;

  // Validate score values
  const h = parseInt(predHome), a = parseInt(predAway);
  if (isNaN(h) || isNaN(a) || h < 0 || h > 20 || a < 0 || a > 20)
    return { error: 'Invalid score values.' };

  // Validate penalty winner — only allowed if draw predicted on a knockout match
  const pen = (penaltyWinner === 'Home' || penaltyWinner === 'Away') ? penaltyWinner : null;

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
      sheet.getRange(i + 1, 4, 1, 4)
        .setValues([[h, a, new Date().toISOString(), pen || '']]);
      return { success: true };
    }
  }
  sheet.appendRow([Utilities.getUuid(), participantId, matchId,
    h, a, new Date().toISOString(), pen || '']);
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
function getLeaderboard(groupName, token) {
  const participant = lookupParticipant(token);
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
      const homePen = r[11] !== '' && r[11] !== undefined ? Number(r[11]) : null;
      const awayPen = r[12] !== '' && r[12] !== undefined ? Number(r[12]) : null;
      const wentToPenalties = homePen !== null && awayPen !== null && (homePen > 0 || awayPen > 0);
      matchMap[r[0]] = {
        homeScore:      Number(r[7]),
        awayScore:      Number(r[8]),
        stage:          r[1] || 'Other',
        group:          r[3],
        penaltyWinner:  wentToPenalties ? (homePen > awayPen ? 'Home' : 'Away') : null,
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

  // Proximity scoring for knockout rounds.
  // Penalty winner is part of outcome: predicted draw must have correct pen winner.
  function _koScore(predHome, predAway, actualHome, actualAway, predPenWinner, actualPenWinner) {
    const isExact = predHome === actualHome && predAway === actualAway;
    const predDraw = predHome === predAway;
    const actualDraw = actualHome === actualAway;

    if (isExact && actualDraw) {
      if (actualPenWinner) {
        return predPenWinner === actualPenWinner ? SCORING.KO_EXACT_PEN : SCORING.KO_WRONG_PEN;
      }
      return SCORING.KO_EXACT;
    }

    if (isExact) return SCORING.KO_EXACT;

    const predWinner = predDraw
      ? (actualPenWinner ? predPenWinner : 'draw')
      : (predHome > predAway ? 'Home' : 'Away');
    const actualWinner = actualDraw
      ? actualPenWinner
      : (actualHome > actualAway ? 'Home' : 'Away');

    if (predWinner !== actualWinner) return SCORING.KO_WRONG;

    const diff = Math.abs(predHome - actualHome) + Math.abs(predAway - actualAway);
    if (diff === 1) return SCORING.KO_1GOAL;
    if (diff === 2) return SCORING.KO_2GOAL;
    return SCORING.KO_FLOOR;
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
    const predPenWinner = r[6] || null;
    const isKnockout = !match.group;
    const stage = match.stage;
    _ensureStage(pid, stage);

    let pts;
    if (isKnockout) {
      pts = _koScore(predHome, predAway, match.homeScore, match.awayScore, predPenWinner, match.penaltyWinner);
    } else {
      if (predHome === match.homeScore && predAway === match.awayScore) {
        pts = SCORING.CORRECT_SCORE;
      } else {
        const predResult = Math.sign(predHome - predAway);
        const realResult = Math.sign(match.homeScore - match.awayScore);
        pts = predResult === realResult ? SCORING.CORRECT_RESULT : 0;
      }
    }

    scores[pid] += pts;
    stageScores[pid][stage].score += pts;

    if (predHome === match.homeScore && predAway === match.awayScore) {
      breakdown[pid].exactScore++;
      stageScores[pid][stage].exactScore++;
    } else if (pts > 0) {
      breakdown[pid].correctResult++;
      stageScores[pid][stage].correctResult++;
    } else {
      breakdown[pid].incorrect++;
      stageScores[pid][stage].incorrect++;
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
