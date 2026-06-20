const SHEETS = {
  CONFIG:       'CONFIG',
  MATCHES:      'MATCHES',
  PARTICIPANTS: 'PARTICIPANTS',
  MEMBERSHIPS:  'MEMBERSHIPS',
  PREDICTIONS:  'PREDICTIONS',
  GROUPS:       'GROUPS',
};

// Scoring weights — tunable in CONFIG sheet later
const SCORING = {
  CORRECT_RESULT: 1,   // right W/D/L
  CORRECT_SCORE:  3,   // exact scoreline (includes result point)
  KNOCKOUT_MULT:  2,   // multiplier from Round of 32 onward
};

// FIFA public API — no key required
const FIFA = {
  BASE:        'https://api.fifa.com/api/v3',
  COMPETITION: 17,      // FIFA World Cup
  SEASON:      285023,  // 2026
};
