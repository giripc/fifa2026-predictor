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
  CORRECT_RESULT: 1,   // right W/D/L (group stage)
  CORRECT_SCORE:  3,   // exact scoreline (group stage)
  KNOCKOUT_MULT:  2,   // multiplier from Round of 32 onward (applied to group-stage formula)

  // Proximity scoring — Round of 32 and beyond only
  KO_EXACT:       6,   // exact score
  KO_WRONG:      -1,   // wrong result
  KO_FLOOR:       1,   // minimum for correct result (however far off)
};

// FIFA public API — no key required
const FIFA = {
  BASE:        'https://api.fifa.com/api/v3',
  COMPETITION: 17,      // FIFA World Cup
  SEASON:      285023,  // 2026
};
