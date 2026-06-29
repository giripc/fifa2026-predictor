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

  // Proximity scoring — Round of 32 and beyond only
  KO_EXACT_PEN:      8,   // exact score + correct penalty winner
  KO_EXACT:          6,   // exact score, no penalties
  KO_1GOAL:          3,   // correct outcome, 1 goal off
  KO_2GOAL:          2,   // correct outcome, 2 goals off
  KO_FLOOR:          1,   // correct outcome, 3+ goals off
  KO_WRONG_PEN:      3,   // exact draw score but wrong penalty winner
  KO_DRAW_WRONG_PEN: 1,   // non-exact draw predicted, wrong penalty winner
  KO_WRONG:         -1,   // wrong result
};

// FIFA public API — no key required
const FIFA = {
  BASE:        'https://api.fifa.com/api/v3',
  COMPETITION: 17,      // FIFA World Cup
  SEASON:      285023,  // 2026
};
