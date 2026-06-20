function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚽ FIFA Predictor')
    .addItem('1 · Setup Sheets',    'setupSheets')
    .addItem('2 · Fetch Matches',   'fetchMatches')
    .addSeparator()
    .addItem('Add Participant',     'showAddParticipantDialog')
    .addItem('Submit Prediction',   'showPredictionDialog')
    .addToUi();
}

// ── Participants ─────────────────────────────────────────────
function showAddParticipantDialog() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body{font-family:sans-serif;padding:16px;font-size:13px}
      label{display:block;margin-top:12px}
      input,select{width:100%;padding:6px;margin-top:3px;box-sizing:border-box}
      button{margin-top:16px;padding:8px 24px;background:#1a4e8c;color:#fff;
             border:none;border-radius:4px;cursor:pointer;font-size:13px}
    </style>
    <h3 style="margin-top:0">Add Participant</h3>
    <label>Name       <input id="n"/></label>
    <label>Email      <input id="e" type="email"/></label>
    <label>Pool Group <input id="g" placeholder="e.g. Office, Family"/></label>
    <button onclick="go()">Add</button>
    <script>
      function go(){
        google.script.run
          .withSuccessHandler(()=>google.script.host.close())
          .addParticipant(n.value, e.value, g.value);
      }
    </script>`)
    .setWidth(320).setHeight(270);
  SpreadsheetApp.getUi().showModalDialog(html, 'Add Participant');
}

function addParticipant(name, email, poolGroup) {
  SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(SHEETS.PARTICIPANTS)
    .appendRow([Utilities.getUuid(), name, email, poolGroup, new Date().toISOString()]);
}

// ── Predictions ──────────────────────────────────────────────
function showPredictionDialog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mSh = ss.getSheetByName(SHEETS.MATCHES);
  const pSh = ss.getSheetByName(SHEETS.PARTICIPANTS);
  if (!mSh || !pSh) {
    SpreadsheetApp.getUi().alert('Run Setup and Fetch Matches first.');
    return;
  }

  const upcoming = mSh.getDataRange().getValues().slice(1)
    .filter(r => r[9] === 'Upcoming')
    .map(r => `<option value="${r[0]}">${r[4].toString().substring(0,10)} · ${r[5]} vs ${r[6]}</option>`)
    .join('');

  const people = pSh.getDataRange().getValues().slice(1)
    .map(r => `<option value="${r[0]}">${r[1]} (${r[3]})</option>`)
    .join('');

  if (!upcoming) { SpreadsheetApp.getUi().alert('No upcoming matches found.'); return; }
  if (!people)   { SpreadsheetApp.getUi().alert('Add participants first.'); return; }

  const html = HtmlService.createHtmlOutput(`
    <style>
      body{font-family:sans-serif;padding:16px;font-size:13px}
      label{display:block;margin-top:12px}
      select,input{width:100%;padding:6px;margin-top:3px;box-sizing:border-box}
      .pair{display:flex;gap:8px}
      .pair input{width:50%}
      button{margin-top:16px;padding:8px 24px;background:#1a4e8c;color:#fff;
             border:none;border-radius:4px;cursor:pointer;font-size:13px}
    </style>
    <h3 style="margin-top:0">Submit Prediction</h3>
    <label>Participant  <select id="pid">${people}</select></label>
    <label>Match        <select id="mid">${upcoming}</select></label>
    <label>Predicted Score
      <div class="pair">
        <input id="h" type="number" min="0" placeholder="Home"/>
        <input id="a" type="number" min="0" placeholder="Away"/>
      </div>
    </label>
    <button onclick="go()">Save</button>
    <script>
      function go(){
        google.script.run
          .withSuccessHandler(()=>google.script.host.close())
          .savePrediction(pid.value, mid.value, +h.value, +a.value);
      }
    </script>`)
    .setWidth(360).setHeight(320);
  SpreadsheetApp.getUi().showModalDialog(html, 'Submit Prediction');
}

function savePrediction(participantId, matchId, predHome, predAway) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
                  .getSheetByName(SHEETS.PREDICTIONS);
  const data  = sheet.getDataRange().getValues();

  // Overwrite if prediction already exists for this person+match
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === participantId && data[i][2] === matchId) {
      sheet.getRange(i + 1, 4, 1, 3)
        .setValues([[predHome, predAway, new Date().toISOString()]]);
      return;
    }
  }
  sheet.appendRow([
    Utilities.getUuid(), participantId, matchId,
    predHome, predAway, new Date().toISOString()
  ]);
}
