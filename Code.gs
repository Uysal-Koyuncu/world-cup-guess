/**
 * World Cup Guess — Google Apps Script Backend
 * Deploy as Web App: Execute as "Me", Access "Anyone"
 *
 * Required Sheets (create these tabs in your Google Sheet):
 *   1. "Matches"    — columns: matchId | date | time | teamA | teamB | group | scoreA | scoreB | status | oddsA | oddsDraw | oddsB
 *   2. "Predictions" — columns: matchId | userName | predictedA | predictedB | betAmount | timestamp | settled
 *   3. "Users"      — columns: userName | balance
 */

const SHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function doGet(e) {
  const action = e.parameter.action;
  let result;
  try {
    switch (action) {
      case 'getMatches':
        result = getMatches();
        break;
      case 'getPredictions':
        result = getPredictions(e.parameter.user);
        break;
      case 'getLeaderboard':
        result = getLeaderboard();
        break;
      case 'getUsers':
        result = getUsers();
        break;
      default:
        result = { error: 'Unknown action' };
    }
  } catch (err) {
    result = { error: err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  let result;
  try {
    switch (data.action) {
      case 'submitPrediction':
        result = submitPrediction(data);
        break;
      case 'setResult':
        result = setResult(data);
        break;
      case 'addMatch':
        result = addMatch(data);
        break;
      case 'settleMatch':
        result = settleMatch(data.matchId);
        break;
      case 'addUser':
        result = addUser(data.userName, data.balance || 1000);
        break;
      default:
        result = { error: 'Unknown action' };
    }
  } catch (err) {
    result = { error: err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ── Matches ───────────────────────────────────────────── */

function getMatches() {
  const sheet = getSheet('Matches');
  const rows = sheet.getDataRange().getValues();
  const headers = rows.shift();
  return rows.map(r => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = r[i]);
    return obj;
  });
}

function addMatch(data) {
  const sheet = getSheet('Matches');
  sheet.appendRow([
    data.matchId, data.date, data.time,
    data.teamA, data.teamB, data.group,
    '', '', 'upcoming',
    data.oddsA || '', data.oddsDraw || '', data.oddsB || ''
  ]);
  return { success: true };
}

function setResult(data) {
  const sheet = getSheet('Matches');
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.matchId)) {
      sheet.getRange(i + 1, 7).setValue(data.scoreA);
      sheet.getRange(i + 1, 8).setValue(data.scoreB);
      sheet.getRange(i + 1, 9).setValue('finished');
      return { success: true };
    }
  }
  return { error: 'Match not found' };
}

/* ── Predictions ───────────────────────────────────────── */

function getPredictions(user) {
  const sheet = getSheet('Predictions');
  const rows = sheet.getDataRange().getValues();
  const headers = rows.shift();
  const preds = rows.map(r => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = r[i]);
    return obj;
  });
  if (user) return preds.filter(p => p.userName === user);
  return preds;
}

function submitPrediction(data) {
  const predSheet = getSheet('Predictions');
  const userSheet = getSheet('Users');

  const betAmount = Number(data.betAmount) || 100;

  // check balance
  const users = userSheet.getDataRange().getValues();
  let userRow = -1;
  for (let i = 1; i < users.length; i++) {
    if (users[i][0] === data.userName) { userRow = i + 1; break; }
  }
  if (userRow === -1) return { error: 'User not found' };

  const balance = Number(users[userRow - 1][1]);
  if (balance < betAmount) return { error: 'Insufficient balance', balance };

  // check for existing prediction
  const preds = predSheet.getDataRange().getValues();
  for (let i = 1; i < preds.length; i++) {
    if (String(preds[i][0]) === String(data.matchId) && preds[i][1] === data.userName) {
      return { error: 'Already predicted this match' };
    }
  }

  // deduct balance
  userSheet.getRange(userRow, 2).setValue(balance - betAmount);

  // save prediction
  predSheet.appendRow([
    data.matchId, data.userName,
    data.predictedA, data.predictedB,
    betAmount, new Date().toISOString(), 'pending'
  ]);

  return { success: true, newBalance: balance - betAmount };
}

/* ── Settlement ────────────────────────────────────────── */

function settleMatch(matchId) {
  const matchSheet = getSheet('Matches');
  const predSheet = getSheet('Predictions');
  const userSheet = getSheet('Users');

  // find match
  const matches = matchSheet.getDataRange().getValues();
  let match = null;
  for (let i = 1; i < matches.length; i++) {
    if (String(matches[i][0]) === String(matchId)) {
      match = {
        scoreA: Number(matches[i][6]),
        scoreB: Number(matches[i][7]),
        status: matches[i][8],
        oddsA: Number(matches[i][9]) || 2,
        oddsDraw: Number(matches[i][10]) || 3,
        oddsB: Number(matches[i][11]) || 2
      };
      break;
    }
  }
  if (!match || match.status !== 'finished') return { error: 'Match not finished' };

  const actualResult = match.scoreA > match.scoreB ? 'A' :
                       match.scoreA < match.scoreB ? 'B' : 'D';

  // settle predictions
  const preds = predSheet.getDataRange().getValues();
  const settled = [];
  for (let i = 1; i < preds.length; i++) {
    if (String(preds[i][0]) !== String(matchId)) continue;
    if (preds[i][6] !== 'pending') continue;

    const predA = Number(preds[i][2]);
    const predB = Number(preds[i][3]);
    const bet = Number(preds[i][4]);
    const userName = preds[i][1];
    const predResult = predA > predB ? 'A' : predA < predB ? 'B' : 'D';

    let payout = 0;
    let resultType = 'wrong';

    if (predA === match.scoreA && predB === match.scoreB) {
      // exact score — 3× odds
      const odds = actualResult === 'A' ? match.oddsA :
                   actualResult === 'D' ? match.oddsDraw : match.oddsB;
      payout = bet * odds * 3;
      resultType = 'exact';
    } else if (predResult === actualResult) {
      // correct result — 1× odds
      const odds = actualResult === 'A' ? match.oddsA :
                   actualResult === 'D' ? match.oddsDraw : match.oddsB;
      payout = bet * odds;
      resultType = 'result';
    }

    // update user balance
    const users = userSheet.getDataRange().getValues();
    for (let u = 1; u < users.length; u++) {
      if (users[u][0] === userName) {
        userSheet.getRange(u + 1, 2).setValue(Number(users[u][1]) + payout);
        break;
      }
    }

    // mark settled
    predSheet.getRange(i + 1, 7).setValue(resultType);
    settled.push({ userName, predA, predB, bet, payout, resultType });
  }

  return { success: true, settled };
}

/* ── Users & Leaderboard ───────────────────────────────── */

function getUsers() {
  const sheet = getSheet('Users');
  const rows = sheet.getDataRange().getValues();
  rows.shift();
  return rows.map(r => ({ userName: r[0], balance: Number(r[1]) }));
}

function addUser(userName, balance) {
  const sheet = getSheet('Users');
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === userName) return { error: 'User already exists' };
  }
  sheet.appendRow([userName, balance]);
  return { success: true };
}

function getLeaderboard() {
  const userSheet = getSheet('Users');
  const predSheet = getSheet('Predictions');

  const users = userSheet.getDataRange().getValues();
  users.shift();

  const preds = predSheet.getDataRange().getValues();
  preds.shift();

  return users.map(u => {
    const userName = u[0];
    const balance = Number(u[1]);
    const userPreds = preds.filter(p => p[1] === userName);
    const totalBets = userPreds.length;
    const exactWins = userPreds.filter(p => p[6] === 'exact').length;
    const resultWins = userPreds.filter(p => p[6] === 'result').length;
    const totalWagered = userPreds.reduce((s, p) => s + Number(p[4]), 0);
    return { userName, balance, totalBets, exactWins, resultWins, totalWagered };
  }).sort((a, b) => b.balance - a.balance);
}
