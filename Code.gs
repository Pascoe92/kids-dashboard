// =====================================================================
//  Kids Dashboard — Apps Script JSON API (v2)
//  Frontend hosted on GitHub Pages; this script is the data backend.
//  Google Sheet: https://docs.google.com/spreadsheets/d/1g1gohK6yKtyGPwdSd85tM-UvxhtaJsUhDGUOVq8RTfc
// =====================================================================

const SS_ID = '1g1gohK6yKtyGPwdSd85tM-UvxhtaJsUhDGUOVq8RTfc';

// ── Entry point — supports both JSON and JSONP (for CORS bypass) ────
function doGet(e) {
  const result = handleRequest(e);
  const json = JSON.stringify(result);
  const cb = (e.parameter || {}).callback;
  return ContentService
    .createTextOutput(cb ? `${cb}(${json})` : json)
    .setMimeType(cb ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}

function handleRequest(e) {
  const p = e.parameter || {};
  const action = p.action;
  try {
    switch (action) {
      case 'getData':      return getData();
      case 'logChore':     return logChore(p.kid, dec(p.chore));
      case 'addStrike':    return addStrike(p.kid, dec(p.reason || ''));
      case 'removeStrike': return removeStrike(p.kid);
      case 'awardBonus':   return awardBonus(p.kid, parseInt(p.points), dec(p.reason || ''));
      case 'redeemReward': return redeemReward(p.kid, dec(p.reward), parseInt(p.cost));
      case 'checkPin':     return checkPin(p.pin);
      case 'getActivity':  return { activity: getRecentActivity() };
      case 'initSheets':   return initializeSheets();
      default:             return { status: 'Kids Dashboard API v2', ok: true };
    }
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function dec(s) { return s ? decodeURIComponent(s) : ''; }

// ── One-time sheet initialisation ────────────────────────────────────
function initializeSheets() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const defs = {
    Config: [
      ['Setting', 'Value'],
      ['points_per_chore', 10],
      ['strike_limit', 5],
      ['strike_consequence', 'No ice cream this weekend! 🍦'],
      ['streak_bonus_days', 3],
      ['streak_bonus_points', 15],
      ['parent_pin', '1234'],
      ['last_strike_reset', new Date().toISOString()]
    ],
    KidsData: [
      ['Name', 'Age', 'Avatar', 'Points', 'Strikes', 'Streak', 'LastChoreDate', 'TotalEarned'],
      ['Myles',  6, '🦸', 0, 0, 0, '', 0],
      ['Finley', 3, '🐸', 0, 0, 0, '', 0]
    ],
    ChoresList: [
      ['Chore', 'Emoji', 'Kids'],
      ['Put shoes away',           '👟', 'Both'],
      ['Hang coat up',             '🧥', 'Both'],
      ['Make bed',                 '🛏️', 'Both'],
      ['Tidy toys',                '🧸', 'Both'],
      ['Dirty clothes in basket',  '🧺', 'Both'],
      ['Brush teeth (morning)',    '🪥', 'Both'],
      ['Brush teeth (night)',      '🪥', 'Both'],
      ['Wash hands before meals',  '🧼', 'Both'],
      ['Pack school bag',          '🎒', 'Myles'],
      ['Read for 10 minutes',      '📚', 'Myles'],
      ['Help set the table',       '🍽️', 'Myles'],
      ['Help tidy lounge',         '🛋️', 'Myles'],
      ['Pick up 5 toys',           '🪀', 'Finley'],
      ['Help with washing up',     '🫧', 'Finley']
    ],
    Rewards: [
      ['Reward', 'Points', 'Emoji', 'Active'],
      ['Extra screen time (30 min)', 50,  '📱', true],
      ["Choose what's for dinner",   75,  '🍕', true],
      ['Stay up 30 minutes later',   100, '🌙', true],
      ['Pick a movie or show',       120, '🎬', true],
      ['Trip to soft play',          200, '🏋️', true],
      ['New small toy or book',      300, '🎁', true],
      ['Cinema trip',                400, '🍿', true],
      ['Big adventure day out',      500, '🎢', true]
    ],
    ChoresLog:  [['Date', 'Kid', 'Chore', 'Points', 'Notes']],
    StrikesLog: [['Date', 'Kid', 'Reason']]
  };
  Object.entries(defs).forEach(([name, rows]) => {
    let sheet = ss.getSheetByName(name) || ss.insertSheet(name);
    sheet.clearContents();
    sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  });
  return { success: true, message: 'All sheets initialised!' };
}

// ── Main data payload ────────────────────────────────────────────────
function getData() {
  maybeResetStrikes_();
  const ss = SpreadsheetApp.openById(SS_ID);

  const kids = rows_(ss, 'KidsData', 'A2:H10').filter(r => r[0]).map(r => ({
    name: r[0], age: r[1], avatar: r[2], points: r[3]||0,
    strikes: r[4]||0,
    streak: typeof r[5] === 'number' ? Math.round(r[5]) : 0,
    lastChoreDate: r[6] ? new Date(r[6]).toISOString() : '',
    totalEarned: r[7]||0
  }));

  const chores = rows_(ss, 'ChoresList', 'A2:C30').filter(r => r[0])
    .map(r => ({ name: r[0], emoji: r[1], kids: r[2] }));

  const rewards = rows_(ss, 'Rewards', 'A2:D20').filter(r => r[0] && r[3])
    .map(r => ({ name: r[0], cost: r[1], emoji: r[2] }));

  const config = {};
  rows_(ss, 'Config', 'A2:B20').filter(r => r[0]).forEach(r => { config[r[0]] = r[1]; });
  // Send PIN as simple XOR so it's not plaintext but still usable client-side
  const rawPin = String(config.parent_pin || '1234');
  config.pin_check = rawPin.split('').map((c,i)=>c.charCodeAt(0)^(42+i)).join(',');
  delete config.parent_pin;

  // Today's completed chores per kid
  const today = new Date().toDateString();
  const logSheet = ss.getSheetByName('ChoresLog');
  const lastRow = logSheet.getLastRow();
  const completedToday = {};
  if (lastRow > 1) {
    logSheet.getRange('A2:C' + lastRow).getValues()
      .filter(r => r[0] && new Date(r[0]).toDateString() === today)
      .forEach(r => { completedToday[r[1] + '|' + r[2]] = true; });
  }

  return { success: true, kids, chores, rewards, config, completedToday };
}

// ── Log chore ────────────────────────────────────────────────────────
function logChore(kidName, choreName) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const config = config_(ss);
  const pts = parseInt(config.points_per_chore) || 10;

  // Duplicate check
  const today = new Date().toDateString();
  const logSheet = ss.getSheetByName('ChoresLog');
  const lastRow = logSheet.getLastRow();
  if (lastRow > 1) {
    const dup = logSheet.getRange('A2:C' + lastRow).getValues()
      .some(r => r[0] && new Date(r[0]).toDateString() === today && r[1] === kidName && r[2] === choreName);
    if (dup) return { success: false, message: 'Already done today!' };
  }
  logSheet.appendRow([new Date(), kidName, choreName, pts, '']);
  return updatePoints_(ss, kidName, pts, config, false);
}

// ── Strike management ────────────────────────────────────────────────
function addStrike(kidName, reason) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const config = config_(ss);
  const limit = parseInt(config.strike_limit) || 5;
  ss.getSheetByName('StrikesLog').appendRow([new Date(), kidName, reason || 'Not following the rules']);
  const kidsSheet = ss.getSheetByName('KidsData');
  const data = kidsSheet.getRange('A2:H10').getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] !== kidName) continue;
    const newStrikes = (data[i][4] || 0) + 1;
    kidsSheet.getRange(i + 2, 5).setValue(newStrikes);
    return { success: true, strikes: newStrikes, hitLimit: newStrikes >= limit, consequence: config.strike_consequence };
  }
  return { success: false };
}

function removeStrike(kidName) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const kidsSheet = ss.getSheetByName('KidsData');
  const data = kidsSheet.getRange('A2:H10').getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] !== kidName) continue;
    const newStrikes = Math.max(0, (data[i][4] || 0) - 1);
    kidsSheet.getRange(i + 2, 5).setValue(newStrikes);
    return { success: true, strikes: newStrikes };
  }
  return { success: false };
}

// ── Bonus points ─────────────────────────────────────────────────────
function awardBonus(kidName, points, reason) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const config = config_(ss);
  ss.getSheetByName('ChoresLog').appendRow([new Date(), kidName, '⭐ BONUS: ' + reason, points, 'Bonus']);
  return updatePoints_(ss, kidName, parseInt(points), config, true);
}

// ── Redeem reward ────────────────────────────────────────────────────
function redeemReward(kidName, rewardName, cost) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const kidsSheet = ss.getSheetByName('KidsData');
  const data = kidsSheet.getRange('A2:H10').getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] !== kidName) continue;
    const current = data[i][3] || 0;
    if (current < cost) return { success: false, message: 'Not enough points! Keep going! 💪' };
    kidsSheet.getRange(i + 2, 4).setValue(current - cost);
    ss.getSheetByName('ChoresLog').appendRow([new Date(), kidName, '🎁 REWARD: ' + rewardName, -cost, 'Redeemed']);
    return { success: true, newTotal: current - cost };
  }
  return { success: false };
}

// ── PIN check ────────────────────────────────────────────────────────
function checkPin(pin) {
  const stored = config_(SpreadsheetApp.openById(SS_ID)).parent_pin || '1234';
  return { valid: String(pin) === String(stored) };
}

// ── Activity feed ─────────────────────────────────────────────────────
function getRecentActivity() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('ChoresLog');
  const last = sheet.getLastRow();
  if (last <= 1) return [];
  const start = Math.max(2, last - 29);
  return sheet.getRange('A' + start + ':E' + last).getValues()
    .filter(r => r[0]).reverse()
    .map(r => ({
      date: Utilities.formatDate(new Date(r[0]), Session.getScriptTimeZone(), 'dd MMM HH:mm'),
      kid: r[1], chore: r[2], points: r[3]
    }));
}

// ── Private helpers ───────────────────────────────────────────────────
function rows_(ss, sheet, range) {
  const s = ss.getSheetByName(sheet);
  return s ? s.getRange(range).getValues() : [];
}

function config_(ss) {
  const cfg = {};
  rows_(ss, 'Config', 'A2:B20').filter(r => r[0]).forEach(r => { cfg[r[0]] = r[1]; });
  return cfg;
}

function updatePoints_(ss, kidName, pts, config, isBonus) {
  const kidsSheet = ss.getSheetByName('KidsData');
  const data = kidsSheet.getRange('A2:H10').getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] !== kidName) continue;
    const row = i + 2;
    const currentPts  = data[i][3] || 0;
    const currentStrk = data[i][5] || 0;
    const lastDate    = data[i][6];
    const totalEarned = data[i][7] || 0;

    let newStreak = currentStrk;
    if (!isBonus) {
      const todayStr = new Date().toDateString();
      const lastStr  = lastDate ? new Date(lastDate).toDateString() : '';
      if (lastStr !== todayStr) {
        const yesterday = new Date(); yesterday.setDate(yesterday.getDate() - 1);
        newStreak = (lastStr === yesterday.toDateString()) ? currentStrk + 1 : 1;
      }
    }

    const bonusDays = parseInt(config.streak_bonus_days)   || 3;
    const bonusPts  = parseInt(config.streak_bonus_points) || 15;
    const streakBonus = (!isBonus && newStreak > 0 && newStreak % bonusDays === 0) ? bonusPts : 0;
    const earned = pts + streakBonus;

    kidsSheet.getRange(row, 4).setValue(currentPts + earned);   // Points
    kidsSheet.getRange(row, 6).setValue(newStreak);             // Streak
    kidsSheet.getRange(row, 7).setValue(new Date());            // LastChoreDate
    kidsSheet.getRange(row, 8).setValue(totalEarned + earned);  // TotalEarned

    return { success: true, points: pts, bonusPoints: streakBonus,
             newTotal: currentPts + earned, streak: newStreak, streakBonus: streakBonus > 0 };
  }
  return { success: false };
}

function maybeResetStrikes_() {
  if (new Date().getDay() !== 1) return; // not Monday
  const ss = SpreadsheetApp.openById(SS_ID);
  const configSheet = ss.getSheetByName('Config');
  const data = configSheet.getRange('A2:B20').getValues();
  let resetRow = -1, lastReset = null;
  data.forEach((r, i) => {
    if (r[0] === 'last_strike_reset') { resetRow = i + 2; lastReset = r[1] ? new Date(r[1]) : null; }
  });
  const today = new Date().toDateString();
  if (lastReset && lastReset.toDateString() === today) return;
  ss.getSheetByName('KidsData').getRange('A2:H10').getValues()
    .forEach((r, i) => { if (r[0]) ss.getSheetByName('KidsData').getRange(i + 2, 5).setValue(0); });
  if (resetRow > 0) configSheet.getRange(resetRow, 2).setValue(new Date().toISOString());
}
