// =====================================================================
//  Kids Behaviour Dashboard — Apps Script Backend
//  Google Sheet: https://docs.google.com/spreadsheets/d/1g1gohK6yKtyGPwdSd85tM-UvxhtaJsUhDGUOVq8RTfc
// =====================================================================

const SS_ID = '1g1gohK6yKtyGPwdSd85tM-UvxhtaJsUhDGUOVq8RTfc';

// ── Entry point ─────────────────────────────────────────────────────
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('⭐ Kids Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── One-time setup: creates all required sheets ──────────────────────
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
      ['Put shoes away',            '👟', 'Both'],
      ['Hang coat up',              '🧥', 'Both'],
      ['Make bed',                  '🛏️', 'Both'],
      ['Tidy toys',                 '🧸', 'Both'],
      ['Dirty clothes in basket',   '🧺', 'Both'],
      ['Brush teeth (morning)',      '🪥', 'Both'],
      ['Brush teeth (night)',        '🪥', 'Both'],
      ['Wash hands before meals',   '🧼', 'Both'],
      ['Pack school bag',           '🎒', 'Myles'],
      ['Read for 10 minutes',       '📚', 'Myles'],
      ['Help set the table',        '🍽️', 'Myles'],
      ['Help tidy lounge',          '🛋️', 'Myles'],
      ['Pick up 5 toys',            '🪀', 'Finley'],
      ['Help with washing up',      '🫧', 'Finley']
    ],
    Rewards: [
      ['Reward', 'Points', 'Emoji', 'Active'],
      ['Extra screen time (30 min)', 50,  '📱', true],
      ['Choose what\'s for dinner',  75,  '🍕', true],
      ['Stay up 30 minutes later',   100, '🌙', true],
      ['Pick a movie/show',          120, '🎬', true],
      ['Trip to soft play',          200, '🏋️', true],
      ['New small toy or book',      300, '🎁', true],
      ['Cinema trip',                400, '🍿', true],
      ['Big adventure day out',      500, '🎢', true]
    ],
    ChoresLog:  [['Date', 'Kid', 'Chore', 'Points', 'Notes']],
    StrikesLog: [['Date', 'Kid', 'Reason']]
  };

  Object.entries(defs).forEach(([name, rows]) => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    else sheet.clearContents();
    sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  });

  return { success: true, message: 'All sheets initialised!' };
}

// ── Main data fetch ──────────────────────────────────────────────────
function getData() {
  maybeResetStrikes_();
  const ss = SpreadsheetApp.openById(SS_ID);

  const kids = readRows_(ss, 'KidsData', 'A2:H10')
    .filter(r => r[0])
    .map(r => ({
      name: r[0], age: r[1], avatar: r[2],
      points: r[3] || 0, strikes: r[4] || 0,
      streak: r[5] || 0, lastChoreDate: r[6] || '',
      totalEarned: r[7] || 0
    }));

  const chores = readRows_(ss, 'ChoresList', 'A2:C30')
    .filter(r => r[0])
    .map(r => ({ name: r[0], emoji: r[1], kids: r[2] }));

  const rewards = readRows_(ss, 'Rewards', 'A2:D20')
    .filter(r => r[0] && r[3])
    .map(r => ({ name: r[0], cost: r[1], emoji: r[2] }));

  const config = {};
  readRows_(ss, 'Config', 'A2:B20')
    .filter(r => r[0])
    .forEach(r => { config[r[0]] = r[1]; });

  // Which chores has each kid already completed today?
  const today = new Date().toDateString();
  const logSheet = ss.getSheetByName('ChoresLog');
  const lastRow = logSheet.getLastRow();
  const completedToday = {};
  if (lastRow > 1) {
    const logData = logSheet.getRange('A2:C' + lastRow).getValues();
    logData
      .filter(r => r[0] && new Date(r[0]).toDateString() === today)
      .forEach(r => { completedToday[r[1] + '|' + r[2]] = true; });
  }

  return { kids, chores, rewards, config, completedToday };
}

// ── Log a chore completion ───────────────────────────────────────────
function logChore(kidName, choreName) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const config = getConfig_(ss);
  const pts = parseInt(config.points_per_chore) || 10;

  // Duplicate check
  const today = new Date().toDateString();
  const logSheet = ss.getSheetByName('ChoresLog');
  const lastRow = logSheet.getLastRow();
  if (lastRow > 1) {
    const already = logSheet.getRange('A2:C' + lastRow).getValues()
      .some(r => r[0] && new Date(r[0]).toDateString() === today
                      && r[1] === kidName && r[2] === choreName);
    if (already) return { success: false, message: 'Already done today! ✅' };
  }

  logSheet.appendRow([new Date(), kidName, choreName, pts, '']);
  return updateKidPoints_(ss, kidName, pts, config);
}

// ── Add a strike ─────────────────────────────────────────────────────
function addStrike(kidName, reason) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const config = getConfig_(ss);
  const limit = parseInt(config.strike_limit) || 5;

  ss.getSheetByName('StrikesLog').appendRow([new Date(), kidName, reason || 'Not following the rules']);

  const kidsSheet = ss.getSheetByName('KidsData');
  const rows = kidsSheet.getRange('A2:H10').getValues();
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0] === kidName) {
      const newStrikes = (rows[i][4] || 0) + 1;
      kidsSheet.getRange(i + 2, 5).setValue(newStrikes);
      return {
        success: true, strikes: newStrikes,
        hitLimit: newStrikes >= limit,
        consequence: config.strike_consequence
      };
    }
  }
  return { success: false };
}

// ── Remove a strike (parent undo) ────────────────────────────────────
function removeStrike(kidName) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const kidsSheet = ss.getSheetByName('KidsData');
  const rows = kidsSheet.getRange('A2:H10').getValues();
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0] === kidName) {
      const newStrikes = Math.max(0, (rows[i][4] || 0) - 1);
      kidsSheet.getRange(i + 2, 5).setValue(newStrikes);
      return { success: true, strikes: newStrikes };
    }
  }
  return { success: false };
}

// ── Award bonus points ───────────────────────────────────────────────
function awardBonus(kidName, points, reason) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const config = getConfig_(ss);
  ss.getSheetByName('ChoresLog')
    .appendRow([new Date(), kidName, '⭐ BONUS: ' + reason, points, 'Bonus']);
  return updateKidPoints_(ss, kidName, parseInt(points), config, true);
}

// ── Redeem a reward ──────────────────────────────────────────────────
function redeemReward(kidName, rewardName, cost) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const kidsSheet = ss.getSheetByName('KidsData');
  const rows = kidsSheet.getRange('A2:H10').getValues();
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0] === kidName) {
      const current = rows[i][3] || 0;
      if (current < cost) return { success: false, message: 'Not enough points! Keep going! 💪' };
      kidsSheet.getRange(i + 2, 4).setValue(current - cost);
      ss.getSheetByName('ChoresLog')
        .appendRow([new Date(), kidName, '🎁 REWARD: ' + rewardName, -cost, 'Redeemed']);
      return { success: true, newTotal: current - cost };
    }
  }
  return { success: false };
}

// ── Verify parent PIN ────────────────────────────────────────────────
function checkPin(pin) {
  const config = getConfig_(SpreadsheetApp.openById(SS_ID));
  return { valid: String(pin) === String(config.parent_pin || '1234') };
}

// ── Recent activity feed ─────────────────────────────────────────────
function getRecentActivity() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('ChoresLog');
  const last = sheet.getLastRow();
  if (last <= 1) return [];
  const start = Math.max(2, last - 24);
  return sheet.getRange('A' + start + ':E' + last).getValues()
    .filter(r => r[0])
    .reverse()
    .map(r => ({
      date: Utilities.formatDate(new Date(r[0]), Session.getScriptTimeZone(), 'dd MMM HH:mm'),
      kid: r[1], chore: r[2], points: r[3]
    }));
}

// ── Private helpers ──────────────────────────────────────────────────
function readRows_(ss, sheetName, range) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  return sheet.getRange(range).getValues();
}

function getConfig_(ss) {
  const config = {};
  readRows_(ss, 'Config', 'A2:B20')
    .filter(r => r[0])
    .forEach(r => { config[r[0]] = r[1]; });
  return config;
}

function updateKidPoints_(ss, kidName, pts, config, isBonus) {
  const kidsSheet = ss.getSheetByName('KidsData');
  const rows = kidsSheet.getRange('A2:H10').getValues();
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0] !== kidName) continue;
    const row = i + 2;
    const currentPts = rows[i][3] || 0;
    const currentStreak = rows[i][5] || 0;
    const lastDate = rows[i][6];
    const totalEarned = rows[i][7] || 0;

    let newStreak = currentStreak;
    if (!isBonus) {
      const todayStr = new Date().toDateString();
      const lastStr = lastDate ? new Date(lastDate).toDateString() : '';
      if (lastStr !== todayStr) {
        const yesterday = new Date();
        yesterday.setDate(yesterday.getDate() - 1);
        newStreak = lastStr === yesterday.toDateString() ? currentStreak + 1 : 1;
        kidsSheet.getRange(row, 6).setValue(new Date());
      }
    }

    const bonusDays = parseInt(config.streak_bonus_days) || 3;
    const bonusPts  = parseInt(config.streak_bonus_points) || 15;
    const streakBonus = (!isBonus && newStreak > 0 && newStreak % bonusDays === 0) ? bonusPts : 0;
    const earned = pts + streakBonus;

    kidsSheet.getRange(row, 4).setValue(currentPts + earned);
    kidsSheet.getRange(row, 6).setValue(new Date());
    kidsSheet.getRange(row, 7).setValue(totalEarned + earned);
    if (!isBonus) kidsSheet.getRange(row, 6).setValue(newStreak); // streak col is 6 (F)

    // Fix: streak is col 6 (index 5), lastChoreDate is col 7 (index 6)
    kidsSheet.getRange(row, 6).setValue(newStreak);
    kidsSheet.getRange(row, 7).setValue(new Date());
    kidsSheet.getRange(row, 4).setValue(currentPts + earned);
    kidsSheet.getRange(row, 8).setValue(totalEarned + earned);

    return {
      success: true, points: pts, bonusPoints: streakBonus,
      newTotal: currentPts + earned, streak: newStreak,
      streakBonus: streakBonus > 0
    };
  }
  return { success: false };
}

function maybeResetStrikes_() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const configSheet = ss.getSheetByName('Config');
  const rows = configSheet.getRange('A2:B20').getValues();

  let lastResetRow = -1, lastReset = null;
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0] === 'last_strike_reset') {
      lastResetRow = i + 2;
      lastReset = rows[i][1] ? new Date(rows[i][1]) : null;
      break;
    }
  }

  const now = new Date();
  if (now.getDay() !== 1) return; // not Monday
  if (lastReset && lastReset.toDateString() === now.toDateString()) return; // already reset today

  const kidsSheet = ss.getSheetByName('KidsData');
  const kidsData = kidsSheet.getRange('A2:H10').getValues();
  kidsData.forEach((r, i) => { if (r[0]) kidsSheet.getRange(i + 2, 5).setValue(0); });

  if (lastResetRow > 0) configSheet.getRange(lastResetRow, 2).setValue(now.toISOString());
}
