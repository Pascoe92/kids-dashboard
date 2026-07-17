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
      case 'getAll':       return getAll();
      case 'logChore':     return logChore(p.kid, dec(p.chore));
      case 'undoChore':    return undoChore(p.kid, dec(p.chore));
      case 'addStrike':    return addStrike(p.kid, dec(p.reason || ''));
      case 'removeStrike': return removeStrike(p.kid);
      case 'awardBonus':   return awardBonus(p.kid, parseInt(p.points), dec(p.reason || ''));
      case 'redeemReward': return redeemReward(p.kid, dec(p.reward), parseInt(p.cost));
      case 'checkPin':     return checkPin(p.pin);
      case 'getActivity':       return { activity: getRecentActivity() };
      case 'getRecentBonuses':  return { bonuses: getRecentBonuses(parseInt(p.limit) || 8) };
      case 'initSheets':        return initializeSheets();
      // Calendar
      case 'getCalendarList': return getCalendarList();
      case 'getCalEvents':    return getCalEvents(p.calId||'');
      case 'addCalEvent':     return addCalEvent(p.id, dec(p.title), p.date, p.time||'', p.color||'gold', dec(p.notes||''), p.endTime||'', p.calId||'');
      case 'deleteCalEvent':  return deleteCalEvent(p.id);
      // Meal Planner
      case 'getMealPlan':       return getMealPlan(p.week);
      case 'saveMeal':          return saveMeal(p.week, p.day, p.type, dec(p.text));
      // Shopping List
      case 'getShoppingList':   return getShoppingList();
      case 'addShopItem':       return addShopItem(dec(p.item), dec(p.cat));
      case 'toggleShopItem':    return toggleShopItem(p.id);
      case 'deleteShopItem':    return deleteShopItem(p.id);
      case 'clearDoneShop':     return clearDoneShop();
      case 'clearAllShop':      return clearAllShop();
      // Family Notes
      case 'getNotes':          return getNotes();
      case 'saveNote':          return saveNote(p.id, dec(p.text), dec(p.color));
      case 'deleteNote':        return deleteNote(p.id);
      default:                  return { status: 'Kids Dashboard API v2', ok: true };
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
      ['Chore', 'Emoji', 'Kids', 'Period', 'Days'],
      ['Make bed',                 '🛏️', 'Both',   'Morning',      ''],
      ['Brush teeth (morning)',    '🪥', 'Both',   'Morning',      ''],
      ['Pack school bag',          '🎒', 'Myles',  'Morning',      ''],
      ['Put shoes away',           '👟', 'Both',   'After School', ''],
      ['Hang coat up',             '🧥', 'Both',   'After School', ''],
      ['Tidy toys',                '🧸', 'Both',   'After School', ''],
      ['Wash hands before meals',  '🧼', 'Both',   'After School', ''],
      ['Help set the table',       '🍽️', 'Myles',  'After School', ''],
      ['Pick up 5 toys',           '🪀', 'Finley', 'After School', ''],
      ['Dirty clothes in basket',  '🧺', 'Both',   'Bedtime',      ''],
      ['Help tidy lounge',         '🛋️', 'Myles',  'Bedtime',      ''],
      ['Help with washing up',     '🫧', 'Finley', 'Bedtime',      ''],
      ['Brush teeth (night)',      '🪥', 'Both',   'Bedtime',      ''],
      ['Read for 10 minutes',      '📚', 'Myles',  'Bonus',        '']
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
    ChoresLog:   [['Date', 'Kid', 'Chore', 'Points', 'Notes']],
    StrikesLog:  [['Date', 'Kid', 'Reason']],
    CalendarEvents:[['ID', 'Title', 'Date', 'Time', 'Color', 'Notes']],
    MealPlan:    [['WeekStart', 'Day', 'MealType', 'Text']],
    ShoppingList:[['ID', 'Item', 'Category', 'Checked']],
    FamilyNotes: [['ID', 'Text', 'Color', 'Created']]
  };
  Object.entries(defs).forEach(([name, rows]) => {
    let sheet = ss.getSheetByName(name) || ss.insertSheet(name);
    sheet.clearContents();
    sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  });
  return { success: true, message: 'All sheets initialised!' };
}

// ── Combined startup payload (home + kids + shopping count + meal) ───
function getAll() {
  const base = getData();
  base.activity = getRecentActivity();
  // Shopping count
  const ss = SpreadsheetApp.openById(SS_ID);
  const shopSheet = ss.getSheetByName('ShoppingList');
  const shopLast = shopSheet.getLastRow();
  let unchecked = 0;
  if (shopLast > 1) {
    unchecked = shopSheet.getRange('A2:D' + shopLast).getValues()
      .filter(r => r[0] && !(r[3] === true || r[3] === 'TRUE')).length;
  }
  base.shopCount = unchecked;
  // Today's dinner
  const wk = Utilities.formatDate(getMondayDate_(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const todayIdx = new Date().getDay() === 0 ? 6 : new Date().getDay() - 1;
  const mealSheet = ss.getSheetByName('MealPlan');
  const mealLast = mealSheet.getLastRow();
  let dinner = '';
  if (mealLast > 1) {
    const meals = mealSheet.getRange('A2:D' + mealLast).getValues();
    const row = meals.find(r => r[0] === wk && String(r[1]) === String(todayIdx) && r[2] === 'dinner');
    if (row) dinner = row[3];
  }
  base.todayDinner = dinner;
  return base;
}

function getMondayDate_() {
  const d = new Date();
  const day = d.getDay();
  const diff = day === 0 ? -6 : 1 - day;
  d.setDate(d.getDate() + diff);
  d.setHours(0,0,0,0);
  return d;
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

  const chores = rows_(ss, 'ChoresList', 'A2:E30').filter(r => r[0])
    .map(r => ({ name: r[0], emoji: r[1], kids: r[2], period: r[3] || '', days: r[4] || '' }));

  const rewards = rows_(ss, 'Rewards', 'A2:D20').filter(r => r[0] && r[3])
    .map(r => ({ name: r[0], cost: r[1], emoji: r[2] }));

  const config = {};
  rows_(ss, 'Config', 'A2:B20').filter(r => r[0]).forEach(r => { config[r[0]] = r[1]; });
  config.pin_key = String(config.parent_pin || '1234');
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

// ── Undo a chore ─────────────────────────────────────────────────────
function undoChore(kidName, choreName) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const today = new Date().toDateString();
  const logSheet = ss.getSheetByName('ChoresLog');
  const lastRow = logSheet.getLastRow();
  if (lastRow <= 1) return { success: false, message: 'Nothing to undo' };

  const data = logSheet.getRange('A2:D' + lastRow).getValues();
  // Find most recent matching entry today
  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i][0] && new Date(data[i][0]).toDateString() === today
        && data[i][1] === kidName && data[i][2] === choreName) {
      const pts = data[i][3] || 0;
      logSheet.deleteRow(i + 2);
      // Deduct points
      const kidsSheet = ss.getSheetByName('KidsData');
      const kidsData = kidsSheet.getRange('A2:H10').getValues();
      for (let j = 0; j < kidsData.length; j++) {
        if (kidsData[j][0] === kidName) {
          const newPts = Math.max(0, (kidsData[j][3] || 0) - pts);
          const newTotal = Math.max(0, (kidsData[j][7] || 0) - pts);
          kidsSheet.getRange(j + 2, 4).setValue(newPts);
          kidsSheet.getRange(j + 2, 8).setValue(newTotal);
          return { success: true, newTotal: newPts, pointsRemoved: pts };
        }
      }
    }
  }
  return { success: false, message: 'Chore not found in today\'s log' };
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

// Bonus-only feed, scanning the whole log (not just the tail window
// getRecentActivity uses) since bonuses are rare and would otherwise get
// crowded out by ordinary chore check-offs.
function getRecentBonuses(limit) {
  limit = limit || 8;
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('ChoresLog');
  const last = sheet.getLastRow();
  if (last <= 1) return [];
  return sheet.getRange('A2:E' + last).getValues()
    .filter(r => r[0] && r[4] === 'Bonus')
    .slice(-limit).reverse()
    .map(r => ({
      date: Utilities.formatDate(new Date(r[0]), Session.getScriptTimeZone(), 'dd MMM HH:mm'),
      kid: r[1], reason: String(r[2]).replace(/^⭐ BONUS: /, ''), points: r[3]
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

// ── Calendar (Google Calendar API via CalendarApp) ────────────────────
function getCalendarList() {
  try {
    const cals = CalendarApp.getAllCalendars().map(c => ({
      id: c.getId(),
      name: c.getName(),
      color: c.getColor(),
      isDefault: c.isMyPrimaryCalendar()
    }));
    return { success: true, calendars: cals };
  } catch(err) {
    return { success: false, error: err.message };
  }
}
function getCalEvents(calId) {
  try {
    const cal = (calId && calId !== 'default') ? CalendarApp.getCalendarById(calId) : CalendarApp.getDefaultCalendar();
    if (!cal) return { success: false, error: 'Calendar not found' };
    const now = new Date();
    const start = new Date(now.getFullYear(), now.getMonth() - 1, 1); // prev month
    const end   = new Date(now.getFullYear(), now.getMonth() + 3, 0); // 2 months ahead
    const events = cal.getEvents(start, end).map(e => ({
      id: e.getId(),
      title: e.getTitle(),
      date: Utilities.formatDate(e.getStartTime(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      time: e.isAllDayEvent() ? '' : Utilities.formatDate(e.getStartTime(), Session.getScriptTimeZone(), 'HH:mm'),
      endTime: e.isAllDayEvent() ? '' : Utilities.formatDate(e.getEndTime(), Session.getScriptTimeZone(), 'HH:mm'),
      allDay: e.isAllDayEvent(),
      notes: e.getDescription() || '',
      color: gcalColorToApp_(e.getColor())
    }));
    return { success: true, events, calName: cal.getName() };
  } catch(err) {
    return { success: false, error: err.message };
  }
}

function addCalEvent(id, title, date, time, color, notes, endTime, calId) {
  try {
    const cal = (calId && calId !== 'default') ? CalendarApp.getCalendarById(calId) : CalendarApp.getDefaultCalendar();
    if (!cal) return { success: false, error: 'Calendar not found' };
    const parts = date.split('-').map(Number);
    let event;
    if (time) {
      const timeParts = time.split(':').map(Number);
      const start = new Date(parts[0], parts[1]-1, parts[2], timeParts[0], timeParts[1]);
      const end = endTime
        ? (() => { const ep = endTime.split(':').map(Number); return new Date(parts[0], parts[1]-1, parts[2], ep[0], ep[1]); })()
        : new Date(start.getTime() + 60*60*1000);
      event = cal.createEvent(title, start, end, { description: notes || '' });
    } else {
      event = cal.createAllDayEvent(title, new Date(parts[0], parts[1]-1, parts[2]), { description: notes || '' });
    }
    const gcalColor = appColorToGcal_(color);
    if (gcalColor) event.setColor(gcalColor);
    return { success: true, id: event.getId() };
  } catch(err) {
    return { success: false, error: err.message };
  }
}

function deleteCalEvent(id) {
  try {
    const events = CalendarApp.getDefaultCalendar().getEventById(id)
      || CalendarApp.getEventById(id);
    if (events) { events.deleteEvent(); return { success: true }; }
    // Try searching across all calendars
    const allCals = CalendarApp.getAllCalendars();
    for (const cal of allCals) {
      const e = cal.getEventById(id);
      if (e) { e.deleteEvent(); return { success: true }; }
    }
    return { success: false, error: 'Event not found' };
  } catch(err) {
    return { success: false, error: err.message };
  }
}

function gcalColorToApp_(color) {
  const map = {
    [CalendarApp.EventColor.YELLOW]: 'gold',
    [CalendarApp.EventColor.GREEN]:  'green',
    [CalendarApp.EventColor.BLUE]:   'blue',
    [CalendarApp.EventColor.RED]:    'red',
    [CalendarApp.EventColor.PURPLE]: 'purple',
    [CalendarApp.EventColor.CYAN]:   'blue',
    [CalendarApp.EventColor.TEAL]:   'green',
    [CalendarApp.EventColor.PINK]:   'red',
    [CalendarApp.EventColor.GRAY]:   'purple',
  };
  return map[color] || 'blue';
}

function appColorToGcal_(color) {
  const map = {
    gold:   CalendarApp.EventColor.YELLOW,
    green:  CalendarApp.EventColor.GREEN,
    blue:   CalendarApp.EventColor.BLUE,
    red:    CalendarApp.EventColor.RED,
    purple: CalendarApp.EventColor.PURPLE,
  };
  return map[color] || null;
}

// ── Meal Planner ──────────────────────────────────────────────────────
function getMealPlan(week) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('MealPlan');
  const last = sheet.getLastRow();
  const meals = {};
  if (last > 1) {
    sheet.getRange('A2:D' + last).getValues()
      .filter(r => r[0] === week)
      .forEach(r => { meals[r[1] + '_' + r[2]] = r[3]; }); // key: "dayIndex_mealType"
  }
  return { success: true, week, meals };
}

function saveMeal(week, day, type, text) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('MealPlan');
  const last = sheet.getLastRow();
  // Update existing row or append
  if (last > 1) {
    const data = sheet.getRange('A2:D' + last).getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === week && String(data[i][1]) === String(day) && data[i][2] === type) {
        if (text) {
          sheet.getRange(i + 2, 4).setValue(text);
        } else {
          sheet.deleteRow(i + 2);
        }
        return { success: true };
      }
    }
  }
  if (text) sheet.appendRow([week, day, type, text]);
  return { success: true };
}

// ── Shopping List ─────────────────────────────────────────────────────
function getShoppingList() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('ShoppingList');
  const last = sheet.getLastRow();
  if (last <= 1) return { success: true, items: [] };
  const items = sheet.getRange('A2:D' + last).getValues()
    .filter(r => r[0])
    .map(r => ({ id: String(r[0]), name: r[1], cat: r[2], checked: r[3] === true || r[3] === 'TRUE' }));
  return { success: true, items };
}

function addShopItem(item, cat) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const id = String(Date.now());
  ss.getSheetByName('ShoppingList').appendRow([id, item, cat || 'Other', false]);
  return { success: true, id };
}

function toggleShopItem(id) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('ShoppingList');
  const last = sheet.getLastRow();
  if (last <= 1) return { success: false };
  const data = sheet.getRange('A2:D' + last).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      const newVal = !(data[i][3] === true || data[i][3] === 'TRUE');
      sheet.getRange(i + 2, 4).setValue(newVal);
      return { success: true, checked: newVal };
    }
  }
  return { success: false };
}

function deleteShopItem(id) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('ShoppingList');
  const last = sheet.getLastRow();
  if (last <= 1) return { success: false };
  const data = sheet.getRange('A2:A' + last).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) { sheet.deleteRow(i + 2); return { success: true }; }
  }
  return { success: false };
}

function clearDoneShop() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('ShoppingList');
  const last = sheet.getLastRow();
  if (last <= 1) return { success: true };
  const data = sheet.getRange('A2:D' + last).getValues();
  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i][3] === true || data[i][3] === 'TRUE') sheet.deleteRow(i + 2);
  }
  return { success: true };
}

function clearAllShop() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('ShoppingList');
  const last = sheet.getLastRow();
  if (last > 1) sheet.deleteRows(2, last - 1);
  return { success: true };
}

// ── Family Notes ──────────────────────────────────────────────────────
function getNotes() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('FamilyNotes');
  const last = sheet.getLastRow();
  if (last <= 1) return { success: true, notes: [] };
  const notes = sheet.getRange('A2:D' + last).getValues()
    .filter(r => r[0])
    .map(r => ({ id: String(r[0]), text: r[1], color: r[2] || 'yellow', created: r[3] }))
    .reverse();
  return { success: true, notes };
}

function saveNote(id, text, color) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('FamilyNotes');
  const last = sheet.getLastRow();
  if (last > 1) {
    const data = sheet.getRange('A2:D' + last).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.getRange(i + 2, 2).setValue(text);
        if (color) sheet.getRange(i + 2, 3).setValue(color);
        return { success: true };
      }
    }
  }
  // New note
  const newId = id || String(Date.now());
  const created = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  sheet.appendRow([newId, text, color || 'yellow', created]);
  return { success: true, id: newId };
}

function deleteNote(id) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('FamilyNotes');
  const last = sheet.getLastRow();
  if (last <= 1) return { success: false };
  const data = sheet.getRange('A2:A' + last).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) { sheet.deleteRow(i + 2); return { success: true }; }
  }
  return { success: false };
}

// Fallback safety net: catches the reset if the scheduled trigger below
// (setupWeeklyStrikeReset) hasn't been installed, or a run was missed.
function maybeResetStrikes_() {
  if (new Date().getDay() !== 1) return; // not Monday
  const ss = SpreadsheetApp.openById(SS_ID);
  const lastReset = getLastStrikeReset_(ss);
  const today = new Date().toDateString();
  if (lastReset && lastReset.toDateString() === today) return;
  resetAllStrikes_();
}

function getLastStrikeReset_(ss) {
  const data = ss.getSheetByName('Config').getRange('A2:B20').getValues();
  const row = data.find(r => r[0] === 'last_strike_reset');
  return row && row[1] ? new Date(row[1]) : null;
}

// The actual reset. Called by the weekly trigger (see setupWeeklyStrikeReset)
// and as a fallback from maybeResetStrikes_.
function resetAllStrikes_() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const kidsSheet = ss.getSheetByName('KidsData');
  kidsSheet.getRange('A2:H10').getValues()
    .forEach((r, i) => { if (r[0]) kidsSheet.getRange(i + 2, 5).setValue(0); });
  const configSheet = ss.getSheetByName('Config');
  const data = configSheet.getRange('A2:B20').getValues();
  const rowIdx = data.findIndex(r => r[0] === 'last_strike_reset');
  if (rowIdx >= 0) configSheet.getRange(rowIdx + 2, 2).setValue(new Date().toISOString());
}

// ── Run this ONCE from the Apps Script editor (select it in the function
// dropdown and click Run) to install the automatic weekly reset. Without
// this, strikes only reset if someone happens to open the app on a Monday.
function setupWeeklyStrikeReset() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'resetAllStrikes_') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('resetAllStrikes_').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(0).create();
  return { success: true, message: 'Weekly strike reset trigger installed — runs every Monday around midnight.' };
}
