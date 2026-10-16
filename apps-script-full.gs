// ===================== 5 Minute Dog – waitlist receiver =====================
// Deploy → Manage deployments → edit → New version after pasting (keeps the same /exec URL).
var FIELDS = ['email', 'whatsapp', 'delivery', 'dogName', 'age', 'problem', 'consent'];

function clean_(v) {
  v = String(v || '').slice(0, 200);
  return /^[=+\-@]/.test(v) ? "'" + v : v; // block formula injection
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Signups');
    var p = (e && e.parameter) || {};
    ensureCols_(s, ['consent']);
    var c = cols_(s);
    var row = new Array(s.getLastColumn()).fill('');
    row[c.timestamp] = new Date();
    FIELDS.forEach(function (k) { if (k in c) row[c[k]] = clean_(p[k]); });
    row[c.consent] = p.consent === 'yes' ? 'yes' : 'no';
    if ('lastDaySent' in c) row[c.lastDaySent] = 0;
    if ('status' in c) row[c.status] = 'waitlist'; // not emailed until Rui sets 'active'
    s.appendRow(row);
    return ContentService.createTextOutput('OK');
  } finally {
    lock.releaseLock();
  }
}

// ===================== 5 Minute Dog – daily sender =====================
// Adds: Content tab (Day 1-7), Signups columns startDate/lastDaySent/status/lastSentAt/unsubToken,
// sendDaily() on a ~07:00 Europe/Amsterdam daily trigger, and an unsubscribe link via doGet.
// SAFE BY DEFAULT: DRY_RUN = true only logs. Set false (Script Properties: DRY_RUN=false) to send.

var SENDER_NAME = '5 Minute Dog';
var REPLY_TO = 'rlemospereira9@gmail.com';
var TOTAL_DAYS = 7;
var WEBAPP_URL = 'https://script.google.com/macros/s/AKfycbxUwF0BmPLpFlJAlv4GtPljF189QlPRzKyU6r2HqTsDCu9Yq21PPyDb63Uj593Dt37W/exec';
var TZ = 'Europe/Amsterdam';
var EXTRA_COLS = ['consent', 'startDate', 'lastDaySent', 'status', 'lastSentAt', 'unsubToken'];
// Statuses: waitlist (default for new signups, NOT sent), active, paused, done, unsubscribed.

function isDryRun_() {
  var v = PropertiesService.getScriptProperties().getProperty('DRY_RUN');
  return v !== 'false'; // anything except the literal 'false' = dry run
}

var CONTENT = [
 [
  1,
  "Day 1 with {{dogName}}: Name game: \"my name means good things\"",
  "Good morning! Today with {{dogName}}: Name game: \"my name means good things\" 🐾 (5 min)\n\n🧰 You need: 15–20 small treats, a quiet room.\n\n1. Stand near your dog with treats in a pouch or pocket. Don't ask for anything yet.\n2. Say your dog's name once, in a happy, normal voice.\n3. The moment they turn their head towards you, say \"yes!\" and give a treat.\n4. If they don't turn, make a small kissy noise or step back — then \"yes!\" and treat as soon as they look.\n5. Wait until they wander off or sniff, then repeat. Aim for 10–15 reps, then stop.\n\n✅ Success looks like:\n• Your dog turns towards you within a second or two of hearing their name.\n• They start hanging around you, expecting the game.\n• If not, that's fine — tomorrow starts one step easier.\n\n💡 Why it helps: Every other exercise starts with getting your dog's attention, and this makes their name a reliable \"look at me\" signal."
 ],
 [
  2,
  "Day 2 with {{dogName}}: Check-in: rewarding eye contact",
  "Good morning! Today with {{dogName}}: Check-in: rewarding eye contact 🐾 (5 min)\n\n🧰 You need: 15–20 small treats, a quiet room (later: hallway or garden).\n\n1. Stand or sit with treats ready. Stay quiet — no name, no cues.\n2. Wait. The moment your dog glances at your face, say \"yes!\" and treat.\n3. Toss the next treat a little way onto the floor so they move away, then wait for the next glance.\n4. Once glances come quickly, wait for a slightly longer look (one second) before marking.\n5. Do 10–15 reps, then end with a little play or a sniff.\n\n✅ Success looks like:\n• Your dog looks back up at you soon after eating each tossed treat.\n• They offer eye contact without you saying anything.\n• If not, that's fine — tomorrow starts one step easier.\n\n💡 Why it helps: A dog who checks in with you on their own is easier to walk, call back and calm down."
 ],
 [
  3,
  "Day 3 with {{dogName}}: Hand target: \"touch\"",
  "Good morning! Today with {{dogName}}: Hand target: \"touch\" 🐾 (5 min)\n\n🧰 You need: 15–20 small treats, a quiet room.\n\n1. Hold your flat open palm a few centimetres from your dog's nose, at their nose height.\n2. Most dogs sniff a new hand. When their nose touches your palm, \"yes!\" and treat from your other hand.\n3. Remove your hand behind your back, then present it again. Repeat 5 times.\n4. When they're bumping your palm every time, say \"touch\" just before presenting it.\n5. Move your hand slightly to the left, right, or a little further away so they take a step to reach it.\n\n✅ Success looks like:\n• Your dog boops your palm with their nose 8 out of 10 times.\n• They'll take a step or two to reach your hand.\n• If not, that's fine — tomorrow starts one step easier.\n\n💡 Why it helps: \"Touch\" is a gentle way to move your dog anywhere — away from a door, back to your side, or off the sofa — without pulling or pushing."
 ],
 [
  4,
  "Day 4 with {{dogName}}: Sit: from lure to hand signal",
  "Good morning! Today with {{dogName}}: Sit: from lure to hand signal 🐾 (5 min)\n\n🧰 You need: 15–20 small treats, a quiet room.\n\n1. Hold a treat at your dog's nose, then slowly lift it up and slightly back over their head. As their head goes up, their bottom usually goes down.\n2. The moment their bottom touches the floor, \"yes!\" and give the treat.\n3. Repeat 4–5 times with the treat in your hand.\n4. Now do the exact same hand movement with an empty hand. When they sit, \"yes!\" and treat from your other hand or pocket.\n5. Gradually make the hand movement smaller until it's just a small upward flick of your palm.\n\n✅ Success looks like:\n• Your dog sits for the empty-hand signal at least 3 times in a row.\n• The hand movement is getting smaller without them losing it.\n• If not, that's fine — tomorrow starts one step easier.\n\n💡 Why it helps: A quick sit is a polite alternative to jumping up, rushing doors or pulling — your dog can't do both at once."
 ],
 [
  5,
  "Day 5 with {{dogName}}: Loose-leash: the 1-step game",
  "Good morning! Today with {{dogName}}: Loose-leash: the 1-step game 🐾 (5 min)\n\n🧰 You need: lead, 20 pea-sized treats, a quiet hallway or garden.\n\n1. Clip on the lead and stand still with your dog beside you. Wait until the lead hangs in a soft \"J\" shape.\n2. Take one step forward. If the lead stays loose, say \"yes!\" and feed a treat right next to your leg.\n3. If your dog pulls ahead, simply stop. Don't tug. Wait for them to look back or step towards you — the moment the lead slackens, \"yes!\" and treat.\n4. Repeat one step at a time. After 5 loose steps in a row, try two steps between treats.\n5. Stop after 5 minutes, ideally on a success. Finish with a quick game or a sniff.\n\n✅ Success looks like:\n• Your dog stays beside you for 5 one-step repetitions in a row.\n• They check in by looking up at you without being asked.\n• If not, that's fine — tomorrow starts one step easier.\n\n💡 Why it helps: Your dog learns that a loose lead is what makes the walk move forward, and pulling makes it stop."
 ],
 [
  6,
  "Day 6 with {{dogName}}: Recall: the ping-pong game",
  "Good morning! Today with {{dogName}}: Recall: the ping-pong game 🐾 (5 min)\n\n🧰 You need: 20 small treats (extra tasty), a room or hallway; a second person helps but isn't required.\n\n1. Start close: stand 2 steps from your dog. Say their name and \"come!\" once, in a happy voice, and step backwards.\n2. When they reach you, \"yes!\", gently touch their collar, then treat. Make it a party.\n3. Toss a treat away from you to reset, and call again when they've eaten it.\n4. With a helper, take turns calling your dog back and forth between you, a few steps apart.\n5. Slowly increase the distance, or call from around a corner. Never call your dog to something they dislike.\n\n✅ Success looks like:\n• Your dog runs straight to you on the first call, 5 times in a row.\n• They're happy to have their collar touched when they arrive.\n• If not, that's fine — tomorrow starts one step easier.\n\n💡 Why it helps: Coming back becomes the best game in the house, so it's built up before you need it outside."
 ],
 [
  7,
  "Day 7 with {{dogName}}: Settle on a mat",
  "Good morning! Today with {{dogName}}: Settle on a mat 🐾 (5 min)\n\n🧰 You need: a mat, towel or dog bed; 20 small treats; a calm moment (after a walk is ideal).\n\n1. Put the mat on the floor near you. When your dog steps on it, \"yes!\" and drop a treat on the mat.\n2. Repeat, rewarding all four paws on the mat. Use \"touch\" from Day 3 to guide them on if needed.\n3. When they're standing on it, wait — if they sit or lie down, reward calmly between their front paws.\n4. Once they're lying down, drop a treat every few seconds, then slowly stretch to every 10–15 seconds. Keep your voice low and calm.\n5. End by saying \"free!\" and inviting them off the mat. Pick the mat up so it stays special.\n\n✅ Success looks like:\n• Your dog heads to the mat by themselves when you put it down.\n• They lie on it for 30 seconds while you sit nearby.\n• If not, that's fine — tomorrow starts one step easier.\n\n💡 Why it helps: A mat gives your dog a clear \"this is where to relax\" spot, which helps with barking, visitors and restless evenings."
 ]
];

/** Run ONCE manually from the editor: creates Content tab, extra columns, secret, trigger. Sends nothing. */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.getActive().setSpreadsheetTimeZone(TZ);
  var c = ss.getSheetByName('Content') || ss.insertSheet('Content');
  c.clear();
  c.getRange(1, 1, 1, 3).setValues([['day', 'subject', 'body']]);
  c.getRange(2, 1, CONTENT.length, 3).setValues(CONTENT);
  c.setFrozenRows(1);

  var s = ss.getSheetByName('Signups');
  var head = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
  EXTRA_COLS.forEach(function (h) {
    if (head.indexOf(h) === -1) { s.getRange(1, s.getLastColumn() + 1).setValue(h); head.push(h); }
  });

  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('UNSUB_SECRET')) props.setProperty('UNSUB_SECRET', Utilities.getUuid() + Utilities.getUuid());
  if (!props.getProperty('DRY_RUN')) props.setProperty('DRY_RUN', 'true');

  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendDaily') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendDaily').timeBased().everyDays(1).atHour(7).inTimezone(TZ).create();
  Logger.log('Setup done. DRY_RUN=' + isDryRun_());
}

function ensureCols_(sheet, names) {
  var head = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  names.forEach(function (h) {
    if (head.indexOf(h) === -1) { sheet.getRange(1, head.length + 1).setValue(h); head.push(h); }
  });
}

function cols_(sheet) {
  var head = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var m = {}; head.forEach(function (h, i) { m[h] = i; }); return m;
}

function token_(email) {
  var secret = PropertiesService.getScriptProperties().getProperty('UNSUB_SECRET');
  var sig = Utilities.computeHmacSha256Signature(String(email).toLowerCase().trim(), secret);
  return Utilities.base64EncodeWebSafe(sig).replace(/=+$/, '').slice(0, 24);
}

function unsubUrl_(email) {
  return WEBAPP_URL + '?unsub=' + encodeURIComponent(email) + '&t=' + token_(email);
}

function esc_(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }

/** Daily trigger target. Sends each ACTIVE email-delivery signup their next day. */
function sendDaily() {
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var dry = isDryRun_();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var s = ss.getSheetByName('Signups'), c = cols_(s);
    var content = ss.getSheetByName('Content').getRange(2, 1, TOTAL_DAYS, 3).getValues();
    var last = s.getLastRow(); if (last < 2) return;
    var rows = s.getRange(2, 1, last - 1, s.getLastColumn()).getValues();
    var today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
    var sent = 0;
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i], rowNum = i + 2;
      var email = String(r[c.email] || '').trim();
      if (!email || String(r[c.status]) !== 'active') continue;
      if (String(r[c.consent]) !== 'yes') continue; // explicit opt-in only
      if (String(r[c.delivery]) !== 'email') continue; // WhatsApp-only: skip until WhatsApp launches
      var lastSent = r[c.lastSentAt] ? Utilities.formatDate(new Date(r[c.lastSentAt]), TZ, 'yyyy-MM-dd') : '';
      if (lastSent === today) continue; // never twice a day
      var day = Number(r[c.lastDaySent] || 0) + 1;
      if (day > TOTAL_DAYS) { s.getRange(rowNum, c.status + 1).setValue('done'); continue; }
      if (!dry && MailApp.getRemainingDailyQuota() < 1) { Logger.log('Quota reached; stopping at row ' + rowNum); break; }

      var dog = String(r[c.dogName] || 'your dog');
      var subject = String(content[day - 1][1]).replace(/\{\{dogName\}\}/g, dog);
      var body = String(content[day - 1][2]).replace(/\{\{dogName\}\}/g, dog);
      var url = unsubUrl_(email);
      var footer = '\n\n—\nYou get this because you signed up at 5 Minute Dog. Unsubscribe: ' + url;
      var html = '<div style="font-family:system-ui,sans-serif;font-size:15px;line-height:1.6;max-width:560px">' +
        esc_(body).replace(/\n/g, '<br>') +
        '<hr style="border:0;border-top:1px solid #eee;margin:24px 0"><small style="color:#888">You get this because you signed up at 5 Minute Dog. <a href="' + url + '">Unsubscribe</a></small></div>';

      if (dry) {
        Logger.log('[DRY RUN] row ' + rowNum + ' → ' + email + ' | Day ' + day + ' | ' + subject);
      } else {
        GmailApp.sendEmail(email, subject, body + footer, { name: SENDER_NAME, replyTo: REPLY_TO, htmlBody: html });
        var upd = {}; upd[c.lastDaySent] = day; upd[c.lastSentAt] = new Date();
        if (!r[c.startDate]) upd[c.startDate] = new Date();
        if (day === TOTAL_DAYS) upd[c.status] = 'done';
        Object.keys(upd).forEach(function (k) { s.getRange(rowNum, Number(k) + 1).setValue(upd[k]); });
        sent++;
        Utilities.sleep(300);
      }
    }
    Logger.log((dry ? '[DRY RUN] ' : '') + 'sendDaily finished, sent=' + sent);
  } finally { lock.releaseLock(); }
}

/** Unsubscribe link handler: GET ?unsub=<email>&t=<token> */
function doGet(e) {
  var p = (e && e.parameter) || {};
  var msg = 'Nothing to do here.';
  if (p.unsub) {
    var email = String(p.unsub).toLowerCase().trim();
    if (p.t && p.t === token_(email)) {
      var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Signups'), c = cols_(s);
      var last = s.getLastRow(), n = 0;
      if (last >= 2) {
        var emails = s.getRange(2, c.email + 1, last - 1, 1).getValues();
        emails.forEach(function (row, i) {
          if (String(row[0]).toLowerCase().trim() === email) { s.getRange(i + 2, c.status + 1).setValue('unsubscribed'); n++; }
        });
      }
      msg = n ? 'You\'re unsubscribed. You won\'t get any more 5 Minute Dog emails.' : 'This address isn\'t on our list.';
    } else {
      msg = 'This unsubscribe link is invalid. Reply to any of our emails and we\'ll remove you by hand.';
    }
  }
  return HtmlService.createHtmlOutput('<div style="font-family:system-ui,sans-serif;max-width:480px;margin:60px auto;text-align:center"><h2>🐾 5 Minute Dog</h2><p>' + esc_(msg) + '</p></div>');
}

/** Manual helper: mark a signup active (starts Day 1 at the next 07:00 run). */
function activateRow(rowNum) {
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Signups'), c = cols_(s);
  s.getRange(rowNum, c.status + 1).setValue('active');
}
