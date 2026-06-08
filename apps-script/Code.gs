// ============================================================
// SEVEN PRESIDENTS LIFEGUARD SCHEDULING SYSTEM — Phase 1
// Code.gs — Google Apps Script Backend
// ============================================================

const SS = () => SpreadsheetApp.getActiveSpreadsheet();
const SH = (name) => SS().getSheetByName(name);

const SHEETS = {
  GUARDS:         'Guards',
  POSTS:          'Posts',
  TEMPLATES:      'ShiftTemplates',
  SHIFTS:         'Shifts',
  AVAILABILITY:   'Availability',
  REQUESTS:       'ShiftRequests',
  SWAPS:          'SwapRequests',
  PERIODS:        'PayPeriods',
  CONFIG:         'Config',
  NOTIFICATIONS:  'Notifications',
  TIME_RECORDS:   'TimeRecords',
  SHIFT_STATS:    'ShiftStats',
  SESSIONS:       'Sessions',
  LOCATIONS:      'Locations',
  BADGES:         'Badges'
};

// ── ENTRY POINT ─────────────────────────────────────────────

function doGet(e) {
  if (e && e.parameter && e.parameter.board === '1') {
    return HtmlService.createTemplateFromFile('Board')
      .evaluate()
      .setTitle('7P Beach Assignments')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  const tmpl = HtmlService.createTemplateFromFile('index');
  tmpl.inviteToken = (e && e.parameter && e.parameter.invite) ? String(e.parameter.invite) : '';
  tmpl.sessionToken = (e && e.parameter && e.parameter.st) ? String(e.parameter.st) : '';
  return tmpl.evaluate()
    .setTitle('7 Presidents STS')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getBoardData(dateStr) {
  const date = dateStr || toYMD(new Date());
  const guards = sheetToObjects(SHEETS.GUARDS).filter(g => g.status !== 'inactive');
  const posts = getAllPosts();
  const shifts = sheetToObjects(SHEETS.SHIFTS).filter(s =>
    toYMD(s.date) === date && s.assigned_guard_id && s.status !== 'cancelled'
  );
  const guardMap = {};
  guards.forEach(g => { guardMap[String(g.id)] = g; });
  const adminPostIds = new Set(posts.filter(p => p.name.trim().toLowerCase() === 'admin').map(p => p.id));
  const BOARD_RANK_ORDER = [
    'lifeguard supervisor','crew captain','beach captain','captain',
    'lieutenant','vet','guard','rookie','cadet','replacement','volunteer'
  ];
  const boardRankIdx = r => { const i = BOARD_RANK_ORDER.indexOf((r||'').toLowerCase()); return i < 0 ? 99 : i; };
  const boardSort = (a, b) => boardRankIdx(a.rank) - boardRankIdx(b.rank) || parseFloat(b.pay_rate||0) - parseFloat(a.pay_rate||0);
  const officers = shifts
    .filter(s => adminPostIds.has(s.post_id))
    .map(s => guardMap[String(s.assigned_guard_id)])
    .filter(Boolean)
    .map(g => ({ name: g.name, rank: g.rank || '', pay_rate: g.pay_rate || 0 }))
    .sort(boardSort);
  const postData = posts
    .filter(p => !adminPostIds.has(p.id))
    .map(post => {
      const postGuards = shifts
        .filter(s => s.post_id === post.id)
        .map(s => guardMap[String(s.assigned_guard_id)])
        .filter(Boolean)
        .map(g => ({ name: g.name, rank: g.rank || '', pay_rate: g.pay_rate || 0 }))
        .sort(boardSort);
      return { name: post.name, color: post.color || '#2176ae', guards: postGuards };
    }).filter(p => p.guards.length > 0);
  const rankColors = {
    'Lifeguard Supervisor':'#d4a017','Crew Captain':'#d4a017',
    'Captain':'#ef4444','Beach Captain':'#ef4444',
    'Lieutenant':'#f97316','Vet':'#3b82f6',
    'Guard':'#a78bfa','Rookie':'#a78bfa',
    'Replacement':'#94a3b8','Volunteer':'#94a3b8',
    'Cadet':'#6ee7b7'
  };
  getAllConfig().filter(c => c.key && c.key.startsWith('rank_color_')).forEach(c => {
    const k = c.key.slice(11).replace(/_/g, ' ');
    const match = Object.keys(rankColors).find(r => r.toLowerCase() === k.toLowerCase());
    if (match && c.value) rankColors[match] = c.value;
  });
  return { date, posts: postData, officers, rankColors };
}

function clientGetBoardUrl(token) {
  return ScriptApp.getService().getUrl() + '?board=1';
}

// ── AUTH ─────────────────────────────────────────────────────

// ── PASSWORD HASHING ─────────────────────────────────────────

function hashPassword(password, salt) {
  // SHA-256 with salt using Apps Script Utilities
  salt = salt || Utilities.base64Encode(Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(Math.random())
  )).slice(0, 16);
  const hash = Utilities.base64Encode(
    Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      salt + ':' + password
    )
  );
  return { hash: salt + ':' + hash, salt };
}

function verifyPassword(password, storedHash) {
  if (!storedHash || !storedHash.includes(':')) return false;
  const salt = storedHash.split(':')[0];
  const { hash } = hashPassword(password, salt);
  return hash === storedHash;
}

// ── SESSION (Sheet-based, 90-day expiry) ─────────────────────

function createSession(guard) {
  const token = Utilities.base64Encode(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
      guard.id + ':' + Date.now() + ':' + Math.random())
  ).replace(/[^a-zA-Z0-9]/g,'').slice(0,32);

  const expires = new Date(Date.now() + 90 * 24 * 3600 * 1000).toISOString(); // 90 days

  // Ensure Sessions sheet exists
  let sheet = SH(SHEETS.SESSIONS);
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(SHEETS.SESSIONS);
    sheet.getRange(1,1,1,5).setValues([['token','guard_id','role','created_at','expires_at']]);
  }

  sheet.appendRow([token, String(guard.id), roleFor(guard.rank), new Date().toISOString(), expires]);
  return token;
}

function validateSession(token) {
  if (!token) return null;
  const sheet = SH(SHEETS.SESSIONS);
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(token)) {
      const expiresAt = data[i][4];
      if (expiresAt && new Date(expiresAt) < new Date()) {
        // Expired — delete row
        sheet.deleteRow(i + 1);
        return null;
      }
      return { guardId: String(data[i][1]), role: String(data[i][2]), expires: expiresAt };
    }
  }
  return null;
}

function destroySession(token) {
  if (!token) return;
  const sheet = SH(SHEETS.SESSIONS);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(token)) {
      sheet.deleteRow(i + 1);
      return;
    }
  }
}

// Run once from the Apps Script editor to install time-based triggers
function addTournamentTemplate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tmpl = ss.getSheetByName('ShiftTemplates');
  const data = tmpl.getDataRange().getValues();
  if (data.some(row => row[0] === 'TOURN')) { Logger.log('TOURN already exists.'); return; }
  tmpl.appendRow(['TOURN', 'Tournament', '6:00 PM', '8:30 PM', 2.5, 0]);
  Logger.log('Tournament template added.');
}

function installTriggers() {
  // Remove any existing triggers for these functions to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === 'autoClockOutCheck' || fn === 'cleanExpiredSessions') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('autoClockOutCheck')
    .timeBased().everyHours(1).create();
  ScriptApp.newTrigger('cleanExpiredSessions')
    .timeBased().everyDays(1).create();
}

// Clean up expired sessions (run periodically as trigger)
function cleanExpiredSessions() {
  const sheet = SH(SHEETS.SESSIONS);
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  for (let i = data.length - 1; i >= 1; i--) {
    const exp = data[i][4];
    if (exp && new Date(exp) < now) {
      sheet.deleteRow(i + 1);
    }
  }
}

// ── AUTH FUNCTIONS ───────────────────────────────────────────

function getUserSession() {
  // Password-only auth — no Google OAuth
  return { authenticated: false };
}

function getSessionFromToken(token) {
  if (!token) return null;
  const sess = validateSession(token);
  if (!sess) return null;
  const guard = findGuardById(sess.guardId);
  if (!guard) return null;
  return { authenticated: true, guard, role: sess.role, token };
}

function loginWithPassword(email, password) {
  const guard = findGuardByEmail(email.toLowerCase().trim());
  if (!guard) return { success: false, message: 'No account found for that email.' };
  if (guard.status === 'inactive') return { success: false, message: 'Account is inactive.' };

  // Support both old temp_password and new password_hash
  const storedHash = guard.password_hash || '';
  const tempPw = guard.temp_password || '';
  let verified = false;

  if (storedHash) {
    verified = verifyPassword(password, storedHash);
  } else if (tempPw) {
    // Legacy plaintext password — migrate on successful login
    verified = String(tempPw) === String(password);
    if (verified) {
      // Upgrade to hashed
      const { hash } = hashPassword(password);
      updateById(SHEETS.GUARDS, { id: guard.id, password_hash: hash, temp_password: '' });
    }
  }

  if (!verified) return { success: false, message: 'Incorrect password.' };

  const mustChange = guard.must_change_pw === 'true' || guard.must_change_pw === true;
  const token = createSession(guard);
  return {
    success: true,
    token,
    guard: sanitizeGuard(guard),
    role: roleFor(guard.rank),
    must_change_pw: mustChange
  };
}

function changePassword(token, currentPassword, newPassword) {
  const sess = validateSession(token);
  if (!sess) return { success: false, message: 'Session expired. Please log in again.' };
  const guard = findGuardById(sess.guardId);
  if (!guard) return { success: false, message: 'Guard not found.' };

  // Verify current password (skip if must_change_pw — admin set it)
  if (guard.must_change_pw !== 'true') {
    const storedHash = guard.password_hash || guard.temp_password || '';
    if (!verifyPassword(currentPassword, storedHash) && storedHash !== currentPassword) {
      return { success: false, message: 'Current password is incorrect.' };
    }
  }

  if (!newPassword || newPassword.length < 6) return { success: false, message: 'Password must be at least 6 characters.' };

  const { hash } = hashPassword(newPassword);
  updateById(SHEETS.GUARDS, { id: guard.id, password_hash: hash, temp_password: '', must_change_pw: 'false' });

  if (guard.email) {
    const loginUrl = 'https://baconlt.github.io/7p-sts/';
    const htmlBody = `<div style="font-family:Arial,sans-serif;max-width:540px;margin:0 auto">
  <div style="background:#0d2137;padding:20px 24px;border-radius:8px 8px 0 0">
    <div style="font-size:1.3rem;font-weight:800;color:#fff;letter-spacing:1px">7 Presidents STS</div>
    <div style="color:#57a0d3;font-size:.8rem;letter-spacing:2px;text-transform:uppercase;margin-top:3px">Lifeguard Scheduling System</div>
  </div>
  <div style="background:#fff;padding:28px 24px;border:1px solid #d0e4ef;border-top:none">
    <p style="font-size:1rem;margin-bottom:12px">Hi ${guard.name},</p>
    <p style="margin-bottom:18px">Your password has been updated successfully.</p>
    <div style="text-align:center;margin:24px 0">
      <a href="${loginUrl}" style="background:#0d2137;color:#fff;text-decoration:none;padding:14px 28px;border-radius:8px;font-weight:700;font-size:1rem;display:inline-block;letter-spacing:.5px">Log In Here →</a>
    </div>
    <p style="color:#6b8a9a;font-size:.82rem;margin-top:18px">If the button doesn't work, copy and paste this link into your browser:</p>
    <p style="font-size:.78rem;word-break:break-all;color:#2176ae;margin-bottom:18px">${loginUrl}</p>
    <hr style="border:none;border-top:1px solid #d0e4ef;margin:18px 0">
    <p style="color:#6b8a9a;font-size:.78rem">If you did not make this change, contact your administrator immediately.</p>
    <p style="color:#6b8a9a;font-size:.78rem;margin-top:8px">— 7 Presidents STS</p>
  </div>
</div>`;
    const body = `Hi ${guard.name},\n\nYour password has been updated successfully.\n\nLog in here:\n${loginUrl}\n\nIf you did not make this change, contact your administrator immediately.\n\n— 7 Presidents STS`;
    try {
      MailApp.sendEmail({ to: guard.email, subject: 'Your 7 Presidents STS password has been updated', htmlBody, body });
    } catch(e) {
      Logger.log('Password change email error: ' + e.message);
    }
  }

  return { success: true };
}

function adminSetPassword(guardId, newPassword) {
  requireAdmin();
  if (!newPassword || newPassword.length < 4) return { success: false, message: 'Password too short.' };
  const { hash } = hashPassword(newPassword);
  updateById(SHEETS.GUARDS, { id: guardId, password_hash: hash, temp_password: '', must_change_pw: 'true' });
  return { success: true };
}

function requestPasswordReset(email) {
  const guard = findGuardByEmail(email.toLowerCase().trim());
  // Always return success to avoid email enumeration
  if (!guard || guard.status === 'inactive') return { success: true };

  // Generate 6-digit code
  const code = String(Math.floor(100000 + Math.random() * 900000));
  const expires = new Date(Date.now() + 1800000).toISOString(); // 30 minutes

  ensureGuardColumns_();
  updateById(SHEETS.GUARDS, { id: guard.id, reset_token: code, reset_token_expires: expires });

  try {
    MailApp.sendEmail({
      to: guard.email,
      subject: '[7 Presidents STS] Password Reset Code',
      htmlBody: `<p>Hi ${guard.name},</p>
        <p>Your password reset code is:</p>
        <p style="font-size:2rem;font-weight:bold;letter-spacing:6px;color:#2176ae;margin:16px 0">${code}</p>
        <p>Enter this code on the login page to set a new password. This code expires in 30 minutes.</p>
        <p>If you did not request this, ignore this email.</p>
        <p>— 7 Presidents STS</p>`,
      body: `Hi ${guard.name},\n\nYour password reset code is: ${code}\n\nEnter this code on the login page to set a new password.\nThis code expires in 30 minutes.\n\nIf you did not request this, ignore this email.\n\n— 7 Presidents STS`
    });
  } catch(e) { Logger.log('Reset email error: ' + e.message); }
  return { success: true };
}

function ensureGuardColumns_() {
  // Add reset_token and reset_token_expires columns to Guards sheet if missing
  const sheet = SH(SHEETS.GUARDS);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const needed = ['password_hash','must_change_pw','reset_token','reset_token_expires','invite_token','invite_token_expires','locker_number','parking_pass'];
  needed.forEach(col => {
    if (!headers.includes(col)) {
      const newCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, newCol).setValue(col)
        .setBackground('#0d2137').setFontColor('#ffffff').setFontWeight('bold');
      Logger.log('Added column: ' + col);
    }
  });
}

function generateInviteToken_() {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Date.now() + ':' + Math.random() + ':' + Math.random())
    .map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('').slice(0, 48);
}

function sendInvite(adminToken, guardId) {
  requireAdmin(adminToken);
  ensureGuardColumns_();
  const guard = findGuardById(guardId);
  if (!guard) return { success: false, message: 'Guard not found.' };
  if (guard.status === 'inactive') return { success: false, message: 'Guard is inactive.' };

  const inviteToken = generateInviteToken_();
  const expires = new Date(Date.now() + 7 * 24 * 3600 * 1000).toISOString();
  const appUrl = ScriptApp.getService().getUrl();
  const inviteUrl = appUrl + '?invite=' + inviteToken;

  updateById(SHEETS.GUARDS, { id: guardId, invite_token: inviteToken, invite_token_expires: expires });

  const htmlBody = `<div style="font-family:Arial,sans-serif;max-width:540px;margin:0 auto">
  <div style="background:#0d2137;padding:20px 24px;border-radius:8px 8px 0 0">
    <div style="font-size:1.3rem;font-weight:800;color:#fff;letter-spacing:1px">7 Presidents STS</div>
    <div style="color:#57a0d3;font-size:.8rem;letter-spacing:2px;text-transform:uppercase;margin-top:3px">Lifeguard Scheduling System</div>
  </div>
  <div style="background:#fff;padding:28px 24px;border:1px solid #d0e4ef;border-top:none">
    <p style="font-size:1rem;margin-bottom:12px">Hi ${guard.name},</p>
    <p style="margin-bottom:14px">You've been invited to join the <strong>7 Presidents Lifeguard Scheduling System</strong> — the hub for shift scheduling, time tracking, and availability management for Ocean Rescue &amp; Events.</p>
    <p style="margin-bottom:18px">Click the button below to set up your account. This link is valid for <strong>7 days</strong> and can only be used once.</p>
    <div style="text-align:center;margin:24px 0">
      <a href="${inviteUrl}" style="background:#0d2137;color:#fff;text-decoration:none;padding:14px 28px;border-radius:8px;font-weight:700;font-size:1rem;display:inline-block;letter-spacing:.5px">Set Up My Account →</a>
    </div>
    <p style="color:#6b8a9a;font-size:.82rem;margin-top:18px">If the button doesn't work, copy and paste this link into your browser:</p>
    <p style="font-size:.78rem;word-break:break-all;color:#2176ae;margin-bottom:18px">${inviteUrl}</p>
    <hr style="border:none;border-top:1px solid #d0e4ef;margin:18px 0">
    <p style="color:#4a6a7a;font-size:.83rem;margin-bottom:8px">Already set up your account?</p>
    <div style="text-align:center;margin:10px 0 18px">
      <a href="https://baconlt.github.io/7p-sts/" style="background:#2176ae;color:#fff;text-decoration:none;padding:10px 22px;border-radius:8px;font-weight:700;font-size:.9rem;display:inline-block">Log In Here →</a>
    </div>
    <hr style="border:none;border-top:1px solid #d0e4ef;margin:18px 0">
    <p style="color:#6b8a9a;font-size:.78rem">If you weren't expecting this invitation, you can safely ignore this email. The link will expire in 7 days.</p>
    <p style="color:#6b8a9a;font-size:.78rem;margin-top:8px">— 7 Presidents STS</p>
  </div>
</div>`;
  const body = `Hi ${guard.name},\n\nYou've been invited to join the 7 Presidents Lifeguard Scheduling System.\n\nClick the link below to set up your account (valid for 7 days):\n${inviteUrl}\n\nAlready set up your account? Log in here:\nhttps://baconlt.github.io/7p-sts/\n\nIf you weren't expecting this, you can safely ignore this email.\n\n— 7 Presidents STS`;

  try {
    MailApp.sendEmail({ to: guard.email, subject: "You're invited to 7 Presidents STS", htmlBody, body });
  } catch(e) {
    Logger.log('Invite email error: ' + e.message);
    return { success: false, message: 'Guard updated but email failed: ' + e.message };
  }
  return { success: true, email: guard.email };
}

function acceptInvite(inviteToken, newPassword) {
  if (!inviteToken) return { success: false, message: 'Missing invite token.' };
  if (!newPassword || newPassword.length < 6) return { success: false, message: 'Password must be at least 6 characters.' };

  const all = sheetToObjects(SHEETS.GUARDS);
  const guard = all.find(g => String(g.invite_token) === String(inviteToken).trim());
  if (!guard) return { success: false, message: 'Invalid or already-used invite link.' };
  if (guard.status === 'inactive') return { success: false, message: 'This account has been deactivated.' };
  if (new Date(guard.invite_token_expires) < new Date()) {
    return { success: false, message: 'This invite link has expired. Ask your admin to send a new invitation.' };
  }

  const { hash } = hashPassword(newPassword);
  updateById(SHEETS.GUARDS, {
    id: guard.id,
    password_hash: hash,
    auth_type: 'password',
    temp_password: '',
    must_change_pw: 'false',
    invite_token: '',
    invite_token_expires: ''
  });

  const freshGuard = findGuardById(guard.id);
  const sessionToken = createSession(freshGuard);
  return { success: true, token: sessionToken, guard: sanitizeGuard(freshGuard), role: roleFor(freshGuard.rank) };
}

function resetPasswordWithToken(code, newPassword) {
  if (!code || !newPassword) return { success: false, message: 'Invalid request.' };
  if (newPassword.length < 6) return { success: false, message: 'Password must be at least 6 characters.' };

  const all = sheetToObjects(SHEETS.GUARDS);
  const guard = all.find(g => String(g.reset_token) === String(code).trim());
  if (!guard) return { success: false, message: 'Invalid code. Please request a new one.' };
  if (new Date(guard.reset_token_expires) < new Date()) return { success: false, message: 'Code has expired. Please request a new one.' };

  const { hash } = hashPassword(newPassword);
  updateById(SHEETS.GUARDS, { id: guard.id, password_hash: hash, temp_password: '', must_change_pw: 'false', reset_token: '', reset_token_expires: '' });

  const appUrl = 'https://baconlt.github.io/7p-sts/';
  try {
    MailApp.sendEmail({
      to: guard.email,
      subject: '[7 Presidents STS] Your password has been reset',
      htmlBody: `<div style="font-family:Arial,sans-serif;max-width:540px;margin:0 auto">
  <div style="background:#0d2137;padding:20px 24px;border-radius:8px 8px 0 0">
    <div style="font-size:1.3rem;font-weight:800;color:#fff;letter-spacing:1px">7 Presidents STS</div>
    <div style="color:#57a0d3;font-size:.8rem;letter-spacing:2px;text-transform:uppercase;margin-top:3px">Lifeguard Scheduling System</div>
  </div>
  <div style="background:#fff;padding:28px 24px;border:1px solid #d0e4ef;border-top:none">
    <p style="font-size:1rem;margin-bottom:12px">Hi ${guard.name},</p>
    <p style="margin-bottom:18px">Your password for the 7 Presidents Lifeguard Scheduling System has been successfully reset. You can now log in with your new password.</p>
    <div style="text-align:center;margin:24px 0">
      <a href="${appUrl}" style="background:#0d2137;color:#fff;text-decoration:none;padding:14px 28px;border-radius:8px;font-weight:700;font-size:1rem;display:inline-block;letter-spacing:.5px">Go to App →</a>
      <p style="margin-top:10px;font-size:.82rem;color:#6b8a9a">or <a href="${appUrl}" style="color:#2176ae">login here</a></p>
    </div>
    <hr style="border:none;border-top:1px solid #d0e4ef;margin:18px 0">
    <p style="color:#6b8a9a;font-size:.78rem">If you did not reset your password, please contact your supervisor immediately.</p>
    <p style="color:#6b8a9a;font-size:.78rem;margin-top:8px">— 7 Presidents STS</p>
  </div>
</div>`,
      body: `Hi ${guard.name},\n\nYour password for the 7 Presidents Lifeguard Scheduling System has been successfully reset.\n\nYou can now log in with your new password at:\n${appUrl}\n\nIf you did not reset your password, please contact your supervisor immediately.\n\n— 7 Presidents STS`
    });
  } catch(e) { Logger.log('Password reset confirmation email error: ' + e.message); }

  return { success: true };
}

function sanitizeGuard(g) {
  // Never send password fields to client
  const { password_hash, temp_password, reset_token, reset_token_expires, invite_token, invite_token_expires, ...safe } = g;
  return safe;
}

function logoutSession(token) {
  destroySession(token);
  return { success: true };
}

function roleFor(rank) {
  if (['Lifeguard Supervisor', 'Crew Captain'].includes(rank)) return 'admin';
  return 'guard';
}

// ── SHEET UTILITIES ──────────────────────────────────────────

function sheetToObjects(name) {
  const sheet = SH(name);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0].map(h => String(h).trim());
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = serializeCell(row[i]); });
    return obj;
  });
}

function serializeCell(val) {
  if (val instanceof Date) {
    const y = val.getFullYear();
    // Year ≤ 1900 = time-only cell (Sheets internal representation)
    if (y <= 1900) {
      let h = val.getHours(), m = val.getMinutes();
      const ap = h >= 12 ? 'PM' : 'AM';
      h = h % 12 || 12;
      return h + ':' + (m < 10 ? '0' : '') + m + ' ' + ap;
    }
    const mo = String(val.getMonth() + 1).padStart(2, '0');
    const d = String(val.getDate()).padStart(2, '0');
    return `${y}-${mo}-${d}`;
  }
  return (val === null || val === undefined) ? '' : val;
}

function toYMD(val) {
  if (!val) return '';
  if (val instanceof Date) return serializeCell(val);
  const s = String(val).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(s)) {
    const [m, d, y] = s.split('/');
    return `${y}-${m.padStart(2,'0')}-${d.padStart(2,'0')}`;
  }
  return s;
}

function appendRow(sheetName, obj) {
  const sheet = SH(sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
  sheet.appendRow(headers.map(h => obj[h] !== undefined ? obj[h] : ''));
}

function updateById(sheetName, updates) {
  const sheet = SH(sheetName);
  if (!sheet) return { success: false, message: 'Sheet not found.' };
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const idCol = headers.indexOf('id');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(updates.id)) {
      Object.keys(updates).forEach(k => {
        const col = headers.indexOf(k);
        if (col >= 0) sheet.getRange(i + 1, col + 1).setValue(updates[k]);
      });
      return { success: true };
    }
  }
  return { success: false, message: 'Row not found.' };
}

function uid(prefix) {
  return `${prefix}-${Date.now()}-${Math.random().toString(36).slice(2,7).toUpperCase()}`;
}

function requireAdmin(token) {
  if (token) {
    const sess = validateSession(token);
    if (!sess || sess.role !== 'admin') throw new Error('Admin access required.');
    return sess;
  }
  // Fallback: check Google session for backward compat during transition
  const email = (() => { try { return Session.getActiveUser().getEmail(); } catch(e) { return null; } })();
  if (email) {
    const guard = findGuardByEmail(email);
    if (guard && roleFor(guard.rank) === 'admin') return { guardId: guard.id, role: 'admin' };
  }
  throw new Error('Admin access required.');
}

// ── GUARDS ───────────────────────────────────────────────────

function findGuardByEmail(email) {
  return sheetToObjects(SHEETS.GUARDS).find(g => g.email === email && g.status !== 'inactive') || null;
}
function findGuardById(id) {
  return sheetToObjects(SHEETS.GUARDS).find(g => g.id === id) || null;
}
function getAllGuards() {
  return sheetToObjects(SHEETS.GUARDS).filter(g => g.status !== 'inactive');
}
function createGuard(d) {
  const id = uid('G');
  appendRow(SHEETS.GUARDS, { id, name: d.name, email: d.email, phone: d.phone||'', rank: d.rank,
    pay_rate: d.pay_rate||0, post_eligibility: d.post_eligibility||'', status:'active',
    auth_type: d.auth_type||'google', temp_password: d.temp_password||'',
    locker_number: d.locker_number||'', parking_pass: d.parking_pass||'', created_at: new Date().toISOString() });
  return { success: true, id };
}
function updateGuard(d) { return updateById(SHEETS.GUARDS, d); }
function deactivateGuard(id) { return updateById(SHEETS.GUARDS, { id, status:'inactive' }); }

// ── POSTS ────────────────────────────────────────────────────

function getAllPosts() {
  const posts = sheetToObjects(SHEETS.POSTS).filter(p => String(p.active).toUpperCase() === 'TRUE');
  return posts.sort((a,b) => (parseFloat(a.sort_order)||99) - (parseFloat(b.sort_order)||99));
}
function createPost(d) {
  const id = uid('P');
  const allPosts = sheetToObjects(SHEETS.POSTS);
  const nextOrder = allPosts.length + 1;
  SH(SHEETS.POSTS).appendRow([id, d.name, d.rank_eligibility||'', 'TRUE', d.color||'#2176ae', d.sort_order||nextOrder]);
  return { success: true, id };
}
function updatePost(d) {
  return updateById(SHEETS.POSTS, d);
}

// ── SHIFT TEMPLATES ──────────────────────────────────────────

function getAllTemplates() { return sheetToObjects(SHEETS.TEMPLATES); }
function templateByCode(code) { return getAllTemplates().find(t => t.code === code) || null; }

// ── PAY PERIODS ──────────────────────────────────────────────

function getAllPeriods() { return sheetToObjects(SHEETS.PERIODS); }
function periodById(id) { return getAllPeriods().find(p => p.id === id) || null; }

function getActivePeriod() {
  const today = toYMD(new Date());
  const periods = getAllPeriods();
  return periods.find(p => toYMD(p.start_date) <= today && toYMD(p.schedule_thru) >= today)
    || (periods.length ? periods[periods.length - 1] : null);
}

function createPeriod(d) {
  const id = uid('PP');
  appendRow(SHEETS.PERIODS, { id, start_date: d.start_date, end_date: d.end_date,
    schedule_thru: d.schedule_thru, locked: 'false', payroll_due: d.payroll_due||'',
    availability_deadline: d.availability_deadline||'', time_report_due: d.time_report_due||'' });
  return { success: true, id };
}
function updatePeriod(d) { return updateById(SHEETS.PERIODS, d); }
function lockPeriod(id) { return updateById(SHEETS.PERIODS, { id, locked:'true' }); }

// ── SHIFTS ───────────────────────────────────────────────────

function getShiftsForPeriod(periodId) {
  const period = periodById(periodId);
  if (period && period.start_date && period.schedule_thru) {
    const start = toYMD(period.start_date), end = toYMD(period.schedule_thru);
    return sheetToObjects(SHEETS.SHIFTS).filter(s => {
      const d = toYMD(s.date);
      return d >= start && d <= end;
    });
  }
  return sheetToObjects(SHEETS.SHIFTS).filter(s => s.pay_period_id === periodId);
}
function getShiftsForGuard(guardId) {
  return sheetToObjects(SHEETS.SHIFTS).filter(s => s.assigned_guard_id === guardId && s.status !== 'cancelled');
}
function getOpenShiftsForGuard(guardId) {
  const guard = findGuardById(guardId);
  if (!guard) return [];
  const eligible = guard.post_eligibility ? String(guard.post_eligibility).split(',').map(p=>p.trim()).filter(Boolean) : [];
  return sheetToObjects(SHEETS.SHIFTS).filter(s => {
    if (s.status !== 'open') return false;
    if (eligible.length && !eligible.includes(s.post_id)) return false;
    return true;
  });
}
function createShift(d) {
  const qty = parseInt(d.quantity)||1;
  const seriesId = (d.type==='recurring') ? uid('SR') : '';
  const created = [];
  for (let i=0; i<qty; i++) {
    const id = uid('S');
    const status = d.assigned_guard_id ? 'filled' : (d.type==='floating'?'floating':'open');
    appendRow(SHEETS.SHIFTS, { id, pay_period_id:d.pay_period_id, date:d.date,
      post_id:d.post_id, template_code:d.template_code, type:d.type,
      assigned_guard_id:d.assigned_guard_id||'', status:d.status||status,
      notes:d.notes||'', created_at:new Date().toISOString(), series_id:seriesId });
    created.push(id);
  }
  if (d.type==='recurring' && d.recur_through) recurringShifts(d, seriesId, qty);
  return { success:true, id:created[0], series_id:seriesId, count:created.length };
}

function assignShift(d) {
  // Fill-or-create: find an open slot for this date/post/template, else create new
  const all = sheetToObjects(SHEETS.SHIFTS);
  const open = all.find(s =>
    s.date && toYMD(s.date)===toYMD(d.date) &&
    s.post_id===d.post_id &&
    s.template_code===d.template_code &&
    (s.status==='open'||s.status==='floating') &&
    !s.assigned_guard_id
  );
  if (open) {
    updateById(SHEETS.SHIFTS, { id:open.id, assigned_guard_id:d.assigned_guard_id, status:'filled',
      custom_start:d.custom_start||'', custom_end:d.custom_end||'' });
    return { success:true, id:open.id, created:false };
  }
  // No open slot — create new (caller confirmed overstaffing)
  const id = uid('S');
  appendRow(SHEETS.SHIFTS, { id, pay_period_id:d.pay_period_id, date:d.date,
    post_id:d.post_id, template_code:d.template_code, type:'assigned',
    assigned_guard_id:d.assigned_guard_id, status:'filled',
    notes:d.notes||'', created_at:new Date().toISOString(), series_id:'',
    custom_start:d.custom_start||'', custom_end:d.custom_end||'' });
  return { success:true, id, created:true };
}

function checkOpenSlot(date, postId, templateCode) {
  // Returns whether an open slot exists — used before assigning to warn admin
  const all = sheetToObjects(SHEETS.SHIFTS);
  const open = all.find(s =>
    s.date && toYMD(s.date)===toYMD(date) &&
    s.post_id===postId &&
    s.template_code===templateCode &&
    (s.status==='open'||s.status==='floating') &&
    !s.assigned_guard_id
  );
  return { hasOpen: !!open, openId: open?open.id:null };
}

function deleteShiftSeries(seriesId, fromDate, scope) {
  // scope: 'one' (handled by cancelShift), 'future', 'all'
  const sheet = SH(SHEETS.SHIFTS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const sidIdx = headers.indexOf('series_id');
  const dateIdx = headers.indexOf('date');
  const statusIdx = headers.indexOf('status');
  if (sidIdx<0) return { success:false, message:'series_id column not found' };
  let count=0;
  for (let i=data.length-1; i>=1; i--) {
    if (String(data[i][sidIdx])!==seriesId) continue;
    const rowDate = toYMD(data[i][dateIdx]);
    if (scope==='future' && rowDate < fromDate) continue;
    sheet.getRange(i+1, statusIdx+1).setValue('cancelled');
    count++;
  }
  return { success:true, count };
}

function getNthWeekday(year, month, nth, wday) {
  // month is 0-indexed; nth=-1 means last occurrence
  if (nth === -1) {
    const dt = new Date(year, month+1, 0, 12); // last day of month
    while (dt.getDay() !== wday) dt.setDate(dt.getDate()-1);
    return dt;
  }
  const first = new Date(year, month, 1);
  const diff = (wday - first.getDay() + 7) % 7;
  const dayNum = 1 + diff + (nth-1)*7;
  if (dayNum > new Date(year, month+1, 0).getDate()) return null; // e.g. 5th Friday may not exist
  return new Date(year, month, dayNum, 12);
}

function recurringShifts(d, seriesId, qty) {
  qty = parseInt(qty)||1;
  const freq = d.recur_freq || 'weekly';
  const [y,m,day] = d.date.split('-').map(Number);
  const startDate = new Date(y, m-1, day, 12);
  const [ey,em,ed] = d.recur_through.split('-').map(Number);
  const end = new Date(ey, em-1, ed, 12);

  const dates = [];

  if (freq === 'daily') {
    const cur = new Date(startDate);
    cur.setDate(cur.getDate()+1);
    while (cur <= end) { dates.push(new Date(cur)); cur.setDate(cur.getDate()+1); }

  } else if (freq === 'weekly') {
    const recurDays = (d.recur_days && d.recur_days.length)
      ? d.recur_days.map(Number)
      : [startDate.getDay()];
    const cur = new Date(startDate);
    cur.setDate(cur.getDate()+1);
    while (cur <= end) {
      if (recurDays.indexOf(cur.getDay()) >= 0) dates.push(new Date(cur));
      cur.setDate(cur.getDate()+1);
    }

  } else if (freq === 'monthly') {
    const mode = d.recur_month_mode || 'date';
    if (mode === 'weekday') {
      const nth = parseInt(d.recur_nth)||1;
      const wday = parseInt(d.recur_wday)||0;
      let mn = m, yr = y;
      mn++; if (mn > 12) { mn = 1; yr++; }
      while (yr < ey || (yr === ey && mn <= em)) {
        const dt = getNthWeekday(yr, mn-1, nth, wday);
        if (dt && dt > startDate && dt <= end) dates.push(dt);
        mn++; if (mn > 12) { mn = 1; yr++; }
      }
    } else {
      const dayOfMonth = parseInt(d.recur_month_day)||day;
      let mn = m, yr = y;
      mn++; if (mn > 12) { mn = 1; yr++; }
      while (yr < ey || (yr === ey && mn <= em)) {
        const maxDay = new Date(yr, mn, 0).getDate();
        const dt = new Date(yr, mn-1, Math.min(dayOfMonth, maxDay), 12);
        if (dt > startDate && dt <= end) dates.push(dt);
        mn++; if (mn > 12) { mn = 1; yr++; }
      }
    }
  }

  dates.forEach(function(dt) {
    for (var i=0; i<qty; i++) {
      var id = uid('S');
      appendRow(SHEETS.SHIFTS, { id: id, pay_period_id: d.pay_period_id, date: serializeCell(dt),
        post_id: d.post_id, template_code: d.template_code, type: 'recurring',
        assigned_guard_id: '', status: 'open',
        notes: d.notes||'', created_at: new Date().toISOString(), series_id: seriesId });
    }
  });
}
function updateShift(d) { return updateById(SHEETS.SHIFTS, d); }
function cancelShift(id) { return updateById(SHEETS.SHIFTS, { id, status:'cancelled' }); }

// ── AVAILABILITY ─────────────────────────────────────────────

function getAvailabilityForPeriod(periodId) {
  const p = periodById(periodId);
  if (!p) return [];
  const start = toYMD(p.start_date), thru = toYMD(p.schedule_thru);
  return sheetToObjects(SHEETS.AVAILABILITY).filter(a => {
    const d = toYMD(a.date);
    return d >= start && d <= thru;
  });
}
function getAvailabilityForGuard(guardId) {
  return sheetToObjects(SHEETS.AVAILABILITY).filter(a => a.guard_id === guardId);
}
function submitAvailability(token, entries) {
  const session = getSessionFromToken(token);
  if (!session) return { success: false, message: 'Not authenticated.' };

  const existing = sheetToObjects(SHEETS.AVAILABILITY);
  for (const e of entries) {
    if (session.role === 'guard' && e.guard_id !== session.guard.id) continue;
    if (session.role === 'guard') {
      const p = getAllPeriods().find(pp => toYMD(e.date) >= toYMD(pp.start_date) && toYMD(e.date) <= toYMD(pp.schedule_thru));
      if (p && p.availability_deadline && p.locked === 'true')
        return { success: false, message: 'This pay period is locked.' };
      // Allow updates if period is active (started) even if deadline passed
      // Only enforce deadline if the period hasn't started yet
      const periodStarted = toYMD(new Date()) >= toYMD(p?.start_date);
      if (p && p.availability_deadline && !periodStarted && toYMD(new Date()) > toYMD(p.availability_deadline))
        return { success: false, message: 'Availability deadline has passed.' };
    }
    const ex = existing.find(x => x.guard_id === e.guard_id && toYMD(x.date) === toYMD(e.date));
    if (ex) {
      updateById(SHEETS.AVAILABILITY, { id: ex.id, status: e.status,
        custom_start: e.custom_start||'', custom_end: e.custom_end||'', ot_willing: e.ot_willing||false, notes: e.notes||'' });
    } else {
      appendRow(SHEETS.AVAILABILITY, { id: uid('AV'), guard_id: e.guard_id, date: e.date, status: e.status,
        custom_start: e.custom_start||'', custom_end: e.custom_end||'', ot_willing: e.ot_willing||false, notes: e.notes||'',
        submitted_at: new Date().toISOString() });
    }
  }
  return { success: true };
}

// ── SHIFT REQUESTS ───────────────────────────────────────────

function getPendingRequests() {
  return sheetToObjects(SHEETS.REQUESTS).filter(r => r.status === 'pending');
}
function getRequestsForGuard(guardId) {
  return sheetToObjects(SHEETS.REQUESTS).filter(r => r.guard_id === guardId);
}
function requestShift(shiftId) {
  const session = getUserSession();
  if (!session.authenticated) return { success: false, message: 'Not authenticated.' };
  const shift = sheetToObjects(SHEETS.SHIFTS).find(s => s.id === shiftId);
  if (!shift || shift.status !== 'open') return { success: false, message: 'Shift not available.' };
  const tmpl = templateByCode(shift.template_code);
  if (tmpl && parseFloat(tmpl.paid_hours) > 0) {
    const wk = weeklyHoursForGuard(session.guard.id, shift.date);
    if (wk + parseFloat(tmpl.paid_hours) > 40)
      return { success: false, message: `Exceeds 40hr weekly limit (${(40-wk).toFixed(1)}h remaining).` };
  }
  appendRow(SHEETS.REQUESTS, { id: uid('SR'), guard_id: session.guard.id, shift_id: shiftId,
    requested_at: new Date().toISOString(), status: 'pending', admin_notes: '' });
  return { success: true };
}
function approveRequest(requestId) {
  requireAdmin();
  const req = sheetToObjects(SHEETS.REQUESTS).find(r => r.id === requestId);
  if (!req) return { success: false };
  updateById(SHEETS.REQUESTS, { id: requestId, status: 'approved' });
  updateById(SHEETS.SHIFTS, { id: req.shift_id, assigned_guard_id: req.guard_id, status: 'filled' });
  sheetToObjects(SHEETS.REQUESTS)
    .filter(r => r.shift_id === req.shift_id && r.id !== requestId && r.status === 'pending')
    .forEach(r => updateById(SHEETS.REQUESTS, { id: r.id, status: 'denied', admin_notes: 'Shift filled.' }));
  const shift = sheetToObjects(SHEETS.SHIFTS).find(s => s.id === req.shift_id);
  notify(req.guard_id, 'shift_approved', `Your shift request for ${shift ? shift.date : ''} was approved.`);
  return { success: true };
}
function denyRequest(requestId, reason) {
  requireAdmin();
  const req = sheetToObjects(SHEETS.REQUESTS).find(r => r.id === requestId);
  if (!req) return { success: false };
  updateById(SHEETS.REQUESTS, { id: requestId, status: 'denied', admin_notes: reason||'' });
  notify(req.guard_id, 'shift_denied', `Your shift request was not approved.${reason ? ' '+reason : ''}`);
  return { success: true };
}

// ── SWAPS ────────────────────────────────────────────────────

function getPendingSwaps() {
  return sheetToObjects(SHEETS.SWAPS).filter(s => s.status === 'pending_admin');
}
function getSwapsForGuard(guardId) {
  return sheetToObjects(SHEETS.SWAPS).filter(s => s.requestor_id === guardId || s.target_id === guardId);
}
function proposeSwap(myShiftId, targetGuardId) {
  const session = getUserSession();
  if (!session.authenticated) return { success: false };
  const shift = sheetToObjects(SHEETS.SHIFTS).find(s => s.id === myShiftId);
  if (!shift || shift.assigned_guard_id !== session.guard.id) return { success: false, message: 'Not your shift.' };
  appendRow(SHEETS.SWAPS, { id: uid('SW'), requestor_id: session.guard.id, target_id: targetGuardId,
    shift_id: myShiftId, status: 'pending_target', admin_notes: '', target_response: '',
    created_at: new Date().toISOString() });
  notify(targetGuardId, 'swap_proposed', `${session.guard.name} wants to swap a shift with you.`);
  return { success: true };
}
function respondSwap(swapId, accept) {
  const session = getUserSession();
  const swap = sheetToObjects(SHEETS.SWAPS).find(s => s.id === swapId);
  if (!swap || swap.target_id !== session.guard.id) return { success: false };
  if (!accept) {
    updateById(SHEETS.SWAPS, { id: swapId, status: 'target_declined', target_response: 'declined' });
    notify(swap.requestor_id, 'swap_declined', 'Your swap request was declined.');
    return { success: true };
  }
  updateById(SHEETS.SWAPS, { id: swapId, status: 'pending_admin', target_response: 'accepted' });
  getAllGuards().filter(g => roleFor(g.rank) === 'admin')
    .forEach(a => notify(a.id, 'swap_needs_approval', 'A shift swap needs your approval.'));
  return { success: true };
}
function approveSwap(swapId) {
  requireAdmin();
  const swap = sheetToObjects(SHEETS.SWAPS).find(s => s.id === swapId);
  if (!swap) return { success: false };
  const shift = sheetToObjects(SHEETS.SHIFTS).find(s => s.id === swap.shift_id);
  updateById(SHEETS.SHIFTS, { id: swap.shift_id, assigned_guard_id: swap.target_id });
  updateById(SHEETS.SWAPS, { id: swapId, status: 'approved' });
  notify(swap.requestor_id, 'swap_approved', `Your swap for ${shift ? shift.date : 'your shift'} was approved.`);
  notify(swap.target_id, 'swap_approved', `You've been assigned the swapped shift${shift ? ' on '+shift.date : ''}.`);
  return { success: true };
}
function denySwap(swapId, reason) {
  requireAdmin();
  const swap = sheetToObjects(SHEETS.SWAPS).find(s => s.id === swapId);
  if (!swap) return { success: false };
  updateById(SHEETS.SWAPS, { id: swapId, status: 'denied' });
  notify(swap.requestor_id, 'swap_denied', `Your swap was denied.${reason ? ' '+reason : ''}`);
  return { success: true };
}

// ── HOURS / OT ───────────────────────────────────────────────

function weeklyHoursForGuard(guardId, dateStr) {
  const bounds = workWeekBounds(dateStr);
  return getShiftsForGuard(guardId).reduce((total, s) => {
    const d = toYMD(s.date);
    if (d >= bounds.start && d <= bounds.end) {
      const t = templateByCode(s.template_code);
      if (t) total += parseFloat(t.paid_hours) || 0;
    }
    return total;
  }, 0);
}

function workWeekBounds(dateStr) {
  const s = toYMD(dateStr);
  if (!s) return { start: dateStr, end: dateStr };
  const [y,m,d] = s.split('-').map(Number);
  const dt = new Date(y, m-1, d, 12);
  const day = dt.getDay(); // 0=Sun,6=Sat
  const toSat = day === 6 ? 0 : -(day + 1);
  const sat = new Date(dt); sat.setDate(dt.getDate() + toSat);
  const fri = new Date(sat); fri.setDate(sat.getDate() + 6);
  return { start: serializeCell(sat), end: serializeCell(fri) };
}

function getWeeklySummary(periodId) {
  requireAdmin();
  const period = periodById(periodId);
  if (!period) return {};
  const shifts = getShiftsForPeriod(periodId);
  const summary = {};
  getAllGuards().forEach(g => { summary[g.id] = { guard: g, weeks: {} }; });
  shifts.forEach(s => {
    if (!s.assigned_guard_id || s.status === 'cancelled') return;
    const t = templateByCode(s.template_code);
    if (!t) return;
    const wk = workWeekBounds(s.date).start;
    if (!summary[s.assigned_guard_id]) return;
    summary[s.assigned_guard_id].weeks[wk] = (summary[s.assigned_guard_id].weeks[wk] || 0) + (parseFloat(t.paid_hours) || 0);
  });
  return summary;
}

// ── CONFIG ───────────────────────────────────────────────────

function getAllConfig() { return sheetToObjects(SHEETS.CONFIG); }
function getConfig(key) {
  const sheet = SH(SHEETS.CONFIG);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === key) return String(data[i][1]);
  }
  return null;
}
function setConfig(key, value) {
  requireAdmin();
  const sheet = SH(SHEETS.CONFIG);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === key) { sheet.getRange(i+1,2).setValue(value); return { success: true }; }
  }
  sheet.appendRow([key, value]);
  return { success: true };
}

// ── NOTIFICATIONS ────────────────────────────────────────────

function notify(guardId, type, message) {
  const guard = findGuardById(guardId);
  if (!guard) return;
  try {
    SH(SHEETS.NOTIFICATIONS).appendRow([uid('N'), guardId, type, message, new Date().toISOString(), 'email']);
    MailApp.sendEmail({
      to: guard.email,
      subject: `[7 Presidents STS] ${notifyLabel(type)}`,
      body: `Hi ${guard.name},\n\n${message}\n\nLog in: ${ScriptApp.getService().getUrl()}\n\n— 7 Presidents STS – Scheduling Hub for Ocean Rescue & Events`
    });
  } catch(e) { Logger.log('notify error: ' + e.message); }
}
function notifyLabel(type) {
  return ({ shift_approved:'Shift Approved', shift_denied:'Shift Denied', swap_proposed:'Swap Request',
    swap_approved:'Swap Approved', swap_declined:'Swap Declined', swap_denied:'Swap Denied',
    swap_needs_approval:'Swap Needs Approval', schedule_published:'Schedule Published',
    auto_clocked_out:'Auto Clock-Out Notice' })[type] || type;
}

// Default time-tracking config values (written on first run if missing)
function ensureTimeTrackingConfig() {
  const defaults = {
    hq_lat:             '40.2171',
    hq_lng:             '-74.0060',
    hq_name:            'Seven Presidents LG HQ',
    max_shift_hours:    '12',
    shift_reminder_min: '15',
  };
  const sheet = SH(SHEETS.CONFIG);
  const data = sheet.getDataRange().getValues();
  Object.entries(defaults).forEach(([k,v]) => {
    const exists = data.some(row => String(row[0]) === k);
    if (!exists) sheet.appendRow([k, v]);
  });
  return { success: true, message: 'Time tracking config initialized.' };
}
function publishSchedule(periodId) {
  requireAdmin();
  const p = periodById(periodId);
  if (!p) return { success: false };
  getAllGuards().forEach(g => notify(g.id, 'schedule_published',
    `The schedule for ${p.start_date} – ${p.schedule_thru} has been published.`));
  return { success: true };
}

// ── SETUP ────────────────────────────────────────────────────

function setupSpreadsheet() {
  const ss = SS();
  const schemas = {
    Guards:        ['id','name','email','phone','rank','pay_rate','post_eligibility','status','auth_type','password_hash','temp_password','must_change_pw','reset_token','reset_token_expires','invite_token','invite_token_expires','locker_number','parking_pass','created_at'],
    Posts:         ['id','name','rank_eligibility','active','color','sort_order'],
    ShiftTemplates:['code','name','start_time','end_time','paid_hours','break_minutes'],
    Shifts:        ['id','pay_period_id','date','post_id','template_code','type','assigned_guard_id','status','notes','created_at','series_id','custom_start','custom_end'],
    Availability:  ['id','guard_id','date','status','custom_start','custom_end','ot_willing','notes','submitted_at'],
    ShiftRequests: ['id','guard_id','shift_id','requested_at','status','admin_notes'],
    SwapRequests:  ['id','requestor_id','target_id','shift_id','status','admin_notes','target_response','created_at'],
    PayPeriods:    ['id','start_date','end_date','schedule_thru','locked','payroll_due','availability_deadline','time_report_due'],
    Config:        ['key','value'],
    Notifications: ['id','guard_id','type','message','sent_at','channel'],
    TimeRecords:   ['id','guard_id','guard_name','shift_id','date','clock_in','clock_out','break_minutes',
                    'total_minutes','status','edited','edited_by','edited_at','locked',
                    'clock_in_lat','clock_in_lng','clock_in_dist_ft','clock_in_location',
                    'clock_out_lat','clock_out_lng','clock_out_dist_ft','clock_out_location',
                    'auto_clocked_out','notes','created_at','stats_dismissed'],
    Locations:     ['id','name','polygon','active','notes'],
    ShiftStats:    ['id','guard_id','guard_name','time_record_id','shift_id','date','submitted_at',
                    'beach_location',
                    'preventive_actions','bather_assists','rfd_rescues','rfd_line_rescues',
                    'seabob_rescues','first_aid',
                    'incidents','incidents_desc',
                    'ten_55','ten_55_desc','ten_24','ten_24_desc',
                    'ten_34','ten_34_desc',
                    'patron_harassment','patron_harassment_desc',
                    'vandalism','vandalism_desc',
                    'training','surf_condition','wind_direction','air_temp','weather','flags_flown','notes'],
    Sessions:      ['token','guard_id','role','created_at','expires_at'],
    Badges:        ['id','guard_id','badge_key','earned_at','meta']
  };
  for (const [name, headers] of Object.entries(schemas)) {
    let sheet = ss.getSheetByName(name) || ss.insertSheet(name);
    sheet.getRange(1,1,1,headers.length).setValues([headers])
      .setBackground('#0d2137').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  const tmpl = ss.getSheetByName('ShiftTemplates');
  if (tmpl.getLastRow() <= 1) {
    tmpl.getRange(2,1,7,6).setValues([
      ['8HR', 'Regular 8hr',    '8:45 AM',  '5:15 PM', 8,  30],
      ['LS8', 'Late Shift 8hr', '10:45 AM','7:15 PM', 8,  30],
      ['LSO', 'Late Shift Only','5:15 PM', '7:15 PM', 2,  0 ],
      ['LS10','Late Shift 10hr','8:45 AM', '7:15 PM', 10, 30],
      ['SPEC', 'Special',    'TBD',     'TBD',     0,   0 ],
      ['TOURN','Tournament', '6:00 PM', '8:30 PM', 2.5, 0 ],
      ['ATH',  'Athlete Custom', 'Custom', 'Custom', 0, 0 ],
      ['MISC', 'Other',     'Custom',  'Custom',  0,   0 ],
    ]);
  }
  const cfg = ss.getSheetByName('Config');
  if (cfg.getLastRow() <= 1) {
    cfg.getRange(2,1,8,2).setValues([
      ['ot_threshold_hours','40'],['work_week_start','Saturday'],
      ['season_start','2026-05-23'],['season_end','2026-09-07'],
      ['app_name','7 Presidents STS'],['geolocation_required_phase2','true'],
      ['swap_requires_admin_approval','true'],['guard_self_select_max_hours','40'],
    ]);
  }
  Logger.log('Setup complete.');
}

// ── CLIENT API ───────────────────────────────────────────────

function clientGetSession()                         { return getUserSession(); }
function clientLoginWithPassword(email, password)   { 
  try {
    return loginWithPassword(email, password); 
  } catch(e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}
function clientChangePassword(token, cur, nw)       { return changePassword(token, cur, nw); }
function clientAdminSetPassword(token,guardId,password) { requireAdmin(token); return adminSetPassword(guardId,password); }
function clientRequestPasswordReset(email)          { return requestPasswordReset(email); }
function clientResetPasswordWithToken(token, pw)    { return resetPasswordWithToken(token, pw); }
function clientLogout(token)                        { return logoutSession(token); }
function clientValidateToken(token) {
  const session = getSessionFromToken(token);
  if (!session) return { authenticated: false };
  return { authenticated: true, guard: sanitizeGuard(session.guard), role: session.role };
}
function clientEnsureGuardColumns()                 { ensureGuardColumns_(); return {success:true}; }

// ── SHIFT STATS ───────────────────────────────────────────────

function submitShiftStats(token, d) {
  const session = getSessionFromToken(token);
  if (!session) return { success: false, message: 'Not authenticated.' };
  const guardId = String(session.guard.id);

  // Check if stats already exist for this time_record_id
  const existing = sheetToObjects(SHEETS.SHIFT_STATS).find(s =>
    s.time_record_id === d.time_record_id && String(s.guard_id) === guardId
  );
  if (existing) {
    // Update existing record
    updateById(SHEETS.SHIFT_STATS, { id: existing.id, ...d,
      guard_id: guardId, guard_name: session.guard.name,
      submitted_at: new Date().toISOString() });
    return { success: true, id: existing.id, updated: true };
  }

  const id = uid('SS');
  appendRow(SHEETS.SHIFT_STATS, {
    id,
    guard_id:              guardId,
    guard_name:            session.guard.name,
    time_record_id:        d.time_record_id || '',
    shift_id:              d.shift_id || '',
    date:                  d.date || '',
    submitted_at:          new Date().toISOString(),
    beach_location:        d.beach_location || '',
    preventive_actions:    d.preventive_actions || 0,
    bather_assists:        d.bather_assists || 0,
    rfd_rescues:           d.rfd_rescues || 0,
    rfd_line_rescues:      d.rfd_line_rescues || 0,
    seabob_rescues:        d.seabob_rescues || 0,
    first_aid:             d.first_aid || 0,
    incidents:             d.incidents || 0,
    incidents_desc:        d.incidents_desc || '',
    ten_55:                d.ten_55 || 0,
    ten_55_desc:           d.ten_55_desc || '',
    ten_24:                d.ten_24 || 0,
    ten_24_desc:           d.ten_24_desc || '',
    ten_34:                d.ten_34 || 0,
    ten_34_desc:           d.ten_34_desc || '',
    patron_harassment:     d.patron_harassment || 0,
    patron_harassment_desc:d.patron_harassment_desc || '',
    vandalism:             d.vandalism || 0,
    vandalism_desc:        d.vandalism_desc || '',
    training:              Array.isArray(d.training) ? d.training.join(', ') : (d.training || ''),
    surf_condition:        d.surf_condition || '',
    wind_direction:        d.wind_direction || '',
    air_temp:              d.air_temp || '',
    weather:               d.weather || '',
    flags_flown:           Array.isArray(d.flags_flown) ? d.flags_flown.join(', ') : (d.flags_flown || ''),
    notes:                 d.notes || ''
  });
  return { success: true, id };
}

function getMyShiftStats(token, periodId) {
  const session = getSessionFromToken(token);
  if (!session) return [];
  const guardId = String(session.guard.id);
  // Get all stats for this guard
  const all = sheetToObjects(SHEETS.SHIFT_STATS).filter(s => String(s.guard_id) === guardId);
  if (!periodId) return all;
  // Filter by period dates
  const p = periodById(periodId);
  if (!p) return all;
  const start = toYMD(p.start_date), end = toYMD(p.end_date);
  return all.filter(s => s.date >= start && s.date <= end);
}

function getPendingStatsAlerts(token) {
  // Returns time records that are complete but have no stats form
  const session = getSessionFromToken(token);
  if (!session) return [];
  const guardId = String(session.guard.id);

  const completedRecords = sheetToObjects(SHEETS.TIME_RECORDS).filter(r =>
    String(r.guard_id) === guardId && r.status === 'complete'
  );
  const submittedIds = new Set(
    sheetToObjects(SHEETS.SHIFT_STATS)
      .filter(s => String(s.guard_id) === guardId)
      .map(s => s.time_record_id)
  );

  return completedRecords.filter(r => !submittedIds.has(r.id) && !r.stats_dismissed);
}

function getAllShiftStats(token, periodId) {
  requireAdmin(token);
  if (!periodId) return sheetToObjects(SHEETS.SHIFT_STATS);
  const p = periodById(periodId);
  if (!p) return [];
  const start = toYMD(p.start_date), end = toYMD(p.end_date);
  return sheetToObjects(SHEETS.SHIFT_STATS).filter(s => s.date >= start && s.date <= end);
}

function exportShiftStatsCSV(token, periodId) {
  requireAdmin(token);
  const stats = getAllShiftStats(token, periodId);
  if (!stats.length) return { success: false, message: 'No stats found.' };
  const period = periodId ? periodById(periodId) : null;

  const headers = ['Guard','Date','Beach','Preventive','Bather Assists','RFD Rescues',
    'RFD+Line','Seabob','First Aid','Incidents','Incidents Desc',
    '10-55','10-55 Desc','10-24','10-24 Desc','10-34','10-34 Desc',
    'Patron Harassment','Harassment Desc','Vandalism','Vandalism Desc',
    'Training','Surf','Wind','Air Temp','Weather','Flags','Notes','Submitted'];

  const rows = [headers, ...stats.map(s => [
    s.guard_name||s.guard_id, s.date, s.beach_location,
    s.preventive_actions||0, s.bather_assists||0, s.rfd_rescues||0,
    s.rfd_line_rescues||0, s.seabob_rescues||0, s.first_aid||0,
    s.incidents||0, s.incidents_desc||'',
    s.ten_55||0, s.ten_55_desc||'',
    s.ten_24||0, s.ten_24_desc||'',
    s.ten_34||0, s.ten_34_desc||'',
    s.patron_harassment||0, s.patron_harassment_desc||'',
    s.vandalism||0, s.vandalism_desc||'',
    s.training||'', s.surf_condition||'', s.wind_direction||'',
    s.air_temp||'', s.weather||'', s.flags_flown||'', s.notes||'',
    s.submitted_at||''
  ])];

  const csv = rows.map(r => r.map(c => `"${String(c).replace(/"/g,'""')}"`).join(',')).join('\n');
  return { success: true, csv, period: period ? `${period.start_date}_${period.end_date}` : 'all' };
}

function getAdminStatsData(token, date) {
  requireAdmin(token);
  const today = toYMD(new Date());
  const queryDate = date || today;

  const allStats = sheetToObjects(SHEETS.SHIFT_STATS);
  const statsForDate = allStats.filter(s => toYMD(s.date) === queryDate);
  const statsToday   = allStats.filter(s => toYMD(s.date) === today);

  const submittedTRIds = new Set(allStats.map(s => s.time_record_id).filter(Boolean));
  // Fallback: guard+date pairs — covers any stats regardless of whether time_record_id is linked correctly
  const submittedGuardDates = new Set(allStats.map(s => `${s.guard_id}|${toYMD(s.date)}`));
  const allTimeRecords = sheetToObjects(SHEETS.TIME_RECORDS);

  // missingToday always reflects today regardless of selected date
  const dateCompleteRecords = allTimeRecords.filter(r =>
    r.status === 'complete' && toYMD(r.date) === today
  );
  const missingToday = dateCompleteRecords
    .filter(r => !submittedTRIds.has(r.id) && !submittedGuardDates.has(`${r.guard_id}|${today}`))
    .map(r => ({
      guard_id:       r.guard_id,
      guard_name:     r.guard_name || r.guard_id,
      time_record_id: r.id,
      clock_out:      r.clock_out ? new Date(r.clock_out).toISOString() : '',
      date:           toYMD(r.date)
    }));

  const latestByGuard = {};
  allTimeRecords.filter(r => r.status === 'complete').forEach(r => {
    const gid = String(r.guard_id);
    const co = r.clock_out ? new Date(r.clock_out).getTime() : 0;
    if (!latestByGuard[gid] || co > latestByGuard[gid]._t) {
      latestByGuard[gid] = { ...r, _t: co };
    }
  });
  const missingLastShift = Object.values(latestByGuard)
    .filter(r => !submittedTRIds.has(r.id) && !submittedGuardDates.has(`${r.guard_id}|${toYMD(r.date)}`) && r.stats_dismissed !== true && r.stats_dismissed !== 'true')
    .map(r => {
      const guard = findGuardById(r.guard_id);
      return {
        guard_id:       r.guard_id,
        guard_name:     guard ? guard.name : (r.guard_name || r.guard_id),
        time_record_id: r.id,
        clock_out:      r.clock_out ? new Date(r.clock_out).toISOString() : '',
        date:           toYMD(r.date)
      };
    });

  const statKeys = ['preventive_actions','bather_assists','rfd_rescues','rfd_line_rescues',
                    'seabob_rescues','first_aid','incidents','ten_55','ten_24','ten_34',
                    'patron_harassment','vandalism'];
  const seasonTotals = {};
  statKeys.forEach(k => { seasonTotals[k] = 0; });
  const guardTotalsMap = {};
  allStats.forEach(s => {
    statKeys.forEach(k => { seasonTotals[k] += parseInt(s[k]) || 0; });
    const gid = String(s.guard_id);
    if (!guardTotalsMap[gid]) {
      guardTotalsMap[gid] = { guard_id: gid, guard_name: s.guard_name || gid, cards: 0 };
      statKeys.forEach(k => { guardTotalsMap[gid][k] = 0; });
    }
    guardTotalsMap[gid].cards++;
    statKeys.forEach(k => { guardTotalsMap[gid][k] += parseInt(s[k]) || 0; });
  });
  const guardTotals = Object.values(guardTotalsMap)
    .sort((a, b) => (b.bather_assists + b.rfd_rescues) - (a.bather_assists + a.rfd_rescues));

  return {
    date:           queryDate,
    today,
    statsForDate,
    statsToday,
    missingToday,
    missingLastShift,
    seasonTotals,
    guardTotals,
    totalCards:     allStats.length
  };
}

function clientSubmitShiftStats(token, d)          { return submitShiftStats(token, d); }
function clientGetMyShiftStats(token, periodId)    { return getMyShiftStats(token, periodId); }
function clientGetPendingStatsAlerts(token)        { return getPendingStatsAlerts(token); }
function clientDismissStatsAlert(token, timeRecordId) {
  const session = getSessionFromToken(token);
  if (!session) return { success: false };
  const guardId = String(session.guard.id);
  const rec = sheetToObjects(SHEETS.TIME_RECORDS).find(r => String(r.id) === String(timeRecordId));
  if (!rec || String(rec.guard_id) !== guardId) return { success: false };
  updateById(SHEETS.TIME_RECORDS, { id: timeRecordId, stats_dismissed: true });
  return { success: true };
}
function clientGetActiveClockins(token) {
  requireAdmin(token);
  const open = sheetToObjects(SHEETS.TIME_RECORDS).filter(r => r.status === 'open');
  return open.map(r => {
    const guard = findGuardById(r.guard_id);
    return {
      time_record_id: r.id,
      guard_id:   r.guard_id,
      guard_name: guard ? guard.name : (r.guard_name || r.guard_id),
      guard_rank: guard ? guard.rank : '',
      clock_in:   r.clock_in,
      date:       r.date,
      shift_id:   r.shift_id || ''
    };
  });
}
function clientGetAllShiftStats(token, periodId)   { return getAllShiftStats(token, periodId); }
function clientExportShiftStatsCSV(token, periodId){ return exportShiftStatsCSV(token, periodId); }
function clientGetAdminStatsData(token, date)      { return getAdminStatsData(token, date); }

function nudgeGuardStats(token, guardId, date) {
  requireAdmin(token);
  const guard = findGuardById(guardId);
  if (!guard || !guard.email) return { success: false, message: 'No email on file for this guard.' };

  const d = new Date(date + 'T12:00:00');
  const DAYS_L  = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
  const MONTHS_L = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  const dateFmt = `${DAYS_L[d.getDay()]}, ${MONTHS_L[d.getMonth()]} ${d.getDate()}`;
  const name = guard.name.split(' ')[0];

  const phrases = [
    `Psst, ${name} — your shift on ${dateFmt} is almost officially in the books... just need that stats card from you!`,
    `${name}, the seagulls filed their report. The piping plovers filed their report. Just need yours for ${dateFmt}! 📋`,
    `Tower's closed, whistle's down, ${name} — just your stats card for ${dateFmt} left to go! 🏄`,
    `${name} — beaches are clear, guards are home, but your ${dateFmt} stats card is 10-55! Let's start the search.`,
    `Quick one, ${name} — your stats card for ${dateFmt} is MIA. The Supe would love to see it when you get a chance.`,
    `All clear on the beach, ${name}! Just your stats card for ${dateFmt} left to call it a day. 🏖️`,
    `${name}, consider this a friendly tap on the shoulder — your stats card for ${dateFmt} hasn't made it in yet!`,
    `Your shift rocked, ${name}. Now your stats card for ${dateFmt} just needs to find its way home. 📋`,
    `${name} — one small thing between a great shift and a complete one: your stats card for ${dateFmt}. You've got this. 🚩`
  ];
  const body = phrases[Math.floor(Math.random() * phrases.length)];
  const htmlBody = `<p>${body}</p><p style="margin-top:16px"><a href="https://baconlt.github.io/7p-sts/">Submit your stats card →</a></p><p style="color:#888;font-size:.8em;margin-top:24px">— 7 Presidents STS</p>`;

  try {
    MailApp.sendEmail({
      to: guard.email,
      subject: `Stats card reminder — ${dateFmt}`,
      body: body + '\n\nSubmit here: https://baconlt.github.io/7p-sts/\n\n— 7 Presidents STS',
      htmlBody
    });
    return { success: true };
  } catch(e) {
    return { success: false, message: e.message };
  }
}
function clientNudgeGuardStats(token, guardId, date) { return nudgeGuardStats(token, guardId, date); }

// ── BADGES / ACHIEVEMENTS ─────────────────────────────────────
// All criteria are derived from data the app already stores (TimeRecords,
// ShiftStats, Shifts/ShiftTemplates). Earned badges are persisted to the
// Badges sheet so we can detect newly-earned ones and keep a permanent record.

// Registry — single source of truth. Each def's earned()/progress() receive the
// per-guard ctx built by buildBadgeCtx_(). To add a badge, add one object here.
const BADGE_DEFS = [
  { key:'stats_streak_5',  name:'On the Record', emoji:'🗒️', category:'Diligence',
    desc:'5 shifts in a row with stats submitted',
    earned: c => c.statsStreak >= 5,
    progress: c => ({ current: Math.min(c.statsStreak, 5), target: 5 }) },
  { key:'stats_streak_10', name:'Diligent', emoji:'📋', category:'Diligence',
    desc:'10 shifts in a row with stats submitted',
    earned: c => c.statsStreak >= 10,
    progress: c => ({ current: Math.min(c.statsStreak, 10), target: 10 }) },
  { key:'stats_streak_20', name:'Meticulous', emoji:'📊', category:'Diligence',
    desc:'20 shifts in a row with stats submitted',
    earned: c => c.statsStreak >= 20,
    progress: c => ({ current: Math.min(c.statsStreak, 20), target: 20 }) },
  { key:'stats_streak_50', name:'Statistician', emoji:'🏆', category:'Diligence',
    desc:'50 shifts in a row with stats submitted',
    earned: c => c.statsStreak >= 50,
    progress: c => ({ current: Math.min(c.statsStreak, 50), target: 50 }) },
  { key:'late_shift_10hr', name:'Closing Time', emoji:'🌙', category:'Endurance',
    desc:'Worked a 10-hour late shift without an auto clock-out',
    earned: c => c.tenHourNoAuto >= 1,
    progress: c => ({ current: Math.min(c.tenHourNoAuto, 1), target: 1 }) },
  { key:'iron_week', name:'Iron Week', emoji:'💪', category:'Endurance',
    desc:'5 ten-hour shifts in a single work week',
    earned: c => c.maxTenHourWeek >= 5,
    progress: c => ({ current: Math.min(c.maxTenHourWeek, 5), target: 5 }) },
  { key:'ten_in_a_row', name:'Unbroken', emoji:'🔥', category:'Endurance',
    desc:'Worked 10 days in a row',
    earned: c => c.maxConsecutiveDays >= 10,
    progress: c => ({ current: Math.min(c.maxConsecutiveDays, 10), target: 10 }) }
];

// Build the per-guard context once, then evaluate every def against it.
function buildBadgeCtx_(guardId) {
  const gid = String(guardId);
  const completed = sheetToObjects(SHEETS.TIME_RECORDS)
    .filter(r => String(r.guard_id) === gid && r.status === 'complete')
    .sort((a, b) => {
      const d = toYMD(a.date).localeCompare(toYMD(b.date));
      return d !== 0 ? d : String(a.clock_in).localeCompare(String(b.clock_in));
    });

  const statsTRIds = new Set(
    sheetToObjects(SHEETS.SHIFT_STATS)
      .filter(s => String(s.guard_id) === gid)
      .map(s => s.time_record_id)
  );

  // Map shift_id → 10-hour? (template LS10 or paid_hours >= 10)
  const shiftsById = {};
  sheetToObjects(SHEETS.SHIFTS).forEach(s => { shiftsById[s.id] = s; });
  const isTenHour = (rec) => {
    const shift = rec.shift_id ? shiftsById[rec.shift_id] : null;
    if (shift) {
      if (shift.template_code === 'LS10') return true;
      const tmpl = templateByCode(shift.template_code);
      if (tmpl && parseFloat(tmpl.paid_hours) >= 10) return true;
    }
    // Fallback: actual logged time ≥ ~10h (paid minutes, breaks already removed)
    return (parseInt(rec.total_minutes) || 0) >= 600;
  };

  // Current stats streak — most recent consecutive completed records that have
  // stats. Trailing records still awaiting a card (no stats, not dismissed) are
  // skipped rather than breaking the streak, so finishing a shift doesn't briefly
  // zero out the count before the guard fills the card in. A dismissed/forgotten
  // shift does end the streak.
  let i = completed.length - 1;
  while (i >= 0 && !statsTRIds.has(completed[i].id)
         && String(completed[i].stats_dismissed) !== 'true') i--;
  let statsStreak = 0;
  for (; i >= 0; i--) {
    if (statsTRIds.has(completed[i].id)) statsStreak++;
    else break;
  }

  // 10-hour shifts with no auto clock-out (count, for the milestone badge).
  const tenHourNoAuto = completed.filter(r =>
    isTenHour(r) && String(r.auto_clocked_out) !== 'true').length;

  // Most 10-hour shifts in any single Sat–Fri work week.
  const tenHourByWeek = {};
  completed.forEach(r => {
    if (!isTenHour(r)) return;
    const wk = workWeekBounds(r.date).start;
    tenHourByWeek[wk] = (tenHourByWeek[wk] || 0) + 1;
  });
  const maxTenHourWeek = Object.values(tenHourByWeek).reduce((m, v) => Math.max(m, v), 0);

  // Longest run of consecutive calendar days worked.
  const workedDays = [...new Set(completed.map(r => toYMD(r.date)))].filter(Boolean).sort();
  let maxConsecutiveDays = 0, run = 0, prev = null;
  workedDays.forEach(day => {
    if (prev && isNextDay_(prev, day)) run++; else run = 1;
    if (run > maxConsecutiveDays) maxConsecutiveDays = run;
    prev = day;
  });

  return { statsStreak, tenHourNoAuto, maxTenHourWeek, maxConsecutiveDays };
}

// True if dayB is the calendar day immediately after dayA (both YYYY-MM-DD).
function isNextDay_(dayA, dayB) {
  const [ay, am, ad] = dayA.split('-').map(Number);
  const next = new Date(ay, am - 1, ad, 12);
  next.setDate(next.getDate() + 1);
  return serializeCell(next) === dayB;
}

function ensureBadgesSheet_() {
  let sheet = SH(SHEETS.BADGES);
  if (!sheet) {
    sheet = SS().insertSheet(SHEETS.BADGES);
    const headers = ['id','guard_id','badge_key','earned_at','meta'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground('#0d2137').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Evaluate all badges for the logged-in guard, persist any newly-earned ones,
// and return the full display state plus the keys that were just earned.
function getBadgeState(token) {
  const session = getSessionFromToken(token);
  if (!session) return { badges: [], newlyEarned: [] };
  const guardId = String(session.guard.id);
  ensureBadgesSheet_();

  const ctx = buildBadgeCtx_(guardId);
  const owned = {};
  sheetToObjects(SHEETS.BADGES)
    .filter(b => String(b.guard_id) === guardId)
    .forEach(b => { owned[b.badge_key] = b.earned_at; });

  const newlyEarned = [];
  const badges = BADGE_DEFS.map(def => {
    const isEarned = !!def.earned(ctx);
    let earnedAt = owned[def.key] || '';
    if (isEarned && !owned.hasOwnProperty(def.key)) {
      earnedAt = new Date().toISOString();
      appendRow(SHEETS.BADGES, {
        id: uid('BDG'), guard_id: guardId, badge_key: def.key,
        earned_at: earnedAt, meta: ''
      });
      owned[def.key] = earnedAt;
      newlyEarned.push(def.key);
    }
    return {
      key: def.key, name: def.name, emoji: def.emoji,
      desc: def.desc, category: def.category,
      earned: isEarned, earned_at: earnedAt,
      progress: def.progress ? def.progress(ctx) : null
    };
  });

  return { badges, newlyEarned };
}

// Admin: badge counts per active guard, most badges first.
function getBadgeLeaderboard(token) {
  requireAdmin(token);
  ensureBadgesSheet_();
  const rows = sheetToObjects(SHEETS.BADGES);
  const byGuard = {};
  rows.forEach(b => {
    const gid = String(b.guard_id);
    (byGuard[gid] = byGuard[gid] || []).push(b.badge_key);
  });
  const defByKey = {};
  BADGE_DEFS.forEach(d => { defByKey[d.key] = d; });

  return getAllGuards().map(g => {
    const keys = byGuard[String(g.id)] || [];
    return {
      guard_id: g.id,
      name: g.name,
      count: keys.length,
      badges: keys.map(k => defByKey[k]
        ? { key: k, name: defByKey[k].name, emoji: defByKey[k].emoji }
        : { key: k, name: k, emoji: '🏅' })
    };
  }).sort((a, b) => b.count - a.count || a.name.localeCompare(b.name));
}

function clientGetBadgeState(token)       { return getBadgeState(token); }
function clientGetBadgeLeaderboard(token) { return getBadgeLeaderboard(token); }

function clientAdminDismissLastShift(token, timeRecordId) {
  requireAdmin(token);
  updateById(SHEETS.TIME_RECORDS, { id: timeRecordId, stats_dismissed: true });
  return { success: true };
}
function clientGetGuards()                 { return getAllGuards(); }
function clientGetPosts()                  { return getAllPosts(); }
function clientGetTemplates()              { return getAllTemplates(); }
function clientGetRawTemplate(code) {
  // Debug: returns the raw cell values for a template row
  const sheet = SH(SHEETS.TEMPLATES);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const codeIdx = headers.indexOf('code');
  for (let i=1; i<data.length; i++) {
    if (String(data[i][codeIdx]) === code) {
      const row = {};
      headers.forEach((h,j) => { row[h] = { value: data[i][j], type: typeof data[i][j], str: String(data[i][j]) }; });
      return row;
    }
  }
  return null;
}
function clientMigrateAvailability() {
  // Run once to migrate old availability status values to new shift-type based values
  const STATUS_MAP = {
    'WORK':      'S_8HR',
    'OT_AVAIL':  'S_8HR_OT',
    'S_8HR_LS8': 'S_LS8',
    'S_8HR_LSO': 'S_LSO',
    'SPECIAL':   'S_SPEC',
    'FLEX_OFF':  'FLEX_OFF',
    'FIXED_OFF': 'FIXED_OFF',
    'UNAVAIL':   'FIXED_OFF',
  };
  const sheet = SH(SHEETS.AVAILABILITY);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const statusIdx = headers.indexOf('status');
  if (statusIdx < 0) return { success: false, message: 'status column not found' };
  let count = 0;
  for (let i = 1; i < data.length; i++) {
    const old = String(data[i][statusIdx]).trim();
    const newVal = STATUS_MAP[old];
    if (newVal && newVal !== old) {
      sheet.getRange(i + 1, statusIdx + 1).setValue(newVal);
      count++;
    }
  }
  return { success: true, message: `Migrated ${count} rows.` };
}
function clientFixTemplateTimes() {
  // Run once to correct template times in the sheet
  requireAdmin();
  const correct = {
    '8HR': {name:'Regular 8hr',   start:'8:45 AM', end:'5:15 PM', hours:8,  brk:30},
    'LS8': {name:'Late Shift 8hr',start:'10:45 AM',end:'7:15 PM', hours:8,  brk:30},
    'LSO': {name:'Late Shift Only',start:'5:15 PM', end:'7:15 PM', hours:2,  brk:0},
    'LS10':{name:'Late Shift 10hr',start:'8:45 AM', end:'7:15 PM', hours:10, brk:30},
  };
  const sheet = SH(SHEETS.TEMPLATES);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const codeIdx = headers.indexOf('code');
  for (let i=1; i<data.length; i++) {
    const code = data[i][codeIdx];
    if (correct[code]) {
      const c = correct[code];
      const row = data[i];
      // Update name, start_time, end_time, paid_hours, break_minutes
      ['name','start_time','end_time','paid_hours','break_minutes'].forEach((col,j) => {
        const idx = headers.indexOf(col);
        if (idx>=0) sheet.getRange(i+1, idx+1).setValue([c.name,c.start,c.end,c.hours,c.brk][j]);
      });
    }
  }
  return {success:true, message:'Template times updated.'};
}
function clientGetPeriods()                { return getAllPeriods(); }
function clientGetActivePeriod()           { return getActivePeriod(); }
function clientGetShifts(pid)             { return getShiftsForPeriod(pid); }
function clientGetMyShifts(gid)           { return getShiftsForGuard(gid); }
function clientGetOpenShifts(gid)         { return getOpenShiftsForGuard(gid); }
function clientGetAvailability(pid)       { return getAvailabilityForPeriod(pid); }
function clientGetMyAvailability(gid)     { return getAvailabilityForGuard(gid); }
function clientGetPendingRequests()       { return getPendingRequests(); }
function clientGetPendingSwaps()          { return getPendingSwaps(); }
function clientGetMyRequests(gid)         { return getRequestsForGuard(gid); }
function clientGetMySwaps(gid)            { return getSwapsForGuard(gid); }
function clientGetWeeklySummary(pid)      { return getWeeklySummary(pid); }
function clientGetConfig()                { return getAllConfig(); }
function clientCreateGuard(token,d)        { requireAdmin(token); return createGuard(d); }
function clientUpdateGuard(token,d)        { requireAdmin(token); return updateGuard(d); }
function clientDeactivateGuard(token,id)   { requireAdmin(token); return deactivateGuard(id); }
function clientCreatePost(token,d)         { requireAdmin(token); return createPost(d); }
function clientUpdatePost(token,d)         { requireAdmin(token); return updatePost(d); }
function clientCreateShift(token,d)        { requireAdmin(token); return createShift(d); }
function clientAssignShift(token,d)        { requireAdmin(token); return assignShift(d); }
function clientCheckOpenSlot(token,date,pid,tc) { requireAdmin(token); return checkOpenSlot(date,pid,tc); }
function clientUpdateShift(token,d)        { requireAdmin(token); return updateShift(d); }
function clientCancelShift(token,id)       { requireAdmin(token); return cancelShift(id); }
function clientDeleteShiftSeries(token,sid,from,scope) { requireAdmin(token); return deleteShiftSeries(sid,from,scope); }
function clientBulkCreateShifts(token,d) { requireAdmin(token);
  const { pay_period_id, post_id, template_code, start_date, end_date,
          days_of_week, qty_per_day, notes } = d;
  if (!pay_period_id||!post_id||!template_code||!start_date||!end_date)
    return { success: false, message: 'Missing required fields.' };
  const dows = days_of_week || [1,2,3,4,5];
  const qty = parseInt(qty_per_day)||1;
  const [sy,sm,sday] = start_date.split('-').map(Number);
  const [ey,em,eday] = end_date.split('-').map(Number);
  let cur = new Date(sy,sm-1,sday,12);
  const endDt = new Date(ey,em-1,eday,12);
  let count = 0;
  while (cur <= endDt) {
    if (dows.includes(cur.getDay())) {
      for (let i=0; i<qty; i++) {
        const id = uid('S');
        appendRow(SHEETS.SHIFTS, {
          id, pay_period_id, date: serializeCell(cur),
          post_id, template_code, type: 'open',
          assigned_guard_id: '', status: 'open',
          notes: notes||'', created_at: new Date().toISOString(), series_id: ''
        });
        count++;
      }
    }
    cur.setDate(cur.getDate()+1);
  }
  return { success: true, count };
}
function clientCreatePeriod(token,d)       { requireAdmin(token); return createPeriod(d); }
function clientUpdatePeriod(token,d)       { requireAdmin(token); return updatePeriod(d); }
function clientLockPeriod(token,id)        { requireAdmin(token); return lockPeriod(id); }
function clientPublishSchedule(token,id)   { requireAdmin(token); return publishSchedule(id); }
function clientSubmitAvailability(token,e) { return submitAvailability(token, e); }
function clientRequestShift(sid)          { return requestShift(sid); }
function clientApproveRequest(rid)        { return approveRequest(rid); }
function clientDenyRequest(rid,reason)    { return denyRequest(rid,reason); }
function clientProposeSwap(sid,tid)       { return proposeSwap(sid,tid); }
function clientRespondSwap(sid,accept)    { return respondSwap(sid,accept); }
function clientApproveSwap(sid)           { return approveSwap(sid); }
function clientDenySwap(sid,reason)       { return denySwap(sid,reason); }
function clientSetConfig(token,k,v)        { requireAdmin(token); return setConfig(k,v); }
function clientSendInvite(token,guardId)   { try { return sendInvite(token,guardId); } catch(e) { return { success:false, message:e.message }; } }
function clientAcceptInvite(tok,pw)        { try { return acceptInvite(tok,pw); } catch(e) { return { success:false, message:e.message }; } }
// clientLoginWithPassword defined above

// ── TIME TRACKING ─────────────────────────────────────────────

// Config helpers for time tracking
function ttConfig(key, fallback) {
  const cfg = getConfig(key);
  return cfg !== null && cfg !== '' ? cfg : fallback;
}

// Returns today's open time record for this guard if exists
function getOpenTimeRecord(guardId) {
  const today = toYMD(new Date());
  const records = sheetToObjects(SHEETS.TIME_RECORDS);
  return records.find(r =>
    String(r.guard_id) === String(guardId) &&
    toYMD(r.date) === today &&
    r.status === 'open'
  ) || null;
}

// Get all time records for a pay period
function getTimeRecordsForPeriod(periodId) {
  const p = periodById(periodId);
  if (!p) return [];
  const start = toYMD(p.start_date);
  const end   = toYMD(p.schedule_thru) || toYMD(p.end_date); // use schedule_thru, fall back to end_date
  return sheetToObjects(SHEETS.TIME_RECORDS).filter(r => {
    const d = toYMD(r.date);
    return d >= start && d <= end;
  });
}

// Get time records for a specific guard
function getTimeRecordsForGuard(guardId, periodId) {
  const all = getTimeRecordsForPeriod(periodId);
  return all.filter(r => String(r.guard_id) === String(guardId));
}

// Calculate distance in feet between two lat/lng points
function distanceFeet(lat1, lng1, lat2, lng2) {
  if (!lat1 || !lat2) return null;
  const R = 20902231; // Earth radius in feet
  const dLat = (lat2 - lat1) * Math.PI / 180;
  const dLng = (lng2 - lng1) * Math.PI / 180;
  const a = Math.sin(dLat/2) * Math.sin(dLat/2) +
    Math.cos(lat1 * Math.PI/180) * Math.cos(lat2 * Math.PI/180) *
    Math.sin(dLng/2) * Math.sin(dLng/2);
  return Math.round(R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a)));
}

// ── GEOFENCE HELPERS ─────────────────────────────────────────

function pointInPolygon(lat, lng, polygonJson) {
  let poly;
  try { poly = JSON.parse(polygonJson); } catch(e) { return false; }
  if (!Array.isArray(poly) || poly.length < 3) return false;
  let inside = false;
  for (let i = 0, j = poly.length - 1; i < poly.length; j = i++) {
    const xi = poly[i][0], yi = poly[i][1];
    const xj = poly[j][0], yj = poly[j][1];
    const intersect = ((yi > lng) !== (yj > lng)) &&
      (lat < (xj - xi) * (lng - yi) / (yj - yi) + xi);
    if (intersect) inside = !inside;
  }
  return inside;
}

function resolveLocation(lat, lng) {
  if (!lat || !lng) return '';
  const locs = sheetToObjects(SHEETS.LOCATIONS).filter(l => l.active !== 'false');
  for (const loc of locs) {
    if (pointInPolygon(parseFloat(lat), parseFloat(lng), loc.polygon)) return loc.name;
  }
  return '';
}

// Locations CRUD
function getLocations() {
  return sheetToObjects(SHEETS.LOCATIONS);
}

function createLocation(d) {
  const id = uid('LOC');
  appendRow(SHEETS.LOCATIONS, {
    id, name: d.name, polygon: d.polygon, active: 'true', notes: d.notes || ''
  });
  return { success: true, id };
}

function updateLocation(d) {
  const updates = { id: d.id };
  if (d.name    !== undefined) updates.name    = d.name;
  if (d.polygon !== undefined) updates.polygon = d.polygon;
  if (d.active  !== undefined) updates.active  = d.active;
  if (d.notes   !== undefined) updates.notes   = d.notes;
  updateById(SHEETS.LOCATIONS, updates);
  return { success: true };
}

function deleteLocation(id) {
  const sheet = SH(SHEETS.LOCATIONS);
  const data  = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(id)) { sheet.deleteRow(i + 1); break; }
  }
  return { success: true };
}

function clientGetLocations(token)         { requireAdmin(token); return getLocations(); }
function clientCreateLocation(token, d)    { requireAdmin(token); return createLocation(d); }
function clientUpdateLocation(token, d)    { requireAdmin(token); return updateLocation(d); }
function clientDeleteLocation(token, id)   { requireAdmin(token); return deleteLocation(id); }

// Clock in
function clockIn(d) {
  const session = getSessionFromToken(d.token);
  if (!session) return { success: false, message: 'Not authenticated. Please log in again.' };

  // Admin can clock in on behalf of another guard
  let guardId, guard;
  if (d.target_guard_id && session.role === 'admin') {
    guard = findGuardById(d.target_guard_id);
    if (!guard) return { success: false, message: 'Guard not found.' };
    guardId = String(guard.id);
  } else {
    guardId = String(session.guard.id);
    guard   = session.guard;
  }

  // Use a script lock to prevent duplicate clock-ins from simultaneous requests
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    // Block if already clocked in today (checked inside lock to prevent race)
    const open = getOpenTimeRecord(guardId);
    if (open) return { success: false, message: 'You are already clocked in.', record: open };

    // Find matching shift if any
    const today = toYMD(new Date());
    const shifts = sheetToObjects(SHEETS.SHIFTS).filter(s =>
      toYMD(s.date) === today &&
      String(s.assigned_guard_id) === guardId &&
      s.status !== 'cancelled'
    );
    const shift = shifts[0] || null;

    // Reference location from config
    const hqLat = parseFloat(ttConfig('hq_lat', '40.2171'));
    const hqLng = parseFloat(ttConfig('hq_lng', '-74.0060'));

    const distFt = d.lat && d.lng ? distanceFeet(d.lat, d.lng, hqLat, hqLng) : null;
    const inLocation = d.lat && d.lng ? resolveLocation(d.lat, d.lng) : '';

    const now = new Date();
    const id = uid('TR');
    appendRow(SHEETS.TIME_RECORDS, {
      id,
      guard_id:          guardId,
      guard_name:        guard.name || '',
      shift_id:          shift ? shift.id : '',
      date:              today,
      clock_in:          now.toISOString(),
      clock_out:         '',
      break_minutes:     '',
      total_minutes:     '',
      status:            'open',
      edited:            'false',
      edited_by:         '',
      edited_at:         '',
      locked:            'false',
      clock_in_lat:      d.lat || '',
      clock_in_lng:      d.lng || '',
      clock_in_dist_ft:  distFt !== null ? distFt : '',
      clock_in_location: inLocation,
      clock_out_lat:     '',
      clock_out_lng:     '',
      clock_out_dist_ft: '',
      clock_out_location:'',
      auto_clocked_out:  'false',
      notes:             d.notes || '',
      created_at:        now.toISOString()
    });

    return { success: true, id, clock_in: now.toISOString(), shift_id: shift ? shift.id : '' };
  } finally {
    lock.releaseLock();
  }
}

// Clock out
function clockOut(d) {
  const session = getSessionFromToken(d.token);
  if (!session) return { success: false, message: 'Not authenticated. Please log in again.' };

  // Admin can clock out on behalf of another guard
  let guardId;
  if (d.target_guard_id && session.role === 'admin') {
    guardId = String(d.target_guard_id);
  } else {
    guardId = String(session.guard.id);
  }

  const record = d.record_id
    ? sheetToObjects(SHEETS.TIME_RECORDS).find(r => r.id === d.record_id)
    : getOpenTimeRecord(guardId);

  if (!record) return { success: false, message: 'No open clock-in found.' };
  if (record.locked === 'true') return { success: false, message: 'This record is locked.' };

  const now = new Date();
  const clockIn = d.clock_in_override
    ? new Date(d.clock_in_override)
    : new Date(record.clock_in);

  // Use provided time or now
  const clockOutTime = d.clock_out_override
    ? new Date(d.clock_out_override)
    : now;

  // If clock-in was changed, mark as edited
  const clockInChanged = !!d.clock_in_override;

  // Reference location
  const hqLat = parseFloat(ttConfig('hq_lat', '40.2171'));
  const hqLng = parseFloat(ttConfig('hq_lng', '-74.0060'));
  const distFt = d.lat && d.lng ? distanceFeet(d.lat, d.lng, hqLat, hqLng) : null;
  const outLocation = d.lat && d.lng ? resolveLocation(d.lat, d.lng) : '';

  // Calculate total minutes
  const rawMinutes = Math.round((clockOutTime - clockIn) / 60000);

  // Break: use provided break_minutes, else auto-deduct 30 if worked 4+ hours
  let breakMins = d.break_minutes !== undefined && d.break_minutes !== ''
    ? parseInt(d.break_minutes) || 0
    : (rawMinutes >= 240 ? 30 : 0); // 30 min auto if 4+ hours

  const totalMinutes = Math.max(0, rawMinutes - breakMins);
  const isEdited = !!d.clock_out_override || clockInChanged;

  updateById(SHEETS.TIME_RECORDS, {
    id:               record.id,
    clock_in:         clockIn.toISOString(),
    clock_out:        clockOutTime.toISOString(),
    break_minutes:    breakMins,
    total_minutes:    totalMinutes,
    status:           'complete',
    edited:           isEdited ? 'true' : 'false',
    edited_by:        isEdited ? guardId : '',
    edited_at:        isEdited ? now.toISOString() : '',
    clock_out_lat:      d.lat || '',
    clock_out_lng:      d.lng || '',
    clock_out_dist_ft:  distFt !== null ? distFt : '',
    clock_out_location: outLocation,
    notes:              d.notes || record.notes || ''
  });

  // Compute rounded_minutes: round each timestamp to nearest 15, then diff
  const cinRnd = new Date(clockIn); cinRnd.setMinutes(Math.round(cinRnd.getMinutes()/15)*15, 0, 0);
  const coutRnd = new Date(clockOutTime); coutRnd.setMinutes(Math.round(coutRnd.getMinutes()/15)*15, 0, 0);
  const roundedMinutes = Math.max(0, Math.round((coutRnd - cinRnd)/60000) - breakMins);

  return { success: true, total_minutes: totalMinutes, break_minutes: breakMins, rounded_minutes: roundedMinutes };
}

// Admin: edit a time record
function editTimeRecord(d) {
  // token passed via requireAdmin in client wrapper
  const session = { guard: { name: 'Admin' } };
  const record = sheetToObjects(SHEETS.TIME_RECORDS).find(r => r.id === d.id);
  if (!record) return { success: false, message: 'Record not found.' };
  if (record.locked === 'true') return { success: false, message: 'Record is locked.' };

  const clockIn  = new Date(d.clock_in  || record.clock_in);
  const clockOut = new Date(d.clock_out || record.clock_out);
  const breakMins = parseInt(d.break_minutes) || 0;
  const totalMinutes = Math.max(0, Math.round((clockOut - clockIn) / 60000) - breakMins);

  updateById(SHEETS.TIME_RECORDS, {
    id:            d.id,
    clock_in:      clockIn.toISOString(),
    clock_out:     clockOut.toISOString(),
    break_minutes: breakMins,
    total_minutes: totalMinutes,
    status:        'complete',
    edited:        'true',
    edited_by:     session.guard ? session.guard.name : 'Admin',
    edited_at:     new Date().toISOString(),
    notes:         d.notes !== undefined ? d.notes : record.notes
  });

  return { success: true, total_minutes: totalMinutes };
}

// Admin: lock/unlock a time record
function lockTimeRecord(id, lock) {
  requireAdmin();
  updateById(SHEETS.TIME_RECORDS, { id, locked: lock ? 'true' : 'false' });
  return { success: true };
}

// Admin: lock all records for a pay period (after payroll)
function lockPeriodTimeRecords(periodId) {
  requireAdmin();
  const sheet = SH(SHEETS.TIME_RECORDS);
  const data  = sheet.getDataRange().getValues();
  const heads = data[0];
  const idIdx     = heads.indexOf('id');
  const lockedIdx = heads.indexOf('locked');
  const records   = getTimeRecordsForPeriod(periodId);
  const ids       = new Set(records.map(r => r.id));
  for (let i = 1; i < data.length; i++) {
    if (ids.has(data[i][idIdx])) sheet.getRange(i+1, lockedIdx+1).setValue('true');
  }
  return { success: true, count: ids.size };
}

// Auto clock-out check — run via time-based trigger or on page load
function autoClockOutCheck() {
  const maxHours = parseFloat(ttConfig('max_shift_hours', '12'));
  const now = new Date();
  const records = sheetToObjects(SHEETS.TIME_RECORDS).filter(r => r.status === 'open');

  records.forEach(r => {
    const clockIn = new Date(r.clock_in);
    const elapsedHours = (now - clockIn) / 3600000;

    // Check if guard has a shift with an end time
    let limitHours = maxHours;
    if (r.shift_id) {
      const shift = sheetToObjects(SHEETS.SHIFTS).find(s => s.id === r.shift_id);
      if (shift) {
        const tmpl = sheetToObjects(SHEETS.TEMPLATES).find(t => t.code === shift.template_code);
        if (tmpl && tmpl.end_time && tmpl.end_time !== 'Custom' && tmpl.end_time !== 'TBD') {
          // Calculate hours from clock_in to scheduled end
          const [h,m,ap] = tmpl.end_time.match(/(\d+):(\d+)\s*(AM|PM)/i).slice(1);
          let endH = parseInt(h) + (ap.toUpperCase() === 'PM' && h !== '12' ? 12 : 0);
          const scheduled = new Date(clockIn);
          scheduled.setHours(endH, parseInt(m), 0, 0);
          const scheduledHours = (scheduled - clockIn) / 3600000;
          limitHours = Math.max(scheduledHours + 1, maxHours); // 1hr grace past schedule
        }
      }
    }

    if (elapsedHours >= limitHours) {
      const autoOut = new Date(clockIn.getTime() + limitHours * 3600000);
      const breakMins = (limitHours * 60) >= 240 ? 30 : 0;
      const totalMinutes = Math.max(0, Math.round(limitHours * 60) - breakMins);
      updateById(SHEETS.TIME_RECORDS, {
        id:               r.id,
        clock_out:        autoOut.toISOString(),
        break_minutes:    breakMins,
        total_minutes:    totalMinutes,
        status:           'complete',
        auto_clocked_out: 'true',
        edited:           'false',
        notes:            (r.notes || '') + ' [Auto clocked out — please update your clock-out time]'
      });
      // Notify guard
      notify(r.guard_id, 'auto_clocked_out',
        `You were automatically clocked out after ${limitHours} hours. Please log in and correct your clock-out time if needed.`);
    }
  });
}

// Get current clock status for a guard
function getClockStatus(guardId) {
  const id = guardId;
  if (!id) return { clocked_in: false };
  const open = getOpenTimeRecord(String(id));
  return open
    ? { clocked_in: true, record: open, clock_in: open.clock_in }
    : { clocked_in: false };
}

function getAllClockedIn() {
  const today = toYMD(new Date());
  return sheetToObjects(SHEETS.TIME_RECORDS).filter(r =>
    toYMD(r.date) === today && r.status === 'open'
  );
}

// Get time records summary for a guard in a period (for reporting)
function getTimeReportForGuard(guardId, periodId) {
  const records = getTimeRecordsForGuard(guardId, periodId);
  const totalMins = records.reduce((sum, r) => sum + (parseInt(r.total_minutes)||0), 0);
  return { records, total_minutes: totalMins, total_hours: (totalMins/60).toFixed(2) };
}

// Round minutes to nearest 15 (for reporting)
function roundToQuarter(minutes) {
  return Math.round(minutes / 15) * 15;
}

// Export time records for a period as CSV data
function exportTimeRecordsCSV(periodId) {
  requireAdmin();
  const period = periodById(periodId);
  if (!period) return { success: false, message: 'Period not found.' };

  const records   = getTimeRecordsForPeriod(periodId);
  const guards    = getAllGuards();
  const today     = toYMD(new Date());

  const allShifts = getShiftsForPeriod(periodId);
  const templates = sheetToObjects(SHEETS.TEMPLATES);
  const tmplMap   = {};
  templates.forEach(t => { tmplMap[t.code] = t; });

  // Helper: get projected minutes for a shift
  function projectedMins(shift) {
    const tmpl = shift ? tmplMap[shift.template_code] : null;
    return tmpl ? Math.round((parseFloat(tmpl.paid_hours) || 0) * 60) : 0;
  }
  function projectedBreak(shift) {
    const tmpl = shift ? tmplMap[shift.template_code] : null;
    return tmpl ? parseInt(tmpl.break_minutes) || 0 : 0;
  }

  const rows = [['Guard','Rank','Date','Clock In','Clock Out','Break (min)',
                 'Raw Hours','Rounded Hours','Edited','Locked','Notes']];

  // Track which guard+date combos already have a time record
  const recordSet = new Set(records.map(r => `${r.guard_id}|${toYMD(r.date)}`));

  records.forEach(r => {
    const g        = guards.find(x => String(x.id) === String(r.guard_id));
    const cin      = r.clock_in  ? new Date(r.clock_in)  : null;
    const cout     = r.clock_out ? new Date(r.clock_out) : null;
    const recDate  = toYMD(r.date);
    const isOpen   = cin && !cout;

    let rawMins, roundedMins, clockOutStr, breakMins, notesStr;

    if (isOpen) {
      // Clocked in but not yet out — use scheduled shift hours as projection
      const shift = allShifts.find(s =>
        (r.shift_id && s.id === r.shift_id) ||
        (String(s.assigned_guard_id) === String(r.guard_id) && toYMD(s.date) === recDate)
      );
      rawMins     = projectedMins(shift);
      roundedMins = roundToQuarter(rawMins);
      clockOutStr = 'Open';
      breakMins   = projectedBreak(shift) || parseInt(r.break_minutes) || 0;
      notesStr    = 'projected hours';
    } else {
      rawMins     = parseInt(r.total_minutes) || 0;
      roundedMins = roundToQuarter(rawMins);
      clockOutStr = cout ? Utilities.formatDate(cout, Session.getScriptTimeZone(), 'h:mm a') : 'Open';
      breakMins   = parseInt(r.break_minutes) || 0;
      notesStr    = r.notes || '';
    }

    rows.push([
      g ? g.name : r.guard_id,
      g ? g.rank : '',
      recDate,
      cin ? Utilities.formatDate(cin, Session.getScriptTimeZone(), 'h:mm a') : '',
      clockOutStr,
      breakMins,
      (rawMins / 60).toFixed(2),
      (roundedMins / 60).toFixed(2),
      r.edited === 'true' ? 'Yes' : '',
      r.locked === 'true' ? 'Yes' : '',
      notesStr
    ]);
  });

  // Add projected rows for future assigned shifts with no time record yet
  allShifts
    .filter(s => {
      const d = toYMD(s.date);
      return s.assigned_guard_id &&
             s.status === 'filled' &&
             d > today &&
             !recordSet.has(`${s.assigned_guard_id}|${d}`);
    })
    .sort((a, b) => toYMD(a.date).localeCompare(toYMD(b.date)))
    .forEach(s => {
      const g         = guards.find(x => String(x.id) === String(s.assigned_guard_id));
      const rawMins   = projectedMins(s);
      const breakMins = projectedBreak(s);
      rows.push([
        g ? g.name : s.assigned_guard_id,
        g ? g.rank : '',
        toYMD(s.date),
        '',
        '',
        breakMins,
        (rawMins / 60).toFixed(2),
        (roundToQuarter(rawMins) / 60).toFixed(2),
        '',
        '',
        'projected hours'
      ]);
    });

  return { success: true, csv: rows.map(r => r.map(c => `"${String(c).replace(/"/g,'""')}"`).join(',')).join('\n'),
           period: `${period.start_date}_${period.end_date}` };
}

// Client-facing functions
function clientClockIn(d)                         { return clockIn(d); }
function clientClockOut(d)                        { return clockOut(d); }
function clientGetClockStatus(token)              { const s=getSessionFromToken(token); return s ? getClockStatus(s.guard.id) : {clocked_in:false}; }
function clientGetClockStatusFor(token,guardId)   {
  // Admin-only: get clock status for any guard
  requireAdmin();
  return getClockStatus(guardId);
}
function clientGetAllClockedIn(token)             { requireAdmin(token); return getAllClockedIn(); }
function clientGetClockedOutToday(token) {
  requireAdmin(token);
  const today = toYMD(new Date());
  return sheetToObjects(SHEETS.TIME_RECORDS)
    .filter(r => toYMD(r.date) === today && r.status === 'complete')
    .sort((a, b) => new Date(b.clock_out) - new Date(a.clock_out));
}
function clientGetClockedOutYesterday(token) {
  requireAdmin(token);
  const d = new Date(); d.setDate(d.getDate() - 1);
  const yesterday = toYMD(d);
  return sheetToObjects(SHEETS.TIME_RECORDS)
    .filter(r => toYMD(r.date) === yesterday && r.status === 'complete')
    .sort((a, b) => new Date(b.clock_out) - new Date(a.clock_out));
}
function clientGetTimeRecords(token,periodId)      { requireAdmin(token); return getTimeRecordsForPeriod(periodId); }
function clientGetMyTimeRecords(token, periodId, targetGuardId) {
  const s = getSessionFromToken(token);
  if (!s) return [];
  // Admin can pass a targetGuardId to view another guard's records (impersonation)
  const guardId = (targetGuardId && (s.role === 'admin')) ? targetGuardId : s.guard.id;
  return getTimeRecordsForGuard(guardId, periodId);
}
function clientGetMyTimeRecordsByRange(token, startDate, endDate, targetGuardId) {
  const s = getSessionFromToken(token);
  if (!s) return [];
  const guardId = (targetGuardId && s.role === 'admin') ? targetGuardId : s.guard.id;
  return sheetToObjects(SHEETS.TIME_RECORDS).filter(r => {
    if (String(r.guard_id) !== String(guardId)) return false;
    const d = toYMD(r.date);
    return d >= startDate && d <= endDate && r.status !== 'cancelled';
  });
}

function clientGuardEditTimeRecord(token, d) {
  // Guards can edit their own completed, unlocked records
  const session = getSessionFromToken(token);
  if (!session) return { success: false, message: 'Not authenticated.' };
  const guardId = String(session.guard.id);
  const record = sheetToObjects(SHEETS.TIME_RECORDS).find(r => r.id === d.id);
  if (!record) return { success: false, message: 'Record not found.' };
  if (String(record.guard_id) !== guardId) return { success: false, message: 'Not your record.' };
  if (record.locked === 'true') return { success: false, message: 'Record is locked by admin.' };
  if (record.status === 'open') return { success: false, message: 'Cannot edit an open record — clock out first.' };
  const clockIn  = new Date(d.clock_in);
  const clockOut = new Date(d.clock_out);
  const breakMins = parseInt(d.break_minutes) || 0;
  const totalMinutes = Math.max(0, Math.round((clockOut - clockIn) / 60000) - breakMins);
  updateById(SHEETS.TIME_RECORDS, {
    id:            d.id,
    clock_in:      clockIn.toISOString(),
    clock_out:     clockOut.toISOString(),
    break_minutes: breakMins,
    total_minutes: totalMinutes,
    edited:        'true',
    edited_by:     session.guard.name,
    edited_at:     new Date().toISOString(),
    notes:         d.notes || record.notes || ''
  });
  return { success: true, total_minutes: totalMinutes };
}
function clientGuardDeleteTimeRecord(token, recordId) {
  const session = getSessionFromToken(token);
  if (!session) return { success: false, message: 'Not authenticated.' };
  const guardId = String(session.guard.id);
  const sheet = SH(SHEETS.TIME_RECORDS);
  if (!sheet) return { success: false, message: 'Sheet not found.' };
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const idCol = headers.indexOf('id');
  const guardCol = headers.indexOf('guard_id');
  const lockedCol = headers.indexOf('locked');
  const statusCol = headers.indexOf('status');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(recordId)) {
      if (String(data[i][guardCol]) !== guardId) return { success: false, message: 'Not your record.' };
      if (String(data[i][lockedCol]) === 'true') return { success: false, message: 'Record is locked by admin.' };
      if (String(data[i][statusCol]) === 'open') return { success: false, message: 'Cannot delete an open record — clock out first.' };
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, message: 'Record not found.' };
}
function clientEditTimeRecord(token,d)             { requireAdmin(token); return editTimeRecord(d); }
function clientAddManualTimeRecord(token,d) {
  // Admin creates a time record from scratch for any guard on any date
  const session = requireAdmin(token);
  const guard = findGuardById(d.guard_id);
  if (!guard) return { success: false, message: 'Guard not found.' };

  // Check for existing open record on same date for this guard
  const existing = sheetToObjects(SHEETS.TIME_RECORDS).find(r =>
    String(r.guard_id) === String(d.guard_id) &&
    toYMD(r.date) === toYMD(d.date) &&
    r.status !== 'cancelled'
  );
  if (existing) return { success: false, message: 'A time record already exists for this guard on this date.' };

  // Find matching shift if any
  const shift = sheetToObjects(SHEETS.SHIFTS).find(s =>
    toYMD(s.date) === toYMD(d.date) &&
    String(s.assigned_guard_id) === String(d.guard_id) &&
    s.status !== 'cancelled'
  );

  const clockIn  = new Date(d.clock_in);
  const clockOut = new Date(d.clock_out);
  const breakMins = parseInt(d.break_minutes) || 0;
  const totalMinutes = Math.max(0, Math.round((clockOut - clockIn) / 60000) - breakMins);

  const id = uid('TR');
  appendRow(SHEETS.TIME_RECORDS, {
    id,
    guard_id:         String(d.guard_id),
    guard_name:       guard.name || '',
    shift_id:         shift ? shift.id : '',
    date:             toYMD(d.date),
    clock_in:         clockIn.toISOString(),
    clock_out:        clockOut.toISOString(),
    break_minutes:    breakMins,
    total_minutes:    totalMinutes,
    status:           'complete',
    edited:           'true',
    edited_by:        (session && session.guardId ? (findGuardById(session.guardId)||{}).name : null) || 'Admin',
    edited_at:        new Date().toISOString(),
    locked:           'false',
    clock_in_lat:     '',
    clock_in_lng:     '',
    clock_in_dist_ft: '',
    clock_out_lat:    '',
    clock_out_lng:    '',
    clock_out_dist_ft:'',
    auto_clocked_out: 'false',
    notes:            d.notes || 'Manually entered by admin',
    created_at:       new Date().toISOString()
  });

  return { success: true, id, total_minutes: totalMinutes };
}
function clientLockTimeRecord(token,id,lock)       { requireAdmin(token); return lockTimeRecord(id,lock); }
function clientLockPeriodTimeRecords(token,periodId) { requireAdmin(token); return lockPeriodTimeRecords(periodId); }
function clientExportTimeRecordsCSV(token,periodId) { requireAdmin(token); return exportTimeRecordsCSV(periodId); }
function clientExportPayrollUpdateCSV(token,periodId) { requireAdmin(token); return exportPayrollUpdateCSV(periodId); }
function clientAutoClockOutCheck()                { return autoClockOutCheck(); }

// Export payroll update report: changes on/after payroll_due date.
// Covers (1) tail days where actual ≠ scheduled, (2) no-shows on tail days,
// (3) any record edited on/after payroll_due.
function exportPayrollUpdateCSV(periodId) {
  const period = periodById(periodId);
  if (!period) return { success: false, message: 'Period not found.' };

  const payrollDue = toYMD(period.payroll_due);
  if (!payrollDue) return { success: false, message: 'No payroll due date set for this period.' };

  // One read per sheet — avoid helper functions that each re-read sheets
  const periodStart = toYMD(period.start_date);
  const periodEnd   = toYMD(period.schedule_thru) || toYMD(period.end_date);

  const guards    = sheetToObjects(SHEETS.GUARDS).filter(g => g.status !== 'inactive');
  const templates = sheetToObjects(SHEETS.TEMPLATES);
  const allShifts = sheetToObjects(SHEETS.SHIFTS).filter(s => {
    const d = toYMD(s.date); return d >= periodStart && d <= periodEnd;
  });
  const records   = sheetToObjects(SHEETS.TIME_RECORDS).filter(r => {
    const d = toYMD(r.date); return d >= periodStart && d <= periodEnd;
  });

  // Lookup maps — O(1) access instead of repeated .find()
  const guardMap   = {};
  guards.forEach(g => { guardMap[String(g.id)] = g; });
  const tmplMap    = {};
  templates.forEach(t => { tmplMap[t.code] = t; });
  // record lookup: guardId|date → record
  const recByKey   = {};
  records.forEach(r => { recByKey[`${r.guard_id}|${toYMD(r.date)}`] = r; });
  // shift lookup by id
  const shiftById  = {};
  allShifts.forEach(s => { shiftById[s.id] = s; });

  const schedMins = shift => {
    if (!shift) return null;
    const t = tmplMap[shift.template_code];
    return t ? Math.round((parseFloat(t.paid_hours) || 0) * 60) : null;
  };

  // Fast time format from ISO string — avoids Utilities.formatDate overhead
  const fmtTime = iso => {
    if (!iso) return '';
    // iso may be "2026-05-28T13:00:00.000Z" or a locale string
    const d = new Date(iso);
    if (isNaN(d)) return String(iso);
    let h = d.getHours(), m = d.getMinutes();
    const ap = h >= 12 ? 'PM' : 'AM';
    h = h % 12 || 12;
    return h + ':' + (m < 10 ? '0' : '') + m + ' ' + ap;
  };
  const hh = mins => (mins / 60).toFixed(2);

  const rows = [['Guard','Rank','Date','Change Type','Scheduled Hrs','Actual Hrs','Difference','Clock In','Clock Out','Edited By','Notes']];
  const includedRecordIds = new Set();

  // 1. Tail-day shifts (date >= payrollDue)
  allShifts
    .filter(s => s.assigned_guard_id && s.status === 'filled' && toYMD(s.date) >= payrollDue)
    .forEach(shift => {
      const shiftDate = toYMD(shift.date);
      const record    = recByKey[`${shift.assigned_guard_id}|${shiftDate}`]
                     || (shift.id ? records.find(r => r.shift_id === shift.id) : null);
      const g         = guardMap[String(shift.assigned_guard_id)];
      const sm        = schedMins(shift);
      const smH       = sm !== null ? hh(sm) : 'N/A';

      if (!record) {
        rows.push([g ? g.name : shift.assigned_guard_id, g ? g.rank : '',
          shiftDate, 'No-show', smH, '0.00', sm !== null ? hh(-sm) : 'N/A', '', '', '', '']);
        return;
      }

      includedRecordIds.add(record.id);
      if (record.status === 'open') return; // still clocked in, hours not final

      const am = roundToQuarter(parseInt(record.total_minutes) || 0);
      if (sm !== null && am === sm && record.edited !== 'true') return; // no change vs estimate

      const changeType = record.edited === 'true' ? 'Tail day – edited' : 'Tail day – hours differ';
      rows.push([
        g ? g.name : record.guard_id, g ? g.rank : '',
        shiftDate, changeType, smH, hh(am),
        sm !== null ? hh(am - sm) : 'N/A',
        fmtTime(record.clock_in), fmtTime(record.clock_out),
        record.edited === 'true' ? (record.edited_by || 'admin') : '',
        record.notes || ''
      ]);
    });

  // 2. Earlier records edited on/after payrollDue
  records.forEach(r => {
    if (includedRecordIds.has(r.id)) return;
    const recDate = toYMD(r.date);
    if (recDate >= payrollDue) return;
    if (r.edited !== 'true') return;
    const editDate = r.edited_at ? toYMD(r.edited_at) : '';
    if (!editDate || editDate < payrollDue) return;

    const g     = guardMap[String(r.guard_id)];
    const shift = (r.shift_id && shiftById[r.shift_id])
               || allShifts.find(s => String(s.assigned_guard_id) === String(r.guard_id)
                                   && toYMD(s.date) === recDate && s.status === 'filled');
    const sm    = schedMins(shift);
    const am    = roundToQuarter(parseInt(r.total_minutes) || 0);
    rows.push([
      g ? g.name : r.guard_id, g ? g.rank : '',
      recDate, 'Edited after payroll due',
      sm !== null ? hh(sm) : 'N/A', hh(am),
      sm !== null ? hh(am - sm) : 'N/A',
      fmtTime(r.clock_in), fmtTime(r.clock_out),
      r.edited_by || 'admin', r.notes || ''
    ]);
  });

  const header = rows.shift();
  rows.sort((a, b) => (a[2] < b[2] ? -1 : a[2] > b[2] ? 1 : 0) || (a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0));
  rows.unshift(header);

  return {
    success: true,
    csv: rows.map(r => r.map(c => `"${String(c).replace(/"/g,'""')}"`).join(',')).join('\n'),
    period: `${period.start_date}_${period.end_date}`,
    payroll_due: payrollDue,
    row_count: rows.length - 1
  };
}
