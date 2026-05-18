// ============================================================
// HertzQuest LINE 學員身分重建與報名系統
// Google Apps Script — 主程式 Code.gs
// ============================================================
// 設定區（部署前請填入）
// ============================================================
const CONFIG = {
  SHEET_ID:            'YOUR_GOOGLE_SHEET_ID',        // Google Sheet 的 ID（網址列中那串）
  LINE_ACCESS_TOKEN:   'YOUR_LINE_CHANNEL_ACCESS_TOKEN', // LINE Messaging API Channel access token
  LINE_LIFF_REGISTER:  'YOUR_LIFF_ID_REGISTER',       // 報名頁面的 LIFF ID
  LINE_LIFF_EVENT:     'YOUR_LIFF_ID_EVENT',           // 活動頁面的 LIFF ID
  ADMIN_TOKEN:         'hertzquest_admin_2024',         // 教練後台驗證 token（自行更改）
  COACH_LINE_GROUP_ID: 'YOUR_LINE_GROUP_OR_USER_ID',   // 收通知的教練 LINE UID 或群組 ID
};

// Sheet 欄位定義
const COL = {
  MEMBERS: {
    ID: 1, LINE_UID: 2, NAME: 3, PHONE: 4,
    COURSE: 5, COACH: 6, STATUS: 7,
    REGISTERED_AT: 8, APPROVED_AT: 9, NOTES: 10,
  },
  EVENTS: {
    ID: 1, TITLE: 2, DATE: 3, LOCATION: 4,
    DESCRIPTION: 5, MAX_SPOTS: 6, SPOTS_LEFT: 7,
    PRICE: 8, IS_ACTIVE: 9, CREATED_AT: 10,
  },
  ENROLLMENTS: {
    ID: 1, EVENT_ID: 2, MEMBER_ID: 3, LINE_UID: 4,
    NAME: 5, PHONE: 6, STATUS: 7, ENROLLED_AT: 8,
  },
};

// ─────────────────────────────────────────────────────────────
// HTTP 路由
// ─────────────────────────────────────────────────────────────
function doGet(e) {
  const page = (e.parameter.page || 'register').toLowerCase();

  const pages = {
    register: () => buildPage('register', '學員身分驗證登記'),
    admin:    () => buildPage('admin',    '教練管理後台'),
    event:    () => buildPage('event',    '活動報名'),
  };

  const builder = pages[page] || pages.register;
  return builder();
}

function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action;

    const handlers = {
      register:       () => registerMember(body),
      approve:        () => approveMember(body),
      reject:         () => rejectMember(body),
      get_members:    () => getMembers(body),
      get_events:     () => getEvents(body),
      create_event:   () => createEvent(body),
      check_whitelist:() => checkWhitelist(body),
      enroll_event:   () => enrollEvent(body),
      get_enrollments:() => getEnrollments(body),
    };

    const handler = handlers[action];
    if (!handler) return json({ ok: false, error: '未知操作' });
    return json(handler());

  } catch (err) {
    console.error('doPost error:', err);
    return json({ ok: false, error: err.message });
  }
}

function buildPage(file, title) {
  return HtmlService.createTemplateFromFile(file)
    .evaluate()
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .addMetaTag('Content-Type', 'text/html; charset=utf-8')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function json(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─────────────────────────────────────────────────────────────
// 初始化試算表（第一次執行）
// ─────────────────────────────────────────────────────────────
function setupSheets() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);

  // Members
  let sh = ss.getSheetByName('Members') || ss.insertSheet('Members');
  if (sh.getLastRow() === 0) {
    sh.appendRow([
      'ID', 'LINE_UID', '真實姓名', '電話', '報名課程',
      '指定教練', '狀態', '登記時間', '核准時間', '備註',
    ]);
    sh.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#0D47A1').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
  }

  // Events
  sh = ss.getSheetByName('Events') || ss.insertSheet('Events');
  if (sh.getLastRow() === 0) {
    sh.appendRow([
      'ID', '標題', '日期', '地點', '說明',
      '名額上限', '剩餘名額', '費用(NT$)', '開放報名', '建立時間',
    ]);
    sh.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#1B5E20').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
  }

  // Enrollments
  sh = ss.getSheetByName('Enrollments') || ss.insertSheet('Enrollments');
  if (sh.getLastRow() === 0) {
    sh.appendRow([
      'ID', '活動ID', '學員ID', 'LINE_UID',
      '姓名', '電話', '狀態', '報名時間',
    ]);
    sh.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#4A148C').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
  }

  return '試算表初始化完成！';
}

// ─────────────────────────────────────────────────────────────
// 學員註冊
// ─────────────────────────────────────────────────────────────
function registerMember({ lineUid, name, phone, course, coach }) {
  if (!lineUid || !name || !phone) {
    return { ok: false, error: '請填寫所有必填欄位' };
  }

  const sheet = getSheet('Members');
  const rows  = sheet.getDataRange().getValues();

  // 防止重複登記同一 LINE UID
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][COL.MEMBERS.LINE_UID - 1] === lineUid) {
      const status = rows[i][COL.MEMBERS.STATUS - 1];
      return {
        ok: false,
        error: status === 'approved'
          ? '您的身分已通過驗證！請直接使用報名功能。'
          : '您已提交過申請，請等待教練審核。',
        status,
      };
    }
  }

  const id  = 'M' + Date.now();
  const now = new Date();
  sheet.appendRow([
    id, lineUid, name, phone,
    course || '未填寫', coach || '未指定',
    'pending', now, '', '',
  ]);

  // 通知教練有新申請
  notifyCoach(`新學員申請驗證\n姓名：${name}\n電話：${phone}\n課程：${course || '未填'}\n教練：${coach || '未指定'}\n請至教練後台審核。`);

  return { ok: true, message: '申請成功！教練審核後將通知您。', id };
}

// ─────────────────────────────────────────────────────────────
// 核准學員
// ─────────────────────────────────────────────────────────────
function approveMember({ memberId, adminToken, notes }) {
  if (adminToken !== CONFIG.ADMIN_TOKEN) return { ok: false, error: '驗證失敗' };

  const sheet = getSheet('Members');
  const rows  = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][COL.MEMBERS.ID - 1] === memberId) {
      const lineUid = rows[i][COL.MEMBERS.LINE_UID - 1];
      const name    = rows[i][COL.MEMBERS.NAME - 1];

      sheet.getRange(i + 1, COL.MEMBERS.STATUS).setValue('approved');
      sheet.getRange(i + 1, COL.MEMBERS.APPROVED_AT).setValue(new Date());
      if (notes) sheet.getRange(i + 1, COL.MEMBERS.NOTES).setValue(notes);

      // LINE 通知學員
      sendLineMessage(lineUid,
        `✅ 身分驗證成功！\n\n您好，${name}！\n\n您的 HertzQuest 學員身分已通過驗證。\n現在可以開啟活動連結進行報名囉！`
      );

      return { ok: true, message: `已核准 ${name}` };
    }
  }
  return { ok: false, error: '找不到該學員' };
}

// ─────────────────────────────────────────────────────────────
// 拒絕學員
// ─────────────────────────────────────────────────────────────
function rejectMember({ memberId, adminToken, reason }) {
  if (adminToken !== CONFIG.ADMIN_TOKEN) return { ok: false, error: '驗證失敗' };

  const sheet = getSheet('Members');
  const rows  = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][COL.MEMBERS.ID - 1] === memberId) {
      const lineUid = rows[i][COL.MEMBERS.LINE_UID - 1];
      const name    = rows[i][COL.MEMBERS.NAME - 1];

      sheet.getRange(i + 1, COL.MEMBERS.STATUS).setValue('rejected');
      if (reason) sheet.getRange(i + 1, COL.MEMBERS.NOTES).setValue(reason);

      sendLineMessage(lineUid,
        `⚠️ 身分驗證未通過\n\n您好，${name}。\n\n原因：${reason || '資料不符'}\n\n如有疑問請聯絡教練。`
      );

      return { ok: true, message: `已拒絕 ${name}` };
    }
  }
  return { ok: false, error: '找不到該學員' };
}

// ─────────────────────────────────────────────────────────────
// 取得學員名單
// ─────────────────────────────────────────────────────────────
function getMembers({ adminToken, statusFilter }) {
  if (adminToken !== CONFIG.ADMIN_TOKEN) return { ok: false, error: '驗證失敗' };

  const rows = getSheet('Members').getDataRange().getValues();
  const members = [];

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const status = r[COL.MEMBERS.STATUS - 1];
    if (statusFilter && statusFilter !== 'all' && status !== statusFilter) continue;

    members.push({
      id:           r[COL.MEMBERS.ID - 1],
      lineUid:      r[COL.MEMBERS.LINE_UID - 1],
      name:         r[COL.MEMBERS.NAME - 1],
      phone:        r[COL.MEMBERS.PHONE - 1],
      course:       r[COL.MEMBERS.COURSE - 1],
      coach:        r[COL.MEMBERS.COACH - 1],
      status:       status,
      registeredAt: r[COL.MEMBERS.REGISTERED_AT - 1]?.toString() || '',
      approvedAt:   r[COL.MEMBERS.APPROVED_AT - 1]?.toString() || '',
      notes:        r[COL.MEMBERS.NOTES - 1],
    });
  }

  return { ok: true, members };
}

// ─────────────────────────────────────────────────────────────
// 白名單檢查
// ─────────────────────────────────────────────────────────────
function checkWhitelist({ lineUid }) {
  if (!lineUid) return { ok: false, approved: false };

  const rows = getSheet('Members').getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][COL.MEMBERS.LINE_UID - 1] === lineUid &&
        rows[i][COL.MEMBERS.STATUS - 1] === 'approved') {
      return {
        ok: true, approved: true,
        name: rows[i][COL.MEMBERS.NAME - 1],
        id:   rows[i][COL.MEMBERS.ID - 1],
      };
    }
  }
  return { ok: true, approved: false };
}

// ─────────────────────────────────────────────────────────────
// 活動管理
// ─────────────────────────────────────────────────────────────
function createEvent({ adminToken, title, date, location, description, maxSpots, price }) {
  if (adminToken !== CONFIG.ADMIN_TOKEN) return { ok: false, error: '驗證失敗' };
  if (!title || !date) return { ok: false, error: '標題和日期為必填' };

  const id  = 'E' + Date.now();
  const now = new Date();
  const spots = Number(maxSpots) || 20;

  getSheet('Events').appendRow([
    id, title, date, location || '',
    description || '', spots, spots,
    Number(price) || 0, true, now,
  ]);

  return { ok: true, eventId: id, message: `活動「${title}」已建立` };
}

function getEvents({ activeOnly }) {
  const rows   = getSheet('Events').getDataRange().getValues();
  const events = [];

  for (let i = 1; i < rows.length; i++) {
    const r        = rows[i];
    const isActive = r[COL.EVENTS.IS_ACTIVE - 1];
    if (activeOnly && !isActive) continue;

    events.push({
      id:          r[COL.EVENTS.ID - 1],
      title:       r[COL.EVENTS.TITLE - 1],
      date:        r[COL.EVENTS.DATE - 1]?.toString() || '',
      location:    r[COL.EVENTS.LOCATION - 1],
      description: r[COL.EVENTS.DESCRIPTION - 1],
      maxSpots:    r[COL.EVENTS.MAX_SPOTS - 1],
      spotsLeft:   r[COL.EVENTS.SPOTS_LEFT - 1],
      price:       r[COL.EVENTS.PRICE - 1],
      isActive:    isActive,
    });
  }

  return { ok: true, events };
}

// ─────────────────────────────────────────────────────────────
// 活動報名（需通過白名單）
// ─────────────────────────────────────────────────────────────
function enrollEvent({ lineUid, eventId }) {
  // 白名單驗證
  const wl = checkWhitelist({ lineUid });
  if (!wl.approved) {
    return { ok: false, error: '您尚未通過身分驗證，無法報名活動。' };
  }

  const eSheet = getSheet('Events');
  const eRows  = eSheet.getDataRange().getValues();
  let eventRow = -1, eventTitle = '', spotsLeft = 0;

  for (let i = 1; i < eRows.length; i++) {
    if (eRows[i][COL.EVENTS.ID - 1] === eventId) {
      eventRow   = i + 1;
      eventTitle = eRows[i][COL.EVENTS.TITLE - 1];
      spotsLeft  = Number(eRows[i][COL.EVENTS.SPOTS_LEFT - 1]);
      break;
    }
  }

  if (eventRow === -1) return { ok: false, error: '找不到該活動' };
  if (spotsLeft <= 0)  return { ok: false, error: '名額已滿' };

  // 防重複報名
  const nRows = getSheet('Enrollments').getDataRange().getValues();
  for (let i = 1; i < nRows.length; i++) {
    if (nRows[i][COL.ENROLLMENTS.EVENT_ID - 1] === eventId &&
        nRows[i][COL.ENROLLMENTS.LINE_UID - 1] === lineUid) {
      return { ok: false, error: '您已報名此活動' };
    }
  }

  // 寫入報名
  const enId = 'EN' + Date.now();
  getSheet('Enrollments').appendRow([
    enId, eventId, wl.id, lineUid,
    wl.name, '', 'confirmed', new Date(),
  ]);

  // 扣除名額
  eSheet.getRange(eventRow, COL.EVENTS.SPOTS_LEFT).setValue(spotsLeft - 1);

  // LINE 確認通知
  sendLineMessage(lineUid,
    `✅ 報名成功！\n\n活動：${eventTitle}\n姓名：${wl.name}\n\n教練會再與您確認細節，請留意訊息通知！`
  );

  return { ok: true, message: `報名成功！活動：${eventTitle}` };
}

function getEnrollments({ adminToken, eventId }) {
  if (adminToken !== CONFIG.ADMIN_TOKEN) return { ok: false, error: '驗證失敗' };

  const rows        = getSheet('Enrollments').getDataRange().getValues();
  const enrollments = [];

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (eventId && r[COL.ENROLLMENTS.EVENT_ID - 1] !== eventId) continue;
    enrollments.push({
      id:         r[COL.ENROLLMENTS.ID - 1],
      eventId:    r[COL.ENROLLMENTS.EVENT_ID - 1],
      memberId:   r[COL.ENROLLMENTS.MEMBER_ID - 1],
      lineUid:    r[COL.ENROLLMENTS.LINE_UID - 1],
      name:       r[COL.ENROLLMENTS.NAME - 1],
      status:     r[COL.ENROLLMENTS.STATUS - 1],
      enrolledAt: r[COL.ENROLLMENTS.ENROLLED_AT - 1]?.toString() || '',
    });
  }

  return { ok: true, enrollments };
}

// ─────────────────────────────────────────────────────────────
// LINE Messaging API
// ─────────────────────────────────────────────────────────────
function sendLineMessage(toUid, text) {
  if (!CONFIG.LINE_ACCESS_TOKEN || CONFIG.LINE_ACCESS_TOKEN.includes('YOUR_')) {
    console.warn('LINE_ACCESS_TOKEN 未設定，跳過發訊息');
    return;
  }
  try {
    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method: 'post',
      headers: {
        'Authorization': 'Bearer ' + CONFIG.LINE_ACCESS_TOKEN,
        'Content-Type':  'application/json',
      },
      payload: JSON.stringify({
        to: toUid,
        messages: [{ type: 'text', text }],
      }),
      muteHttpExceptions: true,
    });
  } catch (err) {
    console.error('LINE sendMessage error:', err);
  }
}

function notifyCoach(text) {
  if (CONFIG.COACH_LINE_GROUP_ID && !CONFIG.COACH_LINE_GROUP_ID.includes('YOUR_')) {
    sendLineMessage(CONFIG.COACH_LINE_GROUP_ID, text);
  }
}

// ─────────────────────────────────────────────────────────────
// 工具函式
// ─────────────────────────────────────────────────────────────
function getSheet(name) {
  return SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(name);
}
