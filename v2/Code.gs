// ================================================================
// HertzQuest 團練系統 — Backend (Google Apps Script)
// ================================================================
const CFG = {
  SHEET_ID:   '1Vc0d6T0q-7NQav4ZJpfrJZ7e1BA-vbERMMmsFsyOMyE',
  LINE_TOKEN: 'NlJBb+6+N7ia+qn9TdbgMMy/Xs9GyPgZZoPfowe7xUAYhm6wxgupfK3shAZ+je8zzIt6eEBAvmGR+I+mqXvw/yLtl+m25Kj9Erh9q2q6JLhZvoZl+jK0jiP6unr1kIh4VG1UrhVj2I7k4autSZ52sgdB04t89/1O/w1cDnyilFU=',
  COACH_PWD:  'hertz2024',
  LIFF_ID:    '2010027559-58ppHoAr',
};

// Column indices (1-based) — Enrollments 含 CERT 欄（第6欄）
const C = {
  M: { ID:1,UID:2,NAME:3,PHONE:4,DISPLAY:5,PIC:6,STATUS:7,AT:8,
       FIRST_CLASS:9,COACH_NAME:10,LINE_ID:11,DEPTH:12,BREATH:13,CERT:14,APPROVAL:15 },
  S: { ID:1,TITLE:2,DATE:3,TIME:4,LOC:5,MEET:6,DESC:7,COACH:8,MAX:9,LEFT:10,PRICE:11,ACTIVE:12,AT:13,UNCERT_MAX:14 },
  E: { ID:1,SID:2,UID:3,NAME:4,PHONE:5,CERT:6,AT:7,TRANSFER:8,PAYMENT:9 },
};

// ── 路由 ──────────────────────────────────────────────────────
function doGet(e) {
  const pg = (e.parameter.page || 'index');
  const file = pg === 'coach' ? 'coach' : 'index';
  return HtmlService.createHtmlOutputFromFile(file)
    .setTitle(pg === 'coach' ? 'HertzQuest 教練後台' : 'HertzQuest 團練')
    .addMetaTag('viewport','width=device-width,initial-scale=1,maximum-scale=1')
    .addMetaTag('Content-Type', 'text/html; charset=utf-8')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    return res(dispatch(d));
  } catch(err) {
    return res({ok:false,error:err.message});
  }
}

function dispatch(d) {
  const map = {
    initMember, getSessions, enrollSession, myEnrollments, cancelEnroll,
    coachLogin, allSessions, createSession, updateSession, deleteSession,
    getEnrollList, notifyStudents, confirmPayment, getMembers, deleteMember,
    approveStudent, updateStudentStats,
    getAnnouncements, createAnnouncement, deleteAnnouncement,
    createTrip, getTrips, enrollTrip, cancelTripEnroll, myTripEnrollments,
    getTripEnrollList, updateTrip, deleteTrip,
  };
  const fn = map[d.action];
  if (!fn) return {ok:false,error:'Unknown action: '+d.action};
  return fn(d);
}

function res(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── 學員 API ─────────────────────────────────────────────────
function initMember({lineUid, displayName, pictureUrl, realName, phone, firstClassDate, coachName, lineId}) {
  const sh = sh_('Members');
  const rows = sh.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][C.M.UID-1] === lineUid) {
      // 更新 LINE 暱稱/頭像
      sh.getRange(i+1, C.M.DISPLAY).setValue(displayName||'');
      sh.getRange(i+1, C.M.PIC).setValue(pictureUrl||'');
      const r = rows[i];
      return {ok:true, member:{
        id:           r[C.M.ID-1],
        lineUid:      r[C.M.UID-1],
        realName:     r[C.M.NAME-1],
        phone:        r[C.M.PHONE-1],
        displayName:  displayName||r[C.M.DISPLAY-1],
        firstClassDate: String(r[C.M.FIRST_CLASS-1]||''),
        coachName:    r[C.M.COACH_NAME-1]||'',
        lineId:       r[C.M.LINE_ID-1]||'',
        depth:        String(r[C.M.DEPTH-1]||''),
        breathHold:   formatBreath(r[C.M.BREATH-1]),
        certLevel:    r[C.M.CERT-1]||'',
        approvalStatus: r[C.M.APPROVAL-1]||'pending',
      }, isNew:false};
    }
  }

  // 新學員 — 如果沒有提交真實資料，先回傳 isNew:true 讓前端顯示表單
  if (!realName || !phone) return {ok:true, member:null, isNew:true};

  const id = 'M'+Date.now();
  sh.appendRow([id,lineUid,realName,phone,displayName||'',pictureUrl||'','active',new Date(),firstClassDate||'',coachName||'',lineId||'','','','','pending']);
  return {ok:true, member:{id,lineUid,realName,phone,displayName,firstClassDate:firstClassDate||'',coachName:coachName||'',lineId:lineId||'',depth:'',breathHold:'',certLevel:'',approvalStatus:'pending'}, isNew:true, justRegistered:true};
}

function getSessions({lineUid}) {
  const sRows = sh_('Sessions').getDataRange().getValues();
  const eRows = sh_('Enrollments').getDataRange().getValues();
  const today = new Date(); today.setHours(0,0,0,0);

  const enrolled = new Set();
  const uncertCnt = {};
  eRows.slice(1).forEach(r => {
    if (r[C.E.UID-1]===lineUid) enrolled.add(r[C.E.SID-1]);
    const sid = r[C.E.SID-1];
    if (r[C.E.CERT-1]==='無證') uncertCnt[sid] = (uncertCnt[sid]||0)+1;
  });

  const sessions = sRows.slice(1)
    .filter(r => String(r[C.S.ACTIVE-1]).toLowerCase()==='true')
    .filter(r => { try { return new Date(r[C.S.DATE-1]) >= today; } catch(e){ return false; } })
    .sort((a,b) => new Date(a[C.S.DATE-1])-new Date(b[C.S.DATE-1]))
    .map(r => ({...toSession(r), enrolled: enrolled.has(r[C.S.ID-1]), uncertCount: uncertCnt[r[C.S.ID-1]]||0}));

  return {ok:true, sessions};
}

function enrollSession({lineUid, sessionId, certLevel, transferCode}) {
  // 確認學員已登記
  const mRows = sh_('Members').getDataRange().getValues();
  const member = mRows.slice(1).find(r => r[C.M.UID-1]===lineUid);
  if (!member) return {ok:false, error:'請先填寫真實姓名和電話完成學員登記'};

  // 確認未重複報名
  const eSheet = sh_('Enrollments');
  const eRows = eSheet.getDataRange().getValues();
  if (eRows.slice(1).find(r => r[C.E.SID-1]===sessionId && r[C.E.UID-1]===lineUid)) {
    return {ok:false, error:'您已報名此活動'};
  }

  // 確認名額
  const sSheet = sh_('Sessions');
  const sRows = sSheet.getDataRange().getValues();
  let sRowIdx = -1, sData = null;
  for (let i=1;i<sRows.length;i++) {
    if (sRows[i][C.S.ID-1]===sessionId) { sRowIdx=i+1; sData=sRows[i]; break; }
  }
  if (!sData) return {ok:false,error:'找不到此活動'};
  // 分別計算無照/有照已報人數
  const allForSession = eRows.slice(1).filter(r => r[C.E.SID-1]===sessionId);
  const totalEnrolled = allForSession.length;
  const uncertEnrolled = allForSession.filter(r => r[C.E.CERT-1]==='無證').length;
  const uncertMax = Number(sData[C.S.UNCERT_MAX-1])||4;
  const certMax = Number(sData[C.S.MAX-1]) - uncertMax;
  if (certLevel==='無證') {
    if (uncertEnrolled >= uncertMax) return {ok:false, error:`無照名額已滿（${uncertMax}/${uncertMax}）`};
  } else {
    const certEnrolled = totalEnrolled - uncertEnrolled;
    if (certMax > 0 && certEnrolled >= certMax) return {ok:false, error:'有照名額已滿'};
  }
  if (Number(sData[C.S.LEFT-1]) <= 0) return {ok:false,error:'名額已滿'};

  // 寫入報名
  const id='E'+Date.now();
  eSheet.appendRow([id,sessionId,lineUid,member[C.M.NAME-1],member[C.M.PHONE-1],certLevel||'',new Date(),transferCode||'','pending']);
  sSheet.getRange(sRowIdx, C.S.LEFT).setValue(Number(sData[C.S.LEFT-1])-1);

  // LINE 確認通知
  const dateStr = String(sData[C.S.DATE-1]).slice(0,10);
  sendLine(lineUid,
    `✅ 報名成功！\n\n` +
    `📌 ${sData[C.S.TITLE-1]}\n` +
    `📅 ${dateStr} ${sData[C.S.TIME-1]||''}\n` +
    `📍 地點：${sData[C.S.LOC-1]||''}\n` +
    `🗺 集合點：${sData[C.S.MEET-1]||''}\n\n` +
    `教練 ${sData[C.S.COACH-1]} 期待您的到來！`
  );

  return {ok:true, message:'報名成功！收到 LINE 確認通知了嗎？'};
}

function myEnrollments({lineUid}) {
  const eRows = sh_('Enrollments').getDataRange().getValues();
  const sRows = sh_('Sessions').getDataRange().getValues();
  const sMap = {};
  sRows.slice(1).forEach(r => { sMap[r[C.S.ID-1]]=r; });

  const list = eRows.slice(1)
    .filter(r => r[C.E.UID-1]===lineUid)
    .map(r => {
      const s = sMap[r[C.E.SID-1]];
      if (!s) return null;
      return {
        ...toSession(s),
        enrollId:      String(r[C.E.ID-1]||''),
        transferCode:  r[C.E.TRANSFER-1]||'',
        paymentStatus: r[C.E.PAYMENT-1]||'pending',
      };
    })
    .filter(Boolean)
    .sort((a,b) => new Date(a.date)-new Date(b.date));

  return {ok:true, enrollments:list};
}

function cancelEnroll({lineUid, sessionId}) {
  const eSheet = sh_('Enrollments');
  const eRows = eSheet.getDataRange().getValues();
  for (let i=1;i<eRows.length;i++) {
    if (eRows[i][C.E.SID-1]===sessionId && eRows[i][C.E.UID-1]===lineUid) {
      eSheet.deleteRow(i+1);
      // 還名額
      const sSheet = sh_('Sessions');
      const sRows = sSheet.getDataRange().getValues();
      for (let j=1;j<sRows.length;j++) {
        if (sRows[j][C.S.ID-1]===sessionId) {
          sSheet.getRange(j+1, C.S.LEFT).setValue(Number(sRows[j][C.S.LEFT-1])+1);
          break;
        }
      }
      return {ok:true, message:'已取消報名'};
    }
  }
  return {ok:false, error:'找不到報名紀錄'};
}

// ── 教練 API ─────────────────────────────────────────────────
function coachLogin({password}) {
  return password===CFG.COACH_PWD ? {ok:true} : {ok:false,error:'密碼錯誤'};
}

function allSessions({password}) {
  if (!auth(password)) return noAuth();
  const sRows = sh_('Sessions').getDataRange().getValues();
  const eRows = sh_('Enrollments').getDataRange().getValues();
  const cnt = {};
  const uncertCnt = {};
  eRows.slice(1).forEach(r => {
    if(r[C.E.UID-1]) {
      cnt[r[C.E.SID-1]]=(cnt[r[C.E.SID-1]]||0)+1;
      if(r[C.E.CERT-1]==='無證') uncertCnt[r[C.E.SID-1]]=(uncertCnt[r[C.E.SID-1]]||0)+1;
    }
  });

  const sessions = sRows.slice(1)
    .sort((a,b) => new Date(b[C.S.DATE-1])-new Date(a[C.S.DATE-1]))
    .map(r => ({...toSession(r), enrollCount:cnt[r[C.S.ID-1]]||0, uncertCount:uncertCnt[r[C.S.ID-1]]||0, active: String(r[C.S.ACTIVE-1]).toLowerCase()==='true'}));

  return {ok:true, sessions};
}

function createSession({password, title, date, time, location, meetingPoint, description, coach, uncertMax, certMaxSpots, price}) {
  if (!auth(password)) return noAuth();
  if (!title||!date) return {ok:false,error:'標題和日期為必填'};
  const id='S'+Date.now();
  const uMax = Number(uncertMax)||4;
  const cMax = Number(certMaxSpots)||0;
  const max = uMax + cMax;
  sh_('Sessions').appendRow([id,title,date,time||'',location||'',meetingPoint||'',description||'',coach||'',max,max,Number(price)||0,true,new Date(),uMax]);
  return {ok:true,id,message:`「${title}」已建立`};
}

function updateSession({password, sessionId, ...fields}) {
  if (!auth(password)) return noAuth();
  const sh = sh_('Sessions');
  const rows = sh.getDataRange().getValues();
  for (let i=1;i<rows.length;i++) {
    if (rows[i][C.S.ID-1]===sessionId) {
      const map = {title:C.S.TITLE,date:C.S.DATE,time:C.S.TIME,location:C.S.LOC,meetingPoint:C.S.MEET,description:C.S.DESC,coach:C.S.COACH,maxSpots:C.S.MAX,price:C.S.PRICE,active:C.S.ACTIVE,uncertMax:C.S.UNCERT_MAX};
      Object.entries(fields).forEach(([k,v]) => { if(map[k]) sh.getRange(i+1,map[k]).setValue(v); });
      return {ok:true};
    }
  }
  return {ok:false,error:'找不到活動'};
}

function deleteSession({password, sessionId}) {
  if (!auth(password)) return noAuth();
  const sh = sh_('Sessions');
  const rows = sh.getDataRange().getValues();
  for (let i=1;i<rows.length;i++) {
    if (rows[i][C.S.ID-1]===sessionId) { sh.deleteRow(i+1); return {ok:true}; }
  }
  return {ok:false,error:'找不到活動'};
}

function getMembers({password}) {
  if (!auth(password)) return noAuth();
  const rows = sh_('Members').getDataRange().getValues();
  const members = rows.slice(1).filter(r=>r[0]).map(r=>({
    id:           r[C.M.ID-1],
    lineUid:      r[C.M.UID-1],
    realName:     r[C.M.NAME-1],
    phone:        r[C.M.PHONE-1],
    displayName:  r[C.M.DISPLAY-1],
    status:       r[C.M.STATUS-1],
    joinedAt:     String(r[C.M.AT-1]).slice(0,10),
    firstClassDate: String(r[C.M.FIRST_CLASS-1]||''),
    coachName:    r[C.M.COACH_NAME-1]||'',
    lineId:       r[C.M.LINE_ID-1]||'',
    depth:        String(r[C.M.DEPTH-1]||''),
    breathHold:   formatBreath(r[C.M.BREATH-1]),
    certLevel:    r[C.M.CERT-1]||'',
    approvalStatus: r[C.M.APPROVAL-1]||'pending',
  }));
  return {ok:true, members};
}

function approveStudent({password, memberId, approved}) {
  if (!auth(password)) return noAuth();
  const sh = sh_('Members');
  const rows = sh.getDataRange().getValues();
  for (let i=1;i<rows.length;i++) {
    if (rows[i][C.M.ID-1]===memberId) {
      sh.getRange(i+1, C.M.APPROVAL).setValue(approved ? 'approved' : 'rejected');
      return {ok:true};
    }
  }
  return {ok:false, error:'找不到學員'};
}

function updateStudentStats({password, memberId, depth, breathHold, certLevel}) {
  if (!auth(password)) return noAuth();
  const sh = sh_('Members');
  const rows = sh.getDataRange().getValues();
  for (let i=1;i<rows.length;i++) {
    if (rows[i][C.M.ID-1]===memberId) {
      if (depth !== undefined) sh.getRange(i+1, C.M.DEPTH).setValue(depth||'');
      if (breathHold !== undefined) sh.getRange(i+1, C.M.BREATH).setValue(breathHold||'');
      if (certLevel !== undefined) sh.getRange(i+1, C.M.CERT).setValue(certLevel||'');
      return {ok:true};
    }
  }
  return {ok:false, error:'找不到學員'};
}

function deleteMember({password, memberId}) {
  if (!auth(password)) return noAuth();
  const sh = sh_('Members');
  const rows = sh.getDataRange().getValues();
  for (let i=1;i<rows.length;i++) {
    if (rows[i][C.M.ID-1]===memberId) { sh.deleteRow(i+1); return {ok:true}; }
  }
  return {ok:false,error:'找不到學員'};
}

function getEnrollList({password, sessionId}) {
  if (!auth(password)) return noAuth();
  const rows = sh_('Enrollments').getDataRange().getValues();
  // 支援新舊兩種欄位格式：有 STATUS 欄或無
  const list = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (r[C.E.SID-1]===sessionId && r[0] && String(r[0]).length>0) {
      list.push({
        id:           String(r[C.E.ID-1]||''),
        rowIndex:     i+1, // 實際的 1-based Sheet 行號
        name:         r[C.E.NAME-1],
        phone:        r[C.E.PHONE-1],
        at:           String(r[C.E.AT-1]).slice(0,10),
        lineUid:      String(r[C.E.UID-1]||''),
        certLevel:    r[C.E.CERT-1]||'',
        transferCode: r[C.E.TRANSFER-1]||'',
        paymentStatus:r[C.E.PAYMENT-1]||'pending',
      });
    }
  }
  return {ok:true, list};
}

function confirmPayment({password, enrollId, lineUid, sessionId, rowIndex}) {
  if (!auth(password)) return noAuth();
  const sh = sh_('Enrollments');
  const rows = sh.getDataRange().getValues();
  const PAYMENT_COL = C.E.PAYMENT; // 9

  // 1. 直接用行號（最準確）
  if (rowIndex && rowIndex > 1 && rowIndex <= rows.length) {
    sh.getRange(rowIndex, PAYMENT_COL).setValue('confirmed');
    return {ok:true};
  }
  // 2. 用 enrollId 查
  if (enrollId) {
    for (let i=1;i<rows.length;i++) {
      if (String(rows[i][C.E.ID-1])===String(enrollId)) {
        sh.getRange(i+1, PAYMENT_COL).setValue('confirmed');
        return {ok:true};
      }
    }
  }
  // 3. 用 lineUid + sessionId 查（備援）
  if (lineUid && sessionId) {
    for (let i=1;i<rows.length;i++) {
      if (String(rows[i][C.E.UID-1])===lineUid && String(rows[i][C.E.SID-1])===sessionId) {
        sh.getRange(i+1, PAYMENT_COL).setValue('confirmed');
        return {ok:true};
      }
    }
  }
  return {ok:false, error:'找不到報名紀錄'};
}

function notifyStudents({password, sessionId, message}) {
  if (!auth(password)) return noAuth();
  const rows = sh_('Enrollments').getDataRange().getValues();
  const uids = rows.slice(1)
    .filter(r => r[C.E.SID-1]===sessionId && r[C.E.UID-1])
    .map(r => r[C.E.UID-1]);
  uids.forEach(uid => sendLine(uid, message));
  return {ok:true, sent:uids.length};
}

// ── 工具 ─────────────────────────────────────────────────────
function auth(p) { return p===CFG.COACH_PWD; }
function noAuth() { return {ok:false,error:'未授權'}; }
function checkPassword(p) { return p===CFG.COACH_PWD ? {ok:true} : {ok:false,error:'未授權'}; }
function sh_(name) { return SpreadsheetApp.openById(CFG.SHEET_ID).getSheetByName(name); }

function toSession(r) {
  // 時間欄位可能被 Sheets 解析為 Date 物件，用 UTC 方法避免時區偏移，並轉成 AM/PM
  var rawTime = r[3];
  var timeStr = '';
  if (rawTime instanceof Date) {
    var h = rawTime.getUTCHours();
    var m = rawTime.getUTCMinutes();
    timeStr = String(h).padStart(2,'0') + ':' + String(m).padStart(2,'0');
  } else if (rawTime) {
    var parts = String(rawTime).split(':');
    if (parts.length >= 2) {
      var hh = parseInt(parts[0], 10);
      var mm = parseInt(parts[1], 10);
      if (!isNaN(hh) && !isNaN(mm)) {
        timeStr = String(hh).padStart(2,'0') + ':' + String(mm).padStart(2,'0');
      } else {
        timeStr = String(rawTime);
      }
    } else {
      timeStr = String(rawTime);
    }
  }
  // 日期欄位同樣可能是 Date 物件，用 formatDate 確保格式為 YYYY-MM-DD
  var rawDate = r[2];
  var dateStr = rawDate instanceof Date ? formatDate(rawDate) : String(rawDate).slice(0,10);
  return {
    id:r[0], title:r[1],
    date: dateStr,
    time: timeStr, location:r[4]||'', meetingPoint:r[5]||'',
    description:r[6]||'', coach:r[7]||'',
    maxSpots:r[8], spotsLeft:r[9], price:r[10],
    active: String(r[11]).toLowerCase()==='true',
    uncertMax: Number(r[C.S.UNCERT_MAX-1])||4,
  };
}

function toMember(r) {
  return {
    id:r[C.M.ID-1], lineUid:r[C.M.UID-1], realName:r[C.M.NAME-1],
    phone:r[C.M.PHONE-1], displayName:r[C.M.DISPLAY-1],
    firstClassDate: String(r[C.M.FIRST_CLASS-1]||''),
    coachName: r[C.M.COACH_NAME-1]||'',
    lineId: r[C.M.LINE_ID-1]||'',
    depth: String(r[C.M.DEPTH-1]||''),
    breathHold: formatBreath(r[C.M.BREATH-1]),
    certLevel: r[C.M.CERT-1]||'',
    approvalStatus: r[C.M.APPROVAL-1]||'pending',
  };
}

function sendLine(toUid, text) {
  if (!CFG.LINE_TOKEN||CFG.LINE_TOKEN.includes('REPLACE')) { console.log('LINE token not set'); return; }
  try {
    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push',{
      method:'post',
      headers:{'Authorization':'Bearer '+CFG.LINE_TOKEN,'Content-Type':'application/json'},
      payload:JSON.stringify({to:toUid,messages:[{type:'text',text}]}),
      muteHttpExceptions:true,
    });
  } catch(e) { console.error('LINE error:',e); }
}

// ── 公告 API ─────────────────────────────────────────────────
function getAnnouncements() {
  try {
    const ss = SpreadsheetApp.openById(CFG.SHEET_ID);
    let sheet;
    try { sheet = ss.getSheetByName('Announcements'); } catch(e) { sheet = null; }
    if (!sheet) {
      sheet = ss.insertSheet('Announcements');
      sheet.getRange(1,1,1,3).setValues([['id','text','active']]);
      return { ok:true, announcements:[] };
    }
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idCol = headers.indexOf('id');
    const textCol = headers.indexOf('text');
    const activeCol = headers.indexOf('active');
    const announcements = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][activeCol] === true || data[i][activeCol] === 'TRUE' || data[i][activeCol] === 1) {
        announcements.push({ id: data[i][idCol], text: data[i][textCol] });
      }
    }
    return { ok:true, announcements };
  } catch(e) {
    return { ok:false, error: e.message };
  }
}

function createAnnouncement({password, text, active}) {
  if (!auth(password)) return noAuth();
  if (!text) return {ok:false, error:'公告內容不可為空'};
  try {
    const ss = SpreadsheetApp.openById(CFG.SHEET_ID);
    let sheet = ss.getSheetByName('Announcements');
    if (!sheet) {
      sheet = ss.insertSheet('Announcements');
      sheet.getRange(1,1,1,3).setValues([['id','text','active']]);
    }
    const id = 'A' + Date.now();
    sheet.appendRow([id, text, active !== false ? true : false]);
    return {ok:true, id, message:'公告已建立'};
  } catch(e) {
    return {ok:false, error: e.message};
  }
}

function deleteAnnouncement({password, announcementId}) {
  if (!auth(password)) return noAuth();
  try {
    const sheet = SpreadsheetApp.openById(CFG.SHEET_ID).getSheetByName('Announcements');
    if (!sheet) return {ok:false, error:'找不到公告表'};
    const data = sheet.getDataRange().getValues();
    const idCol = data[0].indexOf('id');
    for (let i = 1; i < data.length; i++) {
      if (data[i][idCol] === announcementId) {
        sheet.deleteRow(i + 1);
        return {ok:true};
      }
    }
    return {ok:false, error:'找不到此公告'};
  } catch(e) {
    return {ok:false, error: e.message};
  }
}

// 第一次執行：初始化 Sheets
function setupSheets() {
  const ss = SpreadsheetApp.openById(CFG.SHEET_ID);
  const init = (name,hdrs,bg) => {
    let sh=ss.getSheetByName(name)||ss.insertSheet(name);
    if(sh.getLastRow()===0){
      sh.appendRow(hdrs);
      sh.getRange(1,1,1,hdrs.length).setFontWeight('bold').setBackground(bg).setFontColor('#fff');
      sh.setFrozenRows(1);
    }
  };
  init('Members',     ['ID','LINE_UID','真實姓名','電話','LINE暱稱','頭像','狀態','加入時間','初次上課日','教練','LINE_ID','深度','閉氣','證照','審核狀態'],'#1565C0');
  init('Sessions',    ['ID','標題','日期','時間','地點','集合地點','說明','教練','名額上限','剩餘名額','費用','開放報名','建立時間','無照名額'],'#2E7D32');
  init('Enrollments', ['ID','活動ID','LINE_UID','姓名','電話','證照等級','報名時間','轉帳末5碼','收款狀態'],'#6A1B9A');
  init('Announcements',['id','text','active'],'#4A148C');
  initTripSheets();
  return '✅ Sheets 初始化完成！';
}

// ── 旅行功能 ────────────────────────────────────────────────

function formatTime(val) {
  if (!val) return '';
  if (val instanceof Date) {
    var h = val.getHours().toString().padStart(2,'0');
    var m = val.getMinutes().toString().padStart(2,'0');
    return h + ':' + m;
  }
  return String(val);
}

// 閉氣時間可能被 Sheets 解析為 Date 物件（1899-12-30 時間格式）
// 或純文字 "1:24"，或數字（小數天）——統一轉成易讀字串
function formatBreath(val) {
  if (!val && val !== 0) return '';
  // Date 物件（Sheets Time 格式）
  // 必須用 getUTC* — GAS 把 Sheets 時間存成 UTC epoch，用 getHours() 會多加時區偏移（台灣 +8）
  if (val instanceof Date) {
    // Sheets 把使用者輸入的 "3:00" 解讀成 3 小時 0 分
    // 所以要把 getUTCHours() 當「分鐘」、getUTCMinutes() 當「秒數」讀回來
    var mm = val.getUTCHours();
    var ss = val.getUTCMinutes();
    return mm + ':' + String(ss).padStart(2,'0');
  }
  // 數字（Sheets 以小數天儲存）："3:00" = 3 小時 = 3/24
  // 同理，把「小時部分」當分鐘，「分鐘部分」當秒數
  if (typeof val === 'number') {
    var totalHoursFrac = val * 24;
    var mm2 = Math.floor(totalHoursFrac);
    var ss2 = Math.round((totalHoursFrac - mm2) * 60);
    return mm2 + ':' + String(ss2).padStart(2,'0');
  }
  return String(val);
}

function formatDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    var y = val.getFullYear();
    var mo = (val.getMonth()+1).toString().padStart(2,'0');
    var d = val.getDate().toString().padStart(2,'0');
    return y + '-' + mo + '-' + d;
  }
  return String(val);
}

function initTripSheets() {
  var ss = SpreadsheetApp.openById(CFG.SHEET_ID);
  if (!ss.getSheetByName('Trips')) {
    var sh = ss.insertSheet('Trips');
    sh.appendRow(['id','country','title','itinerary','inclusions','priceSingle','priceCouple','startDate','endDate','maxSpots','minSpots','active','notes','createdAt']);
  } else {
    // Add minSpots column if missing (backward compat)
    var sh = ss.getSheetByName('Trips');
    var headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
    if (headers.indexOf('minSpots') === -1) {
      sh.getRange(1, sh.getLastColumn()+1).setValue('minSpots');
    }
  }
  if (!ss.getSheetByName('TripEnrollments')) {
    var sh2 = ss.insertSheet('TripEnrollments');
    sh2.appendRow(['id','tripId','memberId','memberName','memberPhone','plan','enrolledAt']);
  }
}

function createTrip(data) {
  try {
    var p = checkPassword(data.password);
    if (!p.ok) return p;
    initTripSheets();
    var sh = SpreadsheetApp.openById(CFG.SHEET_ID).getSheetByName('Trips');
    var id = 'trip_' + Date.now();
    sh.appendRow([id, data.country||'', data.title||'', data.itinerary||'', data.inclusions||'',
      data.priceSingle||0, data.priceCouple||0, data.startDate||'', data.endDate||'',
      data.maxSpots||20, data.minSpots||0, true, data.notes||'', new Date().toISOString()]);
    return { ok: true, message: '旅行已建立', id: id };
  } catch(e) { return { ok: false, error: e.message }; }
}

function getTrips(data) {
  try {
    initTripSheets();
    var ss = SpreadsheetApp.openById(CFG.SHEET_ID);
    var sh = ss.getSheetByName('Trips');
    var rows = sh.getDataRange().getValues();
    if (rows.length <= 1) return { ok: true, trips: [] };
    var headers = rows[0];
    var trips = rows.slice(1).filter(function(r){ return r[0]; }).map(function(r) {
      var obj = {};
      headers.forEach(function(h,i){ obj[h] = r[i]; });
      if (obj.startDate instanceof Date) obj.startDate = formatDate(obj.startDate);
      if (obj.endDate instanceof Date) obj.endDate = formatDate(obj.endDate);
      return obj;
    });
    var enrollSh = ss.getSheetByName('TripEnrollments');
    var enrollRows = enrollSh ? enrollSh.getDataRange().getValues().slice(1).filter(function(r){return r[0];}) : [];
    var lineUid = data.lineUid;
    trips.forEach(function(t) {
      var myEnrolls = enrollRows.filter(function(r){ return String(r[1]) === String(t.id); });
      t.enrollCount = myEnrolls.length;
      t.spotsLeft = Math.max(0, Number(t.maxSpots||20) - t.enrollCount);
      t.minSpots = Number(t.minSpots||0);
      t.minReached = t.minSpots > 0 ? t.enrollCount >= t.minSpots : true;
      if (lineUid) {
        var myRow = myEnrolls.find(function(r){ return String(r[2]) === String(lineUid); });
        t.enrolled = !!myRow;
        t.myPlan = myRow ? myRow[5] : null;
      }
    });
    var result = data.password
      ? trips
      : trips.filter(function(t){ return t.active === true || t.active === 'TRUE'; });
    return { ok: true, trips: result };
  } catch(e) { return { ok: false, error: e.message }; }
}

function enrollTrip(data) {
  try {
    var lineUid = data.lineUid, tripId = data.tripId, plan = data.plan;
    if (!lineUid || !tripId || !plan) return { ok: false, error: '參數不完整' };
    var ss = SpreadsheetApp.openById(CFG.SHEET_ID);
    var memberSh = ss.getSheetByName('Members');
    if (!memberSh) return { ok: false, error: '找不到學員資料表' };
    var memberRows = memberSh.getDataRange().getValues();
    var memberRow = null;
    for (var i = 1; i < memberRows.length; i++) {
      if (String(memberRows[i][C.M.UID-1]) === String(lineUid)) { memberRow = memberRows[i]; break; }
    }
    if (!memberRow) return { ok: false, error: '請先完成學員登記' };
    initTripSheets();
    var enrollSh = ss.getSheetByName('TripEnrollments');
    var enrollRows = enrollSh.getDataRange().getValues().slice(1).filter(function(r){ return r[0]; });
    if (enrollRows.some(function(r){ return String(r[1])===String(tripId) && String(r[2])===String(lineUid); })) {
      return { ok: false, error: '您已報名此旅行' };
    }
    var tripSh = ss.getSheetByName('Trips');
    var tripRows = tripSh.getDataRange().getValues();
    var tripRow = null;
    for (var j = 1; j < tripRows.length; j++) {
      if (String(tripRows[j][0]) === String(tripId)) { tripRow = tripRows[j]; break; }
    }
    if (!tripRow) return { ok: false, error: '找不到旅行' };
    // Find column indices by header
    var tripHeaders = tripRows[0];
    var idxMax = tripHeaders.indexOf('maxSpots');
    var idxMin = tripHeaders.indexOf('minSpots');
    var maxSpots = Number(tripRow[idxMax >= 0 ? idxMax : 9] || 20);
    var minSpots = idxMin >= 0 ? Number(tripRow[idxMin] || 0) : 0;
    var tripTitle = String(tripRow[2] || '');
    var currentCount = enrollRows.filter(function(r){ return String(r[1])===String(tripId); }).length;
    if (currentCount >= maxSpots) return { ok: false, error: '名額已滿' };
    enrollSh.appendRow(['te_'+Date.now(), tripId, lineUid,
      memberRow[C.M.NAME-1]||memberRow[C.M.DISPLAY-1]||'',
      memberRow[C.M.PHONE-1]||'', plan, new Date().toISOString()]);
    var newCount = currentCount + 1;
    // If minSpots just reached, notify all enrolled students for payment
    if (minSpots > 0 && newCount === minSpots) {
      var allEnrollRows = enrollSh.getDataRange().getValues().slice(1).filter(function(r){
        return r[0] && String(r[1]) === String(tripId);
      });
      // Get UIDs of all enrolled members from Members sheet
      var memberShAll = ss.getSheetByName('Members');
      var allMemberRows = memberShAll ? memberShAll.getDataRange().getValues().slice(1) : [];
      allEnrollRows.forEach(function(er) {
        var enrolledUid = String(er[2]);
        var phone = String(er[4]);
        var lastFive = phone.length >= 5 ? phone.slice(-5) : phone;
        sendLine(enrolledUid,
          '🎉 【出團確認】' + tripTitle + '\n\n' +
          '人數已達出團標準（' + minSpots + ' 人），旅行確認出發！\n\n' +
          '請完成報名費收款，匯款後請告知末五碼：' + lastFive + '\n\n' +
          '如有疑問請聯繫教練，謝謝！');
      });
    }
    return { ok: true, message: '報名成功' };
  } catch(e) { return { ok: false, error: e.message }; }
}

function cancelTripEnroll(data) {
  try {
    var lineUid = data.lineUid, tripId = data.tripId;
    if (!lineUid || !tripId) return { ok: false, error: '參數不完整' };
    initTripSheets();
    var sh = SpreadsheetApp.openById(CFG.SHEET_ID).getSheetByName('TripEnrollments');
    var rows = sh.getDataRange().getValues();
    for (var i = rows.length - 1; i >= 1; i--) {
      if (String(rows[i][1])===String(tripId) && String(rows[i][2])===String(lineUid)) {
        sh.deleteRow(i + 1);
        return { ok: true, message: '已取消報名' };
      }
    }
    return { ok: false, error: '找不到報名紀錄' };
  } catch(e) { return { ok: false, error: e.message }; }
}

function myTripEnrollments(data) {
  try {
    var lineUid = data.lineUid;
    if (!lineUid) return { ok: false, error: '參數不完整' };
    initTripSheets();
    var ss = SpreadsheetApp.openById(CFG.SHEET_ID);
    var enrollRows = ss.getSheetByName('TripEnrollments').getDataRange().getValues()
      .slice(1).filter(function(r){ return r[0] && String(r[2])===String(lineUid); });
    if (!enrollRows.length) return { ok: true, enrollments: [] };
    var tripData = ss.getSheetByName('Trips').getDataRange().getValues();
    var tripHeaders = tripData[0];
    var tripMap = {};
    tripData.slice(1).filter(function(r){return r[0];}).forEach(function(r){
      var obj = {};
      tripHeaders.forEach(function(h,i){ obj[h] = r[i]; });
      if (obj.startDate instanceof Date) obj.startDate = formatDate(obj.startDate);
      if (obj.endDate instanceof Date) obj.endDate = formatDate(obj.endDate);
      tripMap[String(r[0])] = obj;
    });
    var result = enrollRows.map(function(r) {
      return Object.assign({}, tripMap[String(r[1])]||{}, { myPlan: r[5], enrolledAt: r[6] });
    });
    return { ok: true, enrollments: result };
  } catch(e) { return { ok: false, error: e.message }; }
}

function getTripEnrollList(data) {
  try {
    var p = checkPassword(data.password);
    if (!p.ok) return p;
    var tripId = data.tripId;
    if (!tripId) return { ok: false, error: '參數不完整' };
    initTripSheets();
    var rows = SpreadsheetApp.openById(CFG.SHEET_ID).getSheetByName('TripEnrollments')
      .getDataRange().getValues().slice(1)
      .filter(function(r){ return r[0] && String(r[1])===String(tripId); });
    var list = rows.map(function(r) {
      return { id:r[0], tripId:r[1], name:r[3], phone:r[4], plan:r[5], at:r[6] };
    });
    return { ok: true, list: list };
  } catch(e) { return { ok: false, error: e.message }; }
}

function updateTrip(data) {
  try {
    var p = checkPassword(data.password);
    if (!p.ok) return p;
    initTripSheets();
    var sh = SpreadsheetApp.openById(CFG.SHEET_ID).getSheetByName('Trips');
    var rows = sh.getDataRange().getValues();
    var headers = rows[0];
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.tripId)) {
        if (data.active !== undefined) {
          var col = headers.indexOf('active');
          if (col >= 0) sh.getRange(i+1, col+1).setValue(data.active);
        }
        if (data.minSpots !== undefined) {
          var colMin = headers.indexOf('minSpots');
          if (colMin >= 0) sh.getRange(i+1, colMin+1).setValue(Number(data.minSpots)||0);
        }
        return { ok: true };
      }
    }
    return { ok: false, error: '找不到旅行' };
  } catch(e) { return { ok: false, error: e.message }; }
}

function deleteTrip(data) {
  try {
    var p = checkPassword(data.password);
    if (!p.ok) return p;
    var tripId = data.tripId;
    if (!tripId) return { ok: false, error: '參數不完整' };
    initTripSheets();
    var ss = SpreadsheetApp.openById(CFG.SHEET_ID);
    var sh = ss.getSheetByName('Trips');
    var rows = sh.getDataRange().getValues();
    var found = false;
    for (var i = rows.length - 1; i >= 1; i--) {
      if (String(rows[i][0]) === String(tripId)) { sh.deleteRow(i + 1); found = true; break; }
    }
    if (!found) return { ok: false, error: '找不到旅行' };
    var esh = ss.getSheetByName('TripEnrollments');
    if (esh) {
      var erows = esh.getDataRange().getValues();
      for (var j = erows.length - 1; j >= 1; j--) {
        if (String(erows[j][1]) === String(tripId)) esh.deleteRow(j + 1);
      }
    }
    return { ok: true };
  } catch(e) { return { ok: false, error: e.message }; }
}
