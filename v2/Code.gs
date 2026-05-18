// ================================================================
// HertzQuest 團練系統 — Backend (Google Apps Script)
// ================================================================
const CFG = {
  SHEET_ID:   'REPLACE_WITH_SHEET_ID',
  LINE_TOKEN: 'REPLACE_WITH_LINE_TOKEN',
  COACH_PWD:  'hertz2024',
  LIFF_ID:    'REPLACE_WITH_LIFF_ID',
};

// Column indices (1-based)
const C = {
  M: { ID:1,UID:2,NAME:3,PHONE:4,DISPLAY:5,PIC:6,STATUS:7,AT:8 },
  S: { ID:1,TITLE:2,DATE:3,TIME:4,LOC:5,MEET:6,DESC:7,COACH:8,MAX:9,LEFT:10,PRICE:11,ACTIVE:12,AT:13 },
  E: { ID:1,SID:2,UID:3,NAME:4,PHONE:5,STATUS:6,AT:7,TRANSFER:8,PAYMENT:9 },
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
    getEnrollList, notifyStudents, confirmPayment,
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
function initMember({lineUid, displayName, pictureUrl, realName, phone}) {
  const sh = sh_('Members');
  const rows = sh.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][C.M.UID-1] === lineUid) {
      // 更新 LINE 暱稱/頭像
      sh.getRange(i+1, C.M.DISPLAY).setValue(displayName||'');
      sh.getRange(i+1, C.M.PIC).setValue(pictureUrl||'');
      return {ok:true, member:toMember(rows[i]), isNew:false};
    }
  }

  // 新學員 — 如果沒有提交真實資料，先回傳 isNew:true 讓前端顯示表單
  if (!realName || !phone) return {ok:true, member:null, isNew:true};

  const id = 'M'+Date.now();
  sh.appendRow([id,lineUid,realName,phone,displayName||'',pictureUrl||'','active',new Date()]);
  return {ok:true, member:{id,lineUid,realName,phone,displayName}, isNew:true, justRegistered:true};
}

function getSessions({lineUid}) {
  const sRows = sh_('Sessions').getDataRange().getValues();
  const eRows = sh_('Enrollments').getDataRange().getValues();
  const today = new Date(); today.setHours(0,0,0,0);

  const enrolled = new Set();
  eRows.slice(1).forEach(r => {
    if (r[C.E.UID-1]===lineUid && r[C.E.STATUS-1]==='confirmed') enrolled.add(r[C.E.SID-1]);
  });

  const sessions = sRows.slice(1)
    .filter(r => String(r[C.S.ACTIVE-1]).toLowerCase()==='true')
    .filter(r => { try { return new Date(r[C.S.DATE-1]) >= today; } catch(e){ return false; } })
    .sort((a,b) => new Date(a[C.S.DATE-1])-new Date(b[C.S.DATE-1]))
    .map(r => ({...toSession(r), enrolled: enrolled.has(r[C.S.ID-1])}));

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
  if (eRows.slice(1).find(r => r[C.E.SID-1]===sessionId && r[C.E.UID-1]===lineUid && r[C.E.STATUS-1]==='confirmed')) {
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
  if (Number(sData[C.S.LEFT-1]) <= 0) return {ok:false,error:'名額已滿'};

  // 寫入報名
  const id='E'+Date.now();
  eSheet.appendRow([id,sessionId,lineUid,member[C.M.NAME-1],member[C.M.PHONE-1],'confirmed',new Date(),transferCode||'','pending']);
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
    .filter(r => r[C.E.UID-1]===lineUid && r[C.E.STATUS-1]==='confirmed')
    .map(r => { const s=sMap[r[C.E.SID-1]]; return s ? toSession(s) : null; })
    .filter(Boolean)
    .sort((a,b) => new Date(a.date)-new Date(b.date));

  return {ok:true, enrollments:list};
}

function cancelEnroll({lineUid, sessionId}) {
  const eSheet = sh_('Enrollments');
  const eRows = eSheet.getDataRange().getValues();
  for (let i=1;i<eRows.length;i++) {
    if (eRows[i][C.E.SID-1]===sessionId && eRows[i][C.E.UID-1]===lineUid && eRows[i][C.E.STATUS-1]==='confirmed') {
      eSheet.getRange(i+1, C.E.STATUS).setValue('cancelled');
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
  eRows.slice(1).forEach(r => { if(r[C.E.STATUS-1]==='confirmed') cnt[r[C.E.SID-1]]=(cnt[r[C.E.SID-1]]||0)+1; });

  const sessions = sRows.slice(1)
    .sort((a,b) => new Date(b[C.S.DATE-1])-new Date(a[C.S.DATE-1]))
    .map(r => ({...toSession(r), enrollCount:cnt[r[C.S.ID-1]]||0, active: String(r[C.S.ACTIVE-1]).toLowerCase()==='true'}));

  return {ok:true, sessions};
}

function createSession({password, title, date, time, location, meetingPoint, description, coach, maxSpots, price}) {
  if (!auth(password)) return noAuth();
  if (!title||!date) return {ok:false,error:'標題和日期為必填'};
  const id='S'+Date.now();
  const max=Number(maxSpots)||20;
  sh_('Sessions').appendRow([id,title,date,time||'',location||'',meetingPoint||'',description||'',coach||'',max,max,Number(price)||0,true,new Date()]);
  return {ok:true,id,message:`「${title}」已建立`};
}

function updateSession({password, sessionId, ...fields}) {
  if (!auth(password)) return noAuth();
  const sh = sh_('Sessions');
  const rows = sh.getDataRange().getValues();
  for (let i=1;i<rows.length;i++) {
    if (rows[i][C.S.ID-1]===sessionId) {
      const map = {title:C.S.TITLE,date:C.S.DATE,time:C.S.TIME,location:C.S.LOC,meetingPoint:C.S.MEET,description:C.S.DESC,coach:C.S.COACH,maxSpots:C.S.MAX,price:C.S.PRICE,active:C.S.ACTIVE};
      Object.entries(fields).forEach(([k,v]) => { if(map[k]) sh.getRange(i+1,map[k]).setValue(v); });
      return {ok:true};
    }
  }
  return {ok:false,error:'找不到活動'};
}

function deleteSession({password, sessionId}) {
  if (!auth(password)) return noAuth();
  return updateSession({password, sessionId, active:false});
}

function getEnrollList({password, sessionId}) {
  if (!auth(password)) return noAuth();
  const rows = sh_('Enrollments').getDataRange().getValues();
  const list = rows.slice(1)
    .filter(r => r[C.E.SID-1]===sessionId && r[C.E.STATUS-1]==='confirmed')
    .map(r => ({
      id:           r[C.E.ID-1],
      name:         r[C.E.NAME-1],
      phone:        r[C.E.PHONE-1],
      at:           String(r[C.E.AT-1]).slice(0,10),
      lineUid:      r[C.E.UID-1],
      transferCode: r[C.E.TRANSFER-1]||'',
      paymentStatus:r[C.E.PAYMENT-1]||'pending',
    }));
  return {ok:true, list};
}

function confirmPayment({password, enrollId}) {
  if (!auth(password)) return noAuth();
  const sh = sh_('Enrollments');
  const rows = sh.getDataRange().getValues();
  for (let i=1;i<rows.length;i++) {
    if (rows[i][C.E.ID-1]===enrollId) {
      sh.getRange(i+1, C.E.PAYMENT).setValue('confirmed');
      sh.getRange(i+1, C.E.STATUS).setValue('confirmed');
      return {ok:true};
    }
  }
  return {ok:false, error:'找不到報名紀錄'};
}

function notifyStudents({password, sessionId, message}) {
  if (!auth(password)) return noAuth();
  const rows = sh_('Enrollments').getDataRange().getValues();
  const uids = rows.slice(1)
    .filter(r => r[C.E.SID-1]===sessionId && r[C.E.STATUS-1]==='confirmed')
    .map(r => r[C.E.UID-1]);
  uids.forEach(uid => sendLine(uid, message));
  return {ok:true, sent:uids.length};
}

// ── 工具 ─────────────────────────────────────────────────────
function auth(p) { return p===CFG.COACH_PWD; }
function noAuth() { return {ok:false,error:'未授權'}; }
function sh_(name) { return SpreadsheetApp.openById(CFG.SHEET_ID).getSheetByName(name); }

function toSession(r) {
  return {
    id:r[0], title:r[1],
    date: String(r[2]).slice(0,10),
    time:r[3]||'', location:r[4]||'', meetingPoint:r[5]||'',
    description:r[6]||'', coach:r[7]||'',
    maxSpots:r[8], spotsLeft:r[9], price:r[10],
    active: String(r[11]).toLowerCase()==='true',
  };
}

function toMember(r) {
  return {id:r[0],lineUid:r[1],realName:r[2],phone:r[3],displayName:r[4]};
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
  init('Members',  ['ID','LINE_UID','真實姓名','電話','LINE暱稱','頭像','狀態','加入時間'],'#1565C0');
  init('Sessions', ['ID','標題','日期','時間','地點','集合地點','說明','教練','名額上限','剩餘名額','費用','開放報名','建立時間'],'#2E7D32');
  init('Enrollments',['ID','活動ID','LINE_UID','姓名','電話','狀態','報名時間'],'#6A1B9A');
  return '✅ Sheets 初始化完成！共建立 Members、Sessions、Enrollments 三張表。';
}
