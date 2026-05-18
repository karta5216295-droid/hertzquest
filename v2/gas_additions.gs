// ============================================================
// 以下是要加入 程式碼.gs 的新程式碼
// 請將 dispatch() 函數中加入新的 case，
// 並在檔案尾端加入下列輔助函數
// ============================================================

// ── 在 dispatch() 的 switch/if 區塊內新增以下 case ──────────
/*
  case 'enrollSession':   return enrollSession(data);
  case 'myTripEnrollments': return myTripEnrollments(data);
  case 'createTrip':      return createTrip(data);
  case 'getTrips':        return getTrips(data);
  case 'enrollTrip':      return enrollTrip(data);
  case 'cancelTripEnroll': return cancelTripEnroll(data);
  case 'getTripEnrollList': return getTripEnrollList(data);
  case 'updateTrip':      return updateTrip(data);
*/

// ── getSessions / allSessions ─────────────────────────────
// 在現有的 getSessions 函數中，需要額外回傳 uncertCount 和 certLevel（學員自己的）
// 修改：在 getSessions 最後處理 session 物件時加入：
//   s.uncertCount = enrollRows.filter(r => r[1]===s.id && r[5]==='無證').length;
//   if (lineUid) { const myRow = enrollRows.find(r=>r[1]===s.id&&r[2]===lineUid); s.certLevel = myRow ? myRow[5] : null; }
//
// ── enrollSession ─────────────────────────────────────────
// 修改現有 enrollSession 以接受 certLevel 並存入工作表
// Enrollments 欄位：id, sessionId, memberId, memberName, memberPhone, certLevel, enrolledAt

// ── 以下是完整的新函數，直接貼到 程式碼.gs 尾端 ───────────

// 格式化 GAS 時間（修正 1899 問題）
function formatTime(val) {
  if (!val) return '';
  if (val instanceof Date) {
    var h = val.getHours().toString().padStart(2,'0');
    var m = val.getMinutes().toString().padStart(2,'0');
    return h + ':' + m;
  }
  return String(val);
}

// 格式化日期
function formatDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    var y = val.getFullYear();
    var m = (val.getMonth()+1).toString().padStart(2,'0');
    var d = val.getDate().toString().padStart(2,'0');
    return y + '-' + m + '-' + d;
  }
  return String(val);
}

// 初始化旅行相關工作表
function initTripSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName('Trips')) {
    var sh = ss.insertSheet('Trips');
    sh.appendRow(['id','country','title','itinerary','inclusions','priceSingle','priceCouple','startDate','endDate','maxSpots','active','notes','createdAt']);
  }
  if (!ss.getSheetByName('TripEnrollments')) {
    var sh2 = ss.insertSheet('TripEnrollments');
    sh2.appendRow(['id','tripId','memberId','memberName','memberPhone','plan','enrolledAt']);
  }
}

// 建立旅行
function createTrip(data) {
  try {
    var p = checkPassword(data.password);
    if (!p.ok) return p;
    initTripSheets();
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Trips');
    var id = 'trip_' + Date.now();
    sh.appendRow([
      id,
      data.country || '',
      data.title || '',
      data.itinerary || '',
      data.inclusions || '',
      data.priceSingle || 0,
      data.priceCouple || 0,
      data.startDate || '',
      data.endDate || '',
      data.maxSpots || 20,
      true,
      data.notes || '',
      new Date().toISOString()
    ]);
    return { ok: true, message: '旅行已建立', id: id };
  } catch(e) { return { ok: false, error: e.message }; }
}

// 取得旅行列表
function getTrips(data) {
  try {
    initTripSheets();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('Trips');
    var rows = sh.getDataRange().getValues();
    if (rows.length <= 1) return { ok: true, trips: [] };
    var headers = rows[0];
    var trips = rows.slice(1).filter(function(r){ return r[0]; }).map(function(r) {
      var obj = {};
      headers.forEach(function(h,i){ obj[h] = r[i]; });
      // 格式化日期字串
      if (obj.startDate instanceof Date) obj.startDate = formatDate(obj.startDate);
      if (obj.endDate instanceof Date) obj.endDate = formatDate(obj.endDate);
      return obj;
    });

    // 加入報名數量
    var enrollSh = ss.getSheetByName('TripEnrollments');
    var enrollRows = enrollSh ? enrollSh.getDataRange().getValues().slice(1).filter(function(r){return r[0];}) : [];
    var lineUid = data.lineUid;

    trips.forEach(function(t) {
      var myEnrolls = enrollRows.filter(function(r){ return String(r[1]) === String(t.id); });
      t.enrollCount = myEnrolls.length;
      t.spotsLeft = Math.max(0, Number(t.maxSpots||20) - t.enrollCount);
      if (lineUid) {
        var myRow = myEnrolls.find(function(r){ return String(r[2]) === String(lineUid); });
        t.enrolled = !!myRow;
        t.myPlan = myRow ? myRow[5] : null;
      }
    });

    // 教練後台顯示全部，學員端只顯示開放中
    var result = data.password
      ? trips
      : trips.filter(function(t){ return t.active !== false && t.active !== 'FALSE' && t.active !== false; });

    return { ok: true, trips: result };
  } catch(e) { return { ok: false, error: e.message }; }
}

// 旅行報名
function enrollTrip(data) {
  try {
    var lineUid = data.lineUid;
    var tripId  = data.tripId;
    var plan    = data.plan;
    if (!lineUid || !tripId || !plan) return { ok: false, error: '參數不完整' };

    // 取得學員資料
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var memberSh = ss.getSheetByName('Members');
    if (!memberSh) return { ok: false, error: '找不到學員資料表' };
    var memberRows = memberSh.getDataRange().getValues();
    var memberRow = null;
    for (var i = 1; i < memberRows.length; i++) {
      if (String(memberRows[i][1]) === String(lineUid)) { memberRow = memberRows[i]; break; }
    }
    if (!memberRow) return { ok: false, error: '請先完成學員登記' };

    initTripSheets();
    var enrollSh = ss.getSheetByName('TripEnrollments');
    var enrollData = enrollSh.getDataRange().getValues();
    var enrollRows = enrollData.slice(1).filter(function(r){ return r[0]; });

    // 檢查是否已報名
    if (enrollRows.some(function(r){ return String(r[1])===String(tripId) && String(r[2])===String(lineUid); })) {
      return { ok: false, error: '您已報名此旅行' };
    }

    // 檢查名額
    var tripSh = ss.getSheetByName('Trips');
    var tripRows = tripSh.getDataRange().getValues();
    var tripRow = null;
    for (var j = 1; j < tripRows.length; j++) {
      if (String(tripRows[j][0]) === String(tripId)) { tripRow = tripRows[j]; break; }
    }
    if (!tripRow) return { ok: false, error: '找不到旅行' };
    var maxSpots = Number(tripRow[9] || 20);
    var currentCount = enrollRows.filter(function(r){ return String(r[1])===String(tripId); }).length;
    if (currentCount >= maxSpots) return { ok: false, error: '名額已滿' };

    var id = 'te_' + Date.now();
    enrollSh.appendRow([
      id, tripId, lineUid,
      memberRow[4] || memberRow[2] || '',  // realName fallback displayName
      memberRow[5] || '',                   // phone
      plan,
      new Date().toISOString()
    ]);
    return { ok: true, message: '報名成功' };
  } catch(e) { return { ok: false, error: e.message }; }
}

// 取消旅行報名
function cancelTripEnroll(data) {
  try {
    var lineUid = data.lineUid;
    var tripId  = data.tripId;
    if (!lineUid || !tripId) return { ok: false, error: '參數不完整' };
    initTripSheets();
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TripEnrollments');
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

// 我的旅行報名（學員端）
function myTripEnrollments(data) {
  try {
    var lineUid = data.lineUid;
    if (!lineUid) return { ok: false, error: '參數不完整' };
    initTripSheets();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var enrollSh = ss.getSheetByName('TripEnrollments');
    var enrollRows = enrollSh.getDataRange().getValues().slice(1).filter(function(r){ return r[0] && String(r[2])===String(lineUid); });
    if (!enrollRows.length) return { ok: true, enrollments: [] };

    var tripSh = ss.getSheetByName('Trips');
    var tripData = tripSh.getDataRange().getValues();
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
      var trip = tripMap[String(r[1])] || {};
      return Object.assign({}, trip, { myPlan: r[5], enrolledAt: r[6] });
    });
    return { ok: true, enrollments: result };
  } catch(e) { return { ok: false, error: e.message }; }
}

// 旅行報名名單（教練端）
function getTripEnrollList(data) {
  try {
    var p = checkPassword(data.password);
    if (!p.ok) return p;
    var tripId = data.tripId;
    if (!tripId) return { ok: false, error: '參數不完整' };
    initTripSheets();
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TripEnrollments');
    var rows = sh.getDataRange().getValues().slice(1).filter(function(r){ return r[0] && String(r[1])===String(tripId); });
    var list = rows.map(function(r) {
      return { id:r[0], tripId:r[1], name:r[3], phone:r[4], plan:r[5], at:r[6] };
    });
    return { ok: true, list: list };
  } catch(e) { return { ok: false, error: e.message }; }
}

// 更新旅行（開關狀態等）
function updateTrip(data) {
  try {
    var p = checkPassword(data.password);
    if (!p.ok) return p;
    initTripSheets();
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Trips');
    var rows = sh.getDataRange().getValues();
    var headers = rows[0];
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.tripId)) {
        if (data.active !== undefined) {
          var col = headers.indexOf('active');
          if (col >= 0) sh.getRange(i+1, col+1).setValue(data.active);
        }
        return { ok: true };
      }
    }
    return { ok: false, error: '找不到旅行' };
  } catch(e) { return { ok: false, error: e.message }; }
}

// ── 以下是修改現有函數的說明 ────────────────────────────────
// 1. 在 dispatch() 加入以下 else if：
//    } else if (action === 'createTrip')       { return createTrip(data);
//    } else if (action === 'getTrips')          { return getTrips(data);
//    } else if (action === 'enrollTrip')        { return enrollTrip(data);
//    } else if (action === 'cancelTripEnroll')  { return cancelTripEnroll(data);
//    } else if (action === 'myTripEnrollments') { return myTripEnrollments(data);
//    } else if (action === 'getTripEnrollList') { return getTripEnrollList(data);
//    } else if (action === 'updateTrip')        { return updateTrip(data);

// 2. 在現有 enrollSession() 中，加入 certLevel 到 Enrollments 表：
//    Enrollments 欄位改為: id, sessionId, memberId, memberName, memberPhone, certLevel, enrolledAt
//    並在 append 時加入 data.certLevel || ''

// 3. 在現有 getSessions() / allSessions() 中加入：
//    s.uncertCount = enrollRows.filter(r => r[1]===s.id && r[5]==='無證').length;
//    s.certLevel (學員自己的) = myRow ? myRow[5] : null;

// 4. 修正時間格式（1899 bug）：
//    在讀取 sessions 時，time 欄位用 formatTime() 處理：
//    s.time = formatTime(r[timeIndex]);
//    s.date = formatDate(r[dateIndex]);
