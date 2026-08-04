/**
 * 재활운동센터 회원관리 프로그램 - 백엔드 (Google Apps Script)
 *
 * 사용법:
 * 1. Google Sheets에서 새 스프레드시트를 만든다.
 * 2. 확장 프로그램 > Apps Script 를 연다.
 * 3. 이 파일 내용을 Code.gs 에 붙여넣고 저장한다.
 * 4. 배포 > 새 배포 > 유형: 웹 앱
 *    - 실행 계정: 나
 *    - 액세스 권한: 전체 허용(익명 포함)
 *    로 배포한다.
 * 5. 발급된 웹 앱 URL을 프런트엔드(index.html) 설정 화면에 입력한다.
 * 6. 설정 화면에서 "시트 초기화" 를 눌러 필요한 시트를 자동 생성한다.
 */

var SHEETS = {
  members: {
    name: '회원',
    headers: ['id', 'name', 'phone', 'birth', 'joinDate', 'passType', 'totalCount', 'remainingCount', 'startDate', 'expireDate', 'status', 'trainer', 'note'],
    labels: ['ID', '이름', '연락처', '생년월일', '등록일', '이용권종류', '총횟수', '잔여횟수', '시작일', '만료일', '상태', '담당트레이너', '비고']
  },
  visits: {
    name: '출석',
    headers: ['id', 'memberId', 'memberName', 'visitAt', 'trainer', 'note'],
    labels: ['ID', '회원ID', '회원이름', '방문일시', '담당트레이너', '비고']
  },
  rehab: {
    name: '재활기록',
    headers: ['id', 'memberId', 'memberName', 'date', 'area', 'exercise', 'painLevel', 'note', 'trainer'],
    labels: ['ID', '회원ID', '회원이름', '날짜', '통증부위', '운동내용', '통증강도', '메모', '담당트레이너']
  },
  payments: {
    name: '결제',
    headers: ['id', 'memberId', 'memberName', 'date', 'item', 'amount', 'method', 'staff', 'note'],
    labels: ['ID', '회원ID', '회원이름', '결제일', '항목', '금액', '결제수단', '담당자', '비고']
  }
};

function doGet(e) {
  return handle(e, 'GET');
}

function doPost(e) {
  return handle(e, 'POST');
}

function handle(e, method) {
  var action, data;
  try {
    if (method === 'GET') {
      action = e.parameter.action;
      data = e.parameter;
    } else {
      var body = JSON.parse((e.postData && e.postData.contents) || '{}');
      action = body.action;
      data = body.data || {};
    }
    var result = route(action, data);
    return jsonOut({ ok: true, result: result });
  } catch (err) {
    return jsonOut({ ok: false, error: err.message });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function route(action, data) {
  switch (action) {
    case 'ping': return ping();
    case 'setup': return setupSheets();
    case 'listMembers': return listMembers();
    case 'addMember': return addMember(data);
    case 'updateMember': return updateMember(data);
    case 'deleteMember': return deleteMember(data);
    case 'checkIn': return checkIn(data);
    case 'listVisits': return listVisits(data);
    case 'addRehabRecord': return addRehabRecord(data);
    case 'listRehabRecords': return listRehabRecords(data);
    case 'addPayment': return addPayment(data);
    case 'listPayments': return listPayments(data);
    case 'summary': return getSummary();
    default: throw new Error('알 수 없는 action: ' + action);
  }
}

function ping() {
  return { pong: true, time: new Date().toISOString(), sheets: Object.keys(SHEETS).map(function (k) { return SHEETS[k].name; }) };
}

function getSheet(key) {
  var cfg = SHEETS[key];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(cfg.name);
  if (!sheet) {
    sheet = ss.insertSheet(cfg.name);
    sheet.appendRow(cfg.labels);
    sheet.setFrozenRows(1);
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(cfg.labels);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function setupSheets() {
  Object.keys(SHEETS).forEach(getSheet);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var defaultNames = ['Sheet1', '시트1'];
  defaultNames.forEach(function (n) {
    var s = ss.getSheetByName(n);
    if (s && ss.getSheets().length > 1) {
      try { ss.deleteSheet(s); } catch (e) { }
    }
  });
  return { sheets: Object.keys(SHEETS).map(function (k) { return SHEETS[k].name; }) };
}

function readAll(key) {
  var cfg = SHEETS[key];
  var sheet = getSheet(key);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, cfg.headers.length).getValues();
  return values.map(function (row) {
    var obj = {};
    cfg.headers.forEach(function (h, i) { obj[h] = row[i]; });
    return obj;
  }).filter(function (o) { return o.id; });
}

function findRowIndexById(sheet, cfg, id) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  var idCol = cfg.headers.indexOf('id') + 1;
  var ids = sheet.getRange(2, idCol, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

function todayStr() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function getMemberById(id) {
  if (!id) return null;
  var members = readAll('members');
  for (var i = 0; i < members.length; i++) {
    if (String(members[i].id) === String(id)) return members[i];
  }
  return null;
}

/* ---------- 회원 ---------- */

function listMembers() {
  return readAll('members').sort(function (a, b) { return String(a.name).localeCompare(String(b.name), 'ko'); });
}

function addMember(data) {
  var cfg = SHEETS.members;
  var sheet = getSheet('members');
  var id = Utilities.getUuid();
  var total = (data.totalCount === '' || data.totalCount === undefined || data.totalCount === null) ? '' : Number(data.totalCount);
  var obj = {
    id: id,
    name: data.name || '',
    phone: data.phone || '',
    birth: data.birth || '',
    joinDate: data.joinDate || todayStr(),
    passType: data.passType || '',
    totalCount: total,
    remainingCount: total,
    startDate: data.startDate || todayStr(),
    expireDate: data.expireDate || '',
    status: '활성',
    trainer: data.trainer || '',
    note: data.note || ''
  };
  sheet.appendRow(cfg.headers.map(function (h) { return obj[h]; }));
  return obj;
}

function updateMember(data) {
  var cfg = SHEETS.members;
  var sheet = getSheet('members');
  var rowIdx = findRowIndexById(sheet, cfg, data.id);
  if (rowIdx === -1) throw new Error('회원을 찾을 수 없습니다.');
  var rowValues = sheet.getRange(rowIdx, 1, 1, cfg.headers.length).getValues()[0];
  var current = {};
  cfg.headers.forEach(function (h, i) { current[h] = rowValues[i]; });
  cfg.headers.forEach(function (h) {
    if (h === 'id') return;
    if (data[h] !== undefined) current[h] = data[h];
  });
  ['totalCount', 'remainingCount'].forEach(function (k) {
    if (current[k] !== '' && current[k] !== null && current[k] !== undefined && !isNaN(Number(current[k]))) {
      current[k] = Number(current[k]);
    }
  });
  sheet.getRange(rowIdx, 1, 1, cfg.headers.length).setValues([cfg.headers.map(function (h) { return current[h]; })]);
  return current;
}

function deleteMember(data) {
  var cfg = SHEETS.members;
  var sheet = getSheet('members');
  var rowIdx = findRowIndexById(sheet, cfg, data.id);
  if (rowIdx === -1) throw new Error('회원을 찾을 수 없습니다.');
  sheet.deleteRow(rowIdx);
  return { deleted: true, id: data.id };
}

/* ---------- 출석 ---------- */

function checkIn(data) {
  if (!data.memberId) throw new Error('memberId가 필요합니다.');
  var mCfg = SHEETS.members;
  var mSheet = getSheet('members');
  var rowIdx = findRowIndexById(mSheet, mCfg, data.memberId);
  if (rowIdx === -1) throw new Error('회원을 찾을 수 없습니다.');
  var rowValues = mSheet.getRange(rowIdx, 1, 1, mCfg.headers.length).getValues()[0];
  var member = {};
  mCfg.headers.forEach(function (h, i) { member[h] = rowValues[i]; });

  var vCfg = SHEETS.visits;
  var vSheet = getSheet('visits');
  var visit = {
    id: Utilities.getUuid(),
    memberId: member.id,
    memberName: member.name,
    visitAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
    trainer: data.trainer || '',
    note: data.note || ''
  };
  vSheet.appendRow(vCfg.headers.map(function (h) { return visit[h]; }));

  if (member.totalCount !== '' && member.totalCount !== null && !isNaN(Number(member.remainingCount))) {
    var remaining = Number(member.remainingCount);
    if (remaining > 0) remaining -= 1;
    member.remainingCount = remaining;
    if (remaining <= 0) member.status = '만료';
    mSheet.getRange(rowIdx, 1, 1, mCfg.headers.length).setValues([mCfg.headers.map(function (h) { return member[h]; })]);
  }
  return { visit: visit, member: member };
}

function listVisits(data) {
  var all = readAll('visits');
  if (data && data.memberId) all = all.filter(function (v) { return String(v.memberId) === String(data.memberId); });
  all.sort(function (a, b) { return String(b.visitAt).localeCompare(String(a.visitAt)); });
  var limit = (data && data.limit) ? Number(data.limit) : 100;
  return all.slice(0, limit);
}

/* ---------- 재활기록 ---------- */

function addRehabRecord(data) {
  var cfg = SHEETS.rehab;
  var sheet = getSheet('rehab');
  var member = getMemberById(data.memberId);
  var obj = {
    id: Utilities.getUuid(),
    memberId: data.memberId || '',
    memberName: member ? member.name : (data.memberName || ''),
    date: data.date || todayStr(),
    area: data.area || '',
    exercise: data.exercise || '',
    painLevel: (data.painLevel === undefined || data.painLevel === '') ? '' : Number(data.painLevel),
    note: data.note || '',
    trainer: data.trainer || ''
  };
  sheet.appendRow(cfg.headers.map(function (h) { return obj[h]; }));
  return obj;
}

function listRehabRecords(data) {
  var all = readAll('rehab');
  if (data && data.memberId) all = all.filter(function (r) { return String(r.memberId) === String(data.memberId); });
  all.sort(function (a, b) { return String(b.date).localeCompare(String(a.date)); });
  return all;
}

/* ---------- 결제 ---------- */

function addPayment(data) {
  var cfg = SHEETS.payments;
  var sheet = getSheet('payments');
  var member = getMemberById(data.memberId);
  var obj = {
    id: Utilities.getUuid(),
    memberId: data.memberId || '',
    memberName: member ? member.name : (data.memberName || ''),
    date: data.date || todayStr(),
    item: data.item || '',
    amount: data.amount === undefined ? 0 : Number(data.amount),
    method: data.method || '',
    staff: data.staff || '',
    note: data.note || ''
  };
  sheet.appendRow(cfg.headers.map(function (h) { return obj[h]; }));
  return obj;
}

function listPayments(data) {
  var all = readAll('payments');
  if (data && data.memberId) all = all.filter(function (p) { return String(p.memberId) === String(data.memberId); });
  if (data && data.month) all = all.filter(function (p) { return String(p.date).indexOf(data.month) === 0; });
  all.sort(function (a, b) { return String(b.date).localeCompare(String(a.date)); });
  return all;
}

/* ---------- 대시보드 요약 ---------- */

function daysBetween(a, b) {
  var d1 = new Date(a), d2 = new Date(b);
  return Math.round((d2 - d1) / (1000 * 60 * 60 * 24));
}

function getSummary() {
  var members = readAll('members');
  var visits = readAll('visits');
  var payments = readAll('payments');
  var today = todayStr();
  var thisMonth = today.slice(0, 7);

  var activeCount = members.filter(function (m) { return m.status === '활성'; }).length;
  var newThisMonth = members.filter(function (m) { return String(m.joinDate).indexOf(thisMonth) === 0; }).length;
  var revenueThisMonth = payments
    .filter(function (p) { return String(p.date).indexOf(thisMonth) === 0; })
    .reduce(function (sum, p) { return sum + (Number(p.amount) || 0); }, 0);
  var todayVisits = visits.filter(function (v) { return String(v.visitAt).indexOf(today) === 0; }).length;

  var expiringSoon = members.filter(function (m) {
    if (!m.expireDate || m.status !== '활성') return false;
    var diff = daysBetween(today, m.expireDate);
    return diff >= 0 && diff <= 7;
  }).sort(function (a, b) { return String(a.expireDate).localeCompare(String(b.expireDate)); });

  var lowRemaining = members.filter(function (m) {
    return m.status === '활성' && m.remainingCount !== '' && m.remainingCount !== null && Number(m.remainingCount) <= 2;
  });

  return {
    activeCount: activeCount,
    newThisMonth: newThisMonth,
    revenueThisMonth: revenueThisMonth,
    todayVisits: todayVisits,
    expiringSoon: expiringSoon,
    lowRemaining: lowRemaining
  };
}
