/**
 * 재활운동센터 - 수업 리마인드 자동 문자 발송
 *
 * 전제:
 *  - "회원" 시트(Code.gs가 만드는 시트)에 이름/연락처가 이미 등록되어 있어야 발송 대상이 됩니다.
 *  - "스케줄표" 시트에 기존에 쓰시던 요일별/시간대별 표를 그대로 붙여넣어 사용합니다.
 *      1행: TIME/DATE | 일 | 월 | 화 | 수 | 목 | 금 | 토  (헤더)
 *      이후 시간마다 2행 1세트로 반복:
 *        - "9:00" 같은 시간이 적힌 행 = 회원 이름이 적힌 행
 *        - 그 바로 아래 행 = 메모(통증부위 등)가 적힌 행
 *      예: "SMC 김단비 회원님" 처럼 이름 앞뒤에 다른 글자가 붙어있어도,
 *          회원 시트의 이름이 셀 텍스트 안에 "포함"되어 있으면 매칭됩니다. (동명이인 없다는 전제)
 *  - 문자 발송은 알리고(Aligo, https://smartsms.aligo.in) API를 사용합니다.
 *    사업자등록 없이 개인 인증만으로 가입 가능하며, 발신번호 인증 후 API 키를 발급받습니다.
 *    나중에 사업자등록 후 카카오 알림톡으로 바꾸고 싶다면 sendSms() 함수 내부만 교체하면 됩니다.
 *
 * 동작:
 *  - "자동발송 트리거 설치"를 한 번 실행해두면 매시 정각마다 checkAndSendReminders()가 실행되어,
 *    "지금부터 12시간 뒤" 시간대에 스케줄표에 적힌 회원에게 리마인드 문자를 자동 발송합니다.
 *  - 같은 수업에 중복 발송되지 않도록 "발송이력" 시트에 기록을 남기고 확인합니다.
 */

var REMINDER_CONFIG = {
  SCHEDULE_SHEET: '스케줄표',
  LOG_SHEET: '발송이력',
  TIME_COL: 1,       // TIME/DATE 열 = A열
  FIRST_DAY_COL: 2,  // 일요일 열 = B열 (순서: 일 월 화 수 목 금 토). 실제 시트에서 밀려있으면 이 숫자만 조정하세요.
  REMIND_HOURS_BEFORE: 12
};

var REMINDER_DAY_NAMES = ['일', '월', '화', '수', '목', '금', '토'];

/* ---------- 메뉴 & 트리거 설치 ---------- */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔔 수업 리마인더')
    .addItem('지금 한 번 확인/발송 테스트', 'testCheckReminders')
    .addSeparator()
    .addItem('자동발송 트리거 설치(매시 정각)', 'installReminderTrigger')
    .addItem('자동발송 트리거 해제', 'removeReminderTriggers')
    .addToUi();
}

function installReminderTrigger() {
  removeReminderTriggers();
  ScriptApp.newTrigger('checkAndSendReminders').timeBased().everyHours(1).create();
  SpreadsheetApp.getUi().alert('자동발송 트리거를 설치했습니다. 매시 정각마다 12시간 뒤 수업을 확인해 문자를 보냅니다.');
}

function removeReminderTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'checkAndSendReminders') ScriptApp.deleteTrigger(t);
  });
}

/* ---------- 핵심 로직 ---------- */

function checkAndSendReminders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var scheduleSheet = ss.getSheetByName(REMINDER_CONFIG.SCHEDULE_SHEET);
  if (!scheduleSheet) return { sent: 0, reason: '스케줄표 시트를 찾을 수 없습니다.' };

  var target = new Date(Date.now() + REMINDER_CONFIG.REMIND_HOURS_BEFORE * 60 * 60 * 1000);
  var targetDay = REMINDER_DAY_NAMES[target.getDay()];
  var targetHour = target.getHours();
  var targetDateStr = Utilities.formatDate(target, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  var members = listMembers().filter(function (m) { return m.name && m.phone; });
  var slots = findScheduleSlots(scheduleSheet, targetDay, targetHour);

  var sentCount = 0;
  slots.forEach(function (slot) {
    matchMembersInText(slot.text, members).forEach(function (member) {
      if (wasAlreadySent(ss, targetDateStr, targetHour, member.name)) return;
      var msg = buildReminderMessage(member.name, targetDateStr, targetDay, targetHour);
      var ok = sendSms(member.phone, msg);
      logReminderSend(ss, targetDateStr, targetDay, targetHour, member.name, member.phone, ok);
      if (ok) sentCount++;
    });
  });
  return { sent: sentCount, checkedDay: targetDay, checkedHour: targetHour, slotsFound: slots.length };
}

function testCheckReminders() {
  var result = checkAndSendReminders();
  Logger.log(JSON.stringify(result));
  SpreadsheetApp.getUi().alert(
    '확인 대상: ' + result.checkedDay + '요일 ' + result.checkedHour + '시\n' +
    '스케줄표에서 찾은 칸: ' + (result.slotsFound || 0) + '개\n' +
    '문자 발송: ' + result.sent + '건\n' +
    (result.reason ? ('오류: ' + result.reason) : '(자세한 내용은 실행 로그에서 확인하세요)')
  );
  return result;
}

/* ---------- 스케줄표 파싱 ---------- */

function findScheduleSlots(sheet, targetDay, targetHour) {
  var dayColIndex = REMINDER_DAY_NAMES.indexOf(targetDay);
  if (dayColIndex === -1) return [];
  var col = REMINDER_CONFIG.FIRST_DAY_COL + dayColIndex;

  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return [];
  var timeValues = sheet.getRange(1, REMINDER_CONFIG.TIME_COL, lastRow, 1).getValues();

  var slots = [];
  for (var r = 0; r < timeValues.length; r++) {
    var hour = parseHourCell(timeValues[r][0]);
    if (hour === null || hour !== targetHour) continue;
    var rowNum = r + 1;
    var nameCell = sheet.getRange(rowNum, col, 1, 1).getValue();
    var noteCell = (rowNum + 1 <= lastRow) ? sheet.getRange(rowNum + 1, col, 1, 1).getValue() : '';
    var text = [nameCell, noteCell].filter(String).join(' ');
    if (text) slots.push({ row: rowNum, text: String(text) });
  }
  return slots;
}

function parseHourCell(value) {
  if (value === null || value === undefined || value === '') return null;
  if (Object.prototype.toString.call(value) === '[object Date]') return value.getHours();
  var m = String(value).trim().match(/^(\d{1,2}):00$/);
  return m ? Number(m[1]) : null;
}

function matchMembersInText(text, members) {
  return members.filter(function (m) { return text.indexOf(m.name) !== -1; });
}

/* ---------- 발송 이력 (중복 발송 방지) ---------- */

function getOrCreateLogSheet(ss) {
  var sheet = ss.getSheetByName(REMINDER_CONFIG.LOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(REMINDER_CONFIG.LOG_SHEET);
    sheet.appendRow(['발송일', '요일', '시간', '회원명', '연락처', '발송시각', '결과']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function wasAlreadySent(ss, dateStr, hour, memberName) {
  var sheet = getOrCreateLogSheet(ss);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  var values = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  return values.some(function (row) {
    return String(row[0]) === dateStr && Number(row[2]) === hour && String(row[3]) === memberName;
  });
}

function logReminderSend(ss, dateStr, day, hour, memberName, phone, ok) {
  getOrCreateLogSheet(ss).appendRow([dateStr, day, hour + ':00', memberName, phone, new Date(), ok ? '성공' : '실패(설정 확인)']);
}

function buildReminderMessage(memberName, dateStr, day, hour) {
  var hh = (hour < 10 ? '0' : '') + hour;
  return '[재활운동센터] ' + memberName + '님, ' + dateStr + '(' + day + ') ' + hh + ':00 예약된 수업 12시간 전 안내입니다. 시간 확인 부탁드려요 :)';
}

/* ---------- 문자 발송 (알리고) ----------
 * 나중에 카카오 알림톡으로 바꿀 때는 이 함수 내부만 교체하면 됩니다. */

function sendSms(phone, message) {
  var props = PropertiesService.getScriptProperties();
  var apiKey = props.getProperty('ALIGO_API_KEY');
  var userId = props.getProperty('ALIGO_USER_ID');
  var sender = props.getProperty('ALIGO_SENDER');

  if (!apiKey || !userId || !sender) {
    Logger.log('[SMS 미발송 - Script Properties 미설정] ' + phone + ' : ' + message);
    return false;
  }

  var options = {
    method: 'post',
    payload: {
      key: apiKey,
      user_id: userId,
      sender: sender,
      receiver: String(phone).replace(/-/g, ''),
      msg: message
    },
    muteHttpExceptions: true
  };
  try {
    var res = UrlFetchApp.fetch('https://apis.aligo.in/send/', options);
    var json = JSON.parse(res.getContentText());
    return String(json.result_code) === '1';
  } catch (err) {
    Logger.log('SMS 발송 실패: ' + err.message);
    return false;
  }
}
