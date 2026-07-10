const TIMEZONE = "Asia/Seoul";

const SOLAPI_API_KEY =
  PropertiesService.getScriptProperties().getProperty("SOLAPI_API_KEY");

const SOLAPI_API_SECRET =
  PropertiesService.getScriptProperties().getProperty("SOLAPI_API_SECRET");

const SENDER =
  PropertiesService.getScriptProperties().getProperty("SOLAPI_SENDER");

function onOpen() {
SpreadsheetApp.getUi()
.createMenu("한국의집 예약관리")
.addItem("DB 초기 세팅", "setupKoreaHouseDB")
.addToUi();
}

function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || "";

  if (page === "customer-manager") {
    return HtmlService.createTemplateFromFile("customer-manager")
      .evaluate()
      .setTitle("한국의집 고객관리 CRM")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (page === "sms-manager") {
    return HtmlService.createTemplateFromFile("sms-manager-v2")
      .evaluate()
      .setTitle("한국의집 문자관리")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setTitle("한국의집 스마트 예약관리")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function setupKoreaHouseDB() {
const ss = SpreadsheetApp.getActiveSpreadsheet();

setupSheet_(ss, "Reservations", [
"예약번호",
"등록일시",
"예약일",
"예약시간",
"고객명",
"전화번호",
"인원",
"좌석",
"행사구분",
"메뉴",
"예약금",
"입금상태",
"예약상태",
"담당자",
"요청사항",
"예약경로",
"예약확정문자",
"전날안내문자",
"방문완료",
"방문완료시간",
"수정일시"
]);

setupSheet_(ss, "Customers", [
  "고객명",
  "전화번호",
  "첫예약일",
  "최근예약일",
  "예약횟수",
  "취소횟수",
  "노쇼횟수",
  "고객등급",
  "고객태그",
  "즐겨찾는메뉴",
  "알레르기",
  "메모",
  "최근수정일시"
]);

setupSheet_(ss, "Seats", [
"예약일",
"예약시간",
"좌석",
"예약번호",
"고객명",
"상태"
]);

setupSheet_(ss, "SMSHistory", [
  "발송일시",
  "예약번호",
  "고객명",
  "전화번호",
  "문자종류",
  "문자내용",
  "발송자",
  "발송결과",
  "Solapi Message ID",
  "비고",
  "재발송횟수"
]);

setupSheet_(ss, "ReservationHistory", [
"기록일시",
"예약번호",
"작업",
"변경항목",
"변경전",
"변경후",
"작업자"
]);

setupSheet_(ss, "Settings", [
"구분",
"항목",
"값",
"비고"
]);

setupSheet_(ss, "Users", [
"이름",
"권한",
"사용여부",
"비고"
]);

setupSheet_(ss, "Logs", [
"기록일시",
"작업자",
"작업",
"상세내용"
]);

setupDefaultSettings_();

SpreadsheetApp.getUi().alert("한국의집 예약관리 DB 세팅 완료");
}

function setupSheet_(ss, sheetName, headers) {
let sheet = ss.getSheetByName(sheetName);
if (!sheet) sheet = ss.insertSheet(sheetName);

sheet.clear();
sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
sheet.setFrozenRows(1);

sheet.getRange(1, 1, 1, headers.length)
.setFontWeight("bold")
.setHorizontalAlignment("center")
.setBackground("#18252d")
.setFontColor("#ffffff");

sheet.autoResizeColumns(1, headers.length);
}

function setupDefaultSettings_() {
const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");

const values = [
["매장", "브랜드", "한국의집", ""],
["매장", "매장명", "롯데월드몰점", ""],
["매장", "예약번호코드", "KH-LWM", ""],
["문자", "발신번호", "0232134560", ""],
["계좌", "은행명", "하나은행", ""],
["계좌", "계좌번호", "602-910043-80104", ""],
["계좌", "예금주", "한국의집 롯데월드몰점", ""],
["좌석", "좌석목록", "홀,청우정,룸,별실,기타", ""],
["예약", "운영시작시간", "10:00", ""],
["예약", "운영종료시간", "20:00", ""],
["예약", "시간간격", "30", "분 단위"],
["정책", "자동문자발송", "사용안함", "직원 수동 선택 발송"]
];

sheet.getRange(2, 1, values.length, 4).setValues(values);
}
function saveReservation(data) {
  const res = createReservation(data);
  rebuildCustomerStatistics();
  return res;
}

function getReservationList() {
  return getAllReservations();
}

function updateReservationApi(data) {
  const res = updateReservation(String(data.reservationNo || "").trim(), data);
  rebuildCustomerStatistics();
  return res;
}

function getReservedSeatStatusApi(data) {
  return getReservedSeatStatus(
    data.reserveDate,
    data.reserveTime,
    data.reservationNoToIgnore || ""
  );
}

function cancelReservationApi(reservationNo, staff, reason) {
  const res = cancelReservation(reservationNo, staff, reason);
  rebuildCustomerStatistics();
  return res;
}

function markDepositDoneApi(reservationNo, staff) {
  return markDepositDone(reservationNo, staff);
}

function confirmReservationApi(reservationNo, staff) {
  return confirmReservation(reservationNo, staff);
}

function markNoShowApi(reservationNo, staff, reason) {
  const res = markNoShow(reservationNo, staff, reason);
  rebuildCustomerStatistics();
  return res;
}
function searchCustomerApi(phone) {
  return searchCustomerByPhone(phone);
}
function rebuildCustomerStatistics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resSheet = ss.getSheetByName("Reservations");
  const cusSheet = ss.getSheetByName("Customers");

  if (!resSheet || !cusSheet) return;

  const resValues = resSheet.getDataRange().getValues();
  if (resValues.length < 2) return;

  const oldValues = cusSheet.getDataRange().getValues();
  const oldProfileMap = {};

for (let i = 1; i < oldValues.length; i++) {
  const phone = normalizePhone_(oldValues[i][1]);
  if (!phone) continue;

  oldProfileMap[phone] = {
    tag: oldValues[i][8] || "",
    favoriteMenu: oldValues[i][9] || "",
    allergy: oldValues[i][10] || "",
    memo: oldValues[i][11] || ""
  };
}

  const map = {};

  for (let i = 1; i < resValues.length; i++) {
    const row = resValues[i];

    const reserveDate = row[2];
    const customerName = row[4];
    const phone = normalizePhone_(row[5]);
    const status = String(row[12] || "");

    if (!phone) continue;

    if (!map[phone]) {
      map[phone] = {
        customerName: customerName || "",
        phone,
        firstVisit: "",
        recentVisit: "",
        visitCount: 0,
        cancelCount: 0,
        noShowCount: 0
      };
    }

    map[phone].customerName = customerName || map[phone].customerName;

    if (status.includes("취소")) {
      map[phone].cancelCount++;
      continue;
    }

    if (status.includes("노쇼")) {
      map[phone].noShowCount++;
      continue;
    }

    map[phone].visitCount++;

    if (!map[phone].firstVisit || String(reserveDate) < String(map[phone].firstVisit)) {
      map[phone].firstVisit = reserveDate;
    }

    if (!map[phone].recentVisit || String(reserveDate) > String(map[phone].recentVisit)) {
      map[phone].recentVisit = reserveDate;
    }
  }

  const now = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd HH:mm:ss");

  const output = Object.values(map).map(function(c) {
  const old = oldProfileMap[c.phone] || {};

  return [
    c.customerName,
    formatPhone_(c.phone),
    c.firstVisit,
    c.recentVisit,
    c.visitCount,
    c.cancelCount,
    c.noShowCount,
    getCustomerGrade_(c.visitCount),
    old.tag || "",
    old.favoriteMenu || "",
    old.allergy || "",
    old.memo || "",
    now
  ];
});

  cusSheet.getRange(2, 1, Math.max(cusSheet.getLastRow() - 1, 1), 13).clearContent();

  if (output.length) {
    cusSheet.getRange(2, 1, output.length, 13).setValues(output);
  }
}

function normalizePhone_(phone) {
  return String(phone || "").replace(/[^0-9]/g, "");
}

function formatPhone_(phone) {
  const p = normalizePhone_(phone);
  if (p.length === 11) {
    return p.slice(0, 3) + "-" + p.slice(3, 7) + "-" + p.slice(7);
  }
  return phone;
}

function getCustomerGrade_(visitCount) {
  const n = Number(visitCount || 0);

  if (n >= 30) return "VVIP";
  if (n >= 10) return "VIP";
  if (n >= 3) return "일반";
  if (n >= 1) return "신규";

  return "";
}

function getMostUsedKey_(obj) {
  let bestKey = "";
  let bestCount = 0;

  Object.keys(obj || {}).forEach(function(key) {
    if (obj[key] > bestCount) {
      bestKey = key;
      bestCount = obj[key];
    }
  });

  return bestKey;
}



function testSendSMS() {

  const result = sendSolapiSMS_(
    "01097998275",   // ← 본인 휴대폰 번호 입력
    "한국의집 예약관리 문자 발송 테스트입니다."
  );

  Logger.log(result);

  return result;
}


