const DB_ID = "1OI8FL060ASLokmdHHEr8E1_bkwIeeE1FvFZ362TNqOA";

function getDB_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getCustomerListApi() {
  const ss = getDB_();
  const sheet = ss.getSheetByName("Customers");
  if (!sheet) return [];

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  return values.slice(1)
    .filter(r => r[0] || r[1])
    .map(r => ({
      customerName: formatText_(r[0]),
      phone: formatText_(r[1]),
      firstVisit: formatDateText_(r[2]),
      recentVisit: formatDateText_(r[3]),
      visitCount: Number(r[4] || 0),
      cancelCount: Number(r[5] || 0),
      noShowCount: Number(r[6] || 0),
      grade: formatText_(r[7]),
      tag: formatText_(r[8]),
      favoriteMenu: formatText_(r[9]),
      allergy: formatText_(r[10]),
      memo: formatText_(r[11]),
      updatedAt: formatDateTimeText_(r[12])
    }))
    .sort((a, b) => String(b.recentVisit).localeCompare(String(a.recentVisit)));
}

function getCustomerDetailApi(phone) {
  const targetPhone = normalizePhone_(phone);
  if (!targetPhone) return { ok:false, message:"전화번호가 없습니다." };

  const customers = getCustomerListApi();
  const customer = customers.find(c => normalizePhone_(c.phone) === targetPhone);

  if (!customer) {
    return { ok:false, message:"고객정보를 찾을 수 없습니다." };
  }

  const reservations = getCustomerReservations_(targetPhone);
  const histories = getCustomerHistories_(reservations.map(r => r.reservationNo));

  return {
    ok:true,
    customer,
    reservations,
    histories
  };
}

function saveCustomerMemoApi(phone, memo) {
  const ss = getDB_();
  const sheet = ss.getSheetByName("Customers");
  if (!sheet) return { ok:false, message:"Customers 시트가 없습니다." };

  const targetPhone = normalizePhone_(phone);
  if (!targetPhone) return { ok:false, message:"전화번호가 없습니다." };

  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (normalizePhone_(values[i][1]) === targetPhone) {
      sheet.getRange(i + 1, 12).setValue(memo || "");
      sheet.getRange(i + 1, 13).setValue(nowText_());

      return {
        ok:true,
        message:"고객 메모가 저장되었습니다."
      };
    }
  }

  return { ok:false, message:"고객을 찾을 수 없습니다." };
}

function getCustomerReservations_(phone) {
  const ss = getDB_();
  const sheet = ss.getSheetByName("Reservations");
  if (!sheet) return [];

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  return values.slice(1)
    .filter(r => normalizePhone_(r[5]) === phone)
    .map(r => ({
      reservationNo: formatText_(r[0]),
      createdAt: formatDateTimeText_(r[1]),
      reserveDate: formatDateText_(r[2]),
      reserveTime: formatTimeText_(r[3]),
      customerName: formatText_(r[4]),
      phone: formatText_(r[5]),
      people: formatText_(r[6]),
      seat: formatText_(r[7]),
      eventType: formatText_(r[8]),
      menu: formatText_(r[9]),
      deposit: formatText_(r[10]),
      depositStatus: formatText_(r[11]),
      reserveStatus: formatText_(r[12]),
      staff: formatText_(r[13]),
      memo: formatText_(r[14])
    }))
    .sort((a, b) =>
      String(b.reserveDate + " " + b.reserveTime)
        .localeCompare(String(a.reserveDate + " " + a.reserveTime))
    );
}

function getCustomerHistories_(reservationNos) {
  const ss = getDB_();
  const sheet = ss.getSheetByName("ReservationHistory");

  if (!sheet || !reservationNos || !reservationNos.length) return [];

  const noMap = {};
  reservationNos.forEach(no => {
    if (no) noMap[String(no)] = true;
  });

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  return values.slice(1)
    .filter(r => noMap[String(r[1] || "")])
    .map(r => ({
      loggedAt: formatDateTimeText_(r[0]),
      reservationNo: formatText_(r[1]),
      action: formatText_(r[2]),
      field: formatText_(r[3]),
      before: formatText_(r[4]),
      after: formatText_(r[5]),
      staff: formatText_(r[6])
    }))
    .sort((a, b) => String(b.loggedAt).localeCompare(String(a.loggedAt)));
}

function normalizePhone_(value) {
  return String(value || "").replace(/[^0-9]/g, "");
}

function formatText_(value) {
  if (value === null || value === undefined) return "";
  return String(value);
}

function formatDateText_(value) {
  if (!value) return "";

  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }

  return String(value);
}

function formatDateTimeText_(value) {
  if (!value) return "";

  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  }

  return String(value);
}

function formatTimeText_(value) {
  if (!value) return "";

  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "HH:mm");
  }

  return String(value);
}

function nowText_() {
  return Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd HH:mm:ss"
  );
}
function saveCustomerProfileApi(phone, data) {
     data = data || {};
     
  const ss = getDB_();
  const sheet = ss.getSheetByName("Customers");
  if (!sheet) return { ok:false, message:"Customers 시트가 없습니다." };

  const targetPhone = normalizePhone_(phone);
  if (!targetPhone) return { ok:false, message:"전화번호가 없습니다." };

  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (normalizePhone_(values[i][1]) === targetPhone) {

      sheet.getRange(i + 1, 9).setValue(data.tag || "");          // I 고객태그
      sheet.getRange(i + 1, 10).setValue(data.favoriteMenu || ""); // J 즐겨찾는메뉴
      sheet.getRange(i + 1, 11).setValue(data.allergy || "");      // K 알레르기
      sheet.getRange(i + 1, 13).setValue(nowText_());              // M 최근수정일시

      return { ok:true, message:"고객 프로필이 저장되었습니다." };
    }
  }

  return { ok:false, message:"고객을 찾을 수 없습니다." };
}
function searchCustomerByPhone(phone) {

  const targetPhone = normalizePhone_(phone);
  if (!targetPhone) return { found: false };

  const customers = getCustomerListApi();

  const customer = customers.find(function(c) {
    return normalizePhone_(c.phone) === targetPhone;
  });

  if (!customer) {
    return { found: false };
  }

  return {
    found: true,
    customerName: customer.customerName,
    phone: customer.phone,
    grade: customer.grade,
    tag: customer.tag,
    favoriteMenu: customer.favoriteMenu,
    allergy: customer.allergy,
    memo: customer.memo
  };
}