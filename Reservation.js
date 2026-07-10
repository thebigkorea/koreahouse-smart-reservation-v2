const RESERVATION_TIMEZONE = "Asia/Seoul";
const SEAT_OCCUPANCY_MINUTES = 120;

const RESERVATION_STATUS = {
  RECEIVED: "예약접수",
  CONFIRMED: "예약확정",
  CANCELED: "예약취소",
  NO_SHOW: "노쇼"
};

const SEAT_STATUS = {
  RESERVED: "예약",
  RELEASED: "해제",
  CANCELED: "취소"
};

function createReservation(data) {
  validateReservationData_(data);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reservationSheet = ss.getSheetByName("Reservations");
  const seatsSheet = ss.getSheetByName("Seats");
  const reservationNo = generateReservationNo_();
  const now = new Date();
  const seatIds = normalizeSeatIds_(data.seats || data.seat || "");
  const conflict = findSeatConflict_(seatIds, data.reserveDate, data.reserveTime, "");

  if (conflict) {
    return {
      ok: false,
      message: "이미 예약된 좌석입니다: " + conflict.seat + " (" + conflict.reservationNo + ")"
    };
  }

  const row = [
    reservationNo,
    now,
    data.reserveDate || "",
    data.reserveTime || "",
    data.customerName || "",
    normalizePhone_(data.phone || ""),
    Number(data.people || 0),
    seatIds.join(", "),
    data.eventType || "",
    data.menu || "",
    Number(data.deposit || 0),
    data.depositStatus || "미입금",
    data.reserveStatus || RESERVATION_STATUS.RECEIVED,
    data.staff || "",
    data.memo || "",
    data.channel || "전화",
    "",
    "",
    "",
    "",
    now
  ];

  reservationSheet.appendRow(row);
  writeSeatsForReservation_(seatsSheet, data.reserveDate, data.reserveTime, seatIds, reservationNo, data.customerName || "", SEAT_STATUS.RESERVED);
  upsertCustomerFromReservation_(data);
  sortReservations_();
  writeReservationLog_(reservationNo, "예약등록", "", "", "", data.staff || "");

  return {
    ok: true,
    reservationNo: reservationNo,
    message: "예약이 등록되었습니다."
  };
}

function getAllReservations() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Reservations");
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, 21).getValues();

  return values
    .filter(function(row) {
      return row[12] !== RESERVATION_STATUS.CANCELED;
    })
    .map(function(row) {
      return {
        reservationNo: row[0],
        createdAt: formatDateTime_(row[1]),
        reserveDate: formatDate_(row[2]),
        reserveTime: formatTime_(row[3]),
        customerName: row[4],
        phone: formatPhoneDisplay_(row[5]),
        people: row[6],
        seat: row[7],
        seats: normalizeSeatIds_(row[7]),
        eventType: row[8],
        menu: row[9],
        deposit: row[10],
        depositStatus: row[11],
        reserveStatus: row[12],
        staff: row[13],
        memo: row[14],
        channel: row[15],
        confirmSms: row[16],
        beforeSms: row[17],
        visitDone: row[18],
        visitDoneAt: formatDateTime_(row[19]),
        updatedAt: formatDateTime_(row[20])
      };
    });
}

function getTodayReservationData() {
  const today = Utilities.formatDate(new Date(), RESERVATION_TIMEZONE, "yyyy-MM-dd");
  const list = getAllReservations().filter(function(r) {
    return r.reserveDate === today && r.reserveStatus !== RESERVATION_STATUS.CANCELED;
  });
  const people = list.reduce(function(sum, r) {
    return sum + Number(r.people || 0);
  }, 0);
  const depositWait = list.filter(function(r) {
    return r.depositStatus !== "입금완료";
  }).length;
  const confirmed = list.filter(function(r) {
    return r.reserveStatus === RESERVATION_STATUS.CONFIRMED;
  }).length;

  return {
    today: today,
    count: list.length,
    people: people,
    depositWait: depositWait,
    confirmed: confirmed,
    reservations: list
  };
}

function updateReservation(reservationNo, data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Reservations");
  const row = findReservationRow_(reservationNo);

  if (!row) {
    return { ok: false, message: "예약을 찾을 수 없습니다." };
  }

  const oldData = sheet.getRange(row, 1, 1, 21).getValues()[0];
  const reserveDate = data.reserveDate || oldData[2];
  const reserveTime = data.reserveTime || oldData[3];
  const seatIds = normalizeSeatIds_(data.seats || data.seat || oldData[7]);
  const conflict = findSeatConflict_(seatIds, reserveDate, reserveTime, reservationNo);

  if (conflict) {
    return {
      ok: false,
      message: "이미 예약된 좌석입니다: " + conflict.seat + " (" + conflict.reservationNo + ")"
    };
  }

  const newValues = [
    reservationNo,
    oldData[1],
    reserveDate,
    reserveTime,
    data.customerName || oldData[4],
    normalizePhone_(data.phone || oldData[5]),
    Number(data.people || oldData[6] || 0),
    seatIds.join(", "),
    data.eventType || oldData[8],
    data.menu || oldData[9],
    Number(data.deposit || oldData[10] || 0),
    data.depositStatus || oldData[11],
    data.reserveStatus || oldData[12],
    data.staff || oldData[13],
    data.memo || oldData[14],
    data.channel || oldData[15],
    oldData[16],
    oldData[17],
    oldData[18],
    oldData[19],
    new Date()
  ];

  sheet.getRange(row, 1, 1, 21).setValues([newValues]);
  refreshSeatForReservation_(reservationNo, newValues);
  updateCustomerFromReservation_(data);
  sortReservations_();
  writeReservationLog_(reservationNo, "예약수정", "전체", JSON.stringify(oldData), JSON.stringify(newValues), data.staff || "");

  return {
    ok: true,
    message: "예약이 수정되었습니다."
  };
}

function cancelReservation(reservationNo, staff, reason) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Reservations");
  const row = findReservationRow_(reservationNo);

  if (!row) {
    return { ok: false, message: "예약을 찾을 수 없습니다." };
  }

  sheet.getRange(row, 13).setValue(RESERVATION_STATUS.CANCELED);
  sheet.getRange(row, 21).setValue(new Date());
  setSeatRowsStatus_(reservationNo, SEAT_STATUS.CANCELED);
  writeReservationLog_(reservationNo, "예약취소", "예약상태", "", reason || "취소", staff || "");

  return {
    ok: true,
    message: "예약이 취소되었습니다."
  };
}

function releaseSeatsForReservation(reservationNo, staff) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Reservations");
  const row = findReservationRow_(reservationNo);

  if (!row) {
    return { ok: false, message: "예약을 찾을 수 없습니다." };
  }

  const oldSeats = sheet.getRange(row, 8).getValue();
  sheet.getRange(row, 8).setValue("");
  sheet.getRange(row, 21).setValue(new Date());
  setSeatRowsStatus_(reservationNo, SEAT_STATUS.RELEASED);
  writeReservationLog_(reservationNo, "좌석해제", "좌석", oldSeats || "", "", staff || "");

  return {
    ok: true,
    message: "좌석이 해제되었습니다."
  };
}

function isSeatsAvailable(reserveDate, reserveTime, seats, reservationNoToIgnore) {
  return !findSeatConflict_(normalizeSeatIds_(seats), reserveDate, reserveTime, reservationNoToIgnore || "");
}

function markDepositDone(reservationNo, staff) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Reservations");
  const row = findReservationRow_(reservationNo);

  if (!row) {
    return { ok: false, message: "예약을 찾을 수 없습니다." };
  }

  sheet.getRange(row, 12).setValue("입금완료");
  sheet.getRange(row, 21).setValue(new Date());
  writeReservationLog_(reservationNo, "입금완료", "입금상태", "미입금", "입금완료", staff || "");

  return {
    ok: true,
    message: "입금완료 처리되었습니다."
  };
}

function confirmReservation(reservationNo, staff) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Reservations");
  const row = findReservationRow_(reservationNo);

  if (!row) {
    return { ok: false, message: "예약을 찾을 수 없습니다." };
  }

  const depositStatus = sheet.getRange(row, 12).getValue();

  if (depositStatus !== "입금완료") {
    return {
      ok: false,
      message: "예약금을 먼저 입금완료 처리하세요."
    };
  }

  sheet.getRange(row, 13).setValue(RESERVATION_STATUS.CONFIRMED);
  sheet.getRange(row, 21).setValue(new Date());

  const data = sheet.getRange(row, 1, 1, 21).getValues()[0];

  const message = buildSmsMessage_("예약확정", {
    customerName: data[4],
    reserveDate: data[2],
    reserveTime: data[3],
    people: data[6],
    seat: data[7],
    menu: data[9]
  });

  const smsResult = sendCustomerSMS(
    data[5],
    data[4],
    message,
    reservationNo
  );

  sheet.getRange(row, 17).setValue("완료");

  writeReservationLog_(reservationNo, "예약확정", "예약상태", "", RESERVATION_STATUS.CONFIRMED, staff || "");

  return {
    ok: true,
    message: "예약확정 처리 및 문자 발송 완료",
    smsResult: smsResult
  };
}

function markNoShow(reservationNo, staff, reason) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Reservations");
  const row = findReservationRow_(reservationNo);

  if (!row) {
    return { ok: false, message: "예약을 찾을 수 없습니다." };
  }

  const phone = sheet.getRange(row, 6).getValue();
  const oldStatus = sheet.getRange(row, 13).getValue();
  sheet.getRange(row, 13).setValue(RESERVATION_STATUS.NO_SHOW);
  sheet.getRange(row, 21).setValue(new Date());
  setSeatRowsStatus_(reservationNo, SEAT_STATUS.RELEASED);
  increaseCustomerNoShowCount(phone);
  writeReservationLog_(reservationNo, "노쇼", "예약상태", oldStatus, RESERVATION_STATUS.NO_SHOW, (staff || "") + " / " + (reason || ""));

  return {
    ok: true,
    message: "노쇼 처리되었습니다."
  };
}

function findReservationRow_(reservationNo) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Reservations");
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) return null;

  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(reservationNo)) {
      return i + 2;
    }
  }

  return null;
}

function generateReservationNo_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Reservations");
  const today = Utilities.formatDate(new Date(), RESERVATION_TIMEZONE, "yyyyMMdd");
  const prefix = "KH-LWM-" + today + "-";
  const lastRow = sheet.getLastRow();
  let maxNo = 0;

  if (lastRow > 1) {
    const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

    values.forEach(function(row) {
      const no = row[0] ? row[0].toString() : "";

      if (no.indexOf(prefix) === 0) {
        const tail = Number(no.replace(prefix, ""));

        if (!isNaN(tail) && tail > maxNo) {
          maxNo = tail;
        }
      }
    });
  }

  return prefix + (maxNo + 1).toString().padStart(6, "0");
}

function validateReservationData_(data) {
  if (!data) throw new Error("예약 데이터가 없습니다.");
  if (!data.reserveDate) throw new Error("예약일은 필수입니다.");
  if (!data.reserveTime) throw new Error("예약시간은 필수입니다.");
  if (!data.customerName) throw new Error("고객명은 필수입니다.");
  if (!data.phone) throw new Error("전화번호는 필수입니다.");
  if (!data.people) throw new Error("인원은 필수입니다.");
}

function findSeatConflict_(seatIds, reserveDate, reserveTime, reservationNoToIgnore) {
  if (!seatIds || seatIds.length === 0) return null;

  const wantedStart = buildReservationDateTime_(reserveDate, reserveTime);
  const wantedEnd = addMinutes_(wantedStart, SEAT_OCCUPANCY_MINUTES);
  const wantedSeats = {};
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Reservations");
  const lastRow = sheet.getLastRow();

  seatIds.forEach(function(seat) {
    wantedSeats[seat] = true;
  });

  if (lastRow <= 1) return null;

  const values = sheet.getRange(2, 1, lastRow - 1, 21).getValues();

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const reservationNo = String(row[0] || "");
    const status = row[12];

    if (reservationNoToIgnore && reservationNo === String(reservationNoToIgnore)) continue;
    if (!isActiveReservationStatus_(status)) continue;

    const existingSeats = normalizeSeatIds_(row[7]);
    const overlappingSeat = existingSeats.find(function(seat) {
      return wantedSeats[seat];
    });

    if (!overlappingSeat) continue;

    const existingStart = buildReservationDateTime_(row[2], row[3]);
    const existingEnd = addMinutes_(existingStart, SEAT_OCCUPANCY_MINUTES);

    if (isOverlapping_(wantedStart, wantedEnd, existingStart, existingEnd)) {
      return {
        seat: overlappingSeat,
        reservationNo: reservationNo,
        reserveDate: formatDate_(row[2]),
        reserveTime: formatTime_(row[3])
      };
    }
  }

  return null;
}

function isActiveReservationStatus_(status) {
  return status !== RESERVATION_STATUS.CANCELED && status !== RESERVATION_STATUS.NO_SHOW;
}

function isOverlapping_(startA, endA, startB, endB) {
  return startA.getTime() < endB.getTime() && startB.getTime() < endA.getTime();
}

function buildReservationDateTime_(reserveDate, reserveTime) {
  const dateText = formatDate_(reserveDate);
  const timeText = formatTime_(reserveTime);
  const date = new Date(dateText + "T" + timeText + ":00+09:00");

  if (isNaN(date.getTime())) {
    throw new Error("예약일 또는 예약시간 형식이 올바르지 않습니다.");
  }

  return date;
}

function addMinutes_(date, minutes) {
  return new Date(date.getTime() + minutes * 60 * 1000);
}

function normalizeSeatIds_(seats) {
  const values = Array.isArray(seats) ? seats : String(seats || "").split(/[,/|、\n]+/);
  const seen = {};
  const result = [];

  values.forEach(function(value) {
    const seat = String(value || "").trim();

    if (!seat) return;
    if (seen[seat]) return;

    seen[seat] = true;
    result.push(seat);
  });

  return result;
}

function writeSeatsForReservation_(sheet, reserveDate, reserveTime, seatIds, reservationNo, customerName, status) {
  seatIds.forEach(function(seat) {
    sheet.appendRow([
      reserveDate || "",
      reserveTime || "",
      seat,
      reservationNo,
      customerName || "",
      status || SEAT_STATUS.RESERVED
    ]);
  });
}

function refreshSeatForReservation_(reservationNo, reservationValues) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Seats");
  setSeatRowsStatus_(reservationNo, SEAT_STATUS.RELEASED);
  writeSeatsForReservation_(
    sheet,
    reservationValues[2],
    reservationValues[3],
    normalizeSeatIds_(reservationValues[7]),
    reservationNo,
    reservationValues[4],
    reservationValues[12] === RESERVATION_STATUS.CANCELED ? SEAT_STATUS.CANCELED : SEAT_STATUS.RESERVED
  );
}

function setSeatRowsStatus_(reservationNo, status) {
  const seatsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Seats");
  const lastRow = seatsSheet.getLastRow();

  if (lastRow <= 1) return;

  const values = seatsSheet.getRange(2, 1, lastRow - 1, 6).getValues();

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][3]) === String(reservationNo)) {
      seatsSheet.getRange(i + 2, 6).setValue(status);
    }
  }
}

function updateCustomerFromReservation_(data) {
  upsertCustomerFromReservation_(data);
}

function upsertCustomerFromReservation_(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Customers");
  const phone = normalizePhone_(data.phone || "");

  if (!phone) return;

  const lastRow = sheet.getLastRow();

  if (lastRow > 1) {
    const values = sheet.getRange(2, 1, lastRow - 1, 11).getValues();

    for (let i = 0; i < values.length; i++) {
      const savedPhone = values[i][0] ? values[i][0].toString() : "";

      if (savedPhone === phone) {
        const row = i + 2;
        const count = Number(values[i][2] || 0) + 1;
        let grade = "일반";

        if (count >= 20) grade = "VIP";
        else if (count >= 10) grade = "Gold";
        else if (count >= 5) grade = "Silver";

        sheet.getRange(row, 2).setValue(data.customerName || values[i][1]);
        sheet.getRange(row, 3).setValue(count);
        sheet.getRange(row, 4).setValue(data.reserveDate || "");
        sheet.getRange(row, 5).setValue(grade);
        sheet.getRange(row, 8).setValue(normalizeSeatIds_(data.seats || data.seat || "").join(", "));
        sheet.getRange(row, 11).setValue(new Date());
        return;
      }
    }
  }

  sheet.appendRow([
    phone,
    data.customerName || "",
    1,
    data.reserveDate || "",
    "일반",
    0,
    0,
    normalizeSeatIds_(data.seats || data.seat || "").join(", "),
    "",
    "",
    new Date()
  ]);
}

function sortReservations_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Reservations");
  const lastRow = sheet.getLastRow();

  if (lastRow <= 2) return;

  sheet.getRange(2, 1, lastRow - 1, 21).sort([
    { column: 3, ascending: true },
    { column: 4, ascending: true }
  ]);
}

function writeReservationLog_(reservationNo, action, field, beforeValue, afterValue, staff) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ReservationHistory");

  sheet.appendRow([
    new Date(),
    reservationNo,
    action,
    field || "",
    beforeValue || "",
    afterValue || "",
    staff || "시스템"
  ]);
}

function normalizePhone_(phone) {
  let p = (phone || "").toString().replace(/[^0-9]/g, "");

  if (p.length === 10 && p.charAt(0) !== "0") {
    p = "0" + p;
  }

  return p;
}

function formatPhoneDisplay_(phone) {
  let p = (phone || "").toString().replace(/[^0-9]/g, "");

  if (p.length === 10 && p.charAt(0) !== "0") {
    p = "0" + p;
  }

  if (p.length === 11) {
    return p.slice(0, 3) + "-" + p.slice(3, 7) + "-" + p.slice(7);
  }

  return p;
}

function formatDate_(value) {
  if (!value) return "";

  if (value instanceof Date) {
    return Utilities.formatDate(value, RESERVATION_TIMEZONE, "yyyy-MM-dd");
  }

  return String(value).trim();
}

function formatTime_(value) {
  if (!value) return "";

  if (value instanceof Date) {
    return Utilities.formatDate(value, RESERVATION_TIMEZONE, "HH:mm");
  }

  return String(value).trim();
}

function formatDateTime_(value) {
  if (!value) return "";

  if (value instanceof Date) {
    return Utilities.formatDate(value, RESERVATION_TIMEZONE, "yyyy-MM-dd HH:mm:ss");
  }

  return String(value).trim();
}

function testCreateReservation() {
  const result = createReservation({
    reserveDate: "2026-06-29",
    reserveTime: "11:00",
    customerName: "테스트",
    phone: "010-1234-5678",
    people: 4,
    seat: "A1, A2",
    eventType: "일반행사",
    menu: "한우정식",
    deposit: 50000,
    staff: "담당자",
    memo: "테스트 예약",
    channel: "전화"
  });

  Logger.log(result);
}
function getReservedSeatStatus(reserveDate, reserveTime, reservationNoToIgnore) {
  const seatMap = {};
  if (!reserveDate || !reserveTime) return seatMap;

  const wantedStart = buildReservationDateTime_(reserveDate, reserveTime);
  const wantedEnd = addMinutes_(wantedStart, SEAT_OCCUPANCY_MINUTES);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Reservations");
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return seatMap;

  const values = sheet.getRange(2, 1, lastRow - 1, 21).getValues();

  values.forEach(function(row) {
    const reservationNo = String(row[0] || "");
    const status = row[12];

    if (reservationNoToIgnore && reservationNo === String(reservationNoToIgnore)) return;
    if (!isActiveReservationStatus_(status)) return;

    const existingStart = buildReservationDateTime_(row[2], row[3]);
    const existingEnd = addMinutes_(existingStart, SEAT_OCCUPANCY_MINUTES);

    if (!isOverlapping_(wantedStart, wantedEnd, existingStart, existingEnd)) return;

    normalizeSeatIds_(row[7]).forEach(function(seat) {
      seatMap[seat] = {
        reservationNo: reservationNo,
        customerName: row[4] || "",
        reserveTime: formatTime_(row[3]),
        people: row[6] || "",
        status: status || ""
      };
    });
  });

  return seatMap;
}
let calendarDate = new Date();

function changeCalendarMonth(offset){
  calendarDate.setMonth(calendarDate.getMonth() + offset);
  renderCalendar();
}

function renderCalendar(){

  const area = document.getElementById("calendarArea");
  const title = document.getElementById("calendarTitle");

  if(!area || !title) return;

  const year = calendarDate.getFullYear();
  const month = calendarDate.getMonth();

  title.textContent =
    year + "년 " + (month + 1) + "월";

  const firstDay = new Date(year, month, 1).getDay();
  const lastDate = new Date(year, month + 1, 0).getDate();

  let html = `
    <table class="calendar-table">
      <thead>
        <tr>
          <th>일</th>
          <th>월</th>
          <th>화</th>
          <th>수</th>
          <th>목</th>
          <th>금</th>
          <th>토</th>
        </tr>
      </thead>
      <tbody>
  `;

  let day = 1;

  for(let r=0;r<6;r++){

    html += "<tr>";

    for(let c=0;c<7;c++){

      if((r===0 && c<firstDay) || day>lastDate){

        html += "<td></td>";

      }else{

        html += `
          <td>
            <div class="calendar-day">
              ${day}
            </div>
          </td>
        `;

        day++;
      }

    }

    html += "</tr>";

    if(day>lastDate) break;

  }

  html += "</tbody></table>";

  area.innerHTML = html;
}