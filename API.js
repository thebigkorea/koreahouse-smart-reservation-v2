function saveReservation(data) {
  try {
    return createReservation(data);
  } catch (err) {
    return {
      ok: false,
      message: "서버 오류: " + err.message + "\n" + err.stack
    };
  }
}

function getReservationList() {
  return getAllReservations();
}

function getTodayDashboard() {
  return getTodayReservationData();
}

function checkSeatConflictApi(data) {
  const list = getAllReservations();

  const conflicts = list.filter(r =>
    r.reserveDate === data.reserveDate &&
    r.reserveTime === data.reserveTime &&
    r.seat === data.seat &&
    r.reserveStatus !== "예약취소"
  );

  return {
    conflict: conflicts.length > 0,
    count: conflicts.length,
    list: conflicts
  };
}

function cancelReservationApi(reservationNo, staff, reason) {
  return cancelReservation(reservationNo, staff, reason);
}

function updateReservationApi(data){
  return updateReservation(data);
}
function markDepositDoneApi(reservationNo, staff){
  return markDepositDone(reservationNo, staff);
}

function confirmReservationApi(reservationNo, staff){
  return confirmReservation(reservationNo, staff);
}
function markNoShowApi(reservationNo, staff, reason){
  return markNoShow(reservationNo, staff, reason);
}
function getSeatHoldMinutesApi() {
  return getSeatHoldMinutes_();
}

function saveSeatHoldMinutesApi(minutes) {
  return saveSeatHoldMinutes_(minutes);
}
function getReservedSeatStatusApi(data) {
  return getReservedSeatStatus(
    data.reserveDate,
    data.reserveTime,
    data.reservationNoToIgnore || ""
  );
}
function getCustomerListApi() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Customers");
  if (!sheet) return [];

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  return values.slice(1)
    .filter(r => r[0] || r[1])
    .map(r => ({
      customerName: r[0] || "",
      phone: r[1] || "",
      firstVisit: r[2] || "",
      recentVisit: r[3] || "",
      visitCount: r[4] || 0,
      cancelCount: r[5] || 0,
      noShowCount: r[6] || 0,
      grade: r[7] || "",
      memo: r[8] || "",
      updatedAt: r[9] || ""
    }));
}