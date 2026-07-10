function sendSolapiSMS_(to, text) {
  const url = "https://api.solapi.com/messages/v4/send-many/detail";
  const date = new Date().toISOString();
  const salt = createSalt_();
  const signature = createSignature_(date, salt);

  const payload = {
    messages: [
      {
        to: cleanPhone_(to),
        from: SENDER,
        text: text
      }
    ]
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization:
        "HMAC-SHA256 apiKey=" + SOLAPI_API_KEY +
        ", date=" + date +
        ", salt=" + salt +
        ", signature=" + signature
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    const body = response.getContentText();

    Logger.log(body);

    if (code >= 200 && code < 300) {
  let messageId = "";

  try {
    const json = JSON.parse(body);

    if (json.groupInfo && json.groupInfo._id) {
      messageId = json.groupInfo._id;
    } else if (json.messageId) {
      messageId = json.messageId;
    } else if (json.messages && json.messages[0] && json.messages[0].messageId) {
      messageId = json.messages[0].messageId;
    }

  } catch(e) {
    messageId = "";
  }

  return {
    ok: true,
    message: "성공",
    messageId: messageId
  };
}

    return { ok: false, message: body };

  } catch (err) {
    return { ok: false, message: String(err.message || err) };
  }
}
function createSignature_(date, salt) {
  const message = date + salt;

  const byteSignature = Utilities.computeHmacSha256Signature(
    message,
    SOLAPI_API_SECRET
  );

  return byteSignature.reduce(function(str, chr) {
    chr = (chr < 0 ? chr + 256 : chr).toString(16);
    return str + (chr.length === 1 ? "0" : "") + chr;
  }, "");
}

function createSalt_() {
  return Utilities.getUuid().replace(/-/g, "");
}
function cleanPhone_(phone) {
  let p = String(phone || "").replace(/[^0-9]/g, "");

  if (p.length === 8) {
    p = "010" + p;
  }

  if (p.length === 10 && p.startsWith("10")) {
    p = "0" + p;
  }

  return p;
}
function sendCustomerSMS(phone, customerName, message, reservationNo) {

  Logger.log("전화번호 = " + phone);
  Logger.log("이름 = " + customerName);
  Logger.log("메시지 = " + message);

  const result = sendSolapiSMS_(phone, message);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("SMSHistory");

    if (sheet) {
      sheet.appendRow([
        Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd HH:mm:ss"),
        reservationNo || "",
        customerName || "",
        phone || "",
        detectSmsType_(message),
        message || "",
        Session.getActiveUser().getEmail() || "시스템",
        result && result.ok ? "성공" : "실패",
        result && result.messageId ? result.messageId : "",
        "",
        0
      ]);
    }
  } catch (err) {
    Logger.log("SMSHistory 저장 오류: " + err.message);
  }

  return result;
}
function detectSmsType_(message) {

  const text = String(message || "");

  if (text.includes("예약금")) return "예약금안내";

  if (text.includes("예약을 접수") || text.includes("예약이 접수"))
    return "예약접수";

  if (text.includes("예약이 확정") || text.includes("예약 확정"))
    return "예약확정";

  if (text.includes("내일 예약"))
    return "내일예약";

  if (text.includes("방문해 주셔서") || text.includes("방문 감사"))
    return "방문감사";

  if (text.includes("예약이 취소") || text.includes("예약 취소"))
    return "예약취소";

  if (text.includes("노쇼"))
    return "노쇼안내";

  return "직접작성";
}
function getSmsHistoryApi(filters) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SMSHistory");

  if (!sheet) {
    return { list: [], stats: { total: 0, today: 0, success: 0, fail: 0 } };
  }

  const values = sheet.getDataRange().getValues();

  if (values.length < 2) {
    return { list: [], stats: { total: 0, today: 0, success: 0, fail: 0 } };
  }

  let list = values.slice(1).map((r, i) => ({
    rowNo: i + 2,
    sentAt: r[0] instanceof Date
     ? Utilities.formatDate(r[0], TIMEZONE, "yyyy-MM-dd HH:mm:ss")
     : String(r[0] || ""),
    reservationNo: r[1] || "",
    customerName: r[2] || "",
    phone: r[3] || "",
    smsType: r[4] || "",
    message: r[5] || "",
    sender: r[6] || "",
    result: r[7] || "",
    messageId: r[8] || "",
    memo: r[9] || "",
    resendCount: r[10] || 0
  }));

  filters = filters || {};

  if (filters.startDate) {
    list = list.filter(v => String(v.sentAt).slice(0, 10) >= filters.startDate);
  }

  if (filters.endDate) {
    list = list.filter(v => String(v.sentAt).slice(0, 10) <= filters.endDate);
  }

  if (filters.customerName) {
    list = list.filter(v => String(v.customerName).includes(filters.customerName));
  }

  if (filters.phone) {
    const q = String(filters.phone).replace(/[^0-9]/g, "");
    list = list.filter(v => String(v.phone).replace(/[^0-9]/g, "").includes(q));
  }

  if (filters.smsType) {
    list = list.filter(v => String(v.smsType) === String(filters.smsType));
  }

  if (filters.result) {
    list = list.filter(v => String(v.result).includes(filters.result));
  }

  const today = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd");

  const stats = {
    total: list.length,
    today: list.filter(v => String(v.sentAt).slice(0, 10) === today).length,
    success: list.filter(v => String(v.result).includes("성공")).length,
    fail: list.filter(v => String(v.result).includes("실패")).length
  };

  return {
    list: list.reverse(),
    stats: stats
  };
}

function resendSmsApi(rowNo) {
  const ss = getDB_();
  const sheet = ss.getSheetByName("SMSHistory");
  if (!sheet) throw new Error("SMSHistory 시트를 찾을 수 없습니다.");

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);

  function idx(name) {
    return headers.indexOf(name) + 1;
  }

  const row = Number(rowNo);
  const phone = sheet.getRange(row, idx("전화번호")).getValue();
  const text = sheet.getRange(row, idx("문자내용")).getValue();

  if (!phone || !text) {
    throw new Error("전화번호 또는 문자내용이 없습니다.");
  }

  const result = sendSolapiSMS_(phone, text);

  const resendCol = idx("재발송횟수");
  const oldCount = Number(sheet.getRange(row, resendCol).getValue() || 0);
  sheet.getRange(row, resendCol).setValue(oldCount + 1);

  return {
  ok: true,
  message: "재발송 완료",
  result
};
}

function buildSmsMessage_(type, data) {
  data = data || {};

  const storeName = "한국의집 롯데월드몰점";

  function v(x) {
    return String(x || "").trim();
  }

  function money(x) {
    const n = Number(String(x || "").replace(/[^0-9]/g, ""));
    if (!n) return "";
    return n.toLocaleString("ko-KR") + "원";
  }

  function line(label, value) {
    value = v(value);
    if (!value) return "";
    return label + ": " + value;
  }

  function joinLines(lines) {
    return lines
      .filter(function(x) {
        return x !== null && x !== undefined && String(x).trim() !== "";
      })
      .join("\n");
  }

  const customerName = v(data.customerName || data.name);
  const reserveDate = formatSmsDate_(data.reserveDate || data.date);
  const reserveTime = formatSmsTime_(data.reserveTime || data.time);
  const people = v(data.people);
  const seat = v(data.seat);
  const menu = v(data.menu);
  const depositAmount = money(data.depositAmount || data.deposit || data.totalPrice);

  if (type === "예약금안내") {
    return joinLines([
      "[한국의집 예약금 안내]",
      "",
      customerName + " 고객님,",
      "한국의집 예약이 접수되었습니다.",
      "",
      line("예약일", reserveDate),
      line("예약시간", reserveTime),
      line("인원", people ? people + "명" : ""),
      line("좌석", seat),
      line("메뉴", menu),
      line("예약금", depositAmount),
      "",
      "아래 계좌로 예약금 입금을 부탁드립니다.",
      "하나은행 602-910043-80104",
      "예금주: 한국의집 롯데월드몰점",
      "",
      "입금 확인 후 예약이 확정됩니다.",
      "",
      "감사합니다.",
      storeName + " 드림"
    ]);
  }

  if (type === "예약접수") {
    return joinLines([
      "[한국의집 예약 안내]",
      "",
      customerName + " 고객님,",
      "한국의집 예약을 접수하였습니다.",
      "",
      line("예약일", reserveDate),
      line("예약시간", reserveTime),
      line("인원", people ? people + "명" : ""),
      line("좌석", seat),
      line("메뉴", menu),
      "",
      "예약금 입금 확인 후 예약이 확정됩니다.",
      "",
      "감사합니다.",
      storeName + " 드림"
    ]);
  }

  if (type === "예약확정") {
    return joinLines([
      "[한국의집 예약 확정]",
      "",
      customerName + " 고객님,",
      "예약이 확정되었습니다.",
      "",
      line("예약일", reserveDate),
      line("예약시간", reserveTime),
      line("인원", people ? people + "명" : ""),
      line("좌석", seat),
      line("메뉴", menu),
      "",
      "예약 당일 정성껏 준비하겠습니다.",
      "즐거운 시간 되시길 바랍니다.",
      "",
      "감사합니다.",
      storeName + " 드림"
    ]);
  }

  if (type === "방문감사") {
    return joinLines([
      "[한국의집 방문 감사]",
      "",
      customerName + " 고객님,",
      "오늘 방문해 주셔서 진심으로 감사드립니다.",
      "",
      "만족스러운 시간이 되셨기를 바라며,",
      "다음에도 한식의 품격을 전해드릴 수 있도록",
      "정성껏 준비하겠습니다.",
      "",
      "가까운 시일 내 다시 뵐 수 있기를 기대하겠습니다.",
      "",
      "감사합니다.",
      storeName + " 드림"
    ]);
  }

  return "";
}
function getSmsTemplate(templateId) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("SmsTemplates");

  if (!sheet) {
    throw new Error("SmsTemplates 시트를 찾을 수 없습니다.");
  }

  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === String(templateId).trim()) {
      return {
        templateId: values[i][0],
        title: values[i][1],
        content: values[i][2]
      };
    }
  }

  return null;
}

function saveSmsTemplate(data) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("SmsTemplates");

  if (!sheet) {
    throw new Error("SmsTemplates 시트를 찾을 수 없습니다.");
  }

  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === String(data.templateId).trim()) {
      sheet.getRange(i + 1, 2).setValue(data.title);
      sheet.getRange(i + 1, 3).setValue(data.content);
      sheet.getRange(i + 1, 4).setValue(new Date());

      return { ok: true };
    }
  }

  sheet.appendRow([
    data.templateId,
    data.title,
    data.content,
    new Date()
  ]);

  return { ok: true };
}
function renderSmsTemplate(templateId, item) {

  const t = getSmsTemplate(templateId);
  if (!t) return "";

  let text = String(t.content || "").replaceAll("\\n", "\n");

  const reserveDate = formatSmsDate_(item.reserveDate);
  const reserveTime = formatSmsTime_(item.reserveTime);

  text = text.replaceAll("{고객명}", item.customerName || "");
  text = text.replaceAll("{예약일}", reserveDate);
  text = text.replaceAll("{예약시간}", reserveTime);
  text = text.replaceAll("{인원}", item.people ? item.people + "명" : "");
  text = text.replaceAll("{좌석}", item.seat || "");
  text = text.replaceAll("{메뉴}", item.menu || "");

  return text;
}
function formatSmsDate_(value) {

  if (!value) return "";

  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }

  return String(value).split(" ")[0];
}

function formatSmsTime_(value) {

  if (!value) return "";

  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "HH:mm");
  }

  const text = String(value);

  const m = text.match(/(\d{1,2}):(\d{2})/);
  if (m) {
    return m[1].padStart(2, "0") + ":" + m[2];
  }

  return text;
}