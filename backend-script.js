/**
 * 논술고사 감독위원 신청 시스템 API
 */
const MASTER_ADMIN_KEY = "EXAM_PROCTOR_AES_SECURE_KEY_2026";

function doGet(e) {
  const action = e.parameter.action;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  
  // 1. 코드 인증 및 기존 제출 여부 조회 (index.html)
  if (action === "verify") {
    const inputCode = String(e.parameter.code || "").trim().toUpperCase();
    const data = sheet.getDataRange().getValues();
    
    let assignedName = "";
    let assignedPhone = "";
    let alreadySubmitted = false;
    let codeFound = false;

    for (let i = 0; i < data.length; i++) {
      const cellCode = String(data[i][0] || "").trim().toUpperCase();
      if (cellCode === "코드") continue;
      
      if (cellCode === inputCode) {
        codeFound = true;
        assignedName = data[i][1];
        assignedPhone = data[i][2];

        // 6열(암호화 데이터)이나 8열(제출일시)이 존재하면 이미 제출한 이력이 있음
        if (data[i][5] || data[i][7]) {
          alreadySubmitted = true;
        }
      }
    }

    if (codeFound) {
      return createResponse({
        success: true,
        assignedTo: assignedName,
        assignedPhone: assignedPhone,
        alreadySubmitted: alreadySubmitted
      });
    }

    return createResponse({ success: false, message: "존재하지 않는 코드입니다." });
  }

  // 2. 관리자 데이터 조회 (admin.html)
  if (action === "get_data") {
    const adminKey = e.parameter.admin_key;
    if (adminKey !== MASTER_ADMIN_KEY) {
      return createResponse({ success: false, message: "관리자 인증 실패" });
    }

    const data = sheet.getDataRange().getValues();
    const list = [];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === "코드") continue;
      list.push({
        code: data[i][0],
        assigned_name: data[i][1],
        assigned_phone: data[i][2],
        actual_name: data[i][3],
        actual_phone: data[i][4],
        encrypted_payload: data[i][5],
        status: data[i][6],
        submitted_at: data[i][7]
      });
    }
    return createResponse({ success: true, data: list });
  }

  return createResponse({ success: false, message: "잘못된 요청입니다." });
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    const action = body.action;

    // 관리자: 코드 등록
    if (action === "register_codes") {
      if (body.admin_key !== MASTER_ADMIN_KEY) {
        return createResponse({ success: false, message: "관리자 인증 실패" });
      }

      body.codes.forEach(item => {
        sheet.appendRow([
          item.code,
          item.name,
          item.phone,
          "", // actual_name
          "", // actual_phone
          "", // encrypted_payload
          "PENDING", // status
          ""  // submitted_at
        ]);
      });
      return createResponse({ success: true });
    }

    // 지원자: 설문 제출
    if (action === "submit") {
      const code = body.code;
      const data = sheet.getDataRange().getValues();
      let assignedName = "";
      let assignedPhone = "";
      let targetRowIndex = -1;

      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim().toUpperCase() === code) {
          assignedName = data[i][1];
          assignedPhone = data[i][2];

          // 최초 미제출 상태의 행을 찾음
          if (!data[i][5] && !data[i][7] && targetRowIndex === -1) {
            targetRowIndex = i + 1;
          }
        }
      }

      // 유출 여부 판정
      let status = "NORMAL";
      if (assignedName && body.actual_name && assignedName.trim() !== body.actual_name.trim()) {
        status = "SHARED_SUSPECT";
      }

      // 처음 작성하는 경우: 기존 빈 행에 저장
      if (targetRowIndex !== -1) {
        sheet.getRange(targetRowIndex, 4).setValue(body.actual_name);
        sheet.getRange(targetRowIndex, 5).setValue(body.actual_phone);
        sheet.getRange(targetRowIndex, 6).setValue(body.encrypted_payload);
        sheet.getRange(targetRowIndex, 7).setValue(status);
        sheet.getRange(targetRowIndex, 8).setValue(new Date().toISOString());
      } else {
        // 이미 제출 이력이 있어 재작성하는 경우: 덮어쓰지 않고 새로운 행을 아래에 추가
        if (status === "NORMAL") status = "RESUBMITTED";
        sheet.appendRow([
          code,
          assignedName,
          assignedPhone,
          body.actual_name,
          body.actual_phone,
          body.encrypted_payload,
          status,
          new Date().toISOString()
        ]);
      }

      return createResponse({ success: true });
    }

  } catch (err) {
    return createResponse({ success: false, message: err.toString() });
  }
}

function createResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
