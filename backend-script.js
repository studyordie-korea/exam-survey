/**
 * 시험감독 설문 시스템 - Apps Script Backend API
 * 구글 스프레드시트에 연동하여 실행합니다.
 */

const MASTER_ADMIN_KEY = "EXAM_PROCTOR_AES_SECURE_KEY_2026"; // index.html과 admin.html의 키와 일치

function doGet(e) {
  const action = e.parameter.action;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // 1. 코드 유효성 검증 (index.html)
  if (action === "verify") {
    const code = (e.parameter.code || "").toUpperCase();
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == code) {
        return createResponse({
          success: true,
          assignedTo: data[i][1] // 원래 배정된 이름
        });
      }
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
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const action = body.action;

    // 1. 관리자의 코드 일괄 등록
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

    // 2. 지원자의 설문 제출 (유출 감지 로직 탑재)
    if (action === "submit") {
      const code = body.code;
      const data = sheet.getDataRange().getValues();

      for (let i = 1; i < data.length; i++) {
        if (data[i][0] == code) {
          const assignedName = data[i][1];
          const actualName = body.actual_name;

          // 발급 대상자와 실제 입력자 이름이 다른 경우 유출로 판정
          let status = "NORMAL";
          if (assignedName && actualName && assignedName.trim() !== actualName.trim()) {
            status = "SHARED_SUSPECT"; // 초대받지 않은 제3자가 제출함 (원 배정자 추적 가능)
          }

          // DB 업데이트 (암호문 저장)
          sheet.getRange(i + 1, 4).setValue(body.actual_name);
          sheet.getRange(i + 1, 5).setValue(body.actual_phone);
          sheet.getRange(i + 1, 6).setValue(body.encrypted_payload);
          sheet.getRange(i + 1, 7).setValue(status);
          sheet.getRange(i + 1, 8).setValue(new Date().toISOString());

          return createResponse({ success: true });
        }
      }
      return createResponse({ success: false, message: "해당 코드를 찾을 수 없습니다." });
    }

  } catch (err) {
    return createResponse({ success: false, message: err.toString() });
  }
}

function createResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// 시트 초기 세팅 함수 (최초 1회 실행)
function setupSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.clear();
  sheet.appendRow([
    "코드", "원래배정자명", "원래배정연락처", 
    "실제입력자명", "실제입력연락처", "암호화데이터(AES)", 
    "상태(NORMAL/SHARED_SUSPECT)", "제출일시"
  ]);
}
