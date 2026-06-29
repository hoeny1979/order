// ══════════════════════════════════════════════════════════════════════════════
//  사진·영상·AI 창작 동호회 — 회원가입 자동화 시스템
//  Google Apps Script (Code.gs)
// ══════════════════════════════════════════════════════════════════════════════

// ── CONFIG ────────────────────────────────────────────────────────────────────
var CONFIG = {
  SHEET_NAME : "회원 신청 목록",
  ADMIN_EMAIL: "admin@example.com",       // ⚠️ 실제 관리자 이메일로 변경하세요
  CLUB_NAME  : "사진·영상·AI 창작 동호회",
  ADMIN_NAME : "운영진",
};

// ── HEADERS ───────────────────────────────────────────────────────────────────
var HEADERS = [
  "번호", "신청일시", "이름", "연락처", "이메일", "생년월일", "태어난 요일",
  "소속기관", "직책·직급", "관심분야", "사진·영상 수준", "AI 도구 수준",
  "가입동기", "이용약관", "개인정보동의", "창작물공유", "알림수신", "처리상태"
];

// ══════════════════════════════════════════════════════════════════════════════
//  doPost — POST 요청 수신 진입점
// ══════════════════════════════════════════════════════════════════════════════
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = getOrCreateSheet();
    var rowNum = appendRow(sheet, data);
    sendConfirmEmail(data);
    sendAdminEmail(data, rowNum);
    return jsonResponse({ status: "success", message: "신청이 접수되었습니다.", row: rowNum });
  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  }
}

// ══════════════════════════════════════════════════════════════════════════════
//  getOrCreateSheet — 시트 가져오기 / 없으면 생성
// ══════════════════════════════════════════════════════════════════════════════
function getOrCreateSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);

    // 헤더 행 삽입
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);

    // 헤더 스타일
    var headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setBackground("#1e3a5f")
               .setFontColor("#ffffff")
               .setFontWeight("bold")
               .setHorizontalAlignment("center")
               .setVerticalAlignment("middle");

    // 1행 고정
    sheet.setFrozenRows(1);

    // 열 너비 조정
    sheet.setColumnWidth(1,  50);   // 번호
    sheet.setColumnWidth(2,  150);  // 신청일시
    sheet.setColumnWidth(3,  80);   // 이름
    sheet.setColumnWidth(4,  120);  // 연락처
    sheet.setColumnWidth(5,  180);  // 이메일
    sheet.setColumnWidth(6,  110);  // 생년월일
    sheet.setColumnWidth(7,  90);   // 태어난 요일
    sheet.setColumnWidth(8,  120);  // 소속기관
    sheet.setColumnWidth(9,  90);   // 직책·직급
    sheet.setColumnWidth(10, 200);  // 관심분야
    sheet.setColumnWidth(11, 110);  // 사진·영상수준
    sheet.setColumnWidth(12, 110);  // AI수준
    sheet.setColumnWidth(13, 200);  // 가입동기
    sheet.setColumnWidth(14, 80);   // 이용약관
    sheet.setColumnWidth(15, 90);   // 개인정보
    sheet.setColumnWidth(16, 90);   // 창작물공유
    sheet.setColumnWidth(17, 90);   // 알림수신
    sheet.setColumnWidth(18, 100);  // 처리상태
  }

  return sheet;
}

// ══════════════════════════════════════════════════════════════════════════════
//  appendRow — 데이터 행 추가
// ══════════════════════════════════════════════════════════════════════════════
function appendRow(sheet, data) {
  var lastRow = sheet.getLastRow();
  var rowNum  = lastRow + 1;
  var seq     = lastRow; // 헤더 제외 순번

  var now = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
  var interests = Array.isArray(data.interests) ? data.interests.join(", ") : (data.interests || "");

  var rowData = [
    seq,
    now,
    data.name      || "",
    data.phone     || "",
    data.email     || "",
    data.birthDate || "",
    data.birthDay  || "",
    data.org       || "",
    data.pos    || "",
    interests,
    data.skillPhoto || "",
    data.skillAi    || "",
    data.intro  || "",
    data.terms1 ? "동의" : "미동의",
    data.terms2 ? "동의" : "미동의",
    data.terms3 ? "동의" : "미동의",
    data.terms4 ? "동의" : "미동의",
    "신청 접수",
  ];

  sheet.appendRow(rowData);

  // 짝수 행 배경색
  if (rowNum % 2 === 0) {
    sheet.getRange(rowNum, 1, 1, HEADERS.length).setBackground("#f0f4f8");
  }

  // 처리상태 셀 스타일
  var statusCell = sheet.getRange(rowNum, HEADERS.length);
  statusCell.setBackground("#fef9c3")
            .setFontColor("#854d0e")
            .setFontWeight("bold")
            .setHorizontalAlignment("center");

  return rowNum;
}

// ══════════════════════════════════════════════════════════════════════════════
//  sendConfirmEmail — 신청자 확인 메일
// ══════════════════════════════════════════════════════════════════════════════
function sendConfirmEmail(data) {
  if (!data.email) return;

  var interests = Array.isArray(data.interests) ? data.interests.join(", ") : (data.interests || "");
  var now = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");

  var subject = "[" + CONFIG.CLUB_NAME + "] 가입 신청이 접수되었습니다";

  var html = [
    '<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8">',
    '<style>',
    'body{font-family:"Noto Sans KR",sans-serif;background:#f8fafc;margin:0;padding:24px}',
    '.wrap{max-width:560px;margin:0 auto;background:#fff;border-radius:16px;overflow:hidden;box-shadow:0 2px 12px rgba(0,0,0,.1)}',
    '.hd{background:linear-gradient(135deg,#1e293b 0%,#0f4c81 60%,#14532d 100%);padding:32px;text-align:center;color:#fff}',
    '.hd .emoji{font-size:36px;margin-bottom:12px}',
    '.hd h1{font-size:20px;margin:0 0 6px;font-weight:700}',
    '.hd p{font-size:13px;color:#94a3b8;margin:0}',
    '.bd{padding:28px}',
    '.greeting{font-size:15px;color:#1e293b;margin-bottom:20px;line-height:1.7}',
    'table{width:100%;border-collapse:collapse;margin-bottom:20px}',
    'tr{border-bottom:1px solid #f1f5f9}',
    'td{padding:10px 8px;font-size:13px}',
    'td:first-child{color:#94a3b8;width:120px;font-weight:500}',
    'td:last-child{color:#1e293b;font-weight:600}',
    '.notice{background:#eff6ff;border:1.5px solid #bfdbfe;border-radius:10px;padding:16px;font-size:13px;color:#1d4ed8;line-height:1.7}',
    '.ft{text-align:center;padding:20px;font-size:12px;color:#94a3b8;border-top:1px solid #f1f5f9}',
    '</style></head><body>',
    '<div class="wrap">',
    '  <div class="hd">',
    '    <div class="emoji">📷🎬✨</div>',
    '    <h1>' + CONFIG.CLUB_NAME + '</h1>',
    '    <p>가입 신청 접수 확인</p>',
    '  </div>',
    '  <div class="bd">',
    '    <p class="greeting"><strong>' + (data.name || "회원") + '</strong>님, 안녕하세요!<br>',
    '    <strong>' + CONFIG.CLUB_NAME + '</strong>에 가입 신청해 주셔서 감사합니다.<br>',
    '    아래와 같이 신청 내용이 정상 접수되었습니다.</p>',
    '    <table>',
    '      <tr><td>이름</td><td>' + (data.name || "—") + '</td></tr>',
    '      <tr><td>이메일</td><td>' + (data.email || "—") + '</td></tr>',
    '      <tr><td>연락처</td><td>' + (data.phone || "—") + '</td></tr>',
    '      <tr><td>관심 분야</td><td>' + (interests || "—") + '</td></tr>',
    '      <tr><td>신청 일시</td><td>' + now + '</td></tr>',
    '    </table>',
    '    <div class="notice">',
    '      📋 <strong>운영진 검토 후 1~2일 내</strong> 승인 안내 메일을 보내드립니다.<br>',
    '      문의사항은 동호회 채널을 통해 ' + CONFIG.ADMIN_NAME + '에게 연락해 주세요.',
    '    </div>',
    '  </div>',
    '  <div class="ft">' + CONFIG.CLUB_NAME + ' &mdash; 렌즈와 AI로 창작의 경계를 함께 넓혀요</div>',
    '</div>',
    '</body></html>',
  ].join("\n");

  GmailApp.sendEmail(data.email, subject, "", { htmlBody: html, name: CONFIG.CLUB_NAME });
}

// ══════════════════════════════════════════════════════════════════════════════
//  sendAdminEmail — 관리자 알림 메일
// ══════════════════════════════════════════════════════════════════════════════
function sendAdminEmail(data, rowNum) {
  var interests = Array.isArray(data.interests) ? data.interests.join(", ") : (data.interests || "");
  var now = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
  var sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();

  var subject = "[신규 신청] " + (data.name || "이름 없음") + "님이 가입을 신청했습니다";

  var html = [
    '<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8">',
    '<style>',
    'body{font-family:"Noto Sans KR",sans-serif;background:#f8fafc;margin:0;padding:24px}',
    '.wrap{max-width:580px;margin:0 auto;background:#fff;border-radius:16px;overflow:hidden;box-shadow:0 2px 12px rgba(0,0,0,.1)}',
    '.hd{background:#1e293b;padding:24px 28px;color:#fff}',
    '.hd h1{font-size:18px;margin:0 0 4px;font-weight:700}',
    '.hd p{font-size:12px;color:#94a3b8;margin:0}',
    '.bd{padding:24px 28px}',
    '.row-badge{display:inline-block;background:#fef9c3;color:#854d0e;border-radius:6px;padding:3px 10px;font-size:12px;font-weight:700;margin-bottom:16px}',
    'table{width:100%;border-collapse:collapse;margin-bottom:20px}',
    'tr{border-bottom:1px solid #f1f5f9}',
    'td{padding:10px 8px;font-size:13px;vertical-align:top}',
    'td:first-child{color:#64748b;width:130px;font-weight:500;white-space:nowrap}',
    'td:last-child{color:#1e293b;font-weight:600;word-break:break-word}',
    '.intro-box{background:#f8fafc;border-radius:8px;padding:12px;font-size:13px;color:#374151;line-height:1.7;margin-bottom:20px}',
    '.btn-sheet{display:inline-block;background:#2563eb;color:#fff;text-decoration:none;border-radius:8px;padding:11px 22px;font-size:14px;font-weight:700;margin-bottom:8px}',
    '.ft{text-align:center;padding:16px;font-size:12px;color:#94a3b8;border-top:1px solid #f1f5f9}',
    '</style></head><body>',
    '<div class="wrap">',
    '  <div class="hd">',
    '    <h1>👥 신규 가입 신청 알림</h1>',
    '    <p>' + CONFIG.CLUB_NAME + ' &mdash; 관리자 알림</p>',
    '  </div>',
    '  <div class="bd">',
    '    <span class="row-badge">📋 스프레드시트 ' + rowNum + '행</span>',
    '    <table>',
    '      <tr><td>이름</td><td>' + (data.name  || "—") + '</td></tr>',
    '      <tr><td>이메일</td><td>' + (data.email || "—") + '</td></tr>',
    '      <tr><td>연락처</td><td>' + (data.phone || "—") + '</td></tr>',
    '      <tr><td>소속기관</td><td>' + (data.org   || "—") + '</td></tr>',
    '      <tr><td>직책·직급</td><td>' + (data.pos   || "—") + '</td></tr>',
    '      <tr><td>관심 분야</td><td>' + (interests  || "—") + '</td></tr>',
    '      <tr><td>사진·영상 수준</td><td>' + (data.skillPhoto || "—") + '</td></tr>',
    '      <tr><td>AI 도구 수준</td><td>' + (data.skillAi    || "—") + '</td></tr>',
    '      <tr><td>이용약관</td><td>' + (data.terms1 ? "✅ 동의" : "❌ 미동의") + '</td></tr>',
    '      <tr><td>개인정보 동의</td><td>' + (data.terms2 ? "✅ 동의" : "❌ 미동의") + '</td></tr>',
    '      <tr><td>창작물 공유</td><td>' + (data.terms3 ? "✅ 동의" : "—") + '</td></tr>',
    '      <tr><td>알림 수신</td><td>' + (data.terms4 ? "✅ 동의" : "—") + '</td></tr>',
    '      <tr><td>신청 일시</td><td>' + now + '</td></tr>',
    '    </table>',
    (data.intro
      ? '    <p style="font-size:13px;font-weight:600;color:#374151;margin-bottom:8px">📝 가입 동기</p>\n    <div class="intro-box">' + data.intro + '</div>'
      : ''),
    '    <p style="text-align:center">',
    '      <a class="btn-sheet" href="' + sheetUrl + '" target="_blank">📊 스프레드시트 바로가기</a>',
    '    </p>',
    '  </div>',
    '  <div class="ft">' + CONFIG.CLUB_NAME + ' &mdash; 자동 발송 메일입니다.</div>',
    '</div>',
    '</body></html>',
  ].join("\n");

  GmailApp.sendEmail(CONFIG.ADMIN_EMAIL, subject, "", { htmlBody: html, name: CONFIG.CLUB_NAME });
}

// ══════════════════════════════════════════════════════════════════════════════
//  approveSelected — 선택 행 승인 처리
// ══════════════════════════════════════════════════════════════════════════════
function approveSelected() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) { SpreadsheetApp.getUi().alert("시트를 찾을 수 없습니다."); return; }

  var selection = sheet.getActiveRange();
  var startRow  = selection.getRow();
  var endRow    = selection.getLastRow();

  if (startRow <= 1) { SpreadsheetApp.getUi().alert("헤더 행은 처리할 수 없습니다."); return; }

  var count = 0;
  for (var r = startRow; r <= endRow; r++) {
    var statusCell = sheet.getRange(r, HEADERS.length);
    var nameCell   = sheet.getRange(r, 3); // 이름
    var emailCell  = sheet.getRange(r, 5); // 이메일

    statusCell.setValue("✅ 승인 완료")
              .setBackground("#dcfce7")
              .setFontColor("#166534")
              .setFontWeight("bold")
              .setHorizontalAlignment("center");

    // 승인 이메일 발송
    var name  = nameCell.getValue();
    var email = emailCell.getValue();
    if (email) sendApprovalEmail(name, email);
    count++;
  }

  SpreadsheetApp.getUi().alert(count + "건이 승인 처리되었습니다.");
}

// ══════════════════════════════════════════════════════════════════════════════
//  rejectSelected — 선택 행 반려 처리
// ══════════════════════════════════════════════════════════════════════════════
function rejectSelected() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) { SpreadsheetApp.getUi().alert("시트를 찾을 수 없습니다."); return; }

  var selection = sheet.getActiveRange();
  var startRow  = selection.getRow();
  var endRow    = selection.getLastRow();

  if (startRow <= 1) { SpreadsheetApp.getUi().alert("헤더 행은 처리할 수 없습니다."); return; }

  var count = 0;
  for (var r = startRow; r <= endRow; r++) {
    var statusCell = sheet.getRange(r, HEADERS.length);
    statusCell.setValue("❌ 반려")
              .setBackground("#fee2e2")
              .setFontColor("#991b1b")
              .setFontWeight("bold")
              .setHorizontalAlignment("center");
    count++;
  }

  SpreadsheetApp.getUi().alert(count + "건이 반려 처리되었습니다.");
}

// ══════════════════════════════════════════════════════════════════════════════
//  sendApprovalEmail — 승인 안내 메일
// ══════════════════════════════════════════════════════════════════════════════
function sendApprovalEmail(name, email) {
  var subject = "[" + CONFIG.CLUB_NAME + "] 가입이 승인되었습니다 🎉";

  var html = [
    '<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8">',
    '<style>',
    'body{font-family:"Noto Sans KR",sans-serif;background:#f8fafc;margin:0;padding:24px}',
    '.wrap{max-width:560px;margin:0 auto;background:#fff;border-radius:16px;overflow:hidden;box-shadow:0 2px 12px rgba(0,0,0,.1)}',
    '.hd{background:linear-gradient(135deg,#1e293b 0%,#0f4c81 60%,#14532d 100%);padding:32px;text-align:center;color:#fff}',
    '.hd .emoji{font-size:40px;margin-bottom:12px}',
    '.hd h1{font-size:20px;margin:0 0 6px;font-weight:700}',
    '.bd{padding:28px;text-align:center}',
    '.badge{display:inline-block;background:#dcfce7;color:#166534;border-radius:20px;padding:6px 18px;font-size:14px;font-weight:700;margin-bottom:20px}',
    '.msg{font-size:15px;color:#1e293b;line-height:1.8;margin-bottom:24px;text-align:left}',
    '.welcome-box{background:#eff6ff;border:1.5px solid #bfdbfe;border-radius:10px;padding:16px;font-size:13px;color:#1d4ed8;line-height:1.7;text-align:left;margin-bottom:20px}',
    '.ft{text-align:center;padding:20px;font-size:12px;color:#94a3b8;border-top:1px solid #f1f5f9}',
    '</style></head><body>',
    '<div class="wrap">',
    '  <div class="hd">',
    '    <div class="emoji">🎉</div>',
    '    <h1>' + CONFIG.CLUB_NAME + '</h1>',
    '  </div>',
    '  <div class="bd">',
    '    <span class="badge">✅ 가입 승인 완료</span>',
    '    <p class="msg">',
    '      <strong>' + (name || "회원") + '</strong>님, 환영합니다!<br><br>',
    '      <strong>' + CONFIG.CLUB_NAME + '</strong> 가입이 최종 승인되었습니다. 🎊<br>',
    '      이제 동호회의 공식 회원으로 모든 활동에 참여하실 수 있습니다.',
    '    </p>',
    '    <div class="welcome-box">',
    '      📷 <strong>사진 촬영</strong> — 풍경·인물·스트릿 출사 활동<br>',
    '      🎬 <strong>영상 제작</strong> — 촬영·편집·유튜브 워크숍<br>',
    '      ✨ <strong>생성형 AI</strong> — Midjourney, FLUX, Sora 실습<br>',
    '      🗺️ <strong>출사·여행</strong> — 정기 그룹 출사 기획 참여<br><br>',
    '      운영진이 곧 상세 안내를 드릴 예정입니다. 잘 부탁드립니다!',
    '    </div>',
    '  </div>',
    '  <div class="ft">' + CONFIG.CLUB_NAME + ' &mdash; 렌즈와 AI로 창작의 경계를 함께 넓혀요</div>',
    '</div>',
    '</body></html>',
  ].join("\n");

  GmailApp.sendEmail(email, subject, "", { htmlBody: html, name: CONFIG.CLUB_NAME });
}

// ══════════════════════════════════════════════════════════════════════════════
//  showStats — 통계 팝업
// ══════════════════════════════════════════════════════════════════════════════
function showStats() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) { SpreadsheetApp.getUi().alert("시트를 찾을 수 없습니다."); return; }

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    SpreadsheetApp.getUi().alert("📊 통계\n\n신청 데이터가 없습니다.");
    return;
  }

  var statusCol = HEADERS.length;
  var statuses  = sheet.getRange(2, statusCol, lastRow - 1, 1).getValues();

  var total    = statuses.length;
  var approved = 0;
  var pending  = 0;
  var rejected = 0;
  var other    = 0;

  statuses.forEach(function(row) {
    var v = String(row[0]);
    if (v.indexOf("승인") !== -1)  approved++;
    else if (v.indexOf("신청") !== -1) pending++;
    else if (v.indexOf("반려") !== -1) rejected++;
    else other++;
  });

  var msg = [
    "📊 회원 신청 현황 통계",
    "─────────────────────",
    "전체 신청  : " + total    + "명",
    "✅ 승인 완료 : " + approved + "명",
    "⏳ 검토 대기 : " + pending  + "명",
    "❌ 반려     : " + rejected + "명",
    "기타       : " + other    + "명",
    "─────────────────────",
    "승인율 : " + (total > 0 ? Math.round(approved / total * 100) : 0) + "%",
  ].join("\n");

  SpreadsheetApp.getUi().alert(msg);
}

// ══════════════════════════════════════════════════════════════════════════════
//  onOpen — 스프레드시트 열 때 커스텀 메뉴 등록
// ══════════════════════════════════════════════════════════════════════════════
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("👥 회원 관리")
    .addItem("✅ 선택 행 승인",   "approveSelected")
    .addItem("❌ 선택 행 반려",   "rejectSelected")
    .addSeparator()
    .addItem("📊 통계 보기",      "showStats")
    .addToUi();
}

// ══════════════════════════════════════════════════════════════════════════════
//  jsonResponse — JSON ContentService 반환 헬퍼
// ══════════════════════════════════════════════════════════════════════════════
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
