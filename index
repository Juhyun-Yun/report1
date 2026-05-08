/**
 * 과학탐구 활동 보고서 - Google Apps Script Backend
 *
 * [설정 순서]
 * 1) Google Sheets 새 스프레드시트 생성.
 * 2) 시트 두 개 만들기:
 *    - "학생명단" : 1행 헤더 = 이름  (또는  번호 | 이름  /  이름 | 팀  /  번호 | 이름 | 팀)
 *      · "이름" 칼럼만 필수. "번호"·"팀"은 있어도 되고 없어도 됩니다.
 *      · "팀" 값이 비어 있는 학생은 개인 활동 모드로 시작합니다.
 *    - "활동기록" : 자동 생성됩니다(직접 만들 필요 없음).
 * 3) 메뉴 [확장 프로그램] → [Apps Script] 열기.
 * 4) 이 파일의 내용을 전부 복사해서 Apps Script Code.gs에 붙여넣기.
 * 5) 우상단 [배포] → [새 배포] → 유형 [웹 앱] 선택.
 *    - 설명: 아무거나
 *    - 다음 사용자로 실행: "나"
 *    - 액세스 권한: "모든 사용자"  (학생도 접근해야 하므로 필수)
 * 6) 배포 후 나오는 "웹 앱 URL"을 복사.
 * 7) 탐구활동.html 파일 상단의 SHEET_API_URL 값에 그 URL을 붙여넣기.
 *
 * 코드 수정 후 다시 배포할 때는 [배포] → [배포 관리] → 연필 아이콘 → 버전 "새 버전".
 */

const ROSTER_SHEET = '학생명단';
const RECORD_SHEET = '활동기록';
const DRIVE_FOLDER_NAME = '탐구활동_보고서';

const RECORD_HEADERS = [
  '타임스탬프', '번호', '이름', '팀', '모드',
  '활동유형', '제목', '내용', '보고서 PDF'
];

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'getRoster';
    if (action === 'getRoster') return getRoster_();
    return jsonResponse_({ ok: false, error: 'unknown action: ' + action });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    const action = body.action;
    if (action === 'uploadImage') return uploadImage_(body);
    if (action === 'saveReport') return saveReport_(body);
    return jsonResponse_({ ok: false, error: 'unknown action: ' + action });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  }
}

function getRoster_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(ROSTER_SHEET);
  if (!sheet) {
    return jsonResponse_({ ok: false, error: '"' + ROSTER_SHEET + '" 시트가 없습니다.' });
  }
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return jsonResponse_({ ok: true, students: [] });

  const headers = values[0].map(function (h) { return String(h).trim(); });
  let idIdx = headers.indexOf('번호');
  if (idIdx < 0) idIdx = headers.indexOf('학번'); // 이전 명칭도 호환
  const nameIdx = headers.indexOf('이름');
  const teamIdx = headers.indexOf('팀');
  if (nameIdx < 0) {
    return jsonResponse_({ ok: false, error: '학생명단 1행에 "이름" 칼럼이 필요합니다.' });
  }

  const students = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const name = String(row[nameIdx] || '').trim();
    if (!name) continue;
    students.push({
      id: idIdx >= 0 ? String(row[idIdx] || '').trim() : '',
      name: name,
      team: teamIdx >= 0 ? String(row[teamIdx] || '').trim() : ''
    });
  }
  return jsonResponse_({ ok: true, students: students });
}

function uploadImage_(body) {
  if (!body.dataBase64) {
    return jsonResponse_({ ok: false, error: 'dataBase64 누락' });
  }
  const folder = getOrCreateFolder_(DRIVE_FOLDER_NAME);
  const decoded = Utilities.base64Decode(body.dataBase64);
  const mime = body.mimeType || 'image/png';
  const ext = mime === 'application/pdf' ? '.pdf'
            : mime === 'image/jpeg' ? '.jpg'
            : mime === 'image/png' ? '.png'
            : '';
  const filename = (body.filename || ('file_' + Date.now())) + ext;
  const blob = Utilities.newBlob(decoded, mime, filename);
  const file = folder.createFile(blob);
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (err) {
    // 도메인 정책으로 외부 공유가 막혀 있으면 무시 — 선생님은 폴더 안에서 직접 보면 됨
  }
  return jsonResponse_({ ok: true, url: file.getUrl(), id: file.getId() });
}

function saveReport_(body) {
  const sheet = getOrCreateRecordSheet_();
  // 새 형식: reportUrl 단일 값 / 옛 형식: imageUrls 배열도 호환
  let reportLink = '';
  if (body.reportUrl) {
    reportLink = String(body.reportUrl);
  } else if (Array.isArray(body.imageUrls) && body.imageUrls.length) {
    reportLink = body.imageUrls.join('\n');
  }
  sheet.appendRow([
    new Date(),
    body.studentId || '',
    body.studentName || '',
    body.team || '',
    body.mode || '',
    body.activityType || '',
    body.title || '',
    body.content || '',
    reportLink
  ]);
  return jsonResponse_({ ok: true });
}

function getOrCreateRecordSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(RECORD_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(RECORD_SHEET);
    sheet.appendRow(RECORD_HEADERS);
    sheet.getRange(1, 1, 1, RECORD_HEADERS.length)
         .setFontWeight('bold')
         .setBackground('#e8eaf6');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(8, 400);
    sheet.setColumnWidth(9, 300);
  }
  return sheet;
}

function getOrCreateFolder_(name) {
  const folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(name);
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
