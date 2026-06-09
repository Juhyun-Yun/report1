/**
 * 과학탐구 활동 보고서 - Google Apps Script Backend
 * ---------------------------------------------------------------------------
 * 이 스크립트는 "스프레드시트에 연결된(컨테이너 바운드)" 스크립트입니다.
 * 선생님은 이 스프레드시트 "사본"을 만든 뒤, 시트 상단 메뉴
 *   [📋 탐구활동 보고서] → [① 처음 설정]
 * 을 누르면 '선생님 가이드 / 학생명단 / 활동기록' 시트가 자동으로 만들어지고,
 * 가이드 시트의 안내대로 웹앱을 배포해 앱과 연결합니다.
 *
 * ※ 보안: 특정 선생님의 웹앱 URL은 앱(공개 GitHub) 코드에 넣지 않습니다.
 *   각 선생님이 자기 사본을 배포하고, 앱의 "선생님 설정"에 URL을 넣거나
 *   학생용 링크(앱주소?api=웹앱URL)로 학생을 연결합니다.
 * ---------------------------------------------------------------------------
 */

// ▼▼▼ (배포 담당자만) 이 마스터 시트를 나눠주기 전에 한 번만 채워주세요. ▼▼▼
// GitHub 등에 올린 앱(index.html)의 공개 주소.
// 예) 'https://your-id.github.io/science-report/'
// 비워 두면 "학생용 링크 만들기"가 웹앱 URL만 안내합니다(앱 주소는 선생님이 직접 붙임).
const APP_URL = '';
// ▲▲▲ ----------------------------------------------------------------- ▲▲▲

const ROSTER_SHEET = '학생명단';
const RECORD_SHEET = '활동기록';
const GUIDE_SHEET  = '선생님 가이드';
const DRIVE_FOLDER_NAME = '탐구활동_보고서';

const RECORD_HEADERS = [
  '타임스탬프', '번호', '이름', '팀', '모드',
  '활동유형', '제목', '내용', '보고서 PDF'
];

// ===========================================================================
//  메뉴
// ===========================================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📋 탐구활동 보고서')
    .addItem('① 처음 설정 (가이드·명단·기록 만들기)', 'setupWorkbook')
    .addItem('② 학생용 링크 만들기', 'showStudentLink')
    .addSeparator()
    .addItem('선생님 가이드 다시 보기', 'openGuide')
    .addToUi();
}

// ===========================================================================
//  처음 설정: 필요한 시트 생성 + 가이드 작성
// ===========================================================================
function setupWorkbook() {
  const ss = SpreadsheetApp.getActive();
  createGuideSheet_(ss);
  ensureRosterSheet_(ss);
  getOrCreateRecordSheet_();

  // 가이드 시트를 맨 앞으로, 활성화
  const guide = ss.getSheetByName(GUIDE_SHEET);
  ss.setActiveSheet(guide);
  ss.moveActiveSheet(1);

  SpreadsheetApp.getUi().alert(
    '설정 완료! ✅',
    '「선생님 가이드」 시트를 먼저 읽어 주세요.\n' +
    '그 다음 「학생명단」에 학생을 입력하고, 안내대로 웹앱을 배포하면 됩니다.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function openGuide() {
  const ss = SpreadsheetApp.getActive();
  let guide = ss.getSheetByName(GUIDE_SHEET);
  if (!guide) { createGuideSheet_(ss); guide = ss.getSheetByName(GUIDE_SHEET); }
  ss.setActiveSheet(guide);
}

// 학생명단 시트가 없으면 헤더 + 예시와 함께 생성
function ensureRosterSheet_(ss) {
  let sheet = ss.getSheetByName(ROSTER_SHEET);
  if (sheet) return sheet;

  sheet = ss.insertSheet(ROSTER_SHEET);
  sheet.getRange(1, 1, 1, 3).setValues([['번호', '이름', '팀']]);
  sheet.getRange(1, 1, 1, 3)
       .setFontWeight('bold')
       .setBackground('#e8eaf6');
  sheet.setFrozenRows(1);
  // 예시 데이터 (선생님이 지우고 우리 반으로 교체)
  sheet.getRange(2, 1, 3, 3).setValues([
    ['1', '김과학', '1모둠'],
    ['2', '이탐구', '1모둠'],
    ['3', '박관찰', '']  // 팀이 비면 개인 활동
  ]);
  sheet.getRange(5, 1, 1, 3).setValues([['', '※ 위 예시는 지우고 우리 반 학생을 입력하세요. (이름만 필수)', '']]);
  sheet.setColumnWidth(2, 140);
  sheet.setColumnWidth(3, 140);
  return sheet;
}

// ===========================================================================
//  선생님 가이드 시트 작성
// ===========================================================================
function createGuideSheet_(ss) {
  let sheet = ss.getSheetByName(GUIDE_SHEET);
  if (sheet) ss.deleteSheet(sheet); // 항상 최신 내용으로 다시 작성
  sheet = ss.insertSheet(GUIDE_SHEET, 0);

  sheet.setColumnWidth(1, 880);
  sheet.setHiddenGridlines(true);

  // [텍스트, 종류] — 종류에 따라 서식 적용
  const rows = [
    ['📋 과학탐구활동 보고서 — 선생님 가이드', 'title'],
    ['이 스프레드시트는 우리 반 학생들의 탐구활동 보고서를 모으는 곳이에요. 아래 순서대로 한 번만 설정하면 됩니다.', 'lead'],
    ['', 'gap'],

    ['STEP 1.  처음 설정 (이미 보고 있다면 완료!)', 'h'],
    ['• 상단 메뉴 [📋 탐구활동 보고서] → [① 처음 설정]을 누르면 「학생명단」과 「활동기록」 시트가 자동으로 만들어져요.', 'body'],
    ['', 'gap'],

    ['STEP 2.  학생 명단 입력', 'h'],
    ['• 「학생명단」 시트에 우리 반 학생을 입력하세요.', 'body'],
    ['• "이름" 칸만 필수예요. "번호"와 "팀"은 선택입니다.', 'body'],
    ['• 모둠(팀) 활동이면 같은 팀끼리 "팀" 칸에 같은 이름을 적어요. 비워 두면 개인 활동으로 시작합니다.', 'body'],
    ['', 'gap'],

    ['STEP 3.  앱과 연결하기 (웹앱 배포)', 'h'],
    ['① 상단 메뉴 [확장 프로그램] → [Apps Script] 를 엽니다.', 'body'],
    ['② Apps Script 화면 오른쪽 위 [배포] → [새 배포] 를 누릅니다.', 'body'],
    ['③ 톱니바퀴(유형 선택) → [웹 앱] 을 고릅니다.', 'body'],
    ['④ "다음 사용자로 실행": 나   /   "액세스 권한": 모든 사용자  ← 학생도 접속하므로 꼭 이렇게!', 'warn'],
    ['⑤ [배포] 를 누르고 권한을 승인하면 "웹 앱 URL"(끝이 /exec)이 나옵니다. 복사해 두세요.', 'body'],
    ['', 'gap'],

    ['STEP 4.  학생들에게 나눠주기', 'h'],
    ['• 상단 메뉴 [📋 탐구활동 보고서] → [② 학생용 링크 만들기] 를 누르면 학생용 링크가 나와요.', 'body'],
    ['• 그 링크를 복사해 클래스룸/메신저로 학생들에게 전달하세요. 학생이 링크로 들어오면 자동으로 우리 반 시트에 연결됩니다.', 'body'],
    ['• (또는 학생이 앱 첫 화면의 "선생님 설정"에 위 웹앱 URL을 직접 붙여넣어도 됩니다.)', 'body'],
    ['', 'gap'],

    ['STEP 5.  학생 보고서 확인', 'h'],
    ['• 학생이 앱에서 [제출]을 누르면 「활동기록」 시트에 한 줄씩 쌓여요.', 'body'],
    ['• "보고서 PDF" 칸의 링크로 제출한 보고서를 볼 수 있어요.', 'body'],
    ['• 사진·그래프 등은 내 Google Drive의 「탐구활동_보고서」 폴더에 저장됩니다.', 'body'],
    ['', 'gap'],

    ['❗ 꼭 기억하세요', 'h'],
    ['• 웹앱은 반드시 "모든 사용자"로 배포해야 학생이 로그인 없이 접속할 수 있어요.', 'body'],
    ['• 코드(Apps Script)를 수정했다면 [배포] → [배포 관리] → 연필 아이콘 → 버전 "새 버전"으로 다시 배포하세요.', 'body'],
    ['• 이 시트와 링크는 "우리 반" 전용이에요. 다른 선생님은 각자 이 스프레드시트 "사본"을 만들어 같은 순서로 설정하면 됩니다.', 'body'],
    ['• 학생의 개인정보가 담기므로, 시트와 학생용 링크는 우리 반에게만 공유해 주세요.', 'warn'],
  ];

  const values = rows.map(function (r) { return [r[0]]; });
  sheet.getRange(1, 1, values.length, 1).setValues(values);

  // 서식 적용
  for (let i = 0; i < rows.length; i++) {
    const type = rows[i][1];
    const cell = sheet.getRange(i + 1, 1);
    cell.setVerticalAlignment('middle').setWrap(true);
    if (type === 'title') {
      cell.setFontSize(18).setFontWeight('bold').setFontColor('#ffffff').setBackground('#3f51b5');
      sheet.setRowHeight(i + 1, 48);
    } else if (type === 'lead') {
      cell.setFontSize(11).setFontColor('#374151').setBackground('#eef2ff');
      sheet.setRowHeight(i + 1, 40);
    } else if (type === 'h') {
      cell.setFontSize(13).setFontWeight('bold').setFontColor('#1e293b').setBackground('#e0e7ff');
      sheet.setRowHeight(i + 1, 34);
    } else if (type === 'warn') {
      cell.setFontSize(11).setFontWeight('bold').setFontColor('#b91c1c').setBackground('#fef2f2');
      sheet.setRowHeight(i + 1, 30);
    } else if (type === 'gap') {
      sheet.setRowHeight(i + 1, 10);
    } else {
      cell.setFontSize(11).setFontColor('#334155');
      sheet.setRowHeight(i + 1, 26);
    }
  }
  return sheet;
}

// ===========================================================================
//  학생용 링크 만들기 (배포 후)
// ===========================================================================
function showStudentLink() {
  const ui = SpreadsheetApp.getUi();
  let webAppUrl = '';
  try { webAppUrl = ScriptApp.getService().getUrl() || ''; } catch (e) {}

  if (!webAppUrl) {
    ui.alert(
      '아직 웹앱이 배포되지 않았어요',
      '먼저 [확장 프로그램] → [Apps Script] → [배포] → [새 배포] → [웹 앱]으로 배포한 뒤\n다시 이 메뉴를 눌러 주세요. (가이드 STEP 3 참고)',
      ui.ButtonSet.OK
    );
    return;
  }

  const hasApp = APP_URL && APP_URL.indexOf('http') === 0;
  const base = hasApp ? APP_URL.replace(/[?#].*$/, '') : '';
  const studentLink = hasApp ? (base + '?api=' + encodeURIComponent(webAppUrl)) : '';

  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:Pretendard,Apple SD Gothic Neo,Malgun Gothic,sans-serif;padding:6px 4px;color:#1e293b;">' +
      '<p style="margin:0 0 6px;font-weight:700;">① 내 웹앱 URL</p>' +
      '<input style="width:100%;box-sizing:border-box;padding:8px;border:1px solid #cbd5e1;border-radius:8px;font-size:12px;" ' +
        'value="' + escapeHtml_(webAppUrl) + '" onclick="this.select()" readonly />' +
      (hasApp
        ? ('<p style="margin:14px 0 6px;font-weight:700;">② 학생용 링크 (이 링크를 학생에게 전달하세요)</p>' +
           '<input style="width:100%;box-sizing:border-box;padding:8px;border:1px solid #34d399;border-radius:8px;font-size:12px;background:#ecfdf5;" ' +
             'value="' + escapeHtml_(studentLink) + '" onclick="this.select()" readonly />' +
           '<p style="margin:10px 0 0;font-size:12px;color:#64748b;">학생이 이 링크로 들어오면 우리 반 시트에 자동 연결돼요.</p>')
        : ('<p style="margin:14px 0 0;font-size:12px;color:#64748b;">앱 주소가 설정되어 있지 않아요.<br>' +
           '학생에게 <b>앱 주소</b>를 알려주고, 앱 첫 화면의 <b>"선생님 설정"</b>에 위 웹앱 URL을 넣게 하거나,<br>' +
           '<b>앱주소?api=웹앱URL</b> 형태의 링크를 만들어 전달하세요.</p>')
      ) +
      '<p style="margin:14px 0 0;font-size:11px;color:#94a3b8;">칸을 클릭하면 전체 선택돼요. Ctrl+C(또는 ⌘+C)로 복사하세요.</p>' +
    '</div>'
  ).setWidth(460).setHeight(hasApp ? 300 : 280);

  ui.showModalDialog(html, '학생용 링크 만들기');
}

function escapeHtml_(s) {
  return String(s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

// ===========================================================================
//  웹앱 엔드포인트
// ===========================================================================
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
    return jsonResponse_({ ok: false, error: '"' + ROSTER_SHEET + '" 시트가 없습니다. 메뉴 [📋 탐구활동 보고서] → [① 처음 설정]을 먼저 실행하세요.' });
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
