/**
 * 2026 직원 건강관리 설문조사 - 응답 수집 백엔드
 *
 * 사용 방법
 * 1) 대상 구글 스프레드시트를 열고 [확장 프로그램] > [Apps Script] 클릭
 * 2) 기본 생성된 코드를 지우고 이 파일 내용 전체를 붙여넣기
 * 3) SPREADSHEET_ID 값이 비어있으면 바인딩된 스프레드시트를 자동으로 사용합니다.
 *    (스프레드시트에 직접 연결한 스크립트라면 비워둬도 됩니다)
 * 4) 상단 [배포] > [새 배포] > 유형: "웹 앱" 선택
 *    - 실행 사용자: 나
 *    - 액세스 권한이 있는 사용자: 전체 허용(Anyone)  ← 로그인 없이 깃허브 페이지에서 호출 가능해야 하므로 반드시 "Anyone"으로 설정
 * 5) 배포 후 나오는 웹 앱 URL을 복사해서, store/index.html (및 추후 본사용 index.html) 상단의
 *    SCRIPT_URL 값에 붙여넣기
 * 6) 매장용/본사용을 같은 시트로 모을지, 탭을 분리할지는 아래 SHEET_NAME_PREFIX 로 제어됩니다.
 *    (기본: 요청으로 들어온 site 값 그대로 탭 이름 사용, 예: "매장", "본사")
 */

// 비워두면 이 스크립트가 바인딩된 스프레드시트를 사용합니다.
const SPREADSHEET_ID = '1ODz3LIxyHucAq-4L_FfZCl9RKU0w__43E3EoYRFmnkA';

const HEADERS = [
  '제출시각',
  '구분(매장/본사)',
  'Q1_건강관리활동_도움정도(1-5)',
  'Q2_건강개선여부',
  'Q2_개선된항목',
  'Q3_회사지원인식(1-5)',
  'Q4_만족활동',
  'Q5_걷기프로그램_참여',
  'Q5_걷기프로그램_만족도(1-5)',
  'Q5_그림컨테스트_참여',
  'Q5_그림컨테스트_만족도(1-5)',
  'Q5_참여안함',
  'Q6_희망프로그램',
  'Q6_기타의견'
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getOrCreateSheet_(data.site);

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.setFrozenRows(1);
    } else {
      // 설문 문항이 수정되어 헤더 구성이 바뀐 경우, 기존 헤더 행을 최신 구성으로 갱신
      const currentHeaderRow = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
      const isSame = HEADERS.every((h, i) => currentHeaderRow[i] === h);
      if (!isSame) {
        sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      }
    }

    sheet.appendRow([
      new Date(),
      data.site || '',
      data.q1 || '',
      data.q2 || '',
      (data.q2_sub || []).join(', '),
      data.q3 || '',
      (data.q4 || []).join(', '),
      data.q5_walk ? 'Y' : 'N',
      data.q5_walk_sat || '',
      data.q5_draw ? 'Y' : 'N',
      data.q5_draw_sat || '',
      data.q5_none ? 'Y' : 'N',
      (data.q6 || []).join(', '),
      data.q6_other || ''
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: '설문 제출용 웹앱이 정상 동작 중입니다.' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet_(site) {
  const ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = site === '본사' ? '본사_응답' : '매장_응답';
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}
