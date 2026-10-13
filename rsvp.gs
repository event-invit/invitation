/**
 * 모바일 초대장 참석 회신 → 구글 시트 저장
 *
 * [설정 방법]
 * 1. 구글 시트를 새로 만들고, 메뉴 [확장 프로그램] → [Apps Script] 를 엽니다.
 * 2. 기존 코드를 지우고 이 파일 내용을 그대로 붙여넣고 저장합니다.
 * 3. 오른쪽 위 [배포] → [새 배포] → 유형 선택(톱니바퀴) → [웹 앱]
 *      - 다음 사용자 인증 정보로 실행: 나
 *      - 액세스 권한이 있는 사용자: 모든 사용자
 *    → [배포] → 권한 승인 → 나오는 "웹 앱 URL"(…/exec)을 복사합니다.
 * 4. index.html 의 CONFIG.rsvpUrl 에 그 주소를 붙여넣습니다.
 *
 * ※ 코드를 고친 뒤에는 [배포] → [배포 관리] → 연필 아이콘 → 버전 "새 버전" 으로 다시 배포해야 반영됩니다.
 */
const SHEET_NAME = '참석회신';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) {
      sh = ss.insertSheet(SHEET_NAME);
      sh.appendRow(['접수일시', '참석여부', '성함', '소속', '연락처', '인원', '남기신 말씀']);
      sh.setFrozenRows(1);
      sh.getRange('1:1').setFontWeight('bold').setBackground('#f0e2d8');
    }
    const p = e.parameter;
    const clean = v => String(v || '').replace(/^[=+\-@]/, "'$&").slice(0, 500); // 수식 주입 방지
    sh.appendRow([
      new Date(),
      clean(p.attend),
      clean(p.name),
      clean(p.org),
      clean(p.tel),
      Number(p.count) || 0,
      clean(p.message),
    ]);
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

/** 시트 맨 아래에 참석 인원 합계를 보고 싶다면: =SUMIF(B:B,"참석",F:F) */
