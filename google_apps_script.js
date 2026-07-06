/**
 * 한남대학교 창업지원단 스타트업 모니터링 시스템 - 구글 스프레드시트 연동용 Apps Script
 * 
 * [설치 방법]
 * 1. 구글 스프레드시트(새 시트)를 하나 생성합니다.
 * 2. 상단 메뉴에서 [확장 프로그램] -> [Apps Script]를 클릭합니다.
 * 3. 기존에 있던 myFunction 코드를 지우고 이 파일의 전체 코드를 복사하여 붙여넣습니다.
 * 4. 상단 저장 버튼(디스크 아이콘)을 누릅니다.
 * 5. 우측 상단의 [배포] -> [새 배포]를 클릭합니다.
 * 6. 유형 선택(톱니바퀴)에서 [웹 앱]을 선택합니다.
 * 7. 아래와 같이 설정하고 [배포] 버튼을 누릅니다:
 *    - 설명: 스타트업 모니터링 DB
 *    - 웹앱을 실행할 사용자: 나 (대표자 이메일)
 *    - 액세스 권한이 있는 사용자: 모든 사용자 (가장 중요!)
 * 8. 승인 요청이 뜨면 [권한 검토] -> 본인 계정 선택 -> [고급] -> [제목 없는 프로젝트(이동)] -> [허용]을 진행합니다.
 * 9. 생성된 "웹 앱 URL"을 복사하여 changup/app.js 파일의 `GOOGLE_SCRIPT_URL` 변수에 넣으시면 됩니다.
 */

// 데이터가 저장될 시트 이름
const SHEET_NAME = "DB_STORE";

// GET 요청 처리 (데이터 불러오기)
function doGet(e) {
  const sheet = getOrCreateSheet();
  const jsonStr = sheet.getRange(1, 1).getValue() || "{}";
  
  return ContentService.createTextOutput(JSON.stringify({
    status: "success",
    data: JSON.parse(jsonStr)
  }))
  .setMimeType(ContentService.MimeType.JSON)
  .setHeader("Access-Control-Allow-Origin", "*");
}

// POST 요청 처리 (데이터 저장하기)
function doPost(e) {
  try {
    let postData;
    if (e.postData.type === "application/json") {
      postData = JSON.parse(e.postData.contents);
    } else {
      // CORS preflight 및 simple request 대응을 위해 text/plain으로 넘어온 경우 처리
      postData = JSON.parse(e.postData.contents);
    }
    
    if (postData.action === "syncData") {
      const sheet = getOrCreateSheet();
      
      // 저장할 데이터 오브젝트 구성
      const dbData = {
        USERS: postData.USERS,
        COMPANIES: postData.companies,
        MILESTONES: postData.milestones,
        coachName: postData.coachName,
        eduNames: postData.eduNames,
        notices: postData.notices,
        lastUpdated: new Date().toISOString(),
        updatedBy: postData.userEmail || "System"
      };
      
      // JSON 스트링으로 시트의 A1 셀에 저장 (50,000자 제한 내 안전 저장)
      sheet.getRange(1, 1).setValue(JSON.stringify(dbData));
      
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        message: "Data synchronized successfully"
      }))
      .setMimeType(ContentService.MimeType.JSON)
      .setHeader("Access-Control-Allow-Origin", "*");
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: "Unknown action"
    }))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeader("Access-Control-Allow-Origin", "*");
    
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    }))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeader("Access-Control-Allow-Origin", "*");
  }
}

// 시트가 없을 경우 자동 생성하는 헬퍼 함수
function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    // 첫 행 넓이와 높이를 늘려 데이터 가독성 확보
    sheet.setRowHeight(1, 100);
    sheet.setColumnWidth(1, 500);
    // 자동 줄바꿈 설정
    sheet.getRange(1, 1).setWrap(true);
  }
  return sheet;
}
