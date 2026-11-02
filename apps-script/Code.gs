// 상하이 여행 페이지 일정 댓글 저장소
// 구글 시트 > 확장 프로그램 > Apps Script 에 붙여 넣고 웹 앱으로 배포한다
// (실행: 나, 액세스: 모든 사용자)

const SHEET_NAME = '댓글';
const DAYS = ['1', '2', '3'];

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
}

function json(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// = + - @ 로 시작하면 시트가 수식으로 해석하므로 작은따옴표를 붙여 글자로 저장한다
function asText(value) {
  return /^[=+\-@]/.test(value) ? "'" + value : value;
}

function doGet() {
  const rows = getSheet().getDataRange().getValues();
  const comments = rows
    .filter(row => DAYS.includes(String(row[1])))
    .map(row => ({ time: row[0], day: String(row[1]), name: String(row[2]), text: String(row[3]) }));
  return json(comments);
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const day = String(body.day);
  const name = String(body.name || '').trim().slice(0, 20);
  const text = String(body.text || '').trim().slice(0, 500);
  if (!DAYS.includes(day) || !name || !text) {
    return json({ ok: false });
  }
  getSheet().appendRow([new Date(), day, asText(name), asText(text)]);
  return json({ ok: true });
}
