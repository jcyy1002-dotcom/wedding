/**
 * 모바일 청첩장 - 축하 메시지 저장소 (Google Sheets)
 * 구글 시트 → 확장 프로그램 → Apps Script 에 이 코드를 붙여넣고 배포하세요.
 * (자세한 단계는 채팅의 안내를 참고)
 */

function doGet(e)  { return handle(e); }
function doPost(e) { return handle(e); }

function handle(e) {
  var sheet = getSheet();
  var p = (e && e.parameter) || {};
  var action = p.action || 'list';
  var out;

  if (action === 'add') {
    var id = String(Date.now());
    sheet.appendRow([id, p.text || '', p.password || '', kstTime()]);
    out = { ok: true, id: id, time: kstTime() };

  } else if (action === 'delete') {
    var data = sheet.getDataRange().getValues();
    out = { ok: false, error: 'notfound' };
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(p.id)) {
        if (String(data[i][2]) === String(p.password)) {
          sheet.deleteRow(i + 1);
          out = { ok: true };
        } else {
          out = { ok: false, error: 'password' };
        }
        break;
      }
    }

  } else { // list
    var data = sheet.getDataRange().getValues();
    var list = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      list.push({ id: String(data[i][0]), text: data[i][1], time: data[i][3] });
    }
    out = { ok: true, messages: list };
  }

  return ContentService
    .createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('messages');
  if (!sheet) {
    sheet = ss.insertSheet('messages');
    sheet.appendRow(['id', 'text', 'password', 'time']);
  }
  return sheet;
}

function kstTime() {
  var d = new Date(Date.now() + 9 * 3600 * 1000); // 한국 시간(KST)
  var h = d.getUTCHours(), m = d.getUTCMinutes();
  var ampm = h < 12 ? '오전' : '오후';
  var hh = h % 12; if (hh === 0) hh = 12;
  return ampm + ' ' + hh + ':' + (m < 10 ? '0' + m : m);
}
