var COL_NAME    = 8;  // H열 = 수령자
var COL_UNIPASS = 12; // L열 = 통관번호
var COL_PHONE   = 13; // M열 = 휴대폰1
var COL_PHONE2  = 14; // N열 = 휴대폰2
var COL_RESULT  = 15; // O열 = 검증결과


function stopVerify() {
  PropertiesService.getScriptProperties().setProperty('STOP', 'true');
  SpreadsheetApp.getUi().alert('중지 요청됨! 현재 행 완료 후 멈춥니다.');
}

function resetStop() {
  PropertiesService.getScriptProperties().setProperty('STOP', 'false');
}

function isStopped() {
  return PropertiesService.getScriptProperties().getProperty('STOP') === 'true';
}

function verifySelected() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var row = sheet.getActiveCell().getRow();
  if (row <= 1) {
    SpreadsheetApp.getUi().alert('데이터 행을 선택해주세요');
    return;
  }
  verifyRow(sheet, row);
}

function verifyAll() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var ui = SpreadsheetApp.getUi();

  // 통관번호 열 전체를 한번에 읽기
  var unipassValues = sheet.getRange(2, COL_UNIPASS, lastRow - 1, 1).getValues();
  var targetRows = [];
  for (var r = 0; r < unipassValues.length; r++) {
    if (unipassValues[r][0].toString().trim()) {
      targetRows.push(r + 2);
    }
  }

  if (targetRows.length === 0) {
    ui.alert('통관번호가 있는 행이 없습니다.');
    return;
  }

  var confirm = ui.alert('일괄 검증', '통관번호가 있는 ' + targetRows.length + '건만 검증합니다. 계속?', ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  resetStop();

  for (var i = 0; i < targetRows.length; i++) {
    if (isStopped()) {
      ui.alert(i + '건까지 검증 완료 후 중지됐습니다.');
      resetStop();
      return;
    }
    verifyRow(sheet, targetRows[i]);
    Utilities.sleep(600);
  }
  ui.alert('검증 완료! 총 ' + targetRows.length + '건 처리됐습니다.');
}

function verifyRow(sheet, row) {
  var name    = sheet.getRange(row, COL_NAME).getValue().toString().trim();
  var unipass = sheet.getRange(row, COL_UNIPASS).getValue().toString().trim().toUpperCase();
  var phoneM  = sheet.getRange(row, COL_PHONE).getValue().toString().trim();
  var phoneN  = sheet.getRange(row, COL_PHONE2).getValue().toString().trim();

  // M열이 비어있거나 0504/0502 등 비표준 번호면 N열 우선 사용
  var phone = phoneM;
  if (!phone || /^0[^1]/.test(phone.replace(/-/g, ''))) {
    phone = phoneN;
  }

  if (!name || !unipass || !phone) {
    sheet.getRange(row, COL_RESULT).setValue('⚠️ 데이터없음');
    return;
  }

  if (!/^P[0-9]{12}$/.test(unipass)) {
    sheet.getRange(row, COL_RESULT).setValue('❌ 형식오류');
    return;
  }

  try {
    var url = 'https://sellochomes.co.kr/api/v1/sellerlife/unipass/unipass';

    // 1차 시도 (선택된 전화번호)
    var result = callApi(url, unipass, name, phone);

    // 1차 실패하고 M열과 N열이 다르면 나머지 번호로 재시도
    if (result !== '✅ 정상' && phoneM && phoneN && phoneM !== phoneN) {
      var phone2 = (phone === phoneM) ? phoneN : phoneM;
      // 재시도할 번호가 표준 010 번호일 때만
      if (/^010/.test(phone2.replace(/-/g, ''))) {
        var result2 = callApi(url, unipass, name, phone2);
        if (result2 === '✅ 정상') result = result2;
      }
    }

    sheet.getRange(row, COL_RESULT).setValue(result);

  } catch(e) {
    sheet.getRange(row, COL_RESULT).setValue('오류:' + e.message);
  }
}

function callApi(url, unipass, name, phone) {
  var cleanPhone = phone.replace(/-/g, '');
  var payload = 'persEcm=' + unipass +
                '&pltxNm=' + encodeURIComponent(name) +
                '&cralTelno=' + cleanPhone +
                '&addrNo=';
  var options = {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: payload,
    muteHttpExceptions: true,
    headers: {
      Cookie: getLoginCookie(),
      Referer: 'https://sellochomes.co.kr/sellerlife/unipass/'
    }
  };
  var response = UrlFetchApp.fetch(url, options);
  var text = response.getContentText();
  var json = JSON.parse(text);
  var data = json.data || '';

  var tCntMatch = data.match(/<tCnt>(\d+)<\/tCnt>/);
  var tCnt = tCntMatch ? parseInt(tCntMatch[1]) : 0;

  if (tCnt >= 1) return '✅ 정상';
  if (data.indexOf('존재하지 않습니다') > -1) return '❌ 번호없음';
  if (data.indexOf('불일치') > -1 || data.indexOf('일치하지 않습니다') > -1) return '❌ 불일치';
  return '⚠️ 확인필요';
}

function getLoginCookie() {
  return 'connect.sid=s%3Agu4c_Li3M-EbBV0g-b6bnWdTKDUkWM91.gn4BvUAq50TAFydPCKKaOhb1qIW8fWiaO1DJQ5xeCTw';
}

function testApi() {
  var url = 'https://sellochomes.co.kr/api/v1/sellerlife/unipass/unipass';
  var options = {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: 'persEcm=P260000491863&pltxNm=' + encodeURIComponent('정금동') + '&cralTelno=01087870127&addrNo=',
    muteHttpExceptions: true,
    headers: {
      Cookie: getLoginCookie(),
      Referer: 'https://sellochomes.co.kr/sellerlife/unipass/'
    }
  };
  var response = UrlFetchApp.fetch(url, options);
  Logger.log('상태코드:' + response.getResponseCode());
  Logger.log('응답전체:' + response.getContentText());
}
