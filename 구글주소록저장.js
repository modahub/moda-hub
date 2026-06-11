var COL = {
  status  : 0,
  source  : 5,
  orderNum: 6,
  name    : 7,
  product : 8,
  option  : 9,
  phone   : 13,
  address : 15
};

function saveSelectedContact() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row   = sheet.getActiveCell().getRow();

  if (row <= 1) {
    SpreadsheetApp.getUi().alert('데이터 행을 선택해주세요. (1행은 헤더입니다)');
    return;
  }

  var rows = sheet.getDataRange().getValues();
  var r    = rows[row - 1];

  var nameH  = String(r[COL.name]).trim();
  var status = String(r[COL.status]).trim();
  var phone  = String(r[COL.phone]).trim();

  if (!nameH) {
    SpreadsheetApp.getUi().alert('이름이 없는 행입니다.');
    return;
  }
  if (status && status !== '정상') {
    SpreadsheetApp.getUi().alert('정상 상태가 아닌 행입니다. (현재: ' + status + ')');
    return;
  }
  if (!phone) {
    SpreadsheetApp.getUi().alert('전화번호(N열)가 없는 행입니다.');
    return;
  }

  // 중복 체크
  try {
    var searchResult = People.People.searchContacts({
      query: phone.replace(/\D/g, ''),
      readMask: 'phoneNumbers,names'
    });
    if (searchResult.results && searchResult.results.length > 0) {
      SpreadsheetApp.getUi().alert('⏭ 이미 주소록에 존재합니다.\n이름: ' + nameH + '\n전화번호: ' + phone);
      return;
    }
  } catch(e) {}

  var sourceRaw = String(r[COL.source]).trim();
  var lastName  = '';
  var dashIdx   = sourceRaw.indexOf('-');
  if (dashIdx > 0) {
    lastName = sourceRaw.substring(0, dashIdx).trim();
  }

  var nameI     = String(r[COL.product]).trim();
  var firstName = nameH + '/' + nameI;
  var option    = String(r[COL.option]).trim();
  var source    = dashIdx > 0 ? sourceRaw.substring(dashIdx + 1).trim() : sourceRaw;
  var orderNum  = String(r[COL.orderNum]).trim();
  var address   = String(r[COL.address]).trim();

  try {
    var memo = '주문번호: ' + orderNum
             + (option  ? '\n옵션: '   + option  : '')
             + (source  ? '\n구입처: ' + source  : '');

    var contactBody = {
      names: [{
        givenName : firstName,
        familyName: lastName
      }],
      phoneNumbers: [{
        value: phone,
        type : 'mobile'
      }],
      biographies: [{
        value      : memo,
        contentType: 'TEXT_PLAIN'
      }]
    };

    if (address) {
      contactBody.addresses = [{
        streetAddress: address,
        type         : 'home'
      }];
    }

    People.People.createContact(contactBody);
    SpreadsheetApp.getUi().alert('✅ 저장 완료!\n이름: ' + lastName + ' ' + nameH + '\n전화번호: ' + phone);

  } catch(err) {
    SpreadsheetApp.getUi().alert('⚠️ 오류: ' + err.message);
  }
}