
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🟠 통관번호 검증')
    .addItem('선택한 행 검증', 'verifySelected')
    .addItem('전체 일괄 검증', 'verifyAll')
    .addItem('🟥 검증 중지', 'stopVerify')
    .addItem('문자발송 시트로 데이터 가져오기', 'copyToSmsSheet')
    .addToUi();

  SpreadsheetApp.getUi()
  .createMenu('🔵 주소록저장')
  .addItem('선택한 행 저장', 'saveSelectedContact')
  .addToUi();
}
