/**
 * 견적서 자동화 - Google Apps Script 백엔드
 */

const SHEET_NAME = '견적서';
const ITEM_START_ROW = 18;
const MAX_ITEMS = 10;

// 시트 메뉴 등록
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('견적서')
    .addItem('견적서 작성', 'openSidebar')
    .addItem('PDF 다운로드', 'exportToPDF')
    .addToUi();
}

// 사이드바 열기
function openSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('견적서 작성')
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

// 시트에 견적 데이터 쓰기
function writeQuoteData(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    throw new Error(`"${SHEET_NAME}" 시트를 찾을 수 없습니다.`);
  }

  // 수신자 정보
  sheet.getRange('C5').setValue(data.recipientCompany);
  sheet.getRange('C6').setValue(data.recipientName);
  sheet.getRange('C7').setValue(data.recipientPhone);
  sheet.getRange('C8').setValue(data.recipientEmail);

  // 견적일자, 유효기간
  sheet.getRange('C9').setValue(data.quoteDate);
  sheet.getRange('H9').setValue(data.validityPeriod || '견적일로 부터 1개월간 유효');

  // 기존 품목 행 초기화
  clearItems(sheet);

  // 품목 입력
  const items = data.items || [];
  let totalSupply = 0;
  let totalTax = 0;

  items.forEach(function(item, index) {
    const row = ITEM_START_ROW + index;
    const qty = Number(item.quantity) || 0;
    const price = Number(item.unitPrice) || 0;
    const supplyAmount = qty * price;
    const tax = Math.round(supplyAmount * 0.1);

    sheet.getRange(row, 2).setValue(item.name);       // B열: 품명
    sheet.getRange(row, 7).setValue(qty);              // G열: 수량
    sheet.getRange(row, 9).setValue(price);            // I열: 단가
    sheet.getRange(row, 12).setValue(supplyAmount);    // L열: 공급가액
    sheet.getRange(row, 16).setValue(tax);             // P열: 세액
    sheet.getRange(row, 19).setValue(item.note || ''); // S열: 비고

    totalSupply += supplyAmount;
    totalTax += tax;
  });

  const totalAmount = totalSupply + totalTax;

  // 하단 합계
  sheet.getRange('P29').setValue(totalSupply);
  sheet.getRange('R29').setValue(totalTax);
  sheet.getRange('T29').setValue(totalAmount);

  // 상단 합계 금액 표시
  sheet.getRange('E15').setValue('\\' + totalAmount.toLocaleString() + '원정');

  SpreadsheetApp.flush();
  return { success: true, totalAmount: totalAmount };
}

// 품목 행 초기화
function clearItems(sheet) {
  for (let row = ITEM_START_ROW; row < ITEM_START_ROW + MAX_ITEMS; row++) {
    sheet.getRange(row, 2).setValue('');  // 품명
    sheet.getRange(row, 7).setValue('');  // 수량
    sheet.getRange(row, 9).setValue('');  // 단가
    sheet.getRange(row, 12).setValue(''); // 공급가액
    sheet.getRange(row, 16).setValue(''); // 세액
    sheet.getRange(row, 19).setValue(''); // 비고
  }
}

// PDF 변환 및 다운로드
function exportToPDF() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  const sheetId = sheet.getSheetId();

  // 수신자 업체명으로 파일명 생성
  const company = sheet.getRange('C5').getValue() || '견적서';
  const today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMdd');
  const fileName = '견적서_' + company + '_' + today + '.pdf';

  // PDF 변환 URL 생성
  const url = ss.getUrl().replace(/\/edit.*$/, '')
    + '/export?exportFormat=pdf'
    + '&format=pdf'
    + '&size=A4'
    + '&portrait=true'
    + '&fitw=true'
    + '&gridlines=false'
    + '&printtitle=false'
    + '&sheetnames=false'
    + '&pagenum=UNDEFINED'
    + '&fzr=false'
    + '&gid=' + sheetId;

  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + token }
  });

  // Google Drive에 PDF 저장
  const blob = response.getBlob().setName(fileName);
  const file = DriveApp.createFile(blob);

  return {
    success: true,
    fileName: fileName,
    fileUrl: file.getUrl(),
    fileId: file.getId()
  };
}

// 사이드바에서 PDF 생성 호출용
function generatePDF() {
  return exportToPDF();
}
