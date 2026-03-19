/**
 * 견적서 자동화 - Google Apps Script 백엔드
 */

const SHEET_NAME = '견적서';
const ITEM_START_ROW = 18;
const MAX_ITEMS = 10;

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('견적서')
    .addItem('견적서 작성', 'openSidebar')
    .addItem('PDF 다운로드', 'exportToPDF')
    .addToUi();
}

function openSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setWidth(800)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, '견적서 작성');
}

function writeQuoteData(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    throw new Error('"' + SHEET_NAME + '" 시트를 찾을 수 없습니다.');
  }

  sheet.getRange('C5').setValue(data.recipientCompany);
  sheet.getRange('C6').setValue(data.recipientName);
  sheet.getRange('C7').setValue(data.recipientPhone);
  sheet.getRange('C8').setValue(data.recipientEmail);
  sheet.getRange('C9').setValue(data.quoteDate);
  sheet.getRange('H9').setValue(data.validityPeriod || '견적일로 부터 1개월간 유효');

  clearItems(sheet);

  var items = data.items || [];
  var totalSupply = 0;
  var totalTax = 0;

  items.forEach(function(item, index) {
    var row = ITEM_START_ROW + index;
    var qty = Number(item.quantity) || 0;
    var price = Number(item.unitPrice) || 0;
    var supplyAmount = qty * price;
    var tax = Math.round(supplyAmount * 0.1);

    sheet.getRange(row, 2).setValue(item.name);
    sheet.getRange(row, 7).setValue(qty);
    sheet.getRange(row, 9).setValue(price);
    sheet.getRange(row, 12).setValue(supplyAmount);
    sheet.getRange(row, 16).setValue(tax);
    sheet.getRange(row, 19).setValue(item.note || '');

    totalSupply += supplyAmount;
    totalTax += tax;
  });

  var totalAmount = totalSupply + totalTax;

  sheet.getRange('P29').setValue(totalSupply);
  sheet.getRange('R29').setValue(totalTax);
  sheet.getRange('T29').setValue(totalAmount);
  sheet.getRange('E15').setValue('\\' + totalAmount.toLocaleString() + '원정');

  SpreadsheetApp.flush();
  return { success: true, totalAmount: totalAmount };
}

function clearItems(sheet) {
  for (var row = ITEM_START_ROW; row < ITEM_START_ROW + MAX_ITEMS; row++) {
    sheet.getRange(row, 2).setValue('');
    sheet.getRange(row, 7).setValue('');
    sheet.getRange(row, 9).setValue('');
    sheet.getRange(row, 12).setValue('');
    sheet.getRange(row, 16).setValue('');
    sheet.getRange(row, 19).setValue('');
  }
}

function exportToPDF() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  var sheetId = sheet.getSheetId();

  var company = sheet.getRange('C5').getValue() || '견적서';
  var today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMdd');
  var fileName = '견적서_' + company + '_' + today + '.pdf';

  var url = ss.getUrl().replace(/\/edit.*$/, '')
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

  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + token }
  });

  var blob = response.getBlob().setName(fileName);
  var file = DriveApp.createFile(blob);

  return {
    success: true,
    fileName: fileName,
    fileUrl: file.getUrl(),
    fileId: file.getId()
  };
}

function generatePDF() {
  return exportToPDF();
}

// 시트 구조 스캔 (디버깅용)
function scanSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('견적서');
  var data = sheet.getDataRange().getValues();
  var result = '';
  for (var r = 0; r < data.length; r++) {
    for (var c = 0; c < data[r].length; c++) {
      var val = data[r][c];
      if (val !== '') {
        result += 'R' + (r+1) + 'C' + (c+1) + ' = ' + val + '\n';
      }
    }
  }
  SpreadsheetApp.getUi().alert(result);
}
