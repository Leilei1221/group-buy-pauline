/**
 * 波琳小姐火烤蝴蝶酥 — 幾隻貓補貨團
 * GAS 訂單接收腳本
 */

const PRODUCT_LIST = [
  { id: 'p01', name: '綜合蝴蝶酥(大)', price: 504 },
  { id: 'p02', name: '原味蝴蝶酥(大)', price: 504 },
  { id: 'p03', name: '雙拼 原味+芝麻(大)', price: 504 },
  { id: 'p04', name: '雙拼 原味+椰子(大)', price: 504 },
  { id: 'p05', name: '原味蝴蝶酥(小)', price: 324 },
  { id: 'p06', name: '雙拼 原味+芝麻(小)', price: 324 },
  { id: 'p07', name: '芝麻蝴蝶酥(小)', price: 324 },
  { id: 'p08', name: '紙袋', price: 18 },
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const params = e.parameter;
    const name = params.name || '';
    const summary = params.summary || '';
    const total = parseInt(params.total || '0', 10);
    const note = params.note || '';

    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheets()[0];

    // Check if headers need setup
    if (sheet.getRange('C1').getValue() === '') {
      setupHeaders(sheet);
    }

    const row = sheet.getLastRow() + 1;
    const timestamp = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');

    // Build row data
    const rowData = [timestamp, name];
    PRODUCT_LIST.forEach(p => {
      const q = parseInt(params[p.id] || '0', 10);
      rowData.push(q > 0 ? q : '');
    });

    // Build reply message
    const itemLines = [];
    PRODUCT_LIST.forEach(p => {
      const q = parseInt(params[p.id] || '0', 10);
      if (q > 0) itemLines.push(`${p.name} ${q}盒 單價${p.price}元`);
    });
    const replyMsg = `${name}你好，你所訂購的產品如下：${itemLines.join('、')}、總金額：${total}元${note ? '\n備註：' + note : ''}`;

    rowData.push(summary, total, replyMsg);

    // Write row
    sheet.getRange(row, 1, 1, rowData.length).setValues([rowData]);

    // Format: center qty columns
    const qtyRange = sheet.getRange(row, 3, 1, PRODUCT_LIST.length);
    qtyRange.setHorizontalAlignment('center');

    // Wrap summary & reply
    sheet.getRange(row, 3 + PRODUCT_LIST.length, 1, 1).setWrap(true);
    sheet.getRange(row, 3 + PRODUCT_LIST.length + 2, 1, 1).setWrap(true);

    // Alternating row colors
    const bgColor = (row % 2 === 0) ? '#e8f5ea' : '#ffffff';
    sheet.getRange(row, 1, 1, rowData.length).setBackground(bgColor);

    return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } finally {
    lock.releaseLock();
  }
}

function setupHeaders(sheet) {
  const headers = ['時間戳記', '姓名'];
  PRODUCT_LIST.forEach(p => headers.push(`${p.name}\n$${p.price}`));
  headers.push('訂購摘要', '總金額（元）', '備註及回覆');

  const range = sheet.getRange(1, 1, 1, headers.length);
  range.setValues([headers]);
  range.setBackground('#5c3317');
  range.setFontColor('#ffffff');
  range.setFontWeight('bold');
  range.setHorizontalAlignment('center');
  range.setVerticalAlignment('middle');
  range.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sheet.setRowHeight(1, 48);

  // Column widths
  sheet.setColumnWidth(1, 155); // 時間
  sheet.setColumnWidth(2, 75);  // 姓名
  for (let i = 0; i < PRODUCT_LIST.length; i++) {
    sheet.setColumnWidth(3 + i, 78);
  }
  const offset = 3 + PRODUCT_LIST.length;
  sheet.setColumnWidth(offset, 220);     // 訂購摘要
  sheet.setColumnWidth(offset + 1, 90);  // 總金額
  sheet.setColumnWidth(offset + 2, 280); // 備註及回覆

  // Freeze
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);
}

function getOrCreateSpreadsheet() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty('SPREADSHEET_ID');

  if (ssId) {
    try { return SpreadsheetApp.openById(ssId); }
    catch(e) { /* recreate */ }
  }

  const ss = SpreadsheetApp.create('波琳小姐蝴蝶酥 — 訂單');
  ssId = ss.getId();
  props.setProperty('SPREADSHEET_ID', ssId);
  setupHeaders(ss.getSheets()[0]);
  return ss;
}

function getSpreadsheetUrl() {
  const props = PropertiesService.getScriptProperties();
  const ssId = props.getProperty('SPREADSHEET_ID');
  if (ssId) {
    Logger.log('試算表網址：https://docs.google.com/spreadsheets/d/' + ssId);
  } else {
    Logger.log('尚未建立試算表，送出第一筆訂單後自動建立');
  }
}
