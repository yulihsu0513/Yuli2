// ==========================================
// 主廚屋｜GAS 雲端點餐後端系統 (升級版)
// Code.gs
// 支援 GitHub Pages 外部跨網域送單 & Google 試算表同步
// ==========================================

const SPREADSHEET_ID = '1qmDbaFGiGT_TNvCvUtDedIu1Ujc1VCgO3o0zsj479ds';

const MENU = [
  { id: 'M001', name: '川味椒麻雞飯', category: 'chicken', price: 135, tag: '👑 人氣冠軍' },
  { id: 'M002', name: '炭燒雞腿排飯', category: 'chicken', price: 135, tag: '🔥 招牌現烤' },
  { id: 'M003', name: '玫瑰油雞腿飯', category: 'chicken', price: 135, tag: '✨ 經典港式' },
  { id: 'M004', name: '東坡肉飯', category: 'pork', price: 135, tag: '🥩 慢火滷製' },
  { id: 'M005', name: '黯然銷魂飯', category: 'pork', price: 125, tag: '🍳 老饕必點' },
  { id: 'M006', name: '香酥豬排飯', category: 'pork', price: 120, tag: '酥脆金黃' },
  { id: 'M007', name: '漁夫鮮魚飯', category: 'seafood', price: 140, priceText: '時價', tag: '🐟 當日現撈' },
  { id: 'M008', name: '農夫鮮蔬飯', category: 'veggie', price: 115, tag: '🥬 清爽少油' }
];

function doGet(e) {
  // 支援 JSON API 格式取得最新菜單
  if (e && e.parameter && e.parameter.action === 'getMenu') {
    return ContentService.createTextOutput(JSON.stringify({ success: true, restaurantName: '主廚屋', menu: MENU }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 預設提供 GAS 內部網頁
  return HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle('主廚屋｜線上點餐系統')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 支援來自 GitHub Pages 的 POST 跨網域送單
function doPost(e) {
  try {
    let orderData = null;
    if (e && e.postData && e.postData.contents) {
      orderData = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      orderData = e.parameter;
    }

    if (!orderData) {
      throw new Error('未收到任何有效訂單資料');
    }

    const result = submitOrder(orderData);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, orderId: result.orderId, total: result.total }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getMenu() {
  return { restaurantName: '主廚屋', menu: MENU };
}

function getOrderSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheets()[0];

  const headers = [
    '訂單編號', '訂單時間', '姓名', '部門', '電話 / 分機', '取餐方式',
    '外送地址', '統一編號', '餐具需求', '期望送達時間', '餐點內容明細',
    '餐點小計', '五穀飯加額', '總金額', '訂單備註', '意見回饋', '訂單狀態'
  ];

  if (sheet.getLastRow() === 0 || sheet.getRange('A1').getValue() === '') {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#87391F')
      .setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 180);
    sheet.setColumnWidth(2, 170);
    sheet.setColumnWidth(3, 100);
    sheet.setColumnWidth(4, 100);
    sheet.setColumnWidth(5, 130);
    sheet.setColumnWidth(6, 100);
    sheet.setColumnWidth(7, 260);
    sheet.setColumnWidth(8, 120);
    sheet.setColumnWidth(9, 110);
    sheet.setColumnWidth(10, 170);
    sheet.setColumnWidth(11, 480);
    sheet.setColumnWidth(12, 110);
    sheet.setColumnWidth(13, 110);
    sheet.setColumnWidth(14, 110);
    sheet.setColumnWidth(15, 240);
    sheet.setColumnWidth(16, 240);
    sheet.setColumnWidth(17, 100);
  }
  return sheet;
}

function submitOrder(order) {
  if (!order) throw new Error('沒有收到訂單資料');
  if (!order.customerName || !String(order.customerName).trim()) throw new Error('請填寫訂購人姓名');
  if (!order.department || !String(order.department).trim()) throw new Error('請選擇部門');
  if (!order.items || !order.items.length) throw new Error('請至少選擇一項餐點');

  let mealSubtotal = 0;
  let multigrainTotal = 0;
  const itemTexts = [];

  order.items.forEach(function(orderItem) {
    const menuItem = MENU.find(function(item) { return item.id === orderItem.id; });
    if (!menuItem) throw new Error('餐點資料不符');

    const qty = Number(orderItem.qty) || 0;
    if (qty <= 0) return;

    const riceType = orderItem.riceType === '五穀飯' ? '五穀飯' : '白飯';
    const riceAmount = orderItem.riceAmount || '正常飯';
    const customize = orderItem.customOptions && orderItem.customOptions.length ? '【' + orderItem.customOptions.join(', ') + '】' : '';

    const price = menuItem.price > 0 ? menuItem.price : (Number(orderItem.price) || 140);
    const baseSubtotal = price * qty;
    const riceExtra = riceType === '五穀飯' ? 15 * qty : 0;

    mealSubtotal += baseSubtotal;
    multigrainTotal += riceExtra;

    let text = menuItem.name + ' × ' + qty + ' 份 ｜ ' + riceType + ' (' + riceAmount + ') ' + customize;
    text += ' ｜ 小計 $' + (baseSubtotal + riceExtra);
    itemTexts.push(text);
  });

  if (!itemTexts.length) throw new Error('請至少選擇一項有效餐點');

  const total = mealSubtotal + multigrainTotal;
  const now = new Date();

  const orderId =
    'ORD-' +
    Utilities.formatDate(now, 'Asia/Taipei', 'yyyyMMdd-HHmmss') +
    '-' +
    Math.floor(Math.random() * 900 + 100);

  const orderTime = Utilities.formatDate(now, 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');

  getOrderSheet_().appendRow([
    orderId,
    orderTime,
    String(order.customerName).trim(),
    String(order.department).trim(),
    order.phone || '',
    order.deliveryMethod || '自取',
    order.deliveryAddress || '',
    order.taxId || '',
    order.tableware || '需要餐具',
    order.deliveryTime || '',
    itemTexts.join('\n'),
    mealSubtotal,
    multigrainTotal,
    total,
    order.note || '',
    order.feedback || '',
    '新訂單'
  ]);

  return { success: true, orderId: orderId, total: total };
}

function testSheetConnection() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  Logger.log('Google 試算表連線成功：' + ss.getUrl());
}
