/* * 🍰 MOODRI 甜點收單系統 v6.0 (API 強化版)
 * 功能：
 * 1. 提供 API 給 Lovable (doGet/doPost)
 * 2. 自動寫入 Google Sheet 訂單紀錄
 * 3. 庫存自動扣除
 * 4. 自動寄送確認信與通知信
 */

/* =========================================
   1. 核心設定與工具函式
   ========================================= */

// 防止 XSS 攻擊 (HTML 跳脫字元)
function escapeHtml(str) {
  if (!str) return '';
  return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

// 驗證 Email 格式
function validateEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

// 取得目前的訂單系統開關狀態 (Open/Closed)
function getOrderStatus() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('系統設定');
  if (!sheet) return 'open';
  return sheet.getRange('C2').getValue();
}

/* =========================================
   2. API 接口 (前後端溝通橋樑)
   ========================================= */

/**
 * 處理前端的 GET 請求 (讀取資料)
 * Lovable 會呼叫這個函式來取得「產品列表」與「商店設定」
 */
function doGet(e) {
  // 若網址帶有 ?action=getData，回傳 JSON 資料
  if (e.parameter.action === 'getData') {
    try {
      var result = {
        status: 'success',
        products: getProductList(),      // 抓取產品清單
        logistics: getLogisticsOptions(),// 抓取物流選項
        config: getConfigData()          // 抓取商店設定 (名稱、公告)
      };
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // 若無參數，回傳舊版 HTML (可保留作為備用或測試頁面)
  checkAndInitialize();
  var template = HtmlService.createTemplateFromFile('index');
  var config = getConfigData();
  template.shopName = config.shopName || "預設店名";
  template.theme = config.theme || "theme-beige";
  template.announcement = config.announcement || "";
  template.formTitle = config.formTitle || "訂購資訊";
  template.formNote = config.formNote || "";

  return template.evaluate()
      .setTitle(template.shopName + " - 線上點單")
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 處理前端的 POST 請求 (接收訂單)
 * Lovable 送出訂單時會呼叫這裡
 */
function doPost(e) {
  try {
    // 解析 JSON 資料
    var data = JSON.parse(e.postData.contents);
    
    // 呼叫主要處理邏輯
    var result = submitOrder(data);

    // 回傳結果
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      "status": "error", 
      "message": "系統錯誤: " + err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/* =========================================
   3. 資料讀取邏輯 (從 Google Sheet 抓資料)
   ========================================= */

// 📦 取得物流選項 (從「系統設定」分頁讀取)
function getLogisticsOptions() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('系統設定');
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(2, 4, lastRow - 1, 3).getValues();
  var options = [];
  data.forEach(function(row) {
    if (row[0] !== "") {
      options.push({ name: escapeHtml(row[0]), price: row[1] || 0, freeThreshold: row[2] || 999999 });
    }
  });
  return options;
}

// 🍰 取得產品列表 (從「產品設定」分頁讀取)
function getProductList() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('產品設定');
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  // 讀取 B~H 欄 (產品資料區塊)
  var data = sheet.getRange(2, 2, lastRow - 1, 7).getValues();
  var products = [];

  data.forEach(function(row) {
    // 欄位對應：[0]名稱, [1]價格, [2]描述, [3]上架?, [4]圖片, [5]售完?, [6]庫存
    var enabled = row[3];
    if (enabled === true || enabled === "TRUE" || enabled === "Yes" || enabled === "上架") {
      products.push({
        name: escapeHtml(row[0]),
        price: row[1],
        desc: escapeHtml(row[2]),
        img: escapeHtml(row[4] || ""),
        soldOut: (row[5] === true || row[5] === "TRUE"),
        stock: Number(row[6]) || 0
      });
    }
  });
  return products;
}

// ⚙️ 取得商店基本設定
function getConfigData() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('系統設定');
  if (!sheet) return {};
  var themeMap = {
    '☕️ 經典文青': 'theme-beige', '🌸 甜美粉紅': 'theme-pink', '🎩 質感黑金': 'theme-dark',
    '🌲 清新森林': 'theme-forest', '☁️ 極簡灰調': 'theme-grey', '👾 賽博龐克': 'theme-cyber',
    '🎸 復古金屬': 'theme-metal', '💎 高雅深藍': 'theme-blue'
  };
  return {
    shopName: sheet.getRange('B1').getValue(),
    theme: themeMap[sheet.getRange('B2').getValue()] || 'theme-beige',
    announcement: escapeHtml(sheet.getRange('B3').getValue()),
    formTitle: escapeHtml(sheet.getRange('B4').getValue()),
    formNote: escapeHtml(sheet.getRange('B5').getValue().toString()).replace(/\n/g, '<br>')
  };
}

/* =========================================
   4. 訂單處理核心 (Submit Order)
   ========================================= */

// 📝 處理訂單：檢查資料、扣庫存、寫入 Sheet、寄信
function submitOrder(formObject) {
  const lock = LockService.getScriptLock();
  try {
    // 鎖定 5 秒，避免多人同時下單導致庫存錯誤
    lock.waitLock(5000);

    // --- A. 驗證欄位 ---
    if (!formObject.customerName || !formObject.customerPhone || !formObject.pickupMethod || !formObject.bankLast5) {
      throw new Error("請確認所有必填欄位已填寫！");
    }
    if (!/^\d{5}$/.test(formObject.bankLast5)) throw new Error("匯款帳號必須填寫 5 位數字！");
    if (formObject.customerEmail && !validateEmail(formObject.customerEmail)) throw new Error("Email 格式錯誤");
    if (!formObject.cartData) throw new Error("購物車內容為空！");

    // --- B. 解析購物車 ---
    let cartItems;
    try {
      cartItems = (typeof formObject.cartData === 'string') ? JSON.parse(formObject.cartData) : formObject.cartData;
      if (!Array.isArray(cartItems) || cartItems.length === 0) throw new Error();
    } catch (e) {
      throw new Error("購物車資料格式錯誤！");
    }

    // --- C. 檢查與扣除庫存 ---
    const productSheet = SpreadsheetApp.getActive().getSheetByName('產品設定');
    const productData = productSheet.getRange(2, 2, productSheet.getLastRow() - 1, 7).getValues(); 

    // 第一次迴圈：純檢查 (避免檢查到一半發現沒貨)
    cartItems.forEach(item => {
      const idx = productData.findIndex(p => p[0] === item.name);
      if (idx === -1) throw new Error(item.name + " 不存在！");
      const stock = Number(productData[idx][6]);
      if (item.qty > stock) throw new Error(item.name + " 庫存不足，剩餘：" + stock);
    });

    // 第二次迴圈：實際扣庫存
    cartItems.forEach(item => {
      const idx = productData.findIndex(p => p[0] === item.name);
      const row = idx + 2;
      const newStock = Number(productData[idx][6]) - item.qty;
      productSheet.getRange(row, 8).setValue(newStock);      // 更新庫存
      productSheet.getRange(row, 7).setValue(newStock <= 0); // 若<=0 自動勾選「售完」
    });

    // --- D. 寫入訂單紀錄 ---
    const ss = SpreadsheetApp.getActive();
    let sheet = ss.getSheetByName('訂單紀錄');
    if (!sheet) { checkAndInitialize(); sheet = ss.getSheetByName('訂單紀錄'); }

    const orderId = Utilities.formatDate(new Date(), "GMT+8", "yyyyMMdd-HHmmss");
    const timestamp = new Date();
    const orderDetails = cartItems.map(i => escapeHtml(i.name) + " x" + i.qty).join("\n");
    const cleanedAddress = (formObject.address || "").replace(/\[.*?\]\s*/, ''); // 清洗地址

    sheet.appendRow([
      orderId, timestamp,
      escapeHtml(formObject.customerName),
      escapeHtml(formObject.customerPhone),
      escapeHtml(formObject.customerEmail || ""),
      escapeHtml(formObject.socialId || ""),
      escapeHtml(formObject.pickupMethod),
      escapeHtml(cleanedAddress),
      orderDetails,
      escapeHtml(formObject.note || ""),
      formObject.totalAmount,
      formObject.bankLast5,
      "未處理" // 預設狀態
    ]);

    // --- E. 寄送信件 ---
    // 1. 通知老闆
    sendAdminNewOrderEmail(orderId, formObject, cartItems);
    
    // 2. 通知客人 (若有 Email)
    if (formObject.customerEmail && validateEmail(formObject.customerEmail)) {
      try {
        sendConfirmationEmail(formObject, orderId, cartItems);
      } catch (err) {
        Logger.log("❌ 顧客確認信寄送失敗：" + err.message);
      }
    }

    return { status: "success", orderId: orderId };

  } catch (e) {
    return { status: "error", message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/* =========================================
   5. Email 寄送模組
   ========================================= */

// 📩 寄給客人：訂單確認信
function sendConfirmationEmail(data, orderId, cartItems) {
  const ss = SpreadsheetApp.getActive();
  const shopName = data.shopName || ss.getSheetByName('系統設定').getRange('B1').getValue() || '商店';
  const subject = `【訂單確認】${shopName}｜訂單編號 ${orderId}`;
  
  const displayAddress = data.address ? data.address.replace(/\[.*?\]\s*/, '') : '';
  const itemsHtml = cartItems.map(item => `<li>${escapeHtml(item.name)} x <b>${item.qty}</b>（$${item.price}）</li>`).join("");

  // ✨ 信件內容樣板 (HTML)
  const body = `
  <div style="font-family: sans-serif; line-height: 1.6; color: #333;">
    <h2 style="color: #2c3e50;">🎉 感謝您的訂購！</h2>
    <p>${escapeHtml(data.customerName)} 您好：</p>
    <div style="background: #f9f9f9; padding: 15px; border-radius: 8px; margin: 20px 0; border: 1px solid #eee;">
      <h3 style="margin-top:0; border-bottom: 2px solid #d35336; display: inline-block;">📋 訂單內容</h3>
      <ul style="margin-top: 15px;">${itemsHtml}</ul>
      <hr style="border:0; border-top:1px solid #ddd; margin: 15px 0;">
      <p><b>取貨方式：</b> ${escapeHtml(data.pickupMethod)}</p>
      ${displayAddress ? `<p><b>地址：</b> ${escapeHtml(displayAddress)}</p>` : ""}
      <p><b>匯款後五碼：</b> ${escapeHtml(data.bankLast5)}</p>
      <p style="font-weight: bold; color: #c0392b;"><b>總金額：</b> $${data.totalAmount}</p>
    </div>
    <div style="text-align: center; margin-top: 30px; font-size: 13px; color: #666;">
      <p>📦 甜點皆為接單後新鮮製作，完成後將盡速安排出貨。<br>謝謝您的支持 🧡</p>
    </div>
  </div>`;

  MailApp.sendEmail({
    to: data.customerEmail.trim(),
    subject: subject,
    htmlBody: body,
    name: shopName
  });
}

// 📩 寄給老闆：新訂單通知
function sendAdminNewOrderEmail(orderId, formObject, cartItems) {
  const configSheet = SpreadsheetApp.getActive().getSheetByName('系統設定');
  const adminEmail = configSheet?.getRange('B6').getValue();
  if (!adminEmail) return;

  const shopName = configSheet.getRange('B1').getValue() || '商店';
  const itemsText = cartItems.map(item => `${item.name} x ${item.qty} ($${item.price})`).join('\n');

  const body = `
新訂單成立 🎉
訂單編號：${orderId}
時間：${Utilities.formatDate(new Date(), "GMT+8", "yyyy/MM/dd HH:mm")}

【顧客】${formObject.customerName} / ${formObject.customerPhone}
【Email】${formObject.customerEmail || '未填'}
【取貨】${formObject.pickupMethod}
【內容】
${itemsText}

【總額】$${formObject.totalAmount}
【後五碼】${formObject.bankLast5}
  `.trim();

  MailApp.sendEmail({ to: adminEmail, subject: `📥【新訂單】${shopName}｜${orderId}`, body: body, name: shopName });
}

// 📩 寄給客人：付款成功通知
function sendPaymentReceivedEmail(order) {
  const shopName = SpreadsheetApp.getActive().getSheetByName('系統設定').getRange('B1').getValue() || '商店';
  const subject = `【付款確認】${shopName} - 訂單 ${order.orderId}`;
  
  // 將訂單內容換行符號轉為清單
  const itemsHtml = order.orderDetails.split("\n").map(l => l.trim() ? `<li>${escapeHtml(l)}</li>` : "").join("");

  const body = `
    <div style="font-family: sans-serif; color: #333;">
      <h2 style="color: #2c3e50;">💰 付款成功通知</h2>
      <p>親愛的 <b>${escapeHtml(order.customerName)}</b> 您好，我們已確認您的匯款。</p>
      <div style="background: #f9f9f9; padding: 15px; border-radius: 8px;">
        <h3>📋 訂單明細</h3>
        <ul>${itemsHtml}</ul>
        <p><b>總金額：</b> $${order.totalAmount}</p>
      </div>
      <p style="text-align:center; color:#666; margin-top:20px;">我們會盡快為您安排製作！🧡</p>
    </div>`;

  MailApp.sendEmail({ to: order.customerEmail, subject: subject, htmlBody: body, name: shopName });
}

// 📩 寄給客人：出貨通知
function sendShippingNotificationEmail(order) {
  if (!order.customerEmail) throw new Error("顧客 Email 為空");
  
  const shopName = SpreadsheetApp.getActive().getSheetByName('系統設定').getRange('B1').getValue() || '商店';
  const subject = `【出貨通知】${shopName} - 訂單 ${order.orderId}`;
  const itemsHtml = order.orderDetails.split("\n").map(l => l.trim() ? `<li>${escapeHtml(l)}</li>` : "").join("");

  const body = `
    <div style="font-family: sans-serif; color: #333;">
      <h2 style="color: #2c3e50;">📦 您的訂單已出貨！</h2>
      <p>親愛的 <b>${escapeHtml(order.customerName)}</b> 您好，您的甜點已經出發囉。</p>
      <div style="background: #f9f9f9; padding: 15px; border-radius: 8px;">
        <h3>📋 訂單資訊</h3>
        <p><b>取貨方式：</b> ${escapeHtml(order.pickupMethod)}</p>
        ${order.trackingNumber ? `<p style="color: #D26900;"><b>物流單號：</b> ${escapeHtml(order.trackingNumber)}</p>` : ''}
        <ul>${itemsHtml}</ul>
      </div>
      <p style="text-align:center; color:#666; margin-top:20px;">祝您有個美好的一天！🍰</p>
    </div>`;

  MailApp.sendEmail({ to: order.customerEmail, subject: subject, htmlBody: body, name: shopName });
}

/* =========================================
   6. 觸發事件：狀態變更自動處理 (寄信/退庫存)
   ========================================= */

// ⚠️ 若要啟用此功能，請在觸發條件設定中，將「編輯時」綁定到此函式
var isHandlingEdit = false;

function processOrderUpdate(e) {
  if (isHandlingEdit) return;
  var sheet = e.source.getSheetByName('訂單紀錄');
  if (!sheet || e.range.getRow() < 2) return;

  var col = e.range.getColumn();
  var newValue = String(e.range.getValue()).trim();
  var row = e.range.getRow();

  // 僅監聽 M 欄 (狀態欄位)
  if (col === 13) {
    
    // 重設為未處理 -> 清除紀錄
    if (newValue === "未處理") {
      sheet.getRange(row, 14, 1, 4).clearContent(); 
      e.source.toast("已重設狀態，相關紀錄已清除。", "系統");
      return;
    }

    // 取得該列資料
    var rowData = sheet.getRange(row, 1, 1, 17).getValues()[0];
    var order = {
      orderId: rowData[0], customerName: rowData[2], customerEmail: rowData[4],
      pickupMethod: rowData[6], address: rowData[7], orderDetails: rowData[8],
      totalAmount: rowData[10], bankLast5: rowData[11],
      paymentEmailStatus: rowData[13], shippingEmailStatus: rowData[14],
      stockRefundStatus: rowData[15], trackingNumber: rowData[16]
    };
    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd HH:mm");

    // 狀態：取消 -> 退庫存
    if (newValue === "取消" && order.stockRefundStatus.indexOf("✅") === -1) {
      try {
        refundStock(order.orderDetails);
        sheet.getRange(row, 16).setValue("✅ (" + timestamp + ")");
        e.source.toast("已加回庫存", "系統");
      } catch (err) { sheet.getRange(row, 16).setValue("❌ " + err.message); }
    }

    // 狀態：已付款 -> 寄信
    if (newValue === "已付款" && order.customerEmail && !order.paymentEmailStatus) {
      try {
        sendPaymentReceivedEmail(order);
        sheet.getRange(row, 14).setValue("✅ (" + timestamp + ")");
      } catch (err) { sheet.getRange(row, 14).setValue("❌ " + err.message); }
    }

    // 狀態：已出貨 -> 寄信
    if (newValue === "已出貨" && order.customerEmail && !order.shippingEmailStatus) {
      try {
        sendShippingNotificationEmail(order);
        sheet.getRange(row, 15).setValue("✅ (" + timestamp + ")");
      } catch (err) { sheet.getRange(row, 15).setValue("❌ " + err.message); }
    }
  }
}

// 📦 退還庫存邏輯
function refundStock(orderDetails) {
  var ss = SpreadsheetApp.getActive();
  var productSheet = ss.getSheetByName('產品設定');
  var productData = productSheet.getRange(2, 2, productSheet.getLastRow() - 1, 7).getValues();

  var lines = orderDetails.split("\n");
  lines.forEach(function(line) {
    var match = line.match(/^(.+?)\s*x\s*(\d+)/); // 解析 "蛋糕 x 2"
    if (match) {
      var itemName = match[1].trim();
      var qtyToRefund = parseInt(match[2]);

      for (var i = 0; i < productData.length; i++) {
        if (productData[i][0] === itemName) {
          var newStock = Number(productData[i][6]) + qtyToRefund;
          productSheet.getRange(i + 2, 8).setValue(newStock); // 加回庫存
          if (newStock > 0) productSheet.getRange(i + 2, 7).setValue(false); // 取消售完勾選
          break;
        }
      }
    }
  });
}

/* =========================================
   7. 系統初始化與選單
   ========================================= */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🍰 蛋糕系統')
      .addItem('💰 自動對帳', 'runAutoReconcile')
      .addItem('📊 產量統計', 'calculateProduction')
      .addItem('📈 營收戰情室', 'createDashboard')
      .addSeparator()
      .addItem('🔄 系統修復', 'checkAndInitialize')
      .addToUi();
}

// 系統初始化 (產生必要分頁)
function checkAndInitialize() {
  var ss = SpreadsheetApp.getActive();
  
  if (!ss.getSheetByName('系統設定')) {
    var s = ss.insertSheet('系統設定');
    s.getRange('A1:A5').setValues([['店鋪名稱'], ['風格主題'], ['公告/副標'], ['表單標題區'], ['訂購須知']]).setBackground('#eaeaea');
    s.getRange('B1').setValue('MOODRI 暮日甜點');
    s.getRange('D1:F1').setValues([['物流名稱', '運費', '免運門檻']]);
  }
  
  if (!ss.getSheetByName('產品設定')) {
    var s = ss.insertSheet('產品設定');
    s.getRange('A1:G1').setValues([['排序', '產品名稱', '價格', '描述', '上架?', '圖片', '售完?']]);
  }

  if (!ss.getSheetByName('訂單紀錄')) {
    var s = ss.insertSheet('訂單紀錄');
    s.getRange('A1:O1').setValues([['訂單編號', '下單時間', '姓名', '電話', 'Email', '社群帳號', '取貨方式', '地址', '內容', '備註', '總金額', '後五碼', '狀態', '付款信', '出貨信']]);
  }
}

// 💰 自動對帳 (比對後五碼與金額)
function runAutoReconcile() {
  var ss = SpreadsheetApp.getActive();
  var orderSheet = ss.getSheetByName('訂單紀錄');
  var bankSheet = ss.getSheetByName('銀行對帳');
  if (!orderSheet || !bankSheet) return;

  var orderData = orderSheet.getDataRange().getValues();
  var bankData = bankSheet.getDataRange().getValues();
  var matchCount = 0;

  for (var i = 1; i < orderData.length; i++) {
    // 若狀態不是已付款，且有填後五碼
    if (orderData[i][12] !== "已付款" && orderData[i][11]) {
      var last5 = String(orderData[i][11]).trim();
      var amount = orderData[i][10];

      for (var j = 1; j < bankData.length; j++) {
        // 銀行資料 C欄金額(2), D欄帳號(3)
        if (bankData[j][2] == amount && String(bankData[j][3]).includes(last5)) {
          orderSheet.getRange(i+1, 13).setValue("已付款");
          bankSheet.getRange(j+1, 5).setValue("✅ 已核銷");
          matchCount++;
          break;
        }
      }
    }
  }
  SpreadsheetApp.getUi().alert('對帳完成，共匹配 ' + matchCount + ' 筆');
}

// 📊 產量統計
function calculateProduction() {
  var ss = SpreadsheetApp.getActive();
  var orderSheet = ss.getSheetByName('訂單紀錄');
  var statSheet = ss.getSheetByName('製作統計') || ss.insertSheet('製作統計');
  statSheet.clear();
  statSheet.getRange('A1:B1').setValues([['產品名稱', '待製作數量']]).setBackground('#fbbc04');

  var orders = orderSheet.getDataRange().getValues();
  var counts = {};

  // 從第 2 列開始讀
  for (var i = 1; i < orders.length; i++) {
    var status = orders[i][12]; // M欄
    // 只有這些狀態才需要製作
    if (["未處理", "已付款", "製作中"].includes(status)) {
      var lines = String(orders[i][8]).split("\n"); // I欄內容
      lines.forEach(function(line) {
        var parts = line.split(" x");
        if (parts.length === 2) {
          var name = parts[0].trim();
          var qty = parseInt(parts[1]);
          counts[name] = (counts[name] || 0) + qty;
        }
      });
    }
  }

  var output = Object.keys(counts).map(function(k) { return [k, counts[k]]; });
  if (output.length) statSheet.getRange(2, 1, output.length, 2).setValues(output);
  statSheet.activate();
}

// 📈 營收戰情室
function createDashboard() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('📊 營收報表');
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet('📊 營收報表', 0);
  
  sheet.getRange('A1').setValue('老板戰情室：即時營收概況').setFontSize(18);
  
  // 設定公式
  sheet.getRange('A4').setValue('📅 本月營收');
  sheet.getRange('A5').setFormula('=SUMIFS(\'訂單紀錄\'!J:J, \'訂單紀錄\'!L:L, "已付款", \'訂單紀錄\'!B:B, ">="&EOMONTH(TODAY(),-1)+1, \'訂單紀錄\'!B:B, "<"&EOMONTH(TODAY(),0)+1)');
  
  sheet.getRange('D4').setValue('⚡ 今日營收');
  sheet.getRange('D5').setFormula('=SUMIFS(\'訂單紀錄\'!J:J, \'訂單紀錄\'!L:L, "已付款", \'訂單紀錄\'!B:B, ">="&TODAY(), \'訂單紀錄\'!B:B, "<"&TODAY()+1)');

  sheet.getRange('G4').setValue('⚠️ 待處理金額');
  sheet.getRange('G5').setFormula('=SUMIFS(\'訂單紀錄\'!J:J, \'訂單紀錄\'!L:L, "未處理")');

  // 美化
  sheet.getRange('A5:H5').setNumberFormat('$0,0').setFontSize(20).setFontWeight('bold');
  SpreadsheetApp.getUi().alert('戰情室已建立！');
}
