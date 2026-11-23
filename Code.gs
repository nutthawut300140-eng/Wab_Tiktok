/**
 * Google Apps Script - Ambliss_style Online TikTok Live Cloud ERP
 * เชื่อมต่อ Google Sheets 100% สด Real-time รองรับครบทุกเฟส 1 - 5 สมบูรณ์แบบ
 */

function getSpreadsheet() {
  try {
    return SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {
    return SpreadsheetApp.openById("1VM96JDZwFdFN9ksNZ45x0hT1Pg0ysa9JPv6wChw5B4Q");
  }
}

function doGet(e) {
  return HtmlService.createTemplateFromFile("Index")
    .evaluate()
    .setTitle("Ambliss_style - TikTok Live Backoffice ERP")
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (err) {
    return '<!-- Error loading ' + filename + ': ' + err.message + ' -->';
  }
}

function sanitizeCell(val) {
  if (val === null || val === undefined) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, "GMT+7", "yyyy-MM-dd HH:mm");
  }
  return String(val);
}

function findSheet(ss, keywords, targetGid) {
  if (!ss) ss = getSpreadsheet();
  var sheets = ss.getSheets();
  if (targetGid !== undefined && targetGid !== null) {
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getSheetId() == targetGid) return sheets[i];
    }
  }
  var kw = Array.isArray(keywords) ? keywords : [keywords];
  for (var i = 0; i < sheets.length; i++) {
    var sName = sheets[i].getName().trim().toLowerCase();
    for (var k = 0; k < kw.length; k++) {
      if (sName.indexOf(String(kw[k]).toLowerCase()) !== -1) return sheets[i];
    }
  }
  return null;
}

function readSheetData(sheet) {
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var headers = data[0];
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var hasContent = false;
    for (var c = 0; c < row.length; c++) {
      if (row[c] !== "" && row[c] !== null && row[c] !== undefined) {
        hasContent = true;
        break;
      }
    }
    if (!hasContent) continue;
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var key = String(headers[j]).trim();
      if (key) obj[key] = sanitizeCell(row[j]);
    }
    var firstVal = sanitizeCell(row[0]).trim();
    if (firstVal !== "") obj['id'] = obj['id'] || firstVal;
    rows.push(obj);
  }
  return rows;
}

/* ==========================================================================
   1. PRODUCTS API
   ========================================================================== */
function getProductsData() {
  var ss = getSpreadsheet();
  var sheet = findSheet(ss, ["สินค้าและสต๊อก", "products", "สินค้า"]);
  var rows = readSheetData(sheet);
  rows.forEach(function(r) {
    var smartId = r['รหัสสินค้า'] || r['รหัส'] || r['code'] || r['sku'] || r['id'] || Object.values(r)[0] || '-';
    r['รหัสสินค้า'] = smartId;
    r['id'] = smartId;
  });
  return { success: true, products: rows };
}

function saveProductToSheet(prod) {
  var ss = getSpreadsheet();
  var sheet = findSheet(ss, ["สินค้าและสต๊อก", "products", "สินค้า"]);
  var headers = ["รหัสสินค้า", "ชื่อสินค้า", "หมวดหมู่", "สัดส่วน/ไซส์", "ราคาขาย", "ต้นทุน", "คงเหลือ", "หน่วยนับ", "สถานะ", "หมายเหตุ", "อัปเดตล่าสุด"];
  
  if (!sheet) {
    sheet = ss.insertSheet("5. สินค้าและสต๊อก (Products)");
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setBackground("#0f172a").setFontColor("#ffffff").setFontWeight("bold");
    sheet.setFrozenRows(1);
  }

  var data = sheet.getDataRange().getValues();
  var foundRow = -1;
  var prodId = String(prod.id || prod['รหัสสินค้า'] || '').trim();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === prodId) { foundRow = i + 1; break; }
  }

  var rowValues = [
    prodId,
    prod.name || prod['ชื่อสินค้า'] || '',
    prod.category || prod['หมวดหมู่'] || '',
    prod.size || prod['สัดส่วน/ไซส์'] || '',
    Number(prod.price || prod['ราคาขาย'] || 0),
    Number(prod.cost || prod['ต้นทุน'] || 0),
    Number(prod.stock || prod['คงเหลือ'] || 0),
    prod.unit || prod['หน่วยนับ'] || 'ชิ้น',
    prod.status || prod['สถานะ'] || 'พร้อมขาย',
    prod.notes || prod['หมายเหตุ'] || '',
    Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm")
  ];

  if (foundRow > 0) sheet.getRange(foundRow, 1, 1, rowValues.length).setValues([rowValues]);
  else sheet.appendRow(rowValues);

  return { success: true, product: prod };
}

function deleteProductFromSheet(prodId) {
  var ss = getSpreadsheet();
  var sheet = findSheet(ss, ["สินค้าและสต๊อก", "products", "สินค้า"]);
  if (!sheet) return { success: false, message: "ไม่พบชีตสินค้า" };
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(prodId).trim()) {
      sheet.deleteRow(i + 1);
      return { success: true, deletedId: prodId };
    }
  }
  return { success: false, message: "ไม่พบสินค้ารหัสนี้ในชีต" };
}

/* ==========================================================================
   2. LIVE SESSIONS & CUSTOMERS API
   ========================================================================== */
function getLiveSessionsData() {
  var ss = getSpreadsheet();
  var rows = readSheetData(findSheet(ss, ["รอบไลฟ์สด", "live sessions", "ไลฟ์สด"]));
  return { success: true, liveSessions: rows };
}

function getCustomersData() {
  var ss = getSpreadsheet();
  var rows = readSheetData(findSheet(ss, ["ข้อมูลลูกค้า", "customers", "ลูกค้า"]));
  return { success: true, customers: rows };
}

/* ==========================================================================
   3. ORDERS API
   ========================================================================== */
function getOrdersData() {
  var ss = getSpreadsheet();
  var sheet = findSheet(ss, ["บิลคำสั่งซื้อ", "orders"]);
  var rows = readSheetData(sheet);
  return { success: true, orders: rows };
}

function saveOrderToSheet(ord) {
  var ss = getSpreadsheet();
  var sheet = findSheet(ss, ["บิลคำสั่งซื้อ", "orders"]);
  var headers = ["รหัสออเดอร์", "วันที่", "เวลา", "รหัสลูกค้า", "ชื่อ TikTok", "ชื่อผู้รับ", "เบอร์โทร", "รายการสินค้า", "ยอดสินค้า", "ค่าส่ง", "ส่วนลด", "ยอดสุทธิ", "สถานะออเดอร์", "ฝากถุง", "สถานะจัดส่ง", "เลขพัสดุ", "หมายเหตุ", "อัปเดตล่าสุด"];
  
  if (!sheet) {
    sheet = ss.insertSheet("1. บิลคำสั่งซื้อ (Orders)");
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setBackground("#1e293b").setFontColor("#ffffff").setFontWeight("bold");
    sheet.setFrozenRows(1);
  }

  var data = sheet.getDataRange().getValues();
  var foundRow = -1;
  var ordId = String(ord.id || ord['รหัสออเดอร์'] || '').trim();

  if (!ordId) {
    var maxNum = 0;
    for (var i = 1; i < data.length; i++) {
      var m = String(data[i][0]).trim().match(/^ORD-(\d+)$/i);
      if (m) maxNum = Math.max(maxNum, parseInt(m[1], 10));
    }
    ordId = "ORD-" + String(maxNum + 1).padStart(3, '0');
    ord.id = ordId;
  }

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === ordId) { foundRow = i + 1; break; }
  }

  var itemsText = "";
  if (Array.isArray(ord.items)) {
    itemsText = ord.items.map(function(it) {
      return (it.name || it.productName || it.productId || 'สินค้า') + ' x ' + (it.quantity || 1) + ' (' + (it.price || 0) + '฿)';
    }).join(', ');
  } else {
    itemsText = String(ord.itemsText || ord['รายการสินค้า'] || '');
  }

  var rowValues = [
    ordId,
    ord.date || Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd"),
    ord.time || Utilities.formatDate(new Date(), "GMT+7", "HH:mm"),
    ord.customerId || ord['รหัสลูกค้า'] || '-',
    ord.tiktokUsername || ord['ชื่อ TikTok'] || '-',
    ord.customerName || ord['ชื่อผู้รับ'] || '-',
    ord.phone ? "'" + ord.phone : "-",
    itemsText,
    Number(ord.subtotal || ord['ยอดสินค้า'] || 0),
    Number(ord.shippingFee || ord['ค่าส่ง'] || 0),
    Number(ord.discount || ord['ส่วนลด'] || 0),
    Number(ord.grandTotal || ord['ยอดสุทธิ'] || 0),
    ord.status || ord['สถานะออเดอร์'] || 'Confirmed',
    ord.isHoldingBag ? 'ฝากถุง' : (ord['ฝากถุง'] || 'ส่งทันที'),
    ord.shippingStatus || ord['สถานะจัดส่ง'] || 'รอดำเนินการ',
    ord.trackingNumber ? "'" + ord.trackingNumber : "-",
    ord.notes || ord['หมายเหตุ'] || '',
    Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm")
  ];

  if (foundRow > 0) sheet.getRange(foundRow, 1, 1, rowValues.length).setValues([rowValues]);
  else sheet.appendRow(rowValues);

  return { success: true, order: ord };
}

function deleteOrderFromSheet(ordId) {
  var ss = getSpreadsheet();
  var sheet = findSheet(ss, ["บิลคำสั่งซื้อ", "orders"]);
  if (!sheet) return { success: false, message: "ไม่พบชีตออเดอร์" };
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(ordId).trim()) {
      sheet.deleteRow(i + 1);
      return { success: true, deletedId: ordId };
    }
  }
  return { success: false, message: "ไม่พบบิลนี้ในชีต" };
}

/* ==========================================================================
   4. HOLDING BAGS & SHIPMENTS API
   ========================================================================== */
function buildBagObject(cid, tiktok, custName, phone, address, itemsDesc, qty, subtotal, shipFee, grandTotal, payStatus, bagStatus, lastOrder) {
  return {
    'รหัสลูกค้า': cid || '-',
    'customerId': cid || '-',
    'id': cid || '-',
    'ชื่อ TikTok': tiktok || '-',
    'tiktokUsername': tiktok || '-',
    'ชื่อลูกค้า': custName || '-',
    'ชื่อผู้รับ': custName || '-',
    'customerName': custName || '-',
    'เบอร์โทร': phone || '-',
    'phone': phone || '-',
    'ที่อยู่จัดส่ง': address || '-',
    'ที่อยู่': address || '-',
    'address': address || '-',
    'รายการเสื้อผ้าในตะกร้า': itemsDesc || '',
    'รายการสินค้าที่ค้างในตะกร้า': itemsDesc || '',
    'รายการสินค้า': itemsDesc || '',
    'itemsDesc': itemsDesc || '',
    'จำนวนชิ้น': qty,
    'จำนวนชิ้นสะสม': qty,
    'จำนวนตัว': qty,
    'totalItemsCount': qty,
    'qty': qty,
    'ยอดรวมค่าเสื้อ': subtotal,
    'ยอดรวมสินค้า (บาท)': subtotal,
    'ยอดสินค้า': subtotal,
    'subtotal': subtotal,
    'ค่าส่งสะสม': shipFee,
    'ค่าส่งรวม (บาท)': shipFee,
    'ค่าส่ง': shipFee,
    'shippingFee': shipFee,
    'ยอดสุทธิ': grandTotal,
    'ยอดสุทธิรวม': grandTotal,
    'grandTotal': grandTotal,
    'สถานะการชำระ': payStatus,
    'paymentStatus': payStatus,
    'สถานะตะกร้า': bagStatus,
    'สถานะถุง': bagStatus,
    'bagStatus': bagStatus,
    'สั่งซื้อล่าสุด': lastOrder || '-',
    'ออเดอร์ล่าสุดเมื่อ': lastOrder || '-',
    'lastOrderDate': lastOrder || '-'
  };
}

function getHoldingBagsData() {
  var ss = getSpreadsheet();
  var bagSheet = findSheet(ss, ["ตะกร้าสินค้าและฝากถุง", "holding bags", "ฝากถุง"]);
  var orderSheet = findSheet(ss, ["บิลคำสั่งซื้อ", "orders"]);
  var custSheet = findSheet(ss, ["ข้อมูลลูกค้า", "customers"]);

  var bagRows = readSheetData(bagSheet);
  var orderRows = readSheetData(orderSheet);
  var custRows = readSheetData(custSheet);

  var custMap = {};
  custRows.forEach(function(c) {
    var cid = String(c['รหัสลูกค้า'] || c['id'] || '').trim();
    if (cid) custMap[cid] = c;
    var cTk = String(c['ชื่อ TikTok'] || c['tiktokUsername'] || '').trim().toLowerCase();
    if (cTk) custMap[cTk] = c;
  });

  var ordersByCust = {};
  orderRows.forEach(function(o) {
    var isCancelled = String(o['สถานะออเดอร์'] || o['status'] || '').toLowerCase().indexOf('cancel') !== -1;
    if (isCancelled) return;

    var bagFlag = String(o['ฝากถุง'] || o['isHoldingBag'] || '').trim();
    var isHolding = bagFlag.indexOf('ฝาก') !== -1 || bagFlag === 'true' || String(o['สถานะออเดอร์'] || '').indexOf('Holding') !== -1;
    
    var shipStatus = String(o['สถานะจัดส่ง'] || o['shippingStatus'] || '').toLowerCase();
    var isShipped = shipStatus.indexOf('shipped') !== -1 || shipStatus.indexOf('ส่งแล้ว') !== -1 || shipStatus.indexOf('delivered') !== -1;

    if (isHolding && !isShipped) {
      var cid = String(o['รหัสลูกค้า'] || o['customerId'] || '').trim();
      var ctk = String(o['ชื่อ TikTok'] || o['tiktokUsername'] || '').trim().toLowerCase();
      var key = cid || ctk || 'UNKNOWN';

      if (!ordersByCust[key]) ordersByCust[key] = [];
      ordersByCust[key].push(o);
      
      if (cid && ctk && cid !== ctk) {
        if (!ordersByCust[ctk]) ordersByCust[ctk] = ordersByCust[key];
      }
    }
  });

  var bagsList = [];
  var processedKeys = {};

  bagRows.forEach(function(r) {
    var cid = String(r['รหัสลูกค้า'] || r['customerId'] || r['id'] || '').trim();
    var tiktok = String(r['ชื่อ TikTok'] || r['tiktokUsername'] || r['tiktok'] || '').trim();
    var custName = String(r['ชื่อลูกค้า'] || r['ชื่อผู้รับ'] || r['ชื่อ-นามสกุล'] || r['customerName'] || r['name'] || '').trim();
    var phone = String(r['เบอร์โทร'] || r['phone'] || r['tel'] || '').trim();
    var address = String(r['ที่อยู่จัดส่ง'] || r['ที่อยู่'] || r['address'] || '').trim();
    var itemsDesc = String(r['รายการเสื้อผ้าในตะกร้า'] || r['รายการสินค้าที่ค้างในตะกร้า'] || r['รายการสินค้า'] || r['itemsDesc'] || '').trim();
    var qty = Number(r['จำนวนชิ้น'] || r['จำนวนชิ้นสะสม'] || r['จำนวนตัว'] || r['จำนวน'] || r['totalItemsCount'] || r['qty'] || 0);
    var subtotal = Number(r['ยอดรวมค่าเสื้อ'] || r['ยอดรวมสินค้า (บาท)'] || r['ยอดสินค้า'] || r['subtotal'] || r['total'] || 0);
    var shipFee = Number(r['ค่าส่งสะสม'] || r['ค่าส่งรวม (บาท)'] || r['ค่าส่ง'] || r['shippingFee'] || 0);
    var grandTotal = Number(r['ยอดสุทธิ'] || r['ยอดสุทธิรวม'] || r['grandTotal'] || 0);
    var payStatus = String(r['สถานะการชำระ'] || r['paymentStatus'] || '').trim();
    var bagStatus = String(r['สถานะตะกร้า'] || r['สถานะถุง'] || r['bagStatus'] || 'กำลังฝาก').trim();
    var lastOrder = String(r['สั่งซื้อล่าสุด'] || r['ออเดอร์ล่าสุดเมื่อ'] || r['lastOrderDate'] || '').trim();

    var cInfo = (cid && custMap[cid]) || (tiktok && custMap[tiktok.toLowerCase()]) || {};
    if (!custName || custName === '-') custName = String(cInfo['ชื่อ-นามสกุล'] || cInfo['ชื่อลูกค้า'] || cInfo['customerName'] || '-').trim();
    if (!phone || phone === '-') phone = String(cInfo['เบอร์โทร'] || cInfo['phone'] || '-').trim();
    if (!address || address === '-') address = String(cInfo['ที่อยู่จัดส่ง'] || cInfo['ที่อยู่'] || cInfo['address'] || '-').trim();

    var matchedOrders = (cid && ordersByCust[cid]) || (tiktok && ordersByCust[tiktok.toLowerCase()]) || [];
    if (matchedOrders.length > 0) {
      var ordSub = 0;
      var ordItemsList = [];
      var ordQty = 0;
      var hasUnpaid = false;

      matchedOrders.forEach(function(mo) {
        ordSub += Number(mo['ยอดสินค้า'] || mo['subtotal'] || mo['ยอดสุทธิ'] || 0);
        var pStatus = String(mo['สถานะออเดอร์'] || mo['paymentStatus'] || mo['status'] || '').toLowerCase();
        if (pStatus.indexOf('paid') === -1 && pStatus.indexOf('ชำระแล้ว') === -1 && pStatus.indexOf('โอนแล้ว') === -1) {
          hasUnpaid = true;
        }
        var iDesc = String(mo['รายการสินค้า'] || mo['itemsDesc'] || mo['items'] || '').trim();
        if (iDesc) {
          ordItemsList.push(iDesc);
          var parts = iDesc.split(/[\n,]+/);
          parts.forEach(function(pt) {
            var m = pt.match(/x\s*(\d+)/i) || pt.match(/(\d+)\s*(ตัว|ชิ้น)/i);
            ordQty += m ? parseInt(m[1], 10) : 1;
          });
        } else {
          ordQty += 1;
        }
      });

      if (qty <= 0 && ordQty > 0) qty = ordQty;
      if (subtotal <= 0 && ordSub > 0) subtotal = ordSub;
      if (!itemsDesc && ordItemsList.length > 0) itemsDesc = ordItemsList.join('\n');
      if (!payStatus) payStatus = hasUnpaid ? 'รอโอน' : 'ชำระแล้ว';
    }

    if (shipFee <= 0 && qty > 0) {
      shipFee = qty <= 5 ? 40 : 40 + ((qty - 5) * 5);
    }
    if (grandTotal <= 0) {
      grandTotal = subtotal + shipFee;
    }
    if (!payStatus) payStatus = 'รอโอน';

    var bagObj = buildBagObject(cid, tiktok, custName, phone, address, itemsDesc, qty, subtotal, shipFee, grandTotal, payStatus, bagStatus, lastOrder);
    bagsList.push(bagObj);

    if (cid) processedKeys[cid] = true;
    if (tiktok) processedKeys[tiktok.toLowerCase()] = true;
  });

  for (var key in ordersByCust) {
    if (processedKeys[key]) continue;
    var ords = ordersByCust[key];
    if (!ords || ords.length === 0) continue;

    var firstOrd = ords[0];
    var cid = String(firstOrd['รหัสลูกค้า'] || firstOrd['customerId'] || '').trim();
    var tiktok = String(firstOrd['ชื่อ TikTok'] || firstOrd['tiktokUsername'] || '').trim();
    
    if ((cid && processedKeys[cid]) || (tiktok && processedKeys[tiktok.toLowerCase()])) continue;

    var cInfo = (cid && custMap[cid]) || (tiktok && custMap[tiktok.toLowerCase()]) || {};
    var custName = String(firstOrd['ชื่อผู้รับ'] || firstOrd['ชื่อลูกค้า'] || cInfo['ชื่อ-นามสกุล'] || cInfo['customerName'] || '-').trim();
    var phone = String(firstOrd['เบอร์โทร'] || cInfo['เบอร์โทร'] || cInfo['phone'] || '-').trim();
    var address = String(firstOrd['ที่อยู่'] || cInfo['ที่อยู่จัดส่ง'] || cInfo['address'] || '-').trim();

    var ordSub = 0;
    var ordItemsList = [];
    var ordQty = 0;
    var hasUnpaid = false;
    var latestDate = '';

    ords.forEach(function(mo) {
      ordSub += Number(mo['ยอดสินค้า'] || mo['subtotal'] || mo['ยอดสุทธิ'] || 0);
      var pStatus = String(mo['สถานะออเดอร์'] || mo['paymentStatus'] || mo['status'] || '').toLowerCase();
      if (pStatus.indexOf('paid') === -1 && pStatus.indexOf('ชำระแล้ว') === -1 && pStatus.indexOf('โอนแล้ว') === -1) {
        hasUnpaid = true;
      }
      var d = String(mo['วันที่'] || mo['date'] || '');
      if (d && (!latestDate || d > latestDate)) latestDate = d;

      var iDesc = String(mo['รายการสินค้า'] || mo['itemsDesc'] || mo['items'] || '').trim();
      if (iDesc) {
        ordItemsList.push(iDesc);
        var parts = iDesc.split(/[\n,]+/);
        parts.forEach(function(pt) {
          var m = pt.match(/x\s*(\d+)/i) || pt.match(/(\d+)\s*(ตัว|ชิ้น)/i);
          ordQty += m ? parseInt(m[1], 10) : 1;
        });
      } else {
        ordQty += 1;
      }
    });

    var qty = ordQty > 0 ? ordQty : 1;
    var shipFee = qty <= 5 ? 40 : 40 + ((qty - 5) * 5);
    var grandTotal = ordSub + shipFee;
    var payStatus = hasUnpaid ? 'รอโอน' : 'ชำระแล้ว';
    var itemsDesc = ordItemsList.join('\n');

    var bagObj = buildBagObject(cid, tiktok, custName, phone, address, itemsDesc, qty, ordSub, shipFee, grandTotal, payStatus, 'กำลังฝาก', latestDate);
    bagsList.push(bagObj);

    if (cid) processedKeys[cid] = true;
    if (tiktok) processedKeys[tiktok.toLowerCase()] = true;
  }

  return { success: true, holdingBags: bagsList };
}

function closeHoldingBag(data) {
  var ss = getSpreadsheet();
  var cid = String(data.customerId || '').trim();
  var tiktok = String(data.tiktokUsername || '').trim();
  var name = String(data.customerName || '').trim();
  var fee = Number(data.shippingFee || 40);

  var shpRes = saveShipmentToSheet({
    orderId: cid || tiktok,
    customerId: cid,
    tiktokUsername: tiktok,
    customerName: name,
    phone: data.phone || '',
    address: data.address || '',
    feeCollected: fee,
    actualCost: fee,
    status: 'Waiting',
    notes: 'ถุงฝากสะสม ' + (data.itemsCount || 1) + ' ตัว'
  });

  var orderSheet = findSheet(ss, ["บิลคำสั่งซื้อ", "orders"]);
  if (orderSheet) {
    var oData = orderSheet.getDataRange().getValues();
    var oHeaders = oData[0];
    var bagCol = -1, shipCol = -1, cCol = -1, tkCol = -1;
    for (var c = 0; c < oHeaders.length; c++) {
      var h = String(oHeaders[c]).trim();
      if (h === "ฝากถุง") bagCol = c;
      if (h === "สถานะจัดส่ง") shipCol = c;
      if (h === "รหัสลูกค้า") cCol = c;
      if (h === "ชื่อ TikTok") tkCol = c;
    }
    for (var r = 1; r < oData.length; r++) {
      var matchC = (cCol >= 0 && cid && String(oData[r][cCol]).trim() === cid);
      var matchTk = (tkCol >= 0 && tiktok && String(oData[r][tkCol]).trim().toLowerCase() === tiktok.toLowerCase());
      if ((matchC || matchTk) && bagCol >= 0) {
        if (String(oData[r][bagCol]).indexOf('ฝาก') !== -1) {
          orderSheet.getRange(r + 1, bagCol + 1).setValue("ปิดถุงแล้ว (ส่งเสาร์)");
          if (shipCol >= 0) orderSheet.getRange(r + 1, shipCol + 1).setValue("Waiting");
        }
      }
    }
  }

  var bagSheet = findSheet(ss, ["ตะกร้าสินค้าและฝากถุง", "holding bags", "ฝากถุง"]);
  if (bagSheet) {
    var bData = bagSheet.getDataRange().getValues();
    var bHeaders = bData[0];
    var statusCol = -1;
    for (var c = 0; c < bHeaders.length; c++) {
      if (String(bHeaders[c]).indexOf('สถานะ') !== -1) statusCol = c;
    }
    for (var r = 1; r < bData.length; r++) {
      var matchC = (cid && String(bData[r][0]).trim() === cid);
      var matchTk = (tiktok && String(bData[r][1]).trim().toLowerCase() === tiktok.toLowerCase());
      if (matchC || matchTk) {
        if (statusCol >= 0) {
          bagSheet.getRange(r + 1, statusCol + 1).setValue("ปิดถุงส่งแล้ว");
        }
      }
    }
  }

  return { success: true, shipment: shpRes };
}

function getShipmentsData() {
  var ss = getSpreadsheet();
  var rows = readSheetData(findSheet(ss, ["ประวัติการจัดส่งพัสดุ", "shipments", "จัดส่งพัสดุ"]));
  return { success: true, shipments: rows };
}

function saveShipmentToSheet(shp) {
  var ss = getSpreadsheet();
  var sheet = findSheet(ss, ["ประวัติการจัดส่งพัสดุ", "shipments", "จัดส่งพัสดุ"]);
  var headers = ["รหัสพัสดุ", "วันที่จัดส่ง", "รหัสออเดอร์ที่ส่ง", "รหัสลูกค้า", "ชื่อ TikTok", "ชื่อผู้รับ", "เบอร์โทร", "ที่อยู่จัดส่ง", "บริษัทขนส่ง", "เลขพัสดุ (Tracking)", "ค่าส่งเก็บจากลูกค้า", "ค่าส่งจริงตามบิล", "กำไร/ส่วนต่างค่าส่ง", "สถานะจัดส่ง", "หมายเหตุ", "อัปเดตล่าสุด"];
  
  if (!sheet) {
    sheet = ss.insertSheet("3. ประวัติการจัดส่งพัสดุ (Shipments)");
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setBackground("#1e293b").setFontColor("#ffffff").setFontWeight("bold");
    sheet.setFrozenRows(1);
  }

  var data = sheet.getDataRange().getValues();
  var foundRow = -1;
  var shpId = String(shp.id || '').trim();

  if (!shpId) {
    var maxNum = 0;
    for (var i = 1; i < data.length; i++) {
      var m = String(data[i][0]).trim().match(/^SHP-(\d+)$/i);
      if (m) maxNum = Math.max(maxNum, parseInt(m[1], 10));
    }
    shpId = "SHP-" + String(maxNum + 1).padStart(3, '0');
    shp.id = shpId;
  }

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === shpId) { foundRow = i + 1; break; }
  }

  var collected = Number(shp.feeCollected || shp.shippingFee || 0);
  var actual = Number(shp.actualCost || 0);
  var variance = collected - actual;

  var rowValues = [
    shpId,
    shp.date || Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd"),
    shp.orderId || "-",
    shp.customerId || "-",
    shp.tiktokUsername || "-",
    shp.customerName || "-",
    shp.phone ? "'" + shp.phone : "-",
    shp.address || "-",
    shp.company || "Flash Express",
    shp.trackingNumber ? "'" + shp.trackingNumber : "-",
    collected,
    actual,
    variance,
    shp.status || "Waiting",
    shp.notes || "",
    Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm")
  ];

  if (foundRow > 0) sheet.getRange(foundRow, 1, 1, rowValues.length).setValues([rowValues]);
  else sheet.appendRow(rowValues);

  return { success: true, shipment: shp };
}

/* ==========================================================================
   5. PAYMENTS API (ระบบการเงิน & บันทึกการชำระเงิน)
   ========================================================================== */
function getPaymentsData() {
  try {
    var ss = getSpreadsheet();
    var pSheet = findSheet(ss, ["ประวัติการชำระเงิน", "payments", "การเงิน", "ชำระเงิน"]);
    var rows = pSheet ? readSheetData(pSheet) : [];

    // Filter out completely blank rows
    var validRows = rows.filter(function(r) {
      return !!(r['รหัสชำระเงิน'] || r['รหัสการชำระ'] || r['รหัสชำระ'] || r['id'] || r['รหัสออเดอร์'] || r['orderId'] || r['จำนวนเงิน'] || r['amount'] || r['ยอดชำระ']);
    });

    if (validRows.length > 0) {
      var normalized = validRows.map(function(r, idx) {
        var rawId = r['รหัสชำระเงิน'] || r['รหัสการชำระ'] || r['รหัสชำระ'] || r['id'] || ('PAY-' + String(idx + 1).padStart(3, '0'));
        var id = String(rawId).trim();
        var ordId = String(r['รหัสออเดอร์'] || r['รหัสคำสั่งซื้อ'] || r['orderId'] || '-').trim();
        var date = r['วันที่'] || r['วันที่ชำระ'] || r['date'] || Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");
        var custName = String(r['ชื่อลูกค้า'] || r['ชื่อผู้รับ'] || r['customerName'] || '-').trim();
        var tiktok = String(r['ชื่อ TikTok'] || r['tiktokUsername'] || '-').trim();
        var amt = Number(r['จำนวนเงิน'] || r['ยอดชำระ'] || r['ยอดเงิน'] || r['ยอดสุทธิ'] || r['ยอดรวม'] || r['amount'] || 0);
        var method = String(r['วิธีชำระ'] || r['ช่องทาง'] || r['ช่องทางการชำระเงิน'] || r['method'] || 'PromptPay').trim();
        var status = String(r['สถานะ'] || r['สถานะการชำระ'] || r['status'] || 'Paid').trim();
        var notes = String(r['หมายเหตุ'] || r['notes'] || '').trim();

        return {
          'รหัสชำระเงิน': id,
          'id': id,
          'วันที่': date,
          'date': date,
          'รหัสออเดอร์': ordId,
          'orderId': ordId,
          'ชื่อลูกค้า': custName,
          'customerName': custName,
          'ชื่อ TikTok': tiktok,
          'tiktokUsername': tiktok,
          'จำนวนเงิน': amt,
          'amount': amt,
          'ยอดชำระ': amt,
          'วิธีชำระ': method,
          'method': method,
          'สถานะ': status,
          'status': status,
          'หมายเหตุ': notes,
          'notes': notes
        };
      });
      return { success: true, payments: normalized, source: 'sheet' };
    }

    // Fallback: derive payments directly from Orders
    var oSheet = findSheet(ss, ["บิลคำสั่งซื้อ", "orders"]);
    var orders = oSheet ? readSheetData(oSheet) : [];
    var payments = [];

    orders.forEach(function(o, idx) {
      var rawOrdId = o['รหัสออเดอร์'] || o['id'] || o['orderId'] || ('ORD-' + (idx + 1));
      var ordIdStr = String(rawOrdId).trim();
      var cleanNum = ordIdStr.replace(/^ORD-?/i, '');
      var payId = 'PAY-' + (cleanNum || String(idx + 1).padStart(3, '0'));

      var grand = Number(o['ยอดสุทธิ'] || o['ยอดรวม'] || o['ยอดเงิน'] || o['ยอดชำระ'] || o['grandTotal'] || o['ยอดสินค้า'] || o['subtotal'] || o['total'] || 0);
      var st = String(o['สถานะออเดอร์'] || o['status'] || o['สถานะ'] || 'Confirmed').toLowerCase();
      var payStatus = 'Paid';
      if (st.indexOf('cancel') !== -1 || st.indexOf('ยกเลิก') !== -1) payStatus = 'Refunded';
      else if (st.indexOf('pending') !== -1 || st.indexOf('รอ') !== -1 || st.indexOf('unpaid') !== -1) payStatus = 'Pending';

      var cust = String(o['ชื่อผู้รับ'] || o['ชื่อลูกค้า'] || o['customerName'] || '-').trim();
      var tk = String(o['ชื่อ TikTok'] || o['tiktokUsername'] || '-').trim();
      var dt = o['วันที่'] || o['date'] || Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");
      var nt = o['หมายเหตุ'] || ('ชำระผ่านบิล ' + ordIdStr);

      payments.push({
        'รหัสชำระเงิน': payId,
        'id': payId,
        'วันที่': dt,
        'date': dt,
        'รหัสออเดอร์': ordIdStr,
        'orderId': ordIdStr,
        'ชื่อลูกค้า': cust,
        'customerName': cust,
        'ชื่อ TikTok': tk,
        'tiktokUsername': tk,
        'จำนวนเงิน': grand,
        'amount': grand,
        'ยอดชำระ': grand,
        'วิธีชำระ': 'PromptPay / โอนเงิน',
        'method': 'PromptPay',
        'สถานะ': payStatus,
        'status': payStatus,
        'หมายเหตุ': nt,
        'notes': nt
      });
    });

    return { success: true, payments: payments, source: 'derived_from_orders' };
  } catch (err) {
    return { success: false, message: err.toString(), payments: [] };
  }
}

function savePaymentToSheet(p) {
  var ss = getSpreadsheet();
  var sheet = findSheet(ss, ["ประวัติการชำระเงิน", "payments", "การเงิน", "ชำระเงิน"]);
  var headers = ["รหัสชำระเงิน", "วันที่", "รหัสออเดอร์", "ชื่อลูกค้า", "ชื่อ TikTok", "จำนวนเงิน", "วิธีชำระ", "สถานะ", "หมายเหตุ", "อัปเดตล่าสุด"];

  if (!sheet) {
    sheet = ss.insertSheet("4. ประวัติการชำระเงิน (Payments)");
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setBackground("#1e293b").setFontColor("#ffffff").setFontWeight("bold");
    sheet.setFrozenRows(1);
  }

  var data = sheet.getDataRange().getValues();
  var foundRow = -1;
  var payId = String(p['รหัสชำระเงิน'] || p.id || '').trim();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === payId) { foundRow = i + 1; break; }
  }

  var rowValues = [
    payId,
    p['วันที่'] || p.date || Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd"),
    p['รหัสออเดอร์'] || p.orderId || '-',
    p['ชื่อลูกค้า'] || p.customerName || '-',
    p['ชื่อ TikTok'] || p.tiktokUsername || '-',
    Number(p['จำนวนเงิน'] || p.amount || 0),
    p['วิธีชำระ'] || p.method || 'PromptPay',
    p['สถานะ'] || p.status || 'Paid',
    p['หมายเหตุ'] || p.notes || '',
    Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm")
  ];

  if (foundRow > 0) sheet.getRange(foundRow, 1, 1, rowValues.length).setValues([rowValues]);
  else sheet.appendRow(rowValues);

  return { success: true, payment: p };
}

/* ==========================================================================
   6. DASHBOARD & SYSTEM MASTER DATA
   ========================================================================== */
function getInitialData() {
  try {
    var pData = [];
    try { pData = getPaymentsData().payments || []; } catch(e) { pData = []; }

    return {
      success: true,
      products: getProductsData().products || [],
      liveSessions: getLiveSessionsData().liveSessions || [],
      customers: getCustomersData().customers || [],
      orders: getOrdersData().orders || [],
      holdingBags: getHoldingBagsData().holdingBags || [],
      shipments: getShipmentsData().shipments || [],
      payments: pData,
      timestamp: Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss")
    };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}