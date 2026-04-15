// =================================================================================
// PACKAGE MANAGEMENT SYSTEM — Code.gs (TEMPLATE VERSION)
// =================================================================================

// =================================================================================
// CONFIGURATION (ดึงข้อมูลจาก Script Properties เพื่อความปลอดภัย)
// =================================================================================
const _props = PropertiesService.getScriptProperties();
const CONFIG = {
  'ORG_NAME': _props.getProperty('ORG_NAME') || 'ชื่อองค์กร/บริษัท/คอนโด ของลูกค้า', 
  
  'SHEET_ID': _props.getProperty('SHEET_ID') || '',
  'CHANNEL_ACCESS_TOKEN': _props.getProperty('CHANNEL_ACCESS_TOKEN') || '',
  
  'ADMIN_LOGIN_SECRET': _props.getProperty('ADMIN_LOGIN_SECRET') || 'ADMIN007',
  'MAIN_ADMIN_SECRET': _props.getProperty('MAIN_ADMIN_SECRET') || 'SUPERADMIN2025',
  
  'WEB_APP_URL': _props.getProperty('WEB_APP_URL') || '',
  'DRIVE_FOLDER_NAME': _props.getProperty('DRIVE_FOLDER_NAME') || 'PackagePhotos',

  // Telegram Bot
  'TELEGRAM_BOT_TOKEN':    _props.getProperty('TELEGRAM_BOT_TOKEN') || '',
  'TELEGRAM_BOT_USERNAME': _props.getProperty('TELEGRAM_BOT_USERNAME') || '',
  'TELEGRAM_CHAT_ID':      _props.getProperty('TELEGRAM_CHAT_ID') || '',

  // LIFF URLs
  'LIFF_REGISTER_URL': _props.getProperty('LIFF_REGISTER_URL') || '', 
  'LIFF_CHECKIN_URL':  _props.getProperty('LIFF_CHECKIN_URL') || '',
  'LIFF_PICKUP_URL':   _props.getProperty('LIFF_PICKUP_URL') || '',

  'TIMEZONE': _props.getProperty('TIMEZONE') || 'Asia/Bangkok',
};


// =================================================================================
// SHEET MANAGEMENT
// =================================================================================
const aSheet            = SpreadsheetApp.openById(CONFIG.SHEET_ID);
const subscribersSheet  = aSheet.getSheetByName('Subscribers') || createSheet('Subscribers');
const packagesSheet     = aSheet.getSheetByName('Packages')    || createSheet('Packages');
const adminLogSheet     = aSheet.getSheetByName('AdminLog')    || createSheet('AdminLog');
const tgSheet           = aSheet.getSheetByName('Subscribers_TG') || createTGSheet();

// =================================================================================
// PERFORMANCE: Sheet data cache
// =================================================================================
const _sheetCache = {};
function getSheetData(sheet) {
  const name = sheet.getName();
  if (!_sheetCache[name]) _sheetCache[name] = sheet.getDataRange().getValues();
  return _sheetCache[name];
}
function clearSheetCache(sheet) {
  if (sheet) delete _sheetCache[sheet.getName()];
  else Object.keys(_sheetCache).forEach(k => delete _sheetCache[k]);
}

// =================================================================================
// KEEP-ALIVE
// =================================================================================
function keepAlive() {
  Logger.log('keepAlive ping ' + new Date().toISOString());
}
function setupKeepAliveTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'keepAlive') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('keepAlive').timeBased().everyMinutes(25).create();
}

function createTGSheet() {
  const s = aSheet.insertSheet('Subscribers_TG');
  s.getRange(1,1,1,6).setValues([['Phone','LINE_UserId','OwnerName','TG_ChatId','TG_Username','Timestamp']]);
  s.getRange(1,1,1,6).setBackground('#1a1a2e').setFontColor('#f97316').setFontWeight('bold');
  s.setFrozenRows(1);
  return s;
}

function createSheet(sheetName) {
  const sheet = aSheet.insertSheet(sheetName);
  const headers = {
    'Subscribers': [['Phone', 'UserId', 'OwnerName', 'Timestamp']],
    'Packages':    [['PackageId', 'TrackingNumber', 'PhoneNumberOnLabel', 'RecipientNameOnLabel',
                     'Status', 'CheckInTimestamp', 'CheckOutTimestamp', 'Carrier',
                     'Notes', 'PackageType', 'AdminUser', 'PhotoUrl', 'SignatureUrl']],
    'AdminLog':    [['UserId', 'LoginTimestamp', 'UserName', 'Status', 'AdminType']]
  };
  if (headers[sheetName]) {
    sheet.getRange(1,1,1,headers[sheetName][0].length).setValues(headers[sheetName]);
    sheet.getRange(1,1,1,headers[sheetName][0].length).setBackground('#1a1a2e').setFontColor('#f59e0b').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// =================================================================================
// HTTP REQUEST HANDLERS
// =================================================================================
function doGet(e) {
  if (e.parameter.page === 'debug') return HtmlService.createHtmlOutput(getDebugInfo());
  if (e.parameter.action === 'check_admin') return jsonResponse(checkAdminForLiff(e.parameter.userId));
  if (e.parameter.action === 'link_tg') return handleLinkTelegramGet(e.parameter.phone, e.parameter.chatId, e.parameter.username);
  return HtmlService.createHtmlOutput(`<h2>📦 Package System : ${CONFIG.ORG_NAME} — Online ✅</h2>`);
}

function doPost(e) {
  try {
    if (!e.postData || !e.postData.contents) return jsonResponse({ status:'error', message:'No post data' });
    let data;
    try { data = JSON.parse(e.postData.contents); }
    catch(_) { data = JSON.parse(e.parameter.payload || e.postData.contents); }

    if (data.action === 'decode_barcode' || data.action === 'ocr_barcode') return decodeBarcodeImage(data);
    if (data.action === 'ocr_label')             return ocrParcelLabel(data);
    if (data.action === 'get_ocr_key')           return handleGetOcrKey(data);
    if (data.action === 'liff_submit')           return handleLiffSubmit(data);
    if (data.action === 'liff_register')         return handleLiffRegister(data);     
    if (data.action === 'get_registration')      return handleGetRegistration(data);   
    if (data.action === 'upload_package_photo')  return handleUploadPackagePhoto(data);
    if (data.action === 'lookup_package')        return handleLookupPackage(data);
    if (data.action === 'pickup_package')        return handlePickupPackage(data);
    if (data.action === 'link_telegram')         return handleLinkTelegram(data);
    if (data.action === 'list_waiting_packages') return handleListWaitingPackages(data);
    if (data.action === 'check_tg_linked')       return handleCheckTGLinked(data);

    // Telegram Bot Webhook
    if (data.update_id !== undefined) {
      if (data.edited_message) return jsonResponse({ status:'ok' });
      const lock = LockService.getScriptLock();
      try { lock.tryLock(5000); } catch(_) { return jsonResponse({ status:'ok' }); }
      const tgCache = CacheService.getScriptCache();
      const tgKey   = 'tgupd_' + data.update_id;
      if (tgCache.get(tgKey)) { lock.releaseLock(); return jsonResponse({ status:'ok' }); }
      tgCache.put(tgKey, '1', 3600);
      lock.releaseLock();
      handleTelegramUpdate(data);
      return jsonResponse({ status:'ok' });
    }

    // LINE Webhook
    if (data.events && data.events.length > 0) {
      const cache = CacheService.getScriptCache();
      for (const event of data.events) {
        const dedupKey = event.webhookEventId || (event.message && event.message.id) || String(event.timestamp);
        if (dedupKey) {
          const cacheKey = 'evt_' + dedupKey;
          if (cache.get(cacheKey)) continue;
          cache.put(cacheKey, '1', 60);
        }
        handleWebhook(event);
      }
    }
    return jsonResponse({ status:'ok' });
  } catch (error) {
    return jsonResponse({ status:'error', message:`Server Error: ${error.message}` });
  }
}

function jsonResponse(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }

// =================================================================================
// LIFF AUTH CHECK
// =================================================================================
function checkAdminForLiff(userId) {
  try {
    if (!userId) return { isAdmin:false, reason:'ไม่พบข้อมูลผู้ใช้' };
    const data = adminLogSheet.getDataRange().getValues();
    const now  = new Date(), TWO_HOURS = 2*60*60*1000;
    for (let i=data.length-1; i>0; i--) {
      const [loggedUserId, loginTime, userName, status, adminType] = data[i];
      if (loggedUserId===userId && status==='ACTIVE') {
        if (adminType==='STAFF' && (now-new Date(loginTime))>TWO_HOURS) {
          adminLogSheet.getRange(i+1,4).setValue('INACTIVE');
          return { isAdmin:false, reason:'เซสชันหมดอายุ (2 ชั่วโมง)\nกรุณาพิมพ์รหัสลับใน LINE Chat ใหม่' };
        }
        return { isAdmin:true, adminType:adminType||'STAFF', userName };
      }
    }
    return { isAdmin:false, reason:'คุณยังไม่ได้เข้าระบบเจ้าหน้าที่\nกรุณาพิมพ์รหัสลับใน LINE Chat ก่อน' };
  } catch(e) { return { isAdmin:false, reason:'เกิดข้อผิดพลาดในระบบ: '+e.message }; }
}

// =================================================================================
// LIFF REGISTER 
// =================================================================================
function handleLiffRegister(data) {
  try {
    const { userId, ownerName, phone } = data;
    if (!userId || !ownerName || !phone) return jsonResponse({ status: 'error', message: 'ข้อมูลไม่ครบ กรุณากรอกให้ครบทุกช่อง' });
    const cleanPhone = String(phone).replace(/\D/g, '');
    if (!/^0[689]\d{8}$/.test(cleanPhone)) return jsonResponse({ status: 'error', message: 'เบอร์โทรไม่ถูกต้อง (ต้องมี 10 หลัก)' });
    const cleanName = String(ownerName).trim();
    if (cleanName.length < 2) return jsonResponse({ status: 'error', message: 'กรุณาระบุชื่อ-สกุล' });

    const rows = subscribersSheet.getDataRange().getValues();
    let foundRow = -1;
    for (let i = 1; i < rows.length; i++) { if (String(rows[i][0]).replace(/\D/g, '') === cleanPhone) { foundRow = i + 1; break; } }
    if (foundRow === -1) { for (let i = 1; i < rows.length; i++) { if (rows[i][1] === userId) { foundRow = i + 1; break; } } }

    if (foundRow > -1) {
      subscribersSheet.getRange(foundRow, 1).setValue(`'${cleanPhone}`);
      subscribersSheet.getRange(foundRow, 2).setValue(userId);
      subscribersSheet.getRange(foundRow, 3).setValue(cleanName);
      subscribersSheet.getRange(foundRow, 4).setValue(new Date());
    } else {
      subscribersSheet.appendRow([`'${cleanPhone}`, userId, cleanName, new Date()]);
    }
    return jsonResponse({ status: 'success', result: { ownerName: cleanName, phone: cleanPhone } });
  } catch (e) { return jsonResponse({ status: 'error', message: e.message }); }
}

function handleGetRegistration(data) {
  try {
    const { userId } = data;
    if (!userId) return jsonResponse({ status: 'error', message: 'ไม่พบ userId' });
    const rows = subscribersSheet.getDataRange().getValues();
    for (let i = rows.length - 1; i > 0; i--) {
      if (rows[i][1] === userId) return jsonResponse({ status: 'ok', result: { phone: String(rows[i][0]).replace(/'/g, ''), ownerName: String(rows[i][2]) } });
    }
    return jsonResponse({ status: 'ok', result: null });
  } catch (e) { return jsonResponse({ status: 'error', message: e.message }); }
}

// =================================================================================
// LOOKUP / LIST PACKAGES
// =================================================================================
function handleLookupPackage(data) {
  try {
    const { identifier, userId } = data;
    const auth = checkAdminForLiff(userId);
    if (!auth.isAdmin) return jsonResponse({ status:'error', message:auth.reason });
    if (!identifier)   return jsonResponse({ status:'error', message:'กรุณาระบุรหัสหรือเลขพัสดุ' });

    const id   = String(identifier).trim().replace(/'/g,'').toLowerCase();
    const rows = packagesSheet.getDataRange().getValues();
    for (let i=rows.length-1; i>0; i--) {
      const pkgId  = String(rows[i][0]).replace(/'/g,'').toLowerCase();
      const trkNum = String(rows[i][1]).toLowerCase();
      if (pkgId===id || trkNum===id) {
        const phone  = String(rows[i][2]).replace(/'/g,'');
        const tgInfo = getTGChatIdByPhone(phone);
        return jsonResponse({
          status:'ok',
          result: {
            packageId:     rows[i][0],
            trackingNumber:rows[i][1],
            phone,
            recipientName: rows[i][3],
            status:        rows[i][4],
            checkIn:       rows[i][5],
            packageType:   rows[i][9],
            photoUrl:      rows[i][11]||'',
            tgChatId:      tgInfo ? tgInfo.chatId : '',
            tgUsername:    tgInfo ? tgInfo.username : '',
          }
        });
      }
    }
    return jsonResponse({ status:'error', message:`ไม่พบพัสดุ "${identifier}"` });
  } catch(e) { return jsonResponse({ status:'error', message:e.message }); }
}

function handleListWaitingPackages(data) {
  try {
    const { userId, search, page } = data;
    const auth = checkAdminForLiff(userId);
    if (!auth.isAdmin) return jsonResponse({ status:'error', message:auth.reason });

    const rows     = packagesSheet.getDataRange().getValues();
    const q        = (search||'').trim().toLowerCase();
    const PER_PAGE = 20;
    const pageNum  = Math.max(1, parseInt(page||1, 10));
    const waiting = [];
    
    for (let i = rows.length-1; i > 0; i--) {
      if (rows[i][4] !== 'รอรับ') continue;
      const pkgId   = String(rows[i][0]).replace(/'/g,'');
      const trkNum  = String(rows[i][1]);
      const phone   = String(rows[i][2]).replace(/'/g,'');
      const name    = String(rows[i][3]);
      const pkgType = String(rows[i][9]||'');
      const checkIn = String(rows[i][5]||'');
      if (q && !pkgId.toLowerCase().includes(q) && !trkNum.toLowerCase().includes(q) &&
          !name.toLowerCase().includes(q) && !phone.includes(q) && !pkgType.includes(q)) continue;
      waiting.push({ packageId:pkgId, trackingNumber:trkNum, phone, recipientName:name, packageType:pkgType, checkIn:checkIn.substring(0,16) });
    }

    const total      = waiting.length;
    const totalPages = Math.max(1, Math.ceil(total / PER_PAGE));
    const start      = (pageNum-1) * PER_PAGE;
    const items      = waiting.slice(start, start + PER_PAGE);

    return jsonResponse({ status:'ok', result:{ items, total, page:pageNum, totalPages } });
  } catch(e) { return jsonResponse({ status:'error', message:e.message }); }
}

// =================================================================================
// TELEGRAM UTILS
// =================================================================================
function handleCheckTGLinked(data) {
  try {
    const { phone } = data;
    if (!phone) return jsonResponse({ status: 'error', message: 'ไม่พบเบอร์' });
    const tgInfo = getTGChatIdByPhone(phone);
    return jsonResponse({ status: 'ok', result: { linked: !!(tgInfo && tgInfo.chatId && tgInfo.chatId.length > 3), tgUsername: tgInfo ? tgInfo.username : '' } });
  } catch(e) { return jsonResponse({ status: 'error', message: e.message }); }
}

function tgBotDeepLinkPhone(chatId, username, firstName, phone) {
  try {
    const _cache   = CacheService.getScriptCache();
    const _coolKey = 'deeplink_cd_' + chatId + '_' + phone;
    if (_cache.get(_coolKey)) return;
    _cache.put(_coolKey, '1', 120);

    if (!/^0[689]\d{8}$/.test(phone)) {
      telegramSendDirectMessage(chatId, '❌ ลิงก์ไม่ถูกต้อง\nกรุณาลองใหม่จาก LINE หรือพิมพ์ /start'); return;
    }

    let foundName = '';
    const subRows = subscribersSheet.getDataRange().getValues();
    for (let i = 1; i < subRows.length; i++) {
      if (String(subRows[i][0]).replace(/\D/g, '') === phone) { foundName = subRows[i][2] || ''; break; }
    }

    const isUpdate = isChatIdLinked(chatId);
    saveTGLink(phone, '', foundName, chatId, username);

    const greet    = firstName ? `คุณ${firstName}` : 'คุณ';
    const action   = isUpdate ? 'อัปเดต' : 'ผูก';
    const nameLine = foundName ? `👤 ชื่อ: <b>${foundName}</b>\n` : '';

    telegramSendDirectMessage(chatId,
      `✅ ${action} Telegram สำเร็จ! สวัสดี ${greet}\n\n` +
      `📞 เบอร์: <code>${phone}</code>\n` + nameLine +
      `\n📦 ระบบจะแจ้งเตือนเมื่อ:\n• มีพัสดุเข้าใหม่\n• พัสดุถูกจ่ายออกแล้ว\n\n` +
      `/mypackages — ดูพัสดุที่รอรับ\n/status — ตรวจสอบข้อมูล`);
  } catch(e) { telegramSendDirectMessage(chatId, '⚠️ เกิดข้อผิดพลาด กรุณาลองใหม่\n\nพิมพ์ /start'); }
}

function getTGChatIdByPhone(phone) {
  const cleanPhone = String(phone).replace(/'/g,'').replace(/\D/g,'');
  if (!cleanPhone || cleanPhone.length < 4) return null;
  const rows = getSheetData(tgSheet);
  for (let i=1; i<rows.length; i++) {
    const rowPhone = String(rows[i][0]).replace(/'/g,'').replace(/\D/g,'');
    if (rowPhone === cleanPhone) return { chatId:String(rows[i][3]), username:String(rows[i][4]) };
  }
  return null;
}

function getTGChatIdFuzzy(phone, recipientName) {
  const cleanPhone = String(phone).replace(/'/g,'').replace(/\D/g,'');
  const tgRows     = getSheetData(tgSheet);

  for (let i=1; i<tgRows.length; i++) {
    const rowPhone = String(tgRows[i][0]).replace(/'/g,'').replace(/\D/g,'');
    if (rowPhone === cleanPhone && rowPhone.length >= 9) return { chatId:String(tgRows[i][3]), username:String(tgRows[i][4]), matchedBy:'phone_exact' };
  }

  // Generic prefix stripping for general use
  if (recipientName && recipientName.length >= 2) {
    const stripped = recipientName
      .replace(/นาย|นาง|นางสาว|ด\.ช\.|ด\.ญ\.|คุณ|จ\.ส\.อ\.?|ส\.อ\.?|ร\.อ\.?|ร\.ท\.?|ร\.ต\.?|พ\.อ\.?|พ\.ท\.?|พ\.ต\.?|จ\.ส\.ท\.?|จ\.ส\.ต\.?|พลทหาร|พลฯ/g, '')
      .trim();
    const keywords = stripped.split(/\s+/).filter(function(k) { return k.length >= 2; });

    if (keywords.length > 0) {
      const subRows = getSheetData(subscribersSheet);
      for (let i=1; i<subRows.length; i++) {
        const subName  = String(subRows[i][2] || '');
        const matched  = keywords.some(function(kw) { return subName.includes(kw); });
        if (matched) {
          const subPhone = String(subRows[i][0]).replace(/'/g,'').replace(/\D/g,'');
          for (let j=1; j<tgRows.length; j++) {
            const tp = String(tgRows[j][0]).replace(/'/g,'').replace(/\D/g,'');
            if (tp === subPhone) return { chatId:String(tgRows[j][3]), username:String(tgRows[j][4]), matchedBy:'name_keyword', resolvedPhone:subPhone };
          }
        }
      }
    }
  }
  return null;
}

// =================================================================================
// PICKUP PACKAGE
// =================================================================================
function handlePickupPackage(data) {
  try {
    const { userId, packageId, signatureBase64, tgChatId, adminName } = data;
    const auth = checkAdminForLiff(userId);
    if (!auth.isAdmin) return jsonResponse({ status:'error', message:auth.reason });

    const rows = packagesSheet.getDataRange().getValues();
    let foundRow=-1, pkgInfo={};
    for (let i=rows.length-1; i>0; i--) {
      const id = String(rows[i][0]).replace(/'/g,'');
      if (id===String(packageId).replace(/'/g,'')) {
        if (rows[i][4]!=='รอรับ') return jsonResponse({ status:'error', message:'พัสดุนี้รับไปแล้ว' });
        foundRow=i+1;
        pkgInfo={ trackingNumber:rows[i][1], recipientName:rows[i][3], phone:String(rows[i][2]).replace(/'/g,''), packageType:rows[i][9] };
        break;
      }
    }
    if (foundRow===-1) return jsonResponse({ status:'error', message:'ไม่พบพัสดุ '+packageId });

    const now       = new Date();
    const dateStr   = Utilities.formatDate(now, CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm:ss');
    const dateShort = Utilities.formatDate(now, CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm');

    let sigUrl = '';
    if (signatureBase64) {
      try {
        const iter   = DriveApp.getFoldersByName(CONFIG.DRIVE_FOLDER_NAME);
        const folder = iter.hasNext() ? iter.next() : DriveApp.createFolder(CONFIG.DRIVE_FOLDER_NAME);
        try { folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(_) {}
        const blob = Utilities.newBlob(Utilities.base64Decode(signatureBase64), 'image/png', `SIG_${packageId}_${Date.now()}.png`);
        const file  = folder.createFile(blob);
        try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(_) {}
        sigUrl = `https://drive.google.com/uc?export=view&id=${file.getId()}`;
      } catch(sigErr) {}
    }

    packagesSheet.getRange(foundRow, 5).setValue('รับแล้ว');
    packagesSheet.getRange(foundRow, 7).setValue(dateStr);
    packagesSheet.getRange(foundRow, 13).setValue(sigUrl);
    const notes = packagesSheet.getRange(foundRow, 9).getValue();
    packagesSheet.getRange(foundRow, 9).setValue(notes + ` | จ่ายโดย: ${adminName}`);

    telegramSendMessage(
      `✅ พัสดุถูกรับแล้ว\n━━━━━━━━━━━━━━━━\n` +
      `🔖 รหัส: ${packageId}\n📫 เลขพัสดุ: ${pkgInfo.trackingNumber}\n` +
      `👤 ผู้รับ: ${pkgInfo.recipientName}\n🕐 เวลารับ: ${dateShort}\n` +
      `👨‍✈️ จ่ายโดย: ${adminName}` + (sigUrl ? `\n✍️ ลายเซ็น: ${sigUrl}` : '')
    );

    let tgSent = false;
    const recipientChatId = tgChatId || '';
    if (recipientChatId && recipientChatId.length > 3) {
      try {
        telegramSendDirectMessage(recipientChatId,
          `📬 <b>พัสดุของท่านถูกรับไปแล้ว</b>\n\n` +
          `🔖 รหัส: <code>${packageId}</code>\n` +
          `📫 เลขพัสดุ: <code>${pkgInfo.trackingNumber}</code>\n` +
          `📦 ลักษณะ: ${pkgInfo.packageType||'—'}\n` +
          `🕐 เวลารับ: ${dateShort}\n` +
          `👨‍✈️ จ่ายโดย: ${adminName}\n\n✅ ดำเนินการเรียบร้อยแล้ว`
        );
        tgSent = true;
      } catch(tgErr) {}
    }

    return jsonResponse({ status:'success', result:{ tgSent, sigUrl } });
  } catch(e) { return jsonResponse({ status:'error', message:e.message }); }
}

// =================================================================================
// LINK TELEGRAM
// =================================================================================
function handleLinkTelegramGet(phone, chatId, username) {
  if (!phone || !chatId) return HtmlService.createHtmlOutput('<h3>❌ ข้อมูลไม่ครบถ้วน</h3>');
  const result = saveTGLink(phone, '', '', chatId, username||'');
  if (result) {
    return HtmlService.createHtmlOutput(`
      <html><head><meta charset="UTF-8"><title>ผูก Telegram</title></head>
      <body style="font-family:sans-serif;text-align:center;padding:40px;background:#0d1117;color:#e6edf3">
        <div style="font-size:60px">✅</div><h2 style="color:#f97316">ผูก Telegram สำเร็จ!</h2>
        <p>เบอร์: ${phone}</p><p style="color:#7d8590;font-size:13px">ระบบจะแจ้งเตือนเมื่อมีพัสดุถึงและจ่ายพัสดุแล้ว</p>
      </body></html>
    `);
  }
  return HtmlService.createHtmlOutput('<h3>⚠️ ไม่พบข้อมูลการลงทะเบียน กรุณาลงทะเบียนผ่าน LINE ก่อน</h3>');
}

function handleLinkTelegram(data) {
  try {
    const { userId, phone, tgChatId, tgUsername } = data;
    if (!phone || !tgChatId) return jsonResponse({ status:'error', message:'ข้อมูลไม่ครบ' });
    const ok = saveTGLink(phone, userId||'', '', tgChatId, tgUsername||'');
    if (!ok) return jsonResponse({ status:'error', message:'ไม่พบเบอร์ในระบบ กรุณาลงทะเบียนก่อน' });
    return jsonResponse({ status:'success' });
  } catch(e) { return jsonResponse({ status:'error', message:e.message }); }
}

function saveTGLink(phone, userId, ownerName, chatId, username) {
  const cleanPhone = phone.replace(/\D/g,'');
  if (!cleanPhone || cleanPhone.length < 9) return false;

  const rows = tgSheet.getDataRange().getValues();
  for (let i=1; i<rows.length; i++) {
    if (String(rows[i][0]).replace(/\D/g,'')===cleanPhone) {
      tgSheet.getRange(i+1,3).setValue(ownerName||rows[i][2]);
      tgSheet.getRange(i+1,4).setValue(chatId);
      tgSheet.getRange(i+1,5).setValue(username);
      tgSheet.getRange(i+1,6).setValue(new Date());
      clearSheetCache(tgSheet);
      return true;
    }
  }

  const subRows = subscribersSheet.getDataRange().getValues();
  for (let i=1; i<subRows.length; i++) {
    if (String(subRows[i][0]).replace(/\D/g,'')===cleanPhone) {
      tgSheet.appendRow([`'${cleanPhone}`, userId||subRows[i][1], ownerName||subRows[i][2], chatId, username, new Date()]);
      clearSheetCache(tgSheet);
      return true;
    }
  }

  tgSheet.appendRow([`'${cleanPhone}`, userId||'', ownerName||'', chatId, username, new Date()]);
  clearSheetCache(tgSheet);
  return true;
}

// =================================================================================
// DAILY REPORT
// =================================================================================
function sendDailyReport() {
  try {
    const now      = new Date();
    const todayStr = Utilities.formatDate(now, CONFIG.TIMEZONE, 'dd/MM/yyyy');
    const rows     = packagesSheet.getDataRange().getValues();

    let newIn=0, pickedUp=0, waiting=0;
    const newList=[], pickList=[], waitList=[];

    for (let i=1; i<rows.length; i++) {
      const checkIn  = String(rows[i][5]||'');
      const checkOut = String(rows[i][6]||'');
      const status   = rows[i][4];
      const name     = rows[i][3];
      const pkgId    = rows[i][0];
      const pkgType  = rows[i][9];

      if (checkIn.startsWith(todayStr)) { newIn++; newList.push(`  • ${pkgId} — ${name} (${pkgType||'—'})`); }
      if (checkOut.startsWith(todayStr) && status==='รับแล้ว') { pickedUp++; pickList.push(`  • ${pkgId} — ${name}`); }
      if (status==='รอรับ') { waiting++; waitList.push(`  • ${pkgId} — ${name} | เข้า: ${checkIn.substring(0,16)}`); }
    }

    const MAX_LIST = 15;
    const fmt = (arr) => arr.length===0 ? '  (ไม่มี)' : arr.slice(0,MAX_LIST).join('\n') + (arr.length>MAX_LIST ? `\n  ...และอีก ${arr.length-MAX_LIST} รายการ` : '');

    const report =
      `📊 รายงานสรุปประจำวัน ${todayStr}\n━━━━━━━━━━━━━━━━━━━━━━\n` +
      `📦 พัสดุเข้าวันนี้: ${newIn} ชิ้น\n${fmt(newList)}\n\n` +
      `✅ พัสดุรับออกวันนี้: ${pickedUp} ชิ้น\n${fmt(pickList)}\n\n` +
      `⏳ คงค้างรอรับ: ${waiting} ชิ้น\n${fmt(waitList)}\n\n` +
      `━━━━━━━━━━━━━━━━━━━━━━\n🕐 สรุป ณ ${Utilities.formatDate(now, CONFIG.TIMEZONE, 'HH:mm')} น.\n📋 ${CONFIG.ORG_NAME}`;

    telegramSendMessage(report);
  } catch(e) { telegramSendMessage('⚠️ เกิดข้อผิดพลาดในการส่งรายงาน: '+e.message); }
}

function setupDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'sendDailyReport') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('sendDailyReport').timeBased().atHour(21).everyDays(1).inTimezone(CONFIG.TIMEZONE).create();
}

// =================================================================================
// NOTIFICATIONS
// =================================================================================
function notifyRecipientTelegram(phone, recipientName, packageId, packageType, trackingNumber, photoUrl) {
  try {
    const _cache      = CacheService.getScriptCache();
    const _notifyKey  = 'notify_' + String(packageId).replace(/'/g,'');
    if (_cache.get(_notifyKey)) return;
    _cache.put(_notifyKey, '1', 3600);

    const tgInfo = getTGChatIdFuzzy(phone, recipientName);
    if (!tgInfo || !tgInfo.chatId || tgInfo.chatId.length < 3) return;

    const now     = new Date();
    const dateStr = Utilities.formatDate(now, CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm');
    const msg =
      `📦 <b>มีพัสดุมาถึงแล้ว!</b>\n\n` +
      `👤 ผู้รับ: ${recipientName}\n` +
      `🔖 รหัสรับพัสดุ: <code>${packageId}</code>\n` +
      `📫 เลขพัสดุ: <code>${trackingNumber||'—'}</code>\n` +
      `📦 ลักษณะ: ${packageType||'—'}\n` +
      `🕐 เวลานำเข้า: ${dateStr}\n\n` +
      `📌 <i>กรุณานำ "รหัสรับพัสดุ" แจ้งเจ้าหน้าที่เพื่อรับพัสดุ</i>`;
    telegramSendDirectMessage(tgInfo.chatId, msg);
  } catch(e) {}
}

// =================================================================================
// UPLOAD PHOTO
// =================================================================================
function handleUploadPackagePhoto(data) {
  try {
    const { packageId, userId, imageBase64, imageMime, trackingNumber, recipientName, packageType, adminName } = data;
    if (!imageBase64 || !packageId) return jsonResponse({ status:'ok', photoUrl:'', skipped:true });

    const photoUrl = savePhotoToDrive(imageBase64, imageMime||'image/jpeg', packageId);

    if (photoUrl) {
      try {
        const rows = packagesSheet.getDataRange().getValues();
        for (let i=rows.length-1; i>0; i--) {
          if (String(rows[i][0]).replace(/'/g,'')===String(packageId).replace(/'/g,'')) {
            packagesSheet.getRange(i+1,12).setValue(photoUrl); break;
          }
        }
      } catch(ue) {}

      try {
        const now     = new Date();
        const dateStr = Utilities.formatDate(now, CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm');
        const caption =
          `📦 พัสดุเข้าใหม่\n━━━━━━━━━━━━━━━━\n` +
          `🔖 รหัส: ${packageId}\n📫 เลขพัสดุ: ${trackingNumber||'—'}\n` +
          `👤 ผู้รับ: ${recipientName||'—'}\n📦 ลักษณะ: ${packageType||'—'}\n` +
          `🕐 เวลา: ${dateStr}\n👨‍✈️ บันทึกโดย: ${adminName||'—'}\n━━━━━━━━━━━━━━━━\n🔗 ${photoUrl}`;
        const match = photoUrl.match(/[?&]id=([^&]+)/);
        if (match) telegramSendPhoto(match[1], caption);
        else       telegramSendMessage(caption);
      } catch(tgErr) {}

      try {
        const rows = packagesSheet.getDataRange().getValues();
        let phone='';
        for (let i=rows.length-1; i>0; i--) {
          if (String(rows[i][0]).replace(/'/g,'')===String(packageId).replace(/'/g,'')) {
            phone=String(rows[i][2]).replace(/'/g,''); break;
          }
        }
        if (phone) notifyRecipientTelegram(phone, recipientName, packageId, packageType, trackingNumber, photoUrl);
      } catch(nr) {}
    }
    return jsonResponse({ status:'ok', photoUrl:photoUrl||'' });
  } catch(e) { return jsonResponse({ status:'ok', photoUrl:'', error:e.message }); }
}

function savePhotoToDrive(base64, mimeType, prefix) {
  try {
    const iter   = DriveApp.getFoldersByName(CONFIG.DRIVE_FOLDER_NAME);
    const folder = iter.hasNext() ? iter.next() : DriveApp.createFolder(CONFIG.DRIVE_FOLDER_NAME);
    try { folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(_) {}
    const ext  = mimeType==='image/png' ? '.png' : '.jpg';
    const file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(base64), mimeType, `${prefix}_${Date.now()}${ext}`));
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(_) {}
    return `https://drive.google.com/uc?export=view&id=${file.getId()}`;
  } catch(e) { return ''; }
}

function handleGetOcrKey(data) {
  try {
    const claudeKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
    if (!claudeKey) return jsonResponse({ status:'ok', result:{ key:null, reason:'CLAUDE_API_KEY ยังไม่ได้ตั้งค่าใน Script Properties' } });
    return jsonResponse({ status:'ok', result:{ key:claudeKey } });
  } catch(e) { return jsonResponse({ status:'error', message:e.message }); }
}

// =================================================================================
// OCR / BARCODE
// =================================================================================
function decodeBarcodeImage(data) {
  try {
    const { imageBase64:b64, mimeType:mime='image/jpeg' } = data;
    if (!b64) return jsonResponse({ status:'error', result:null, message:'no image data' });
    try { const v=callVisionAPI(b64,mime); if(v) return jsonResponse({ status:'ok', tracking:v }); } catch(_) {}
    return jsonResponse({ status:'ok', tracking:null, message:'Vision API not configured' });
  } catch(e) { return jsonResponse({ status:'error', tracking:null, message:e.message }); }
}

function callVisionAPI(b64, mime) {
  const apiKey=PropertiesService.getScriptProperties().getProperty('VISION_API_KEY');
  if (!apiKey) throw new Error('VISION_API_KEY not set');
  const url=`https://vision.googleapis.com/v1/images:annotate?key=${apiKey}`;
  const res=UrlFetchApp.fetch(url,{method:'post',contentType:'application/json',payload:JSON.stringify({requests:[{image:{content:b64},features:[{type:'TEXT_DETECTION',maxResults:1}]}]}),muteHttpExceptions:true});
  const r=JSON.parse(res.getContentText());
  if (r.error) throw new Error(r.error.message);
  const ann=r.responses[0].textAnnotations;
  if (!ann||ann.length===0) return null;
  const t=ann[0].description||'';
  for (const pat of [/([A-Z]{2}\d{9}[A-Z]{2})/,/([A-Z]{1,3}\d{10,15})/,/(\d{12,20})/,/([A-Z0-9]{8,25})/]) { const m=t.match(pat); if(m) return m[1]; }
  const fl=t.split('\n')[0].trim(); return fl.length>=5 ? fl : null;
}

function ocrParcelLabel(data) {
  try {
    const { imageBase64:b64, mimeType:mime='image/jpeg' } = data;
    if (!b64) return jsonResponse({ status:'error', result:null, message:'no image' });
    const claudeKey=PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
    if (!claudeKey) return ocrViaVision(b64,mime);
    const prompt=`คุณคือระบบอ่านป้ายพัสดุ ดูรูปภาพแล้วหาข้อมูลต่อไปนี้:
1. tracking: เลขพัสดุ/บาร์โค้ด
2. name: ชื่อผู้รับ 
3. phone: เบอร์โทร 10 หลัก ขึ้นต้นด้วย 0 (ถ้ามี)
4. type: ลักษณะ — เลือก: กล่อง, ซอง, ถุง, ม้วน, อื่นๆ

ตอบด้วย JSON ล้วน ไม่มีข้อความอื่น ถ้าไม่พบให้ใส่ "ไม่พบ":
{"tracking":"...","name":"...","phone":"...","type":"..."}`;
    const res=UrlFetchApp.fetch('https://api.anthropic.com/v1/messages',{
      method:'post',
      headers:{'x-api-key':claudeKey,'anthropic-version':'2023-06-01','content-type':'application/json'},
      payload:JSON.stringify({ model:'claude-haiku-4-5-20251001', max_tokens:300, messages:[{role:'user',content:[{type:'image',source:{type:'base64',media_type:mime,data:b64}},{type:'text',text:prompt}]}] }),
      muteHttpExceptions:true
    });
    const resp=JSON.parse(res.getContentText());
    if (resp.error) throw new Error(resp.error.message);
    const text=(resp.content&&resp.content[0]&&resp.content[0].text)||'';
    const m=text.match(/\{[\s\S]*?\}/);
    if (!m) throw new Error('ไม่พบ JSON ใน response');
    const result=JSON.parse(m[0]);
    if (result.phone&&result.phone!=='ไม่พบ') result.phone=result.phone.replace(/[^0-9]/g,'').slice(0,10);
    return jsonResponse({ status:'ok', result });
  } catch(e) { return jsonResponse({ status:'error', result:null, message:e.message }); }
}

function ocrViaVision(b64,mimeType) {
  try {
    const apiKey=PropertiesService.getScriptProperties().getProperty('VISION_API_KEY');
    if (!apiKey) throw new Error('ไม่พบ API Key');
    const res=UrlFetchApp.fetch(`https://vision.googleapis.com/v1/images:annotate?key=${apiKey}`,{method:'post',contentType:'application/json',payload:JSON.stringify({requests:[{image:{content:b64},features:[{type:'TEXT_DETECTION',maxResults:1}]}]}),muteHttpExceptions:true});
    const r=JSON.parse(res.getContentText());
    if (r.error) throw new Error(r.error.message);
    const ann=r.responses[0].textAnnotations;
    if (!ann||ann.length===0) return jsonResponse({status:'ok',result:{tracking:'ไม่พบ',name:'ไม่พบ',phone:'ไม่พบ',type:'ไม่พบ'}});
    const full=ann[0].description||'';
    let tracking='ไม่พบ';
    for (const pat of [/([A-Z]{2}\d{9}[A-Z]{2})/,/([A-Z]{1,3}\d{10,15})/,/(\d{12,20})/]) { const m2=full.match(pat); if(m2){tracking=m2[1];break;} }
    const phoneMatch=full.match(/0[689]\d{8}/);
    const phone=phoneMatch?phoneMatch[0]:'ไม่พบ';
    let name='ไม่พบ', type='ไม่พบ';
    const lines=full.split('\n');
    for (const l of lines){const t=l.trim();if(/ผู้รับ|ถึง|To:|คุณ|นาย|นาง/i.test(t)){name=t.replace(/ผู้รับ[:\s]*|ถึง[:\s]*|To:[:\s]*/gi,'').trim();if(name.length>2)break;}}
    if (name==='ไม่พบ') for(const l of lines){const t=l.trim();if(t.length>=5&&t.length<=40&&!/^\d+$/.test(t)&&!/[A-Z]{2}\d{9}/.test(t)){name=t;break;}}
    if(/กล่อง|box/i.test(full))type='กล่อง';else if(/ซอง|envelope/i.test(full))type='ซอง';else if(/ถุง|bag/i.test(full))type='ถุง';
    return jsonResponse({status:'ok',result:{tracking,name,phone,type}});
  } catch(e){return jsonResponse({status:'error',result:null,message:e.message});}
}

// =================================================================================
// LIFF SUBMIT
// =================================================================================
function handleLiffSubmit(data) {
  try {
    const { userId, tracking, name, phone, type } = data;
    const auth = checkAdminForLiff(userId);
    if (!auth.isAdmin) return jsonResponse({ status:'error', message:auth.reason });
    const now=new Date();
    const formattedTimestamp=Utilities.formatDate(now,'GMT+7','dd/MM/yyyy HH:mm:ss');
    const d=String(now.getDate()).padStart(2,'0'), m=String(now.getMonth()+1).padStart(2,'0');
    const sy=String(now.getFullYear()+543).slice(-2), h=String(now.getHours()).padStart(2,'0'), mi=String(now.getMinutes()).padStart(2,'0'), s=String(now.getSeconds()).padStart(2,'0');
    const packageId=`${d}${m}${sy}${h}${mi}${s}`;
    if (isDuplicatePackage(tracking)) return jsonResponse({ status:'error', message:`พัสดุ ${tracking} มีในระบบและยังรอรับอยู่` });
    const packageData={PackageId:packageId,TrackingNumber:tracking,PhoneNumberOnLabel:phone,RecipientNameOnLabel:name,Status:'รอรับ',CheckInTimestamp:formattedTimestamp,CheckOutTimestamp:'',Carrier:'ไม่ระบุ',Notes:`บันทึกโดย ${auth.userName} (LIFF)`,PackageType:type,AdminUser:auth.userName,PhotoUrl:'',SignatureUrl:''};
    const saveResult=savePackageFromChat(packageData);
    if (saveResult.status==='success') return jsonResponse({ status:'success', packageId });
    return jsonResponse({ status:'error', message:saveResult.message });
  } catch(e) { return jsonResponse({ status:'error', message:e.message }); }
}

// =================================================================================
// ADMIN MANAGEMENT & LINE CHAT
// =================================================================================
function processAdminLogin(secret, userId, replyToken) {
  try {
    if (!secret||!userId) { replyMessage(replyToken,'เกิดข้อผิดพลาด'); return; }
    let isValid=false, adminType='STAFF';
    if (secret.trim()===CONFIG.ADMIN_LOGIN_SECRET){isValid=true;adminType='STAFF';}
    else if (secret.trim()===CONFIG.MAIN_ADMIN_SECRET){isValid=true;adminType='MAIN_ADMIN';}
    if (!isValid) { replyMessage(replyToken,'❌ รหัสลับไม่ถูกต้อง'); return; }
    if (adminType==='STAFF') {
      const data=adminLogSheet.getDataRange().getValues();
      for(let i=data.length-1;i>0;i--){if(data[i][4]==='STAFF'&&data[i][3]==='ACTIVE'){adminLogSheet.getRange(i+1,4).setValue('INACTIVE');break;}}
    }
    const profile=getUserProfile(userId);
    const userName=profile?profile.displayName:'Unknown User';
    adminLogSheet.appendRow([userId,new Date(),userName,'ACTIVE',adminType]);
    const txt=adminType==='MAIN_ADMIN'?'แอดมินหลัก':'เจ้าหน้าที่';
    const welcomeText=`✅ เข้าระบบ${txt}สำเร็จ!\n\nสวัสดีคุณ ${userName}`;
    if (adminType==='MAIN_ADMIN') sendMainAdminQuickReplyMenu(replyToken,welcomeText);
    else sendStaffQuickReplyMenu(replyToken,welcomeText);
  } catch(e){replyMessage(replyToken,'เกิดข้อผิดพลาดในการเข้าระบบ');}
}

function isAdmin(userId) {
  try {
    if (!userId) return {isAdmin:false,adminType:null};
    const data=adminLogSheet.getDataRange().getValues(),now=new Date(),TWO_HOURS=2*60*60*1000;
    for(let i=data.length-1;i>0;i--){
      const [uid,lt,,st,at]=data[i];
      if(uid===userId&&st==='ACTIVE'){
        if(at==='STAFF'&&(now-new Date(lt))>TWO_HOURS){adminLogSheet.getRange(i+1,4).setValue('INACTIVE');return{isAdmin:false,adminType:null};}
        return{isAdmin:true,adminType:at||'STAFF'};
      }
    }
    return{isAdmin:false,adminType:null};
  } catch(e){return{isAdmin:false,adminType:null};}
}

function handleLogout(replyToken,userId){
  if(isAdmin(userId).isAdmin){
    const d=adminLogSheet.getDataRange().getValues();
    for(let i=d.length-1;i>0;i--){if(d[i][0]===userId&&d[i][3]==='ACTIVE'){adminLogSheet.getRange(i+1,4).setValue('INACTIVE');break;}}
    replyMessage(replyToken,'✅ ออกจากระบบเรียบร้อย');
  }
  else replyMessage(replyToken,'❌ คุณยังไม่ได้เข้าระบบ');
}

function parsePackageFromText(text, userId) {
  try {
    const lines=text.trim().split('\n');
    if(lines.length<4) return{status:'error',message:'รูปแบบข้อมูลไม่ถูกต้อง\nรูปแบบที่ถูกต้อง:\nเลขพัสดุ\nชื่อผู้รับ\nเบอร์โทรศัพท์\nลักษณะพัสดุ'};
    const trackingNumber=lines[0].trim(), recipientName=lines[1].trim(), phoneInput=lines[2].trim(), packageType=lines[3].trim();
    let cleanPhone=phoneInput.replace(/\D/g,'');
    if(cleanPhone.startsWith('66')&&cleanPhone.length===11)cleanPhone='0'+cleanPhone.substring(2);
    const finalPhone=/^0[689]\d{8}$/.test(cleanPhone)?cleanPhone:(phoneInput.toLowerCase()==='ไม่มีเบอร์'?'ไม่มีเบอร์':phoneInput);
    if(!trackingNumber||trackingNumber.length<5) return{status:'error',message:'เลขพัสดุไม่ถูกต้อง'};
    if(!recipientName||recipientName.length<2) return{status:'error',message:'ชื่อผู้รับไม่ถูกต้อง'};
    const now=new Date(), ts=Utilities.formatDate(now,'GMT+7','dd/MM/yyyy HH:mm:ss');
    const d=String(now.getDate()).padStart(2,'0'),m=String(now.getMonth()+1).padStart(2,'0'),sy=String(now.getFullYear()+543).slice(-2),h=String(now.getHours()).padStart(2,'0'),mi=String(now.getMinutes()).padStart(2,'0'),s=String(now.getSeconds()).padStart(2,'0');
    const packageId=`${d}${m}${sy}${h}${mi}${s}`;
    const adminProfile=getUserProfile(userId);
    const adminName=adminProfile?adminProfile.displayName:'เจ้าหน้าที่';
    return{status:'success',data:{PackageId:packageId,TrackingNumber:trackingNumber,PhoneNumberOnLabel:finalPhone,RecipientNameOnLabel:recipientName,Status:'รอรับ',CheckInTimestamp:ts,CheckOutTimestamp:'',Carrier:'ไม่ระบุ',Notes:`บันทึกโดย ${adminName}`,PackageType:packageType,AdminUser:adminName,PhotoUrl:'',SignatureUrl:''}};
  } catch(e){return{status:'error',message:`เกิดข้อผิดพลาด: ${e.message}`};}
}

function savePackageFromChat(packageData) {
  try {
    packagesSheet.appendRow([`'${packageData.PackageId}`, packageData.TrackingNumber, `'${packageData.PhoneNumberOnLabel}`, packageData.RecipientNameOnLabel, packageData.Status, packageData.CheckInTimestamp, packageData.CheckOutTimestamp, packageData.Carrier, packageData.Notes, packageData.PackageType, packageData.AdminUser, packageData.PhotoUrl||'', packageData.SignatureUrl||'']);
    return{status:'success',message:'บันทึกสำเร็จ',packageId:packageData.PackageId};
  } catch(e){return{status:'error',message:`ข้อผิดพลาด: ${e.message}`};}
}

function handleWebhook(event) {
  try {
    if (event.type==='follow') {
      sendUserQuickReplyMenu(event.replyToken, `🏢 ยินดีต้อนรับสู่ระบบพัสดุ\n${CONFIG.ORG_NAME}\n\n📦 กดปุ่ม "ลงทะเบียนรับพัสดุ" ด้านล่างเพื่อเริ่มต้นรับการแจ้งเตือน`);
    } else if (event.type==='message'&&event.message.type==='text') {
      handleTextMessage(event.replyToken, event.source.userId, event.message.text);
    }
  } catch(e){Logger.log('handleWebhook error: '+e.message);}
}

function handleTextMessage(replyToken, userId, text) {
  try {
    const message=text.trim(), adminStatus=isAdmin(userId), lines=message.split('\n');

    if (message.startsWith('✅ ยืนยันการรับ ')) { replyMessage(replyToken,processPackagePickupFromChat(message.substring('✅ ยืนยันการรับ '.length).trim(),userId).message); return; }
    if (message==='❌ ยกเลิกการรับ') { replyMessage(replyToken,'👍 ยกเลิกแล้ว'); return; }
    if (message.toLowerCase().startsWith('ยืนยันการรับ ')) { if(adminStatus.isAdmin)handlePickupConfirmation(replyToken,message.substring('ยืนยันการรับ '.length).trim());else showUnauthorizedMessage(replyToken,'ยืนยันการรับ'); return; }
    if (message.toLowerCase().startsWith('ดูพัสดุ หน้า ')) { if(adminStatus.isAdmin)showAllPackagesWithQuickReply(replyToken,userId,adminStatus.adminType,parseInt(message.substring('ดูพัสดุ หน้า '.length).trim(),10)); return; }
    if (message==='เมนู'||message.toLowerCase()==='menu') {
      if(adminStatus.isAdmin){ if(adminStatus.adminType==='MAIN_ADMIN') sendMainAdminQuickReplyMenu(replyToken,'เมนูแอดมินหลัก'); else sendStaffQuickReplyMenu(replyToken,'เมนูเจ้าหน้าที่'); } 
      else { sendUserQuickReplyMenu(replyToken,'เมนูผู้รับพัสดุ'); }
      return;
    }
    if (message==='พัสดุของฉัน') { handleFindUserPackages(replyToken,userId); return; }
    if (message.toLowerCase().startsWith('ค้นหาชื่อ ')) { if(adminStatus.isAdmin)replyMessage(replyToken,searchByName(message.substring('ค้นหาชื่อ '.length).trim()).message);else showUnauthorizedMessage(replyToken,'ค้นหาชื่อ'); return; }
    if (message.toLowerCase().startsWith('ค้นหาเบอร์ ')) { if(adminStatus.isAdmin)replyMessage(replyToken,searchByPartialPhone(message.substring('ค้นหาเบอร์ '.length).trim()).message);else showUnauthorizedMessage(replyToken,'ค้นหาเบอร์'); return; }
    if (message.toLowerCase().startsWith('ค้นหา ')) { if(adminStatus.isAdmin)replyMessage(replyToken,searchPackage(message.substring('ค้นหา '.length).trim()).message);else showUnauthorizedMessage(replyToken,'ค้นหา'); return; }
    if (message===CONFIG.ADMIN_LOGIN_SECRET||message===CONFIG.MAIN_ADMIN_SECRET) { processAdminLogin(message,userId,replyToken); return; }
    if (message.toLowerCase().startsWith('ผูก tg ') || message.toLowerCase().startsWith('ผูก telegram ')) { handleLinkTGFromChat(replyToken, userId, message); return; }
    if (message.includes('ออกจากระบบ')||message.includes('ออก')) { handleLogout(replyToken,userId); return; }
    if (message.includes('ดูพัสดุ')||message.includes('รายการ')) { if(adminStatus.isAdmin)showAllPackagesWithQuickReply(replyToken,userId,adminStatus.adminType,1);else showUnauthorizedMessage(replyToken,'ดูพัสดุ'); return; }
    if (message.includes('ตรวจสอบ')||message.includes('สถานะ')) { checkSystemStatus(replyToken,userId); return; }
    if (message.includes('สถิติ')) { if(adminStatus.isAdmin&&adminStatus.adminType==='MAIN_ADMIN')showSystemStats(replyToken,userId,adminStatus.adminType);else showUnauthorizedMessage(replyToken,'สถิติ'); return; }
    
    if (!adminStatus.isAdmin&&lines.length===2) {
      const name=lines[0].trim(),phone=lines[1].trim(),cp=phone.replace(/\D/g,'');
      if(/^0[689]\d{8}$/.test(cp)) registerUser(replyToken,userId,name,cp);
      else replyMessage(replyToken,'🤔 รูปแบบไม่ถูกต้อง\n\nกรุณากด "ลงทะเบียนรับพัสดุ" จากเมนูด้านล่าง');
      return;
    }
    if (adminStatus.isAdmin&&lines.length>=4) { handlePackageConfirmation(replyToken,userId,message); return; }

    if (adminStatus.isAdmin) replyMessage(replyToken,'🤔 ไม่เข้าใจคำสั่ง\n\nพิมพ์ `เมนู` เพื่อดูตัวเลือก');
    else sendUserQuickReplyMenu(replyToken,'🤔 ไม่เข้าใจคำสั่ง\n\nกดปุ่มเมนูด้านล่างเพื่อเริ่มใช้งาน');
  } catch(e){Logger.log('handleTextMessage error: '+e.message);replyMessage(replyToken,'เกิดข้อผิดพลาด');}
}

function handleLinkTGFromChat(replyToken, userId, message) {
  try {
    const parts  = message.split(' '); const chatId = parts[parts.length-1].trim();
    if (!chatId || !/^-?\d+$/.test(chatId)) { replyMessage(replyToken, '❌ รูปแบบไม่ถูกต้อง\n\nพิมพ์: ผูก TG <Chat_ID>'); return; }
    const subRows=subscribersSheet.getDataRange().getValues();
    let phone='', ownerName='';
    for(let i=1;i<subRows.length;i++){if(subRows[i][1]===userId){phone=String(subRows[i][0]).replace(/'/g,'');ownerName=subRows[i][2];break;}}
    if (!phone) { replyMessage(replyToken,'❌ ยังไม่ได้ลงทะเบียนในระบบ\n\nกรุณากด "ลงทะเบียนรับพัสดุ" ก่อน'); return; }
    saveTGLink(phone, userId, ownerName, chatId, '');
    try { telegramSendDirectMessage(chatId, `✅ ผูก Telegram สำเร็จ!\n\nชื่อ: ${ownerName}\nเบอร์: ${phone}\n\nระบบจะแจ้งเตือนเมื่อมีพัสดุถึงและจ่ายพัสดุแล้ว 📦`); } catch(_) {}
    replyMessage(replyToken, `✅ ผูก Telegram สำเร็จ!\n\nชื่อ: ${ownerName}\nChat ID: ${chatId}`);
  } catch(e) { replyMessage(replyToken,'เกิดข้อผิดพลาด: '+e.message); }
}

function sendUserQuickReplyMenu(replyToken, welcomeText) {
  replyMessageAdvanced(replyToken, [{
    type: 'text', text: welcomeText,
    quickReply: { items: [
        { type: 'action', action: { type: 'uri', label: '📋 ลงทะเบียนรับพัสดุ', uri: CONFIG.LIFF_REGISTER_URL } },
        { type: 'action', action: { type: 'message', label: '📦 พัสดุของฉัน', text: 'พัสดุของฉัน' } },
        { type: 'action', action: { type: 'message', label: '🔗 ผูก Telegram', text: 'ผูก TG ' } },
      ] }
  }]);
}

function sendStaffQuickReplyMenu(replyToken, welcomeText) {
  replyMessageAdvanced(replyToken,[{type:'text',text:welcomeText,quickReply:{items:[
    {type:'action',action:{type:'message',label:'📋 ดูพัสดุรอรับ',text:'ดูพัสดุ'}},
    {type:'action',action:{type:'uri',label:'📷 บันทึกพัสดุเข้า',uri:CONFIG.LIFF_CHECKIN_URL}},
    {type:'action',action:{type:'uri',label:'📬 จ่ายพัสดุออก',uri:CONFIG.LIFF_PICKUP_URL}},
    {type:'action',action:{type:'message',label:'🔍 ค้นหาชื่อ',text:'ค้นหาชื่อ '}},
    {type:'action',action:{type:'message',label:'🚪 ออกจากระบบ',text:'ออกจากระบบ'}},
  ]}}]);
}

function sendMainAdminQuickReplyMenu(replyToken, welcomeText) {
  replyMessageAdvanced(replyToken,[{type:'text',text:welcomeText,quickReply:{items:[
    {type:'action',action:{type:'message',label:'📋 ดูพัสดุรอรับ',text:'ดูพัสดุ'}},
    {type:'action',action:{type:'uri',label:'📷 บันทึกพัสดุ',uri:CONFIG.LIFF_CHECKIN_URL}},
    {type:'action',action:{type:'uri',label:'📬 จ่ายพัสดุ',uri:CONFIG.LIFF_PICKUP_URL}},
    {type:'action',action:{type:'message',label:'📊 ตรวจสอบสถานะ',text:'ตรวจสอบสถานะ'}},
    {type:'action',action:{type:'message',label:'🚪 ออกจากระบบ',text:'ออกจากระบบ'}},
  ]}}]);
}

function handlePackageConfirmation(replyToken, userId, text) {
  const parseResult=parsePackageFromText(text,userId);
  if (parseResult.status==='error'){replyMessage(replyToken,`⚠️ ${parseResult.message}`);return;}
  const packageData=parseResult.data;
  if (isDuplicatePackage(packageData.TrackingNumber)){replyMessage(replyToken,`⚠️ พัสดุซ้ำ!`);return;}
  const saveResult=savePackageFromChat(packageData);
  if (saveResult.status==='success') {
    const now=new Date();
    telegramSendMessage(`📦 พัสดุเข้าใหม่\n━━━━━━━━━━━━━━━━\n🔖 รหัส: ${packageData.PackageId}\n📫 เลขพัสดุ: ${packageData.TrackingNumber}\n👤 ผู้รับ: ${packageData.RecipientNameOnLabel}\n📞 เบอร์: ${packageData.PhoneNumberOnLabel}\n📦 ลักษณะ: ${packageData.PackageType}`);
    notifyRecipientTelegram(packageData.PhoneNumberOnLabel, packageData.RecipientNameOnLabel, packageData.PackageId, packageData.PackageType, packageData.TrackingNumber, '');
    replyMessage(replyToken,`✅ บันทึกพัสดุเรียบร้อยแล้ว\nรหัส: ${packageData.PackageId}\nผู้รับ: ${packageData.RecipientNameOnLabel}`);
  }
}

function registerUser(replyToken, userId, ownerName, phoneNumber) {
  try {
    const data=subscribersSheet.getDataRange().getValues();
    let foundRow=-1;
    for(let i=1;i<data.length;i++){if(data[i][0]==phoneNumber){foundRow=i+1;break;}}
    if(foundRow>-1){
      subscribersSheet.getRange(foundRow,2).setValue(userId); subscribersSheet.getRange(foundRow,3).setValue(ownerName);
      sendUserQuickReplyMenu(replyToken,`🔄 อัปเดตข้อมูลเบอร์ ${phoneNumber} เรียบร้อยแล้ว`);
    } else {
      subscribersSheet.appendRow([`'${phoneNumber}`,userId,ownerName,new Date()]);
      sendUserQuickReplyMenu(replyToken,`✅ ลงทะเบียนสำเร็จ!\n\nคุณ "${ownerName}"`);
    }
  } catch(e){replyMessage(replyToken,'เกิดข้อผิดพลาดในการลงทะเบียน');}
}

function handleFindUserPackages(replyToken, userId) {
  const result=findUserPackages(userId);
  if (result.status==='unregistered') sendUserQuickReplyMenu(replyToken,'❗️ ยังไม่ได้ลงทะเบียน\n\nกด "ลงทะเบียนรับพัสดุ" ก่อนครับ');
  else if (result.status==='no_packages') replyMessage(replyToken,'🎉 ไม่มีพัสดุที่รอรับในขณะนี้');
  else if (result.status==='found') {
      let msg='';
      if(result.waiting.length>0){msg+='📦 รายการพัสดุที่รอรับ:\n\n';result.waiting.forEach(p=>{msg+=` • เลขพัสดุ: ${p.trackingNumber}\n • รหัส: ${p.packageId}\n\n`;});}
      if(result.pickedUp.length>0){if(msg!=='')msg+='\n\n---\n\n';msg+='✅ รายการที่รับไปแล้ว (7 วันล่าสุด):\n\n';result.pickedUp.forEach(p=>{const d=new Date(p.checkOutTimestamp).toLocaleDateString('th-TH');msg+=` • รหัส: ${p.packageId}\n • วันที่รับ: ${d}\n\n`;});}
      replyMessage(replyToken,msg);
  }
}

function findUserPackages(userId) {
  try {
    const sData=subscribersSheet.getDataRange().getValues();
    let userPhone=null,userName=null;
    for(let i=1;i<sData.length;i++){if(sData[i][1]===userId){userPhone=sData[i][0];userName=sData[i][2];break;}}
    if(!userPhone||!userName) return{status:'unregistered'};
    
    // แบบใหม่ ตัดคำนำหน้า เพื่อค้นหาเฉพาะชื่อได้แม่นยำขึ้น
    const nameKW=userName.replace(/นาย|นาง|นางสาว|ด\.ช\.|ด\.ญ\./g,'').trim().split(/\s+/);
    const pData=packagesSheet.getDataRange().getValues();
    const waiting=[],pickedUp=[],found=new Set(),now=new Date();
    for(let i=1;i<pData.length;i++){
      const pkgId=pData[i][0];if(found.has(pkgId))continue;
      const pkgPhone=String(pData[i][2]).replace(/'/g,''),pkgName=pData[i][3],pkgStatus=pData[i][4];
      let isMatch=pkgPhone===userPhone;
      if(!isMatch){for(const kw of nameKW){if(kw&&pkgName.includes(kw)){isMatch=true;break;}}}
      if(isMatch){found.add(pkgId);const pk={packageId:pkgId,trackingNumber:pData[i][1],packageType:pData[i][9]};
        if(pkgStatus==='รอรับ')waiting.push(pk);
        else if(pkgStatus==='รับแล้ว'){const co=new Date(pData[i][6]);if(co.getTime()>0&&Math.ceil(Math.abs(now-co)/(1000*60*60*24))<=7){pk.checkOutTimestamp=co.getTime();pickedUp.push(pk);}}
      }
    }
    if(waiting.length===0&&pickedUp.length===0) return{status:'no_packages'};
    pickedUp.sort((a,b)=>b.checkOutTimestamp-a.checkOutTimestamp);
    return{status:'found',waiting,pickedUp};
  } catch(e){return{status:'error'};}
}

function isDuplicatePackage(trackingNumber) {
  if(!trackingNumber||trackingNumber==='ไม่มี') return false;
  const data=packagesSheet.getDataRange().getValues();
  for(let i=1;i<data.length;i++){if(data[i][1]===trackingNumber&&data[i][4]==='รอรับ')return true;}
  return false;
}

function showAllPackagesWithQuickReply(replyToken,userId,adminType,page=1) {
  try {
    const data=packagesSheet.getDataRange().getValues();
    const waiting=[];
    for(let i=1;i<data.length;i++){if(data[i][4]==='รอรับ')waiting.push({packageId:data[i][0],name:data[i][3]});}
    if(waiting.length===0){replyMessage(replyToken,'🎉 ไม่มีพัสดุที่รอรับ');return;}
    const PER_PAGE=11,totalPages=Math.ceil(waiting.length/PER_PAGE);
    const start=(page-1)*PER_PAGE,onPage=waiting.slice(start,start+PER_PAGE);
    let msg=`📦 รายการพัสดุรอรับ (หน้า ${page}/${totalPages}):\n\n`;
    const qrItems=[];
    onPage.forEach(p=>{msg+=`• รหัส: ${p.packageId} (${p.name})\n`;qrItems.push({type:'action',action:{type:'message',label:`รับ ${p.packageId}`,text:`ยืนยันการรับ ${p.packageId}`}});});
    const nav=[];
    if(page>1)nav.push({type:'action',action:{type:'message',label:'⬅️ หน้าก่อน',text:`ดูพัสดุ หน้า ${page-1}`}});
    if(start+PER_PAGE<waiting.length)nav.push({type:'action',action:{type:'message',label:'➡️ หน้าถัดไป',text:`ดูพัสดุ หน้า ${page+1}`}});
    replyMessageAdvanced(replyToken,[{type:'text',text:msg,quickReply:{items:qrItems.concat(nav)}}]);
  } catch(e){replyMessage(replyToken,'เกิดข้อผิดพลาด');}
}

function searchPackage(identifier) {
  try {
    const data=packagesSheet.getDataRange().getValues();
    let info=null;
    for(let i=data.length-1;i>0;i--){
      if(String(data[i][0])===identifier||String(data[i][1]).toLowerCase()===identifier.toLowerCase()){
        info={packageId:data[i][0],trackingNumber:data[i][1],phone:data[i][2],name:data[i][3],status:data[i][4],checkIn:data[i][5],checkOut:data[i][6]||'ยังไม่ได้รับ',notes:data[i][8],packageType:data[i][9],photoUrl:data[i][11]||'',sigUrl:data[i][12]||''};break;
      }
    }
    if(info){
      const msg=`🔍 ผลการค้นหา: ${identifier}\nเลขพัสดุ: ${info.trackingNumber}\nรหัส: ${info.packageId}\nสถานะ: ${info.status}\n\nผู้รับ: ${info.name}\nเบอร์โทร: ${info.phone}\nลักษณะ: ${info.packageType}`+(info.photoUrl?`\n🖼️ รูป: ${info.photoUrl}`:'');
      return{status:'success',message:msg};
    }
    return{status:'not_found',message:`❌ ไม่พบพัสดุ "${identifier}"`};
  } catch(e){return{status:'error',message:'เกิดข้อผิดพลาด'};}
}

function searchByPartialPhone(lastDigits) {
  if(!lastDigits||lastDigits.length<4) return{status:'error',message:'กรุณาระบุเลขท้ายอย่างน้อย 4 ตัว'};
  const data=packagesSheet.getDataRange().getValues(),found=[];
  for(let i=1;i<data.length;i++){if(data[i][4]==='รอรับ'&&String(data[i][2]).replace(/'/g,'').endsWith(lastDigits))found.push({packageId:data[i][0],recipientName:data[i][3]});}
  if(found.length>0){let msg=`📦 ค้นหาเบอร์ลงท้ายด้วย "${lastDigits}":\n\n`;found.forEach(p=>{msg+=`• รหัส: ${p.packageId}\n  ชื่อ: ${p.recipientName}\n\n`;});return{status:'success',message:msg};}
  return{status:'not_found',message:`❌ ไม่พบพัสดุที่เบอร์ลงท้ายด้วย "${lastDigits}"`};
}

function searchByName(nameQuery) {
  if(!nameQuery||nameQuery.length<2) return{status:'error',message:'กรุณาระบุชื่ออย่างน้อย 2 ตัว'};
  const data=packagesSheet.getDataRange().getValues(),q=nameQuery.toLowerCase(),found=[];
  for(let i=1;i<data.length;i++){if(data[i][4]==='รอรับ'&&String(data[i][3]).toLowerCase().includes(q))found.push({packageId:data[i][0],recipientName:data[i][3]});}
  if(found.length>0){let msg=`📦 ค้นหาชื่อ "${nameQuery}":\n\n`;found.forEach(p=>{msg+=`• รหัส: ${p.packageId}\n  ชื่อ: ${p.recipientName}\n\n`;});return{status:'success',message:msg};}
  return{status:'not_found',message:`❌ ไม่พบพัสดุชื่อ "${nameQuery}"`};
}

function handlePickupConfirmation(replyToken, packageId) {
  try {
    const data=packagesSheet.getDataRange().getValues();
    let info=null;
    for(let i=1;i<data.length;i++){if(String(data[i][0])===packageId){info={trackingNumber:data[i][1],recipientName:data[i][3]};break;}}
    if(!info){replyMessage(replyToken,`🔍 ไม่พบพัสดุรหัส ${packageId}`);return;}
    replyMessageAdvanced(replyToken,[{type:'text',
      text:`🤔 ยืนยันการรับพัสดุ?\n\nเลขพัสดุ: ${info.trackingNumber}\nรหัส: ${packageId}\nผู้รับ: ${info.recipientName}`,
      quickReply:{items:[
        {type:'action',action:{type:'message',label:'✅ ยืนยันการรับ',text:`✅ ยืนยันการรับ ${packageId}`}},
        {type:'action',action:{type:'message',label:'❌ ยกเลิก',text:'❌ ยกเลิกการรับ'}},
      ]}}]);
  } catch(e){replyMessage(replyToken,'เกิดข้อผิดพลาด');}
}

function processPackagePickupFromChat(identifier, adminUserId) {
  try {
    const data=packagesSheet.getDataRange().getValues();
    let foundRow=-1,info={};
    for(let i=data.length-1;i>0;i--){
      if(String(data[i][0])===identifier||String(data[i][1]).toLowerCase()===identifier.toLowerCase()){
        if(data[i][4]==='รอรับ'){foundRow=i+1;info={trackingNumber:data[i][1],recipientName:data[i][3],recipientPhone:String(data[i][2]).replace(/'/g,'')};break;}
        else return{message:`❌ พัสดุ ${identifier} ถูกรับไปแล้ว`};
      }
    }
    if(foundRow===-1) return{message:`🔍 ไม่พบพัสดุ ${identifier}`};
    const adminProfile=getUserProfile(adminUserId);
    const adminName=adminProfile?adminProfile.displayName:'เจ้าหน้าที่';
    packagesSheet.getRange(foundRow,5).setValue('รับแล้ว');
    packagesSheet.getRange(foundRow,7).setValue(new Date());
    const notes=packagesSheet.getRange(foundRow,9).getValue();
    packagesSheet.getRange(foundRow,9).setValue(notes+` | รับโดย: ${adminName}`);
    const now=new Date(), dateStr=Utilities.formatDate(now,CONFIG.TIMEZONE,'dd/MM/yyyy HH:mm');
    telegramSendMessage(`✅ พัสดุถูกรับแล้ว\n━━━━━━━━━━━━━━━━\n🔖 รหัส: ${identifier}\n📫 เลขพัสดุ: ${info.trackingNumber}\n👤 ผู้รับ: ${info.recipientName}\n🕐 เวลารับ: ${dateStr}\n👨‍✈️ จ่ายโดย: ${adminName}`);
    try {
      const tgInfo=getTGChatIdFuzzy(info.recipientPhone, info.recipientName);
      if(tgInfo&&tgInfo.chatId)telegramSendDirectMessage(tgInfo.chatId,`📬 <b>พัสดุของท่านถูกรับไปแล้ว</b>\n\n🔖 รหัส: <code>${identifier}</code>\n📫 เลขพัสดุ: <code>${info.trackingNumber}</code>\n🕐 เวลารับ: ${dateStr}\n👨‍✈️ จ่ายโดย: ${adminName}\n\n✅ เรียบร้อยแล้ว`);
    } catch(_) {}
    return{message:`✅ บันทึกการรับพัสดุ ${identifier} (คุณ${info.recipientName}) เรียบร้อย`};
  } catch(e){return{message:`เกิดข้อผิดพลาด: ${e.message}`};}
}

function checkSystemStatus(replyToken, userId) {
  const pkgData=packagesSheet.getDataRange().getValues();
  const waitingCount=pkgData.filter(r=>r[4]==='รอรับ').length;
  const subCount=subscribersSheet.getDataRange().getValues().length-1;
  const tgCount=tgSheet.getDataRange().getValues().length-1;
  replyMessage(replyToken,`📊 สถานะระบบ\n\n📦 พัสดุรอรับ: ${waitingCount} ชิ้น\n👥 ผู้ลงทะเบียน: ${subCount} คน\n📨 ผูก Telegram: ${tgCount} คน`);
}

function showSystemStats(replyToken, userId, adminType) {
  if(adminType!=='MAIN_ADMIN'){showUnauthorizedMessage(replyToken,'สถิติ');return;}
  const pkgs=packagesSheet.getDataRange().getValues();
  const total=pkgs.length-1,pickedUp=pkgs.filter(p=>p[4]==='รับแล้ว').length;
  replyMessage(replyToken,`📈 สถิติระบบ\nพัสดุทั้งหมด: ${total} ชิ้น\nรับไปแล้ว: ${pickedUp} ชิ้น\nรอรับ: ${total-pickedUp} ชิ้น`);
}

function showUnauthorizedMessage(replyToken, command) { replyMessage(replyToken,`❌ คุณไม่มีสิทธิ์ใช้คำสั่ง "${command}"\nกรุณาเข้าระบบเจ้าหน้าที่ก่อน`); }
function getDebugInfo() { return `<h2>📦 Package System : ${CONFIG.ORG_NAME}</h2><p><b>URL:</b> ${CONFIG.WEB_APP_URL}</p><pre>${Logger.getLog()}</pre>`; }

// =================================================================================
// LINE / TELEGRAM API HELPERS
// =================================================================================
function replyMessage(replyToken, text) { replyMessageAdvanced(replyToken,[{type:'text',text:String(text).substring(0,5000)}]); }
function replyMessageAdvanced(replyToken, messages) {
  try {
    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply',{method:'post',contentType:'application/json',muteHttpExceptions:true,payload:JSON.stringify({replyToken,messages}),headers:{Authorization:'Bearer '+CONFIG.CHANNEL_ACCESS_TOKEN}});
  } catch(e){Logger.log('replyMessageAdvanced error: '+e.message);}
}
function getUserProfile(userId) {
  try {
    const res=UrlFetchApp.fetch(`https://api.line.me/v2/bot/profile/${userId}`,{muteHttpExceptions:true,headers:{Authorization:'Bearer '+CONFIG.CHANNEL_ACCESS_TOKEN}});
    return res.getResponseCode()===200?JSON.parse(res.getContentText()):null;
  } catch(_){return null;}
}

function handleTelegramUpdate(update) {
  try {
    const msg = update.message;
    if (!msg || !msg.text) return;
    const chatId    = String(msg.chat.id);
    const text      = msg.text.trim();
    const username  = msg.from ? (msg.from.username  || '') : '';
    const firstName = msg.from ? (msg.from.first_name || '') : '';
    const cache    = CacheService.getScriptCache();
    const stateKey = 'tgstate_' + chatId;
    const stateRaw = cache.get(stateKey);
    const state    = stateRaw ? JSON.parse(stateRaw) : null;

    if (text === '/start' || text.startsWith('/start ')) {
      const parts = text.split(' ');
      if (parts.length === 2 && parts[1].startsWith('phone_')) {
        const phone = parts[1].replace('phone_', '').replace(/\D/g, '');
        tgBotDeepLinkPhone(chatId, username, firstName, phone);
      } else { tgBotStartFlow(chatId, username, firstName, cache, stateKey); }
      return;
    }
    if (text === '/cancel') { cache.remove(stateKey); telegramSendDirectMessage(chatId, '❌ ยกเลิกแล้ว'); return; }
    if (text === '/status')     { tgBotStatus(chatId);     return; }
    if (text === '/mypackages') { tgBotMyPackages(chatId); return; }
    if (state && state.step === 'WAIT_PHONE') { tgBotHandlePhone(chatId, text, username, cache, stateKey, state.isUpdate || false); return; }

    if (isChatIdLinked(chatId)) telegramSendDirectMessage(chatId, '📦 คุณลงทะเบียนแล้ว\n\n/mypackages — ดูพัสดุที่รอรับ\n/status — ตรวจสอบข้อมูล\n/start — เปลี่ยนเบอร์โทร');
    else telegramSendDirectMessage(chatId, `👋 สวัสดีครับ! ระบบพัสดุ ${CONFIG.ORG_NAME}\n\nพิมพ์ /start เพื่อลงทะเบียนรับแจ้งเตือนพัสดุ`);
  } catch(e) {}
}

function tgBotStartFlow(chatId, username, firstName, cache, stateKey) {
  const isUpdate = isChatIdLinked(chatId);
  if (isUpdate) {
    const info = getTGInfoByChatId(chatId);
    telegramSendDirectMessage(chatId, `📋 ข้อมูลปัจจุบัน\n📞 เบอร์: <code>${info.phone}</code>\n👤 ชื่อ: ${info.name || '—'}\n\nหากต้องการเปลี่ยนเบอร์ พิมพ์เบอร์ใหม่ (เช่น 0812345678)`);
  } else {
    telegramSendDirectMessage(chatId, `👋 สวัสดีคุณ${firstName||''}\n\n📦 ระบบแจ้งเตือนพัสดุ ${CONFIG.ORG_NAME}\n\nกรุณาพิมพ์เบอร์โทรศัพท์ 10 หลัก (เช่น 0812345678)`);
  }
  cache.put(stateKey, JSON.stringify({ step: 'WAIT_PHONE', isUpdate }), 600);
}

function tgBotHandlePhone(chatId, text, username, cache, stateKey, isUpdate) {
  let phone = text.replace(/\D/g, '');
  if (phone.startsWith('66') && phone.length === 11) phone = '0' + phone.substring(2);
  if (!/^0[689]\d{8}$/.test(phone)) { telegramSendDirectMessage(chatId, '❌ เบอร์โทรไม่ถูกต้อง (ต้องมี 10 หลัก)'); return; }

  let foundName = '';
  const subRows = subscribersSheet.getDataRange().getValues();
  for (let i = 1; i < subRows.length; i++) { if (String(subRows[i][0]).replace(/\D/g, '') === phone) { foundName = subRows[i][2] || ''; break; } }
  
  saveTGLink(phone, '', foundName, chatId, username);
  cache.remove(stateKey);
  telegramSendDirectMessage(chatId, `✅ ลงทะเบียนสำเร็จ!\n📞 เบอร์: <code>${phone}</code>\n👤 ชื่อ: ${foundName||'—'}\n\n/mypackages — ดูพัสดุที่รอรับ`);
}

function tgBotStatus(chatId) {
  const info = getTGInfoByChatId(chatId);
  if (info) telegramSendDirectMessage(chatId, `✅ ลงทะเบียนแล้ว\n📞 เบอร์: <code>${info.phone}</code>\n👤 ชื่อ: ${info.name||'—'}`);
  else telegramSendDirectMessage(chatId, `❌ ยังไม่ได้ลงทะเบียน พิมพ์ /start`);
}

function tgBotMyPackages(chatId) {
  const info = getTGInfoByChatId(chatId);
  if (!info) { telegramSendDirectMessage(chatId, '❌ ยังไม่ได้ลงทะเบียน พิมพ์ /start'); return; }
  const phone   = info.phone, name = info.name || '';
  const rows    = packagesSheet.getDataRange().getValues(), waiting = [];
  for (let i=1; i<rows.length; i++) {
    if (rows[i][4] !== 'รอรับ') continue;
    const pkgPhone = String(rows[i][2]).replace(/\D/g,'').replace(/^'/, ''), pkgName  = String(rows[i][3]);
    let match = (pkgPhone === phone);
    if (!match && name) {
      const kws = name.replace(/นาย|นาง|นางสาว/g,'').trim().split(/\s+/).filter(k => k.length >= 2);
      match = kws.some(kw => pkgName.includes(kw));
    }
    if (match) waiting.push(rows[i]);
  }
  if (waiting.length === 0) { telegramSendDirectMessage(chatId, '🎉 ไม่มีพัสดุที่รอรับในขณะนี้'); return; }
  let msg = `📦 <b>พัสดุที่รอรับ (${waiting.length} ชิ้น)</b>\n\n`;
  waiting.slice(0, 10).forEach((r, i) => { msg += `${i+1}. รหัส: <code>${r[0]}</code>\nเลขพัสดุ: <code>${r[1]}</code>\nเข้า: ${String(r[5]).substring(0,16)}\n\n`; });
  telegramSendDirectMessage(chatId, msg);
}

function isChatIdLinked(chatId) {
  const rows = getSheetData(tgSheet);
  for (let i=1; i<rows.length; i++) { if (String(rows[i][3]) === String(chatId)) return true; }
  return false;
}
function getTGInfoByChatId(chatId) {
  const rows = getSheetData(tgSheet);
  for (let i=1; i<rows.length; i++) { if (String(rows[i][3]) === String(chatId)) return { phone:String(rows[i][0]).replace(/'/g,''), lineId:String(rows[i][1]), name:String(rows[i][2]), username:String(rows[i][4]) }; }
  return null;
}

function setupTelegramWebhook() {
  const url     = `https://api.telegram.org/bot${CONFIG.TELEGRAM_BOT_TOKEN}/setWebhook`;
  const payload = { url: CONFIG.WEB_APP_URL, allowed_updates: ['message'], drop_pending_updates: true, max_connections: 1 };
  UrlFetchApp.fetch(url, { method:'post', contentType:'application/json', payload:JSON.stringify(payload), muteHttpExceptions:true });
}

function telegramSendMessage(text) {
  try {
    if(!CONFIG.TELEGRAM_BOT_TOKEN) return;
    UrlFetchApp.fetch(`https://api.telegram.org/bot${CONFIG.TELEGRAM_BOT_TOKEN}/sendMessage`,{method:'post',contentType:'application/json',muteHttpExceptions:true,payload:JSON.stringify({chat_id:CONFIG.TELEGRAM_CHAT_ID,text:String(text).substring(0,4096),parse_mode:'HTML'})});
  } catch(e){}
}

function telegramSendDirectMessage(chatId, text) {
  if (!chatId || chatId.length < 3 || !CONFIG.TELEGRAM_BOT_TOKEN) return;
  UrlFetchApp.fetch(`https://api.telegram.org/bot${CONFIG.TELEGRAM_BOT_TOKEN}/sendMessage`,{method:'post',contentType:'application/json',muteHttpExceptions:true,payload:JSON.stringify({chat_id:String(chatId),text:String(text).substring(0,4096),parse_mode:'HTML'})});
}

function telegramSendPhoto(fileId, caption) {
  try {
    if(!CONFIG.TELEGRAM_BOT_TOKEN) return;
    const file=DriveApp.getFileById(fileId);
    UrlFetchApp.fetch(`https://api.telegram.org/bot${CONFIG.TELEGRAM_BOT_TOKEN}/sendPhoto`,{method:'post',muteHttpExceptions:true,payload:{chat_id:CONFIG.TELEGRAM_CHAT_ID,photo:file.getBlob(),caption:caption?String(caption).substring(0,1024):''}});
  } catch(e){ if(caption)telegramSendMessage(caption); }
}