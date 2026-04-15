// =================================================================================
// SMART PARCEL MANAGEMENT SYSTEM — Code.gs (MASTER VERSION 4.0)
// =================================================================================

// =================================================================================
// 1. CONFIGURATION (ดึงข้อมูลจาก Script Properties เพื่อความปลอดภัย)
// =================================================================================
const _props = PropertiesService.getScriptProperties();
const CONFIG = {
  'ORG_NAME':              _props.getProperty('ORG_NAME') || 'Smart Parcel System',
  
  'SHEET_ID':              _props.getProperty('SHEET_ID') || '',
  'CHANNEL_ACCESS_TOKEN':  _props.getProperty('CHANNEL_ACCESS_TOKEN') || '',
  
  'ADMIN_LOGIN_SECRET':    _props.getProperty('ADMIN_LOGIN_SECRET') || 'ADMIN007',
  'MAIN_ADMIN_SECRET':     _props.getProperty('MAIN_ADMIN_SECRET')  || 'SUPERADMIN2025',
  
  'WEB_APP_URL':           _props.getProperty('WEB_APP_URL') || '',
  'DRIVE_FOLDER_NAME':     _props.getProperty('DRIVE_FOLDER_NAME') || 'PackagePhotos',

  // Telegram Bot (ถ้าปล่อยว่าง ระบบจะทำงานแบบแพ็กเกจ Standard อัตโนมัติ)
  'TELEGRAM_BOT_TOKEN':    _props.getProperty('TELEGRAM_BOT_TOKEN') || '',
  'TELEGRAM_BOT_USERNAME': _props.getProperty('TELEGRAM_BOT_USERNAME') || '',
  'TELEGRAM_CHAT_ID':      _props.getProperty('TELEGRAM_CHAT_ID') || '',

  // LIFF URLs
  'LIFF_REGISTER_URL':     _props.getProperty('LIFF_REGISTER_URL') || '', 
  'LIFF_CHECKIN_URL':      _props.getProperty('LIFF_CHECKIN_URL') || '',
  'LIFF_PICKUP_URL':       _props.getProperty('LIFF_PICKUP_URL') || '',

  'TIMEZONE':              _props.getProperty('TIMEZONE') || 'Asia/Bangkok',
};

// =================================================================================
// 2. SETUP SCRIPT PROPERTIES (ฟังก์ชันสำหรับรันครั้งแรก เพื่อตั้งค่าตัวแปร)
// =================================================================================
function setupScriptProperties() {
  const props = PropertiesService.getScriptProperties();
  
  // ✏️ กำหนดค่าต่างๆ ของลูกค้าที่นี่ แล้วกด "รัน" ฟังก์ชันนี้ที่แถบด้านบน
  props.setProperties({
    'ORG_NAME': 'ชื่อองค์กร/บริษัท/คอนโด ของลูกค้า',
    'SHEET_ID': 'ใส่_SHEET_ID_หรือ_ลิงก์_Google_Sheet',
    'CHANNEL_ACCESS_TOKEN': 'ใส่_LINE_TOKEN',
    
    'ADMIN_LOGIN_SECRET': 'ADMIN007',
    'MAIN_ADMIN_SECRET': 'SUPERADMIN2025',
    
    'WEB_APP_URL': 'ใส่_GAS_WEBAPP_URL_ตรงนี้',
    'DRIVE_FOLDER_NAME': 'PackagePhotos',
    
    // สำหรับแพ็กเกจ Premium ให้ใส่ข้อมูล | แพ็กเกจ Standard ปล่อยว่างไว้
    'TELEGRAM_BOT_TOKEN': '',
    'TELEGRAM_BOT_USERNAME': '',
    'TELEGRAM_CHAT_ID': '',
    
    'LIFF_REGISTER_URL': 'https://liff.line.me/ใส่_LIFF_ID_ลงทะเบียน',
    'LIFF_CHECKIN_URL': 'https://liff.line.me/ใส่_LIFF_ID_บันทึกรับ',
    'LIFF_PICKUP_URL': 'https://liff.line.me/ใส่_LIFF_ID_จ่ายพัสดุ',
    
    'TIMEZONE': 'Asia/Bangkok'
  });
  
  Logger.log('✅ ตั้งค่า Script Properties ลงในระบบเรียบร้อยแล้ว!');
}

// =================================================================================
// 3. SHEET MANAGEMENT
// =================================================================================
let aSheet, subscribersSheet, packagesSheet, adminLogSheet, tgSheet;

try {
  let sid = CONFIG.SHEET_ID || '';
  if (sid.includes('/d/')) {
    const match = sid.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (match) sid = match[1];
  }
  sid = sid.trim();
  
  if (sid && !sid.includes('ใส่_')) {
    aSheet            = SpreadsheetApp.openById(sid);
    subscribersSheet  = aSheet.getSheetByName('Subscribers') || createSheet('Subscribers');
    packagesSheet     = aSheet.getSheetByName('Packages')    || createSheet('Packages');
    adminLogSheet     = aSheet.getSheetByName('AdminLog')    || createSheet('AdminLog');
    tgSheet           = aSheet.getSheetByName('Subscribers_TG') || createTGSheet();
  }
} catch (e) {
  Logger.log('Sheet Init Error: ' + e.message);
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

// CACHE
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
// 4. HTTP REQUEST HANDLERS
// =================================================================================
function doGet(e) {
  if (!aSheet) return HtmlService.createHtmlOutput('<h2 style="font-family:sans-serif;padding:20px;">⚠️ ข้อผิดพลาด: ไม่พบ Google Sheet<br><br>กรุณาตรวจสอบว่ารัน setupScriptProperties() ถูกต้อง</h2>');
  if (e.parameter.page === 'debug') return HtmlService.createHtmlOutput(getDebugInfo());
  if (e.parameter.action === 'check_admin') return jsonResponse(checkAdminForLiff(e.parameter.userId));
  return HtmlService.createHtmlOutput(`<h2>📦 Package System : ${CONFIG.ORG_NAME} — Online ✅</h2>`);
}

function doPost(e) {
  try {
    if (!aSheet) return jsonResponse({ status:'error', message:'ระบบยังไม่ได้ตั้งค่า SHEET_ID หรือ ID ไม่ถูกต้อง' });
    if (!e.postData || !e.postData.contents) return jsonResponse({ status:'error', message:'No post data' });
    
    let data;
    try { data = JSON.parse(e.postData.contents); }
    catch(_) { data = JSON.parse(e.parameter.payload || e.postData.contents); }

    // LIFF Endpoints
    if (data.action === 'liff_register')         return handleLiffRegister(data);     
    if (data.action === 'get_registration')      return handleGetRegistration(data);   
    if (data.action === 'liff_submit')           return handleLiffSubmit(data);
    if (data.action === 'upload_package_photo')  return handleUploadPackagePhoto(data);
    if (data.action === 'lookup_package')        return handleLookupPackage(data);
    if (data.action === 'pickup_package')        return handlePickupPackage(data);
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
// 5. LIFF HANDLERS
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

function handleLiffSubmit(data) {
  try {
    const { userId, tracking, name, phone, type, hasPhoto } = data;
    const auth = checkAdminForLiff(userId);
    if (!auth.isAdmin) return jsonResponse({ status:'error', message:auth.reason });
    const now=new Date();
    const formattedTimestamp=Utilities.formatDate(now,'GMT+7','dd/MM/yyyy HH:mm:ss');
    const d=String(now.getDate()).padStart(2,'0'), m=String(now.getMonth()+1).padStart(2,'0');
    const sy=String(now.getFullYear()+543).slice(-2), h=String(now.getHours()).padStart(2,'0'), mi=String(now.getMinutes()).padStart(2,'0'), s=String(now.getSeconds()).padStart(2,'0');
    const packageId=`${d}${m}${sy}${h}${mi}${s}`;
    
    if (isDuplicatePackage(tracking)) return jsonResponse({ status:'error', message:`พัสดุ ${tracking} มีในระบบและยังรอรับอยู่` });
    
    const packageData = {
      PackageId: packageId, TrackingNumber: tracking, PhoneNumberOnLabel: phone,
      RecipientNameOnLabel: name, Status: 'รอรับ', CheckInTimestamp: formattedTimestamp,
      CheckOutTimestamp: '', Carrier: 'ไม่ระบุ', Notes: `บันทึกโดย ${auth.userName} (LIFF)`,
      PackageType: type, AdminUser: auth.userName, PhotoUrl: '', SignatureUrl: ''
    };
    
    const saveResult = savePackageFromChat(packageData);
    
    // ถ้าไม่มีรูปภาพ ให้ส่งข้อความแจ้งเตือนเลย (ถ้ามีรูปภาพ จะส่งตอน Upload Photo)
    if (saveResult.status === 'success' && !hasPhoto) {
      telegramSendMessage(`📦 พัสดุเข้าใหม่\n━━━━━━━━━━━━━━━━\n🔖 รหัส: ${packageId}\n📫 เลขพัสดุ: ${tracking}\n👤 ผู้รับ: ${name}\n📦 ลักษณะ: ${type}`);
      notifyRecipientTelegram(phone, name, packageId, type, tracking, '');
    }
    
    if (saveResult.status==='success') return jsonResponse({ status:'success', packageId });
    return jsonResponse({ status:'error', message:saveResult.message });
  } catch(e) { return jsonResponse({ status:'error', message:e.message }); }
}

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
        sigUrl = savePhotoToDrive(signatureBase64, 'image/png', `SIG_${packageId}`);
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
// 6. UTILS & HELPERS
// =================================================================================

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

function savePackageFromChat(packageData) {
  try {
    packagesSheet.appendRow([`'${packageData.PackageId}`, packageData.TrackingNumber, `'${packageData.PhoneNumberOnLabel}`, packageData.RecipientNameOnLabel, packageData.Status, packageData.CheckInTimestamp, packageData.CheckOutTimestamp, packageData.Carrier, packageData.Notes, packageData.PackageType, packageData.AdminUser, packageData.PhotoUrl||'', packageData.SignatureUrl||'']);
    return{status:'success',message:'บันทึกสำเร็จ',packageId:packageData.PackageId};
  } catch(e){return{status:'error',message:`ข้อผิดพลาด: ${e.message}`};}
}

function isDuplicatePackage(trackingNumber) {
  if(!trackingNumber||trackingNumber==='ไม่มี') return false;
  const data=packagesSheet.getDataRange().getValues();
  for(let i=1;i<data.length;i++){if(data[i][1]===trackingNumber&&data[i][4]==='รอรับ')return true;}
  return false;
}

function getDebugInfo() { return `<h2>📦 Package System : ${CONFIG.ORG_NAME}</h2><p><b>URL:</b> ${CONFIG.WEB_APP_URL}</p><pre>${Logger.getLog()}</pre>`; }

// =================================================================================
// 7. LINE BOT HANDLERS
// =================================================================================

function handleWebhook(event) {
  try {
    if (event.type==='follow') {
      sendUserQuickReplyMenu(event.replyToken, `🏢 ยินดีต้อนรับสู่ระบบพัสดุ\n${CONFIG.ORG_NAME}\n\n📦 กดปุ่ม "ลงทะเบียนรับพัสดุ" ด้านล่างเพื่อเริ่มต้นใช้งาน`);
    } else if (event.type==='message'&&event.message.type==='text') {
      handleTextMessage(event.replyToken, event.source.userId, event.message.text);
    }
  } catch(e){Logger.log('handleWebhook error: '+e.message);}
}

function handleTextMessage(replyToken, userId, text) {
  try {
    const message=text.trim(), adminStatus=isAdmin(userId);

    if (message==='เมนู'||message.toLowerCase()==='menu') {
      if(adminStatus.isAdmin){ if(adminStatus.adminType==='MAIN_ADMIN') sendMainAdminQuickReplyMenu(replyToken,'เมนูแอดมินหลัก'); else sendStaffQuickReplyMenu(replyToken,'เมนูเจ้าหน้าที่'); } 
      else { sendUserQuickReplyMenu(replyToken,'เมนูผู้รับพัสดุ'); }
      return;
    }
    if (message==='พัสดุของฉัน') { handleFindUserPackages(replyToken,userId); return; }
    if (message===CONFIG.ADMIN_LOGIN_SECRET||message===CONFIG.MAIN_ADMIN_SECRET) { processAdminLogin(message,userId,replyToken); return; }
    if (message.includes('ออกจากระบบ')||message.includes('ออก')) { handleLogout(replyToken,userId); return; }
    if (message.includes('ตรวจสอบ')||message.includes('สถานะ')) { checkSystemStatus(replyToken,userId); return; }
    
    if (adminStatus.isAdmin) replyMessage(replyToken,'🤔 ไม่เข้าใจคำสั่ง\n\nพิมพ์ `เมนู` เพื่อดูตัวเลือก');
    else sendUserQuickReplyMenu(replyToken,'🤔 ไม่เข้าใจคำสั่ง\n\nกดปุ่มเมนูด้านล่างเพื่อเริ่มใช้งาน');
  } catch(e){Logger.log('handleTextMessage error: '+e.message);replyMessage(replyToken,'เกิดข้อผิดพลาด');}
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

function handleLogout(replyToken,userId){
  if(isAdmin(userId).isAdmin){
    const d=adminLogSheet.getDataRange().getValues();
    for(let i=d.length-1;i>0;i--){if(d[i][0]===userId&&d[i][3]==='ACTIVE'){adminLogSheet.getRange(i+1,4).setValue('INACTIVE');break;}}
    replyMessage(replyToken,'✅ ออกจากระบบเรียบร้อย');
  } else replyMessage(replyToken,'❌ คุณยังไม่ได้เข้าระบบ');
}

function sendUserQuickReplyMenu(replyToken, welcomeText) {
  replyMessageAdvanced(replyToken, [{
    type: 'text', text: welcomeText,
    quickReply: { items: [
        { type: 'action', action: { type: 'uri', label: '📋 ลงทะเบียนรับพัสดุ', uri: CONFIG.LIFF_REGISTER_URL } },
        { type: 'action', action: { type: 'message', label: '📦 พัสดุของฉัน', text: 'พัสดุของฉัน' } }
      ] }
  }]);
}

function sendStaffQuickReplyMenu(replyToken, welcomeText) {
  replyMessageAdvanced(replyToken,[{type:'text',text:welcomeText,quickReply:{items:[
    {type:'action',action:{type:'uri',label:'📷 บันทึกพัสดุเข้า',uri:CONFIG.LIFF_CHECKIN_URL}},
    {type:'action',action:{type:'uri',label:'📬 จ่ายพัสดุออก',uri:CONFIG.LIFF_PICKUP_URL}},
    {type:'action',action:{type:'message',label:'🚪 ออกจากระบบ',text:'ออกจากระบบ'}},
  ]}}]);
}

function sendMainAdminQuickReplyMenu(replyToken, welcomeText) {
  replyMessageAdvanced(replyToken,[{type:'text',text:welcomeText,quickReply:{items:[
    {type:'action',action:{type:'uri',label:'📷 บันทึกพัสดุ',uri:CONFIG.LIFF_CHECKIN_URL}},
    {type:'action',action:{type:'uri',label:'📬 จ่ายพัสดุ',uri:CONFIG.LIFF_PICKUP_URL}},
    {type:'action',action:{type:'message',label:'📊 ตรวจสอบสถานะ',text:'ตรวจสอบสถานะ'}},
    {type:'action',action:{type:'message',label:'🚪 ออกจากระบบ',text:'ออกจากระบบ'}},
  ]}}]);
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

function checkSystemStatus(replyToken, userId) {
  const pkgData=packagesSheet.getDataRange().getValues();
  const waitingCount=pkgData.filter(r=>r[4]==='รอรับ').length;
  const subCount=subscribersSheet.getDataRange().getValues().length-1;
  const tgCount=tgSheet.getDataRange().getValues().length-1;
  replyMessage(replyToken,`📊 สถานะระบบ\n\n📦 พัสดุรอรับ: ${waitingCount} ชิ้น\n👥 ผู้ลงทะเบียน: ${subCount} คน\n📨 ผูก Telegram: ${tgCount} คน`);
}

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

// =================================================================================
// 8. TELEGRAM HANDLERS
// =================================================================================

function telegramSendMessage(text) {
  try {
    if(!CONFIG.TELEGRAM_BOT_TOKEN || CONFIG.TELEGRAM_BOT_TOKEN.includes('ใส่_')) return;
    UrlFetchApp.fetch(`https://api.telegram.org/bot${CONFIG.TELEGRAM_BOT_TOKEN}/sendMessage`,{method:'post',contentType:'application/json',muteHttpExceptions:true,payload:JSON.stringify({chat_id:CONFIG.TELEGRAM_CHAT_ID,text:String(text).substring(0,4096),parse_mode:'HTML'})});
  } catch(e){}
}

function telegramSendDirectMessage(chatId, text) {
  if (!chatId || chatId.length < 3) return;
  if (!CONFIG.TELEGRAM_BOT_TOKEN || CONFIG.TELEGRAM_BOT_TOKEN.includes('ใส่_')) return;
  try {
    UrlFetchApp.fetch(`https://api.telegram.org/bot${CONFIG.TELEGRAM_BOT_TOKEN}/sendMessage`,{method:'post',contentType:'application/json',muteHttpExceptions:true,payload:JSON.stringify({chat_id:String(chatId),text:String(text).substring(0,4096),parse_mode:'HTML'})});
  } catch(e){}
}

function telegramSendPhoto(fileId, caption) {
  try {
    if(!CONFIG.TELEGRAM_BOT_TOKEN || CONFIG.TELEGRAM_BOT_TOKEN.includes('ใส่_')) return;
    const file=DriveApp.getFileById(fileId);
    UrlFetchApp.fetch(`https://api.telegram.org/bot${CONFIG.TELEGRAM_BOT_TOKEN}/sendPhoto`,{method:'post',muteHttpExceptions:true,payload:{chat_id:CONFIG.TELEGRAM_CHAT_ID,photo:file.getBlob(),caption:caption?String(caption).substring(0,1024):''}});
  } catch(e){ if(caption)telegramSendMessage(caption); }
}

function handleCheckTGLinked(data) {
  try {
    const { phone } = data;
    if (!phone) return jsonResponse({ status: 'error', message: 'ไม่พบเบอร์' });
    const tgInfo = getTGChatIdByPhone(phone);
    return jsonResponse({ status: 'ok', result: { linked: !!(tgInfo && tgInfo.chatId && tgInfo.chatId.length > 3), tgUsername: tgInfo ? tgInfo.username : '' } });
  } catch(e) { return jsonResponse({ status: 'error', message: e.message }); }
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

  if (recipientName && recipientName.length >= 2) {
    const stripped = recipientName
      .replace(/นาย|นาง|นางสาว|ด\.ช\.|ด\.ญ\.|คุณ|แผนก|ห้อง/g, '')
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

function handleTelegramUpdate(update) {
  try {
    const msg = update.message;
    if (!msg || !msg.text) return;
    const chatId    = String(msg.chat.id);
    const text      = msg.text.trim();
    const username  = msg.from ? (msg.from.username  || '') : '';
    const firstName = msg.from ? (msg.from.first_name || '') : '';
    
    if (text === '/start' || text.startsWith('/start ')) {
      const parts = text.split(' ');
      if (parts.length === 2 && parts[1].startsWith('phone_')) {
        const phone = parts[1].replace('phone_', '').replace(/\D/g, '');
        tgBotDeepLinkPhone(chatId, username, firstName, phone);
      } else { 
        telegramSendDirectMessage(chatId, `👋 สวัสดีครับ! ระบบพัสดุ ${CONFIG.ORG_NAME}\nกรุณาผูกเบอร์ผ่าน LINE LIFF ครับ`);
      }
      return;
    }
    if (text === '/mypackages') { tgBotMyPackages(chatId); return; }
  } catch(e) {}
}

function tgBotDeepLinkPhone(chatId, username, firstName, phone) {
  try {
    const _cache   = CacheService.getScriptCache();
    const _coolKey = 'deeplink_cd_' + chatId + '_' + phone;
    if (_cache.get(_coolKey)) return;
    _cache.put(_coolKey, '1', 120);

    if (!/^0[689]\d{8}$/.test(phone)) {
      telegramSendDirectMessage(chatId, '❌ ลิงก์ไม่ถูกต้อง\nกรุณาลองใหม่จาก LINE'); return;
    }

    let foundName = '';
    const subRows = subscribersSheet.getDataRange().getValues();
    for (let i = 1; i < subRows.length; i++) {
      if (String(subRows[i][0]).replace(/\D/g, '') === phone) { foundName = subRows[i][2] || ''; break; }
    }

    saveTGLink(phone, '', foundName, chatId, username);

    const greet    = firstName ? `คุณ${firstName}` : 'คุณ';
    const nameLine = foundName ? `👤 ชื่อ: <b>${foundName}</b>\n` : '';

    telegramSendDirectMessage(chatId,
      `✅ ผูก Telegram สำเร็จ! สวัสดี ${greet}\n\n` +
      `📞 เบอร์: <code>${phone}</code>\n` + nameLine +
      `\n📦 ระบบจะแจ้งเตือนเมื่อ:\n• มีพัสดุเข้าใหม่\n• พัสดุถูกจ่ายออกแล้ว\n\n` +
      `/mypackages — ดูพัสดุที่รอรับ`);
  } catch(e) {}
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

function tgBotMyPackages(chatId) {
  const info = getTGInfoByChatId(chatId);
  if (!info) { telegramSendDirectMessage(chatId, '❌ ยังไม่ได้ลงทะเบียน กรุณาลงทะเบียนผ่าน LINE'); return; }
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

function getTGInfoByChatId(chatId) {
  const rows = getSheetData(tgSheet);
  for (let i=1; i<rows.length; i++) { if (String(rows[i][3]) === String(chatId)) return { phone:String(rows[i][0]).replace(/'/g,''), lineId:String(rows[i][1]), name:String(rows[i][2]), username:String(rows[i][4]) }; }
  return null;
}

function setupTelegramWebhook() {
  if (!CONFIG.TELEGRAM_BOT_TOKEN || CONFIG.TELEGRAM_BOT_TOKEN.includes('ใส่_')) {
    Logger.log('ข้ามการเซ็ต Webhook Telegram เนื่องจากไม่ได้ใส่ Token');
    return;
  }
  const url     = `https://api.telegram.org/bot${CONFIG.TELEGRAM_BOT_TOKEN}/setWebhook`;
  const payload = { url: CONFIG.WEB_APP_URL, allowed_updates: ['message'], drop_pending_updates: true, max_connections: 1 };
  UrlFetchApp.fetch(url, { method:'post', contentType:'application/json', payload:JSON.stringify(payload), muteHttpExceptions:true });
}

// =================================================================================
// 9. DAILY REPORT
// =================================================================================
function sendDailyReport() {
  try {
    if (!CONFIG.TELEGRAM_BOT_TOKEN || CONFIG.TELEGRAM_BOT_TOKEN.includes('ใส่_')) return;
    
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
  } catch(e) {}
}

function setupDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'sendDailyReport') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('sendDailyReport').timeBased().atHour(21).everyDays(1).inTimezone(CONFIG.TIMEZONE).create();
}
