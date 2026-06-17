// ============================================================
// Influencer Work Submission System — Google Apps Script
// ============================================================

// --------------- Config Helper ---------------

function getConfigMap() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('config');
  var data = sheet.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < data.length; i++) {
    var key = data[i][0];
    var val = data[i][1];
    if (key) map[key] = val;
  }
  return map;
}

// --------------- CORS / Response Helper ---------------

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// --------------- doPost ---------------

function doPost(e) {
  try {
    var params = JSON.parse(e.postData.contents);
    var action = params.action;

    if (action === 'log')            return handleLog(params);
    if (action === 'getConfig')      return handleGetConfig();
    if (action === 'getDropdown')    return handleGetDropdown(params);
    if (action === 'saveInfluencer') return handleSaveInfluencer(params);
    if (action === 'uploadFile')     return handleUploadFile(params);
    if (action === 'submitWork')     return handleSubmitWork(params);
    if (action === 'saveBankAccount') return jsonResponse({ success: true });

    return jsonResponse({ success: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

// --------------- 1. log ---------------

function handleLog(params) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cfg = getConfigMap();
  var tz = cfg['time_zone'] || 'Asia/Bangkok';
  var sheet = ss.getSheetByName('log');
  if (!sheet) sheet = ss.insertSheet('log');

  var ts = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
  sheet.appendRow([ts, params.lineId, params.lineName, params.linePhoto]);
  return jsonResponse({ success: true });
}

// --------------- 2. getConfig ---------------

function handleGetConfig() {
  var cfg = getConfigMap();
  return jsonResponse({ success: true, config: cfg });
}

// --------------- 3. getDropdown ---------------

function handleGetDropdown(params) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = params.sheetName;
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return jsonResponse({ success: false, error: 'Sheet not found: ' + sheetName });

  var data = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), 1).getValues();
  var values = [];
  for (var i = 0; i < data.length; i++) {
    if (data[i][0]) values.push(data[i][0]);
  }
  return jsonResponse({ success: true, values: values });
}

// --------------- 4. saveInfluencer ---------------

function handleSaveInfluencer(params) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cfg = getConfigMap();
  var tz = cfg['time_zone'] || 'Asia/Bangkok';

  var sheet = ss.getSheetByName('influencer');
  if (!sheet) {
    sheet = ss.insertSheet('influencer');
    sheet.appendRow([
      'timestamp', 'line_id', 'line_name', 'line_photo', 'nickname', 'tiktok',
      'phone', 'creator_tag', 'folder_id', 'folder_final_id', 'folder_footage_id',
      'folder_stat_id', 'line_profile_drive_url', 'doc_urls', 'bank_name',
      'bank_account', 'account_name', 'gen_code', 'stat_urls',
      'post_link_1_platform', 'post_link_1_url', 'post_link_2_platform', 'post_link_2_url'
    ]);
  }

  var lineId = params.lineId;

  // Check if already exists
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === lineId) {
      return jsonResponse({
        success: true,
        alreadyExists: true,
        folderId: data[i][8],
        finalFolderId: data[i][9],
        footageFolderId: data[i][10],
        statFolderId: data[i][11],
        creatorTag: data[i][7],
        rowIndex: i + 1
      });
    }
  }

  var nickname = params.nickname;
  var tiktokUsername = params.tiktokUsername;
  var phone = params.phone;
  var lineName = params.lineName;
  var linePhoto = params.linePhoto;

  var creatorTag = nickname + ',@' + tiktokUsername;

  // Create Drive folder
  var parentFolderId = cfg['drive_parent_folder_id'];
  var parentFolder = DriveApp.getFolderById(parentFolderId);
  var folderName = nickname + ',@' + tiktokUsername;
  var mainFolder = parentFolder.createFolder(folderName);
  var folderId = mainFolder.getId();

  // Download LINE profile photo and save to Drive
  var profileDriveUrl = '';
  try {
    var response = UrlFetchApp.fetch(linePhoto);
    var blob = response.getBlob();
    blob.setName('line_profile.jpg');
    var profileFile = mainFolder.createFile(blob);
    var profileFileId = profileFile.getId();
    profileDriveUrl = 'https://lh3.googleusercontent.com/d/' + profileFileId;
  } catch (err) {
    profileDriveUrl = linePhoto; // fallback to original URL
  }

  // Create subfolders
  var finalFolder = mainFolder.createFolder('ชิ้นงานสมบูรณ์');
  var footageFolder = mainFolder.createFolder('ฟุตเทจ');
  var statFolder = mainFolder.createFolder('Stat');

  var finalFolderId = finalFolder.getId();
  var footageFolderId = footageFolder.getId();
  var statFolderId = statFolder.getId();

  var ts = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');

  sheet.appendRow([
    ts, lineId, lineName, linePhoto, nickname, tiktokUsername, phone,
    creatorTag, folderId, finalFolderId, footageFolderId, statFolderId,
    profileDriveUrl,
    '', '', '', '', '', '',    // doc_urls, bank_name, bank_account, account_name, gen_code, stat_urls
    '', '', '', ''             // post_link_1_platform, post_link_1_url, post_link_2_platform, post_link_2_url
  ]);

  var rowIndex = sheet.getLastRow();

  return jsonResponse({
    success: true,
    alreadyExists: false,
    folderId: folderId,
    finalFolderId: finalFolderId,
    footageFolderId: footageFolderId,
    statFolderId: statFolderId,
    creatorTag: creatorTag,
    rowIndex: rowIndex
  });
}

// --------------- 5. uploadFile ---------------

function handleUploadFile(params) {
  var folderId = params.folderId;
  var fileName = params.fileName;
  var mimeType = params.mimeType;
  var base64Data = params.base64Data;

  var folder = DriveApp.getFolderById(folderId);
  var decoded = Utilities.base64Decode(base64Data);
  var blob = Utilities.newBlob(decoded, mimeType, fileName);
  var file = folder.createFile(blob);
  var fileId = file.getId();
  var driveUrl = 'https://lh3.googleusercontent.com/d/' + fileId;

  return jsonResponse({ success: true, fileId: fileId, driveUrl: driveUrl });
}

// --------------- 6. submitWork ---------------

function handleSubmitWork(params) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cfg = getConfigMap();

  var sheet = ss.getSheetByName('influencer');
  if (!sheet) return jsonResponse({ success: false, error: 'influencer sheet not found' });

  var lineId = params.lineId;
  var data = sheet.getDataRange().getValues();
  var rowIndex = -1;
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === lineId) {
      rowIndex = i + 1;
      break;
    }
  }
  if (rowIndex === -1) return jsonResponse({ success: false, error: 'influencer not found' });

  // Columns (1-indexed):
  // 1:timestamp 2:line_id 3:line_name 4:line_photo 5:nickname 6:tiktok 7:phone
  // 8:creator_tag 9:folder_id 10:folder_final_id 11:folder_footage_id 12:folder_stat_id
  // 13:line_profile_drive_url 14:doc_urls 15:bank_name 16:bank_account 17:account_name
  // 18:gen_code 19:stat_urls 20:post_link_1_platform 21:post_link_1_url
  // 22:post_link_2_platform 23:post_link_2_url

  sheet.getRange(rowIndex, 14).setValue(params.docUrls || '');
  sheet.getRange(rowIndex, 15).setValue(params.bankName || '');
  sheet.getRange(rowIndex, 16).setValue(params.bankAccount || '');
  sheet.getRange(rowIndex, 17).setValue(params.accountName || '');
  sheet.getRange(rowIndex, 18).setValue(params.genCode || '');
  sheet.getRange(rowIndex, 19).setValue(params.statUrls || '');
  sheet.getRange(rowIndex, 20).setValue(params.post1Platform || '');
  sheet.getRange(rowIndex, 21).setValue(params.post1Url || '');
  sheet.getRange(rowIndex, 22).setValue(params.post2Platform || '');
  sheet.getRange(rowIndex, 23).setValue(params.post2Url || '');

  // Reload row for Flex message data
  var row = sheet.getRange(rowIndex, 1, 1, 23).getValues()[0];
  var influencer = {
    lineId:             row[1],
    lineName:           row[2],
    linePhoto:          row[3],
    nickname:           row[4],
    tiktok:             row[5],
    phone:              row[6],
    creatorTag:         row[7],
    folderId:           row[8],
    finalFolderId:      row[9],
    footageFolderId:    row[10],
    statFolderId:       row[11],
    profileDriveUrl:    row[12],
    docUrls:            row[13],
    bankName:           row[14],
    bankAccount:        row[15],
    accountName:        row[16],
    genCode:            row[17],
    statUrls:           row[18],
    post1Platform:      row[19],
    post1Url:           row[20],
    post2Platform:      row[21],
    post2Url:           row[22]
  };

  sendFlexMessage(cfg, influencer);

  return jsonResponse({ success: true });
}

// --------------- Flex Message ---------------

function sendFlexMessage(cfg, inf) {
  var token = cfg['line_channel_token'];
  var groupId = cfg['line_group_id'];
  var brandTag = cfg['brand_tag'] || '';

  var bodyContents = [];

  // Phone
  bodyContents.push({
    type: 'text',
    text: '📞 ' + inf.phone,
    size: 'sm',
    color: '#555555'
  });

  // Folder buttons
  var folders = [
    { label: '📁 โฟลเดอร์หลัก', id: inf.folderId },
    { label: '🎬 ชิ้นงานสมบูรณ์', id: inf.finalFolderId },
    { label: '🎥 ฟุตเทจ', id: inf.footageFolderId },
    { label: '📊 Stat', id: inf.statFolderId }
  ];

  folders.forEach(function(f) {
    if (f.id) {
      bodyContents.push({
        type: 'button',
        style: 'secondary',
        height: 'sm',
        margin: 'sm',
        action: {
          type: 'uri',
          label: f.label,
          uri: 'https://drive.google.com/drive/folders/' + f.id
        }
      });
    }
  });

  // Doc URLs
  if (inf.docUrls) {
    inf.docUrls.split(',').forEach(function(url) {
      url = url.trim();
      if (url) {
        bodyContents.push({
          type: 'image',
          url: url,
          size: 'full',
          aspectMode: 'cover',
          aspectRatio: '20:13',
          margin: 'sm'
        });
      }
    });
  }

  // Bank info
  bodyContents.push({
    type: 'text',
    text: '🏦 ' + inf.bankName + ' | ' + inf.bankAccount + ' | ' + inf.accountName,
    size: 'sm',
    color: '#333333',
    margin: 'sm',
    wrap: true
  });

  // Post links
  if (inf.post1Platform && inf.post1Url) {
    bodyContents.push({
      type: 'button',
      style: 'link',
      height: 'sm',
      margin: 'sm',
      action: {
        type: 'uri',
        label: '🔗 ' + inf.post1Platform,
        uri: inf.post1Url
      }
    });
  }
  if (inf.post2Platform && inf.post2Url) {
    bodyContents.push({
      type: 'button',
      style: 'link',
      height: 'sm',
      margin: 'sm',
      action: {
        type: 'uri',
        label: '🔗 ' + inf.post2Platform,
        uri: inf.post2Url
      }
    });
  }

  // Stat URLs
  if (inf.statUrls) {
    inf.statUrls.split(',').forEach(function(url) {
      url = url.trim();
      if (url) {
        bodyContents.push({
          type: 'image',
          url: url,
          size: 'full',
          aspectMode: 'cover',
          aspectRatio: '20:13',
          margin: 'sm'
        });
      }
    });
  }

  // Gen code
  if (inf.genCode) {
    bodyContents.push({
      type: 'button',
      style: 'primary',
      color: '#17c950',
      height: 'sm',
      margin: 'sm',
      action: {
        type: 'clipboard',
        label: '📋 เจนโค้ด: ' + inf.genCode,
        clipboardText: inf.genCode
      }
    });
  }

  var heroImage = inf.profileDriveUrl || cfg['header_image_url'] || 'https://via.placeholder.com/400x200';

  var flexMessage = {
    type: 'flex',
    altText: 'ส่งงาน: ' + inf.creatorTag,
    contents: {
      type: 'bubble',
      hero: {
        type: 'image',
        url: heroImage,
        size: 'full',
        aspectRatio: '20:13',
        aspectMode: 'cover'
      },
      header: {
        type: 'box',
        layout: 'vertical',
        contents: [
          {
            type: 'text',
            text: inf.creatorTag,
            weight: 'bold',
            size: 'xl',
            color: '#1DB446'
          }
        ]
      },
      body: {
        type: 'box',
        layout: 'vertical',
        contents: bodyContents,
        spacing: 'sm',
        paddingAll: '13px'
      },
      footer: {
        type: 'box',
        layout: 'vertical',
        contents: [
          {
            type: 'text',
            text: brandTag,
            size: 'xs',
            color: '#aaaaaa',
            align: 'center'
          }
        ]
      }
    }
  };

  var payload = {
    to: groupId,
    messages: [flexMessage]
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + token
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', options);
}
