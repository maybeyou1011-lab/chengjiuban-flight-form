// ============================================================
// 成就班機票資料收集表 → Google 試算表 收件程式（Google Apps Script）
// 用途：把學員填的資料一筆一筆寫進試算表，護照照片存到雲端硬碟資料夾。
//
// 【怎麼裝】（跟著做，約 5 分鐘）
// 1. 開一張新的 Google 試算表（網址列輸入 sheets.new）
// 2. 上方選單「擴充功能」→「Apps Script」
// 3. 把編輯器裡原本的字全部刪掉，貼上這整份程式，按存檔
// 4. 右上「部署」→「新增部署作業」→ 類型選「網頁應用程式」
//    －執行身分：我自己
//    －誰可以存取：所有人
// 5. 第一次會請你授權（會說要存取雲端硬碟），按同意
// 6. 複製產生的網址（.../exec 結尾），貼進 index.html 裡的 SCRIPT_URL
// ============================================================

var SHEET_NAME = '機票資料';
var PHOTO_FOLDER = '成就班護照資料頁';   // 護照照片會存在雲端硬碟這個資料夾

var HEADERS = [
  '送出時間','機票安排','自行安排說明','中文姓名','英文姓氏','英文名字','護照號碼','國籍',
  '出生年月日','性別','護照效期到期日','手機號碼','Email','地址','身分證字號',
  '緊急聯絡人','緊急聯絡電話','飲食習慣','活動同意事項','護照照片連結'
];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getSheet_();
    var p = e.parameter || {};

    // 有照片就存進雲端硬碟，表格只放連結
    var photoUrl = '';
    if (p.photo && p.photo.indexOf('data:image') === 0) {
      photoUrl = savePhoto_(p.photo, p.cnName || '未填姓名');
    }

    sheet.appendRow([
      new Date(),
      p.flightPlan || '', p.selfPlanNote || '',
      p.cnName || '', p.lastName || '', p.firstName || '', p.passportNo || '', p.nationality || '',
      p.birthday || '', p.gender || '', p.passportExpiry || '',
      p.phone || '', p.email || '', p.address || '', p.idNo || '',
      p.emgName || '', p.emgPhone || '', p.diet || '', p.consent || '',
      photoUrl
    ]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function savePhoto_(dataUrl, name) {
  var folders = DriveApp.getFoldersByName(PHOTO_FOLDER);
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(PHOTO_FOLDER);
  var base64 = dataUrl.split(',')[1];
  var blob = Utilities.newBlob(Utilities.base64Decode(base64), 'image/jpeg',
    name + '_' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd_HHmmss') + '.jpg');
  return folder.createFile(blob).getUrl();
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
