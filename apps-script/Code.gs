/**
 * 立德盃少棒報名表：接收各校「完成報名」時上傳的無核章欄 PDF，存入主辦單位的 Google 雲端硬碟。
 * 檔名為「縣市_校名_報名表.pdf」；同校再次更新報名時，舊檔移至垃圾桶並以新檔取代。
 * 聯絡人姓名、行動電話、Email 與上傳時間寫在檔案的「說明」欄位。
 */
const FOLDER_ID = '請貼上雲端硬碟資料夾ID';
const MAX_BYTES = 15 * 1024 * 1024;

/** 開啟網頁應用程式網址時，直接顯示報名表（需在專案中新增名為 Index 的 HTML 檔，貼上 index.html 內容） */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('2027第15屆立德盃少棒錦標賽')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const county = clean(body.county);
    const school = clean(body.school);
    if (!county || !school) return out({ ok: false, error: '缺少縣市或校名' });

    const bytes = Utilities.base64Decode(String(body.pdf || ''));
    if (!bytes.length || bytes.length > MAX_BYTES) return out({ ok: false, error: '檔案大小不符' });
    if (String.fromCharCode(bytes[0], bytes[1], bytes[2], bytes[3]) !== '%PDF') return out({ ok: false, error: '檔案不是 PDF' });

    const name = county + '_' + school + '_報名表.pdf';
    const c = body.contact || {};
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const folder = DriveApp.getFolderById(FOLDER_ID);
      const old = folder.getFilesByName(name);
      while (old.hasNext()) old.next().setTrashed(true);
      const file = folder.createFile(Utilities.newBlob(bytes, 'application/pdf', name));
      file.setDescription(
        '聯絡人：' + clean(c.name) + '　行動電話：' + clean(c.phone) + '　Email：' + clean(c.email) +
        '\n上傳時間：' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm')
      );
    } finally {
      lock.releaseLock();
    }
    return out({ ok: true, name: name });
  } catch (err) {
    return out({ ok: false, error: String(err && err.message || err) });
  }
}

function clean(s) {
  return String(s || '').replace(/[\\/:*?"<>|\r\n]/g, '').trim().slice(0, 60);
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
