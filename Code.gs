// FitTrack Google Apps Script Backend
// 貼到 Google Apps Script 編輯器後，填入你的 Spreadsheet ID，然後部署為 Web App

const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID'  // ← 填入你的 Sheet ID
const SHEET_NAME = 'WorkoutLogs'

function getSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID)
  let sheet = ss.getSheetByName(SHEET_NAME)
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME)
    sheet.appendRow(['id', '日期', '課表', '動作', '時長(分)', 'data_json'])
    sheet.setFrozenRows(1)
    sheet.getRange('A1:F1').setFontWeight('bold')
    sheet.setColumnWidth(6, 600)
  }
  return sheet
}

function doGet(e) {
  try {
    const sheet = getSheet()
    const rows = sheet.getDataRange().getValues()
    const logs = rows.slice(1)
      .filter(r => r[0] && r[5])
      .map(r => { try { return JSON.parse(r[5]) } catch(e) { return null } })
      .filter(Boolean)
      .sort((a, b) => new Date(b.date) - new Date(a.date))
    return ok({ logs })
  } catch(err) {
    return fail(err.message)
  }
}

function doPost(e) {
  try {
    const log = JSON.parse(e.postData.contents)
    if (!log || !log.id) return fail('invalid payload')
    const sheet = getSheet()
    const existing = sheet.getDataRange().getValues().slice(1)
    const alreadyExists = existing.some(r => String(r[0]) === String(log.id))
    if (alreadyExists) return ok({ saved: false, reason: 'duplicate' })
    const date = new Date(log.date).toLocaleDateString('zh-TW')
    const durMin = Math.round((log.durationMs || 0) / 60000)
    const exNames = (log.exercises || []).map(e => e.name.split(' ')[0]).join(', ')
    sheet.appendRow([log.id, date, log.dayLabel + ' ' + log.dayName, exNames, durMin, JSON.stringify(log)])
    return ok({ saved: true })
  } catch(err) {
    return fail(err.message)
  }
}

function ok(data) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, ...data }))
    .setMimeType(ContentService.MimeType.JSON)
}

function fail(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: false, error: msg }))
    .setMimeType(ContentService.MimeType.JSON)
}
