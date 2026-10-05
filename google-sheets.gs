/**
 * 拍照盤點 → Google 試算表
 * 用法（只要做一次）：
 * 1. 開一個新的 Google 試算表 → 上方「擴充功能」→「Apps Script」
 * 2. 把這整段貼上（取代原本的內容）→ 存檔
 * 3. 右上「部署」→「新增部署作業」→ 類型選「網頁應用程式」
 *    執行身分：我　／　誰可以存取：所有人 → 部署 → 第一次會要你授權（選自己的帳號 → 進階 → 前往）
 * 4. 複製「網頁應用程式網址」（https://script.google.com/macros/s/…/exec），貼到拍照盤點 App 的「設定 → Google 試算表」
 * 之後每次在 App 按「完成」，這次的盤點就會寫進來；同一次盤點重複送，會先刪掉舊的再寫，不會重複。
 *
 * 會自動建立四個工作表：
 * - 盤點紀錄：每一次、每一種商品一列（原始資料，不要手動改欄位順序）
 * - 總表：每個品項在每個盤點日期各幾件（自動計算）
 * - 最新一次：最近一次盤點的品項與數量，數量多的排前面（自動計算）
 * - 品項庫：料號、實盤、帳面、差異、安全庫存、該叫貨（每次同步整張更新，不要手動改）
 */
const RAW = '盤點紀錄'
const HEAD = ['盤點日期', '時間', '位置', '品名', '品牌', '型號', '尺寸／規格', '數量', '照片框數', '來源', '盤點ID']

function doPost(e) {
  const data = JSON.parse(e.postData.contents)
  const lock = LockService.getScriptLock()
  lock.waitLock(20000)
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const sh = ensureRaw(ss)
    const rows = (data.rows || []).map((r) => {
      const row = r.slice(0, HEAD.length)
      row[0] = new Date(String(row[0]) + 'T00:00:00') // 盤點日期存成真的日期，才能排序、篩選
      return row
    })
    // 同一次盤點重新同步：先刪掉舊的那幾列
    const ids = {}
    rows.forEach((r) => (ids[String(r[10])] = true))
    const last = sh.getLastRow()
    if (last > 1) {
      const idCol = sh.getRange(2, 11, last - 1, 1).getValues()
      for (let i = idCol.length - 1; i >= 0; i--) if (ids[String(idCol[i][0])]) sh.deleteRow(i + 2)
    }
    if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, HEAD.length).setValues(rows)
    sh.getRange(2, 1, Math.max(1, sh.getLastRow() - 1), 1).setNumberFormat('yyyy/mm/dd')
    ensureSummary(ss)
    if (data.items && data.items.length) writeItems(ss, data.items)
    return json({ ok: true, rows: rows.length, items: data.items ? data.items.length - 1 : 0 })
  } finally {
    lock.releaseLock()
  }
}

/** App 的「測試連線」會呼叫這個 */
function doGet() {
  return json({ ok: true, app: '拍照盤點', sheet: SpreadsheetApp.getActiveSpreadsheet().getName() })
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON)
}

function ensureRaw(ss) {
  const sh = ss.getSheetByName(RAW) || ss.insertSheet(RAW, 0)
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEAD)
    sh.setFrozenRows(1)
    sh.getRange(1, 1, 1, HEAD.length).setFontWeight('bold').setBackground('#e8f0fe')
  }
  return sh
}

/** 品項庫：整張覆蓋（第一列是標題）；盤虧標紅、該叫貨標橘 */
function writeItems(ss, table) {
  const sh = ss.getSheetByName('品項庫') || ss.insertSheet('品項庫', 1)
  sh.clear()
  const width = table[0].length
  const rows = table.map((r) => {
    const row = r.slice(0, width)
    while (row.length < width) row.push('')
    return row
  })
  sh.getRange(1, 1, rows.length, width).setValues(rows)
  sh.setFrozenRows(1)
  sh.getRange(1, 1, 1, width).setFontWeight('bold').setBackground('#e8f0fe')
  const status = table[0].indexOf('狀態')
  if (rows.length < 2 || status < 0) return
  // 一次設定整張的底色（一列一列設會很慢）
  const colors = rows.slice(1).map((r) => {
    const s = String(r[status] || '')
    const c = s.indexOf('盤虧') >= 0 ? '#fde8e8' : s.indexOf('該叫貨') >= 0 ? '#fff4e0' : null
    return r.map(() => c)
  })
  sh.getRange(2, 1, colors.length, width).setBackgrounds(colors)
}

function ensureSummary(ss) {
  if (!ss.getSheetByName('總表')) {
    const s = ss.insertSheet('總表')
    s.getRange('A1').setValue('每個品項在每次盤點日期各幾件（自動計算，不用改）').setFontWeight('bold')
    s.getRange('A3').setFormula('=QUERY(\'盤點紀錄\'!A:K,"select D, G, sum(H) where D is not null group by D, G pivot A",1)')
    s.setFrozenRows(3)
  }
  if (!ss.getSheetByName('最新一次')) {
    const s = ss.insertSheet('最新一次')
    s.getRange('A1').setValue('最近一次盤點日期').setFontWeight('bold')
    s.getRange('B1').setFormula("=MAX('盤點紀錄'!A2:A)").setNumberFormat('yyyy/mm/dd')
    s.getRange('A3').setFormula(
      '=QUERY(\'盤點紀錄\'!A:K,"select D, G, E, F, sum(H) where A = date \'"&TEXT(B1,"yyyy-mm-dd")&"\' group by D, G, E, F order by sum(H) desc label D \'品名\', G \'尺寸／規格\', E \'品牌\', F \'型號\', sum(H) \'數量\'",1)',
    )
    s.setFrozenRows(3)
  }
}
