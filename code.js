/* global Charts, DriveApp, Logger, MimeType, ScriptApp, SpreadsheetApp, UrlFetchApp */

'use strict'

/**
 * フォーム集計用スプレッドシートの Id。
 */
const SpreadSheetsId = '1LlRSVHwjP5nxg_V2X_UV3p7Cf5FS0fwXEIIHxGythl8'

/**
 * スプレッドシートを開いた際にレーダーチャートを作成するメニューをスプレッドシートに追加する。
 */
// eslint-disable-next-line no-unused-vars
function onOpen () {
  SpreadsheetApp
    .getActiveSpreadsheet()
    .addMenu('スキルチェック', [{
      name: 'レーダーチャート作成',
      functionName: 'createRadarChart'
    }])
}

/**
 * ハイパーリンク文字列を作成する。
 *
 * @param {String} title タイトル
 * @param {String} url URL
 * @returns {String} ハイパーリンク文字列
 */
const makeLinkString = (title, url) => `=HYPERLINK("${url}","${title}")`

/**
 * フォルダ ID を取得する。
 *
 * @returns {String} フォルダ ID
 */
const getFolderId = () => SpreadsheetApp.getActiveSpreadsheet()
  .getSheetByName('変数')
  .getDataRange()
  .getValues()
  .filter(row => row[0] === 'フォルダ ID')[0][1]

/**
 * Slack ウェブフックを取得する。
 *
 * @returns {String} Slack ウェブフック
 */
const getSlackWebHook = () => SpreadsheetApp.openById(SpreadSheetsId)
  .getSheetByName('変数')
  .getDataRange()
  .getValues()
  .filter(row => row[0] === 'Slack ウェブフック')[0][1]

/**
 * 新しく作成したスプレッドシートにレーダーチャートのシートを追加する。
 *
 * @param {String} folderId スプレッドシートを作成するフォルダの Id
 * @param {Array.<String>} header 整形済み項目のヘッダ
 * @param {Object} categories スキルカテゴリ一覧
 * @param {Array.<String>} record レーダーチャートを作成したい対象者の情報
 * @returns {Array.<{name:String,url:String}>|null} 追加・削除されたスプレッドシートの URL とファイル名を持つオブジェクトの配列。
 * 追加されたファイルの情報が先頭、削除されたファイルの情報がそれ以降。何もしなかった場合は null
 */
function create (folderId, header, categories, record) {
  if (record.length === 1) {
    return null
  }
  const values = []
  for (let i = 0; i < record.length && record[i] !== ''; i++) {
    // データは縦に並べなければ GAS からはレーダーチャートを適切に表示できない
    values[i] = [header[i], record[i]]
  }
  // シート名は 100 文字まで
  const spreadSheetName = `${values[1][1]} ${values[2][1]}`.substring(0, 100)
  const files = DriveApp.searchFiles(`title = '${spreadSheetName}' and trashed = false and mimeType = '${MimeType.GOOGLE_SHEETS}'`)
  const links = []
  while (files.hasNext()) {
    const file = files.next()
    links.push({ name: spreadSheetName, url: makeLinkString(spreadSheetName, file.getUrl()) })
    file.setTrashed(true) // 古いスプレッドシートは削除
  }
  // 新しいスプレッドシートを作成
  const ss = SpreadsheetApp.create(spreadSheetName)
  const file = DriveApp.getFileById(ss.getId())
  const folder = DriveApp.getFolderById(folderId)
  file.moveTo(folder)
  const s = ss.getActiveSheet()
  s.getRange(1, 1, values.length, values[0].length).setValues(values)

  let row = 4 // レーダーチャートを配置する開始位置
  for (const key in categories) {
    const category = categories[key]
    const range = s.getRange(category.offset, 1, category.count, 2)
    const chart = s.newChart()
      .addRange(range)
      .setChartType(Charts.ChartType.RADAR)
      .setPosition(row, 3, 0, 0)
      // 平滑線
      .setOption('smoothLine', false)
      .setOption('title', key)
      .setOption('curveType', 'none')
      .setOption('vAxis.minValue', 0)
      .setOption('vAxis.maxValue', 5)
    s.insertChart(chart.build())
    row += 18 // レーダーチャートの大きさを考慮した値
  }
  s.insertRows(58, 3)
  s.insertRows(115, 3)
  links.unshift({ name: spreadSheetName, url: makeLinkString(spreadSheetName, ss.getUrl()) })
  return links
}

/**
 * ヘッダを整形して返す。
 *
 * @param {Array.<String>} header 整形前のヘッダ
 * @return {Object} 整形されたヘッダ
 */
function parseHeader (header) {
  const categories = {}
  const newHeader = []
  for (let i = 0; i < header.length; i++) {
    if (i < 3) {
      newHeader[i] = header[i]
      continue
    }
    if (header[i] === '') {
      break
    }
    const splitted = header[i].split(/\[|\]|:/).filter(s => s.length !== 0).map(v => v.trim())
    if (splitted.length !== 3) {
      break
    }
    const category = splitted[0]
    const subCategory = splitted[2]
    newHeader[i] = subCategory
    categories[category] = categories[category] || { offset: null, count: 0 }
    categories[category].offset = categories[category].offset || i + 1
    categories[category].count++
  }
  return { newHeader, categories }
}

/**
 * レーダーチャートを作成する。
 *
 * @param {[Object]} values onSubmit から呼び出された場合は event.values、そうでない場合は undefined
 * @return {[String]} 生成した PDF の URL のリスト
 */
// eslint-disable-next-line no-unused-vars
function createRadarChart (values) {
  const folderId = getFolderId()
  const ss = SpreadsheetApp.getActive()
  const dataSheet = ss.getSheets()[0]
  const header = dataSheet.getRange(1, 1, 1, 256).getValues()[0]
  const records = values || ss.getActiveRangeList().getRanges().reduce((acc, cur) => acc.concat(cur.getValues()), [])
  Logger.log(`header: ${header}`)
  Logger.log(`records: ${records}`)
  const { newHeader, categories } = parseHeader(header)
  const allLinks = records.map(record => create(folderId, newHeader, categories, record)).filter(v => v !== null)
  if (allLinks.length === 0) {
    return []
  }
  const linkSheetName = 'リンク'
  const createSheet = () => {
    const s = ss.insertSheet()
    s.setName(linkSheetName)
    ss.moveActiveSheet(2) // 2 番目に移動
    return s
  }
  const s = ss.getSheetByName(linkSheetName) || createSheet()
  let rows = s.getDataRange()
    .getRichTextValues()
    .map(row => {
      const text = row[0].getText()
      const url = row[0].getLinkUrl()
      const newRow = [makeLinkString(text, url)]
      newRow.text = text
      return newRow
    })
    .filter(row => row[0] !== '' && row[0] !== '=HYPERLINK("null","")')
  const pdfUrls = []
  allLinks.forEach(links => {
    if (links.length > 1) {
      const sublinks = links.slice(1)
      const filter = row => {
        for (const info of sublinks) {
          if (info.url.includes(`"${row.text}"`)) {
            return false
          }
        }
        return true
      }
      rows = rows.filter(filter)
      rows.push([links[0].url])
    } else {
      rows.push([links[0].url])
    }
    const pdfName = `${links[0].name}.pdf`
    deletePdf(pdfName)
    const pdfUrl = savePdf(folderId, links[0].url, pdfName)
    pdfUrls.push(pdfUrl)
  })
  s.clear()
  if (rows.length !== 0) {
    s.getRange(1, 1, rows.length).setValues(rows)
  }
  return pdfUrls
}

// eslint-disable-next-line no-unused-vars
function onSubmit (event) {
  Logger.log(JSON.stringify(event))
  const pdfUrls = createRadarChart([event.values])
  const company = event.values[1]
  const name = event.values[2]
  const slackWebHook = getSlackWebHook()
  const data = {
    blocks: [{
      type: 'section',
      text: {
        type: 'mrkdwn',
        text: `「${company}」の「${name}」さんがスキルチェックシートに入力しました。\n\n<${pdfUrls[0]}}|PDFをダウンロード>`
      }
    }]
  }
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(data)
  }
  UrlFetchApp.fetch(slackWebHook, options)
}

/**
 * 指定された名前の PDF を削除する。
 *
 * @param {String} name 削除する PDF ファイル名
 */
function deletePdf (name) {
  const files = DriveApp.searchFiles(`title = '${name}' and trashed = false and mimeType = '${MimeType.PDF}'`)
  while (files.hasNext()) {
    const file = files.next()
    file.setTrashed(true)
  }
}

/**
 * 指定されたスプレッドシートから指定されたフォルダに指定された名前で PDF を保存する。
 *
 * @param {String} folderId PDF をを作成するフォルダの Id
 * @param {String} url スプレッドシートの URL
 * @param {String} name PDF のファイル名
 * @return {String} PDF のダウンロード URL
 */
function savePdf (folderId, url, name) {
  const ss = SpreadsheetApp.openByUrl(url)
  const baseUrl = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?gid=${ss.getActiveSheet().getSheetId()}`
  const pdfOptions = '&exportFormat=pdf&format=pdf' +
      '&size=A4' + // 用紙サイズ (A4)
      '&portrait=true' + // 用紙の向き true: 縦向き / false: 横向き
      '&scale=2' + // 1= 標準100%, 2= 幅に合わせる, 3= 高さに合わせる,  4= ページに合わせる
      '&fitw=true' + // ページ幅を用紙にフィットさせるか true: フィットさせる / false: 原寸大
      '&top_margin=0.40' + // 上の余白
      '&right_margin=0.50' + // 右の余白
      '&bottom_margin=0.40' + // 下の余白
      '&left_margin=0.50' + // 左の余白
      '&horizontal_alignment=CENTER' + // 水平方向の位置
      '&vertical_alignment=TOP' + // 垂直方向の位置
      '&printtitle=false' + // スプレッドシート名の表示有無
      '&sheetnames=false' + // シート名の表示有無
      '&gridlines=false' + // グリッドラインの表示有無
      '&fzr=false' + // 固定行の表示有無
      '&fzc=false' // 固定列の表示有無
  const pdfUrl = baseUrl + pdfOptions
  const option = {
    headers: {
      Authorization: `Bearer ${ScriptApp.getOAuthToken()}`
    }
  }
  const blob = UrlFetchApp.fetch(pdfUrl, option).getBlob().setName(name)
  const folder = DriveApp.getFolderById(folderId)

  const file = folder.createFile(blob)
  return file.getDownloadUrl()
}
