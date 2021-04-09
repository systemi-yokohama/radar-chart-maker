/* global Charts, DriveApp, Logger, MimeType, SpreadsheetApp */

'use strict'

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
 * 指定されたファイルに対してスプレッドシートの「編集者」シートに記載されている編集者をまとめて追加する。
 *
 * @param {Object} GAS のファイルオブジェクト
 */
function addEditors (file) {
  const emailAddresses = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('編集者')
    .getDataRange()
    .getValues()
    .filter(row => row[0] !== '')
    .map(row => row[0])
  if (emailAddresses.length !== 0) {
    file.addEditors(emailAddresses)
  }
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
 * 新しく作成したスプレッドシートにレーダーチャートのシートを追加する。
 *
 * @param {Array.<String>} header 整形済み項目のヘッダ
 * @param {Object} categories スキルカテゴリ一覧
 * @param {Array.<String>} record レーダーチャートを作成したい対象者の情報
 * @returns {Array.<String>|null} 追加・削除されたスプレッドシートの URL。
 * 追加された URL が先頭、削除された URL がそれ以降。何もしなかった場合は null
 */
function create (header, categories, record) {
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
    links.push(makeLinkString(spreadSheetName, file.getUrl()))
    file.setTrashed(true) // 古いスプレッドシートは削除
  }
  // 新しいスプレッドシートを作成
  const ss = SpreadsheetApp.create(spreadSheetName)
  const file = DriveApp.getFileById(ss.getId())
  addEditors(file)
  const folderName = 'スキルチェックシート'
  const folders = DriveApp.getFoldersByName(folderName)
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName)
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
  links.unshift(makeLinkString(spreadSheetName, ss.getUrl()))
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

// eslint-disable-next-line no-unused-vars
function createRadarChart () {
  const ss = SpreadsheetApp.getActive()
  const dataSheet = ss.getSheets()[0]
  const header = dataSheet.getRange(1, 1, 1, 256).getValues()[0]
  const records = ss.getActiveRangeList().getRanges().reduce((acc, cur) => acc.concat(cur.getValues()), [])
  Logger.log(`header: ${header}`)
  Logger.log(`records: ${records}`)
  const { newHeader, categories } = parseHeader(header)
  const allLinks = records.map(record => create(newHeader, categories, record)).filter(v => v !== null)
  if (allLinks.length === 0) {
    return
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
  allLinks.forEach(links => {
    if (links.length > 1) {
      const sublinks = links.slice(1)
      const filter = row => {
        for (const link of sublinks) {
          if (link.includes(`"${row.text}"`)) {
            return false
          }
        }
        return true
      }
      rows = rows.filter(filter)
      rows.push([links[0]])
    } else {
      rows.push([links[0]])
    }
  })
  s.clear()
  if (rows.length !== 0) {
    s.getRange(1, 1, rows.length).setValues(rows)
  }
}
