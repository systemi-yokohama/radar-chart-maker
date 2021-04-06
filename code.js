/* global Charts, SpreadsheetApp */

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
 * スプレッドシートにレーダーチャートのシートを追加する。
 *
 * @param {Spreadsheet} ss 現在操作しているスプレッドシート
 * @param {Array.<String>} header 整形済み項目のヘッダ
 * @param {Object} categories スキルカテゴリ一覧
 * @param {Array.<String>} record レーダーチャートを作成したい対象者の情報
 */
function create (ss, header, categories, record) {
  if (record.length === 1) {
    return
  }
  const values = []
  for (let i = 0; i < record.length && record[i] !== ''; i++) {
    // データは縦に並べなければ GAS からはレーダーチャートを適切に表示できない
    values[i] = [header[i], record[i]]
  }
  // シート名は 100 文字まで
  const sheetName = `${values[1][1]} ${values[2][1]}`.substring(0, 100)
  const oldSheet = ss.getSheetByName(sheetName)
  if (oldSheet) {
    ss.deleteSheet(oldSheet)
  }
  const insertedSheet = ss.insertSheet()
  insertedSheet.setName(sheetName)
  insertedSheet.getRange(1, 1, values.length, values[0].length).setValues(values)

  let row = 4 // レーダーチャートを配置する開始位置
  for (const key in categories) {
    const category = categories[key]
    const range = insertedSheet.getRange(category.offset, 1, category.count, 2)
    const chart = insertedSheet.newChart()
      .addRange(range)
      .setChartType(Charts.ChartType.RADAR)
      .setPosition(row, 3, 0, 0)
      .setOption('title', key)
    insertedSheet.insertChart(chart.build())
    row += 18 // レーダーチャートの大きさを考慮した値
  }
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
  const { newHeader, categories } = parseHeader(header)
  records.forEach(record => create(ss, newHeader, categories, record))
}

// eslint-disable-next-line no-unused-vars
function debug (...argv) {
  const ss = SpreadsheetApp.getActive()
  const s = ss.getSheetByName('デバッグログ')
  if (s) {
    s.appendRow([
      new Date().toISOString(),
      argv.reduce((prev, cur) => `${prev} ${typeof cur === 'string' ? cur : JSON.stringify(cur)}`)
    ])
  }
}
