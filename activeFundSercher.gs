class Ranking {
  constructor(name, link, termListNum, ignore, isIdeco) {
    this.name = name
    this.link = link
    this.returnList = new Array(termListNum).fill(null)
    this.sharpList =  new Array(termListNum).fill(null)
    this.ignore = ignore
    this.date = null
    this.category = null
    this.isIdeco = isIdeco
  }
  
  targetList() {
    const length = this.returnList.length
    return this.returnList.map((r, i) => {
      // リスクが低すぎるため、期間3ヶ月の評価をさげる
      let m = Math.abs(r) * this.sharpList[i]
      if(length === 6) {
        if(i === 0) {
          m = Math.sign(m) * Math.pow(Math.abs(m), 0.2) // 0.25 は大きい
        } else if(i <= 2) {
          m = Math.sign(m) * Math.pow(Math.abs(m), 0.8) // 採用（0.8 and 上位 3%）  or 0.85 and 上位5%
        }
      }
      return r != null ? m : null
    })
  }
  
  sqrtTargetList() {
    const length = this.returnList.length
    return this.returnList.map((r, i) => {
      // リスクが低すぎるため、期間3ヶ月の評価をさげる
      let m = Math.sqrt(Math.abs(r)) * this.sharpList[i]
      if(length === 6) {
        if(i === 0) {
          m = Math.sign(m) * Math.pow(Math.abs(m), 0.2)
        } else if(i <= 2) {
          m = Math.sign(m) * Math.pow(Math.abs(m), 0.8)
        }
      }
      return r != null ? m : null
    })
  }
}

let sheetNames = [];
const minkabuDataSheetNum = 10
for(let i=0; i<minkabuDataSheetNum; i++) {
 sheetNames.push('みんかぶデータ' + i)
}
const urlSheetName = 'みんかぶURL'
const termList6 = [{n:0, month:3}, {n:1, month:6}, {n:2, month:12}, {n:3, month:36}, {n:4, month:60}, {n:5, month:120}]
const termListNum = 6

// hook
function onOpen() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
  const menu = [
    {name: 'みんかぶクリア', functionName: 'clearMinkabuData'},
    {name: 'みんかぶURL', functionName: 'scrapingFromMinkabuURL'},
    {name: 'みんかぶスクレイピング', functionName: 'scrapingFromMinkabuDataSheet'},
    {name: 'モーニングスタースクレイピング', functionName: 'scrapingFromMorningStar'},
    {name: '自動化シート一括削除', functionName: 'deleteAutoSheet'},
  ]
  for(let i=0; i<minkabuDataSheetNum; i++) {
    menu.push({name: 'みんかぶデータ' + i, functionName: 'scrapingSimpleFromMinkabu' + i})
  }
  sheet.addMenu("Google App Script", menu)
}
 
function deleteAutoSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  const ignoreSheet= ['過去メモ']
  ignoreSheet.forEach(sheetName => {
    SpreadsheetApp.setActiveSheet(spreadsheet.getSheetByName(sheetName))
    spreadsheet.moveActiveSheet(1)
  })
  
  spreadsheet.getSheets().forEach((sheet, i) => {
    if(i >= ignoreSheet.length) {
      spreadsheet.deleteSheet(sheet)
    }
  })
}

function clearMinkabuData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(urlSheetName)
  sheet.clear()

  sheetNames.forEach(name => {
    const sheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name)
    sheet2.clear()
  })
}

function scrapingFromMinkabuURL() {
  const pageNum = 180 // とりあえず3600件まで対応。
  const fundList = new Map()
  getRankingListFromMinkabu(fundList, pageNum)

  const idList = [
    '03311112', '25311177', '93311164', '32311984', '0131F122', '0131102B', '8031114C', 'AA311169', '0231802C', '0131Q154', '89313121', '89315121', '65311058', '89312121', '0231402C',
    '89314121', '0231602C', '68311003', '0231502C', '0231202C', '0331109C', '0231702C', '8931111A', '89311025', '2931116A', '0331110A', '64315021', '64311081', '0431Q169', '89311135',
    '89313135', '0331112A', '89311164', '9C31116A', '7931211C', '79314169', '2931216A', '89312135', '0431X169', '29316153', '8931217C', '8931118A', 'AN31118A', '96311073', '96312073',
    '0431U169', '04316188', '04316186', '2931113C', '29314136', '2931316B', '0131C18A', '03317172', '03318172', '0331C177', '03319172', '0331A172', '03316183', '03312175', '03311187',
  ]
  idList.forEach(id => {
    const url = 'https://itf.minkabu.jp/fund/' + id + '/risk_cont'
    fundList.set('/fund/' + id, new Ranking('', url, termListNum, false, true))
  })

  const data = []
  fundList.forEach(fund => {
    data.push([fund.link, fund.isIdeco])
  })

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(urlSheetName)
  sheet.clear()
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data)
}

function scrapingSimpleFromMinkabu0() {
  const fundList = new Map()
  const n = 0
  _getURLList(fundList, n)
  _scrapingSimpleFromMinkabu(sheetNames[n], fundList)
}

function scrapingSimpleFromMinkabu1() {
  const n = 1
  const fundList = new Map()
  _getURLList(fundList, n)
  _scrapingSimpleFromMinkabu(sheetNames[n], fundList)
}

function scrapingSimpleFromMinkabu2() {
  const n = 2
  const fundList = new Map()
  _getURLList(fundList, n)
  _scrapingSimpleFromMinkabu(sheetNames[n], fundList)
}

function scrapingSimpleFromMinkabu3() {
  const n = 3
  const fundList = new Map()
  _getURLList(fundList, n)
  _scrapingSimpleFromMinkabu(sheetNames[n], fundList)
}

function scrapingSimpleFromMinkabu4() {
  const n = 4
  const fundList = new Map()
  _getURLList(fundList, n)
  _scrapingSimpleFromMinkabu(sheetNames[n], fundList)
}

function scrapingSimpleFromMinkabu5() {
  const n = 5
  const fundList = new Map()
  _getURLList(fundList, n)
  _scrapingSimpleFromMinkabu(sheetNames[n], fundList)
}

function scrapingSimpleFromMinkabu6() {
  const n = 6
  const fundList = new Map()
  _getURLList(fundList, n)
  _scrapingSimpleFromMinkabu(sheetNames[n], fundList)
}

function scrapingSimpleFromMinkabu7() {
  const n = 7
  const fundList = new Map()
  _getURLList(fundList, n)
  _scrapingSimpleFromMinkabu(sheetNames[n], fundList)
}

function scrapingSimpleFromMinkabu8() {
  const n = 8
  const fundList = new Map()
  _getURLList(fundList, n)
  _scrapingSimpleFromMinkabu(sheetNames[n], fundList)
}

function scrapingSimpleFromMinkabu9() {
  const n = 9
  const fundList = new Map()
  _getURLList(fundList, n)
  _scrapingSimpleFromMinkabu(sheetNames[n], fundList)
}

function scrapingFromMinkabuDataSheet () {
  console.log("start")
  const sheetName = 'みんかぶ'
  const ignoreList = [
    'ＤＩＡＭ新興市場日本株ファンド', 'ＳＢＩ中小型成長株ファンドジェイネクスト（ｊｎｅｘｔ）', 'ＦＡＮＧ＋インデックス・オープン', 'ＳＢＩ中小型割安成長株ファンドジェイリバイブ（ｊｒｅｖｉｖｅ）',
    '野村世界業種別投資シリーズ（世界半導体株投資）', '野村クラウドコンピューティング＆スマートグリッド関連株投信Ａコース', 'ＵＢＳ中国株式ファンド', 'ＵＢＳ中国Ａ株ファンド（年１回決算型）（桃源郷）',
    '野村ＳＮＳ関連株投資Ａコース', 'ダイワ／バリュー・パートナーズ・チャイナ・イノベーター・ファンド', 'グローバル全生物ゲノム株式ファンド（１年決算型）', 'テトラ・エクイティ', 'ブラックロック・ゴールド・メタル・オープンＡコース',
    'グローバル・プロスペクティブ・ファンド（イノベーティブ・フューチャー）', '野村米国ブランド株投資（円コース）毎月分配型', '野村クラウドコンピューティング＆スマートグリッド関連株投信Ｂコース',
    '野村ＳＮＳ関連株投資Ｂコース', '野村米国ブランド株投資（円コース）年２回決算型', 'ＵＳテクノロジー・イノベーターズ・ファンド（為替ヘッジあり）', 'ＵＢＳ次世代テクノロジー・ファンド'
  ]
  const sheet = createSheet(sheetName)

  const rankingList = new Map()
  getDetailFromMinkabuDataSheet(rankingList, ignoreList)
  console.log("getDetailFromMinkabuDataSheet")
  const [srdList, medList, medList2, sqrtSrdList, sqrtMedList, sqrtMedList2] = analysis(rankingList, termList6)
  console.log("analysis")
  outputToSheet(sheet, rankingList, srdList, medList, medList2, sqrtSrdList, sqrtMedList, sqrtMedList2)
  console.log("outputToSheet")  
}

function _scrapingSimpleFromMinkabu(sheetName, fundList) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName)
  sheet.clear()
  getDetailFromMinkabu(fundList)
  outputSimpleToSheet(sheet, fundList)    
}

function _getURLList(fundList, targetNum) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(urlSheetName)
  const values = sheet.getDataRange().getValues()
  const n = 360
  const end = Math.min(values.length, n * (targetNum + 1))
  for(let i=n * targetNum; i<end; i++) {
    fundList.set(values[i][0], new Ranking('' , values[i][0], termListNum, false, values[i][1]))
  }
}

function getRankingListFromMinkabu(rankingList, pageNum) {
  const baseUrl = 'https://itf.minkabu.jp'
  for(let p=1; p<=pageNum; p++) {
    const url = baseUrl + '/ranking/popular?page=' + p
    const html = UrlFetchApp.fetch(url).getContentText()
    const table = Parser.data(html).from('<table class="md_table ranking_table">').to("</table>").build()
    const aList = Parser.data(table).from('<a class="fwb"').to("</a>").iterate()
    if(aList.length === 0 || aList[0] === '<!DOCTYPE htm') {
      return
    }
    
    aList.forEach(a => {
      const result = a.match(/href="(.*)">(.*)/)
      const pass = result[1]
      const link = baseUrl + pass + '/risk_cont'
      rankingList.set(pass, new Ranking('', link, termListNum, false, false))        
    })
  }
}

function getDetailFromMinkabu(rankingList) {
  const sharpNum = 12 // 12番目からのspanがシャープレシオ
  console.log('getDetailFromMinkabu:' + rankingList.size)
  
  let i = 0
  rankingList.forEach(ranking => {
    const html = UrlFetchApp.fetch(ranking.link).getContentText()
    const table = Parser.data(html).from('<table class="md_table">').to('</table>').build()
    const spanList = Parser.data(table).from('<span>').to('</span>').iterate()
    
    ranking.name = Parser.data(html).from('<p class="stock_name">').to('</p>').build()
    ranking.returnList = termList6.map(term => {
      const result = spanList[term.n].replace(/%/, '')
      return result != '-' ? Number(result) : null
    })
    ranking.sharpList = termList6.map(term => {
      const result = spanList[sharpNum + term.n].replace(/%/, '')
      return result != '-' ? Number(result) : null
    })
    ranking.date = Parser.data(html).from('<span class="fsm">（').to('）</span>').build()
    
    i++
    if(i%50 === 0) {
      console.log(i)
    }
  })
}

function getDetailFromMinkabuDataSheet(rankingList, ignoreList) { 
  sheetNames.forEach(sheetName => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName)
    console.log(sheetName)
    const values = sheet.getDataRange().getValues()
    values.forEach(value => {
      if(value.length === 1) {
        return
      }
    
      const link = value[1]
      const isIdeco = value[2]
      const name = value[3]
      const ignore = ignoreList.some(j => j === name)
      const ranking = new Ranking(name, link, termListNum, ignore, isIdeco)      
      ranking.date = value[0]
      
      for(let i=0; i<termListNum; i++) {
        const r = value[5 + 3 * i]
        const s = value[6 + 3 * i]
        if(r !== '') {
          ranking.returnList[i] = r
        }
        if(s !== '') {
          ranking.sharpList[i] = s
        }
      }
      rankingList.set(link, ranking)
    })
  })
}
    
function scrapingFromMorningStar() {
  try {
    const sheetName = 'モーニングスター'
    const returnPass = 'FundRankingReturn.do'
    const sharpRatioPass = 'FundRankingSharpRatio.do'
    const characterCode = 'Shift_JIS'
    const ignoreList = ['DIAM 新興市場日本株ファンド', 'SBI 中小型成長株F ジェイネクスト 『愛称：jnext』', 'FANG+インデックス・オープン', 'SBI 中小型割安成長株F ジェイリバイブ 『愛称：jrevive』', 'ゲーム&amp;eスポーツ･オープン', '(NEXT FUNDS)NASDAQ-100(R)連動型上場投信 『愛称：NASDAQ-100ETF』']
    const sheet = createSheet(sheetName)

    const rankingList = new Map()
    const termList4 = [{n:0, year:1}, {n:1, year:3}, {n:2, year:5}, {n:3, year:10}]
    getRankingListFromMorningStar(rankingList, returnPass, termList4, characterCode, ignoreList)
    getRankingListFromMorningStar(rankingList, sharpRatioPass, termList4, characterCode, ignoreList)
    getDetailFromMorningStar(rankingList, termList4, characterCode)

    const [srdList, medList, medList2, sqrtSrdList, sqrtMedList, sqrtMedList2] = analysis(rankingList, termList4)
    outputToSheet(sheet, rankingList, srdList, medList, medList2, sqrtSrdList, sqrtMedList, sqrtMedList2)
  } catch(e) {
    console.error("message:" + e.message + "\nstack:" + e.stack)
    throw e
  }
}

function createSheet(sheetName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  return spreadsheet.insertSheet(sheetName + ":" + (new Date().toLocaleString("ja")), 0)
}

function getRankingListFromMorningStar(rankingList, targetPass, termList, characterCode, ignoreList) {
  const baseUrl = 'http://www.morningstar.co.jp/FundData/'
  termList.forEach(term => {
    const html = UrlFetchApp.fetch(baseUrl + targetPass + '?bunruiCd=all&kikan=' + term.year + 'y').getContentText(characterCode)
    const table = Parser.data(html).from('<table class="table1f">').to('</table>').build()
    const trList = Parser.data(table).from('<tr>').to('</tr>').iterate()
    const date = Parser.data(html).from('<span class="ptdate">').to('</span>').build()
    const category = Parser.data(html).from('<td class="fcate">').to('</td>').build()
    trList.forEach((tr, i) => {
      if(i === 0) return // unknown garbage
  
      const name = Parser.data(tr).from('target="_blank" >').to('</a>').build()
      const link = Parser.data(tr).from('<a\ href="').to('"').build()
      const ignore = ignoreList.some(i => i === name)
      rankingList.set(link, new Ranking(name, baseUrl + link, termList.length, ignore, false))
    })
  })
}

function getDetailFromMorningStar(rankingList, termList, characterCode) {
  const sharpNum = 40 // 40番目からのtdがシャープレシオ
  rankingList.forEach(ranking => {
    const html = UrlFetchApp.fetch(ranking.link).getContentText(characterCode)
    const table = Parser.data(html).from('<table class="table4d mb30 mt20">').to('</table>').build()    
    const tdList = table.match(/<td>(.*?)<\/td>/g)
    
    ranking.date = Parser.data(html).from('<span class="ptdate">').to('</span>').build()
    ranking.category = Parser.data(html).from('<td class="fcate">').to('</td>').build()
    ranking.returnList = termList.map(term => getTdListFromMorningStar(tdList, term.n))
    ranking.sharpList = termList.map(term => getTdListFromMorningStar(tdList, sharpNum + term.n))
  })
}

function getTdListFromMorningStar(tdList, i) {
  const result = /<span class="(plus|minus)">(.*)<\/span>/.exec(tdList[i])
  return result ? Number(result[2].replace(/%/, '')) : null
}

function analysis(rankingList, termList) {
  let targetList = termList.map(_ => [])
  let sqrtTargetList = termList.map(_ => [])
  rankingList.forEach(ranking => {
    const rankingTargetList = ranking.targetList()
    const rankingSqrtTargetList = ranking.sqrtTargetList()

    termList.forEach((_, i) => {
      targetList[i].push(rankingTargetList[i])
      sqrtTargetList[i].push(rankingSqrtTargetList[i])
    })
  })
  targetList = targetList.map(tr => tr.filter(t => t != null))
  sqrtTargetList = sqrtTargetList.map(tr => tr.filter(t => t != null))
  
  const [srdList, medList, medList2] = getStatistics(targetList)
  const [sqrtSrdList, sqrtMedList, sqrtMedList2] = getStatistics(sqrtTargetList)
  return [srdList, medList, medList2, sqrtSrdList, sqrtMedList, sqrtMedList2]
}

function getStatistics(targetList) {
  const sumList = targetList.map(t => t.reduce((acc, v) => acc + v))
  const aveList = sumList.map((s, i) => s/targetList[i].length)
  const srdList = targetList.map((t, i) => {
    let sum = 0
    t.forEach(v => {
      sum += Math.pow(v - aveList[i], 2)
    })
    return Math.sqrt(sum / (t.length - 1))
  })

  // 上位3%
  const top = 3
  const medList = targetList.map(t => {
    t = t.sort((a, b) => a - b)
    return t[parseInt(t.length*(100 - top)/100)]
  })
  
  // 上位50%  iDeCo用 #TODO
  const medList2 = targetList.map(t => {
    t = t.sort((a, b) => a - b)
    return t[parseInt(t.length/2)]
  })
  console.log(sumList, srdList, aveList, medList, medList2)
  return [srdList, medList, medList2]
}1

function outputToSheet(sheet, rankingList, srdList, medList, medList2, sqrtSrdList, sqrtMedList, sqrtMedList2) { 
  const data = []
  rankingList.forEach(ranking => {
    const rankingTargetList = ranking.targetList()
    const finalTargetList = getFinalTarget(rankingTargetList, srdList, medList)
    const finalResult = finalTargetList.reduce((acc, v) => acc + v)
    const finalTargetList2 = getFinalTarget(rankingTargetList, srdList, medList2)
    const finalResult2 = finalTargetList2.reduce((acc, v) => acc + v)

    const rankingSqrtTargetList = ranking.sqrtTargetList()
    const finalSqrtTargetList = getFinalTarget(rankingSqrtTargetList, sqrtSrdList, sqrtMedList)
    const finalSqrtResult = finalSqrtTargetList.reduce((acc, v) => acc + v)
//    const finalSqrtTargetList2 = getFinalTarget(rankingSqrtTargetList, sqrtSrdList, sqrtMedList2)
//    const finalSqrtResult2 = finalSqrtTargetList2.reduce((acc, v) => acc + v)

    const row = [ranking.link, ranking.date, ranking.name, ranking.isIdeco, finalResult, finalTargetList, '', finalSqrtResult, finalSqrtTargetList,　'', finalResult2, finalTargetList2, ''].flat()
    rankingTargetList.map((target, i) => {
      row.push('', target, ranking.returnList[i], ranking.sharpList[i])
    })
    data.push(row)
  })
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data)

  let n = 1
  const targetRow = 5
  rankingList.forEach(ranking => {
    if(ranking.ignore) {
      sheet.getRange(n, targetRow - 2).setBackground('gray')
    }
    if(ranking.isIdeco) {
      sheet.getRange(n, targetRow - 1).setBackground('yellow')
    }
    n++
  })
  
  setColors(sheet, targetRow, 3)
  sheet.getRange(1, targetRow, sheet.getLastRow()).setFontWeight("bold")
  sheet.autoResizeColumns(1, 4)
}

function outputSimpleToSheet(sheet, fundList) {
  const data = []
  fundList.forEach(fund => {
    const row = [fund.date, fund.link, fund.isIdeco, fund.name, '']
    fund.returnList.forEach((r, i) => {
      row.push(r, fund.sharpList[i], '')
    })
    data.push(row)
  })
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data)
  sheet.autoResizeColumns(2, 4)
}

function getFinalTarget(targetList, srdList, medList) {
  return targetList.map((t, i) => ((t || medList[i]) / srdList[i]) || 0)
}

function setColors(sheet, targetRow, listNum) {
  const colorList = ['cyan', 'lime', 'yellow', 'orange', 'pink', 'silver']
  for(let i=targetRow; i < targetRow + listNum*(2 + termListNum) - 1; i++) {
    let c = false
    for(let j=1; j<listNum; j++) {
      if(i === targetRow + j*(termListNum + 2) - 1) {
        c = true
      }
    }
    if (c) {
      continue
    }
    
    sheet.getDataRange().sort({column: i, ascending: false})
    colorList.forEach((color, m) => {
      sheet.getRange(1 + 5*m, i, 5).setBackground(color)
    })
  }
  
  const allRange = sheet.getDataRange()
  
  const idecoRankRow = targetRow + 2*(2 + termListNum)
  allRange.sort({column: idecoRankRow, ascending: false})
  allRange.sort({column: 4, ascending: false})
  const idecoRange = sheet.getRange(1, idecoRankRow, sheet.getLastRow())
  const idecoRgbs = setAquaRankColor(5, idecoRange)
  idecoRange.setBackgrounds(idecoRgbs)
  
  allRange.sort({column: targetRow, ascending: false})
  const range = sheet.getRange(1, targetRow - 2, sheet.getLastRow())
  const rgbs = setAquaRankColor(10, range)
  range.setBackgrounds(rgbs)
}

function setAquaRankColor(max, range) {
  const white = '#ffffff' // needs RGB color
  const aqua = 'aqua'
  let i = 0

  return range.getBackgrounds().map(rows => {
    return rows.map(rgb => {
      if(i >= max || rgb !== white) {
        return rgb
      }
      i++;
      return aqua
    })
  })
}