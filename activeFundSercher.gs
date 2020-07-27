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
      // みんかぶ暫定対応。リスクが低すぎるため、期間3ヶ月の評価をさげる
      let m = Math.abs(r) * this.sharpList[i]
      if (length === 6 && i === 0) {
        m = Math.sqrt(Math.abs(m)) * (this.sharpList[i] > 0 ? 1 : -1)
      }
      return r != null ? m : null
    })
  }
  
  sqrtTargetList() {
    const length = this.returnList.length
    return this.returnList.map((r, i) => {
      // みんかぶ暫定対応。リスクが低すぎるため、期間3ヶ月の評価をさげる
      let m = Math.sqrt(Math.abs(r)) * this.sharpList[i]
      if (length === 6 && i === 0) {
        m = Math.sqrt(Math.abs(m)) * (this.sharpList[i] > 0 ? 1 : -1)
      }
      return r != null ? m : null
    })
  }
}

// hook
function onOpen() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
  sheet.addMenu("Google App Script",  [
    {name: 'みんかぶシンプルスクレイピング', functionName: 'scrapingSimpleFromMinkabu'},
    {name: 'みんかぶスクレイピング', functionName: 'scrapingFromMinkabu'},
    {name: 'モーニングスタースクレイピング', functionName: 'scrapingFromMorningStar'},
    {name: '自動化シート一括削除', functionName: 'deleteAutoSheet'},
  ])
}
 
function deleteAutoSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  const ignoreSheet= ['過去メモ']
  ignoreSheet.forEach(sheetName => {
    SpreadsheetApp.setActiveSheet(spreadsheet.getSheetByName(sheetName))
    spreadsheet.moveActiveSheet(1)
  })
  
  spreadsheet.getSheets().forEach((sheet, i) => {
    if (i >= ignoreSheet.length) {
      spreadsheet.deleteSheet(sheet)
    }
  })
}

function scrapingSimpleFromMinkabu1() {
  const sheetName = 'みんかぶデータ1'
  const pass = '/return'
  const page = {start:1, end:7}
  const fundList = new Map()
  
  getRankingListFromMinkabu(fundList, pass, page, [])
  _scrapingSimpleFromMinkabu(sheetName, fundList)
}

function scrapingSimpleFromMinkabu2() {
  const sheetName = 'みんかぶデータ2'
  const pass = '/return'
  const page = {start:8, end:11}
  const fundList = new Map()
  
  getRankingListFromMinkabu(fundList, pass, page, [])
  _scrapingSimpleFromMinkabu(sheetName, fundList)
}

function scrapingSimpleFromMinkabu3() {
  const sheetName = 'みんかぶデータ3'
  const pass = '/return'
  const page = {start:12, end:15}
  const fundList = new Map()
  
  getRankingListFromMinkabu(fundList, pass, page, [])
  _scrapingSimpleFromMinkabu(sheetName, fundList)
}

function scrapingSimpleFromMinkabu4() {
  const sheetName = 'みんかぶデータ4'
  const fundList = new Map()
  const termListNum = 6
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

  _scrapingSimpleFromMinkabu(sheetName, fundList)
}

    
function _scrapingSimpleFromMinkabu(sheetName, fundList) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName)
  sheet.clear()
  getDetailFromMinkabu(fundList)
  outputSimpleToSheet(sheet, fundList)    
}

function scrapingFromMinkabu () {
  try {
    console.log("start")
    const sheetName = 'みんかぶ'
    const pass = {return: '/return', sharp: '/sharpe_ratio'}
    const page = {return: {start:1, end:5}, sharp: {start:1, end:1}}
    const ignoreList = ['ＤＩＡＭ新興市場日本株ファンド', 'ＳＢＩ中小型成長株ファンドジェイネクスト（ｊｎｅｘｔ）', 'ＦＡＮＧ＋インデックス・オープン', 'ＳＢＩ中小型割安成長株ファンドジェイリバイブ（ｊｒｅｖｉｖｅ）']
    const sheet = createSheet(sheetName)

    const rankingList = new Map()
    getRankingListFromMinkabu(rankingList, pass.return, page.return, ignoreList)
    console.log("getRankingListFromMinkabu:return")
    getRankingListFromMinkabu(rankingList, pass.sharp, page.sharp, ignoreList)
    console.log("getRankingListFromMinkabu:sharpRatio")
    getDetailFromMinkabu(rankingList)
    console.log("getDetailFromMinkabu")

    const [srdList, medianList, sqrtSrdList, sqrtMedianList] = analysis(rankingList, termList)
    console.log("analysis")
    outputToSheet(sheet, rankingList, srdList, medianList, sqrtSrdList, sqrtMedianList)
    console.log("outputToSheet")
  } catch(e) {
    console.error("message:" + e.message + "\nstack:" + e.stack)
    throw e
  }
}

function getRankingListFromMinkabu(rankingList, targetPass, page, ignoreList) {
  const baseUrl = 'https://itf.minkabu.jp'
  const termList = [{n:0, month:3}, {n:1, month:6}, {n:2, month:12}, {n:3, month:36}, {n:4, month:60}, {n:5, month:120}]
  termList.forEach(term => {
    for(let p=page.start; p<=page.end; p++) {
      const url = baseUrl + '/ranking' + targetPass + '?term=' + term.month + '&page=' + p
      const html = UrlFetchApp.fetch(url).getContentText()
      const table = Parser.data(html).from('<table class="md_table ranking_table">').to("</table>").build()
      const aList = Parser.data(table).from('<a class="fwb"').to("</a>").iterate()
      aList.forEach(a => {
        const result = a.match(/href="(.*)">(.*)/)
        const pass = result[1]
        const name = result[2]
        const link = baseUrl + pass + '/risk_cont'
        const ignore = ignoreList.some(i => i === name)
        rankingList.set(pass, new Ranking(name, link, termList.length, ignore, false))
      })
    }
  })
}

function getDetailFromMinkabu(rankingList) {
  const sharpNum = 12 // 12番目からのspanがシャープレシオ
  const termList = [{n:0, month:3}, {n:1, month:6}, {n:2, month:12}, {n:3, month:36}, {n:4, month:60}, {n:5, month:120}]
  console.log('getDetailFromMinkabu:' + rankingList.size)
  
  let i = 0
  rankingList.forEach(ranking => {
    const html = UrlFetchApp.fetch(ranking.link).getContentText()
    const table = Parser.data(html).from('<table class="md_table">').to('</table>').build()
    const spanList = Parser.data(table).from('<span>').to('</span>').iterate()
    
    if (ranking.name === '') {
      ranking.name = Parser.data(html).from('<p class="stock_name">').to('</p>').build()
    }
    ranking.returnList = termList.map(term => {
      const result = spanList[term.n].replace(/%/, '')
      return result != '-' ? Number(result) : null
    })
    ranking.sharpList = termList.map(term => {
      const result = spanList[sharpNum + term.n].replace(/%/, '')
      return result != '-' ? Number(result) : null
    })    
    ranking.date = Parser.data(html).from('<span class="fsm">（').to('）</span>').build()
    
    i++
    if (i%50 === 0) {
      console.log(i)
    }
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
    const termList = [{n:0, year:1}, {n:1, year:3}, {n:2, year:5}, {n:3, year:10}]
    getRankingListFromMorningStar(rankingList, returnPass, termList, characterCode, ignoreList)
    getRankingListFromMorningStar(rankingList, sharpRatioPass, termList, characterCode, ignoreList)
    getDetailFromMorningStar(rankingList, termList, characterCode)

    const [srdList, medianList, sqrtSrdList, sqrtMedianList] = analysis(rankingList, termList)
    outputToSheet(sheet, rankingList, srdList, medianList, sqrtSrdList, sqrtMedianList)
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
      if (i === 0) return // unknown garbage
  
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
  
  const [srdList, medianList] = getStatistics(targetList)
  const [sqrtSrdList, sqrtMedianList] = getStatistics(sqrtTargetList)
  return [srdList, medianList, sqrtSrdList, sqrtMedianList]
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
  const medianList = targetList.map(t => {
    t = t.sort((a, b) => a - b)
    return t[parseInt(t.length/2)]
  })
  console.log(sumList, srdList, medianList)
  return [srdList, medianList]
}1

function outputToSheet(sheet, rankingList, srdList, medianList, sqrtSrdList, sqrtMedianList) { 
  const data = []
  let n = 1
  rankingList.forEach(ranking => {
    const rankingTargetList = ranking.targetList()
    const finalTargetList = getFinalTarget(rankingTargetList, srdList, medianList)
    const finalResult = finalTargetList.reduce((acc, v) => acc + v)

    const rankingSqrtTargetList = ranking.sqrtTargetList()
    const finalSqrtTargetList = getFinalTarget(rankingSqrtTargetList, sqrtSrdList, sqrtMedianList)
    const finalSqrtResult = finalSqrtTargetList.reduce((acc, v) => acc + v)

    const row = [ranking.date, ranking.link, ranking.category, ranking.name, finalResult, finalTargetList, '', finalSqrtResult, finalSqrtTargetList].flat()
//    rankingTargetList.map((target, i) => {
//     row.push('', target, ranking.returnList[i], ranking.sharpList[i])
//    })
    data.push(row)

    if(ranking.ignore) {
      sheet.getRange(n, 4).setBackground('gray')
    }
    n++
  })
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data)

  setColors(sheet, srdList, sqrtSrdList)
}

function outputSimpleToSheet(sheet, fundList) {
  const data = []
  fundList.forEach(fund => {
    const row = [fund.date, fund.link, fund.isIdeco ? 1 : 0, fund.name, '']
    fund.returnList.forEach((r, i) => {
      row.push(r, fund.sharpList[i], '')
    })
    data.push(row)
  })
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data)

  data[0].forEach((_, i) => {
    sheet.autoResizeColumn(i + 1)
  })
}

function getFinalTarget(targetList, srdList, medianList) {
  return targetList.map((t, i) => (t || medianList[i]) / srdList[i])
}

function setColors(sheet, srdList, sqrtSrdList) {
  const colorList = ['lime', 'yellow', 'orange', 'pink']
  const targetNum = 5
  for(let i=targetNum; i < targetNum + 3 + srdList.length + sqrtSrdList.length; i++) {
    if (i === targetNum + 1 + srdList.length) continue
    sheet.getDataRange().sort({column: i, ascending: false})
    colorList.forEach((color, m) => {
      sheet.getRange(1 + 5*m, i, 5).setBackground(color)
    })
  }
  sheet.getDataRange().sort({column: targetNum, ascending: false})
  
  sheet.autoResizeColumn(3)
  sheet.autoResizeColumn(4)

  let i = 0
  const max = 10
  const white = '#ffffff' // needs RGB color
  const aqua = 'aqua'
  const range = sheet.getRange(1, 4, sheet.getLastRow() - 1)
  const rgbs = range.getBackgrounds().map(rows => {
    return rows.map(rgb => {
      if (i >= max || rgb !== white) {
        return rgb
      }
      i++;
      return aqua
    })
  })
  range.setBackgrounds(rgbs)
}
