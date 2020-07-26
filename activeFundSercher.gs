class Ranking {
  constructor(name, link, termListNum, ignore) {
    this.name = name
    this.link = link
    this.returnList = new Array(termListNum).fill(null)
    this.sharpList =  new Array(termListNum).fill(null)
    this.ignore = ignore
    this.date = null
    this.category = null
  }
  
  targetList() {
    const length = this.returnList.length
    return this.returnList.map((r, i) => {
      // みんかぶ暫定対応。リスクが低すぎるため、期間3ヶ月の評価をさげる
      if (i === 0 && length === 6) {
        return this.sharpList[i]
      }
      return r != null ? Math.abs(r) * this.sharpList[i] : null  
    })
  }
  
  sqrtTargetList() {
    const length = this.returnList.length
    return this.returnList.map((r, i) => {
      // みんかぶ暫定対応。リスクが低すぎるため、期間3ヶ月の評価をさげる
      if (i === 0 && length === 6) {
        return Number(this.sharpList[i])
      }
      return r != null ? Math.sqrt(Math.abs(r)) * this.sharpList[i] : null  
    })
  }
}

// hook
function onOpen() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
  sheet.addMenu("Google App Script",  [
    {name: 'みんかぶスクレイピング', functionName: 'scrapingFromMinkabu'},
    {name: 'モーニングスタースクレイピング', functionName: 'scrapingFromMorningStar'},
    {name: '自動化シート一括削除', functionName: 'deleteAutoSheet'},    
  ])
}

function scrapingAll() {
  scrapingFromMorningStar()
  scrapingFromMinkabu()
}

    
function deleteAutoSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  const ignoreSheet= ['過去メモ', 'iDeCo']
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
    

function scrapingFromMinkabu () {
  try {
    console.log("start")
    const sheetName = 'みんかぶ'
    const pass = {return: '/return', sharp: '/sharpe_ratio'}
    const pageNum = {return: 5, sharp: 1}
    const termList = [{n:0, month:3}, {n:1, month:6}, {n:2, month:12}, {n:3, month:36}, {n:4, month:60}, {n:5, month:120}]
    const ignoreList = ['ＤＩＡＭ新興市場日本株ファンド', 'ＳＢＩ中小型成長株ファンドジェイネクスト（ｊｎｅｘｔ）', 'ＦＡＮＧ＋インデックス・オープン', 'ＳＢＩ中小型割安成長株ファンドジェイリバイブ（ｊｒｅｖｉｖｅ）']
    const sheet = createSheet(sheetName)

    const rankingList = new Map()
    getRankingListFromMinkabu(sheet, rankingList, pass.return, termList, pageNum.return, ignoreList)
    console.log("getRankingListFromMinkabu:return")
    getRankingListFromMinkabu(sheet, rankingList, pass.sharp, termList, pageNum.sharp, ignoreList)
    console.log("getRankingListFromMinkabu:sharpRatio")
    getDetailFromMinkabu(rankingList, termList)
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

function getRankingListFromMinkabu(sheet, rankingList, targetPass, termList, pageNum, ignoreList) {
  const baseUrl = 'https://itf.minkabu.jp'
  termList.forEach(term => {
    for(let page=0; page<pageNum; page++) {
      const url = baseUrl + '/ranking' + targetPass + '?term=' + term.month + '&page=' + (page + 1)
      const html = UrlFetchApp.fetch(url).getContentText()
      const table = Parser.data(html).from('<table class="md_table ranking_table">').to("</table>").build()
      const aList = Parser.data(table).from('<a class="fwb"').to("</a>").iterate()
      aList.forEach(a => {
        const result = a.match(/href="(.*)">(.*)/)
        const pass = result[1]
        const name = result[2]
        const link = baseUrl + pass + '/risk_cont'
        const ignore = ignoreList.some(i => i === name)
        rankingList.set(pass, new Ranking(name, link, termList.length, ignore))
      })
    }
  })
}

function getDetailFromMinkabu(rankingList, termList) {
  const sharpNum = 12 // 12番目からのspanがシャープレシオ
  console.log('getDetailFromMinkabu:' + rankingList.size)
  rankingList.forEach(ranking => {
    const html = UrlFetchApp.fetch(ranking.link).getContentText()
    const table = Parser.data(html).from('<table class="md_table">').to('</table>').build()
    const spanList = Parser.data(table).from('<span>').to('</span>').iterate()
    ranking.returnList = termList.map(term => {
      const result = spanList[term.n].replace(/%/, '')
      return result != '-' ? result : null
    })
    ranking.sharpList = termList.map(term => {
      const result = spanList[sharpNum + term.n].replace(/%/, '')
      return result != '-' ? result : null
    })    
    ranking.date = Parser.data(html).from('<span class="fsm">（').to('）</span>').build()
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
    getRankingListFromMorningStar(sheet, rankingList, returnPass, termList, characterCode, ignoreList)
    getRankingListFromMorningStar(sheet, rankingList, sharpRatioPass, termList, characterCode, ignoreList)
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

function getRankingListFromMorningStar(sheet, rankingList, targetPass, termList, characterCode, ignoreList) {
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
      rankingList.set(link, new Ranking(name, baseUrl + link, termList.length, ignore))
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
  return result ? result[2].replace(/%/, '') : null
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

    const row = [ranking.date, ranking.link, ranking.category, ranking.name, finalSqrtResult, finalSqrtTargetList, '', finalResult, finalTargetList].flat()
    rankingSqrtTargetList.map((target, i) => {
     row.push('', target, ranking.returnList[i], ranking.sharpList[i])
    })
    data.push(row)

    if(ranking.ignore) {
      sheet.getRange(n, 4).setBackground('gray')
    }
    n++
  })
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data)

  const colorList = ['lime', 'yellow', 'orange', 'pink']
  const resultNum = 5
  for(let i=resultNum; i < resultNum + 3 + srdList.length + sqrtSrdList.length; i++) {
    if (i === resultNum + 1 + srdList.length) continue
    sheet.getDataRange().sort({column: i, ascending: false})
    colorList.forEach((color, m) => {
      sheet.getRange(1 + 5*m, i, 5).setBackground(color)
    })
  }
  sheet.getDataRange().sort({column: resultNum, ascending: false})
  sheet.autoResizeColumn(3)
  sheet.autoResizeColumn(4)
}

function getFinalTarget(targetList, srdList, medianList) {
  return targetList.map((t, i) => (t || medianList[i]) / srdList[i])
}
