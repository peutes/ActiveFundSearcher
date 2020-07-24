class Ranking {
  constructor(name, link, yearListNum) {
    this.name = name
    this.link = link
    this.returnList = new Array(yearListNum).fill(null)
    this.sharpList =  new Array(yearListNum).fill(null)
  }
  
  targetList() {
    return this.returnList.map((r, i) => {
      return r != null ? Math.abs(r) * this.sharpList[i] : null  
    })
  }
  
  sqrtTargetList() {
    return this.returnList.map((r, i) => {
      return r != null ? Math.sqrt(Math.abs(r)) * this.sharpList[i] : null  
    })
  }
}

// hook
function onOpen() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
  sheet.addMenu("Google App Script",  [
    {name: 'モーニングスタースクレイピング', functionName: 'scrapingMorningStar'},
    {name: 'みんかぶスクレイピング', functionName: 'scrapingMinkabu'},
  ])
}

function scrapingMorningStar () {
  try {
    const sheetName = 'モーニングスター'
    const returnUrl = 'FundRankingReturn.do'
    const sharpRatioUrl = 'FundRankingSharpRatio.do'
    const characterCode = 'Shift_JIS'

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName)
    sheet.clear()
  
    const rankingList = new Map()  
    const yearList = [{n:0, y:1}, {n:1, y:3}, {n:2, y:5}, {n:3, y:10}]
    getRankingList(sheet, rankingList, returnUrl, yearList)
    getRankingList(sheet, rankingList, sharpRatioUrl, yearList)
    getDetail(rankingList, characterCode, yearList)
  
    const [aveList, srdList, sqrtAveList, sqrtSrdList] = analysis(rankingList, yearList)
    outputToSheet(sheet, rankingList, aveList, srdList, sqrtAveList, sqrtSrdList)
  } catch(e) {
    console.error("message:" + e.message + "\nstack:" + e.stack)
    throw e
  }
}

function getRankingList(sheet, rankingList, targetUrl, yearList) {
  const baseUrl = 'http://www.morningstar.co.jp/FundData/'
  yearList.forEach(year => {
    const html = UrlFetchApp.fetch(baseUrl + targetUrl + '?bunruiCd=all&kikan=' + year.y + 'y').getContentText('Shift_JIS')
    const table = Parser.data(html).from('<table class="table1f">').to("</table>").build()
    const trList = Parser.data(table).from('<tr>').to('</tr>').iterate()
    trList.forEach((tr, i) => {
      if (i == 0) return // unknown garbage
  
      const name = Parser.data(tr).from('target="_blank" >').to('</a>').build()
      const link = Parser.data(tr).from('<a\ href="').to('"').build()
      rankingList.set(link, new Ranking(name, baseUrl + link, yearList.length))
    })
  })
}

function getDetail(rankingList, characterCode, yearList) {
  rankingList.forEach(ranking => {
    const linkHtml = UrlFetchApp.fetch(ranking.link).getContentText(characterCode)
    const lintTable = Parser.data(linkHtml).from('<table class="table4d mb30 mt20">').to("</table>").build()    
    const tdList = lintTable.match(/<td>(.*?)<\/td>/g)
  
    const sharpNum = 40 // 40番目からのtdがシャープレシオ
    ranking.returnList = yearList.map(year => getTdList(tdList, year.n))
    ranking.sharpList = yearList.map(year => getTdList(tdList, sharpNum + year.n))
  })
}

function getTdList(tdList, i) {
  const result = /<span class="(plus|minus)">(.*)<\/span>/.exec(tdList[i])
  return result ? result[2].replace(/%/, '') : null
}

function analysis(rankingList, yearList) {
  let targetList = yearList.map(_ => [])
  let sqrtTargetList = yearList.map(_ => [])
  rankingList.forEach(ranking => {
    const rankingTargetList = ranking.targetList()
    const rankingSqrtTargetList = ranking.sqrtTargetList()

    yearList.forEach((_, i) => {
      targetList[i].push(rankingTargetList[i])
      sqrtTargetList[i].push(rankingSqrtTargetList[i])
    })
  })
  targetList = targetList.map(tr => tr.filter(t => t != null))
  sqrtTargetList = sqrtTargetList.map(tr => tr.filter(t => t != null))
  
  const [aveList, srdList] = getStatistics(targetList)
  const [sqrtAveList, sqrtSrdList] = getStatistics(sqrtTargetList)
  return [aveList, srdList, sqrtAveList, sqrtSrdList]
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
  console.log(sumList, aveList, srdList)
  return [aveList, srdList]
}

function outputToSheet(sheet, rankingList, aveList, srdList, sqrtAveList, sqrtSrdList) { 
  const data = []
  rankingList.forEach(ranking => {
    const rankingTargetList = ranking.targetList()
    const finalTargetList = getFinalTarget(rankingTargetList, aveList, srdList)
    const finalResult = finalTargetList.reduce((acc, v) => acc + v)

    const rankingSqrtTargetList = ranking.sqrtTargetList()
    const finalSqrtTargetList = getFinalTarget(rankingSqrtTargetList, sqrtAveList, sqrtSrdList)
    const finalSqrtResult = finalSqrtTargetList.reduce((acc, v) => acc + v)

    const row = [ranking.link, ranking.name, finalResult, finalTargetList, '', finalSqrtResult, finalSqrtTargetList].flat()
    rankingTargetList.map((target, i) => {
     row.push('', target, ranking.returnList[i], ranking.sharpList[i])
    })
    data.push(row)
  })
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data)
  
  const colorList = ["lime", "yellow", "orange", "pink"]
  const resultNum = 3
  for(let i=resultNum; i<6 + aveList.length + sqrtAveList.length; i++) {
    if (i == 4 + aveList.length) continue
    sheet.getDataRange().sort({column: i, ascending: false})
    colorList.forEach((color, m) => {
      sheet.getRange(1 + 5*m, i, 5).setBackground(color)
    })
  }
  sheet.getDataRange().sort({column: resultNum, ascending: false})
}

function getFinalTarget(targetList, aveList, srdList) {
  return targetList.map((t, i) => (t || aveList[i]) / srdList[i])
}
