class Ranking {
  constructor(name, link, yearsNum) {
    this.name = name
    this.link = link
    this.returnList = new Array(yearsNum).fill(null)
    this.sharpList =  new Array(yearsNum).fill(null)
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

//hook
function onOpen() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
  sheet.addMenu("Google App Script",  [
    {name: 'モーニングスタースクレイピング', functionName: 'scrapingMorningStar'},
    {name: 'みんかぶスクレイピング', functionName: 'scrapingMinkabu'},
  ])
}

function scrapingMorningStar () {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('モーニングスター')
    sheet.clear()
  
    const rankingList = new Map()
    const returnUrl = 'FundRankingReturn.do'
    const sharpRatioUrl = 'FundRankingSharpRatio.do'
  
    const years = [1, 3, 5, 10]
    getRankingList(sheet, rankingList, returnUrl, years)
    getRankingList(sheet, rankingList, sharpRatioUrl, years)
                
    getDetails(rankingList)
  
    const [aveList, srdList, sqrtAveList, sqrtSrdList] = analysis(rankingList, years)
    output(sheet, rankingList, aveList, srdList, sqrtAveList, sqrtSrdList)
  } catch(e) {
    console.error("message:" + e.message + "\nstack:" + e.stack)
    throw e
  }
}

function getRankingList(sheet, rankingList, targetUrl, years) {
  const baseUrl = 'http://www.morningstar.co.jp/FundData/'
  years.forEach(year => {
    const html = UrlFetchApp.fetch(baseUrl + targetUrl + '?bunruiCd=all&kikan=' + year + 'y').getContentText('Shift_JIS')
    const table = Parser.data(html).from('<table class="table1f">').to("</table>").build()
    const trList = Parser.data(table).from('<tr>').to('</tr>').iterate()
    trList.forEach((tr, i) => {
      if (i == 0) return // unknown garbage
  
      const name = Parser.data(tr).from('target="_blank" >').to('</a>').build()
      const link = Parser.data(tr).from('<a\ href="').to('"').build()
      rankingList.set(link, new Ranking(name, baseUrl + link, years.length))
    })
  })
}

function getDetails(rankingList) {
  rankingList.forEach(ranking => {
    const linkHtml = UrlFetchApp.fetch(ranking.link).getContentText('Shift_JIS')
    const lintTable = Parser.data(linkHtml).from('<table class="table4d mb30 mt20">').to("</table>").build()    
    const tdList = lintTable.match(/<td>(.*?)<\/td>/g)
    if (!tdList) {
      return
    }
  
    const sharpNum = 40 // 40番目からのtdがシャープレシオ    
    ranking.returnList = ranking.returnList.map((_, i) => getTdList(tdList, i))
    ranking.sharpList = ranking.sharpList.map((_, i) => getTdList(tdList, sharpNum + i))
  })
}

function getTdList(tdList, i) {
  const result = /<span class="(plus|minus)">(.*)<\/span>/.exec(tdList[i])
  return result ? result[2].replace(/%/, '') : null
}

function analysis(rankingList, years) {
  let targetList = years.map(_ => [])
  let sqrtTargetList = years.map(_ => [])
  rankingList.forEach(ranking => {
    const rankingTargetList = ranking.targetList()
    const rankingSqrtTargetList = ranking.sqrtTargetList()

    years.forEach((_, i) => {
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

function output(sheet, rankingList, aveList, srdList, sqrtAveList, sqrtSrdList) { 
  const data = []
  rankingList.forEach(ranking => {
    const rankingTargetList = ranking.targetList()
    const finalTargetList = getFinalTarget(rankingTargetList, aveList, srdList)    
    const finalResult = finalTargetList.reduce((acc, v) => acc + v)

    const rankingSqrtTargetList = ranking.sqrtTargetList()
    const finalSqrtTargetList = getFinalTarget(rankingSqrtTargetList, sqrtAveList, sqrtSrdList)
    const finalSqrtResult = finalSqrtTargetList.reduce((acc, v) => acc + v)

    data.push([
      ranking.link,
      ranking.name,
      finalResult,
      finalTargetList[0], 
      finalTargetList[1],
      finalTargetList[2],
      finalTargetList[3],
      '',
      finalSqrtResult,
      finalSqrtTargetList[0], 
      finalSqrtTargetList[1],
      finalSqrtTargetList[2],
      finalSqrtTargetList[3],
      '',
      rankingTargetList[0],
      ranking.returnList[0],
      ranking.sharpList[0],
      '',
      rankingTargetList[1],
      ranking.returnList[1],
      ranking.sharpList[1],
      '',
      rankingTargetList[2],
      ranking.returnList[2],
      ranking.sharpList[2],
      '',
      rankingTargetList[3],
      ranking.returnList[3],
      ranking.sharpList[3]
    ])
  })
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data)
  
  for(let i=13; i>=3; i--) {
    if (i==8) continue
    sheet.getDataRange().sort({column: i, ascending: false})
    sheet.getRange(1, i, 5).setBackground("lime")
    sheet.getRange(6, i, 5).setBackground("yellow")
    sheet.getRange(11, i, 5).setBackground("orange")
    sheet.getRange(16, i, 5).setBackground("pink")
  }
}

function getFinalTarget(targetList, aveList, srdList) {
  return targetList.map((t, i) => (t || aveList[i]) / srdList[i])
}
