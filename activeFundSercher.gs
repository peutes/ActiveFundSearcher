class Ranking {
  constructor(name, link) {
    this.name = name
    this.link = link
    this.return1 = null
    this.return3 = null
    this.return5 = null
    this.return10 = null
    this.sharp1 = null
    this.sharp3 = null
    this.sharp5 = null
    this.sharp10 = null
  }
  
  get target1(){
    return this.target(this.return1, this.sharp1)
  }
  get target3(){
    return this.target(this.return3, this.sharp3)
  }
  get target5(){
    return this.target(this.return5, this.sharp5)
  }
  get target10(){
    return this.target(this.return10, this.sharp10)
  }

  get sqrtTarget1(){
    return this.sqrtTarget(this.return1, this.sharp1)
  }
  get sqrtTarget3(){
    return this.sqrtTarget(this.return3, this.sharp3)
  }
  get sqrtTarget5(){
    return this.sqrtTarget(this.return5, this.sharp5)
  }
  get sqrtTarget10(){
    return this.sqrtTarget(this.return10, this.sharp10)
  }

  target(r, s) {
    return r != null ? Math.abs(r) * s : null  
  }

  sqrtTarget(r, s) {
    return r != null ? Math.sqrt(Math.abs(r)) * s : null  
  }
}

//hook
function onOpen() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();  
  sheet.addMenu("Google App Script",  [{name: 'モーニングスタースクレイピング', functionName: 'exec'}]);
}

function exec () {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('自動化')
  sheet.clear()
  
  const rankingList = new Map()
  const returnUrl = 'FundRankingReturn.do'
  const sharpRatioUrl = 'FundRankingSharpRatio.do'

  let year = 1
  getRankingList(sheet, rankingList, returnUrl, year)
  getRankingList(sheet, rankingList, sharpRatioUrl, year)

  year = 3  
  getRankingList(sheet, rankingList, returnUrl, year)
  getRankingList(sheet, rankingList, sharpRatioUrl, year)
  
  year = 5
  getRankingList(sheet, rankingList, returnUrl, year)
  getRankingList(sheet, rankingList, sharpRatioUrl, year)

  year = 10
  getRankingList(sheet, rankingList, returnUrl, year)
  getRankingList(sheet, rankingList, sharpRatioUrl, year)
  getDetails(rankingList)
  
  const targets = [[], [], [], []]
  const sqrtTargets = [[], [], [], []]
  rankingList.forEach(ranking => {
    targets[0].push(ranking.target1)
    targets[1].push(ranking.target3)
    targets[2].push(ranking.target5)
    targets[3].push(ranking.target10)
    sqrtTargets[0].push(ranking.sqrtTarget1)
    sqrtTargets[1].push(ranking.sqrtTarget3)
    sqrtTargets[2].push(ranking.sqrtTarget5)
    sqrtTargets[3].push(ranking.sqrtTarget10)
  })

  for(let i=0; i<4; i++) {
    targets[i] = targets[i].filter(t => t != null)
    sqrtTargets[i] = sqrtTargets[i].filter(t => t != null)
  }

  const [aveList, srdList] = getStatistics(targets)
  const [sqrtAveList, sqrtSrdList] = getStatistics(sqrtTargets)
  
  const data = []
  rankingList.forEach(ranking => {
    const finalTarget1 = getFinalTarget(ranking.target1, aveList, srdList, 0)
    const finalTarget3 = getFinalTarget(ranking.target3, aveList, srdList, 1)
    const finalTarget5 = getFinalTarget(ranking.target5, aveList, srdList, 2)
    const finalTarget10 = getFinalTarget(ranking.target10, aveList, srdList, 3)
    const finalResult = finalTarget1 + finalTarget3 + finalTarget5 + finalTarget10

    const finalSqrtTarget1 = getFinalTarget(ranking.sqrtTarget1, sqrtAveList, sqrtSrdList, 0)
    const finalSqrtTarget3 = getFinalTarget(ranking.sqrtTarget3, sqrtAveList, sqrtSrdList, 1)
    const finalSqrtTarget5 = getFinalTarget(ranking.sqrtTarget5, sqrtAveList, sqrtSrdList, 2)
    const finalSqrtTarget10 = getFinalTarget(ranking.sqrtTarget10, sqrtAveList, sqrtSrdList, 3)
    const finalSqrtResult = finalSqrtTarget1 + finalSqrtTarget3 + finalSqrtTarget5 + finalSqrtTarget10

    data.push([
      ranking.link,
      ranking.name,
      finalResult,
      finalTarget1, 
      finalTarget3,
      finalTarget5,
      finalTarget10,
      '',
      finalSqrtResult,
      finalSqrtTarget1, 
      finalSqrtTarget3,
      finalSqrtTarget5,
      finalSqrtTarget10,
      '',
      ranking.target1,
      ranking.return1,
      ranking.sharp1,
      '',
      ranking.target3,
      ranking.return3,
      ranking.sharp3,
      '',
      ranking.target5,
      ranking.return5,
      ranking.sharp5,
      '',
      ranking.target10,
      ranking.return10,
      ranking.sharp10
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

function getStatistics(targets) { 
  const sumList = targets.map(t => t.reduce((acc, v) => acc + v))
  const aveList = sumList.map((s, i) => s/targets[i].length)
  const srdList = [0,0,0,0].map((s, i) => {
    targets[i].forEach(v => {
      s += Math.pow(v - aveList[i], 2)
    })
    return Math.sqrt(s / (targets[i].length - 1))
  }) 
  return [aveList, srdList]
}

function getFinalTarget(target, aveList, srdList, n) {
  return (target || aveList[n]) / srdList[n]
}

function getRankingList(sheet, rankingList, targetUrl, year) {
  const baseUrl = 'http://www.morningstar.co.jp/FundData/'
  const html = UrlFetchApp.fetch(baseUrl + targetUrl + '?bunruiCd=all&kikan=' + year + 'y').getContentText('Shift_JIS')
  const table = Parser.data(html).from('<table class="table1f">').to("</table>").build()
  const trs = Parser.data(table).from('<tr>').to('</tr>').iterate()
    
  for(let i=1; i<trs.length; i++) {
    const name = Parser.data(trs[i]).from('target="_blank" >').to('</a>').build()
    const link = Parser.data(trs[i]).from('<a\ href="').to('"').build()
    rankingList.set(link, new Ranking(name, baseUrl + link))
  }
}

function getDetails(rankingList) {
  rankingList.forEach(ranking => {
    const linkHtml = UrlFetchApp.fetch(ranking.link).getContentText('Shift_JIS')
    const lintTable = Parser.data(linkHtml).from('<table class="table4d mb30 mt20">').to("</table>").build()    
    const tds = lintTable.match(/<td>(.*?)<\/td>/g)
    if (!tds) {
      return
    }

    const sharpNum = 40 // 40番目からのtdがシャープレシオ
    let m = 0
    ranking.return1 = getTdsValue(tds, m)
    ranking.sharp1 = getTdsValue(tds, sharpNum + m)
  
    m++
    ranking.return3 = getTdsValue(tds, m)
    ranking.sharp3 = getTdsValue(tds, sharpNum + m)

    m++
    ranking.return5 = getTdsValue(tds, m)
    ranking.sharp5 = getTdsValue(tds, sharpNum + m)

    m++
    ranking.return10 = getTdsValue(tds, m)
    ranking.sharp10 = getTdsValue(tds, sharpNum + m)
  })
}


function getTdsValue(tds, m) {
  const result = /<span class="(plus|minus)">(.*)<\/span>/.exec(tds[m])
  return result ? result[2].replace(/%/, '') : null
}