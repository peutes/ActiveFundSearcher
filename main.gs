const termSize = 6
const scoresSize = 3
const fundsSheetMax = 12
const purchaseNum = 20

class SheetInfo {
  constructor() {
    this.linkSheetName = 'Link'
    this.fundsSheetNames = [];
    for (let i=0; i<fundsSheetMax; i++) {
      this.fundsSheetNames.push('Funds' + i)
    }
    this.downsideRiskSheetName = 'Downside'
    this.logSheetName = 'Log'
    
    this.minkabuSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  }
  
  getScoreSheetName() {
    const prefix = 'スコア：'
    return prefix + (new Date().toLocaleString("ja")) 
  }
  
  getSheet(name) {
    return this.minkabuSpreadSheet.getSheetByName(name)
  }

  insertSheet(name) {
    return this.minkabuSpreadSheet.insertSheet(name, 0)
  }

}

class Fund {
  constructor(link, isIdeco) {
    this.link = link
    this.isIdeco = isIdeco
    
    this.date = null
    this.name = null
    this.ignore = false
    this.returns = new Array(termSize).fill(null)
    this.sharps = new Array(termSize).fill(null)

    this.scores = new Array(scoresSize)
    for (let i=0; i<scoresSize; i++) {
      this.scores[i] = new Array(termSize).fill(null)
    }
    this.totalScores = new Array(scoresSize).fill(0)
  }
}

// hook
function onOpen() {
  const menu = [{name: 'ランキング', functionName: 'scrapingMinkabuRanking'}]
  for (let i=0; i<fundsSheetMax; i++) {
    menu.push({name: 'ファンド' + i, functionName: 'scrapingMinkabuFunds' + i})
  }
  menu.push({name: 'ファンドスコア', functionName: 'calcMinkabuFundsScore'})
  SpreadsheetApp.getActiveSpreadsheet().addMenu("みんかぶ", menu)
}
